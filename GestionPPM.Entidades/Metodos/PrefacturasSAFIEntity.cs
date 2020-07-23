using GestionPPM.Entidades.Helpers;
using GestionPPM.Entidades.Modelo;
using GestionPPM.Entidades.Modelo.SistemaContable;
using GestionPPM.Repositorios;
using Newtonsoft.Json;
using NLog;
using RestSharp;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data.Entity;
using System.Linq;
using System.Net.Http; 

namespace GestionPPM.Entidades.Metodos
{
    public static class PrefacturasSAFIEntity
    {
        private static readonly GestionPPMEntities db = new GestionPPMEntities();
        private static Logger logger = LogManager.GetCurrentClassLogger();
         
        public static RespuestaTransaccion PrefacturarSAFI(int idCodigoCotizacion)
        {
            using (var transaction = db.Database.BeginTransaction())
            {
                try
                {
                    //Obtener listado facturas para SaFi 
                    var ListadoPrefacturaSAFI = db.ListadoPresupuestoPrefacturaSAFI(idCodigoCotizacion).ToList();
                    int codigoCliente = 0;
                    var codigoProducto = "";
                    var RespuestaPrincipal = new Respuesta();
                    var mensaje = "";

                    //Validar que el cliente este creado en el safi
                    foreach (var cliente in ListadoPrefacturaSAFI)
                    {
                        codigoCliente = Convert.ToInt32(db.ConsultarDatosClienteSAFI(cliente.Identificacion).First().CodigoCliente);
                        try
                        {
                            codigoProducto = cliente.CodigoProdcuto.ToString();
                        }
                        catch (Exception)
                        {

                            return new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeProductoNoExiste };
                        }
                        
                    }

                    //Si el Cliente no existe no puede prefacturar
                    if (codigoCliente == 0)
                    {
                        return new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeClienteNoExiste};                        
                    }
                    else if(codigoProducto.Length < 3)
                    {
                        return new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeProductoNoExiste };
                    }
                    else
                    {
                        //Obtener el secuencial 
                        var secuencial = db.usp_g_codigo_documento("CodigoCotizacion").First().secuencial; 

                        //Generar Prefactura
                        var data = ProcesaCotizacionesSAFI(ListadoPrefacturaSAFI, secuencial.ToString());

                        foreach (var prefactura in data)
                        {
                            var tRespuesta = new Respuesta();
                            var estado = "OK";
                            var trama = JsonConvert.SerializeObject(prefactura);
                            var respuesta = InsertCostumerRequest(trama);
                            Respuesta2 obj = JsonConvert.DeserializeObject<Respuesta2>(respuesta.Content);

                            if (obj != null)
                            {
                                if (Convert.ToInt32(obj.codigoRetorno) != 200)
                                {
                                    if (obj.mensaje.Any() && obj.mensaje != null)
                                    mensaje = obj.mensaje;
                                    estado = "ERROR";
                                    tRespuesta = new Respuesta { mensaje = mensaje, codigoRetorno = obj.codigoRetorno.ToString(), estado = "ERROR", numeroDocumento = "" };
                                    return new RespuestaTransaccion { Estado = false, Respuesta = tRespuesta.mensaje };
                                }
                                else
                                {
                                    estado = "OK";
                                    tRespuesta = new Respuesta { mensaje = "PROCESO OK", codigoRetorno = obj.codigoRetorno.ToString(), estado = "OK", numeroDocumento = obj.numeroDocumento };

                                    if (Convert.ToInt32(obj.codigoRetorno) != 200)
                                    { 
                                        mensaje += obj.mensaje;
                                        estado = "ERROR";
                                        tRespuesta = new Respuesta { mensaje = mensaje, codigoRetorno = obj.codigoRetorno.ToString(), estado = "ERROR", numeroDocumento = obj.numeroDocumento };
                                        return new RespuestaTransaccion { Estado = false, Respuesta = tRespuesta.mensaje };
                                    }
                                    else
                                    {
                                        //Insertar el registro de Prefactura
                                        SAFIGeneral prefacturaSAFI = new SAFIGeneral();

                                        foreach (var ele in ListadoPrefacturaSAFI)
                                        {
                                            prefacturaSAFI.id_codigo_cotizacion = idCodigoCotizacion;
                                            prefacturaSAFI.detalle_cotizacion = ele.Comentario1;
                                            prefacturaSAFI.correos_facturacion = ele.Correos;
                                            prefacturaSAFI.numero_pago = ele.Pago;
                                            prefacturaSAFI.id_codigo_producto = ele.Id_Producto;
                                            prefacturaSAFI.id_forma_pago = ele.IdFormaPago;
                                            prefacturaSAFI.id_centro_costos = ele.IdCentroCosto;
                                            prefacturaSAFI.cantidad = ele.Cantidad;
                                            prefacturaSAFI.precio_unitario = ele.PrecioUnitario;
                                            prefacturaSAFI.subtotal_pago = ele.Subtotal;
                                            prefacturaSAFI.iva_pago = (ele.Total - ele.PrecioUnitario);
                                            prefacturaSAFI.descuento_pago = ele.Descuento;
                                            prefacturaSAFI.total_pago = ele.Total;
                                            prefacturaSAFI.fecha_prefactura = DateTime.Now;
                                            prefacturaSAFI.numero_prefactura = obj.numeroDocumento;
                                            prefacturaSAFI.estado = true;
                                        }

                                        db.SAFIGeneral.Add(prefacturaSAFI);
                                        db.SaveChanges();

                                        //Insertar los detalles

                                        SAFIGeneralDetalle detalle = new SAFIGeneralDetalle();

                                        foreach (var ele in ListadoPrefacturaSAFI)
                                        {
                                            detalle = new SAFIGeneralDetalle();

                                            detalle.id_facturacion_safi = prefacturaSAFI.id_facturacion_safi;
                                            detalle.id_codigo_producto = ele.Id_Producto;
                                            detalle.id_forma_pago = ele.IdFormaPago;
                                            detalle.id_centro_costos = ele.IdCentroCosto;
                                            detalle.cantidad = ele.Cantidad;
                                            detalle.precio_unitario = ele.PrecioUnitario;
                                            detalle.subtotal_pago = ele.Subtotal;
                                            detalle.iva_pago = (ele.Total - ele.PrecioUnitario);
                                            detalle.descuento_pago = ele.Descuento;
                                            detalle.total_pago = ele.Total;
                                            detalle.estado = true;

                                            db.SAFIGeneralDetalle.Add(detalle);
                                            db.SaveChanges();
                                        }

                                        estado = "OK";
                                        tRespuesta = new Respuesta { mensaje = "PROCESO OK", codigoRetorno = obj.codigoRetorno.ToString(), estado = "OK", numeroDocumento = obj.numeroDocumento };

                                        transaction.Commit();
                                        return new RespuestaTransaccion { Estado = true, Respuesta = Mensajes.MensajeTransaccionExitosa };
                                    }
                                }
                            }
                            else
                            {
                                estado = "ERROR";
                                tRespuesta = new Respuesta { mensaje = "Servicios caídos, consulte con su proveedor.", codigoRetorno = "400", estado = "ERROR", numeroDocumento = "" };
                                return new RespuestaTransaccion { Estado = false, Respuesta = tRespuesta.mensaje };
                            }
                        }

                        return new RespuestaTransaccion { Estado = false, Respuesta = mensaje};
                    }
                }catch (Exception ex)
                {
                    transaction.Rollback();
                    return new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeTransaccionFallida + " ;" + ex.Message.ToString() };
                }
            }
        }

        public static RespuestaTransaccion EliminarPrefactura(int id)
        {
            try
            {
                var safiGeneral = db.SAFIGeneral.Find(id);

                if (safiGeneral.estado == true)
                {
                    safiGeneral.estado = false;
                }
                else
                {
                    safiGeneral.estado = true;
                }

                db.Entry(safiGeneral).State = EntityState.Modified;
                db.SaveChanges();

                return new RespuestaTransaccion { Estado = true, Respuesta = Mensajes.MensajeTransaccionExitosa };
            }
            catch (Exception ex)
            {
                return new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeTransaccionFallida + " ;" + ex.Message.ToString() };
            }
        }
         
        public static List<ListadoCodigosPrefacturar> ListadoCodigosPrefacturar()
        {
            try
            {
                return db.ListadoCodigosPrefacturar().ToList();
            }
            catch (Exception ex)
            {
                throw;
            }
        }
          
        public static List<Wrapper> ProcesaCotizacionesSAFI(List<ListadoPresupuestoPrefacturaSAFI> ListadoPrefactura, string secuencial)
        {
            var facturas = new List<Wrapper>();
            try
            {
                if (ListadoPrefactura.Any())
                {
                    foreach (var ele in ListadoPrefactura)
                    {
                        var w = new Wrapper();
                        var cab = new Cabecera();
                        var det = new Detalle();
                        
                        //Cliente
                        var cliente = new ClienteSAFI
                        {
                            CodigoCliente = ele.Identificacion.ToString(),
                            NombreCliente = ele.Cliente,
                            Identificacion = ele.Identificacion,
                            Direccion = ele.Direccion,
                            Mail = ele.Correos,
                            Telefono = "",
                            Segmento = ConfigurationManager.AppSettings.Get("segmentoComercial")
                        };
                        cab.Cliente = cliente; 

                        //Detalle de la cotizacion
                        var cotizacion = new DetallesCotizacion
                        {
                            Vencimiento = ele.Fecha,
                            Comentario1 = "-",
                            PlazoDias = "0",
                            FhComent = ele.Comentario1.ToString(),
                            FhComent1 = ele.Comentario2.ToString(),
                            FhComent2 = "-",
                            Bodega = ele.Bodega,
                            FormaPago = ele.CodigoFormaPago,
                        };

                        cab.CotizacionDetalle = cotizacion;
                        w.Cabecera = cab; 
                           
                        //Datos del detalle de la factura
                        var listDet = new List<DetalleFactura>();
                        var detItem = new DetalleFactura
                        {
                            Cantidad = Convert.ToInt32(ele.Cantidad.ToString()),
                            Detalle = "", 
                            Valor = Convert.ToDecimal(ele.Subtotal.ToString()), //subtotal+iva
                            SubTotal = Convert.ToDecimal(ele.Subtotal.ToString()),// !string.IsNullOrEmpty(item.SubTotal) ? decimal.Parse(item.SubTotal) : 0,
                            Descuento = Convert.ToDecimal(ele.Descuento.ToString()),
                            Total = Convert.ToDecimal(ele.Total.ToString()),
                            CodigoCategoria = ConfigurationManager.AppSettings.Get("codigoCategoria"),
                            CodigoProducto = ele.CodigoProdcuto,
                            NombreProducto = ele.NombreProducto,
                            RUCProveedor = "",
                            Proveedor = "",
                            CostoUnitario = Convert.ToDecimal(ele.Subtotal.ToString()),
                            FechaVenta = Convert.ToDateTime(ele.Fecha.ToString()),
                            PorcentajeIva = Convert.ToDecimal(ele.PorcentajeIva.ToString()),
                        };
                        listDet.Add(detItem);

                        //Datos Factura
                        var fact = new Factura
                        {
                            Secuencial = secuencial,
                            Observacion = "",
                            SubTotal = Convert.ToDecimal(ele.Subtotal.ToString()),
                            Total = Convert.ToDecimal(ele.Total.ToString()),
                            Descuento = Convert.ToDecimal(ele.Descuento.ToString()),
                            PorcentajeIva = Convert.ToDecimal(ele.PorcentajeIva.ToString()),
                        };


                        fact.DetalleFactura.AddRange(listDet);

                        det.Factura = fact;
                        w.Detalle = det;

                        facturas.Add(w);
                    }
                }
            }
            catch (Exception ex)
            {
                logger.Error(ex, "Stopped program because of exception");
            }

            return facturas;
        }

        public static IRestResponse InsertCostumerRequest(string json)
        {
            using (var client = new HttpClient())
            {
                try
                {
                    var client1 = new RestClient(ConfigurationManager.AppSettings.Get("wsCotizaciones"));
                    var request = new RestRequest(Method.POST);  
                    request.AddParameter("application/json", json, ParameterType.RequestBody);
                    IRestResponse response = client1.Execute(request);
                     
                    return response;
                }
                catch (Exception ex)
                {
                    logger.Info("Llamada Web Service.");
                    logger.Error(ex, "Stopped program because of exception");
                    //var texto = ex.Message;
                }
                return null;
            }
        }

        public static List<ListadoPresupuestosAprobacionEjecutivo> ListadoPresupuestoAprobadosEjecutivo(int id_usuario)
        {
            List<ListadoPresupuestosAprobacionEjecutivo> listado = new List<ListadoPresupuestosAprobacionEjecutivo>();
            try
            {
                listado = db.ListadoPresupuestosAprobacionEjecutivo(id_usuario).ToList();
                return listado;
            }
            catch (Exception ex)
            {
                return listado;
            }
        }

        public static List<PrefacturaSAFIInfo> LitadoReversosPrefacturas()
        {
            List<PrefacturaSAFIInfo> listado = new List<PrefacturaSAFIInfo>();
            try
            {
                listado = db.ListadoPrefacturasAprobadasReversar().ToList();
                return listado;
            }
            catch (Exception ex)
            {
                return listado;
            }
        }

        public static List<PrefacturaSAFIInfo> ListadoFacturasAprobadas()
        {
            using (var context = new GestionPPMEntities())
            {
                List<PrefacturaSAFIInfo> listado = new List<PrefacturaSAFIInfo>();
                try
                {
                    listado = db.ListadoFacturasAprobadas().ToList();
                    return listado;
                }
                catch (Exception ex)
                {
                    return listado;
                }
            }

            
        }

    }
}
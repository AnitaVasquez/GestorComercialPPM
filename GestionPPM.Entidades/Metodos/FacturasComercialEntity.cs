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
    public static class FacturasComercialEntity
    {
        private static readonly GestionPPMEntities db = new GestionPPMEntities();
        private static Logger logger = LogManager.GetCurrentClassLogger();
         
        public static RespuestaTransaccion FacturarSAFI(int id_prefactura_id)
        {
            using (var transaction = db.Database.BeginTransaction())
            {
                try
                { 
                    //Obtener listado facturas para SaFi 
                    var ListadoFacturarSAFI = db.ListadoFacturarSAFI(id_prefactura_id).ToList();

                    //Obtener el secuencial 
                    var secuencial = db.usp_g_codigo_documento("Factura").First().secuencial; 

                    //Generar Prefactura
                    var data = ProcesaFacturaSAFI(ListadoFacturarSAFI, secuencial.ToString(), id_prefactura_id);

                    var tRespuesta = new Respuesta();

                    foreach (var prefactura in data)
                    {
                        var estado = "OK";
                        var trama = JsonConvert.SerializeObject(prefactura);
                        var respuesta = InsertCostumerRequest(trama);
                        Respuesta2 obj = JsonConvert.DeserializeObject<Respuesta2>(respuesta.Content);

                        if (obj != null)
                        {
                            if (Convert.ToInt32(obj.codigoRetorno) != 200)
                            {
                                var mensaje = "";
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
                                    var mensaje = "";
                                    mensaje += obj.mensaje;

                                    estado = "ERROR";
                                    tRespuesta = new Respuesta { mensaje = mensaje, codigoRetorno = obj.codigoRetorno.ToString(), estado = "ERROR", numeroDocumento = obj.numeroDocumento };
                                    return new RespuestaTransaccion { Estado = false, Respuesta = tRespuesta.mensaje };
                                }
                                else
                                { 
                                    //Insertar el registro de Prefactura
                                    SAFIGeneral prefacturaSAFI = new SAFIGeneral();

                                    prefacturaSAFI = db.SAFIGeneral.Find(id_prefactura_id);
                                     
                                    prefacturaSAFI.fecha_factura = DateTime.Now;
                                    prefacturaSAFI.numero_factura = obj.numeroDocumento;  
                                     
                                    db.Entry(prefacturaSAFI).State = EntityState.Modified;
                                    db.SaveChanges();

                                    ActualizacionSAFI actualizacion = new ActualizacionSAFI();
                                    actualizacion.id_prefactura_safi = id_prefactura_id;
                                    actualizacion.id_estado_safi = false;

                                    db.ActualizacionSAFI.Add(actualizacion);
                                    db.SaveChanges();

                                    //Actualizar la Factura
                                    //var valor = db.ActualizarFacturaSAFI();
                                    //db.SaveChanges();

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

                    return new RespuestaTransaccion { Estado = false, Respuesta = tRespuesta.mensaje };
                }
                catch (Exception ex)
                {
                    transaction.Rollback();
                    return new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeTransaccionFallida + " ;" + ex.Message.ToString() };
                }
            }
        }

        public static RespuestaTransaccion ActualizarDetallePrefactura(SAFIGeneral prefactura)
        {
            try
            {
                // Por si queda el Attach de la entidad y no deja actualizar
                var local = db.SAFIGeneral.FirstOrDefault(f => f.id_facturacion_safi == prefactura.id_facturacion_safi);
                if (local != null)
                {
                    db.Entry(local).State = EntityState.Detached;
                }
                 
                db.Entry(prefactura).State = EntityState.Modified;
                db.SaveChanges();

                return new RespuestaTransaccion { Estado = true, Respuesta = Mensajes.MensajeTransaccionExitosa };
            }
            catch (Exception ex)
            {
                return new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeTransaccionFallida + " ;" + ex.Message.ToString() };
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

        public static SAFIGeneral ConsultarPrefactura(int id)
        {
            try
            {
                SAFIGeneral prefactura = db.SAFIGeneral.Find(id);
                return prefactura;
            }
            catch (Exception)
            {
                throw;
            }
        }

        public static List<ListadoPresupuestosFacturar> ListadoPrefacturasFacturar()
        {
            try
            {
                return db.ListadoPresupuestosFacturar().ToList();
            }
            catch (Exception ex)
            {
                throw;
            }
        }
          
        public static List<Wrapper> ProcesaFacturaSAFI(List<ListadoFacturarSAFI> ListadoFactura, string secuencial, int id_prefactura)
        {
            var facturas = new List<Wrapper>();
            try
            {
                if (ListadoFactura.Any())
                {
                    foreach (var ele in ListadoFactura)
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
                            Vencimiento = DateTime.Now.ToString(),
                            Comentario1 = "-",
                            PlazoDias = "0",
                            FhComent = ele.Detalle.ToString(),
                            FhComent1 = ele.Detalle.ToString(),
                            FhComent2 = "-",
                            Bodega = ele.CodigoBodega,
                            UGE = ele.CentroCosto,
                            FormaPago = ele.FormaPago
                        };

                        cab.CotizacionDetalle = cotizacion;
                        w.Cabecera = cab;

                        //Datos Factura
                        var fact = new Factura
                        {
                            Secuencial = secuencial,
                            Observacion = "",
                            SubTotal = Convert.ToDecimal(ele.Subtotal.ToString()),
                            Total = Convert.ToDecimal(ele.Total.ToString()),
                            Descuento = Convert.ToDecimal(ele.Descuento.ToString())
                        };

                        List<ConsultarPrefacturaSAFIDetalle> detallePrefactura = CotizacionEntity.ConsultarPrefacturaSAFIDetalle(id_prefactura);
                        
                        //Datos del detalle de la factura
                        var listDet = new List<DetalleFactura>();

                        foreach (var detallado in detallePrefactura)
                        {

                            var detItem = new DetalleFactura
                            {
                                Cantidad = Convert.ToInt32(detallado.cantidad.ToString()),
                                Detalle = ele.Detalle,
                                Valor = Convert.ToDecimal(detallado.subtotal_pago.ToString()), //subtotal+iva
                                SubTotal = Convert.ToDecimal(detallado.subtotal_pago.ToString()),// !string.IsNullOrEmpty(item.SubTotal) ? decimal.Parse(item.SubTotal) : 0,
                                Descuento = Convert.ToDecimal(detallado.descuento_pago.ToString()),
                                Total = Convert.ToDecimal(detallado.total_pago.ToString()),
                                CodigoCategoria = ConfigurationManager.AppSettings.Get("codigoCategoriaComercial"),
                                CodigoProducto = detallado.codigo_producto,
                                NombreProducto = detallado.nombre_producto,
                                RUCProveedor = "",
                                Proveedor = "",
                                CostoUnitario = Convert.ToDecimal(detallado.precio_unitario.ToString()),
                                FechaVenta = Convert.ToDateTime(DateTime.Now.ToString()),
                            };
                            listDet.Add(detItem);
                        } 

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
                    var client1 = new RestClient(ConfigurationManager.AppSettings.Get("wsFacturasPresupuestos"));
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

    }
}
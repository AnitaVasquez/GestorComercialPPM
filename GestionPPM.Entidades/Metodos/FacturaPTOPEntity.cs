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
    public static class FacturaPTOPEntity
    {
        private static readonly GestionPPMEntities db = new GestionPPMEntities();
        private static Logger logger = LogManager.GetCurrentClassLogger();
         
        public static List<ListadoFacturasPTOPSAFI> ListarListadoPTOPSAFI()
        {
            try
            {
                return db.ListadoFacturasPTOPSAFI().ToList();
            }
            catch (Exception ex)
            {
                throw;
            }
        }

        public static List<GeneraFactruasPTOPSAFI> ListadoFacturasPTOPSAFI()
        {
            try
            {
                return db.GeneraFactruasPTOPSAFI().ToList();
            }
            catch (Exception ex)
            {
                throw;
            }
        }

        public static List<CalculoValoresFacturaPTOP> CalcularValoresPTOP(string nombrePlan, int numeroTransacciones, decimal valorTransacciones, bool esCertificacion, decimal porcentajeIva, decimal valorCertificado, decimal descuento) 
        {
            try
            {
                return db.CalculoValoresFacturaPTOP(nombrePlan, numeroTransacciones, valorTransacciones, esCertificacion, porcentajeIva, valorCertificado, descuento).ToList();
            }
            catch (Exception)
            {
                throw;
            }
        }

        public static List<FacturaPTOPSAFI> ConsultarFactura(int id)
        {
            try
            { 
                return db.FacturaPTOPSAFI(id).ToList();  
            }
            catch (Exception)
            {
                throw;
            }
        }

        public static RespuestaTransaccion ActualizarFactura(FacturaPTOPSAFI factura)
        {
            using (var transaction = db.Database.BeginTransaction())
            {
                try
                {
                    //nota de credito null
                    var notaCredito = factura.numero_nota_credito;
                    if (notaCredito == null)
                    {
                        notaCredito = "";
                    }
                    db.Actualizar_FacturaPTOP(Convert.ToInt32(factura.secuencial), factura.numero_factura.ToString(), notaCredito);
                    transaction.Commit();

                    return new RespuestaTransaccion { Estado = true, Respuesta = Mensajes.MensajeTransaccionExitosa };
                }
                catch (Exception ex)
                {
                    transaction.Rollback();
                    return new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeTransaccionFallida + " ;" + ex.Message.ToString() };
                } 
            }
        }

        public static RespuestaTransaccion AnularFacturaNormal(FacturaPTOPSAFI factura, int accion)
        {
            using (var transaction = db.Database.BeginTransaction())
            {
                try
                {
                    db.Anular_FacturaPTOP(Convert.ToInt32(factura.id_factura_PTOP), accion);
                    db.SaveChanges();
                    transaction.Commit();

                    //refacturar
                    if (accion == 1)
                    {
                        //Mandar a generar las facturas 
                        db.GeneraFactruasPTOP();
                        db.SaveChanges(); 

                        //Obtener listado facturas para SAFI 
                        var ListadoPTOPSAFI = ListadoFacturasPTOPSAFI();

                        var data = CargaPTOPEntity.ProcesaFacturaSAFI(ListadoPTOPSAFI);

                        foreach (var facturar in data)
                        {
                            var tRespuesta = new Respuesta();
                            var estado = "OK";
                            var trama = JsonConvert.SerializeObject(facturar);
                            var respuesta = CargaPTOPEntity.InsertCostumerRequest(trama);
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
                                    }
                                    else
                                    {
                                        estado = "OK";
                                        tRespuesta = new Respuesta { mensaje = "PROCESO OK", codigoRetorno = obj.codigoRetorno.ToString(), estado = "OK", numeroDocumento = obj.numeroDocumento };

                                        //Actualizar el codigo de la factura 
                                        db.Actualizar_FacturaPTOP(Convert.ToInt32(facturar.Detalle.Factura.Secuencial), obj.numeroDocumento,"");
                                        db.SaveChanges();
                                         
                                    }
                                }
                            }
                            else
                            {
                                estado = "ERROR";
                                tRespuesta = new Respuesta { mensaje = "Servicios caídos, consulte con su proveedor.", codigoRetorno = "400", estado = "ERROR", numeroDocumento = "" };
                            }
                        }
                    }

                    return new RespuestaTransaccion { Estado = true, Respuesta = Mensajes.MensajeTransaccionExitosa };
                }
                catch (Exception ex)
                {
                    transaction.Rollback();
                    return new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeTransaccionFallida + " ;" + ex.Message.ToString() };
                }
            }
        }
          
        public static List<Wrapper> ProcesaFacturaSAFI(List<GeneraFactruasPTOPSAFI> ListadoSAFI)
        {
            var facturas = new List<Wrapper>();
            try
            {
                if (ListadoSAFI.Any())
                {
                    foreach (var ele in ListadoSAFI)
                    {
                        var w = new Wrapper();
                        var cab = new Cabecera();
                        var det = new Detalle();
                        
                        //Cliente
                        var cliente = new ClienteSAFI
                        {
                            CodigoCliente = ele.idCliente.ToString(),
                            NombreCliente = ele.Cliente,
                            Identificacion = ele.Identificacion,
                            Direccion = ele.Direccion,
                            Mail = ele.Correos,
                            Telefono = ele.Telefono,
                            Segmento = ConfigurationManager.AppSettings.Get("segmento")
                        };
                        cab.Cliente = cliente;
                        w.Cabecera = cab;
                          
                        //Datos Factura
                        var fact = new Factura
                        {
                            Secuencial = ele.ID.ToString(),
                            Observacion = ele.observaciones.ToString(),
                            SubTotal = Convert.ToDecimal(ele.subtotal.ToString()),
                            Total = Convert.ToDecimal(ele.total.ToString()), 
                            Descuento = Convert.ToDecimal(ele.descuento.ToString())
                        };

                        //Datos del detalle de la factura
                        var listDet = new List<DetalleFactura>();
                        var detItem = new DetalleFactura
                        {
                            Cantidad = Convert.ToInt32(ele.cantidad.ToString()),
                            Detalle = ele.detalle.ToString(), 
                            Valor = Convert.ToDecimal(ele.subtotal.ToString()), //subtotal+iva
                            SubTotal = Convert.ToDecimal(ele.subtotal.ToString()),// !string.IsNullOrEmpty(item.SubTotal) ? decimal.Parse(item.SubTotal) : 0,
                            Descuento = Convert.ToDecimal(ele.descuento.ToString()),
                            Total = Convert.ToDecimal(ele.total.ToString()), 
                            CodigoCategoria = "7",
                            CodigoProducto = ConfigurationManager.AppSettings.Get("codProducto"),
                            RUCProveedor = "",
                            Proveedor = "",
                            CostoUnitario = Convert.ToDecimal(ele.subtotal.ToString()),
                            FechaVenta = Convert.ToDateTime(ele.fecha_factura.ToString()),
                        };
                        listDet.Add(detItem);
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
                    var client1 = new RestClient("http://172.19.3.72/p2ppruebas/api/factura");
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
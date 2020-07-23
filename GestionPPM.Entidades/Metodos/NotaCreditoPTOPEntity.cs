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
    public static class NotaCreditoPTOPEntity
    {
        private static readonly GestionPPMEntities db = new GestionPPMEntities();
        private static Logger logger = LogManager.GetCurrentClassLogger();
         
        public static RespuestaTransaccion CrearNotaCredito(int id_factura, string motivo)
        {
            using (var transaction = db.Database.BeginTransaction())
            {
                try
                {
                    //Datos de la nota de la factura
                    FacturaPTOPSAFI factura = db.FacturaPTOPSAFI(id_factura).FirstOrDefault();
                    Cliente cliente = db.Cliente.Find(factura.id_cliente);
                    var secuencial = db.GeneraCodigoNotaCredito().FirstOrDefault();

                    NotaCreditoPTOP notaCredito = new NotaCreditoPTOP();
                    notaCredito.id_factura = id_factura;
                    notaCredito.motivo = motivo;
                    notaCredito.costo_unitario = factura.precio_unitario;
                    notaCredito.subtotal = factura.subtotal;
                    notaCredito.total = factura.total;
                    notaCredito.fecha_nota_credito = DateTime.Now;
                    notaCredito.estado_nota_credito = true;
                    notaCredito.secuencial = secuencial.Secuencial;

                    db.NotaCreditoPTOP.Add(notaCredito);
                    db.SaveChanges();  
                    transaction.Commit();
                     
                    //Mandar a generar las facturas 
                    db.GeneraFactruasPTOP();
                    db.SaveChanges(); 

                    //Procesar data de nota de credito
                    var data = ProcesarNotaCreditoSAFI(notaCredito, factura.numero_factura, cliente.ruc_ci_cliente);

                    foreach (var facturar in data)
                    {
                        var tRespuesta = new Respuesta();
                        var estado = "OK";
                        var trama = JsonConvert.SerializeObject(facturar);
                        var respuesta = EnviarNotaCredito(trama);
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

                                    //Actualizar el numero de nota de credito
                                    NotaCreditoPTOP ActualizarNC = db.NotaCreditoPTOP.Find(notaCredito.id_nota_credito_PTOP);
                                    ActualizarNC.numero_nota_credito = obj.numeroDocumento;

                                    db.Entry(ActualizarNC).State = EntityState.Modified;
                                    db.SaveChanges(); 

                                    //actualizar en el registro de la factura
                                    db.Actualizar_FacturaPTOP_NC(Convert.ToInt32(factura.id_factura_PTOP), obj.numeroDocumento);
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
                return new RespuestaTransaccion { Estado = true, Respuesta = Mensajes.MensajeTransaccionExitosa };
                }
                catch (Exception ex)
                {
                    transaction.Rollback();
                    return new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeTransaccionFallida + " ;" + ex.Message.ToString() };
                }
            }
        }
         
        //Inactivar NC Fisica
        public static RespuestaTransaccion InactivarNotaCreditoPTOP(int idFactura)
        {
            try
            {
                //obtenerl la nota de credio
                var NotaCredito = db.NotaCreditoPTOP.Where( nc => nc.id_factura == idFactura).FirstOrDefault();

                NotaCredito.estado_nota_credito = false;
                  
                db.Entry(NotaCredito).State = EntityState.Modified;
                db.SaveChanges();

                //Activar la factura nuevamente
                db.Activar_FacturaPTOP(idFactura);
                db.SaveChanges();

                return new RespuestaTransaccion { Estado = true, Respuesta = Mensajes.MensajeTransaccionExitosa };
            }
            catch (Exception ex)
            {
                return new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeTransaccionFallida + " ;" + ex.Message.ToString() };
            }
        }
         
        public static List<WraperNotaCreditoWs> ProcesarNotaCreditoSAFI(NotaCreditoPTOP notaCredito, string numero_factura, string cedula_ruc)
        {
            var notasDeCredito = new List<WraperNotaCreditoWs>();

            try
            {
                if (notaCredito != null)
                {
                    var w = new WraperNotaCreditoWs();
                    var cab = new CabeceraWs();

                    //Datos Factura
                    var datosNotaCredito = new NotaCreditoWs
                    {
                        Idfactura = numero_factura,
                        Motivo = notaCredito.motivo,
                        Valor = Convert.ToDecimal(notaCredito.total.ToString()),
                        Secuencial = notaCredito.secuencial.ToString(),
                    };

                    //Datos del detalle de la factura
                    var listDet = new List<DetalleNotaCreditoWs>();
                    var detItem = new DetalleNotaCreditoWs
                    {
                        Cantidad = Convert.ToInt32(1),
                        Detalle = notaCredito.motivo,
                        Valor = Convert.ToDecimal(notaCredito.total.ToString()),
                        SubTotal = Convert.ToDecimal(notaCredito.subtotal.ToString()),
                        CodigoCategoria = ConfigurationManager.AppSettings.Get("codigoCategoria"),
                        CodigoProducto = ConfigurationManager.AppSettings.Get("codProducto"),
                        RUCProveedor = cedula_ruc,
                        Proveedor = "",
                        CostoUnitario = Convert.ToDecimal(notaCredito.costo_unitario.ToString()),
                        FechaVenta = Convert.ToDateTime(notaCredito.fecha_nota_credito.ToString()),
                    };

                    listDet.Add(detItem);
                    datosNotaCredito.DetalleNotaCredito.AddRange(listDet);

                    cab.Detalle = datosNotaCredito;
                    w.NotaCredito = cab;

                    notasDeCredito.Add(w);
                }
            }
            catch (Exception ex)
            {
                logger.Error(ex, "Stopped program because of exception");
            }

            return notasDeCredito;
        }

        public static IRestResponse EnviarNotaCredito(string json)
        {
            using (var client = new HttpClient())
            {
                try
                { 
                    var client1 = new RestClient(ConfigurationManager.AppSettings.Get("wsNotaCredito"));
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
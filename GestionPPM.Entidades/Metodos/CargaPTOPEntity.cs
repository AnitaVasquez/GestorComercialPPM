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
    public static class CargaPTOPEntity
    {
        private static readonly GestionPPMEntities db = new GestionPPMEntities();
        private static Logger logger = LogManager.GetCurrentClassLogger();
         
        public static RespuestaTransaccion ActualizarCargaPTOP(CargaPTOP carga)
        {
            using (var transaction = db.Database.BeginTransaction())
            {
                try
                {
                    // Por si queda el Attach de la entidad y no deja actualizar
                    var local = db.CargaPTOP.FirstOrDefault(c => c.id_carga_ptop == carga.id_carga_ptop);
                    if (local != null)
                    {
                        db.Entry(local).State = EntityState.Detached;
                    }

                    CargaPTOP cargaActualizar = db.CargaPTOP.Find(carga.id_carga_ptop);
                    cargaActualizar.detalle = carga.detalle;
                    cargaActualizar.observaciones = carga.observaciones;
                    cargaActualizar.id_facturable_certificacion = carga.id_facturable_certificacion;
                    cargaActualizar.id_facturable_mensual = carga.id_facturable_mensual;
                    cargaActualizar.valor_certificacion = carga.valor_certificacion;
                    
                    db.Entry(cargaActualizar).State = EntityState.Modified;
                    db.SaveChanges();  

                    transaction.Commit();


                    //Mandar a generar las facturas 
                    db.GeneraFactruasPTOP();
                    db.SaveChanges();


                    //Obtener listado facturas para SaFi 
                    var ListadoPTOPSAFI = ListadoFacturasPTOPSAFI();

                    var data = ProcesaFacturaSAFI(ListadoPTOPSAFI);

                    foreach (var facturar in data)
                    {
                        var tRespuesta = new Respuesta();
                        var estado = "OK";
                        var trama = JsonConvert.SerializeObject(facturar);
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
                return new RespuestaTransaccion { Estado = true, Respuesta = Mensajes.MensajeTransaccionExitosa };
                }
                catch (Exception ex)
                {
                    transaction.Rollback();
                    return new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeTransaccionFallida + " ;" + ex.Message.ToString() };
                }
            }
        }
         
        //Eliminación Fisica
        public static RespuestaTransaccion EliminarCargaPTOP(int id)
        {
            try
            {
                var cargaPTOP = db.CargaPTOP.Find(id);
                  
                db.Entry(cargaPTOP).State = EntityState.Deleted;
                db.SaveChanges();

                return new RespuestaTransaccion { Estado = true, Respuesta = Mensajes.MensajeTransaccionExitosa };
            }
            catch (Exception ex)
            {
                return new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeTransaccionFallida + " ;" + ex.Message.ToString() };
            }
        }

        public static List<ListadoCargaPTOP> ListarCargaPTOP()
        {
            try
            {
                return db.ListadoCargaPTOP().ToList();
            }
            catch (Exception ex)
            {
                throw;
            }
        }

        public static List<ListadoCargaPTOPExcel> ListarCargaPTOPExcel()
        {
            try
            {
                return db.ListadoCargaPTOPExcel().ToList();
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

        public static CargaPTOP ConsultarCargaPTOP(int id)
        {
            try
            {
                CargaPTOP carga = new CargaPTOP();
                var cargaOriginal = db.ConsultarCargaPTOP(id).FirstOrDefault();
                carga.id_carga_ptop = cargaOriginal.id_carga_ptop;
                carga.id_comercio = cargaOriginal.id_comercio;
                carga.id_plan = cargaOriginal.id_plan;
                carga.id_facturable_certificacion = cargaOriginal.id_facturable_certificacion;
                carga.id_facturable_mensual = cargaOriginal.id_facturable_mensual;
                carga.observaciones = cargaOriginal.observaciones;
                carga.valor_certificacion = cargaOriginal.valor_certificacion;
                carga.detalle = cargaOriginal.detalle;

                return carga;
            }
            catch (Exception)
            {
                throw;
            }
        }

        public static RespuestaTransaccion CrearActualizarRegistrosPTOPCargaMasiva(List<CargaPTOPExcel> cargarPTOP)
        {
            using (var transaction = db.Database.BeginTransaction())
            {
                try
                {
                    foreach (var item in cargarPTOP)
                    {
                        if (item.id_comercio != null)
                        {
                            item.id_comercio = item.id_comercio.ToUpper();
                        }

                        if (item.plan != null)
                        {
                            item.plan = item.plan.ToUpper();
                        }

                        //Actualizar el correo del cliente
                        var correo = item.email.Split(';');

                        //validar que el comercio exista
                        var comercioActualizar = ComercioEntity.ListarComercios().FirstOrDefault(c => c.ID_Comercio == item.id_comercio);

                        //verificar si existe para crear o actualizar
                        var existenciaCarga = db.CargaPTOP.Where(s => s.id_comercio == comercioActualizar.Codigo && s.mes ==item.mes && s.anio == item.anio 
                        && s.trx_aprobadas == item.transacciones_aprobadas && s.trx_rechazadas == item.transacciones_rechazadas && s.valor_certificacion == item.valor_certificacion
                        && s.monto_vendido_aprobado == item.monto_aprobado && s.monto_vendido_rechazado == item.monto_rechazado && s.estado_transaccion == true).FirstOrDefault();

                        if (existenciaCarga != null) // Actualiza el cliente
                        {
                            transaction.Rollback();
                            return new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeTransaccionFallida + " ;" + "La linea "+ item.id + " ya existe en el sistema, revise el archivo y vuelva a cargalo al sistema" };
                        }
                        else
                        { 
                            // Guarda la nueva linea a facturar PTOP                           
                            var cargaPTOPNuevo = new CargaPTOP();

                            //validar que el comercio exista
                            var comercio = ComercioEntity.ListarComercios().FirstOrDefault(c => c.ID_Comercio == item.id_comercio);

                            //obtener plan                    
                            var plan = TablaPlanesEntity.ListarTablaPlanes().FirstOrDefault(c => c.Nombre_Plan == item.plan.ToUpper()
                            && c.Transaccion_Minima <= item.transacciones_aprobadas && c.Transaccion_Maxima >= item.transacciones_aprobadas && c.Estado_Plan.ToUpper() == "ACTIVO");

                            if (plan == null)
                            {
                                //ver si el plan no tiene valores infinito infinito
                                plan = TablaPlanesEntity.ListarTablaPlanes().FirstOrDefault(c => c.Nombre_Plan == item.plan.ToUpper()
                                && c.Transaccion_Minima <= item.transacciones_aprobadas && c.Transaccion_Maxima == 0 && c.Estado_Plan.ToUpper() == "ACTIVO");
                            }
                            
                            //obtener el codigo de facturable certificacion
                            var facturableCertificacion = CatalogoEntity.ListadoCatalogosPorCodigo("FCR-01").FirstOrDefault(c => c.Text == item.facturable_certificacion.ToUpper());
                            
                            //obtener el codigo de facturable mensual
                            var facturableMensual = CatalogoEntity.ListadoCatalogosPorCodigo("FMS-01").FirstOrDefault(c => c.Text == item.facturable_mensual.ToUpper());

                            cargaPTOPNuevo.id_comercio = comercio.Codigo;
                            cargaPTOPNuevo.id_plan = plan.Codigo;
                            cargaPTOPNuevo.id_facturable_certificacion = Convert.ToInt32(facturableCertificacion.Value.ToString());
                            cargaPTOPNuevo.id_facturable_mensual = Convert.ToInt32(facturableMensual.Value.ToString());
                            cargaPTOPNuevo.mes = item.mes;
                            cargaPTOPNuevo.anio = item.anio;
                            cargaPTOPNuevo.detalle = item.detalle; 
                            cargaPTOPNuevo.valor_certificacion = item.valor_certificacion;
                            cargaPTOPNuevo.trx_aprobadas = item.transacciones_aprobadas;
                            cargaPTOPNuevo.trx_rechazadas = item.transacciones_rechazadas;
                            cargaPTOPNuevo.monto_vendido_aprobado = item.monto_aprobado;
                            cargaPTOPNuevo.monto_vendido_rechazado = item.monto_rechazado;
                            cargaPTOPNuevo.observaciones = item.observaciones;
                            cargaPTOPNuevo.correos = item.email;
                            cargaPTOPNuevo.fecha_carga = DateTime.Now;
                            cargaPTOPNuevo.id_factura_PTOP = 0;
                            cargaPTOPNuevo.estado_transaccion = true;
 
                            db.CargaPTOP.Add(cargaPTOPNuevo);
                            db.SaveChanges();

                            //Hacer el envio una por una al SAFI
                            //var comercioActual = db.Comercio.Find(comercio.ID_Comercio);
                            //var clienteActualizar = db.Cliente.Find(comercioActual.id_cliente);                  

                        }
                    }

                    transaction.Commit();

                    //Mandar a generar las facturas 
                    db.GeneraFactruasPTOP();
                    db.SaveChanges();


                    //Obtener listado facturas para SaFi 
                    var ListadoPTOPSAFI = ListadoFacturasPTOPSAFI();

                    var data = ProcesaFacturaSAFI(ListadoPTOPSAFI);

                    foreach (var facturar in data)
                    { 
                        var tRespuesta = new Respuesta();
                        var estado = "OK";
                        var trama = JsonConvert.SerializeObject(facturar);
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
                                    db.Actualizar_FacturaPTOP(Convert.ToInt32(facturar.Detalle.Factura.Secuencial), obj.numeroDocumento, "");
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
                            Secuencial = ele.Secuencial.ToString(),
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
                            CodigoCategoria = ConfigurationManager.AppSettings.Get("codigoCategoria"),
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
                    var client1 = new RestClient(ConfigurationManager.AppSettings.Get("wsFacturas"));
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
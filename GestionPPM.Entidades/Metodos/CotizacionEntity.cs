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
using System.Threading.Tasks;
using System.Web;
using System.Web.Mvc;

namespace GestionPPM.Entidades.Metodos
{
    public class CotizacionEntity 
    {
        //private static readonly GestionPPMEntities db = new GestionPPMEntities();

        // Define the context here.
        private static GestionPPMEntities db = new GestionPPMEntities();
        private static Logger logger = LogManager.GetCurrentClassLogger();

        //public CotizacionEntity() {
        //    db = new GestionPPMEntities();
        //}

        public static RespuestaTransaccion CrearCotizador(Cotizador cotizadorCabecera, List<DetalleCotizador> cotizadorDetalles, bool aplicaContrato)
        {
            using (var context = new GestionPPMEntities())
            {
                using (var transaction = context.Database.BeginTransaction())
                {
                    try
                    {
                        CodigoCotizacion objeto = new CodigoCotizacion();
                        //var codigoCotizacion1 = db.ConsultarCodigoCotizacion(cotizadorCabecera.id_codigo_cotizacion).FirstOrDefault();

                        // Guardando Aplica Contrato en Código de Cotización
                        //var codigoCotizacion = context.CodigoCotizacion.Find(cotizadorCabecera.id_codigo_cotizacion);

                        var codigoCotizacion = db.ConsultarCodigoCotizacion(cotizadorCabecera.id_codigo_cotizacion).FirstOrDefault();

                        //codigoCotizacion.ejecutivo = codigoCotizacion1.ejecutivo;
                        //codigoCotizacion.id_cliente = codigoCotizacion1.id_cliente;

                        codigoCotizacion.aplica_contrato = aplicaContrato;

                        PropertyCopier<ConsultarCodigoCotizacion, CodigoCotizacion>.Copy(codigoCotizacion, objeto);

                        //context.Entry(codigoCotizacion).State = EntityState.Modified;
                        //context.SaveChanges();

                        context.CodigoCotizacion.Attach(objeto); // State = Unchanged
                        context.SaveChanges();


                        var idCotizador = context.usp_g_cabecera_cotizador(cotizadorCabecera.id_codigo_cotizacion, cotizadorCabecera.numero_dias, cotizadorCabecera.fecha_cotizacion, cotizadorCabecera.fecha_vencimiento, cotizadorCabecera.observacion, cotizadorCabecera.subtotal, cotizadorCabecera.porc_descuento, cotizadorCabecera.valor_descuento, cotizadorCabecera.id_impuesto, Convert.ToDecimal(cotizadorCabecera.porcentaje_iva), cotizadorCabecera.valor_iva, cotizadorCabecera.total).FirstOrDefault().id_cotizador;

                        foreach (var item in cotizadorDetalles)
                        {
                            context.usp_g_detalle_cotizador(idCotizador, item.entregable, item.costo_total_entregable, item.tipo_servicio, item.cantidad_servicio, item.valor_unitario_servicio, item.costo_total_servicio, item.tipo_gestion, item.cantidad_gestion, item.valor_unitario_gestion, item.costo_total_gestion);
                        }

                        transaction.Commit();
                        return new RespuestaTransaccion { Estado = true, Respuesta = Mensajes.MensajeTransaccionExitosa, CotizadorID = idCotizador };
                    }
                    catch (Exception ex)
                    {
                        transaction.Rollback();
                        return new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeTransaccionFallida + " ;" + ex.Message.ToString() };
                    }
                }
            }
        }

        public static RespuestaTransaccion ActualizarCotizador(Cotizador cotizadorCabecera, List<DetalleCotizador> cotizadorDetalles, bool aplicaContrato)
        {
            using (var transaction = db.Database.BeginTransaction())
            {
                try
                {
                    // Guardando Aplica Contrato en Código de Cotización
                    var codigoCotizacion1 = db.ConsultarCodigoCotizacion(cotizadorCabecera.id_codigo_cotizacion).FirstOrDefault();

                    //
                    CodigoCotizacion codigoCotizacion =  db.CodigoCotizacion.Find(cotizadorCabecera.id_codigo_cotizacion);

                    codigoCotizacion.ejecutivo = codigoCotizacion1.ejecutivo;
                    codigoCotizacion.aplica_contrato = aplicaContrato;
                    db.Entry(codigoCotizacion).State = EntityState.Modified;
                    db.SaveChanges();

                    // Por si queda el Attach de la entidad y no deja actualizar
                    var local = db.Cotizador.FirstOrDefault(f => f.id_cotizador == cotizadorCabecera.id_cotizador);
                    if (local != null)
                    {
                        db.Entry(local).State = EntityState.Detached;
                    }

                    db.Entry(cotizadorCabecera).State = EntityState.Modified;
                    db.SaveChanges();

                    // Solo si se agregan detalles a la cotizacion
                    if (cotizadorDetalles != null)
                    {
                        // Eliminando anteriores
                        var detallesCotizadorAnteriores = db.DetalleCotizador.Where(s => s.id_cotizador == cotizadorCabecera.id_cotizador).ToList();
                        foreach (var item in detallesCotizadorAnteriores)
                        {
                            db.DetalleCotizador.Remove(item);
                            db.SaveChanges();
                        }

                        //Agregando nuevos detalles
                        foreach (var item in cotizadorDetalles)
                        {
                            item.id_cotizador = cotizadorCabecera.id_cotizador;
                            db.DetalleCotizador.Add(item);
                            db.SaveChanges();
                        }
                    }
                    else
                    {
                        //Limpiar si viene vacio
                        var detallesCotizadorLimpiar = db.DetalleCotizador.Where(s => s.id_cotizador == cotizadorCabecera.id_cotizador).ToList();
                        foreach (var item in detallesCotizadorLimpiar)
                        {
                            db.DetalleCotizador.Remove(item);
                            db.SaveChanges();
                        }
                    }

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

        //Eliminación Lógica
        public static RespuestaTransaccion EliminarCotizador(int id)
        {
            try
            {
                var Cotizador = db.Cotizador.Find(id);

                Cotizador.estado_cotizador = false;

                db.Entry(Cotizador).State = EntityState.Modified;
                db.SaveChanges();

                return new RespuestaTransaccion { Estado = true, Respuesta = Mensajes.MensajeTransaccionExitosa };
            }
            catch (Exception ex)
            {
                return new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeTransaccionFallida + " ;" + ex.Message.ToString() };
            }
        }

        // Obtiene toda la cotización (Cabecera y Detalle)
        public static List<CabeceraDetalleCotizador_Result> ConsultarCotizacion(int idCotizador)
        {
            try
            {
                return db.CabeceraDetalleCotizador(idCotizador).ToList();
            }
            catch (Exception)
            {
                throw;
            }
        }

        //Obtiene listado de los detalles de cotizaciones 
        public static List<ListadoDetalleCotizador_Result> ListarDetalleCotizador(int idCotizador)
        {
            try
            {
                var listadorDetalle = db.ListadoDetalleCotizador(idCotizador).ToList();
                return listadorDetalle;
            }
            catch (Exception)
            {
                throw;
            }
        }

        //Obtiene listado de cabeceras de cotizaciones
        public static List<ListadoCabeceraCotizador_Result> ListarCabecerasCotizacion()
        {
            try
            {
                var listadorDetalle = db.ListadoCabeceraCotizador().ToList();
                return listadorDetalle;
            }
            catch (Exception ex)
            {
                throw;
            }


        }

        //Obtiene listado de cabeceras de cotizaciones PDF
        public static List<ListadoCabeceraCotizadorPDF> ListarCabecerasCotizacionPDF()
        {
            try
            {
                var listadorDetalle = db.ListadoCabeceraCotizadorPDF().ToList();
                return listadorDetalle;
            }
            catch (Exception)
            {
                throw;
            }
        }

        //Consulta cabecera cotizador
        public static Cotizador ConsultarCotizadorCabecera(int id)
        {
            try
            {
                Cotizador cliente = db.Cotizador.Find(id);
                return cliente;
            }
            catch (Exception)
            {
                throw;
            }
        }

        public static bool ConsultarCotizadorCabeceraByCodigoCotizacion(int idCodigoCotizacion)
        {
            try
            {
                Cotizador cliente = db.Cotizador.FirstOrDefault(s => s.id_codigo_cotizacion == idCodigoCotizacion);
                if (cliente != null)
                {
                    return true;
                }
                else
                    return false;
            }
            catch (Exception ex)
            {
                return false;
            }
        }

        public static void CambiarEstadoCotizador(int id)
        {
            try
            {
                var cotizacion = db.Cotizador.Find(id);
                cotizacion.estado_cotizador = false;

                db.Entry(cotizacion).State = EntityState.Modified;
                db.SaveChanges();
            }
            catch (Exception)
            {
                throw;
            }
        }

        // Consultar Detalle Cotizador
        public static List<DetalleCotizador> ConsultarCotizadorDetalle(int idCabecera)
        {
            try
            {
                    var detalle = db.DetalleCotizador.Where(s => s.id_cotizador == idCabecera).ToList();
                    return detalle;
            }
            catch (Exception)
            {
                throw;
            }
        }

        //Enviar Correo
        public static RespuestaTransaccion EnviarMail(Cotizador cotizador)
        {
            using (var transaction = db.Database.BeginTransaction())
            {
                try
                {
                    ////enviar correo
                    //db.usp_envia_correo_usuario(id_usuario, numero);
                    //db.SaveChanges();

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

        public static RespuestaTransaccion EnviarCorreo(int id_cotizador, string cuerpo, int idUsuario, string usuarioclave)
        {
            using (var transaction = db.Database.BeginTransaction())
            {
                try
                {
                    //enviar correo
                    db.usp_guarda_envio_correo_cotizacion(id_cotizador, cuerpo, idUsuario, usuarioclave);
                    db.SaveChanges();
                     
                    transaction.Commit();
                    return new RespuestaTransaccion { Estado = true, Respuesta = Mensajes.MensajeTransaccionExitosa};
                }
                catch (Exception ex)
                {
                    transaction.Rollback();
                    return new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeTransaccionFallida + " ;" + ex.Message.ToString() };
                }
            }
        }

        public static IEnumerable<SelectListItem> ObtenerListadoVersionesByCodigoCotizacion(string codigo, string seleccionado = null)
        {
            List<SelectListItem> ListadoCatalogo = new List<SelectListItem>();
            try
            {
                int IDCodigoCotizacion = Convert.ToInt32(codigo);

                var padreCatalogo = db.Cotizador.FirstOrDefault(s => s.id_codigo_cotizacion == IDCodigoCotizacion);

                if (padreCatalogo == null)
                {
                    return new List<SelectListItem>();
                }

                ListadoCatalogo = db.Cotizador.Where(s => s.id_codigo_cotizacion == IDCodigoCotizacion && !s.estado_cotizador.Value).OrderBy(c => c.version).Select(c => new SelectListItem
                {
                    Text = c.version.ToString(),
                    Value = c.id_cotizador.ToString(),
                }).ToList();

                if (!string.IsNullOrEmpty(seleccionado))
                {
                    if (ListadoCatalogo.FirstOrDefault(s => s.Value == seleccionado.ToString()) != null)
                        ListadoCatalogo.FirstOrDefault(s => s.Value == seleccionado.ToString()).Selected = true;
                }

                return ListadoCatalogo;
            }
            catch (Exception ex)
            {
                return ListadoCatalogo;
            }

        }

        public static RespuestaTransaccion AprobarCotizacion(AprobacionCotizacion codigo)
        {
            try
            {
                codigo.fecha_aprobacion = DateTime.Now;
                var CodigoCotizacion = db.AprobacionCotizacion.Find(codigo.IDCotizacionAprobada);

                if (CodigoCotizacion != null)
                {
                    CodigoCotizacion.CotizacionID = codigo.CotizacionID; // Actualizacion de la nueva version
                    CodigoCotizacion.estatus_codigo = codigo.estatus_codigo; // Actualizacion de la nueva version
                    db.Entry(CodigoCotizacion).State = EntityState.Modified;
                    db.SaveChanges();
                }
                else {

                    db.AprobacionCotizacion.Add(codigo);
                    db.SaveChanges();
                } 

                //enviar notificaciones
                db.usp_guarda_envio_correo_notificaciones(3, codigo.CodigoCotizacionID, "", 1, "");
                db.SaveChanges();

                return new RespuestaTransaccion { Estado = true, Respuesta = Mensajes.MensajeTransaccionExitosa };
            }
            catch (Exception ex)
            {
                return new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeTransaccionFallida + " ;" + ex.Message.ToString() };
            }
        }

        public static AprobacionCotizacion ConsultarCodigoCotizacion(int id)
        {
            AprobacionCotizacion aprobacion = new AprobacionCotizacion {
                IDCotizacionAprobada = 0,
                CotizacionID = 0,
                CodigoCotizacionID = 0,
                estatus_codigo = 0,
                Observacion = string.Empty
            };

            try
            {
                var aprobacionCotizacion = db.AprobacionCotizacion.FirstOrDefault(s=> s.CodigoCotizacionID == id);

                if (aprobacionCotizacion != null)
                    return aprobacionCotizacion;
                else
                    return aprobacion;
            }
            catch (Exception ex)
            {
                return aprobacion;
            }
        }

        #region Metodos de cotizaciones o prefacturas de SAFI
        // Obtiene toda la cotización (Cabecera y Detalle)
        public static PrefacturaSAFIInfo ConsultarPrefacturaSAFI(int id)
        {
            PrefacturaSAFIInfo objeto = new PrefacturaSAFIInfo();
            try
            {
                objeto = db.ConsultarPrefacturaSAFI(id, null, null).FirstOrDefault();
                return objeto;
            }
            catch (Exception ex)
            {
                return objeto;
            }
        }

        public static List<ConsultarPrefacturaSAFIDetalle> ConsultarPrefacturaSAFIDetalle(int id)
        {
            List<ConsultarPrefacturaSAFIDetalle> objeto = new List<ConsultarPrefacturaSAFIDetalle>();
            try
            {
                objeto = db.ConsultarPrefacturaSAFIDetalle(id).ToList();
                return objeto;
            }
            catch (Exception ex)
            {
                return objeto;
            }
        }

        public static bool AprobacionInicialPrefactura(List<int> ids, int usuarioID)
        {
            using (var transaction = db.Database.BeginTransaction())
            {
                try
                {

                    using (var context = new GestionPPMEntities())
                    {
                        foreach (var id in ids)
                        {
                            var prefactura = context.SAFIGeneral.Find(id);

                            //Prefactura anulada
                            if (prefactura.numero_prefactura.Equals("0"))
                                return false;

                            prefactura.aprobacion_prefactura_ejecutivo = true;
                            prefactura.fecha_aprobacion_prefactura_ejecutivo = DateTime.Now;

                            prefactura.aprobacion_prefactura_ejecutivo_UsuarioID = usuarioID;

                            context.Entry(prefactura).State = EntityState.Modified;
                            context.SaveChanges();
                        }
                    }

                    transaction.Commit();
                    return true;
                }
                catch (Exception ex)
                {
                    transaction.Rollback();
                    return false;
                }
            }
        }

        public static bool AprobacionFinalPrefactura(List<int> ids, int usuarioID)
        {
            using (var transaction = db.Database.BeginTransaction())
            {
                try
                {

                    using (var context = new GestionPPMEntities())
                    {
                        foreach (var id in ids)
                        {
                            var prefactura = context.SAFIGeneral.Find(id);

                            if (prefactura.numero_factura != null)
                                return false;

                            prefactura.aprobacion_final = true;
                            prefactura.fecha_aprobacion_final = DateTime.Now;

                            prefactura.aprobacion_final_UsuarioID = usuarioID;

                            context.Entry(prefactura).State = EntityState.Modified;
                            context.SaveChanges();
                        }
                    }

                    transaction.Commit();
                    return true;
                }
                catch (Exception ex)
                {
                    transaction.Rollback();
                    return false;
                }
            }
        }

        public static RespuestaTransaccion ConsolidarPrefactura(string ids, int usuarioID)
        {
            using (var transaction = db.Database.BeginTransaction())
            {
                using (var context = new GestionPPMEntities())
                {
                    try
                    {
                        var numeroPrefacturas = !string.IsNullOrEmpty(ids) ? ids.Split(',').Select(int.Parse).ToList() : new List<int> { int.Parse(ids) };

                        string detalleConsolidacion = "PRESUPUESTO CONSOLIDADO ( {0} Presupuestos)";
                        detalleConsolidacion = string.Format(detalleConsolidacion, numeroPrefacturas.Count);

                        //string nombreClienteBusqueda = "DINERS CLUB DEL ECUADOR S.A.";

                        //var cliente = context.Cliente.Where(s => s.nombre_comercial_cliente.Equals(nombreClienteBusqueda)).FirstOrDefault();

                        List<ListadoPresupuestoPrefacturaSAFI> ListadoPrefactura = new List<ListadoPresupuestoPrefacturaSAFI>();

                        //Obtener el secuencial 
                        var secuencial = db.usp_g_codigo_documento("CodigoCotizacion").First().secuencial;

                        //Obtener resultado para construir prefactura consolidada
                        var resultadoConsolidacion = db.ResultadoConsolidacionPrefacturas(ids).FirstOrDefault();

                        //Obtener detalle de los Presuspuestos consolidados
                        List<ResultadoConsolidacionPrefacturasDetallado> listadoDetalleConciliacion = db.ResultadoConsolidacionPrefacturasDetallado(ids).ToList();

                        //obtener datos del cliente
                        var cliente = db.Cliente.Find(resultadoConsolidacion.id_cliente);
                         
                        if (resultadoConsolidacion.id_cliente == 0)
                            return new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeTransaccionFallida + string.Format(Mensajes.MensajeClienteConsolidacionInexistente, cliente.nombre_comercial_cliente) };

                        if (resultadoConsolidacion == null)
                            return new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeTransaccionFallida };

                        foreach(var item in listadoDetalleConciliacion)
                        {
                            ListadoPrefactura.Add(new ListadoPresupuestoPrefacturaSAFI
                            {
                                Secuencial = secuencial.Value,
                                Cliente = resultadoConsolidacion.nombre_comercial_cliente,
                                Identificacion = resultadoConsolidacion.ruc_ci_cliente,
                                Correos = resultadoConsolidacion.correos_facturacion,
                                Direccion = resultadoConsolidacion.direccion_cliente,
                                PorcentajeIva = 12,

                                Comentario1 = detalleConsolidacion,
                                Comentario2 = detalleConsolidacion,
                                Comentario3 = detalleConsolidacion,

                                CodigoProdcuto = item.codigo_producto,
                                NombreProducto = item.nombre_producto,
                                Bodega = item.CodigoBodega,
                                IdBodega = item.IdBodega,
                                Id_Producto = item.IdProducto,
                                IdFormaPago = Convert.ToInt32(item.id_forma_pago),
                                CodigoFormaPago = item.FormaPago,
                                IdCentroCosto = item.IdCentroCosto,
                                CodigoCentroCosto = item.CentroCosto,
                                Cantidad = item.Cantidad.Value,
                                PrecioUnitario = item.PrecioUnitario,
                                Subtotal = item.SubtotalGeneral,
                                Descuento = item.DescuentoGeneral,
                                Total = item.TotalPagGeneral,
                                Pago = 1,
                                Fecha = DateTime.Now.ToString("yyyy/MM/dd"),
                                
                            });
                        }
                         
                        var procesoPrefacturacion = PrefacturarSAFI(0, true, ListadoPrefactura);

                        if (procesoPrefacturacion.Estado)
                        {
                            var prefacturasConsolidadas = db.ConsolidarPrefactura(ids, usuarioID, procesoPrefacturacion.DocumentoSAFIID);

                            context.SaveChanges();

                            transaction.Commit();
                            return new RespuestaTransaccion { Estado = true, Respuesta = Mensajes.MensajeTransaccionExitosa };
                        }
                        else
                        {
                            transaction.Rollback();
                            return new RespuestaTransaccion { Estado = false, Respuesta = procesoPrefacturacion.Respuesta };
                        }
                    }
                    catch (Exception ex)
                    {
                        transaction.Rollback();
                        return new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeTransaccionFallida };
                    }
                }
            }
        }

        public static List<PrefacturaSAFIInfo> ListadoPrefacturaSAFI(int? clienteID = null, DateTime? fechaInicio = null, DateTime? fechaFin = null, int? ejecutivoID = null)
        {
            List<PrefacturaSAFIInfo> listado = new List<PrefacturaSAFIInfo>();
            try
            {
                listado = db.ListadoPrefacturasSAFI(clienteID, fechaInicio, fechaFin, ejecutivoID).ToList();
                return listado;
            }
            catch (Exception ex)
            {
                return listado;
            }
        }

        public static List<PrefacturaSAFIInfo> ListadoPrefacturasAprobadas(int usuario)
        {
            List<PrefacturaSAFIInfo> listado = new List<PrefacturaSAFIInfo>();
            try
            {
                listado = db.ListadoPrefacturasAprobadas(usuario).ToList();
                return listado;
            }
            catch (Exception ex)
            {
                return listado;
            }
        }
         
        public static List<PrefacturaSAFIInfo> ListadoCompletoPrefacturaSAFI()
        {
            List<PrefacturaSAFIInfo> listado = new List<PrefacturaSAFIInfo>();
            try
            {
                listado = db.ListadoCompletoDocumentosSAFI().ToList();
                return listado;
            }
            catch (Exception ex)
            {
                return listado;
            }
        }

        public static IEnumerable<SelectListItem> ListadoSeleccionPrefacturaSAFI(List<PrefacturaSAFIInfo> items, string seleccionado = null, bool nuevo = false, bool tipoCliente = false)
        {
            List<SelectListItem> ListadoCatalogo = new List<SelectListItem>();
            try
            {
                if (nuevo)
                {
                    var filtroTipo = tipoCliente ? items.Where(s => !s.UtilizadoEnActaCliente.Value).ToList() : items.Where(s => !s.UtilizadoEnActaContabilidad.Value).ToList();
                    ListadoCatalogo = filtroTipo.OrderBy(c => c.id_facturacion_safi).Select(c => new SelectListItem
                    {
                        Text = c.numero_prefactura + " - " + c.codigo_cotizacion,
                        Value = c.id_facturacion_safi.ToString(),
                        //Disabled = string.IsNullOrEmpty(seleccionado) ? false : c.UtilizadoEnActa.Value && c.id_facturacion_safi != int.Parse((seleccionado ?? "0")) ? true : false,
                        //Disabled = c.UtilizadoEnActa.Value && c.id_facturacion_safi != int.Parse((seleccionado ?? "0")) ? true : false,
                    }).ToList();

                    if (!string.IsNullOrEmpty(seleccionado))
                    {
                        if (ListadoCatalogo.FirstOrDefault(s => s.Value == seleccionado.ToString()) != null)
                            ListadoCatalogo.FirstOrDefault(s => s.Value == seleccionado.ToString()).Selected = true;
                    }
                }
                else
                {
                    ListadoCatalogo = items.OrderBy(c => c.id_facturacion_safi).Select(c => new SelectListItem
                    {
                        Text = c.numero_prefactura + " - " + c.codigo_cotizacion,
                        Value = c.id_facturacion_safi.ToString(),
                        //Disabled = string.IsNullOrEmpty(seleccionado) ? false : c.UtilizadoEnActa.Value && c.id_facturacion_safi != int.Parse((seleccionado ?? "0")) ? true : false,
                        //Disabled = c.UtilizadoEnActa.Value && c.id_facturacion_safi != int.Parse((seleccionado ?? "0")) ? true : false,
                    }).ToList();

                    if (!string.IsNullOrEmpty(seleccionado))
                    {
                        if (ListadoCatalogo.FirstOrDefault(s => s.Value == seleccionado.ToString()) != null)
                            ListadoCatalogo.FirstOrDefault(s => s.Value == seleccionado.ToString()).Selected = true;
                    }
                }

                return ListadoCatalogo;
            }
            catch (Exception ex)
            {
                return ListadoCatalogo;
            }

        }

        public static RespuestaTransaccion PrefacturarSAFI(int idCodigoCotizacion, bool flag = false, List<ListadoPresupuestoPrefacturaSAFI> listado = null)
        {
            int DocumentoSAFIID = 0;
            bool EstadoRespuesta = true;
            string MensajeErrorWs = string.Empty;

            try
            {
                //Obtener listado facturas para SaFi 
                var ListadoPrefacturaSAFI = !flag ? db.ListadoPresupuestoPrefacturaSAFI(idCodigoCotizacion).ToList() : listado;

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
                            var mensaje = "";
                            if (obj.mensaje.Any() && obj.mensaje != null)
                                mensaje = obj.mensaje;
                            estado = "ERROR";
                            tRespuesta = new Respuesta { mensaje = mensaje, codigoRetorno = obj.codigoRetorno.ToString(), estado = "ERROR", numeroDocumento = "" };
                            MensajeErrorWs = tRespuesta.mensaje;
                            EstadoRespuesta = false;
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
                                MensajeErrorWs = tRespuesta.mensaje;
                                EstadoRespuesta = false;
                            }
                            else
                            {
                                decimal subtotalGeneral = 0;
                                decimal totalGeneral = 0;
                                decimal descuentoGenreal = 0;
                                int cantidadGeneral = 0;

                                //Insertar el registro de Prefactura
                                SAFIGeneral prefacturaSAFI = new SAFIGeneral();

                                //insertar la cabecera
                                bool cabecera = false;

                                //obtener totale Generales
                                foreach (var totales in ListadoPrefacturaSAFI)
                                {
                                    cantidadGeneral += totales.Cantidad;
                                    subtotalGeneral += totales.Subtotal.Value;
                                    totalGeneral += totales.Total.Value;
                                    descuentoGenreal += totales.Descuento.Value;
                                }

                                foreach (var ele in ListadoPrefacturaSAFI)
                                {

                                    if (!cabecera)
                                    {
                                        cabecera = true;
                                        prefacturaSAFI.id_codigo_cotizacion = idCodigoCotizacion;
                                        prefacturaSAFI.detalle_cotizacion = ele.Comentario1;
                                        prefacturaSAFI.correos_facturacion = ele.Correos;
                                        prefacturaSAFI.numero_pago = ele.Pago;
                                        prefacturaSAFI.id_codigo_producto = ele.Id_Producto;
                                        prefacturaSAFI.id_forma_pago = ele.IdFormaPago;
                                        prefacturaSAFI.id_centro_costos = ele.IdCentroCosto;
                                        prefacturaSAFI.cantidad = cantidadGeneral;
                                        prefacturaSAFI.precio_unitario = ele.Subtotal;
                                        prefacturaSAFI.subtotal_pago = subtotalGeneral;
                                        prefacturaSAFI.iva_pago = (totalGeneral - subtotalGeneral);
                                        prefacturaSAFI.descuento_pago = descuentoGenreal;
                                        prefacturaSAFI.total_pago = totalGeneral;
                                        prefacturaSAFI.fecha_prefactura = DateTime.Now;
                                        prefacturaSAFI.numero_prefactura = obj.numeroDocumento;
                                        prefacturaSAFI.estado = true;
                                    }
                                    
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
                                    detalle.iva_pago = (totalGeneral - subtotalGeneral);
                                    detalle.descuento_pago = ele.Descuento;
                                    detalle.total_pago = ele.Total;
                                    detalle.estado = true;                                    

                                    db.SAFIGeneralDetalle.Add(detalle);
                                    db.SaveChanges();
                                }

                                DocumentoSAFIID = prefacturaSAFI.id_facturacion_safi;

                                estado = "OK";
                                tRespuesta = new Respuesta { mensaje = "PROCESO OK", codigoRetorno = obj.codigoRetorno.ToString(), estado = "OK", numeroDocumento = obj.numeroDocumento };

                            }
                        }
                    }
                    else
                    {
                        estado = "ERROR";
                        tRespuesta = new Respuesta { mensaje = "Servicios caídos, consulte con su proveedor.", codigoRetorno = "400", estado = "ERROR", numeroDocumento = "" };
                        MensajeErrorWs = "Servicios caídos, consulte con su proveedor.";
                        EstadoRespuesta = false;
                    }
                }
                return new RespuestaTransaccion { Estado = EstadoRespuesta, Respuesta = EstadoRespuesta ? Mensajes.MensajeTransaccionExitosa : Mensajes.MensajeTransaccionFallida + MensajeErrorWs, DocumentoSAFIID = DocumentoSAFIID };
            }
            catch (Exception ex)
            {
                return new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeTransaccionFallida + " " + ex.Message.ToString(), DocumentoSAFIID = DocumentoSAFIID };
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

        //Consolidar varios Presupuetos en un solo Presupuestos
        public static List<Wrapper> ProcesaCotizacionesSAFI(List<ListadoPresupuestoPrefacturaSAFI> ListadoPrefactura, string secuencial)
        {
            var facturas = new List<Wrapper>();
            decimal subtotalGeneral = 0;
            decimal totalGeneral = 0;
            decimal descuentoGenreal = 0 ;
            bool finaliza = false;
            try
            {
                if (ListadoPrefactura.Any())
                {
                    //obtener totale Generales
                    foreach (var ele in ListadoPrefactura)
                    {
                        subtotalGeneral += ele.Subtotal.Value;
                        totalGeneral += ele.Total.Value;
                        descuentoGenreal += ele.Descuento.Value;
                    }
                    
                    //generar la trama de cotizacion
                    foreach (var ele in ListadoPrefactura)
                    {
                        if (!finaliza)
                        {
                            //cambio estado para que solo genere una prefactura con n detalles
                            finaliza = true;

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

                            //Datos Factura
                            var fact = new Factura
                            {
                                Secuencial = secuencial,
                                Observacion = "",
                                SubTotal = Convert.ToDecimal(subtotalGeneral),
                                Total = Convert.ToDecimal(totalGeneral),
                                Descuento = Convert.ToDecimal(descuentoGenreal),
                                PorcentajeIva = Convert.ToDecimal(ele.PorcentajeIva.ToString()),
                            };

                            //Datos del detalle de la factura
                            var listDet = new List<DetalleFactura>();
                            foreach (var detalle in ListadoPrefactura)
                            {
                                var detItem = new DetalleFactura
                                {
                                    Cantidad = Convert.ToInt32(detalle.Cantidad.ToString()),
                                    Detalle = "",
                                    Valor = Convert.ToDecimal(detalle.Subtotal.ToString()), //subtotal+iva
                                    SubTotal = Convert.ToDecimal(detalle.Subtotal.ToString()),// !string.IsNullOrEmpty(item.SubTotal) ? decimal.Parse(item.SubTotal) : 0,
                                    Descuento = Convert.ToDecimal(detalle.Descuento.ToString()),
                                    Total = Convert.ToDecimal(detalle.Total.ToString()),
                                    CodigoCategoria = ConfigurationManager.AppSettings.Get("codigoCategoria"),
                                    CodigoProducto = detalle.CodigoProdcuto,
                                    NombreProducto = detalle.NombreProducto,
                                    RUCProveedor = "",
                                    Proveedor = "",
                                    CostoUnitario = Convert.ToDecimal(detalle.PrecioUnitario.ToString()),
                                    FechaVenta = Convert.ToDateTime(detalle.Fecha.ToString()),
                                    PorcentajeIva = Convert.ToDecimal(detalle.PorcentajeIva.ToString()),
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
            }
            catch (Exception ex)
            {
                logger.Error(ex, "Stopped program because of exception");
            }

            return facturas;
        }

        public static List<WsDocumentoSAFI> ListadoDocumentosSAFI(string url, string puntoEmision, string asc = "true")
        {
            List<WsDocumentoSAFI> listado = new List<WsDocumentoSAFI>();
            try
            {
                var conexion = new RestClient(url);
                var request = new RestRequest(Method.GET);
                request.AddHeader("puntoEmision", puntoEmision);
                request.AddHeader("sort", asc);

                var response = conexion.Execute(request);

                if (response != null)
                {
                    listado = JsonConvert.DeserializeObject<List<WsDocumentoSAFI>>(response.Content);
                }

                return listado;
            }
            catch (Exception ex)
            {
                logger.Error(ex, Mensajes.MensajeExcepcionNLOG);
                return listado;
            }
        }

        public static List<WsDocumentoSAFI> ConsultarDocumentoSAFI(string url, string numeroDocumento)
        {
            List<WsDocumentoSAFI> listado = new List<WsDocumentoSAFI>();
            try
            {
                var conexion = new RestClient(url);
                var request = new RestRequest(Method.GET);
                request.AddHeader("numeroDocumento", numeroDocumento);

                var response = conexion.Execute(request);

                if (response != null)
                {
                    listado = JsonConvert.DeserializeObject<List<WsDocumentoSAFI>>(response.Content);
                }

                return listado;
            }
            catch (Exception ex)
            {
                logger.Error(ex, Mensajes.MensajeExcepcionNLOG);
                return listado;
            }
        }

        public static List<WsDocumentoSAFI> ListadoRetencionesSAFI(string url, int anioInicial, string puntoEmision)
        {
            List<WsDocumentoSAFI> listado = new List<WsDocumentoSAFI>();
            try
            {
                var conexion = new RestClient(url);
                var request = new RestRequest(Method.GET);
                request.AddHeader("anioInicial", anioInicial.ToString());
                request.AddHeader("puntoEmision", puntoEmision);

                var response = conexion.Execute(request);

                if (response != null)
                {
                    listado = JsonConvert.DeserializeObject<List<WsDocumentoSAFI>>(response.Content);
                }

                return listado;
            }
            catch (Exception ex)
            {
                logger.Error(ex, Mensajes.MensajeExcepcionNLOG);
                return listado;
            }
        }

        public static List<WsDocumentoSAFI> ConsultarRetencionSAFI(string url, string numeroDocumento)
        {
            List<WsDocumentoSAFI> listado = new List<WsDocumentoSAFI>();
            try
            {
                var conexion = new RestClient(url);
                var request = new RestRequest(Method.GET);
                request.AddHeader("numeroDocumento", numeroDocumento);

                var response = conexion.Execute(request);

                if (response != null)
                {
                    listado = JsonConvert.DeserializeObject<List<WsDocumentoSAFI>>(response.Content);
                }

                return listado;
            }
            catch (Exception ex)
            {
                logger.Error(ex, Mensajes.MensajeExcepcionNLOG);
                return listado;
            }
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

        #endregion

        public static List<PrefacturaSAFIInfo> ListadoGestionPrefacturaSAFI2(long? pagina = null, string textoBusqueda = null, string filtro = null)
        {
            List<PrefacturaSAFIInfo> listado = new List<PrefacturaSAFIInfo>();
            try
            {
                listado = db.ListadoResumenPrefacturas(pagina, textoBusqueda, filtro).ToList();//db.ListadoPrefacturasSAFI2(null, null, null, null, pagina, textoBusqueda).ToList();//.Where(s => !s.aprobacion_prefactura_ejecutivo.Value).ToList();
                return listado;
            }
            catch (Exception ex)
            {
                return listado;
            }
        }

        public static int ObtenerTotalRegistrosListadoPrefacturaSAFI()
        {
            int total = 0;
            try
            {
                total = db.Database.SqlQuery<int>("SELECT [dbo].[ObtenerTotalRegistrosListadoPrefacturasSAFI]()").Single();
                return total;
            }
            catch (Exception ex)
            {
                return total;
            }
        }
         

    }
}
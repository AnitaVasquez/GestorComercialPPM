using GestionPPM.Entidades.Modelo;
using GestionPPM.Repositorios;
using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.Linq;
using System.Threading.Tasks;
using System.Web;

namespace GestionPPM.Entidades.Metodos
{
    public static class SolicitudClienteEntity
    {
        private static readonly GestionPPMEntities db = new GestionPPMEntities();

        public static RespuestaTransaccion CrearSolicitud(SolicitudCliente solicitud, int codigoUsuario, List<CamposPersonalizadosParcial> camposPersonalizados, List<UrlExternoParcial> urlExterno, List<AdjutoSolicitudParcial> adjuntos, List<UrlExternoParcial> urlSoporte, List<int> portales)
        {
            using (var transaction = db.Database.BeginTransaction())
            {
                try
                {
                    //Registrar la solicitud

                    solicitud.estado = true;
                    solicitud.fecha_hora_solicitud = DateTime.Now;
                    solicitud.id_solicitante = codigoUsuario;

                    //Datos para creado el estado de la solicitud 
                    var Tipo = CatalogoEntity.ObtenerListadoCatalogosByCodigo("SLT-01");
                    var estadoSolicitud = Tipo.Where(t => t.Text == "CREADO").FirstOrDefault().Value;
                    solicitud.id_estado_solicitud = Convert.ToInt32(estadoSolicitud);

                    db.SolicitudCliente.Add(solicitud);
                    db.SaveChanges();
                                       
                    //Validar que exista campos personalizados
                    if (camposPersonalizados.Count > 0 && camposPersonalizados != null)
                    {
                        foreach (var item in camposPersonalizados)
                        {
                            Solicitud_Campos_Personalizados campo = new Solicitud_Campos_Personalizados
                            {
                                id_solicitud = solicitud.id_solicitud,
                                nombre_campo = item.NombrePers, 
                            };

                            db.Solicitud_Campos_Personalizados.Add(campo);
                            db.SaveChanges();
                        }
                    }

                    //Validar que exista urls
                    if (urlExterno.Count > 0 && urlExterno != null)
                    {
                        foreach (var item in urlExterno)
                        {
                            Solicitud_Url_Externo urls = new Solicitud_Url_Externo
                            {
                                id_solicitud = solicitud.id_solicitud,
                                tipo = item.tipoUrl,
                                detalle = item.detalleUrl,
                                NumeroTelefono = item.numeroTelefono,
                                url = item.url
                            };

                            db.Solicitud_Url_Externo.Add(urls);
                            db.SaveChanges();
                        }
                    }

                    //Insertar Adjuntos
                    if (adjuntos.Count > 0 && adjuntos != null)
                    {
                        foreach (var item in adjuntos)
                        {
                            Solicitud_Adjuntos adjuntoArchivo = new Solicitud_Adjuntos
                            {
                                id_solicitud = solicitud.id_solicitud,
                                tipo = item.tipoAdjunto,
                                nombre_adjunto = item.nombreAdjunto
                            };

                            db.Solicitud_Adjuntos.Add(adjuntoArchivo);
                            db.SaveChanges();
                        }
                    }

                    //Validar que exista urls Externo Soporte
                    if (urlSoporte.Count > 0 && urlSoporte != null)
                    {
                        foreach (var item in urlSoporte)
                        {
                            Solicitud_Url_Externo urlsSoporte = new Solicitud_Url_Externo
                            {
                                id_solicitud = solicitud.id_solicitud,
                                tipo = item.tipoUrl,
                                detalle = item.detalleUrl,
                                url = item.url
                            };

                            db.Solicitud_Url_Externo.Add(urlsSoporte);
                            db.SaveChanges();
                        }
                    }

                    //Agregar detalle de Portales
                    if (portales.Count > 0 && portales != null)
                    {
                        foreach (var item in portales)
                        {
                            Solicitud_Portales idsPortales = new Solicitud_Portales
                            {
                                id_solicitud = solicitud.id_solicitud,
                                id_marca = item
                            };

                            db.Solicitud_Portales.Add(idsPortales);
                            db.SaveChanges();
                        }
                    } 

                    //Enviar el correo si tiene asignacion automatica
                    db.AsignacionSolicitudesAutomatico(solicitud.id_solicitud);
                    db.SaveChanges();

                    //enviar notificaciones
                    db.usp_guarda_envio_correo_notificaciones(1, solicitud.id_solicitud, "", 1, "" );
                    db.SaveChanges();

                    transaction.Commit();
                    return new RespuestaTransaccion { Estado = true, Respuesta = Mensajes.MensajeTransaccionExitosa, SolicitudID = solicitud.id_solicitud };
                }
                catch (Exception ex)
                {
                    transaction.Rollback();
                    return new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeTransaccionFallida + " ;" + ex.Message.ToString() };
                }
            }
        }

        public static RespuestaTransaccion ActualizarSolicitud(SolicitudCliente solicitud, List<UrlExternoParcial> urlSoporte)
        {
            using (var transaction = db.Database.BeginTransaction())
            {
                try
                {
                    // Por si queda el Attach de la entidad y no deja actualizar
                    var local = db.SolicitudCliente.FirstOrDefault(f => f.id_solicitud == solicitud.id_solicitud);
                    if (local != null)
                    {
                        db.Entry(local).State = EntityState.Detached;
                    }

                    //Obtener los datos de la Solicitud
                    SolicitudCliente solicitudNueva = db.SolicitudCliente.Find(solicitud.id_solicitud);
                    solicitudNueva.id_tipo = solicitud.id_tipo;
                    solicitudNueva.id_subtipo = solicitud.id_subtipo;
                    solicitudNueva.id_marca = solicitud.id_marca;
                    solicitudNueva.op = solicitud.op;
                    solicitudNueva.mkt = solicitud.mkt;
                    solicitudNueva.cantidad = solicitud.cantidad;

                    db.Entry(solicitudNueva).State = EntityState.Modified;
                    db.SaveChanges();
                     
                    //Validar que exista urls Externo Soporte
                    if (urlSoporte != null)
                    {
                        if(urlSoporte.Count > 0)
                        {
                            // ACTUALIZAR SUBLINEAS DE NEGOCIO CODIGO DE COTIZACION
                            var codigoUrlSoporteAnterior = db.Solicitud_Url_Externo.Where(s => s.id_solicitud == solicitud.id_solicitud).ToList();
                            foreach (var item in codigoUrlSoporteAnterior)
                            {
                                db.Solicitud_Url_Externo.Remove(item);
                                db.SaveChanges();
                            }

                            foreach (var item in urlSoporte)
                            {
                                Solicitud_Url_Externo urlsSoporte = new Solicitud_Url_Externo
                                {
                                    id_solicitud = solicitud.id_solicitud,
                                    tipo = "SOPORTE",
                                    detalle = item.url,
                                    NumeroTelefono = item.numeroTelefono,
                                    url = item.url
                                };

                                db.Solicitud_Url_Externo.Add(urlsSoporte);
                                db.SaveChanges();
                            }
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
         
        public static RespuestaTransaccion CrearCodigoCotizacionCargaMasiva(List<CodigoCotizacionExcel> codigos)
        {
            using (var transaction = db.Database.BeginTransaction())
            {

                try
                {
                    var user = HttpContext.Current.Session["usuario"];
                    var usuarioID = Convert.ToInt16(user);
                    var datosUsuario = UsuarioEntity.ConsultarUsuario(usuarioID);
                    var codigoUsuario = datosUsuario.secu_usua;

                    foreach (var item in codigos)
                    {

                        //Validacion de Mínimo un contacto y una Sublínea de Negocio
                        var sublineas = item.SublineasNegocio;
                        var contactos = item.Contactos;

                        if(!sublineas.Any() || !contactos.Any())
                            return new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeValidacionCodigoCotizacionCargaMasiva };

                        var generacionCodigo = db.GenerarCodigoCotizacion(codigoUsuario).FirstOrDefault();

                        if (generacionCodigo != null)
                        {
                            item.Estado = true;
                            item.codigo_cotizacion = generacionCodigo.CodigoCotizacionGenerado;
                            item.id_empresa = 1;
                        }
                        else
                            return new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeCodigoCotizacionError };

                        var registro = new CodigoCotizacion
                        {
                            //id_codigo_cotizacion = item,
                            codigo_cotizacion = item.codigo_cotizacion,
                            fecha_cotizacion = item.fecha_cotizacion,
                            responsable = item.responsable,
                            estatus_codigo = item.estatus_codigo,
                            id_cliente = item.id_cliente,
                            ejecutivo = item.ejecutivo,
                            tipo_requerido = item.tipo_requerido,
                            tipo_intermediario = item.tipo_intermediario,
                            tipo_proyecto = item.tipo_proyecto,
                            dimension_proyecto = item.dimension_proyecto,
                            aplica_contrato = item.aplica_contrato,
                            forma_pago = item.forma_pago,
                            forma_pago_1 = item.forma_pago_1,
                            forma_pago_2 = item.forma_pago_2,
                            forma_pago_3 = item.forma_pago_3,
                            forma_pago_4 = item.forma_pago_4,
                            etapa_cliente = item.etapa_cliente,
                            etapa_general = item.etapa_general,
                            estatus_detallado = item.estatus_detallado,
                            estatus_general = item.estatus_general,
                            tipo_producto_PtoP = item.tipo_producto_PtoP,
                            tipo_plan = item.tipo_plan,
                            tipo_tarifa = item.tipo_tarifa,
                            tipo_migracion = item.tipo_migracion,
                            tipo_etapa_PtoP = item.tipo_etapa_PtoP,
                            tipo_subsidio = item.tipo_subsidio,
                            nombre_proyecto = item.nombre_proyecto,
                            descripcion_proyecto = item.descripcion_proyecto,
                            tipo_fee = item.tipo_fee,
                            creacion_safi = item.creacion_safi,
                            facturable = item.facturable,
                            area_departamento_usuario = item.area_departamento_usuario,
                            pais = item.pais,
                            ciudad = item.ciudad,
                            direccion = item.direccion,
                            id_empresa = item.id_empresa,
                            estado_cliente = item.estado_cliente,
                            tipo_cliente = item.tipo_cliente,
                            tipo_zoho = item.tipo_zoho,
                            referido = item.referido,
                            Estado = item.Estado,
                        };

                        db.CodigoCotizacion.Add(registro);
                        db.SaveChanges();
                        
                        foreach (var obj in sublineas)
                        {
                            //Verificar si el contacto es de facturación
                            if (!CatalogoEntity.VerificarSublineaNegocio(obj.CodigoCatalogoSublineaNegocio))
                                return new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeValidacionCodigoCotizacionSublineaNegocioCargaMasiva + " Sublínea de Negocio: " + obj.CodigoCatalogoSublineaNegocio + " ."};

                            SublineaNegocioCodigoCotizacion sublinea = new SublineaNegocioCodigoCotizacion
                            {
                                CodigoCatalogoSublineaNegocio = Convert.ToInt32(obj.CodigoCatalogoSublineaNegocio),
                                IdCodigoCotizacion = registro.id_codigo_cotizacion,
                                Valor = obj.Valor,
                                Estado = true,
                            };

                            db.SublineaNegocioCodigoCotizacion.Add(sublinea);
                            db.SaveChanges();
                        }
                        
                        foreach (var obj in contactos)
                        {
                            //Verificar si el contacto es de facturación
                            if(!ContactoClienteEntity.VerificarContactoFacturacion(obj.idContacto.Value))
                                return new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeValidacionCodigoCotizacionContactoFacturacionCargaMasiva + ". Contacto: " + obj.idContacto  };

                            ContactosCodigoCotizacion contacto = new ContactosCodigoCotizacion
                            {
                                idCodigoCotizacion = registro.id_codigo_cotizacion,
                                idContacto = obj.idContacto
                            };

                            db.ContactosCodigoCotizacion.Add(contacto);
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
         
        public static RespuestaTransaccion ActualizarStatusCodigoCotizacion(CodigoCotizacion codigo)
        {
            try
            {
                var CodigoCotizacion = db.CodigoCotizacion.Find(codigo.id_codigo_cotizacion);

                CodigoCotizacion.estatus_codigo = codigo.estatus_codigo;

                db.Entry(CodigoCotizacion).State = EntityState.Modified;
                db.SaveChanges();

                return new RespuestaTransaccion { Estado = true, Respuesta = Mensajes.MensajeTransaccionExitosa };
            }
            catch (Exception ex)
            {
                return new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeTransaccionFallida + " ;" + ex.Message.ToString() };
            }
        }

        public static RespuestaTransaccion ActualizarEstadoSolicitud(SolicitudCliente solicitud)
        {
            try
            {
                var solicitudActual = db.SolicitudCliente.Find(solicitud.id_solicitud);

                solicitudActual.id_estado_solicitud = solicitud.id_estado_solicitud;

                db.Entry(solicitudActual).State = EntityState.Modified;
                db.SaveChanges();

                //enviar notificaciones
                db.usp_guarda_envio_correo_notificaciones(4, solicitud.id_solicitud, "", 1, "");
                db.SaveChanges();

                return new RespuestaTransaccion { Estado = true, Respuesta = Mensajes.MensajeTransaccionExitosa };
            }
            catch (Exception ex)
            {
                return new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeTransaccionFallida + " ;" + ex.Message.ToString() };
            }
        }

        public static RespuestaTransaccion ActualizarANSClienteSolicitud(SolicitudCliente solicitud)
        {
            try
            {
                var solicitudActual = db.SolicitudCliente.Find(solicitud.id_solicitud);

                solicitudActual.id_ans_sla = solicitud.id_ans_sla;

                db.Entry(solicitudActual).State = EntityState.Modified;
                db.SaveChanges();

                //enviar notificaciones
                db.usp_guarda_envio_correo_notificaciones(8, solicitud.id_solicitud, "", 1, "");
                db.SaveChanges();

                return new RespuestaTransaccion { Estado = true, Respuesta = Mensajes.MensajeTransaccionExitosa };
            }
            catch (Exception ex)
            {
                return new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeTransaccionFallida + " ;" + ex.Message.ToString() };
            }
        }

        public static RespuestaTransaccion AgregarAdjuntosMantenimiento(int? idSolicitud, string nombreArchivo)
        {
            using (var transaction = db.Database.BeginTransaction())
            {
                try
                {
                    //Agregar los Adjuntos al Listado de adjuntos  
                    
                    Solicitud_Adjuntos adjuntoArchivo = new Solicitud_Adjuntos
                    {
                        id_solicitud = idSolicitud,
                        tipo = "Adjunto Soporte",
                        nombre_adjunto = nombreArchivo
                    };

                    db.Solicitud_Adjuntos.Add(adjuntoArchivo);
                    db.SaveChanges();
                         
                    transaction.Commit();
                    return new RespuestaTransaccion { Estado = true, Respuesta = Mensajes.MensajeTransaccionExitosa, SolicitudID = idSolicitud };
                }
                catch (Exception ex)
                {
                    transaction.Rollback();
                    return new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeTransaccionFallida + " ;" + ex.Message.ToString() };
                }
            }
        }


        //Eliminación Lógica
        public static RespuestaTransaccion EliminarCodigoCotizacion(int id)
        {
            try
            {
                var CodigoCotizacion = db.CodigoCotizacion.Find(id);

                CodigoCotizacion.Estado = false;

                db.Entry(CodigoCotizacion).State = EntityState.Modified;
                db.SaveChanges();

                return new RespuestaTransaccion { Estado = true, Respuesta = Mensajes.MensajeTransaccionExitosa };
            }
            catch (Exception ex)
            {
                return new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeTransaccionFallida + " ;" + ex.Message.ToString() };
            }
        }

        public static List<CodigoCotizacionInfo> ListarCodigoCotizacion()
        {
            try
            {
                return db.ListadoCodigoCotizacion().ToList();
            }
            catch (Exception)
            {
                throw;
            }
        }

        public static List<SublineaNegocioCodigoCotizacionInfo> ListarSublineaCodigoCotizacion()
        {
            try
            {
                return db.ListadoSublineaNegocioCodigoCotizacion().ToList();
            }
            catch (Exception)
            {
                throw;
            }
        }

        public static SolicitudCliente ConsultarSolicitudCliente(int id)
        {
            try
            {
                SolicitudCliente solicitud = db.SolicitudCliente.Find(id);
                return solicitud;
            }
            catch (Exception)
            {
                throw;
            }
        }

        public static bool VerificarExistenciaCodigoCotizacion(string codigoCotizacion)
        {
            try
            {
                CodigoCotizacion codigo = db.CodigoCotizacion.FirstOrDefault(s => s.codigo_cotizacion == codigoCotizacion);
                if (codigo == null)
                    return true;
                else
                    return false;
            }
            catch (Exception)
            {
                return false;
            }
        }

        public static CodigoCotizacionInfo ConsultarCodigoCotizacionCompleto(int id)
        {
            try
            {
                CodigoCotizacionInfo codigoCotizacion = db.ListadoCodigoCotizacion().FirstOrDefault(s => s.id_codigo_cotizacion == id);
                return codigoCotizacion;
            }
            catch (Exception)
            {
                throw;
            }
        }

        public static bool ValidarPagosCodigoCotizacion(CodigoCotizacion codigoCotizacion)
        {
            bool flag = true;

            if (!codigoCotizacion.forma_pago.Value)
            {
                codigoCotizacion.forma_pago_1 = 100;
                codigoCotizacion.forma_pago_2 = 0;
                codigoCotizacion.forma_pago_3 = 0;
                codigoCotizacion.forma_pago_4 = 0;
            }

            var totalPagos = codigoCotizacion.forma_pago_1 + codigoCotizacion.forma_pago_2 + codigoCotizacion.forma_pago_3 + codigoCotizacion.forma_pago_4;

            if (totalPagos != 100)
                flag = false;

            return flag;
        }

        public static bool ValidarPagosCodigoCotizacion2(CodigoCotizacion codigoCotizacion)
        {
            bool flag = true;

            var totalPagos = codigoCotizacion.forma_pago_1 + codigoCotizacion.forma_pago_2 + codigoCotizacion.forma_pago_3 + codigoCotizacion.forma_pago_4;

            if (totalPagos != 100)
                flag = false;

            return flag;
        }

        public static bool ValidarPagosCodigoCotizacion3(CodigoCotizacion codigoCotizacion)
        {
            bool flag = true;

            if (!codigoCotizacion.forma_pago.Value)
            {
                if (codigoCotizacion.forma_pago_1 == 100 && codigoCotizacion.forma_pago_2 == 0 && codigoCotizacion.forma_pago_3 == 0 && codigoCotizacion.forma_pago_4 == 0)
                    return true;
                else
                    return false;
            }
            else {
                var totalPagos = codigoCotizacion.forma_pago_1 + codigoCotizacion.forma_pago_2 + codigoCotizacion.forma_pago_3 + codigoCotizacion.forma_pago_4;

                if (totalPagos != 100)
                    flag = false;

            }

            return flag;
        }

        public static bool ValidarPagos1(CodigoCotizacion codigoCotizacion)
        {
            bool flag = true;

            if (codigoCotizacion.forma_pago_1 == 0)
            {
                flag = false;
            }

            return flag;
        }

        public static List<SublineaNegocioCodigoCotizacionInfo> ListarSublineaNegocioCodigoCotizacion(int idCodigoCotizacion)
        {
            try
            {
                var listado = db.ListadoSublineaNegocioCodigoCotizacion().Where(s => s.IdCodigoCotizacion == idCodigoCotizacion).ToList();
                return listado;
            }
            catch (Exception)
            {

                throw;
            }

        }

    }
}
using GestionPPM.Entidades.Modelo;
using GestionPPM.Repositorios;
using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.Linq;
using System.Web;

namespace GestionPPM.Entidades.Metodos
{
    public class SolicitudClienteExternoEntity
    {
        private static readonly GestionPPMEntities db = new GestionPPMEntities();

        public static List<SolicitudClienteExternoInfo> ListadoSolicitudClienteExterno()
        {
            List<SolicitudClienteExternoInfo> listado = new List<SolicitudClienteExternoInfo>();
            try
            {
                listado = db.ListadoSolicitudClienteExterno().ToList();
                return listado;
            }
            catch (Exception ex)
            {
                return listado;
            }
        }

        public static SolicitudClienteExternoInfo ConsultarSolicitudClienteExterno(int id)
        {
            SolicitudClienteExternoInfo objeto = new SolicitudClienteExternoInfo();
            try
            {
                objeto = db.ConsultarSolicitudClienteExterno(id).FirstOrDefault();
                return objeto;
            }
            catch (Exception ex)
            {
                return objeto;
            }
        }

        public static TrackingAsignacionProyectoInfo ConsultarTrackingProyecto(int id)
        {
            TrackingAsignacionProyectoInfo objeto = new TrackingAsignacionProyectoInfo();
            try
            {
                objeto = db.ListadoDetalleTrackingAsignacionProyecto().FirstOrDefault(s=> s.id_codigo_cotizacion == id);

                //objeto = objeto ?? new TrackingAsignacionProyectoInfo();

                return objeto;
            }
            catch (Exception ex)
            {
                return objeto;
            }
        }

        public static List<HistorialDetalleAsignacionProyectosInfo> ListadoDetalleHistorialAsignacionProyectos(int id)
        {
            List<HistorialDetalleAsignacionProyectosInfo> listado = new List<HistorialDetalleAsignacionProyectosInfo>();
            try
            {
                listado = db.ListadoHistorialDetalleAsignacionProyectos().Where(s=> s.id_codigo_cotizacion == id).ToList();
                return listado;
            }
            catch (Exception ex)
            {
                return listado;
            }
        }
         
        public static List<ResumenSolicitudesSolicitanteInfo> ListadoResumenSolicitudesSolicitantes()
        {
            List<ResumenSolicitudesSolicitanteInfo> listado = new List<ResumenSolicitudesSolicitanteInfo>();
            try
            {
                listado = db.ListadoResumenSolicitudesSolicitante().ToList();
                return listado;
            }
            catch (Exception ex)
            {
                return listado;
            }
        }

        #region Comentarios Solicitud
        public static ComentariosSolicitudInfo ConsultarComentarioSolicitud(int? id = null)
        {
            ComentariosSolicitudInfo comentario = new ComentariosSolicitudInfo();
            try
            {
                comentario = db.ConsultarComentariosSolicitud(id).FirstOrDefault();
                return comentario;
            }
            catch (Exception ex)
            {
                return comentario;
            }
        }

        public static List<ComentariosSolicitudInfo> ListadoComentarioSolicitud(int? idSolicitud = null)
        {
            List<ComentariosSolicitudInfo> listado = new List<ComentariosSolicitudInfo>();
            try
            {
                listado = db.ListadoComentariosSolicitud(idSolicitud).ToList();
                return listado;
            }
            catch (Exception ex)
            {
                return listado;
            }
        }
        public static RespuestaTransaccion CrearComentarioSolicitud(ComentarioSolicitud comentario)
        {
            using (var transaction = db.Database.BeginTransaction())
            {
                try
                {
                    db.ComentarioSolicitud.Add(comentario);
                    db.SaveChanges();

                    transaction.Commit();

                    //Enviar correo con el comentario 
                    db.usp_guarda_envio_correo_notificaciones(7, Convert.ToInt32(comentario.IDComentarioSolicitud), comentario.Comentario, 1, "");
                    db.SaveChanges();

                    return new RespuestaTransaccion { Estado = true, Respuesta = Mensajes.MensajeTransaccionExitosa };
                }
                catch (Exception ex)
                {
                    transaction.Rollback();
                    return new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeTransaccionFallida + " ;" + ex.Message.ToString() };
                }
            }
        }

        public static RespuestaTransaccion EditarComentarioSolicitud(ComentarioSolicitud comentario)
        {
            using (var transaction = db.Database.BeginTransaction())
            {
                try
                {
                    // assume Entity base class have an Id property for all items
                    var entity = db.ComentarioSolicitud.Find(comentario.IDComentarioSolicitud);
                    if (entity == null)
                    {
                        return new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeTransaccionFallida };
                    }

                    db.Entry(entity).CurrentValues.SetValues(comentario);
                    db.SaveChanges();

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

        public static RespuestaTransaccion EliminarComentarioSolicitud(int id)
        {
            using (var transaction = db.Database.BeginTransaction())
            {
                try
                {
                    var comentario = db.ComentarioSolicitud.Find(id);
                    comentario.Estado = false;
                    db.Entry(comentario).State = EntityState.Modified;
                    db.SaveChanges();

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
        #endregion
    }
}
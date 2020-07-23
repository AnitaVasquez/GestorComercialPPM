using GestionPPM.Entidades.Modelo;
using GestionPPM.Repositorios;
using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.Linq;
using System.Web;

namespace GestionPPM.Entidades.Metodos
{
    public class SolicitudClienteInternoEntity
    {
        private static readonly GestionPPMEntities db = new GestionPPMEntities();

        public static List<SolicitudClienteInternoInfo> ListadoSolicitudClienteInterno()
        {
            List<SolicitudClienteInternoInfo> listado = new List<SolicitudClienteInternoInfo>();
            try
            {
                listado = db.ListadoSolicitudClienteInterno().ToList();
                return listado;
            }
            catch (Exception ex)
            {
                return listado;
            }
        }

        public static SolicitudClienteInternoInfo ConsultarSolicitudClienteInterno(int id)
        {
            SolicitudClienteInternoInfo objeto = new SolicitudClienteInternoInfo();
            try
            {
                objeto = db.ConsultarSolicitudClienteInterno(id).FirstOrDefault();
                return objeto;
            }
            catch (Exception ex)
            {
                return objeto;
            }
        }

        public static SolicitudClienteExternoInfo ConsultarSolicitudClienteByCodigoCotizacion(int idCodigoCotizacion)
        {
            SolicitudClienteExternoInfo objeto = new SolicitudClienteExternoInfo();
            try
            {
                objeto = db.ListadoSolicitudClienteExterno().Where(s=> s.id_codigo_cotizacion == idCodigoCotizacion).FirstOrDefault();
                return objeto;
            }
            catch (Exception ex)
            {
                return objeto;
            }
        }

        public static List<Usuario> ListarUsuariosInternosAsignados(List<int> idUsuariosInternos = null)
        {
            List<Usuario> listado = new List<Usuario>();
            try
            {
                if(idUsuariosInternos == null)
                    listado = db.Usuario.Where(s =>  s.validacion_correo.Value && s.tipo_usuario == 109).ToList();
                else
                    listado = db.Usuario.Where(s => s.validacion_correo.Value && s.tipo_usuario == 109 && idUsuariosInternos.Contains(s.id_usuario)).ToList();

                return listado;
            }
            catch (Exception ex)
            {
                return listado;
            }
        }

        public static List<int> ListarSolicitudesUsuariosBySolicitudID(int solicitudID)
        {
            List<int> listado = new List<int>();
            try
            {
                listado = db.Solicitud_Cliente_Usuario.Where(s => s.id_solicitud == solicitudID).Select(s=> s.id_usuario_interno.Value).ToList();
                return listado;
            }
            catch (Exception ex)
            {
                return listado;
            }
        }

        public static List<int> ListarSolicitudesUsuariosByUsuarioID(int usuarioID)
        {
            List<int> listado = new List<int>();
            try
            {
                listado = db.Solicitud_Cliente_Usuario.Where(s => s.id_usuario_interno == usuarioID).Select(s => s.id_solicitud.Value).ToList();
                return listado;
            }
            catch (Exception ex)
            {
                return listado;
            }
        }

        public static RespuestaTransaccion CrearActualizarAsginacionesSolicitudesUsuarios(int solicitudID, List<int> idsUsuarios)
        {
            using (var transaction = db.Database.BeginTransaction())
            {
                try
                {
                    var usuariosAnteriores = db.Solicitud_Cliente_Usuario.Where(s => s.id_solicitud == solicitudID).ToList();
                    foreach (var item in usuariosAnteriores)
                    {
                        db.Solicitud_Cliente_Usuario.Remove(item);
                        db.SaveChanges();
                    }

                    foreach (var item in idsUsuarios)
                    {
                        db.Solicitud_Cliente_Usuario.Add(new Solicitud_Cliente_Usuario
                        {
                            id_solicitud = solicitudID,
                            id_usuario_interno = item,
                        });
                        db.SaveChanges();
                    }

                    //enviar notificaciones
                    db.usp_guarda_envio_correo_notificaciones(5, solicitudID, "", 1, "");
                    db.SaveChanges();

                    db.usp_guarda_envio_correo_notificaciones(6, solicitudID, "", 1, "");
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

        public static List<int> ConsultarSolicitudesSinAsignacion()
        {
            List<int> ids = new List<int>();

            using (var transaction = db.Database.BeginTransaction())
            {
                try
                {

                    var solicitudes = db.ListadoSolicitudClienteInterno().ToList();

                    var idsSolicitudesAsignadas = db.Solicitud_Cliente_Usuario.Select(s=> s.id_solicitud).ToList();

                    solicitudes = solicitudes.Where(s => !idsSolicitudesAsignadas.Contains(s.id_solicitud)).ToList();

                    ids = solicitudes.Select(s => s.id_solicitud).ToList();

                    return ids;
                }
                catch (Exception ex)
                {
                    return ids;
                }
            }
        }

        public static RespuestaTransaccion ActualizarCodigoCotizacionSolicitudClienteInterno(int id, int codigoCotizacionID)
        {
            using (var transaction = db.Database.BeginTransaction())
            {
                try
                {
                    var solicitud = db.SolicitudCliente.Find(id);
                    solicitud.id_codigo_cotizacion = codigoCotizacionID;
                    db.Entry(solicitud).State = EntityState.Modified;
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

        public static RespuestaTransaccion CrearRespuestaComentarioSolicitud(RespuestaComentario respuestacomentario)
        {
            using (var transaction = db.Database.BeginTransaction())
            {
                try
                {
                    db.RespuestaComentario.Add(respuestacomentario);
                    db.SaveChanges();

                    transaction.Commit();

                    //Enviar correo con la respuesta a comentario 
                    db.usp_guarda_envio_correo_notificaciones(9, respuestacomentario.id_respuesta_comentario, respuestacomentario.Respuesta, 1, "");
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

    }
}
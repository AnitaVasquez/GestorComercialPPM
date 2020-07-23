using GestionPPM.Entidades.Modelo;
using GestionPPM.Repositorios;
using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.Linq;
using System.Threading.Tasks;
using System.Web;
using System.Web.Mvc;

namespace GestionPPM.Entidades.Metodos
{
    public static class AsignacionEntity
    {
        private static readonly GestionPPMEntities db = new GestionPPMEntities();

        public static RespuestaTransaccion CrearAsignacion(AsignacionSolicitudes asignacion, List<int> idUsuarios)
        {
            using (var transaction = db.Database.BeginTransaction())
            {
                try
                { 
                    asignacion.estado = true;
                    db.AsignacionSolicitudes.Add(asignacion);
                    db.SaveChanges();

                    var usuariosAnteriores = db.AsignacionSolicitudUsuario.Where(s => s.id_asignacion_solicitudes == asignacion.id_asignacion_solicitudes).ToList();
                    foreach (var item in usuariosAnteriores)
                    {
                        db.AsignacionSolicitudUsuario.Remove(item);
                        db.SaveChanges();
                    }
                      
                    foreach (var item in idUsuarios)
                    {
                        db.AsignacionSolicitudUsuario.Add(new AsignacionSolicitudUsuario
                        {
                            id_asignacion_solicitudes = asignacion.id_asignacion_solicitudes,
                            id_usuario = item ,
                        });
                        db.SaveChanges();
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

        public static RespuestaTransaccion ActualizarAsignacion(AsignacionSolicitudes asignacion, List<int> idUsuarios)
        {
            try
            {
                // Por si queda el Attach de la entidad y no deja actualizar
                var local = db.AsignacionSolicitudes.FirstOrDefault(f => f.id_asignacion_solicitudes == asignacion.id_asignacion_solicitudes);
                if (local != null)
                {
                    db.Entry(local).State = EntityState.Detached;
                }

                var usuariosAnteriores = db.AsignacionSolicitudUsuario.Where(s => s.id_asignacion_solicitudes == asignacion.id_asignacion_solicitudes).ToList();
                foreach (var item in usuariosAnteriores)
                {
                    db.AsignacionSolicitudUsuario.Remove(item);
                    db.SaveChanges();
                }
                 
                foreach (var item in idUsuarios)
                {
                    db.AsignacionSolicitudUsuario.Add(new AsignacionSolicitudUsuario
                    {
                        id_asignacion_solicitudes = asignacion.id_asignacion_solicitudes,
                        id_usuario = item,
                    });
                    db.SaveChanges();
                }
                 
                db.Entry(asignacion).State = EntityState.Modified;
                db.SaveChanges();

                return new RespuestaTransaccion { Estado = true, Respuesta = Mensajes.MensajeTransaccionExitosa };
            }
            catch (Exception ex)
            {
                return new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeTransaccionFallida + " ;" + ex.Message.ToString() };
            }
        }

        //Eliminación Lógica
        public static RespuestaTransaccion EliminarAsignacion(int id)
        {
            try
            {
                var asignacion = db.AsignacionSolicitudes.Find(id);

                if(asignacion.estado == true)
                {
                    asignacion.estado = false;
                }
                else
                {
                    asignacion.estado = true;
                }                

                db.Entry(asignacion).State = EntityState.Modified;
                db.SaveChanges();

                return new RespuestaTransaccion { Estado = true, Respuesta = Mensajes.MensajeTransaccionExitosa };
            }
            catch (Exception ex)
            {
                return new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeTransaccionFallida + " ;" + ex.Message.ToString() };
            }
        }

        public static List<ListadoAsignacionSolicitudes> ListarAsignacion()
        {
            try
            {
                return db.ListadoAsignacionSolicitudes().ToList();
            }
            catch (Exception e)
            {
                throw;
            }
        } 
         
        public static List<AsignacionSolicitudes> ObtenerListadoAsignados()
        {
            var ListadoAsignados = db.AsignacionSolicitudes.OrderBy(r => r.id_asignacion_solicitudes).ToList();
            return ListadoAsignados;
        }

        public static AsignacionSolicitudes ConsultarAsignacion(int id)
        {
            try
            {
                AsignacionSolicitudes asignacion = db.AsignacionSolicitudes.Find(id);
                return asignacion;
            }
            catch (Exception)
            {
                throw;
            }
        }

        public static List<int> ListadIdsUsuariosByAsignacion(int id)
        {
            try
            {
                var ListadoUsuarios = db.AsignacionSolicitudUsuario.Where(s => s.id_asignacion_solicitudes == id).Select(s=> s.id_usuario.Value).ToList();
                return ListadoUsuarios;
            }
            catch (Exception)
            {
                throw;
            }
        }

    }
}
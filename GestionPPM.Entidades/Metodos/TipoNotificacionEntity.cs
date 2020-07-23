using GestionPPM.Entidades.Modelo;
using GestionPPM.Repositorios;
using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.Linq; 

namespace GestionPPM.Entidades.Metodos
{
    public static class TipoNotificacionEntity
    {
        private static readonly GestionPPMEntities db = new GestionPPMEntities();

        public static RespuestaTransaccion CrearTipoNotificacion(TipoNotificacion tipoNotificacion)
        {
            try
            {
                tipoNotificacion.nombre_notificacion = tipoNotificacion.nombre_notificacion.ToUpper();
                tipoNotificacion.estado_notificacion = true;                
                db.TipoNotificacion.Add(tipoNotificacion);
                db.SaveChanges();

                return new RespuestaTransaccion { Estado = true, Respuesta = Mensajes.MensajeTransaccionExitosa };
            }
            catch (Exception ex)
            {
                return new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeTransaccionFallida + " ;" + ex.Message.ToString() };
            }
        }

        public static RespuestaTransaccion ActualizarTipoNotificacion(TipoNotificacion tipoNotificacion)
        {
            try
            {
                // Por si queda el Attach de la entidad y no deja actualizar
                var local = db.TipoNotificacion.FirstOrDefault(f => f.id_notificacion == tipoNotificacion.id_notificacion);
                if (local != null)
                {
                    db.Entry(local).State = EntityState.Detached;
                }

                tipoNotificacion.nombre_notificacion = tipoNotificacion.nombre_notificacion.ToUpper();
                db.Entry(tipoNotificacion).State = EntityState.Modified;
                db.SaveChanges();

                return new RespuestaTransaccion { Estado = true, Respuesta = Mensajes.MensajeTransaccionExitosa };
            }
            catch (Exception ex)
            {
                return new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeTransaccionFallida + " ;" + ex.Message.ToString() };
            }
        }

        //Eliminación Lógica
        public static RespuestaTransaccion EliminarTipoNotificacion(int id)
        {
            try
            {
                var TipoNotificacion = db.TipoNotificacion.Find(id);

                if(TipoNotificacion.estado_notificacion ==true)
                {
                    TipoNotificacion.estado_notificacion = false;
                }
                else
                {
                    TipoNotificacion.estado_notificacion = true;
                }
                

                db.Entry(TipoNotificacion).State = EntityState.Modified;
                db.SaveChanges();

                return new RespuestaTransaccion { Estado = true, Respuesta = Mensajes.MensajeTransaccionExitosa };
            }
            catch (Exception ex)
            {
                return new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeTransaccionFallida + " ;" + ex.Message.ToString() };
            }
        }

        public static List<ListadoTipoNotificaciones> ListarTipoNotificaciones()
        {
            try
            {
                return db.ListadoTipoNotificaciones().ToList();
            }
            catch (Exception ex)
            {
                throw;
            }
        }

        public static TipoNotificacion ConsultarNotificacion(int id)
        {
            try
            {
                TipoNotificacion tipoNotificacion = db.TipoNotificacion.Find(id);
                return tipoNotificacion;
            }
            catch (Exception)
            {
                throw;
            }
        }

    }
}
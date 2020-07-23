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
    public static class NotificacionEntity
    {

        public static List<NotificacionesCompletasInfo> ListarNotificaciones()
        {
            List<NotificacionesCompletasInfo> notificaciones = new List<NotificacionesCompletasInfo>();
            try
            {
                using (var context = new GestionPPMEntities())
                {
                    notificaciones = context.ListadoNotificacionesCompletas().ToList();
                    return notificaciones;
                }
            }
            catch (Exception ex)
            {
                return notificaciones;
            }
        }

        //Eliminación Lógica
        public static RespuestaTransaccion CancelarNotificacion(int id)
        {
            try
            {
                using (var context = new GestionPPMEntities())
                {
                    var notificacion = context.Notificaciones.Find(id);

                    // Se puede cancelar,  Siempre y cuando la notificacion no haya sido enviada
                    if (!notificacion.EstadoEnviadoNotificacion.Value)
                    {
                        notificacion.EstadoNotificacion = false;
                        notificacion.EstadoEjecucionNotificacion = false;
                        notificacion.EstadoEnviadoNotificacion = false;

                        notificacion.DetalleEstadoEjecucionNotificacion = "Envío cancelado. " + System.DateTime.Now;

                        context.Entry(notificacion).State = EntityState.Modified;
                        context.SaveChanges();

                        return new RespuestaTransaccion { Estado = true, Respuesta = Mensajes.MensajeTransaccionExitosa };
                    }
                    else {
                        return new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeValidacionNotificacionEnviada + " " +  Mensajes.MensajeTransaccionFallida };
                    }
                }
            }
            catch (Exception ex)
            {
                return new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeTransaccionFallida + " ;" + ex.Message.ToString() };
            }
        }

        //public static List<Notificaciones> GetNotificaciones()
        //{
        //    List<Notificaciones> notificaciones = new List<Notificaciones>();
        //    try
        //    {
        //        using (var context = new GestionPPMEntities())
        //        {
        //            notificaciones = context.Notificaciones.Where(s => s.EstadoNotificacion && !s.EstadoEnColaNotificacion).ToList();
        //            return notificaciones;
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        return notificaciones;
        //    }
        //}



        ////Actualizar estado de notificaciones en Cola
        //public static void ActualizarEjecucionNotificacion(Int64 id, string detalle, bool? estado, Int64 identificadorTarea)
        //{
        //    List<Notificaciones> notificaciones = new List<Notificaciones>();
        //    try
        //    {
        //        using (var context = new GestionPPMEntities())
        //        {
        //            context.ActualizarEjecucionNotificacion(id, detalle, estado, identificadorTarea);
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        throw;
        //    }
        //}

        ////Actualizar estado de notificaciones de Envío
        //public static void ActualizarEstadoEnvioNotificacion(Int64 id, bool estado)
        //{
        //    List<Notificaciones> notificaciones = new List<Notificaciones>();
        //    try
        //    {
        //        using (var context = new GestionPPMEntities())
        //        {
        //            context.ActualizarEstadoEnvioNotificacion(id, estado);
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        throw;
        //    }
        //}
    }
}
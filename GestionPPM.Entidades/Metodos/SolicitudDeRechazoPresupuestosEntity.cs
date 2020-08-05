using GestionPPM.Entidades.Modelo;
using GestionPPM.Repositorios;
using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.Linq;
using System.Web.Mvc;
using GestionPPM.Entidades.Modelo.SistemaContable;
using Newtonsoft.Json;
using NLog;
using RestSharp;
using System.Configuration;
using System.Net.Http;

namespace GestionPPM.Entidades.Metodos
{
    public class SolicitudDeRechazoPresupuestosEntity
    {
        private static readonly GestionPPMEntities db = new GestionPPMEntities();
        private static readonly V1791219058001_SAFI_3Entities erp = new V1791219058001_SAFI_3Entities();

        public static RespuestaTransaccion CrearSolicitudRechazoPresupuestoa(SolicitudesDeRechazoPresupuestos solicitudRechazo)
        {
            using (var transaction = db.Database.BeginTransaction())
            {
                try
                {
                    using (var context = new GestionPPMEntities())
                    {
                        //insertar las colicitud de reverso 
                        solicitudRechazo.fecha_solicitud_rechazo = DateTime.Now;
                        solicitudRechazo.estado = true;

                        context.SolicitudesDeRechazoPresupuestos.Add(solicitudRechazo);
                        context.SaveChanges();

                        //enviar notificaciones
                        context.usp_guarda_envio_correo_notificaciones(12, Convert.ToInt32(solicitudRechazo.id_facturacion_safi), "", 1, "");
                        context.SaveChanges();
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

        public static RespuestaTransaccion ActualizarRechazoPresupuesto(int id_facturacion_safi)
        {
            using (var transaction = db.Database.BeginTransaction())
            {
                try
                {
                    using (var context = new GestionPPMEntities())
                    {
                        //insertar las colicitud de reverso
                        SolicitudesDeRechazoPresupuestos reverso = new SolicitudesDeRechazoPresupuestos();
                        reverso = context.SolicitudesDeRechazoPresupuestos.ToList().Where(s => s.id_facturacion_safi == id_facturacion_safi).FirstOrDefault();

                        if (reverso.estado == true)
                        {
                            reverso.estado = false;

                            //cambiar el estado del reverso
                            context.Entry(reverso).State = EntityState.Modified;
                            context.SaveChanges();

                            //enviar notificaciones
                            context.usp_guarda_envio_correo_notificaciones(13, Convert.ToInt32(reverso.id_facturacion_safi), "", 1, "");
                            context.SaveChanges();
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
        public static RespuestaTransaccion AnularPresupuesto(int id)
        {
            using (var transaction = db.Database.BeginTransaction())
            {
                try
                {
                    using (var context = new GestionPPMEntities())
                    {
                        //insertar las colicitud de reverso
                        SolicitudesDeRechazoPresupuestos reverso = new SolicitudesDeRechazoPresupuestos();
                        reverso = context.SolicitudesDeRechazoPresupuestos.ToList().Where(s => s.id_facturacion_safi == id).FirstOrDefault();

                        if (reverso.estado == true)
                        {
                            reverso.estado = false;

                            //cambiar el estado del reverso
                            context.Entry(reverso).State = EntityState.Modified;
                            context.SaveChanges();
                        }

                        //anular el presupuesto en el safi general
                        SAFIGeneral perfactura = context.SAFIGeneral.Where(s => s.id_facturacion_safi == id).FirstOrDefault();
                        if (perfactura.estado == true)
                        {
                            perfactura.estado = false;

                            //cambiar el estado del presupuesto
                            context.Entry(perfactura).State = EntityState.Modified;
                            context.SaveChanges();
                        }

                        //cambiar el estado a ELI en SAFI
                        erp.ActualizarPresupuestosSAFI(1, perfactura.numero_prefactura, "");
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

        public static SolicitudesDeRechazoPresupuestos ConsultarSolicitudRechazo(int id)
        {
            using (var context = new GestionPPMEntities())
            {
                SolicitudesDeRechazoPresupuestos rechazo = new SolicitudesDeRechazoPresupuestos();
                try
                {
                    rechazo = context.SolicitudesDeRechazoPresupuestos.ToList().Where(s => s.id_facturacion_safi == id).FirstOrDefault();
                    return rechazo;
                }
                catch (Exception ex)
                {
                    return rechazo;
                }
            }
        }

        public static List<ListadoPrefacturasRechazadas> ListadoPrefacturasRechazadas()
        {
            List<ListadoPrefacturasRechazadas> listado = new List<ListadoPrefacturasRechazadas>();
            try
            {
                listado = db.ListadoPrefacturasRechazadas().ToList();
                return listado;
            }
            catch (Exception ex)
            {
                return listado;
            }
        }

    }
}
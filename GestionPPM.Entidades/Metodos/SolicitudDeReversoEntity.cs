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
    public class SolicitudDeReversoEntity
    {
        private static readonly GestionPPMEntities db = new GestionPPMEntities();

        public static bool CrearSolicitudReveso(List<int> ids, int usuarioID)
        {
            using (var transaction = db.Database.BeginTransaction())
            {
                try
                { 
                    using (var context = new GestionPPMEntities())
                    {
                        foreach (var id in ids)
                        {
                            //insertar las colicitud de reverso
                            SolicitudesDeReverso reverso = new SolicitudesDeReverso();
                            reverso.id_facturacion_safi = id;
                            reverso.fecha_solicitud_reverso = DateTime.Now;
                            reverso.id_usuario = usuarioID;
                            reverso.estado = true;

                            context.SolicitudesDeReverso.Add(reverso);
                            context.SaveChanges();

                            //enviar notificaciones
                            context.usp_guarda_envio_correo_notificaciones(10, Convert.ToInt32(id), "", 1, "");
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

        public static bool ActualizarSolicitudReveso(List<int> ids)
        {
            using (var transaction = db.Database.BeginTransaction())
            {
                try
                {
                    using (var context = new GestionPPMEntities())
                    {
                        foreach (var id in ids)
                        {
                            //insertar las colicitud de reverso
                            SolicitudesDeReverso reverso = new SolicitudesDeReverso();
                            reverso = context.SolicitudesDeReverso.ToList().Where( s => s.id_facturacion_safi == id).FirstOrDefault();                             

                            if (reverso.estado == true)
                            {
                                reverso.estado = false;

                                //cambiar el estado del reverso
                                context.Entry(reverso).State = EntityState.Modified;
                                context.SaveChanges();

                                //cambiar el estado de las prefacturas
                                context.ReversarConsolidacionPrefactura(id);
                                context.SaveChanges();

                                //enviar notificaciones
                                context.usp_guarda_envio_correo_notificaciones(11, Convert.ToInt32(reverso.id_facturacion_safi), "", 1, "");
                                context.SaveChanges();
                            }

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

        public static SolicitudesDeReverso ConsultarReversoConsolidacion(int id)
        {
            using (var context = new GestionPPMEntities())
            {
                SolicitudesDeReverso reverso = new SolicitudesDeReverso();
                try
                {
                    reverso = context.SolicitudesDeReverso.ToList().Where(s => s.id_facturacion_safi == id).FirstOrDefault();
                    return reverso;
                }
                catch (Exception ex)
                {
                    return reverso;
                }
            }
        }

    }
}
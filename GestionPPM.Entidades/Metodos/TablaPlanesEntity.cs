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
    public static class TablaPlanesEntity
    {
        private static readonly GestionPPMEntities db = new GestionPPMEntities();

        public static RespuestaTransaccion CrearTablaPlanes(TablaPlanes tablaPlanes)
        {
            try
            {
                tablaPlanes.nombre_plan = tablaPlanes.nombre_plan.ToUpper();
                tablaPlanes.estado = true;                
                db.TablaPlanes.Add(tablaPlanes);
                db.SaveChanges();

                return new RespuestaTransaccion { Estado = true, Respuesta = Mensajes.MensajeTransaccionExitosa };
            }
            catch (Exception ex)
            {
                return new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeTransaccionFallida + " ;" + ex.Message.ToString() };
            }
        }

        public static RespuestaTransaccion ActualizarTablaPlanes(TablaPlanes tablaPlanes)
        {
            try
            {
                // Por si queda el Attach de la entidad y no deja actualizar
                var local = db.TablaPlanes.FirstOrDefault(f => f.id_plan == tablaPlanes.id_plan);
                if (local != null)
                {
                    db.Entry(local).State = EntityState.Detached;
                }

                tablaPlanes.nombre_plan = tablaPlanes.nombre_plan.ToUpper();
                db.Entry(tablaPlanes).State = EntityState.Modified;
                db.SaveChanges();

                return new RespuestaTransaccion { Estado = true, Respuesta = Mensajes.MensajeTransaccionExitosa };
            }
            catch (Exception ex)
            {
                return new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeTransaccionFallida + " ;" + ex.Message.ToString() };
            }
        }

        //Eliminación Lógica
        public static RespuestaTransaccion EliminarTablaPlanes(int id)
        {
            try
            {
                var tablaPlanes = db.TablaPlanes.Find(id);

                if(tablaPlanes.estado ==true)
                {
                    tablaPlanes.estado = false;
                }
                else
                {
                    tablaPlanes.estado = true;
                }
                 
                db.Entry(tablaPlanes).State = EntityState.Modified;
                db.SaveChanges();

                return new RespuestaTransaccion { Estado = true, Respuesta = Mensajes.MensajeTransaccionExitosa };
            }
            catch (Exception ex)
            {
                return new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeTransaccionFallida + " ;" + ex.Message.ToString() };
            }
        }

        public static List<ListadoTablaPlanes> ListarTablaPlanes()
        {
            try
            {
                return db.ListadoTablaPlanes().ToList();
            }
            catch (Exception ex)
            {
                throw;
            }
        }
         
        public static TablaPlanes ConsultarTablaPlanes(int id)
        {
            try
            {
                TablaPlanes tablaPlanes = db.TablaPlanes.Find(id);
                return tablaPlanes;
            }
            catch (Exception)
            {
                throw;
            }
        }
          
        public static IEnumerable<SelectListItem> ObtenerListadoTarifario(bool gestion, string seleccionado = null)
        {
            List<SelectListItem> ListadoCatalogo = new List<SelectListItem>();
            try
            {
                using (var context = new GestionPPMEntities())
                {
                    ListadoCatalogo = context.Tarifario.Where(c => c.gestion_tarifario == gestion && c.estado_tarifario.Value).OrderBy(c => c.tipo_tarifario).Select(c => new SelectListItem
                    {
                        Text = c.tipo_tarifario,
                        Value = c.id_tarifario.ToString()
                    }).ToList();

                    if (!string.IsNullOrEmpty(seleccionado))
                    {
                        if (ListadoCatalogo.FirstOrDefault(s => s.Value == seleccionado.ToString()) != null)
                            ListadoCatalogo.FirstOrDefault(s => s.Value == seleccionado.ToString()).Selected = true;
                    }

                    return ListadoCatalogo;
                }
            }
            catch (Exception ex)
            {
                return ListadoCatalogo;
            }
        }

        public static string ObtenerTarifario(int? id)
        {
            var tarifario = "";
            try
            {
                using (var context = new GestionPPMEntities())
                {
                    if (id.HasValue)
                    {
                        var ListadoCatalogo = context.Tarifario.FirstOrDefault(c => c.id_tarifario == id && c.estado_tarifario.Value);
                        if (tarifario != null)
                        {
                            tarifario = ListadoCatalogo.tipo_tarifario;
                        }
                        return tarifario;
                    }
                    else
                        return tarifario;
                }
            }
            catch (Exception ex)
            {
                return tarifario;
            }
        }

        public static List<Tarifario> ObtenerListaTarifarioTablaDinamica(bool gestion)
        {
            List<Tarifario> ListadoTarifarios = new List<Tarifario>();
            try
            {
                using (var context = new GestionPPMEntities())
                {
                    ListadoTarifarios = context.Tarifario.Where(c => c.gestion_tarifario == gestion && c.estado_tarifario.Value).OrderBy(c => c.tipo_tarifario).ToList();

                    return ListadoTarifarios;
                }
            }
            catch (Exception ex)
            {
                return ListadoTarifarios;
            }
        }
    }
}
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
    public static class TarifarioEntity
    {
        private static readonly GestionPPMEntities db = new GestionPPMEntities();

        public static RespuestaTransaccion CrearTarifario(Tarifario Tarifario)
        {
            try
            {
                Tarifario.tipo_tarifario = Tarifario.tipo_tarifario.ToUpper();
                Tarifario.estado_tarifario = true;                
                db.Tarifario.Add(Tarifario);
                db.SaveChanges();

                return new RespuestaTransaccion { Estado = true, Respuesta = Mensajes.MensajeTransaccionExitosa };
            }
            catch (Exception ex)
            {
                return new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeTransaccionFallida + " ;" + ex.Message.ToString() };
            }
        }

        public static RespuestaTransaccion ActualizarTarifario(Tarifario Tarifario)
        {
            try
            {
                // Por si queda el Attach de la entidad y no deja actualizar
                var local = db.Tarifario.FirstOrDefault(f => f.id_tarifario == Tarifario.id_tarifario);
                if (local != null)
                {
                    db.Entry(local).State = EntityState.Detached;
                }

                Tarifario.tipo_tarifario = Tarifario.tipo_tarifario.ToUpper();
                db.Entry(Tarifario).State = EntityState.Modified;
                db.SaveChanges();

                return new RespuestaTransaccion { Estado = true, Respuesta = Mensajes.MensajeTransaccionExitosa };
            }
            catch (Exception ex)
            {
                return new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeTransaccionFallida + " ;" + ex.Message.ToString() };
            }
        }

        //Eliminación Lógica
        public static RespuestaTransaccion EliminarTarifario(int id)
        {
            try
            {
                var Tarifario = db.Tarifario.Find(id);

                if(Tarifario.estado_tarifario ==true)
                {
                    Tarifario.estado_tarifario = false;
                }
                else
                {
                    Tarifario.estado_tarifario = true;
                }
                

                db.Entry(Tarifario).State = EntityState.Modified;
                db.SaveChanges();

                return new RespuestaTransaccion { Estado = true, Respuesta = Mensajes.MensajeTransaccionExitosa };
            }
            catch (Exception ex)
            {
                return new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeTransaccionFallida + " ;" + ex.Message.ToString() };
            }
        }

        public static List<ListadoTablaCostos> ListarTarifario()
        {
            try
            {
                return db.ListadoTablaCostos().ToList();
            }
            catch (Exception ex)
            {
                throw;
            }
        }


        public static Tarifario ConsultarTarifario(int id)
        {
            try
            {
                Tarifario rol = db.Tarifario.Find(id);
                return rol;
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
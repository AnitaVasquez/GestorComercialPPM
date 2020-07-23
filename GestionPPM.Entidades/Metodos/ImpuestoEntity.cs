using GestionPPM.Entidades.Modelo;
using GestionPPM.Repositorios;
using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.Linq; 
using System.Web.Mvc;

namespace GestionPPM.Entidades.Metodos
{
    public static class ImpuestoEntity
    {
        private static readonly GestionPPMEntities db = new GestionPPMEntities();

        public static RespuestaTransaccion CrearImpuesto(Impuesto impuesto)
        {
            try
            {
                impuesto.nombre_impuesto = impuesto.nombre_impuesto.ToUpper();
                impuesto.estado_impuesto = true;                
                db.Impuesto.Add(impuesto);
                db.SaveChanges();

                return new RespuestaTransaccion { Estado = true, Respuesta = Mensajes.MensajeTransaccionExitosa };
            }
            catch (Exception ex)
            {
                return new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeTransaccionFallida + " ;" + ex.Message.ToString() };
            }
        }

        public static RespuestaTransaccion ActualizarImpuestos(Impuesto impuesto)
        {
            try
            {
                // Por si queda el Attach de la entidad y no deja actualizar
                var local = db.Impuesto.FirstOrDefault(f => f.id_impuesto == impuesto.id_impuesto);
                if (local != null)
                {
                    db.Entry(local).State = EntityState.Detached;
                }

                impuesto.nombre_impuesto = impuesto.nombre_impuesto.ToUpper();  
                db.Entry(impuesto).State = EntityState.Modified;
                db.SaveChanges();

                return new RespuestaTransaccion { Estado = true, Respuesta = Mensajes.MensajeTransaccionExitosa };
            }
            catch (Exception ex)
            {
                return new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeTransaccionFallida + " ;" + ex.Message.ToString() };
            }
        }

        //Eliminación Lógica
        public static RespuestaTransaccion EliminarImpuesto(int id)
        {
            try
            {
                var impuesto = db.Impuesto.Find(id);

                if(impuesto.estado_impuesto ==true)
                {
                    impuesto.estado_impuesto = false;
                }
                else
                {
                    impuesto.estado_impuesto = true;
                }
                

                db.Entry(impuesto).State = EntityState.Modified;
                db.SaveChanges();

                return new RespuestaTransaccion { Estado = true, Respuesta = Mensajes.MensajeTransaccionExitosa };
            }
            catch (Exception ex)
            {
                return new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeTransaccionFallida + " ;" + ex.Message.ToString() };
            }
        }

        public static List<ListadoImpuesto> ListarImpuestos()
        {
            try
            {
                return db.ListadoImpuesto().ToList();
            }
            catch (Exception ex)
            {
                throw;
            }
        }

        public static IEnumerable<SelectListItem> ObtenerListadoImpuestos(string seleccionado = null)
        {
            List<SelectListItem> ListadoImpuestos = new List<SelectListItem>();
            try
            { 
                var listadoImpuestos = db.Impuesto.Where(c => c.estado_impuesto == true).OrderBy(c => c.nombre_impuesto).Select(c => new SelectListItem
                {
                    Text = c.nombre_impuesto,
                    Value = c.id_impuesto.ToString(),
                }).ToList();

                if (!string.IsNullOrEmpty(seleccionado))
                {
                    if (listadoImpuestos.FirstOrDefault(s => s.Value == seleccionado.ToString()) != null)
                        listadoImpuestos.FirstOrDefault(s => s.Value == seleccionado.ToString()).Selected = true;
                }
                return listadoImpuestos;
            }
            catch (Exception ex)
            {
                return ListadoImpuestos;
            }
        }

        public static IEnumerable<SelectListItem> ObtenerListadoImpuestosTotales(string seleccionado = null)
        {
            List<SelectListItem> ListadoImpuestos = new List<SelectListItem>();
            try
            {
                var listadoImpuestos = db.Impuesto.OrderBy(c => c.nombre_impuesto).Select(c => new SelectListItem
                {
                    Text = c.nombre_impuesto,
                    Value = c.id_impuesto.ToString(),
                }).ToList();

                if (!string.IsNullOrEmpty(seleccionado))
                {
                    if (listadoImpuestos.FirstOrDefault(s => s.Value == seleccionado.ToString()) != null)
                        listadoImpuestos.FirstOrDefault(s => s.Value == seleccionado.ToString()).Selected = true;
                }
                return listadoImpuestos;
            }
            catch (Exception ex)
            {
                return ListadoImpuestos;
            }
        }

        public static Impuesto ConsultarImpuesto(int id)
        {
            try
            {
                Impuesto impuesto = db.Impuesto.Find(id);
                return impuesto;
            }
            catch (Exception)
            {
                throw;
            }
        }
            
        public static List<Impuesto> ObtenerListaImpuestos()
        {
            List<Impuesto> ListadoImpuestos = new List<Impuesto>();
            try
            {
                using (var context = new GestionPPMEntities())
                {
                    ListadoImpuestos = context.Impuesto.Where(c => c.estado_impuesto == true).ToList();

                    return ListadoImpuestos;
                }
            }
            catch (Exception ex)
            {
                return ListadoImpuestos;
            }
        }

        public static bool ValidarImpuestoActivo(int id)
        {
            bool flag = true;

            List<Impuesto> ListadoImpuestos = new List<Impuesto>();

            ListadoImpuestos = db.Impuesto.Where(c => c.estado_impuesto == true && c.id_impuesto == id).ToList();

            if(ListadoImpuestos.Count()==0)
            {
                flag = false;
            }else
            {
                flag = true;
            }
             
            return flag;
        }
    }
}
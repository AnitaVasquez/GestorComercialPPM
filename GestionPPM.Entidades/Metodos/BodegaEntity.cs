using GestionPPM.Entidades.Modelo;
using GestionPPM.Repositorios;
using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.Linq; 
using System.Web.Mvc;

namespace GestionPPM.Entidades.Metodos
{
    public static class BodegaEntity
    {
        private static readonly GestionPPMEntities db = new GestionPPMEntities();

        public static RespuestaTransaccion CrearBodega(Bodega bodega)
        {
            try
            {
                bodega.nombre_bodega = bodega.nombre_bodega.ToUpper();
                bodega.codigo_bodega = bodega.codigo_bodega.ToUpper();
                bodega.estado = true;                
                db.Bodega.Add(bodega);
                db.SaveChanges();

                return new RespuestaTransaccion { Estado = true, Respuesta = Mensajes.MensajeTransaccionExitosa };
            }
            catch (Exception ex)
            {
                return new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeTransaccionFallida + " ;" + ex.Message.ToString() };
            }
        }

        public static RespuestaTransaccion ActualizarBodega(Bodega bodega)
        {
            try
            {
                // Por si queda el Attach de la entidad y no deja actualizar
                var local = db.Bodega.FirstOrDefault(f => f.id_bodega == bodega.id_bodega);
                if (local != null)
                {
                    db.Entry(local).State = EntityState.Detached;
                }

                bodega.nombre_bodega = bodega.nombre_bodega.ToUpper();
                bodega.codigo_bodega = bodega.codigo_bodega.ToUpper();
                db.Entry(bodega).State = EntityState.Modified;
                db.SaveChanges();

                return new RespuestaTransaccion { Estado = true, Respuesta = Mensajes.MensajeTransaccionExitosa };
            }
            catch (Exception ex)
            {
                return new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeTransaccionFallida + " ;" + ex.Message.ToString() };
            }
        }

        //Eliminación Lógica
        public static RespuestaTransaccion EliminarBodega(int id)
        {
            try
            {
                var bodega = db.Bodega.Find(id);

                if(bodega.estado == true)
                {
                    bodega.estado = false;
                }
                else
                {
                    bodega.estado = true;
                }

                db.Entry(bodega).State = EntityState.Modified;
                db.SaveChanges();

                return new RespuestaTransaccion { Estado = true, Respuesta = Mensajes.MensajeTransaccionExitosa };
            }
            catch (Exception ex)
            {
                return new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeTransaccionFallida + " ;" + ex.Message.ToString() };
            }
        }

        public static List<ListadoBodega> ListarBodegas()
        {
            try
            {
                return db.ListadoBodega().ToList();
            }
            catch (Exception ex)
            {
                throw;
            }
        }

        public static IEnumerable<SelectListItem> ObtenerListadoBodega(string seleccionado = null)
        {
            List<SelectListItem> ListadoBodega = new List<SelectListItem>();
            try
            { 
                var listadoBodegas = db.Bodega.Where(c => c.estado == true).OrderBy(c => c.nombre_bodega).Select(c => new SelectListItem
                {
                    Text = c.nombre_bodega,
                    Value = c.id_bodega.ToString(),
                }).ToList();

                if (!string.IsNullOrEmpty(seleccionado))
                {
                    if (listadoBodegas.FirstOrDefault(s => s.Value == seleccionado.ToString()) != null)
                        listadoBodegas.FirstOrDefault(s => s.Value == seleccionado.ToString()).Selected = true;
                }
                return listadoBodegas;
            }
            catch (Exception ex)
            {
                return ListadoBodega;
            }
        }
          
        public static Bodega ConsultarBodega(int id)
        {
            try
            {
                Bodega bodega = db.Bodega.Find(id);
                return bodega;
            }
            catch (Exception)
            {
                throw;
            }
        } 
    }
}
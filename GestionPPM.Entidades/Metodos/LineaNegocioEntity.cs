using GestionPPM.Entidades.Modelo;
using GestionPPM.Repositorios;
using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.Linq; 
using System.Web.Mvc;

namespace GestionPPM.Entidades.Metodos
{
    public static class LineaNegocioEntity
    {
        private static readonly GestionPPMEntities db = new GestionPPMEntities();

        public static RespuestaTransaccion CrearLineaNegocio(LineaNegocio lineaNegocio)
        {
            try
            {   
                db.LineaNegocio.Add(lineaNegocio);
                db.SaveChanges();

                return new RespuestaTransaccion { Estado = true, Respuesta = Mensajes.MensajeTransaccionExitosa };
            }
            catch (Exception ex)
            {
                return new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeTransaccionFallida + " ;" + ex.Message.ToString() };
            }
        }

        public static RespuestaTransaccion ActualizarLineaNegocio(LineaNegocio lineaNegocio)
        {
            try
            {
                // Por si queda el Attach de la entidad y no deja actualizar
                var local = db.LineaNegocio.FirstOrDefault(f => f.id_linea_negocio == lineaNegocio.id_linea_negocio);
                if (local != null)
                {
                    db.Entry(local).State = EntityState.Detached;
                } 
                db.Entry(lineaNegocio).State = EntityState.Modified;
                db.SaveChanges();

                return new RespuestaTransaccion { Estado = true, Respuesta = Mensajes.MensajeTransaccionExitosa };
            }
            catch (Exception ex)
            {
                return new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeTransaccionFallida + " ;" + ex.Message.ToString() };
            }
        }

        //Eliminación Lógica
        public static RespuestaTransaccion EliminarCodigoProducto(int id)
        {
            try
            {
                var producto = db.CodigoProducto.Find(id);

                if(producto.estado == true)
                {
                    producto.estado = false;
                }
                else
                {
                    producto.estado = true;
                }

                db.Entry(producto).State = EntityState.Modified;
                db.SaveChanges();

                return new RespuestaTransaccion { Estado = true, Respuesta = Mensajes.MensajeTransaccionExitosa };
            }
            catch (Exception ex)
            {
                return new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeTransaccionFallida + " ;" + ex.Message.ToString() };
            }
        }

        public static List<LineaNegocio> ListadoLineasNegocio()
        {
            try
            {
                return db.LineaNegocio.ToList();
            }
            catch (Exception ex)
            {
                throw;
            }
        }

        public static IEnumerable<SelectListItem> ObtenerListadoLineaNegocio(string seleccionado = null)
        {
            List<SelectListItem> ListadoLineasNegocio = new List<SelectListItem>();
            try
            { 
                var listadoLineaNegocio = db.LineaNegocio.Where(c => c.estado == true).OrderBy(c => c.nombre).Select(c => new SelectListItem
                {
                    Text = c.nombre,
                    Value = c.id_linea_negocio.ToString(),
                }).ToList();

                if (!string.IsNullOrEmpty(seleccionado))
                {
                    if (listadoLineaNegocio.FirstOrDefault(s => s.Value == seleccionado.ToString()) != null)
                        listadoLineaNegocio.FirstOrDefault(s => s.Value == seleccionado.ToString()).Selected = true;
                }
                return listadoLineaNegocio;
            }
            catch (Exception ex)
            {
                return ListadoLineasNegocio;
            }
        }

        public static LineaNegocio ConsultarLlinea(int id)
        {
            try
            {
                LineaNegocio producto = db.LineaNegocio.Find(id);
                return producto;
            }
            catch (Exception)
            {
                throw;
            }
        }

        public static IEnumerable<SelectListItem> ConsultarLineaNegocioInverso(int id, string seleccionado = null)
        {
            try
            {
                using (var context = new GestionPPMEntities())
                {
                    //obtener el codigo general
                    var producto = ProductosGestorEntity.ConsultarProducto(id);
                    
                    if (producto != null)
                    {
                        var productoGeneral = ProductosGeneralGestorEntity.ConsultarProductoGeneral(producto.id_producto_general);
                        var sublineaNegocio = SublineaNegocioEntity.ConsultarSublinea(productoGeneral.id_sublinea_negocio);
                        var lineaNegocio = ConsultarLlinea(sublineaNegocio.id_linea_negocio);
                        seleccionado = lineaNegocio.id_linea_negocio.ToString();
                    }

                    var catalogo = context.LineaNegocio.Where(c=> c.estado == true).OrderBy(c => c.nombre).Select(c => new SelectListItem
                    {
                        Text = c.nombre,
                        Value = c.id_linea_negocio.ToString()
                    }).ToList();

                    if (!string.IsNullOrEmpty(seleccionado))
                    {
                        if (catalogo.FirstOrDefault(s => s.Value == seleccionado.ToString()) != null)
                            catalogo.FirstOrDefault(s => s.Value == seleccionado.ToString()).Selected = true;
                    }

                    return catalogo;
                }

            }
            catch (Exception)
            {
                throw;
            }
        }

    }
}
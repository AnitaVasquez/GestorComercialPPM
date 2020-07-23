using GestionPPM.Entidades.Modelo;
using GestionPPM.Repositorios;
using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.Linq; 
using System.Web.Mvc;

namespace GestionPPM.Entidades.Metodos
{
    public static class ProductosGestorEntity
    {
        private static readonly GestionPPMEntities db = new GestionPPMEntities();

        public static RespuestaTransaccion CrearProductosGestor(ProdutosGestor productosGestor)
        {
            try
            {
                productosGestor.estado = true;                
                db.ProdutosGestor.Add(productosGestor);
                db.SaveChanges();

                return new RespuestaTransaccion { Estado = true, Respuesta = Mensajes.MensajeTransaccionExitosa };
            }
            catch (Exception ex)
            {
                return new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeTransaccionFallida + " ;" + ex.Message.ToString() };
            }
        }

        public static RespuestaTransaccion ActualizarProductosGestor(ProdutosGestor productosGestor)
        {
            try
            {
                // Por si queda el Attach de la entidad y no deja actualizar
                var local = db.ProdutosGestor.FirstOrDefault(f => f.id_producto_gestor == productosGestor.id_producto_gestor);
                if (local != null)
                {
                    db.Entry(local).State = EntityState.Detached;
                } 
                db.Entry(productosGestor).State = EntityState.Modified;
                db.SaveChanges();

                return new RespuestaTransaccion { Estado = true, Respuesta = Mensajes.MensajeTransaccionExitosa };
            }
            catch (Exception ex)
            {
                return new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeTransaccionFallida + " ;" + ex.Message.ToString() };
            }
        }

        //Eliminación Lógica
        public static RespuestaTransaccion EliminarProductosGestor(int id)
        {
            try
            {
                var producto = db.ProdutosGestor.Find(id);

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

        public static List<ListadoProductosGestor> ListarProductosGestor()
        {
            try
            {
                return db.ListadoProductosGestor().ToList();
            }
            catch (Exception ex)
            {
                throw;
            }
        }
          
        public static ProdutosGestor ConsultarProducto(int id)
        {
            try
            {
                ProdutosGestor producto = db.ProdutosGestor.Find(id);
                return producto;
            }
            catch (Exception)
            {
                throw;
            }
        }

        public static List<ProdutosGestor> ListadoProductoGestor()
        {
            try
            {
                return db.ProdutosGestor.ToList();
            }
            catch (Exception ex)
            {
                throw;
            }
        }

        public static IEnumerable<SelectListItem> ConsultarProductosGestor(int id, string seleccionado = null)
        {
            try
            {
                using (var context = new GestionPPMEntities())
                {
                    var catalogo = context.ProdutosGestor.Where(s => s.id_producto_general == id && s.estado == true).OrderBy(c => c.nombre).Select(c => new SelectListItem
                    {
                        Text = c.nombre,
                        Value = c.id_producto_gestor.ToString()
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

        public static IEnumerable<SelectListItem> ConsultarProductosGestorInverso(int id, string seleccionado = null)
        {
            try
            {
                using (var context = new GestionPPMEntities())
                {
                    //obtener el codigo general
        
            var producto = ConsultarProducto(id);

                    if(producto != null)
                    { 
                        var catalogo = context.ProdutosGestor.Where(s => s.id_producto_general == producto.id_producto_general && s.estado == true).OrderBy(c => c.nombre).Select(c => new SelectListItem
                        {
                            Text = c.nombre,
                            Value = c.id_producto_gestor.ToString()
                        }).ToList();

                        if (!string.IsNullOrEmpty(seleccionado))
                        {
                            if (catalogo.FirstOrDefault(s => s.Value == seleccionado.ToString()) != null)
                                catalogo.FirstOrDefault(s => s.Value == seleccionado.ToString()).Selected = true;
                        }

                        return catalogo;
                    }
                    else
                    {
                        var catalogo = context.ProdutosGestor.Where(s => s.id_producto_general == 0 && s.estado == true).OrderBy(c => c.nombre).Select(c => new SelectListItem
                        {
                            Text = c.nombre,
                            Value = c.id_producto_gestor.ToString()
                        }).ToList();

                        if (!string.IsNullOrEmpty(seleccionado))
                        {
                            if (catalogo.FirstOrDefault(s => s.Value == seleccionado.ToString()) != null)
                                catalogo.FirstOrDefault(s => s.Value == seleccionado.ToString()).Selected = true;
                        }

                        return catalogo;
                    }
                }

            }
            catch (Exception)
            {
                throw;
            }
        }

    }
}
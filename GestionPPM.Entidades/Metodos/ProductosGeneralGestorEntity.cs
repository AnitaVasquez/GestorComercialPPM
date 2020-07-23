using GestionPPM.Entidades.Modelo;
using GestionPPM.Repositorios;
using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.Linq; 
using System.Web.Mvc;

namespace GestionPPM.Entidades.Metodos
{
    public static class ProductosGeneralGestorEntity
    {
        private static readonly GestionPPMEntities db = new GestionPPMEntities();

        public static RespuestaTransaccion CrearLineaNegocio(ProdutosGeneralGestor produtosGeneralGestor)
        {
            try
            {
                produtosGeneralGestor.estado = true;                
                db.ProdutosGeneralGestor.Add(produtosGeneralGestor);
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

        public static IEnumerable<SelectListItem> ObtenerListadoProdutoGeneralGestor(string seleccionado = null)
        {
            List<SelectListItem> ListadoProductoGeneral = new List<SelectListItem>();
            try
            {
                var listadoProductoGeneral = db.ProdutosGeneralGestor.Where(c => c.estado == true).OrderBy(c => c.nombre).Select(c => new SelectListItem
                {
                    Text = c.nombre,
                    Value = c.id_producto_general.ToString(),
                }).ToList();

                if (!string.IsNullOrEmpty(seleccionado))
                {
                    if (listadoProductoGeneral.FirstOrDefault(s => s.Value == seleccionado.ToString()) != null)
                        listadoProductoGeneral.FirstOrDefault(s => s.Value == seleccionado.ToString()).Selected = true;
                }
                return listadoProductoGeneral;
            }
            catch (Exception ex)
            {
                return ListadoProductoGeneral;
            }
        }

        public static IEnumerable<SelectListItem> ConsultarProductosGenerales(int id)
        {
            try
            {
                using (var context = new GestionPPMEntities())
                {
                    var catalogo = context.ProdutosGeneralGestor.Where(s => s.id_sublinea_negocio == id && s.estado == true).OrderBy(c => c.nombre).Select(c => new SelectListItem
                    {
                        Text = c.nombre,
                        Value = c.id_producto_general.ToString()
                    }).ToList();

                    return catalogo;
                }

            }
            catch (Exception)
            {
                throw;
            }
        }

        public static List<ProdutosGeneralGestor> ListadoProductoGeneral()
        {
            try
            {
                return db.ProdutosGeneralGestor.ToList();
            }
            catch (Exception ex)
            {
                throw;
            }
        }

        public static ProdutosGeneralGestor ConsultarProductoGeneral(int id)
        { 
            try
            {
                ProdutosGeneralGestor producto = db.ProdutosGeneralGestor.Find(id);
                return producto;
            }
            catch (Exception)
            {
                throw;
            }
        }

        public static IEnumerable<SelectListItem> ConsultarProductosGeneralInverso(int id, string seleccionado = null)
        {
            try
            {
                using (var context = new GestionPPMEntities())
                {
                    //obtener el codigo general
                    var producto = ProductosGestorEntity.ConsultarProducto(id);

                    if (producto != null)
                    {

                        var productoGeneral = ConsultarProductoGeneral(producto.id_producto_general);
                        seleccionado = productoGeneral.id_producto_general.ToString();

                        var catalogo = context.ProdutosGeneralGestor.Where(s => s.id_sublinea_negocio == productoGeneral.id_sublinea_negocio && s.estado == true).OrderBy(c => c.nombre).Select(c => new SelectListItem
                        {
                            Text = c.nombre,
                            Value = c.id_producto_general.ToString()
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
                        var catalogo = context.ProdutosGeneralGestor.Where(s => s.id_sublinea_negocio == 0 && s.estado == true).OrderBy(c => c.nombre).Select(c => new SelectListItem
                        {
                            Text = c.nombre,
                            Value = c.id_producto_general.ToString()
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
using GestionPPM.Entidades.Modelo;
using GestionPPM.Repositorios;
using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.Linq; 
using System.Web.Mvc;

namespace GestionPPM.Entidades.Metodos
{
    public static class SublineaNegocioEntity
    {
        private static readonly GestionPPMEntities db = new GestionPPMEntities();

        public static RespuestaTransaccion CrearSublineaNegocio(SublineaNegocio sublineaNegocio)
        {
            try
            {                
                db.SublineaNegocio.Add(sublineaNegocio);
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

        public static List<ListadoProducto> ListarProductos()
        {
            try
            {
                return db.ListadoProducto().ToList();
            }
            catch (Exception ex)
            {
                throw;
            }
        }

        public static IEnumerable<SelectListItem> ObtenerListadoSublineaNegocio(string seleccionado = null)
        {
            List<SelectListItem> ListadoSublineasNegocio = new List<SelectListItem>();
            try
            {
                var listadoSublineaNegocio = db.SublineaNegocio.Where(c => c.estado == true).OrderBy(c => c.nombre).Select(c => new SelectListItem
                {
                    Text = c.nombre,
                    Value = c.id_sublinea_negocio.ToString(),
                }).ToList();

                if (!string.IsNullOrEmpty(seleccionado))
                {
                    if (listadoSublineaNegocio.FirstOrDefault(s => s.Value == seleccionado.ToString()) != null)
                        listadoSublineaNegocio.FirstOrDefault(s => s.Value == seleccionado.ToString()).Selected = true;
                }
                return listadoSublineaNegocio;
            }
            catch (Exception ex)
            {
                return ListadoSublineasNegocio;
            }
        }

        public static IEnumerable<SelectListItem> ConsultarSublineaNegocio(int id)
        {
            try
            {
                using (var context = new GestionPPMEntities())
                {
                    var catalogo = context.SublineaNegocio.Where(s => s.id_linea_negocio == id && s.estado == true).OrderBy(c => c.nombre).Select(c => new SelectListItem
                    {
                        Text = c.nombre,
                        Value = c.id_sublinea_negocio.ToString()
                    }).ToList();

                    return catalogo;
                }

            }
            catch (Exception)
            {
                throw;
            }
        }

        public static List<SublineaNegocio> ListadoSublineasNegocio()
        {
            try
            {
                return db.SublineaNegocio.ToList();
            }
            catch (Exception ex)
            {
                throw;
            }
        }

        public static SublineaNegocio ConsultarSublinea(int id)
        {
            try
            {
                SublineaNegocio producto = db.SublineaNegocio.Find(id);
                return producto;
            }
            catch (Exception)
            {
                throw;
            }
        }

        public static IEnumerable<SelectListItem> ConsultarSublineaNegocioInverso(int id, string seleccionado = null)
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
                        var sublineaNegocio = ConsultarSublinea(productoGeneral.id_sublinea_negocio);
                        seleccionado = sublineaNegocio.id_sublinea_negocio.ToString();

                        var catalogo = context.SublineaNegocio.Where(s => s.id_linea_negocio == sublineaNegocio.id_linea_negocio && s.estado == true).OrderBy(c => c.nombre).Select(c => new SelectListItem
                        {
                            Text = c.nombre,
                            Value = c.id_sublinea_negocio.ToString()
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
                        var catalogo = context.SublineaNegocio.Where(s => s.id_linea_negocio == 0 && s.estado == true).OrderBy(c => c.nombre).Select(c => new SelectListItem
                        {
                            Text = c.nombre,
                            Value = c.id_sublinea_negocio.ToString()
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
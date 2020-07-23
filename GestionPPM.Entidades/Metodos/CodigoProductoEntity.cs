using GestionPPM.Entidades.Modelo;
using GestionPPM.Repositorios;
using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.Linq; 
using System.Web.Mvc;

namespace GestionPPM.Entidades.Metodos
{
    public static class CodigoProductoEntity
    {
        private static readonly GestionPPMEntities db = new GestionPPMEntities();

        public static RespuestaTransaccion CrearCodigoProducto(CodigoProducto producto)
        {
            try
            {
                producto.nombre_producto = producto.nombre_producto.ToUpper();
                producto.codigo_producto = producto.codigo_producto.ToUpper();
                producto.estado = true;                
                db.CodigoProducto.Add(producto);
                db.SaveChanges();

                return new RespuestaTransaccion { Estado = true, Respuesta = Mensajes.MensajeTransaccionExitosa };
            }
            catch (Exception ex)
            {
                return new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeTransaccionFallida + " ;" + ex.Message.ToString() };
            }
        }

        public static RespuestaTransaccion ActualizarCodigoProducto(CodigoProducto producto)
        {
            try
            {
                // Por si queda el Attach de la entidad y no deja actualizar
                var local = db.CodigoProducto.FirstOrDefault(f => f.id_codigo_producto == producto.id_codigo_producto);
                if (local != null)
                {
                    db.Entry(local).State = EntityState.Detached;
                }

                producto.nombre_producto = producto.nombre_producto.ToUpper();
                producto.codigo_producto = producto.codigo_producto.ToUpper();
                db.Entry(producto).State = EntityState.Modified;
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

        public static IEnumerable<SelectListItem> ObtenerListadoBodega(string seleccionado = null)
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
          
        public static CodigoProducto ConsultarProducto(int id)
        {
            try
            {
                CodigoProducto producto = db.CodigoProducto.Find(id);
                return producto;
            }
            catch (Exception)
            {
                throw;
            }
        }

        public static bool TipoCodigoProductoExistente(int bodega, int catalogo, int codigo_producto)
        {
            try
            {
                using (var context = new GestionPPMEntities())
                {
                    //Validar el mismo tarifario y la misma bodega
                    var objeto = context.CodigoProducto.FirstOrDefault(s => s.id_bodega == bodega && s.id_catalogo == catalogo && s.id_codigo_producto != codigo_producto);

                    if (objeto != null)
                        return true;
                    else
                        return false;
                }
            }
            catch (Exception ex)
            {
                return true;
            }
        }
    }
}
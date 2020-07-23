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
    public static class MenuEntity
    {
        private static readonly GestionPPMEntities db = new GestionPPMEntities();

        public static RespuestaTransaccion CrearMenu(Menu Menu)
        {
            try
            {
                Menu.estado_menu = true;
                Menu.orden_menu = 0;
                db.Menu.Add(Menu);
                db.SaveChanges();

                return new RespuestaTransaccion { Estado = true, Respuesta = Mensajes.MensajeTransaccionExitosa };
            }
            catch (Exception ex)
            {
                return new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeTransaccionFallida + " ;" + ex.Message.ToString() };
            }
        }

        public static RespuestaTransaccion ActualizarMenu(Menu Menu)
        {
            try
            {
                // Por si queda el Attach de la entidad y no deja actualizar
                var local = db.Menu.FirstOrDefault(f => f.id_menu == Menu.id_menu);
                if (local != null)
                {
                    db.Entry(local).State = EntityState.Detached;
                }

                db.Entry(Menu).State = EntityState.Modified;
                db.SaveChanges();

                return new RespuestaTransaccion { Estado = true, Respuesta = Mensajes.MensajeTransaccionExitosa };
            }
            catch (Exception ex)
            {
                return new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeTransaccionFallida + " ;" + ex.Message.ToString() };
            }
        }

        //Eliminación Lógica
        public static RespuestaTransaccion EliminarMenu(int id)
        {
            try
            {
                var Menu = db.Menu.Find(id);

                Menu.estado_menu = false;

                db.Entry(Menu).State = EntityState.Modified;
                db.SaveChanges();

                return new RespuestaTransaccion { Estado = true, Respuesta = Mensajes.MensajeTransaccionExitosa };
            }
            catch (Exception ex)
            {
                return new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeTransaccionFallida + " ;" + ex.Message.ToString() };
            }
        }

        public static List<ListadoMenu> ListarMenu()
        {
            try
            {
                //db.Configuration.ProxyCreationEnabled = false;
                return db.ListadoMenu().ToList();
            }
            catch (Exception)
            {
                throw;
            }
        }

        public static List<Menu> ListarMenuHijos()
        {
            try
            {
                //db.Configuration.ProxyCreationEnabled = false;
                return db.Menu.Where(s => s.id_menu_padre != null && s.estado_menu == true).ToList();
            }
            catch (Exception)
            {
                throw;
            }
        }

        public static Menu ConsultarMenu(int id)
        {
            try
            {
                Menu rol = db.Menu.Find(id);
                return rol;
            }
            catch (Exception)
            {
                throw;
            }
        }

        public static IEnumerable<SelectListItem> ObtenerListadoMenusAplicacion()
        {
            List<SelectListItem> listado = new List<SelectListItem>();
            try
            {
                listado = db.ListadoMenu().Select(c => new SelectListItem
                {
                    Text = c.Ruta_Acceso,
                    Value = c.Id.ToString()
                }).ToList();

                return listado;
            }
            catch (Exception)
            {
                return listado;
                throw;
            }

        }
        public static List<MenuRutaAcceso> ConsultarRutaMenu(int id)
        {

            try
            {

                List<MenuRutaAcceso> menu = db.ConsultarMenuRutaAcceso(id).ToList();

                return menu;
            }
            catch (Exception ex)
            {
                throw;
            }

        }

    }
}
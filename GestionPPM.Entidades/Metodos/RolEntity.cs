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
    public static class RolEntity
    {
        private static readonly GestionPPMEntities db = new GestionPPMEntities();

        public static RespuestaTransaccion CrearRol(Rol rol, List<int> idPerfiles)
        {
            using (var transaction = db.Database.BeginTransaction())
            {
                try
                {
                    rol.nombre_rol = rol.nombre_rol.ToUpper();
                    rol.estado_rol = true;
                    db.Rol.Add(rol);
                    db.SaveChanges();

                    var rolesPerfilesAnteriores = db.RolPerfil.Where(s => s.id_rol == rol.id_rol).ToList();
                    foreach (var item in rolesPerfilesAnteriores)
                    {
                        db.RolPerfil.Remove(item);
                        db.SaveChanges();
                    }

                    List<RolPerfil> ListadoRolesPerfiles = new List<RolPerfil>();

                    foreach (var item in idPerfiles)
                    {
                        db.RolPerfil.Add(new RolPerfil
                        {
                            id_rol = rol.id_rol,
                            id_perfil = item ,
                        });
                        db.SaveChanges();
                    }

                    transaction.Commit();
                    return new RespuestaTransaccion { Estado = true, Respuesta = Mensajes.MensajeTransaccionExitosa };
                }

                catch (Exception ex)
                {
                    transaction.Rollback();
                    return new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeTransaccionFallida + " ;" + ex.Message.ToString() };
                }
            }
        }

        public static RespuestaTransaccion ActualizarRol(Rol rol, List<int> idPerfiles)
        {
            try
            {
                // Por si queda el Attach de la entidad y no deja actualizar
                var local = db.Rol.FirstOrDefault(f => f.id_rol == rol.id_rol);
                if (local != null)
                {
                    db.Entry(local).State = EntityState.Detached;
                }

                var rolesPerfilesAnteriores = db.RolPerfil.Where(s => s.id_rol == rol.id_rol).ToList();
                foreach (var item in rolesPerfilesAnteriores)
                {
                    db.RolPerfil.Remove(item);
                    db.SaveChanges();
                }

                List<RolPerfil> ListadoRolesPerfiles = new List<RolPerfil>();

                foreach (var item in idPerfiles)
                {
                    db.RolPerfil.Add(new RolPerfil
                    {
                        id_rol = rol.id_rol,
                        id_perfil = item,
                    });
                    db.SaveChanges();
                }

                rol.nombre_rol = rol.nombre_rol.ToUpper();
                db.Entry(rol).State = EntityState.Modified;
                db.SaveChanges();

                return new RespuestaTransaccion { Estado = true, Respuesta = Mensajes.MensajeTransaccionExitosa };
            }
            catch (Exception ex)
            {
                return new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeTransaccionFallida + " ;" + ex.Message.ToString() };
            }
        }

        //Eliminación Lógica
        public static RespuestaTransaccion EliminarRol(int id)
        {
            try
            {
                var rol = db.Rol.Find(id);

                if(rol.estado_rol == true)
                {
                    rol.estado_rol = false;
                }
                else
                {
                    rol.estado_rol = true;
                }                

                db.Entry(rol).State = EntityState.Modified;
                db.SaveChanges();

                return new RespuestaTransaccion { Estado = true, Respuesta = Mensajes.MensajeTransaccionExitosa };
            }
            catch (Exception ex)
            {
                return new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeTransaccionFallida + " ;" + ex.Message.ToString() };
            }
        }

        public static List<ListadoRol> ListarRol()
        {
            try
            {
                return db.ListadoRol().ToList();
            }
            catch (Exception e)
            {
                throw;
            }
        } 

        //Metodo para listar roles hijos en lista desplegable
        public static IEnumerable<SelectListItem> ObtenerListadoRoles()
        {
            var ListadoRoles = db.Rol.Where(r => r.estado_rol == true).OrderBy(r => r.nombre_rol).Select(x => new SelectListItem
            {
                Text = x.nombre_rol,
                Value = x.id_rol.ToString()
            }).ToList();

            return ListadoRoles;
        }

        public static Rol ConsultarRol(int id)
        {
            try
            {
                Rol rol = db.Rol.Find(id);
                return rol;
            }
            catch (Exception)
            {
                throw;
            }
        }

        public static List<int> ListadIdsPerfilesByRol(int idRol)
        {
            try
            {
                var listadoPerfiles = db.RolPerfil.Where(s => s.id_rol == idRol).Select(s=> s.id_perfil).ToList();
                return listadoPerfiles;
            }
            catch (Exception)
            {
                throw;
            }
        }

    }
}
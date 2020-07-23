using GestionPPM.Entidades.Modelo;
using GestionPPM.Repositorios;
using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.Linq;
using System.Threading.Tasks;
using System.Web;

namespace GestionPPM.Entidades.Metodos
{
    public static class PerfilesEntity
    {
        private static readonly GestionPPMEntities db = new GestionPPMEntities();

        public static RespuestaTransaccion CrearPerfil(Perfil Perfil)
        {
            try
            {
                Perfil.nombre_perfil = Perfil.nombre_perfil.ToUpper();
                Perfil.estado_perfil = true;
                db.Perfil.Add(Perfil);
                db.SaveChanges();

                return new RespuestaTransaccion { Estado = true, Respuesta = Mensajes.MensajeTransaccionExitosa };
            }
            catch (Exception ex)
            {
                return new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeTransaccionFallida + " ;" + ex.Message.ToString() };
            }
        }

        public static RespuestaTransaccion CreacionPerfilesCargaMasiva()
        {
            try
            {
                List<string> perfilesCarga = new List<string> { "CARGA MASIVA DE CODIGO DE COTIZACION", "CARGA MASIVA DE CONTACTOS", "CARGA MASIVA DE CLIENTE" };
                string descripcionPerfiles = "Perfil que permite realizar carga masiva de información.";

                List<Perfil> perfiles = new List<Perfil>();
                foreach (var item in perfilesCarga)
                {
                    var perfil = db.Perfil.Where(s => s.nombre_perfil.Contains(item)).FirstOrDefault();

                    var nombrePerfil = (perfil.nombre_perfil ?? "").Trim();


                    if (nombrePerfil != item) {
                        perfiles.Add(new Perfil {
                            nombre_perfil = item,
                            descripcion_perfil = descripcionPerfiles,
                            estado_perfil = true,
                        });
                    }
                }


                //Guardando perfiles que no existan
                foreach (var item in perfiles)
                {
                    db.Perfil.Add(item);
                    db.SaveChanges();
                }

                return new RespuestaTransaccion { Estado = true, Respuesta = Mensajes.MensajeTransaccionExitosa };
            }
            catch (Exception ex)
            {
                return new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeTransaccionFallida + " ;" + ex.Message.ToString() };
            }
        }

        public static RespuestaTransaccion CrearPerfil(Perfil Perfil, List<int> opcionesMenu)
        {
            using (var transaction = db.Database.BeginTransaction())
            {
                try
                {
                    Perfil.nombre_perfil = Perfil.nombre_perfil.ToUpper();
                    Perfil.estado_perfil = true;
                    db.Perfil.Add(Perfil);
                    db.SaveChanges();

                    var opcionesPerfilMenuAnteriores = db.PerfilMenu.Where(s => s.id_perfil == Perfil.id_perfil).ToList();
                    foreach (var item in opcionesPerfilMenuAnteriores)
                    {
                        db.PerfilMenu.Remove(item);
                        db.SaveChanges();
                    }

                    //List<RolPerfil> ListadoRolesPerfiles = new List<RolPerfil>();

                    foreach (var item in opcionesMenu)
                    {
                        db.PerfilMenu.Add(new PerfilMenu
                        {
                            id_perfil = Perfil.id_perfil,
                            id_menu = item,
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

        public static RespuestaTransaccion ActualizarPerfil(Perfil Perfil)
        {
            try
            {
                // Por si queda el Attach de la entidad y no deja actualizar
                var local = db.Perfil.FirstOrDefault(f => f.id_perfil == Perfil.id_perfil);
                if (local != null)
                {
                    db.Entry(local).State = EntityState.Detached;
                }

                Perfil.nombre_perfil = Perfil.nombre_perfil.ToUpper();
                db.Entry(Perfil).State = EntityState.Modified;
                db.SaveChanges();

                return new RespuestaTransaccion { Estado = true, Respuesta = Mensajes.MensajeTransaccionExitosa };
            }
            catch (Exception ex)
            {
                return new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeTransaccionFallida + " ;" + ex.Message.ToString() };
            }
        }

         public static RespuestaTransaccion ActualizarPerfil(Perfil Perfil, List<int> opcionesMenu)
        {
            using (var transaction = db.Database.BeginTransaction())
            {
                try
                {
                    // Por si queda el Attach de la entidad y no deja actualizar
                    var local = db.Perfil.FirstOrDefault(f => f.id_perfil == Perfil.id_perfil);
                    if (local != null)
                    {
                        db.Entry(local).State = EntityState.Detached;
                    }

                    var opcionesPerfilMenuAnteriores = db.PerfilMenu.Where(s => s.id_perfil == Perfil.id_perfil).ToList();
                    foreach (var item in opcionesPerfilMenuAnteriores)
                    {
                        db.PerfilMenu.Remove(item);
                        db.SaveChanges();
                    }

                    //List<RolPerfil> ListadoRolesPerfiles = new List<RolPerfil>();

                    foreach (var item in opcionesMenu)
                    {
                        db.PerfilMenu.Add(new PerfilMenu
                        {
                            id_perfil = Perfil.id_perfil,
                            id_menu = item,
                        });
                        db.SaveChanges();
                    }

                    Perfil.nombre_perfil = Perfil.nombre_perfil.ToUpper();
                    db.Entry(Perfil).State = EntityState.Modified;
                    db.SaveChanges();

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

        //Eliminación Lógica
        public static RespuestaTransaccion EliminarPerfil(int id)
        {
            try
            {
                var perfil = db.Perfil.Find(id);

                if (perfil.estado_perfil == true)
                {
                    perfil.estado_perfil = false;
                }
                else
                {
                    perfil.estado_perfil = true;
                }

                db.Entry(perfil).State = EntityState.Modified;
                db.SaveChanges();

                return new RespuestaTransaccion { Estado = true, Respuesta = Mensajes.MensajeTransaccionExitosa };
            }
            catch (Exception ex)
            {
                return new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeTransaccionFallida + " ;" + ex.Message.ToString() };
            }
        }

        public static List<ListadoPerfil> ListarPerfil()
        {
            try
            {
                return db.ListadoPerfil().ToList();
            }
            catch (Exception)
            {
                throw;
            }
        }

        public static Perfil ConsultarPerfil(int id)
        {
            try
            {
                Perfil perfil = db.Perfil.Find(id);
                return perfil;
            }
            catch (Exception)
            {
                throw;
            }
        }

        public static List<int> ListadIdsOpcionesMenuByPerfil(int idPerfil)
        {
            try
            {
                var listadoPerfiles = db.PerfilMenu.Where(s => s.id_perfil == idPerfil).Select(s => s.id_menu).ToList();
                return listadoPerfiles;
            }
            catch (Exception)
            {
                throw;
            }
        }

        public static List<ListadoPerfil> ListarPerfilesPorRol(int rolID)
        {
            List<ListadoPerfil> perfiles = new List<ListadoPerfil>();
            try
            {
                var listadoPerfilesRoles = db.RolPerfil.Where(s => s.id_rol == rolID).Select(s => s.id_perfil).ToList();
                perfiles = db.ListadoPerfil().Where(s => listadoPerfilesRoles.Contains(s.Id)).ToList();
                return perfiles;
            }
            catch (Exception ex)
            {
                return perfiles;
            }
        }

    }
}
using GestionPPM.Entidades.Modelo;
using GestionPPM.Repositorios;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.Entity;
using System.Data.SqlClient;
using System.Linq;
using System.Web;

namespace GestionPPM.Entidades.Metodos
{
    public class ManejoPermisosEntity
    {

        private static readonly GestionPPMEntities db = new GestionPPMEntities();


        public static RespuestaTransaccion CrearActualizarPermisos(List<RolMenuPermiso> Permisos, int createby, int updateby, DateTime createat, DateTime updateat, int rolID, int perfilID)
        {
            using (var transaction = db.Database.BeginTransaction())
            {
                try
                {
                    if(Permisos.Count>=1)
                    {
                        foreach (var item2 in Permisos)
                        {

                            db.LimpiarRolMenuPermisos(item2.RolID, item2.PerfilID, item2.MenuID);
                        }

                        foreach (var item in Permisos)
                        {

                            db.GuardarRolMenuPermisos(item.RolID, item.PerfilID, item.MenuID, item.AccionID, createby, updateby, createat, updateat, item.Estado);
                        }
                    }
                    else
                    {
                        db.LimpiarRolMenuPermisosCompleto(rolID, perfilID);

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

        public static List<UsuarioRolMenuPermiso> ConsultarRolMenuPermiso(int usuario, string controlador)
        {
            List<UsuarioRolMenuPermiso> listado = new List<UsuarioRolMenuPermiso>();
            try
            {
                listado = db.ConsultarUsuarioRolMenuPermiso(usuario, controlador).ToList();
                return listado;
            }
            catch (Exception ex)
            {
                return listado;
            }
        }

        public static List<string> ListadoAccionesCatalogoUsuario(int usuario, string controlador)
        {
            List<string> listado = new List<string>();
            try
            {
                listado = db.ConsultarUsuarioRolMenuPermiso(usuario, controlador).Select(s => s.CodigoCatalogo).ToList();
                return listado;
            }
            catch (Exception ex)
            {
                return listado;
            }
        }

        public static List<UsuarioRolMenuPermiso> ListadoRolMenuPermiso(int idRol, int idPerfil)
        {
            List<UsuarioRolMenuPermiso> listado = new List<UsuarioRolMenuPermiso>();
            try
            {
                listado = db.ListadoRolPerfilMenuPermisos(idRol, idPerfil).ToList();
                return listado;
            }
            catch (Exception ex)
            {
                return listado;
            }
        }




    }
}
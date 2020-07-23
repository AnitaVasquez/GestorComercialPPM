using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Web;
using System.Web.UI.WebControls;
using System.Configuration;
using System.Web.Security;
using GestionPPM.Entidades.Metodos;
using GestionPPM.Entidades.Modelo;
using System.Web.Mvc;
using GestionPPM.Repositorios;
using static GestionPPM.Repositorios.Auxiliares;

namespace Seguridad.Helper
{
    public class SessionHelper
    {
        public static bool ExistUserInSession()
        {
            return HttpContext.Current.User.Identity.IsAuthenticated;
        }
        public static string GetUsuarioSession()
        {
            if (!(HttpContext.Current.Session["Usuario"] is string usuario))
                return "";
            else
                return usuario;
        }

        public static string GetPermisosOpciones()
        {
            if (!(HttpContext.Current.Session["PermisosOpcion"] is string permisos))
                return "";
            else
                return permisos;
        }


        public static void DestroyUserSession()
        {
            HttpContext.Current.Session["Usuario"] = null;
            FormsAuthentication.SignOut();
        }
        public static int GetUser()
        {
            int user_id = 0;
            if (HttpContext.Current.User != null && HttpContext.Current.User.Identity is FormsIdentity)
            {
                FormsAuthenticationTicket ticket = ((FormsIdentity)HttpContext.Current.User.Identity).Ticket;
                if (ticket != null)
                {
                    user_id = Convert.ToInt32(ticket.UserData);
                }
            }
            return user_id;
        }
        public static void AddUserToSession(string id)
        {
            bool persist = true;
            var cookie = FormsAuthentication.GetAuthCookie("usuario", persist);

            cookie.Name = FormsAuthentication.FormsCookieName;
            cookie.Expires = DateTime.Now.AddMonths(3);

            var ticket = FormsAuthentication.Decrypt(cookie.Value);
            var newTicket = new FormsAuthenticationTicket(ticket.Version, ticket.Name, ticket.IssueDate, ticket.Expiration, ticket.IsPersistent, id);

            cookie.Value = FormsAuthentication.Encrypt(newTicket);
            HttpContext.Current.Response.Cookies.Add(cookie);
        }
        public static List<UsuarioRolMenuPermisoR> ObtenerRolMenuPermisos(string nombreControlador)
        {
            List<UsuarioRolMenuPermisoR> listadoFinal = new List<UsuarioRolMenuPermisoR>();
            List<UsuarioRolMenuPermiso> listado = new List<UsuarioRolMenuPermiso>();
            try
            {
                var user = System.Web.HttpContext.Current.Session["usuario"];

                if (user == null)
                    return listadoFinal;

                int usuario = int.Parse(user.ToString());

                listado = ManejoPermisosEntity.ConsultarRolMenuPermiso(usuario, nombreControlador);

                foreach (var item in listado)
                {
                    UsuarioRolMenuPermisoR tmp = new UsuarioRolMenuPermisoR();
                    tmp.IDRolMenuPermiso = item.IDRolMenuPermiso;
                    tmp.RolID = item.RolID;
                    tmp.NombreRol = item.NombreRol;
                    tmp.PerfilID = item.PerfilID;
                    tmp.NombrePerfil = item.NombrePerfil;
                    tmp.MenuID = item.MenuID;
                    tmp.NombreMenu = item.NombreMenu;
                    tmp.EnlaceMenu = item.EnlaceMenu;
                    tmp.MenuPadre = item.MenuPadre;
                    tmp.IDCatalogo = item.IDCatalogo;
                    tmp.CodigoCatalogo = item.CodigoCatalogo;
                    tmp.TextoCatalogoAccion = item.TextoCatalogoAccion;
                    tmp.CreadoPorID = item.CreadoPorID;
                    tmp.CreadoPor = item.CreadoPor;
                    tmp.ActualizadoPorID = item.ActualizadoPorID;
                    tmp.ActualizadoPor = item.ActualizadoPor;
                    tmp.CreatedAt = item.CreatedAt;
                    tmp.UpdatedAt = item.UpdatedAt;
                    tmp.Estado = item.Estado;
                    tmp.MetodoControlador = item.MetodoControlador;
                    tmp.NombreControlador = item.NombreControlador;
                    tmp.AccionEnlace = item.AccionEnlace;

                    listadoFinal.Add(tmp);
                }

                return listadoFinal;
            }
            catch (Exception ex)
            {
                return listadoFinal;
            }
        }


    }
}
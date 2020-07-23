
using GestionPPM.Entidades.Metodos;
using GestionPPM.Entidades.Modelo;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.Web.Routing;

namespace Seguridad.Helper
{
    // Si no estamos logeado, regresamos al login
    public class AutenticadoAttribute : ActionFilterAttribute
    {
        //private SeguridadEntities db = new SeguridadEntities();

        public override void OnActionExecuting(ActionExecutingContext filterContext)
        {
            base.OnActionExecuting(filterContext);

            // Si el usuario no ha iniciado sesión o la sesión ya no está activa
            if (string.IsNullOrEmpty(SessionHelper.GetUsuarioSession()))
            {
                filterContext.Result = new RedirectToRouteResult(new RouteValueDictionary(new
                {
                    controller = "Ingreso",
                    action = "Login"
                }));

                //Almacenar en una variable de sesion
                HttpContext.Current.Session["Resultado"] = "Su sesión ha caducado";
                HttpContext.Current.Session["Estado"] = "True";

            }
            else
            {
                var usuarioSesionID = int.Parse(HttpContext.Current.Session["usuario"] as string);

                var tipoRequestControlador = filterContext.RequestContext.HttpContext.Request.AcceptTypes.ToList();
                bool flag = false;

                if (tipoRequestControlador.Any())
                {
                    flag = tipoRequestControlador.Contains("application/json");
                }

                string controlador = filterContext.ActionDescriptor.ControllerDescriptor.ControllerName;
                string accion = filterContext.ActionDescriptor.ActionName;
                List<UsuarioRolMenuPermiso> listado = new List<UsuarioRolMenuPermiso>();
                listado = ManejoPermisosEntity.ConsultarRolMenuPermiso(usuarioSesionID, controlador);

                int rolid = UsuarioEntity.ConsultarUsuario(usuarioSesionID).id_rol;
                //var perfiles = PerfilesEntity.ListarPerfilesPorRol(rolid);

                var ruta = MenuEntity.ConsultarRutaMenu(rolid);


                bool permisoOK = listado.Any(s => s.MetodoControlador == accion);
                bool RutaOK = ruta.Any(s => s.RUTAMENU == controlador);
                bool permisoIndex = false;

                if (RutaOK)
                {
                    permisoIndex = true;
                }

                //if (!permisoOK)
                //{
                //    RutaOK = false;
                //}

                if (!permisoOK && !RutaOK && !accion.Equals("OrdenMenu") && !accion.Equals("CambiarClave") && !accion.Equals("_Create") && !accion.Equals("_GetContactosFacturacion") && !accion.Equals("GetDependientesEtapaCliente") && !accion.Equals("GetSublineaDependientesLineaNegocio") && !accion.Equals("GetDependientesClienteContactos")  && !controlador.Equals("Solicitud") && !controlador.Equals("MatrizPresupuesto") && !flag && !controlador.Equals("Home") && !controlador.Equals("ManejoPermisos"))
                {
                    filterContext.Result = new RedirectToRouteResult(new RouteValueDictionary(new
                    {
                        controller = "Error",
                        action = "NotForbbiden"
                    }));
                }

            }
        }
    }

    // Si estamos logeado ya no podemos acceder a la página de Login
    public class NoLoginAttribute : ActionFilterAttribute
    {
        public override void OnActionExecuting(ActionExecutingContext filterContext)
        {
            base.OnActionExecuting(filterContext);

            if (!string.IsNullOrEmpty(SessionHelper.GetUsuarioSession()))
            {
                filterContext.Result = new RedirectToRouteResult(new RouteValueDictionary(new
                {
                    controller = "Home",
                    action = "Index"
                }));
            }
        }
    }
}
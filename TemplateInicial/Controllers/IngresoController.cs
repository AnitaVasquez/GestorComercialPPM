using GestionPPM.Entidades.Metodos;
using GestionPPM.Entidades.Modelo;
using GestionPPM.Repositorios;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace TemplateInicial.Controllers
{
    public class IngresoController : Controller
    {
        // GET: Ingreso
        public ActionResult Index()
        {

            var respuesta = System.Web.HttpContext.Current.Session["Resultado"] as string;
            var estado = System.Web.HttpContext.Current.Session["Estado"] as string;

            ViewBag.Resultado = respuesta;
            ViewBag.Estado = estado;

            Session["Resultado"] = "";
            Session["Estado"] = "";

            return View();
        }

        // GET: Ingreso/Details/5
        public ActionResult Details(int id)
        {
            return View();
        }

        // GET: Ingreso/Create
        public ActionResult Login()
        {
            var respuesta = System.Web.HttpContext.Current.Session["Resultado"] as string;
            var estado = System.Web.HttpContext.Current.Session["Estado"] as string;

            ViewBag.Resultado = respuesta;
            ViewBag.Estado = estado;

            Session["Resultado"] = "";
            Session["Estado"] = "";

            return View();
        }

        // POST: Ingreso/Create
        [HttpPost]
        public ActionResult Login(LoginViewModel model, string returnUrl)
        {
            var clave = "";
            if (model.Password != null)
            {
                // chequear el usuario
                clave = GestionPPM.Repositorios.Validaciones.GetMD5_1(model.Password);
            }

            var authenticatedUser = UsuarioEntity.ValidarUsuario(model.Login, clave, model.Password); 

            if (authenticatedUser != null)
            {
                System.Web.HttpContext.Current.Session["usuario"] = authenticatedUser.Código.ToString();
                System.Web.HttpContext.Current.Session["nombre"] = authenticatedUser.Nombre.ToString();

                System.Web.HttpContext.Current.Session["rolID"] = authenticatedUser.rol_id.Value;

                //var rolID = Convert.ToInt32(System.Web.HttpContext.Current.Session["rolID"]);
                //ViewBag.PerfilesUsuario = PerfilesEntity.ListarPerfilesPorRol(authenticatedUser.rol_id.Value);

                var nombre = ViewData["nombre"] = System.Web.HttpContext.Current.Session["nombre"] as String;

                var perfiles = PerfilesEntity.ListarPerfilesPorRol(authenticatedUser.rol_id.Value);

                Session["UsuarioUsername"] = authenticatedUser.Nombre + " " + authenticatedUser.Apellido;
                Session["UsuarioEmail"] = authenticatedUser.mail_usuario;
                Session["UsuarioCargo"] = authenticatedUser.Cargo;
                System.Web.HttpContext.Current.Session["usuarioClave"] = model.Password;

                Session["ValidacionRolCreacionSAFI"] = authenticatedUser.descripcion_rol.Equals("No puede modificar la creación SAFI") ? false : true;  

                if(authenticatedUser.reset_clave == false)
                {
                    return RedirectToAction("Index", "Home");
                }
                else
                {
                    return RedirectToAction("CambiarClave/"+ authenticatedUser.Código, "Ingreso");
                }

            }
            else
            {
                ViewBag.Resultado = Mensajes.MensajeCredencialesIncorrectas;
                ViewBag.Estado = "False"; 
                return View(model);
            }
        }

        // GET: Ingreso/Edit/5
        public ActionResult Registrarse()
        {
            var respuesta = System.Web.HttpContext.Current.Session["Resultado"] as string;
            var estado = System.Web.HttpContext.Current.Session["Estado"] as string;

            ViewBag.Resultado = respuesta;
            ViewBag.Estado = estado;

            Session["Resultado"] = "";
            Session["Estado"] = "";
             
            return View();
        }

        // POST: Ingreso/Edit/5
        [HttpPost]
        public ActionResult Registrarse(RegisterViewModel registro)
        {
            try
            {
                //validar contraseña 
                if(registro.Password == registro.ConfirmPassword)
                {
                    //realizar el ingreso del usuario
                    RespuestaTransaccion resultado = UsuarioEntity.CrearUsuarioGenerico(registro);

                    //Almacenar en una variable de sesion
                    Session["Resultado"] = resultado.Respuesta;
                    Session["Estado"] = resultado.Estado.ToString();

                    ViewBag.Resultado = resultado.Respuesta.ToString();
                    ViewBag.Estado = resultado.Estado.ToString();
                     
                    if (resultado.Estado == true)
                    {
                    return RedirectToAction("Login", "Ingreso");
                    }
                    else
                    {
                        return View(registro);
                    }
                }
                else
                { 
                    ViewBag.Resultado = Mensajes.MensajeValidacionContrasenias;
                    ViewBag.Estado = "false";

                    Session["Resultado"] = "";
                    Session["Estado"] = ""; 
                    
                    return View(registro);
                }

            }
            catch(Exception e)
            {
                e.ToString();
                return View(registro);
            }
        }

        // GET: Ingreso/Delete/5
        public ActionResult RecuperarClave()
        {
            var respuesta = System.Web.HttpContext.Current.Session["Resultado"] as string;
            var estado = System.Web.HttpContext.Current.Session["Estado"] as string;

            ViewBag.Resultado = respuesta;
            ViewBag.Estado = estado;

            Session["Resultado"] = "";
            Session["Estado"] = "";

            return View();
        }

        // POST: Ingreso/Delete/5
        [HttpPost]
        public ActionResult RecuperarClave(ForgotViewModel recuperar)
        {
            try
            {
                RespuestaTransaccion resultado = UsuarioEntity.RecuperarClave(recuperar);

                //Almacenar en una variable de sesion
                Session["Resultado"] = resultado.Respuesta;
                Session["Estado"] = resultado.Estado.ToString();

                ViewBag.Resultado = resultado.Respuesta.ToString();
                ViewBag.Estado = resultado.Estado.ToString();

                if (resultado.Estado == true)
                {
                    return RedirectToAction("Login", "Ingreso");
                }
                else
                {
                    return View(recuperar);
                }
            }
            catch
            {
                return View();
            }
        }

        // GET: Cambiar Clave
        public ActionResult CambiarClave(int id)
        {
            CambiarClave cambiaClave = new CambiarClave();
            cambiaClave.UsuaCodi = id;
            return View(cambiaClave);
        }

        [HttpPost]
        public ActionResult CambiarClave(CambiarClave usuario)
        {
            RespuestaTransaccion resultado = UsuarioEntity.CambiarClaveReset(usuario);

            //Almacenar en una variable de sesion
            Session["Resultado"] = resultado.Respuesta;
            Session["Estado"] = resultado.Estado.ToString();

            if (resultado.Estado.ToString() == "True")
            {
                return RedirectToAction("Login", "Ingreso");
            }
            else
            {
                ViewBag.Resultado = resultado.Respuesta;
                ViewBag.Estado = resultado.Estado.ToString();
                Session["Resultado"] = "";
                Session["Estado"] = "";
                return View(usuario);
            }
        }
    }
}

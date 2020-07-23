using GestionPPM.Entidades.Metodos;
using NLog;
using Seguridad.Helper;
using System;
using System.Configuration;
using System.IO;
using System.Web.Mvc;

namespace TemplateInicial.Controllers
{
 

    [Autenticado]
    public class HomeController : Controller
    {
        private static readonly Logger Log = LogManager.GetCurrentClassLogger();
        [HttpGet]
        public ActionResult Index()
        {
            return View();  
        }

        [HttpGet]
        public ActionResult AnotherLink()
        {
            return View("Index");
        }

        public ActionResult Menu()
        {
            //Obtener el codigo de usuario que se logeo
            var user = ViewData["usuario"] = System.Web.HttpContext.Current.Session["usuario"];

            //Obtener listado de opciones del menu
            var list = UsuarioEntity.OpcionesMenuUsuario(Convert.ToInt16(user));
              
            return PartialView("_Menu", list);
        }

        public ActionResult SinPermisos(string tipoRenderizacion = null)
        {

            if (string.IsNullOrEmpty(tipoRenderizacion))
            {
                ViewBag.Layout = "~/Views/Shared/_Layout.cshtml";
                return PartialView("~/Views/Shared/_SinPermisos.cshtml");
            }
            else
            {
                ViewBag.Layout = null;
                return PartialView("~/Views/Shared/_SinPermisos.cshtml");
            }
        }

        public ActionResult DescargarArchivo(string path)
        {
            try
            {
                byte[] fileBytes = System.IO.File.ReadAllBytes(path);
                string fileName = Path.GetFileName(path);
                return File(fileBytes, System.Net.Mime.MediaTypeNames.Application.Octet, fileName);
            }
            catch (Exception ex)
            {
                string mensaje = "Verifique que el servidor tenga instaladas las dependencias necesarias que requiere la aplicación, revisar archivo OfficeToPDF ({0})";
                ViewBag.Excepcion = string.Format(mensaje, ex.Message.ToString());
                return View("~/Views/Error/InternalServerError.cshtml");
            }
        }

    }
}

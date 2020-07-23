using GestionPPM.Entidades.Metodos;
using GestionPPM.Entidades.Modelo;
using GestionPPM.Repositorios;
using OfficeOpenXml;
using Omu.Awem.Helpers;
using Seguridad.Helper;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks; 
using System.Web.Mvc;

namespace TemplateInicial.Controllers
{
    [Autenticado]
    public partial class RolPerfiles
    {
        public RolPerfiles() {
            idsPerfiles_guardar = new List<int>();
            idsPerfiles_editar = new List<int>();
        }
        public int id_rol { get; set; }
        public string nombre_rol { get; set; }
        public string descripcion_rol { get; set; }
        public Nullable<bool> estado_rol { get; set; }
        public List<int> idsPerfiles_guardar { get; set; }
        public List<int> idsPerfiles_editar { get; set; }

    }
    public class RolController : BaseAppController
    {
        private const string XlsxContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

        // GET: Rol
        public ActionResult Index()
        {
            var respuesta = System.Web.HttpContext.Current.Session["Resultado"] as string;
            var estado = System.Web.HttpContext.Current.Session["Estado"] as string;

            ViewBag.Resultado = respuesta;
            ViewBag.Estado = estado;

            Session["Resultado"] = "";
            Session["Estado"] = "";
            //Obtener Ruta PDF
            string path = string.Empty;
            string controllerName = this.ControllerContext.RouteData.Values["controller"].ToString();
            path = "../AdjuntosManual/" + controllerName + ".pdf";

            var absolutePath = HttpContext.Server.MapPath(path);
            bool rutaArchivo = System.IO.File.Exists(absolutePath);

            if (!rutaArchivo)
            {
                string path1 = "../AdjuntosManual/ManualUsuario.pdf";
                ViewBag.Iframe = path1;
            }
            else
            {
                ViewBag.Iframe = path;
            }

            return View();
        }

        [HttpGet]
        public async Task<PartialViewResult> IndexGrid(String search)
        {
            ViewBag.NombreListado = Etiquetas.TituloGridRol;

            //Controlar permisos
            var user = ViewData["usuario"] = System.Web.HttpContext.Current.Session["usuario"];
            var usuario = int.Parse(user.ToString());
            string nombreControlador = ControllerContext.RouteData.Values["controller"].ToString();
            ViewBag.NombreControlador = nombreControlador;

            ViewBag.AccionesUsuario = ManejoPermisosEntity.ListadoAccionesCatalogoUsuario(usuario, nombreControlador);

            //Obtener Acciones del controlador
            ViewBag.AccionesControlador = GetMetodosControlador(nombreControlador);

            var listado = RolEntity.ListarRol();

            search = !string.IsNullOrEmpty(search) ? search.Trim() : "";

            if (!string.IsNullOrEmpty(search))//filter
            {
                var type = listado.GetType().GetGenericArguments()[0];
                var properties = type.GetProperties();

                listado = listado.Where(x => properties
                            .Any(p =>
                            {
                                var value = p.GetValue(x);
                                return value != null && value.ToString().ToLower().Contains(search.ToLower());
                            })).ToList();
            }

            // Only grid query values will be available here.
            return PartialView("_IndexGrid", await Task.Run(() => listado));
        }

        public async Task<ActionResult> Details(int? id)
        {
            if (id == null)
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);

            var obj = RolEntity.ConsultarRol(id.Value);

            if (obj == null)
                return HttpNotFound();
            else
                return View(obj);
        }

        public ActionResult Create()
        {
            return View();
        }        

        [HttpPost]
        public ActionResult Create(RolPerfiles rol, List<int> perfiles)
        {
            try
            {

                string nombreRol = (rol.nombre_rol ?? string.Empty).ToLower().Trim();

                var validacionNombreRolUnico = RolEntity.ListarRol().Where(s => (s.Nombre ?? string.Empty).ToLower().Trim() == nombreRol).ToList();

                if (validacionNombreRolUnico.Count > 0)
                    return Json(new { Resultado = new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeValidacionNombreRol } }, JsonRequestBehavior.AllowGet);

                RespuestaTransaccion resultado = RolEntity.CrearRol(new Rol { nombre_rol = rol.nombre_rol, descripcion_rol = rol.descripcion_rol }, perfiles);


                return Json(new { Resultado = resultado }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                return Json(new { Resultado = new RespuestaTransaccion { Estado = false, Respuesta = ex.Message } }, JsonRequestBehavior.AllowGet);
            }
        }

        public ActionResult Edit(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            var rol = RolEntity.ConsultarRol(id.Value);

            var perfiles = RolEntity.ListadIdsPerfilesByRol(id.Value);
            ViewBag.idsPerfilesRol = string.Join(",", perfiles);

            if (rol == null)
            {
                return HttpNotFound();
            }
            return View(rol);
        }

        [HttpPost]
        public ActionResult Edit(RolPerfiles rol, List<int> perfiles)
        {
            try
            {


                string nombreRol = (rol.nombre_rol ?? string.Empty).ToLower().Trim();

                var validacionNombreRolUnico = RolEntity.ListarRol().Where(s => (s.Nombre ?? string.Empty).ToLower().Trim() == nombreRol && s.Id != rol.id_rol).ToList();

                if (validacionNombreRolUnico.Count > 0)
                    return Json(new { Resultado = new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeValidacionNombreRol } }, JsonRequestBehavior.AllowGet);

                RespuestaTransaccion resultado = RolEntity.ActualizarRol(new Rol { id_rol = rol.id_rol, nombre_rol = rol.nombre_rol, descripcion_rol = rol.descripcion_rol, estado_rol = rol.estado_rol }, perfiles);

                return Json(new { Resultado = resultado }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                return Json(new { Resultado = new RespuestaTransaccion { Estado = false, Respuesta = ex.Message } }, JsonRequestBehavior.AllowGet);
            }
        }

        [HttpPost]
        public ActionResult Eliminar(int id)
        {
            if (id == 0)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            RespuestaTransaccion resultado = RolEntity.EliminarRol(id);// await db.Cabecera.FindAsync(id);

            //return RedirectToAction("Index");
            return Json(new { Resultado = resultado }, JsonRequestBehavior.AllowGet);

        }

        public JsonResult GetPerfiles(string searchTerm)
        {
            var items = PerfilesEntity.ListarPerfil()
                .Select(o => new Oitem(o.Id, o.Nombre));

            return Json(items);
        }

        public JsonResult _GetPerfiles()
        {
            List<MultiSelectJQueryUi> items = new List<MultiSelectJQueryUi>();

            items = PerfilesEntity.ListarPerfil()
.Select(o => new MultiSelectJQueryUi(o.Id, o.Nombre, o.Descripcion)).ToList();
            return Json(items, JsonRequestBehavior.AllowGet);
        }

        public ActionResult DescargarReporteFormatoExcel()
        {
            //Seleccionar las columnas a exportar
            var collection = RolEntity.ListarRol();
            var package = new ExcelPackage();

            package = Reportes.ExportarExcel(collection.Cast<object>().ToList(),"Roles");
            return File(package.GetAsByteArray(), XlsxContentType, "Listado_Roles.xlsx");
        }

        public ActionResult DescargarReporteFormatoPDF()
        {
            // Seleccionar las columnas a exportar
            var results = RolEntity.ListarRol();

            var list = Reportes.SerializeToJSON(results);
            return Content(list, "application/json");
        }

        public ActionResult DescargarReporteFormatoCSV()
        {
            var comlumHeadrs = new string[]
            {
                "ID",
                "NOMBRE",
                "DESCRIPCION",
                "ESTADO"
            };

            var listado = (from item in RolEntity.ListarRol()
                           select new object[]
                           {
                                            item.Id,
                                            item.Nombre,
                                            item.Descripcion,
                                            item.Estado
                           }).ToList();

            // Build the file content
            var employeecsv = new StringBuilder();
            listado.ForEach(line =>
            {
                employeecsv.AppendLine(string.Join(",", line));
            });

            byte[] buffer = Encoding.Default.GetBytes($"{string.Join(",", comlumHeadrs)}\r\n{employeecsv.ToString()}");
            return File(buffer, "text/csv", $"Roles.csv");
        }

    }

}


using System;
using System.Collections.Generic;
using System.Data; 
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks; 
using System.Web.Mvc;
using GestionPPM.Entidades.Metodos;
using GestionPPM.Entidades.Modelo;
using GestionPPM.Repositorios;
using OfficeOpenXml;
using Omu.Awem.Helpers;
using Seguridad.Helper;

namespace TemplateInicial.Controllers
{

    public partial class PerfilesOpcionesMenu
    {
        public PerfilesOpcionesMenu()
        {
            idsOpcionesMenu_guardar = new List<int>();
            idsOpcionesMenu_editar = new List<int>();
        }
        public int id_perfil { get; set; }
        public string nombre_perfil { get; set; }
        public string descripcion_perfil { get; set; }
        public Nullable<bool> estado_perfil { get; set; }
        public List<int> idsOpcionesMenu_guardar { get; set; }
        public List<int> idsOpcionesMenu_editar { get; set; }

    }

    [Autenticado]
    public class PerfilesController : BaseAppController
    {
        private const string XlsxContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

        // GET: Perfiles
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
            ViewBag.NombreListado = Etiquetas.TituloGridPerfil;
            //Controlar permisos
            var user = ViewData["usuario"] = System.Web.HttpContext.Current.Session["usuario"];
            var usuario = int.Parse(user.ToString());
            string nombreControlador = ControllerContext.RouteData.Values["controller"].ToString();
            ViewBag.NombreControlador = nombreControlador;

            ViewBag.AccionesUsuario = ManejoPermisosEntity.ListadoAccionesCatalogoUsuario(usuario, nombreControlador);

            //Obtener Acciones del controlador
            ViewBag.AccionesControlador = GetMetodosControlador(nombreControlador);

            //Búsqueda

            var listado = PerfilesEntity.ListarPerfil();

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

        // GET: Perfiles/Details/5
        public ActionResult Details(int? id)
        {
            if (id == null)
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);

            var obj = PerfilesEntity.ConsultarPerfil(id.Value);

            if (obj == null)
                return HttpNotFound();
            else
                return View(obj);
        }

        // GET: Perfiles/Create
        public ActionResult Create()
        {
            return View();
        }

        [HttpPost]
        public ActionResult Create(PerfilesOpcionesMenu perfil, List<int> opcionesMenu)
        {
            try
            {
                string nombrePerfil = (perfil.nombre_perfil ?? string.Empty).ToLower().Trim();

                var validacionNombreUnico = PerfilesEntity.ListarPerfil().Where(s => (s.Nombre ?? string.Empty).ToLower().Trim() == nombrePerfil).ToList();

                if (validacionNombreUnico.Count > 0)
                    return Json(new { Resultado = new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeResgistroExistente } }, JsonRequestBehavior.AllowGet);



                RespuestaTransaccion resultado = PerfilesEntity.CrearPerfil(new Perfil { nombre_perfil = perfil.nombre_perfil, descripcion_perfil = perfil.descripcion_perfil }, opcionesMenu);

                return Json(new { Resultado = resultado }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                return Json(new { Resultado = new RespuestaTransaccion { Estado = false, Respuesta = ex.Message } }, JsonRequestBehavior.AllowGet);
            }
        }

        // GET: Perfiles/Edit/5
        public ActionResult Edit(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            var rol = PerfilesEntity.ConsultarPerfil(id.Value);

            var opcionesMenu = PerfilesEntity.ListadIdsOpcionesMenuByPerfil(id.Value);
            ViewBag.idsPerfilesOpcionesMenu = string.Join(",", opcionesMenu);  //ContactoClienteEntity.ListadIdsContactosClientesByCliente(id.Value);

            if (rol == null)
            {
                return HttpNotFound();
            }
            return View(rol);
        }

        // POST: Perfiles/Edit/5
        // To protect from overposting attacks, please enable the specific properties you want to bind to, for 
        // more details see https://go.microsoft.com/fwlink/?LinkId=317598.
        //[HttpPost]
        //[ValidateAntiForgeryToken]
        //public ActionResult Edit([Bind(Include = "id_perfil,nombre_perfil,descripcion_perfil,estado_perfil, idsOpcionesMenu_editar")] PerfilesOpcionesMenu perfil)
        //{
        //    if (ModelState.IsValid)
        //    {
        //        string nombrePerfil = (perfil.nombre_perfil ?? string.Empty).ToLower().Trim();

        //        var validacionNombreUnico = PerfilesEntity.ListarPerfil().Where(s => (s.nombre_perfil ?? string.Empty).ToLower().Trim() == nombrePerfil && s.id_perfil != perfil.id_perfil).ToList();

        //        if (validacionNombreUnico.Count > 0)
        //        {

        //            List<RespuestaTransaccion> validacionesRespuesta = new List<RespuestaTransaccion>{
        //            new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeResgistroExistente }
        //            };
        //            ViewBag.idsPerfilesOpcionesMenu = PerfilesEntity.ListadIdsOpcionesMenuByPerfil(perfil.id_perfil);
        //            //Listado de validaciones especificas de la entidad
        //            ViewBag.Resultado = validacionesRespuesta;

        //            return View();
        //        }

        //        RespuestaTransaccion resultado = PerfilesEntity.ActualizarPerfil(new Perfil { id_perfil = perfil.id_perfil, estado_perfil = perfil.estado_perfil, nombre_perfil = perfil.nombre_perfil, descripcion_perfil = perfil.descripcion_perfil }, perfil.idsOpcionesMenu_editar);

        //        //Almacenar en una variable de sesion
        //        Session["Resultado"] = resultado.Respuesta;
        //        Session["Estado"] = resultado.Estado.ToString();

        //        if (resultado.Estado.ToString() == "True")
        //        {
        //            return RedirectToAction("Index");
        //        }
        //        else
        //        {
        //            ViewBag.Resultado = resultado.Respuesta;
        //            ViewBag.Estado = resultado.Estado.ToString();
        //            return View(perfil);
        //        }

        //        //return Json(new { Resultado = resultado }, JsonRequestBehavior.AllowGet);
        //    }

        //    return View(perfil);
        //}
        [HttpPost]
        public ActionResult Edit(PerfilesOpcionesMenu perfil, List<int> opcionesMenu)
        {
            try
            {

                string nombrePerfil = (perfil.nombre_perfil ?? string.Empty).ToLower().Trim();

                var validacionNombreUnico = PerfilesEntity.ListarPerfil().Where(s => (s.Nombre ?? string.Empty).ToLower().Trim() == nombrePerfil && s.Id != perfil.id_perfil).ToList();

                if (validacionNombreUnico.Count > 0)
                    return Json(new { Resultado = new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeResgistroExistente } }, JsonRequestBehavior.AllowGet);

                    RespuestaTransaccion resultado = PerfilesEntity.ActualizarPerfil(new Perfil { id_perfil = perfil.id_perfil, estado_perfil = perfil.estado_perfil, nombre_perfil = perfil.nombre_perfil, descripcion_perfil = perfil.descripcion_perfil }, opcionesMenu);

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
            RespuestaTransaccion resultado = PerfilesEntity.EliminarPerfil(id);// await db.Cabecera.FindAsync(id);

            //return RedirectToAction("Index");
            return Json(new { Resultado = resultado }, JsonRequestBehavior.AllowGet);
        }

        public JsonResult GetOpcionesMenu(string searchTerm)
        {
            var items = MenuEntity.ListarMenuHijos()
                .Select(o => new Oitem(o.id_menu, o.nombre_menu));

            return Json(items);
        }

        public JsonResult _GetOpcionesMenu()
        {
            List<MultiSelectJQueryUi> items = new List<MultiSelectJQueryUi>();

            items = MenuEntity.ListarMenuHijos()
.Select(o => new MultiSelectJQueryUi(o.id_menu, o.nombre_menu, o.nombre_pagina_menu)).ToList();
            return Json(items, JsonRequestBehavior.AllowGet);
        }

        public ActionResult DescargarReporteFormatoExcel()
        {
            //Seleccionar las columnas a exportar
            var collection = PerfilesEntity.ListarPerfil();
            var package = new ExcelPackage();

            package = Reportes.ExportarExcel(collection.Cast<object>().ToList(),"Perfiles");
            return File(package.GetAsByteArray(), XlsxContentType, "Listado_Perfiles.xlsx");
        }

        public ActionResult DescargarReporteFormatoPDF()
        {
            // Seleccionar las columnas a exportar
            var results = PerfilesEntity.ListarPerfil();

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

            var listado = (from item in PerfilesEntity.ListarPerfil()
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
            return File(buffer, "text/csv", $"Perfiles.csv");
        }

    }
}

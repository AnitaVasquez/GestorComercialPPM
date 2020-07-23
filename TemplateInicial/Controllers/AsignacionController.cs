using GestionPPM.Entidades.Metodos;
using GestionPPM.Entidades.Modelo;
using GestionPPM.Repositorios;
using OfficeOpenXml;
using OfficeOpenXml.Style;
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
    public partial class AsignacionUsuarios
    {
        public AsignacionUsuarios() {
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
    public class AsignacionController : BaseAppController
    {
        private const string XlsxContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
         
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

            return View();
        }

        [HttpGet]
        public async Task<PartialViewResult> IndexGrid(String search)
        {
            ViewBag.NombreListado = Etiquetas.TituloGridAsignacion;

            //Controlar permisos
            var user = ViewData["usuario"] = System.Web.HttpContext.Current.Session["usuario"];
            var usuario = int.Parse(user.ToString());
            string nombreControlador = ControllerContext.RouteData.Values["controller"].ToString();
            ViewBag.NombreControlador = nombreControlador;

            ViewBag.AccionesUsuario = ManejoPermisosEntity.ListadoAccionesCatalogoUsuario(usuario, nombreControlador);

            //Obtener Acciones del controlador
            ViewBag.AccionesControlador = GetMetodosControlador(nombreControlador);
            var listado = AsignacionEntity.ListarAsignacion();

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
            //Listado de Tipo 
            var Tipo = CatalogoEntity.ObtenerListadoCatalogosByCodigo("TSL-01");
            ViewBag.ListadoTipo = Tipo;

            //Listado Subtipo
            var Subtipo = CatalogoEntity.ObtenerListadoCatalogosByCodigo("TSL-01");
            ViewBag.ListadoSubtipo = Subtipo;

            return View();
        }        

        [HttpPost]
        public ActionResult Create(AsignacionSolicitudes asignacion, List<int> usuarios)
        {
            try
            {
                //Listado de Tipo 
                var Tipo = CatalogoEntity.ObtenerListadoCatalogosByCodigo("TSL-01");
                ViewBag.ListadoTipo = Tipo;

                //Listado Subtipo
                var Subtipo = CatalogoEntity.ObtenerListadoCatalogosByCodigo("TSL-01-SUBTIPO"); 
                ViewBag.ListadoSubtipo = Subtipo;

                //Validar que no exista
                var validacionUnicidad = AsignacionEntity.ObtenerListadoAsignados().Where(a => a.id_tipo == asignacion.id_tipo && a.id_subtipo == asignacion.id_subtipo).ToList();

                if (validacionUnicidad.Count > 0)
                {
                    return Json(new { Resultado = new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeResgistroExistente } }, JsonRequestBehavior.AllowGet);
                }
                else
                {
                    RespuestaTransaccion resultado = AsignacionEntity.CrearAsignacion(asignacion, usuarios);
                    return Json(new { Resultado = resultado }, JsonRequestBehavior.AllowGet);
                }                 
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

            var asignacion = AsignacionEntity.ConsultarAsignacion(id.Value);

            //Listado de Tipo 
            var Tipo = CatalogoEntity.ObtenerListadoCatalogosByCodigo("TSL-01");
            ViewBag.ListadoTipo = Tipo;

            //Listado Subtipo  
            var Subtipo = CatalogoEntity.ConsultarCatalogoPorPadre(asignacion.id_tipo.HasValue ? asignacion.id_tipo.Value : 0, "SUBTIPO");
            ViewBag.ListadoSubtipo = Subtipo;
             
            var usuarios = AsignacionEntity.ListadIdsUsuariosByAsignacion(id.Value);
            ViewBag.idsUsuarios = string.Join(",", usuarios);

            if (asignacion == null)
            {
                return HttpNotFound();
            }
            return View(asignacion);
        }

        [HttpPost]
        public ActionResult Edit(AsignacionSolicitudes asignacion, List<int> usuarios)
        {
            try
            {
                //Listado de Tipo 
                var Tipo = CatalogoEntity.ObtenerListadoCatalogosByCodigo("TSL-01");
                ViewBag.ListadoTipo = Tipo;

                //Listado Subtipo  
                var Subtipo = CatalogoEntity.ConsultarCatalogoPorPadre(asignacion.id_tipo.HasValue ? asignacion.id_tipo.Value : 0, "SUBTIPO"); 
                ViewBag.ListadoSubtipo = Subtipo;

                //Validar que no exista
                var validacionUnicidad = AsignacionEntity.ObtenerListadoAsignados().Where(a => a.id_tipo == asignacion.id_tipo && a.id_subtipo == asignacion.id_subtipo && a.id_asignacion_solicitudes != asignacion.id_asignacion_solicitudes).ToList();

                if (validacionUnicidad.Count > 0)
                {
                    return Json(new { Resultado = new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeResgistroExistente } }, JsonRequestBehavior.AllowGet);
                }
                else
                {
                    RespuestaTransaccion resultado = AsignacionEntity.ActualizarAsignacion(asignacion, usuarios);
                    return Json(new { Resultado = resultado }, JsonRequestBehavior.AllowGet);
                } 
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
            RespuestaTransaccion resultado = AsignacionEntity.EliminarAsignacion(id); 

            //return RedirectToAction("Index");
            return Json(new { Resultado = resultado }, JsonRequestBehavior.AllowGet);

        }

        public JsonResult GetPerfiles(string searchTerm)
        {
            var items = PerfilesEntity.ListarPerfil()
                .Select(o => new Oitem(o.Id, o.Nombre));

            return Json(items);
        }

        public JsonResult _GetUsuarios()
        {
            List<MultiSelectJQueryUi> items = new List<MultiSelectJQueryUi>();

            items = UsuarioEntity.ListarUsuariosInternos()
            .Select(o => new MultiSelectJQueryUi(o.Id, o.Nombres_Completos, o.Tipo_Usuario)).ToList();
            return Json(items, JsonRequestBehavior.AllowGet);
        }

        public ActionResult GetSubtipoDependiente(int id)
        {
            var subtipo = CatalogoEntity.ConsultarCatalogoPorPadre(id, "SUBTIPO").ToList();
            return Json(subtipo);
        }

        public ActionResult DescargarReporteFormatoExcel()
        {
            var collection = AsignacionEntity.ListarAsignacion();

            ExcelPackage package = new ExcelPackage();
            var workSheet = package.Workbook.Worksheets.Add("Asignación Solicitudes");
            workSheet.TabColor = System.Drawing.Color.Black;
            workSheet.DefaultRowHeight = 12;

            List<string> columnas = new List<string>() {
                "ID",
                "TIPO",
                "SUBTIPO",
                "USUARIOS",
                "ESTADO"
            };

            workSheet.Row(1).Height = 20;
            workSheet.Row(1).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            workSheet.Row(1).Style.Font.Bold = true;

            int contador = 0;
            for (int i = 1; i <= columnas.Count; i++)
            {
                workSheet.Cells[1, i].Value = columnas.ElementAt(contador);
                contador++;
            }

            //Body of table  
            int recordIndex = 2;

            foreach (var item in collection)
            {
                workSheet.Cells[recordIndex, 1].Value = item.Id;
                workSheet.Cells[recordIndex, 2].Value = item.Tipo;
                workSheet.Cells[recordIndex, 3].Value = item.Subtipo;
                workSheet.Cells[recordIndex, 4].Value = item.Usuarios; 
                workSheet.Cells[recordIndex, 5].Value = item.Estado;

                recordIndex++;
            }

            for (int i = 1; i <= columnas.Count; i++)
            {
                workSheet.Column(i).AutoFit();
                workSheet.Column(i).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                if (i == 3)
                {
                    workSheet.Column(i).Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                }
            }

            return File(package.GetAsByteArray(), XlsxContentType, "ListadoAsignaciónSolicitud.xlsx");

        }

        public ActionResult DescargarReporteFormatoPDF()
        {
            // Seleccionar las columnas a exportar
            var results = AsignacionEntity.ListarAsignacion();

            var list = Reportes.SerializeToJSON(results);
            return Content(list, "application/json");
        }

        public ActionResult DescargarReporteFormatoCSV()
        {
            var comlumHeadrs = new string[]
            {
                "ID",
                "TIPO",
                "SUBTIPO",
                "USUARIOS",
                "ESTADO"
            };

            var listado = (from item in AsignacionEntity.ListarAsignacion()
                           select new object[]
                           {
                                            item.Id,
                                            item.Tipo,
                                            item.Subtipo,
                                            item.Usuarios,
                                            item.Estado
                           }).ToList();

            // Build the file content
            var employeecsv = new StringBuilder();
            listado.ForEach(line =>
            {
                employeecsv.AppendLine(string.Join(",", line));
            });

            byte[] buffer = Encoding.Default.GetBytes($"{string.Join(",", comlumHeadrs)}\r\n{employeecsv.ToString()}");
            return File(buffer, "text/csv", $"ListadoAsignaciónSolicitud.csv");
        }

    }

}


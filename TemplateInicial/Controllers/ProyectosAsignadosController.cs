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
using System.Text;
using System.Threading.Tasks;
using System.Web;
using System.Web.Mvc;

namespace TemplateInicial.Controllers
{
    [Autenticado]
    public class ProyectosAsignadosController : BaseAppController
    {
        private const string XlsxContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

        // GET: ProyectosAsignados
        public ActionResult Index()
        {
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
            //obtener codigo usuario logeado 
            var usuarioSesion = ViewData["usuario"] = System.Web.HttpContext.Current.Session["usuario"];
            var idUsuario = Convert.ToInt16(usuarioSesion);

            ViewBag.NombreListado = Etiquetas.TituloGridAvanceProyectos;

            //Controlar permisos
            var user = ViewData["usuario"] = System.Web.HttpContext.Current.Session["usuario"];
            var usuario = int.Parse(user.ToString());
            string nombreControlador = ControllerContext.RouteData.Values["controller"].ToString();
            ViewBag.NombreControlador = nombreControlador;

            ViewBag.AccionesUsuario = ManejoPermisosEntity.ListadoAccionesCatalogoUsuario(usuario, nombreControlador);

            //Obtener Acciones del controlador
            ViewBag.AccionesControlador = GetMetodosControlador(nombreControlador);


            var listado = ProyectosAsigandosEntity.ListarProyectos(idUsuario);

            //var listado = TarifarioEntity.ListadoTarifario();
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

        // GET: ProyectosAsignados/Details/5
        public ActionResult Details(int id)
        {
            return View();
        }

        // GET: ProyectosAsignados/Create
        public ActionResult Create()
        { 
            //obtener codigo usuario logeado 
            var usuarioSesion = ViewData["usuario"] = System.Web.HttpContext.Current.Session["usuario"];
            var idUsuario = Convert.ToInt16(usuarioSesion);

            ViewBag.DatosUsuario = UsuarioEntity.ConsultarUsuario(idUsuario);
            ViewBag.usuario = idUsuario;
             
            return View();
        }

        // POST: ProyectosAsignados/Create
        [HttpPost]
        public ActionResult Create(DetalleProyectosAsignados detalleAvanceProyecto, int id_codigo_cotizacion, int id_proyecto, int etapa_cliente, int etapa_general, int estatus_detallado, int estatus_general, DateTime fecha_inicio_programado, DateTime fecha_fin_programado, DateTime fecha_inicio_real, DateTime fecha_fin_real, int horas_programadas, int horas_reales)
        {
            try
            {
                RespuestaTransaccion resultado = ProyectosAsigandosEntity.CrearAvanceProyecto(detalleAvanceProyecto, id_codigo_cotizacion, id_proyecto, etapa_cliente, etapa_general, estatus_detallado, estatus_general, fecha_inicio_programado, fecha_fin_programado, fecha_inicio_real, fecha_fin_real, horas_programadas, horas_reales);
                return Json(new { Resultado = resultado }, JsonRequestBehavior.AllowGet);
            }
            catch(Exception ex)
            {
                return Json(new { Resultado = new RespuestaTransaccion { Estado = false, Respuesta = ex.Message } }, JsonRequestBehavior.AllowGet);
            }
        }

        // GET: ProyectosAsignados/Edit/5
        public ActionResult Edit(int id)
        {
            //obtener codigo usuario logeado 
            var usuarioSesion = ViewData["usuario"] = System.Web.HttpContext.Current.Session["usuario"];
            var idUsuario = Convert.ToInt16(usuarioSesion);

            ViewBag.DatosUsuario = UsuarioEntity.ConsultarUsuario(idUsuario);
            ViewBag.usuario = idUsuario;

            ViewBag.idProyecto = id;
              
            ViewBag.CodigoCotizacion = ProyectosAsigandosEntity.ConsultarProyectoPorCodigoCotizacion(id); 

            return View();
        }

        // POST: ProyectosAsignados/Edit/5
        [HttpPost]
        public ActionResult Edit(DetalleProyectosAsignados detalleAvanceProyecto, int id_codigo_cotizacion, int id_proyecto, int etapa_cliente, int etapa_general, int estatus_detallado, int estatus_general, DateTime fecha_inicio_programado, DateTime fecha_fin_programado, DateTime fecha_inicio_real, DateTime fecha_fin_real, int horas_programadas, int horas_reales)
        {
            try
            {
                RespuestaTransaccion resultado = ProyectosAsigandosEntity.CrearAvanceProyecto(detalleAvanceProyecto, id_codigo_cotizacion, id_proyecto, etapa_cliente, etapa_general, estatus_detallado, estatus_general, fecha_inicio_programado, fecha_fin_programado, fecha_inicio_real, fecha_fin_real, horas_programadas, horas_reales);
                return Json(new { Resultado = resultado }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                return Json(new { Resultado = new RespuestaTransaccion { Estado = false, Respuesta = ex.Message } }, JsonRequestBehavior.AllowGet);
            }
        }

        // GET: ProyectosAsignados/Delete/5
        public ActionResult Delete(int id)
        {
            return View();
        }

        // POST: ProyectosAsignados/Delete/5
        [HttpPost]
        public ActionResult Delete(int id, FormCollection collection)
        {
            try
            {
                // TODO: Add delete logic here

                return RedirectToAction("Index");
            }
            catch
            {
                return View();
            }
        }

        public JsonResult ObtenerDatosCodigoCotizacion(int id)
        {
            var proyecto = ProyectosAsigandosEntity.ConsultarDatosCodigoCotizacion(id);
            var data = proyecto;
            return Json(data, JsonRequestBehavior.AllowGet);
        }

        public JsonResult ObtenerDetalleUltimoAvance(int id)
        {
            var proyecto = ProyectosAsigandosEntity.ConsultarUltimoAvance(id);
            var data = proyecto;
            return Json(data, JsonRequestBehavior.AllowGet);
        }

        public JsonResult GetCodigoCotizacion(int? id)
        { 
            List<MultiSelectJQueryUi> items = new List<MultiSelectJQueryUi>();

            items = ProyectosAsigandosEntity.ObtenerListadoProyectos(id).Select(o => new MultiSelectJQueryUi(Convert.ToInt64(o.Value), o.Text, "")).ToList();
            return Json(items, JsonRequestBehavior.AllowGet);
        }

        public ActionResult DescargarAvanceProyectoFormmatoExcel()
        { 
            var collection = ProyectosAsigandosEntity.ListarProyectosTotales();

            ExcelPackage package = new ExcelPackage();
            var workSheet = package.Workbook.Worksheets.Add("Listado Avance Proyecto");
            workSheet.TabColor = System.Drawing.Color.Black;
            workSheet.DefaultRowHeight = 12;

            List<string> columnas = new List<string>() {
                "Codigo Cotización",
                "Responsable",
                "Nombre Proyecto",
                "Cliente",
                "Ejecutivo",
                "Fase",
                "Horas Programadas",
                "Inicio Programado",
                "Fin Programado",
                "Horas Reales",
                "Inicio Real",
                "Fin Real",
                "% Avance Programado",
                "% Avance Real",
                "% Avance Registrado",
                "Real Vs Prog. (%)",
                "Real Vs Prog. (Días)",
                "Fecha Avance",
                "Descripción Avance",
                "Ejecutado",
                "En Ejecución",
                "Bloqueantes",
                "Próximos Pasos"};

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
                workSheet.Cells[recordIndex, 1].Value = item.CodigoCotizacion;
                workSheet.Cells[recordIndex, 2].Value = item.Responsable;
                workSheet.Cells[recordIndex, 3].Value = item.NombreProyecto;
                workSheet.Cells[recordIndex, 4].Value = item.Cliente;
                workSheet.Cells[recordIndex, 5].Value = item.Ejecutivo;
                workSheet.Cells[recordIndex, 6].Value = item.Fase;
                workSheet.Cells[recordIndex, 7].Value = item.HorasProg;
                workSheet.Cells[recordIndex, 8].Value = item.FechaInicioProg;
                workSheet.Cells[recordIndex, 9].Value = item.FechaFinProg;
                workSheet.Cells[recordIndex, 10].Value = item.HorasReal;
                workSheet.Cells[recordIndex, 11].Value = item.FechaInicioReal;
                workSheet.Cells[recordIndex, 12].Value = item.FechaFinReal;
                workSheet.Cells[recordIndex, 13].Value = item.AvanceProgramado.ToString()+" %"; 
                workSheet.Cells[recordIndex, 14].Value = item.AvanceReal.ToString() + " %";
                workSheet.Cells[recordIndex, 15].Value = item.Avance.ToString() + " %";
                workSheet.Cells[recordIndex, 16].Value = item.RealProgramado.ToString() + " %";
                workSheet.Cells[recordIndex, 17].Value = item.Dias;
                workSheet.Cells[recordIndex, 18].Value = item.FechaAvance;
                workSheet.Cells[recordIndex, 19].Value = item.Descripcion;
                workSheet.Cells[recordIndex, 20].Value = item.Ejecutado;
                workSheet.Cells[recordIndex, 21].Value = item.En_Ejecucion;
                workSheet.Cells[recordIndex, 22].Value = item.Bloqueantes;
                workSheet.Cells[recordIndex, 23].Value = item.ProximosPasos; 
                 
                recordIndex++;
            }

            for (int i = 1; i <= columnas.Count; i++)
            {
                workSheet.Column(i).AutoFit();
                workSheet.Column(i).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                if (i == columnas.Count)
                {
                    workSheet.Column(i).Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                }
            }


            return File(package.GetAsByteArray(), XlsxContentType, "ListadoAvancesProyectos.xlsx");
        }

        public ActionResult DescargarReporteFormatoCSV()
        {
            var comlumHeadrs = new string[]
            {
                "Codigo Cotización",
                "Responsable",
                "Nombre Proyecto",
                "Cliente",
                "Ejecutivo",
                "Fase",
                "Horas Programadas",
                "Inicio Programado",
                "Fin Programado",
                "Horas Reales",
                "Inicio Real",
                "Fin Real",
                "% Avance Programado",
                "% Avance Real",
                "% Avance Registrado",
                "% Real vs Programado",
                "% Días vs Programado",
                "Fecha Avance",
                "Descripción Avance",
                "Ejecutado",
                "En Ejecución",
                "Bloqueantes",
                "Próximos Pasos" 
            };

            var listado = (from item in ProyectosAsigandosEntity.ListarProyectosTotales()
            select new object[]
                           {

                                item.CodigoCotizacion,
                                item.Responsable,
                                item.NombreProyecto,
                                item.Cliente,
                                item.Ejecutivo,
                                item.Fase,
                                item.HorasProg,
                                item.FechaInicioProg,
                                item.FechaFinProg,
                                item.HorasReal,
                                item.FechaInicioReal,
                                item.FechaFinReal,
                                item.AvanceProgramado.ToString(),
                                item.AvanceReal.ToString(),
                                item.Avance.ToString(),
                                item.RealProgramado.ToString(),
                                item.Dias,
                                item.FechaAvance,
                                item.Descripcion,
                                item.Ejecutado,
                                item.En_Ejecucion,
                                item.Bloqueantes,
                                item.ProximosPasos 
                           }).ToList();

            // Build the file content
            var employeecsv = new StringBuilder();
            listado.ForEach(line =>
            {
                employeecsv.AppendLine(string.Join("¬", line));
            });

            byte[] buffer = Encoding.Default.GetBytes($"{string.Join("¬", comlumHeadrs)}\r\n{employeecsv.ToString()}");

            return File(buffer, "text/csv", $"AvancesProyectos.csv");
        }

        public ActionResult DescargarReporteFormatoPDF()
        {
            // Seleccionar las columnas a exportar
            var results = ProyectosAsigandosEntity.ListarProyectosTotales();

            var list = Reportes.SerializeToJSON(results);
            return Content(list, "application/json");
        }
    }
}

using System;
using System.Collections.Generic;
using System.Data; 
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using System.Web;
using System.Web.Mvc;
using GestionPPM.Entidades.Metodos;
using GestionPPM.Entidades.Modelo;
using GestionPPM.Repositorios;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using Seguridad.Helper;

namespace TemplateInicial.Controllers
{
    [Autenticado]
    public class TipoNotificacionController : BaseAppController
    { 
        private const string XlsxContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
        // GET: Tarifarios
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
            ViewBag.NombreListado = Etiquetas.TituloGridTipoNotificacion;

            //Controlar permisos
            var user = ViewData["usuario"] = System.Web.HttpContext.Current.Session["usuario"];
            var usuario = int.Parse(user.ToString());
            string nombreControlador = ControllerContext.RouteData.Values["controller"].ToString();
            ViewBag.NombreControlador = nombreControlador;

            ViewBag.AccionesUsuario = ManejoPermisosEntity.ListadoAccionesCatalogoUsuario(usuario, nombreControlador);

            //Obtener Acciones del controlador
            ViewBag.AccionesControlador = GetMetodosControlador(nombreControlador);

            //Búsqueda

            var listado = TipoNotificacionEntity.ListarTipoNotificaciones(); 
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

        // GET: Tarifarios/Create
        public ActionResult Create()
        {  
            return View();
        }

        [HttpPost]
        public ActionResult Create(TipoNotificacion tipoNotificacion)
        {
            try
            {
                string nombreTipoNotificacion = (tipoNotificacion.nombre_notificacion ?? string.Empty).ToLower().Trim();

                var tarifariosIguales = TipoNotificacionEntity.ListarTipoNotificaciones().Where(s => (s.Nombre_Notificacion ?? string.Empty).ToLower().Trim() == nombreTipoNotificacion).ToList();

                if (tarifariosIguales.Count > 0)
                    return Json(new { Resultado = new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeValidacionNombreTarifario } }, JsonRequestBehavior.AllowGet);
                 
                RespuestaTransaccion resultado = TipoNotificacionEntity.CrearTipoNotificacion(tipoNotificacion);

                return Json(new { Resultado = resultado }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                return Json(new { Resultado = new RespuestaTransaccion { Estado = false, Respuesta = ex.Message } }, JsonRequestBehavior.AllowGet);
            }
        }

        // GET: Tarifarios/Edit/5
        public ActionResult Edit(int? id)
        {
           
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            var tipoNotificacion = TipoNotificacionEntity.ConsultarNotificacion(id.Value);

            ViewBag.Tiempo = tipoNotificacion.tiempo_espera;

            if (tipoNotificacion == null)
            {
                return HttpNotFound();
            }
            return View(tipoNotificacion);
        }

        [HttpPost]
        public ActionResult Edit(TipoNotificacion tipoNotificacion)
        {
            try
            {
                string nombreTipoNotificacion = (tipoNotificacion.nombre_notificacion ?? string.Empty).ToLower().Trim();

                var tarifariosIguales = TipoNotificacionEntity.ListarTipoNotificaciones().Where(s => (s.Nombre_Notificacion ?? string.Empty).ToLower().Trim() == nombreTipoNotificacion && s.id_notificacion != tipoNotificacion.id_notificacion).ToList();
                
                if (tarifariosIguales.Count > 0)
                    return Json(new { Resultado = new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeValidacionNombreTarifario } }, JsonRequestBehavior.AllowGet);
                 
                RespuestaTransaccion resultado = TipoNotificacionEntity.ActualizarTipoNotificacion(tipoNotificacion);

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
            RespuestaTransaccion resultado = TipoNotificacionEntity.EliminarTipoNotificacion(id);// await db.Cabecera.FindAsync(id);

            //return RedirectToAction("Index");
            return Json(new { Resultado = resultado }, JsonRequestBehavior.AllowGet);
        }          

        public ActionResult DescargarReporteFormatoExcel()
        {
            var collection = TipoNotificacionEntity.ListarTipoNotificaciones();

            ExcelPackage package = new ExcelPackage();
            var workSheet = package.Workbook.Worksheets.Add("Tabla Costos");
            workSheet.TabColor = System.Drawing.Color.Black;
            workSheet.DefaultRowHeight = 12;

            List<string> columnas = new List<string>() {
                "ID",
                "NOMBRE",
                "DESCRIPCIÓN",
                "TIEMPO ESPERA", 
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
                workSheet.Cells[recordIndex, 1].Value = item.id_notificacion;
                workSheet.Cells[recordIndex, 2].Value = item.Nombre_Notificacion;
                workSheet.Cells[recordIndex, 3].Value = item.Descripcion_Tarifario; 
                workSheet.Cells[recordIndex, 4].Value = item.Tiempo_Espera; 
                workSheet.Cells[recordIndex, 5].Value = item.Estado_Notificacion;                 

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

            return File(package.GetAsByteArray(), XlsxContentType, "ListadoTipoNotificaciones.xlsx"); 

        }

        public ActionResult DescargarReporteFormatoPDF()
        {
            // Seleccionar las columnas a exportar
            var results = TipoNotificacionEntity.ListarTipoNotificaciones();

            var list = Reportes.SerializeToJSON(results);
            return Content(list, "application/json");
        }

        public ActionResult DescargarReporteFormatoCSV()
        {
            var comlumHeadrs = new string[]
            {
                 "ID",
                "NOMBRE",
                "DESCRIPCIÓN",
                "TIEMPO ESPERA",
                "ESTADO"
            };

            var listado = (from item in TipoNotificacionEntity.ListarTipoNotificaciones()
                           select new object[]
                           {
                               item.id_notificacion,
                               item.Nombre_Notificacion,
                               item.Descripcion_Tarifario,
                               item.Tiempo_Espera, 
                               item.Estado_Notificacion

                           }).ToList();

            // Build the file content
            var employeecsv = new StringBuilder();
            listado.ForEach(line =>
            {
                employeecsv.AppendLine(string.Join(",", line));
            });

            byte[] buffer = Encoding.Default.GetBytes($"{string.Join(",", comlumHeadrs)}\r\n{employeecsv.ToString()}");
            return File(buffer, "text/csv", $"Tipo_Notificacion.csv");
        }

    }
}

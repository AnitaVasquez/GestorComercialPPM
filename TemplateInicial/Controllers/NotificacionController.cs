using GestionPPM.Entidades.Metodos;
using GestionPPM.Entidades.Modelo;
using GestionPPM.Repositorios;
using NonFactors.Mvc.Grid;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using Seguridad.Helper;
using System;
using System.Drawing;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using System.Web;
using System.Web.Mvc;

namespace TemplateInicial.Controllers
{
    public partial class NotificacionesCompletaPDF
    {
     
        public string NombreTarea { get; set; }
        public string DescripcionTarea { get; set; }
        public string NombreEmisor { get; set; }
        public string CorreoEmisor { get; set; }
       
        public string CorreosDestinarios { get; set; }
        public string AsuntoCorreo { get; set; }
        public string HtmlPlantilla { get; set; }
        
        public string AdjuntosCorreo { get; set; }
        public System.DateTime FechaEnvioCorreo { get; set; }
        
        public string EstadoActivacionNotificacion { get; set; }
 
        public string EstadoEnColaNotificacion { get; set; }
        public Nullable<bool> EstadoEnviadoNotificacion { get; set; }
        public string EstadoEnvioNotificacion { get; set; }
        public string DetalleEstadoEjecucionNotificacion { get; set; }
       
        
   
    }
    [Autenticado]
    public class NotificacionController : BaseAppController
    {

        // GET: Notificacion
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
            ViewBag.NombreListado = Etiquetas.TituloGridNotificaciones;

            //Controlar permisos
            var user = ViewData["usuario"] = System.Web.HttpContext.Current.Session["usuario"];
            var usuario = int.Parse(user.ToString());
            string nombreControlador = ControllerContext.RouteData.Values["controller"].ToString();
            ViewBag.NombreControlador = nombreControlador;

            ViewBag.AccionesUsuario = ManejoPermisosEntity.ListadoAccionesCatalogoUsuario(usuario, nombreControlador);

            //Obtener Acciones del controlador
            ViewBag.AccionesControlador = GetMetodosControlador(nombreControlador);

            //Búsqueda
            var listado = NotificacionEntity.ListarNotificaciones();
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

        [HttpGet]
        public ActionResult DescargarReporteFormatoExcel()
        {
            // Using EPPlus from nuget
            using (ExcelPackage package = new ExcelPackage())
            {
                Int32 row = 2;
                Int32 col = 1;

                package.Workbook.Worksheets.Add("Data");
                IGrid<NotificacionesCompletasInfo> grid = CreateExportableGrid();
                ExcelWorksheet sheet = package.Workbook.Worksheets["Data"];

                foreach (IGridColumn column in grid.Columns)
                {
                    sheet.Cells[1, col].Value = column.Title;
                    sheet.Column(col++).Width = 18;

                    column.IsEncoded = false;
                }

                foreach (IGridRow<NotificacionesCompletasInfo> gridRow in grid.Rows)
                {
                    col = 1;
                    foreach (IGridColumn column in grid.Columns)
                        sheet.Cells[row, col++].Value = column.ValueFor(gridRow);

                    row++;
                }

                col = 1;
                foreach (IGridColumn column in grid.Columns)
                {
                    
                    using (ExcelRange rowRange = sheet.Cells[1, col++])
                    {
                        rowRange.Style.Fill.PatternType = ExcelFillStyle.Solid;
                        rowRange.Style.Fill.BackgroundColor.SetColor(Color.Orange);
                    }
                }

                return File(package.GetAsByteArray(), "application/unknown", "ListadoNotificaciones.xlsx");
            }
        }

        public ActionResult DescargarReporteFormatoPDF()
        {
            // Seleccionar las columnas a exportar
            var results = NotificacionEntity.ListarNotificaciones().Select(s=> new NotificacionesCompletaPDF { 
                
                NombreTarea=s.NombreTarea,
                DescripcionTarea=s.DescripcionTarea,
                NombreEmisor=s.NombreEmisor,
                CorreoEmisor=s.CorreoEmisor,
                CorreosDestinarios=s.CorreosDestinarios,
                AsuntoCorreo= s.AsuntoCorreo,
                HtmlPlantilla =s.NombreArchivoPlantillaCorreo,
                AdjuntosCorreo= s.AdjuntosCorreo,
                FechaEnvioCorreo=s.FechaEnvioCorreo,
                EstadoActivacionNotificacion=s.EstadoActivacionNotificacion,
                EstadoEnColaNotificacion=s.EstadoEnColaNotificacion,
                EstadoEnvioNotificacion=s.EstadoEnvioNotificacion,
                DetalleEstadoEjecucionNotificacion=s.DetalleEstadoEjecucionNotificacion

            }).Take(10).ToList();

            var list = Reportes.SerializeToJSON(results);
            return Content(list, "application/json");
        }

        public ActionResult DescargarReporteFormatoCSV()
        {
            var comlumHeadrs = new string[]
            {
                
                "NOMBRE",
                "DESCRIPCION",
                "EMISOR",
                "CORREO EMISOR",
                "CORREOS DESTINATARIOS",
                "ASUNTO CORREO",
                "HTML PLANTILLA",
                "ADJUNTOS CORREO",
                "FECHA DE ENVÍO",
                "ESTADO ACTIVACIÓN",
                "ESTADO EN COLA",
                "ESTADO ENVÍO",
                "DETALLE"


            };

            var listado = (from item in NotificacionEntity.ListarNotificaciones()
                           select new object[]
                           {
                                            item.NombreTarea,
                                            item.DescripcionTarea,
                                            item.NombreEmisor,
                                            item.CorreoEmisor,
                                            item.CorreosDestinarios,
                                            item.AsuntoCorreo,
                                            item.NombreArchivoPlantillaCorreo,
                                            item.AdjuntosCorreo,
                                            item. FechaEnvioCorreo,
                                            item.EstadoActivacionNotificacion,
                                            item.EstadoEnColaNotificacion,
                                            item.EstadoEnvioNotificacion,
                                            item.DetalleEstadoEjecucionNotificacion
                           }).ToList();

            // Build the file content
            var employeecsv = new StringBuilder();
            listado.ForEach(line =>
            {
                employeecsv.AppendLine(string.Join(",", line));
            });

            byte[] buffer = Encoding.Default.GetBytes($"{string.Join(",", comlumHeadrs)}\r\n{employeecsv.ToString()}");
            return File(buffer, "text/csv", $"Notificaciones.csv");
        }


        private IGrid<NotificacionesCompletasInfo> CreateExportableGrid()
        {
            IGrid<NotificacionesCompletasInfo> grid = new Grid<NotificacionesCompletasInfo>(NotificacionEntity.ListarNotificaciones());
            grid.ViewContext = new ViewContext { HttpContext = HttpContext };
            grid.Query = Request.QueryString;

            grid.Columns.Add(model => model.NombreTarea).Titled("Nombre");
            grid.Columns.Add(model => model.DescripcionTarea).Titled("Descripción");
            grid.Columns.Add(model => model.NombreEmisor).Titled("Emisor");
            grid.Columns.Add(model => model.CorreoEmisor).Titled("Correo Emisor");
            grid.Columns.Add(model => model.CorreosDestinarios).Titled("Correos Destinatarios");
            grid.Columns.Add(model => model.AsuntoCorreo).Titled("Asunto Correo");
            grid.Columns.Add(model => model.NombreArchivoPlantillaCorreo).Titled("Nombre Archivo Plantilla Correo");
            grid.Columns.Add(model => model.CorreosDestinarios).Titled("Correos Destinatarios");
            grid.Columns.Add(model => model.CuerpoCorreo).Titled("Cuerpo Correo");
            grid.Columns.Add(model => model.AdjuntosCorreo).Titled("Adjuntos Correo");
            grid.Columns.Add(model => model.CorreosDestinarios).Titled("Correos Destinatarios");
            grid.Columns.Add(model => model.FechaEnvioCorreo).Titled("Fecha de Envío Correo");
            //grid.Columns.Add(model => model.EstadoNotificacion).Titled("Estado Notificacion");
            grid.Columns.Add(model => model.EstadoActivacionNotificacion).Titled("Estado Activación");
            //grid.Columns.Add(model => model.EstadoEjecucionNotificacion).Titled("Estado Ejecución");
            grid.Columns.Add(model => model.EstadoEnColaNotificacion).Titled("Estado En Cola");
            //grid.Columns.Add(model => model.EstadoEnviadoNotificacion).Titled("Estado Enviado");
            grid.Columns.Add(model => model.EstadoEnvioNotificacion).Titled("Estado Envío");
            grid.Columns.Add(model => model.DetalleEstadoEjecucionNotificacion).Titled("Detalle");

            //grid.Pager = new GridPager<NotificacionesCompletasInfo>(grid);
            //grid.Processors.Add(grid.Pager);
            //grid.Pager.RowsPerPage = 6;

            foreach (IGridColumn column in grid.Columns)
            {
                column.Filter.IsEnabled = true;
                column.Sort.IsEnabled = true;
            }

            return grid;
        }

        [HttpPost]
        public ActionResult Cancelar(int id)
        {
            if (id == 0)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            RespuestaTransaccion resultado = NotificacionEntity.CancelarNotificacion(id);// await db.Cabecera.FindAsync(id);

            return Json(new { Resultado = resultado }, JsonRequestBehavior.AllowGet);
        }



    }
}
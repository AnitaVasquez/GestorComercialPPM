using GestionPPM.Entidades.Metodos;
using GestionPPM.Repositorios;
using GestionPPM.Entidades.Modelo;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Web;
using System.Web.Mvc;
using System.Net;
using Seguridad.Helper;
using System.Text;
using NonFactors.Mvc.Grid;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Drawing;

namespace TemplateInicial.Controllers
{
    public class CostoSublineaNegocioPDF
    {
        public string SublineaNegocio { get; set; }
        public string TipoSolicitud { get; set; }
        public string SubtipoSolicitud { get; set; }
        public decimal Valor { get; set; }
    }
    [Autenticado]
    public class CostoSublineaNegocioController : BaseAppController
    {
        // GET: CostoSublineaNegocio
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
        public async Task<PartialViewResult> IndexGrid(string search)
        {
            ViewBag.NombreListado = Etiquetas.TituloGridSublineaNegocioContacto;
            //Controlar permisos
            var user = ViewData["usuario"] = System.Web.HttpContext.Current.Session["usuario"];
            var usuario = int.Parse(user.ToString());
            string nombreControlador = ControllerContext.RouteData.Values["controller"].ToString();
            ViewBag.NombreControlador = nombreControlador;

            ViewBag.AccionesUsuario = ManejoPermisosEntity.ListadoAccionesCatalogoUsuario(usuario, nombreControlador);

            //Obtener Acciones del controlador
            ViewBag.AccionesControlador = GetMetodosControlador(nombreControlador);

            //Búsqueda
            var listado = CostoSublineaNegocioEntity.ListadoCostosSublineaNegocio();

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

        public ActionResult _Formulario(int? id)
        {
            CostoSublineaNegocio objeto = new CostoSublineaNegocio();
            try
            {
                ViewBag.TituloModal = Etiquetas.TituloPanelSublineaNegocioCreacion;

                if (id.HasValue)
                {
                    var data = CostoSublineaNegocioEntity.ConsultarCostosSublineaNegocio(id.Value);
                    objeto.CatalogoSublineaNegocioID = data.CodigoCatalogoSublineaNegocio;
                    objeto.CatalogoTipoSolicitudID = data.CodigoCatalogoTipoSolicitud;
                    objeto.CatalogoSubTipoSolicitudID = data.CodigoCatalogoSubTipoSolicitud;
                    // Siempre y cuando sean de diferentes tipos los objetos.
                    PropertyCopier<CostosSublineaNegocioInfo, CostoSublineaNegocio>.Copy(data, objeto);
                }

                ViewBag.Modelo = objeto;

                return PartialView();
            }
            catch (Exception ex)
            {
                string mensaje = "Un error ocurrió. {0}";
                ViewBag.Excepcion = string.Format(mensaje, ex.Message.ToString());
                return PartialView("~/Views/Error/_InternalServerError.cshtml");
            }
        }

        // Devuelve los Subtipos de solicitud de acuerdo al Tipo de Solicitud
        public ActionResult GetDependientesTipoSolicitud(int? id)
        {
            return Json(new { Data = CatalogoEntity.ConsultarCatalogoPorPadreByCodigo("SUBTIPO", (id ?? 0)) }, JsonRequestBehavior.AllowGet);
        }

        [HttpPost]
        public ActionResult CreateOrUpdate(CostoSublineaNegocio formulario)
        {
            RespuestaTransaccion resultado = new RespuestaTransaccion();
            try
            {
                //if (CostoSublineaNegocioEntity.SublineaNegocioExistente(formulario.IDCostoSublineaNegocio, formulario.CatalogoSublineaNegocioID))
                //    return Json(new { Resultado = new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeValidacionSublineaNegocioExistente } }, JsonRequestBehavior.AllowGet);

                if (CostoSublineaNegocioEntity.TipoRequerimientoExistente(formulario.IDCostoSublineaNegocio, formulario.CatalogoTipoSolicitudID, formulario.CatalogoSubTipoSolicitudID.Value))
                    return Json(new { Resultado = new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeValidacionTipoRequerimientoExistente } }, JsonRequestBehavior.AllowGet);

                if (formulario.IDCostoSublineaNegocio == 0)
                    resultado = CostoSublineaNegocioEntity.CrearCostoSublineaNegocio(formulario);
                else
                    resultado = CostoSublineaNegocioEntity.EditarCostoSublineaNegocio(formulario);

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
            RespuestaTransaccion resultado = CostoSublineaNegocioEntity.EliminarCostoSublineaNegocio(id);

            return Json(new { Resultado = resultado }, JsonRequestBehavior.AllowGet);
        }

        public ActionResult GetSegmentacion1(int id)
        {
            var subtipo = SublineaNegocioEntity.ConsultarSublineaNegocio(id).ToList();
            return Json(subtipo);
        }

        public ActionResult GetSegmentacion2(int id)
        {
            var subtipo = ProductosGeneralGestorEntity.ConsultarProductosGenerales(id).ToList();
            return Json(subtipo);
        }

        public ActionResult GetSegmentacion2Inversa(int id)
        {
            //obtener codigo producto general
            var producto = ProductosGestorEntity.ConsultarProducto(id);
            var subtipo = ProductosGestorEntity.ConsultarProductosGestor(producto.id_producto_general, id.ToString()).ToList(); 
            return Json(subtipo);
        }

        public ActionResult GetProductoComercial(int id)
        {
            var subtipo = ProductosGestorEntity.ConsultarProductosGestor(id).ToList();
            return Json(subtipo);
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
                IGrid<CostosSublineaNegocioInfo> grid = CreateExportableGrid();
                ExcelWorksheet sheet = package.Workbook.Worksheets["Data"];

                foreach (IGridColumn column in grid.Columns)
                {
                    sheet.Cells[1, col].Value = column.Title;
                    sheet.Column(col++).Width = 18;

                    column.IsEncoded = false;
                }

                foreach (IGridRow<CostosSublineaNegocioInfo> gridRow in grid.Rows)
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

                return File(package.GetAsByteArray(), "application/unknown", "ListadoCostosSublineaNegocio.xlsx");
            }
        }

        private IGrid<CostosSublineaNegocioInfo> CreateExportableGrid()
        {
            IGrid<CostosSublineaNegocioInfo> grid = new Grid<CostosSublineaNegocioInfo>(CostoSublineaNegocioEntity.ListadoCostosSublineaNegocio());
            grid.ViewContext = new ViewContext { HttpContext = HttpContext };
            grid.Query = Request.QueryString;


            grid.Columns.Add(model => model.TextoCatalogoSublineaNegocio).Titled("Sublínea de Negocio").AppendCss("celda-grande");
            grid.Columns.Add(model => model.TextoCatalogoTipoSolicitud).Titled("Tipo de Solicitud").AppendCss("celda-grande");
            grid.Columns.Add(model => model.TextoCatalogoSubTipoSolicitud).Titled("Subtipo de Solicitud").AppendCss("celda-grande");
            grid.Columns.Add(model => (((Math.Round(model.Valor, 2).ToString("N2").Replace(",", "-")).Replace(".", ",")).Replace("-", "."))).AppendCss("alinear-derecha").Titled("Valor (US$)");


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

        public ActionResult DescargarReporteFormatoPDF()
        {
            // Seleccionar las columnas a exportar
            var results = CostoSublineaNegocioEntity.ListadoCostosSublineaNegocio().Select(s => new CostoSublineaNegocioPDF
            {
                SublineaNegocio = s.TextoCatalogoSublineaNegocio,
                TipoSolicitud = s.TextoCatalogoTipoSolicitud,
                SubtipoSolicitud = s.TextoCatalogoSubTipoSolicitud,
                Valor = s.Valor
            });

            var list = Reportes.SerializeToJSON(results);
            return Content(list, "application/json");
        }

        public ActionResult DescargarReporteFormatoCSV()
        {
            var comlumHeadrs = new string[]
            {
                "Sublínea de Negocio",
                "Tipo de Solicitud",
                "Subtipo de Solicitud",
                "Valor (US$)",
            };

            var listado = (from item in CostoSublineaNegocioEntity.ListadoCostosSublineaNegocio()
                           select new object[]
                           {
                                            item.TextoCatalogoSublineaNegocio,
                                            item.TextoCatalogoTipoSolicitud,
                                            item.TextoCatalogoSubTipoSolicitud,
                                            item.Valor
                           }).ToList();

            // Build the file content
            var employeecsv = new StringBuilder();
            listado.ForEach(line =>
            {
                employeecsv.AppendLine(string.Join(",", line));
            });

            byte[] buffer = Encoding.Default.GetBytes($"{string.Join(",", comlumHeadrs)}\r\n{employeecsv.ToString()}");
            return File(buffer, "text/csv", $"CostosSublineaNegocio.csv");
        } 
    }
}
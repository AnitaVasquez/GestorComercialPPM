using GestionPPM.Entidades.Metodos;
using GestionPPM.Repositorios;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Web.Mvc;
using GestionPPM.Entidades.Modelo;
using System.IO;
using iTextSharp.text;
using iTextSharp.text.pdf;
using TemplateInicial.Helper;
using System.Configuration;
using System.Text.RegularExpressions;
using System.Globalization;
using Seguridad.Helper;
using NonFactors.Mvc.Grid;
using OfficeOpenXml.Style;
using System.Drawing;
using OfficeOpenXml;
using System.Text;

namespace TemplateInicial.Controllers
{
    [Autenticado]
    public class AprobarPresupuestoController : BaseAppController
    {
        // GET: Prefactura
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
            ViewBag.NombreListado = Etiquetas.TituloGridPrefacturas;
            //var listado = CotizacionEntity.ListadoPrefacturaSAFI().Where(s => s.aprobacion_prefactura_ejecutivo.Value && !s.aprobacion_final.Value && !s.prefactura_consolidada.Value /*&& string.IsNullOrEmpty(s.numero_factura)*/).ToList();
           
            //Controlar permisos
            var user = ViewData["usuario"] = System.Web.HttpContext.Current.Session["usuario"];
            var usuario = int.Parse(user.ToString());
            string nombreControlador = ControllerContext.RouteData.Values["controller"].ToString();
            ViewBag.NombreControlador = nombreControlador;

            ViewBag.AccionesUsuario = ManejoPermisosEntity.ListadoAccionesCatalogoUsuario(usuario, nombreControlador);

            //Obtener Acciones del controlador
            ViewBag.AccionesControlador = GetMetodosControlador(nombreControlador);

            //Búsqueda
            var listado = PrefacturasSAFIEntity.ListadoPresupuestoAprobadosEjecutivo(usuario);

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
               
        #region Impresion PreFactura Safi - Cotizacion
        public ActionResult GeneracionPrefactura(string listadoIDs, bool descargaDirecta = false)
        {
            List<string> archivos = new List<string>();
            try
            {
                //string test2 = Numalet.ToCardinal("134,40");

                //List<int>ids =  new List<int> {  3 };
                //ids = new List<int> { 1, 2, 3, 4, 5 };
                var ids = !string.IsNullOrEmpty(listadoIDs) ? listadoIDs.Split(',').Select(int.Parse).ToList() : new List<int> { int.Parse(listadoIDs) };

                string nombreFichero = Tools.GetNombreArchivoPlantilla("PREFACTURA");

                string basePath = ConfigurationManager.AppSettings["RepositorioDocumentos"];
                string rutaArchivos = basePath + "\\GESTION_PPM\\SAFI_PREFACTURAS";

                var anioActual = DateTime.Now.Year.ToString();
                var almacenFisicoTemporal = Auxiliares.CrearCarpetasDirectorio(rutaArchivos, new List<string>() { anioActual });

                // Get the complete folder path and store the file inside it.    
                string pathPlantillaPrefactura = Path.Combine(Server.MapPath("~/Plantillas/"), nombreFichero);

                int contador = 1;
                foreach (var id in ids)
                {
                    PrefacturaSAFIInfo prefactura = CotizacionEntity.ConsultarPrefacturaSAFI(id);

                    //string test = Numalet.ToCardinal("134.40", new CultureInfo("en-US"));

                    var NombreDocumentoPrefactura = "Prefactura_" + contador + "_" + prefactura.numero_prefactura + "_" + prefactura.numero_pago + ".pdf";
                    var PathPrefactura = Path.Combine(almacenFisicoTemporal, NombreDocumentoPrefactura);

                    PdfReader reader = new PdfReader(pathPlantillaPrefactura);
                    using (PdfStamper stamper = new PdfStamper(reader, new FileStream(PathPrefactura, FileMode.Create)))
                    {
                        AcroFields form = stamper.AcroFields;
                        var fieldKeys = form.Fields.Keys;
                        foreach (string fieldKey in fieldKeys)
                        {
                            switch (fieldKey)
                            {
                                case "numero_prefactura":
                                    form.SetField(fieldKey, prefactura.numero_prefactura);
                                    break;
                                case "facha_prefactura":
                                    form.SetField(fieldKey, (prefactura.fecha_prefactura.HasValue ? prefactura.fecha_prefactura.Value.ToString("dd/MMMM/yyyy", new CultureInfo("es-ES")) : string.Empty));
                                    break;
                                case "cliente":
                                    form.SetField(fieldKey, prefactura.nombre_comercial_cliente);
                                    break;
                                case "direccion":
                                    form.SetField(fieldKey, prefactura.direccion);
                                    break;
                                case "ciudad":
                                    form.SetField(fieldKey, prefactura.Ciudad);
                                    break;
                                case "telefono":
                                    form.SetField(fieldKey, prefactura.TelefonoEjecutivo);
                                    break;
                                case "numero_pedido":
                                    form.SetField(fieldKey, "0");
                                    break;
                                case "atencion":
                                    form.SetField(fieldKey, string.Empty);
                                    break;
                                case "detalle":
                                    form.SetField(fieldKey, prefactura.detalle_cotizacion);
                                    break;
                                case "ejecutivo":
                                    form.SetField(fieldKey, prefactura.Ejecutivo);
                                    break;
                                case "proyecto":
                                    form.SetField(fieldKey, prefactura.nombre_proyecto);
                                    break;
                                //Detalle
                                case "detalle_numero":


                                    form.SetField(fieldKey, "1");
                                    break;
                                case "detalle_codigo":
                                    form.SetField(fieldKey, prefactura.codigo_producto);
                                    break;
                                case "detalle_descripcion":
                                    form.SetField(fieldKey, prefactura.nombre_producto);
                                    break;
                                case "detalle_cantidad":
                                    form.SetField(fieldKey, prefactura.cantidad.ToString());
                                    break;
                                case "detalle_precio_unitario":
                                    form.SetField(fieldKey, prefactura.precio_unitario.ToString());
                                    break;
                                case "detalle_total":
                                    form.SetField(fieldKey, prefactura.precio_unitario.ToString());
                                    break;
                                case "total_cantidad":
                                    form.SetField(fieldKey, prefactura.cantidad.ToString());
                                    break;
                                case "total_formato_letras":
                                    string valor = prefactura.total_pago.ToString();
                                    valor = Numalet.ToCardinal(valor, new CultureInfo("en-US"));
                                    string formatoLetrasValor = valor + " {0}";
                                    valor = string.Format(formatoLetrasValor, "DOLARES").ToUpper();
                                    form.SetField(fieldKey, valor);
                                    break;
                                case "fecha_vencimiento":
                                    form.SetField(fieldKey, (prefactura.fecha_prefactura.HasValue ? prefactura.fecha_prefactura.Value.ToString("dd/MMMM/yyyy", new CultureInfo("es-ES")) : string.Empty));
                                    break;
                                case "total_suma":
                                    form.SetField(fieldKey, prefactura.precio_unitario.ToString());
                                    break;
                                case "descuento":
                                    form.SetField(fieldKey, prefactura.descuento_pago.ToString());
                                    break;
                                case "subtotal":
                                    form.SetField(fieldKey, prefactura.subtotal_pago.ToString());
                                    break;
                                case "iva":
                                    form.SetField(fieldKey, prefactura.iva_pago.ToString());
                                    break;
                                case "total_final":
                                    form.SetField(fieldKey, prefactura.total_pago.ToString());
                                    break;
                                case "ambiente":
                                    form.SetField(fieldKey, "Gestion PPM Pruebas");
                                    break;
                                default:
                                    break;
                            }
                        }
                        stamper.FormFlattening = true;
                    }

                    //reader.Close();

                    archivos.Add(PathPrefactura);
                    contador++;
                }

                if (!descargaDirecta)
                {
                    // Impresión masiva de prefacturas
                    if (ids.Count > 1)
                    {
                        string pathArchivosConsolidadoPrefacturas = Path.Combine(almacenFisicoTemporal, "ConsolidadoPrefacturas_" + Guid.NewGuid().ToString().Substring(0, 9) + ".pdf");
                        bool generacionCorrecta = MergeArhivosPDF(archivos, pathArchivosConsolidadoPrefacturas);

                        if (generacionCorrecta)
                            return Json(new { Resultado = new RespuestaTransaccion { Estado = true, Respuesta = Mensajes.MensajeTransaccionExitosa }, PathsArchivos = new List<string>() { pathArchivosConsolidadoPrefacturas } }, JsonRequestBehavior.AllowGet);
                        else
                            return Json(new { Resultado = new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeErrorImpresionMasiva }, PathsArchivos = new List<string>() }, JsonRequestBehavior.AllowGet);
                    }
                    else
                    {
                        return Json(new { Resultado = new RespuestaTransaccion { Estado = true, Respuesta = Mensajes.MensajeTransaccionExitosa }, PathsArchivos = archivos }, JsonRequestBehavior.AllowGet);
                    }
                }
                else
                {
                    try
                    {
                        string path = archivos.FirstOrDefault();
                        byte[] fileBytes = System.IO.File.ReadAllBytes(path);
                        string fileName = Path.GetFileName(path);
                        return File(fileBytes, System.Net.Mime.MediaTypeNames.Application.Octet, fileName);
                    }
                    catch (Exception ex)
                    {
                        string mensaje = "Error ({0})";
                        ViewBag.Excepcion = string.Format(mensaje, ex.Message.ToString());
                        return View("~/Views/Error/InternalServerError.cshtml");
                    }
                }
            }
            catch (Exception ex)
            {
                return Json(new { Resultado = new RespuestaTransaccion { Estado = false, Respuesta = ex.Message }, PathsArchivos = archivos }, JsonRequestBehavior.AllowGet);
            }
        }

        // Path Archivos - Destino para Guardar archivo final
        private static bool MergeArhivosPDF(List<string> archivos, string EnRuta)
        {
            try
            {
                PdfReader reader = null;
                Document sourceDocument = null;
                PdfCopy pdfCopyProvider = null;
                PdfImportedPage importedPage;

                sourceDocument = new Document();
                pdfCopyProvider = new PdfCopy(sourceDocument, new FileStream(EnRuta, FileMode.Create));

                //output file Open  
                sourceDocument.Open();

                //files list wise Loop  
                for (int f = 0; f < archivos.Count; f++)
                {
                    int pages = TotalPageCount(archivos[f]);

                    reader = new PdfReader(archivos[f]);
                    //Add pages in new file  
                    for (int i = 1; i <= pages; i++)
                    {
                        importedPage = pdfCopyProvider.GetImportedPage(reader, i);
                        pdfCopyProvider.AddPage(importedPage);
                    }

                    reader.Close();
                }
                //save the output file  
                sourceDocument.Close();

                return true;
            }
            catch (Exception ex)
            {
                return false;
            }

        }

        private static int TotalPageCount(string file)
        {
            using (StreamReader sr = new StreamReader(System.IO.File.OpenRead(file)))
            {
                Regex regex = new Regex(@"/Type\s*/Page[^s]");
                MatchCollection matches = regex.Matches(sr.ReadToEnd());

                return matches.Count;
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
                string mensaje = "Error ({0})";
                ViewBag.Excepcion = string.Format(mensaje, ex.Message.ToString());
                return View("~/Views/Error/InternalServerError.cshtml");
            }
        }

        #endregion 

        public ActionResult _AprobarPrefactura(int? id)
        {
            ViewBag.TituloModal = "Aprobar Presupuesto";
            PrefacturaSAFIInfo modelo = CotizacionEntity.ConsultarPrefacturaSAFI(id.Value);
            return PartialView(modelo);
        }

        public ActionResult AprobacionFinalPrefactura(string listadoIDs, bool ajax = false)
        {
            try
            {
                var ids = !string.IsNullOrEmpty(listadoIDs) ? listadoIDs.Split(',').Select(int.Parse).ToList() : new List<int> { int.Parse(listadoIDs) };

                var user = ViewData["usuario"] = System.Web.HttpContext.Current.Session["usuario"];
                var usuarioID = Convert.ToInt16(user);

                bool aprobado = CotizacionEntity.AprobacionInicialPrefactura(ids, usuarioID);

                if (!aprobado)
                {
                    if (!ajax)
                    {
                        string mensaje = "Error ({0})";
                        ViewBag.Excepcion = string.Format(mensaje, "No se pudo aprobar la prefactura.");
                        return View("~/Views/Error/InternalServerError.cshtml");
                    }
                    else
                        return Json(new { Resultado = new RespuestaTransaccion { Estado = false, Respuesta = "No se pudo aprobar la prefactura." } }, JsonRequestBehavior.AllowGet);
                }  

                if (!ajax)
                    return View("Index");
                else
                    return Json(new { Resultado = new RespuestaTransaccion { Estado = true, Respuesta = Mensajes.MensajeTransaccionExitosa } }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                if (!ajax)
                {
                    string mensaje = "Error ({0})";
                    ViewBag.Excepcion = string.Format(mensaje, ex.Message.ToString());
                    return View("~/Views/Error/InternalServerError.cshtml");
                }
                else
                    return Json(new { Resultado = new RespuestaTransaccion { Estado = false, Respuesta = ex.Message.ToString() } }, JsonRequestBehavior.AllowGet);
            }
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
                IGrid<ListadoPresupuestosAprobacionEjecutivo> grid = CreateExportableGrid();
                ExcelWorksheet sheet = package.Workbook.Worksheets["Data"];

                foreach (IGridColumn column in grid.Columns)
                {
                    sheet.Cells[1, col].Value = column.Title;
                    sheet.Column(col++).Width = 18;

                    column.IsEncoded = false;
                }

                foreach (IGridRow<ListadoPresupuestosAprobacionEjecutivo> gridRow in grid.Rows)
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

                return File(package.GetAsByteArray(), "application/unknown", "ListadoDocumentos.xlsx");
            }
        }

        public IGrid<ListadoPresupuestosAprobacionEjecutivo> CreateExportableGrid()
        {
            //Controlar permisos
            var user = ViewData["usuario"] = System.Web.HttpContext.Current.Session["usuario"];
            var usuario = int.Parse(user.ToString());

            IGrid<ListadoPresupuestosAprobacionEjecutivo> grid = new Grid<ListadoPresupuestosAprobacionEjecutivo>(PrefacturasSAFIEntity.ListadoPresupuestoAprobadosEjecutivo(usuario));
            grid.ViewContext = new ViewContext { HttpContext = HttpContext };
            grid.Query = Request.QueryString;

            //grid.Columns.Add(model => model.id_facturacion_safi).Titled("ID Documento").Css("hidden");
            grid.Columns.Add(model => model.codigo_cotizacion).Titled("Código de Cotización").AppendCss("celda-grande");
            grid.Columns.Add(model => model.numero_prefactura).Titled("Número PreFactura").AppendCss("celda-grande");
            grid.Columns.Add(model => model.MKT).Titled("MKT").AppendCss("celda-grande");
            grid.Columns.Add(model => model.PrefacturaConsolidada).Titled("Consolidada").AppendCss("celda-grande");
            grid.Columns.Add(model => model.nombre_comercial_cliente).Titled("Cliente").AppendCss("celda-grande");
            grid.Columns.Add(model => model.detalle_cotizacion).Titled("Detalle").AppendCss("celda-grande");
            grid.Columns.Add(model => model.fecha_aprobacion_prefactura_ejecutivo).Titled("Fecha Aprobación Ejecutivo").Formatted("{0:d}").AppendCss("celda-grande");
            grid.Columns.Add(model => model.Ejecutivo).Titled("Ejecutivo").AppendCss("celda-grande");
            grid.Columns.Add(model => model.cantidad).Titled("Cantidad").AppendCss("celda-grande");
            grid.Columns.Add(model => (((Math.Round(model.precio_unitario, 2).ToString("N2").Replace(",", "-")).Replace(".", ",")).Replace("-", "."))).Titled("Precio").AppendCss("celda-grande");
            grid.Columns.Add(model => (((Math.Round(model.iva_pago, 2).ToString("N2").Replace(",", "-")).Replace(".", ",")).Replace("-", "."))).Titled("IVA").AppendCss("celda-grande");
            grid.Columns.Add(model => (((Math.Round(model.total_pago, 2).ToString("N2").Replace(",", "-")).Replace(".", ",")).Replace("-", "."))).Titled("Total").AppendCss("celda-grande");

            foreach (IGridColumn column in grid.Columns)
            {
                column.Filter.IsEnabled = true;
                column.Sort.IsEnabled = true;
            }

            return grid;
        }

        public ActionResult DescargarReporteFormatoCSV()
        {
            //Controlar permisos
            var user = ViewData["usuario"] = System.Web.HttpContext.Current.Session["usuario"];
            var usuario = int.Parse(user.ToString());

            var comlumHeadrs = new string[]
            {
                "CODIGO DE COTIZACION",
                "NÚMERO PREFACTURA",
                "MKT",
                "CONSOLIDADA",
                "CLIENTE",
                "DETALLE",
                "FECHA APROBACIN EJECUTIVO",
                "EJECUTIVO",
                "CANTIDAD",
                "PRECIO UNITARIO",
                "IVA",
                "TOTAL",
            };

            var listado = (from item in PrefacturasSAFIEntity.ListadoPresupuestoAprobadosEjecutivo(usuario)
                           select new object[]
                           {
                                item.codigo_cotizacion,
                                item.numero_prefactura,
                                item.MKT,
                                item.PrefacturaConsolidada,
                                item.nombre_comercial_cliente,
                                item.detalle_cotizacion,
                                item.fecha_prefactura.Value.ToString("yyyy-MM-dd"),
                                item.Ejecutivo,
                                item.cantidad,
                                item.precio_unitario,
                                item.iva_pago,
                                item.total_pago,
                           }).ToList();

            // Build the file content
            var employeecsv = new StringBuilder();
            listado.ForEach(line =>
            {
                employeecsv.AppendLine(string.Join(",", line));
            });

            byte[] buffer = Encoding.Default.GetBytes($"{string.Join(",", comlumHeadrs)}\r\n{employeecsv.ToString()}");
            return File(buffer, "text/csv", $"ListadoPrefacturas.csv");
        }

        public ActionResult DescargarReporteFormatoPDF()
        {
            //Controlar permisos
            var user = ViewData["usuario"] = System.Web.HttpContext.Current.Session["usuario"];
            var usuario = int.Parse(user.ToString());

            // Seleccionar las columnas a exportar
            var results = PrefacturasSAFIEntity.ListadoPresupuestoAprobadosEjecutivo(usuario);

            var list = Reportes.SerializeToJSON(results);
            return Content(list, "application/json");
        }

    }
}
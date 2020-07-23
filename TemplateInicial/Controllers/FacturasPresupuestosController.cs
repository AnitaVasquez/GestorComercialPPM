using GestionPPM.Entidades.Metodos;
using GestionPPM.Repositorios;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web.Mvc;
using GestionPPM.Entidades.Modelo;
using Newtonsoft.Json;
using NLog;
using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Style;
using System.Configuration;
using System.Drawing;
using System.IO;
using Seguridad.Helper;

namespace TemplateInicial.Controllers
{
    [Autenticado]
    public class FacturasPresupuestosController : Controller
    {
        // GET: MatrizPresupuesto
        public ActionResult Index()
        {
            return View();
        }

        public ActionResult _SeleccionCodigoCotizacion()
        { 
            ViewBag.TituloModal = "Reporte de Facturas - Presupuesto";
            var usuarioSesion = ViewData["usuario"] = System.Web.HttpContext.Current.Session["usuario"];
            var idUsuario = Convert.ToInt16(usuarioSesion);

            UsuarioCE usuario = UsuarioEntity.ConsultarUsuario(idUsuario);
            if (usuario.cliente_asociado == null)
            {
                usuario.cliente_asociado = 0;
            }
            var ejecutivos = ObtenerListadoEjecutivosCliente(usuario.cliente_asociado.Value,"");
            ViewBag.listadoEjecutivos = ejecutivos;

            return PartialView();
        }

        #region Reporte Matriz de Presupuestos
        public ActionResult ReporteMatriz(int id, string tituloReporte, string ejecutivo, DateTime fechaInicio, DateTime fechaFin)
        {
            try
            { 
                if(ejecutivo.Contains("Seleccione"))
                {
                    ejecutivo = "Todos";
                }

                var usuarioSesion = ViewData["usuario"] = System.Web.HttpContext.Current.Session["usuario"];
                var idUsuario = Convert.ToInt16(usuarioSesion);

                var usuario = UsuarioEntity.ConsultarInformacionPrincipalUsuario(idUsuario);
                var elaboradoPor = usuario != null ? usuario.nombre_usuario : string.Empty;

                var package = new ExcelPackage();
                 

                //PALETA DE COLORES PPM
                var colorGrisOscuroEstiloPPM = Color.FromArgb(60, 66, 87);
                var colorGrisClaroEstiloPPM = Color.FromArgb(240, 240, 240);
                var colorGrisClaro2EstiloPPM = Color.FromArgb(112, 117, 128);
                var colorGrisClaro3EstiloPPM = Color.FromArgb(225, 225, 225);
                var colorBlancoEstiloPPM = Color.FromArgb(255, 255, 255);
                var colorNegroEstiloPPM = Color.FromArgb(0, 0, 0);

                #region Cabecera

                var ws = package.Workbook.Worksheets.Add("Presupuestos-Facturas");

                int columnaFinalDocumentoExcel = 9;
                int columnaInicialDocumentoExcel = 1;

                ws.PrinterSettings.PaperSize = ePaperSize.A4;//ePaperSize.A3;
                ws.PrinterSettings.Orientation = eOrientation.Landscape;
                ws.PrinterSettings.HorizontalCentered = true;
                ws.PrinterSettings.FitToPage = true;
                ws.PrinterSettings.FitToWidth = 1;
                ws.PrinterSettings.FitToHeight = 0;
                ws.PrinterSettings.FooterMargin = 0.70M;//0.5M;
                ws.PrinterSettings.TopMargin = 0.50M;//0.5M;//0.75M;
                ws.PrinterSettings.LeftMargin = 0.70M;//0.5M;//0.25M;
                ws.PrinterSettings.RightMargin = 0.70M; //0.5M;//0.25M;
                ws.Column(9).PageBreak = true;
                ws.PrinterSettings.Scale = 75; // Verificar escala correcta


                ws.Cells.Style.Fill.PatternType = ExcelFillStyle.Solid;
                ws.Cells.Style.Fill.BackgroundColor.SetColor(Color.White);

                var pathUbicacion = Server.MapPath("~/Content/img/LogoPPMPDF.png");

                Image img = Image.FromFile(@pathUbicacion);
                ExcelPicture pic = ws.Drawings.AddPicture("Sample", img);

                pic.SetPosition(1, 1, 0, 40);
                pic.SetSize(184, 52);
                ws.Row(2).Height = 60;

                //LOGO
                using (var range = ws.Cells[1, 1, 2, 4])
                {
                    range.Style.Fill.BackgroundColor.SetColor(colorGrisOscuroEstiloPPM);
                    range.Merge = true;
                }

                // TITULO REPORTE
                using (var range = ws.Cells[1, 5, 2, 8])
                {
                    range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                    range.Style.Fill.BackgroundColor.SetColor(colorGrisOscuroEstiloPPM);

                    range.Style.Font.Color.SetColor(colorBlancoEstiloPPM);

                    range.Value = tituloReporte;
                    //range.Style.Font.Bold = true;
                    range.Style.Font.Size = 18;
                    range.Style.Font.Name = "Raleway";
                    range.Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                    range.Merge = true;
                }

                //FECHA
                using (var range = ws.Cells[1, 9, 2, 9])
                {
                    range.Style.HorizontalAlignment = ExcelHorizontalAlignment.CenterContinuous;
                    range.Style.Fill.BackgroundColor.SetColor(colorGrisClaroEstiloPPM);

                    range.Style.Font.Color.SetColor(colorGrisOscuroEstiloPPM);

                    range.Value = DateTime.Now.ToString("yyyy/MM/dd");//acta.Cabecera.CodigoActa;//actaTitulo + " " + acta.Cabecera.CodigoActa;
                    //range.Style.Font.Bold = true;
                    range.Style.Font.Size = 8;
                    range.Style.Font.Name = "Raleway";
                    range.Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                    range.Merge = true;
                }

                //FORMATO FUENTE TEXTO DE TODO EL DOCUMENTO

                int finCabecera = 4;

                using (var range = ws.Cells[finCabecera, 1, finCabecera, 2])
                {
                    range.Style.Border.Top.Style = ExcelBorderStyle.Hair;
                    range.Style.Border.Left.Style = ExcelBorderStyle.Hair;
                    range.Style.Border.Right.Style = ExcelBorderStyle.Hair;
                    range.Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                    range.Merge = true;

                    range.Value = "Elaborado Por:";

                    range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;


                    range.Style.Fill.BackgroundColor.SetColor(colorGrisOscuroEstiloPPM);
                    range.Style.Font.Color.SetColor(colorBlancoEstiloPPM);
                    range.Style.Font.Bold = true;

                    ws.Cells[range.Start.Row, range.Columns + 1, range.End.Row, range.Columns + 2].Merge = true;

                    ws.Cells[range.Start.Row, range.Columns + 1, range.End.Row, range.Columns + 2].Style.Border.Top.Style = ExcelBorderStyle.Hair;
                    ws.Cells[range.Start.Row, range.Columns + 1, range.End.Row, range.Columns + 2].Style.Border.Left.Style = ExcelBorderStyle.Hair;
                    ws.Cells[range.Start.Row, range.Columns + 1, range.End.Row, range.Columns + 2].Style.Border.Right.Style = ExcelBorderStyle.Hair;
                    ws.Cells[range.Start.Row, range.Columns + 1, range.End.Row, range.Columns + 2].Style.Border.Bottom.Style = ExcelBorderStyle.Hair;

                    ws.Cells[range.Start.Row, range.Columns + 1].Value = elaboradoPor; // Valor campo
                    ws.Cells[range.Start.Row, range.Columns + 1].Style.Indent = 1;

                }

                using (var range = ws.Cells[finCabecera, 5, finCabecera, 6])
                {
                    range.Style.Border.Top.Style = ExcelBorderStyle.Hair;
                    range.Style.Border.Left.Style = ExcelBorderStyle.Hair;
                    range.Style.Border.Right.Style = ExcelBorderStyle.Hair;
                    range.Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                    range.Merge = true;

                    range.Value = "Ejecutivo:";
                    range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;


                    range.Style.Fill.BackgroundColor.SetColor(colorGrisOscuroEstiloPPM);
                    range.Style.Font.Color.SetColor(colorBlancoEstiloPPM);
                    range.Style.Font.Bold = true;

                    ws.Cells[range.Start.Row, range.End.Column + 1, range.End.Row, columnaFinalDocumentoExcel].Merge = true;

                    ws.Cells[range.Start.Row, range.End.Column + 1, range.End.Row, columnaFinalDocumentoExcel].Style.Border.Top.Style = ExcelBorderStyle.Hair;
                    ws.Cells[range.Start.Row, range.End.Column + 1, range.End.Row, columnaFinalDocumentoExcel].Style.Border.Left.Style = ExcelBorderStyle.Hair;
                    ws.Cells[range.Start.Row, range.End.Column + 1, range.End.Row, columnaFinalDocumentoExcel].Style.Border.Right.Style = ExcelBorderStyle.Hair;
                    ws.Cells[range.Start.Row, range.End.Column + 1, range.End.Row, columnaFinalDocumentoExcel].Style.Border.Bottom.Style = ExcelBorderStyle.Hair;

                    ws.Cells[range.Start.Row, range.End.Column + 1].Value = ejecutivo; // Valor campo
                    ws.Cells[range.Start.Row, range.End.Column + 1].Style.Indent = 1;
                }

                finCabecera += 2;
                ws.Row(finCabecera).Height = 20.25;


                #endregion

                Int32 col = 1;
                int contador = 1;

                #region Detalle Acta Cliente


                List<ConsultarPresupuestoFacturados> matriz = SAFIEntity.ConsultarPresupuestosFaturados(id, fechaInicio, fechaFin);

                //totalizadores
                var subtotal = matriz.Sum(m => m.subtotal_pago);
                var iva = matriz.Sum(m => m.iva_pago);
                var total = matriz.Sum(m => m.total_pago);

                var columnas = new List<string> { "N°", "N. PRESUPUESTO", "N. FACTURA","FECHA FACTURA", "CUENTA CONTABLE", "DETALLE", "VALOR", "IVA", "TOTAL" };

                if (matriz.Any())
                {
                    var i = 1;
                    foreach (var item in columnas)
                    {

                        ws.Cells[finCabecera, i].Value = item;
                        ws.Cells[finCabecera, i].Style.Font.Bold = true;

                        ws.Cells[finCabecera, i].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                        ws.Cells[finCabecera, i].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                        ws.Cells[finCabecera, i].Style.Border.Top.Style = ExcelBorderStyle.Hair;
                        ws.Cells[finCabecera, i].Style.Border.Left.Style = ExcelBorderStyle.Hair;
                        ws.Cells[finCabecera, i].Style.Border.Right.Style = ExcelBorderStyle.Hair;
                        ws.Cells[finCabecera, i].Style.Border.Bottom.Style = ExcelBorderStyle.Hair;

                        //worksheet.Cells[1, i].Style.Fill.BackgroundColor. .SetColor(Color.FromArgb(23, 55, 93));
                        i++;
                    }
                    finCabecera++;

                    int numeracion = 1;
                    foreach (var item in matriz)
                    {
                        for (int j = 1; j <= 9; j++)
                        {
                            ws.Cells[finCabecera, j].Style.Border.Top.Style = ExcelBorderStyle.Hair;
                            ws.Cells[finCabecera, j].Style.Border.Left.Style = ExcelBorderStyle.Hair;
                            ws.Cells[finCabecera, j].Style.Border.Right.Style = ExcelBorderStyle.Hair;
                            ws.Cells[finCabecera, j].Style.Border.Bottom.Style = ExcelBorderStyle.Hair;

                            ws.Cells[finCabecera, j].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                            ws.Cells[finCabecera, j].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                            switch (j)
                            {
                                case 1:
                                    ws.Column(j).Width = 7;
                                    ws.Cells[finCabecera, j].Value = numeracion;
                                    break;
                                case 2:
                                    ws.Cells[finCabecera, j].Value = item.numero_prefactura;
                                    break;
                                case 3:
                                    ws.Cells[finCabecera, j].Value = item.numero_factura;
                                    ws.Cells[finCabecera, j].Style.WrapText = true;
                                    break;
                                case 4:
                                    ws.Cells[finCabecera, j].Value = item.fecha_factura;
                                    ws.Cells[finCabecera, j].Style.Numberformat.Format = "yyyy/mm/dd";
                                    ws.Cells[finCabecera, j].Style.WrapText = true;
                                    break;
                                case 5:
                                    ws.Cells[finCabecera, j].Value = item.cuenta_contable;
                                    ws.Cells[finCabecera, j].Style.WrapText = true;
                                    break;
                                case 6:
                                    ws.Cells[finCabecera, j].Value = item.detalle_cotizacion;
                                    ws.Cells[finCabecera, j].Style.WrapText = true;
                                    break;
                                case 7:
                                    ws.Cells[finCabecera, j].Value = item.subtotal_pago;
                                    ws.Cells[finCabecera, j].Style.Numberformat.Format = "$#,##0.00";
                                    break;
                                case 8:
                                    ws.Cells[finCabecera, j].Value = item.iva_pago;
                                    ws.Cells[finCabecera, j].Style.Numberformat.Format = "$#,##0.00";
                                    break;
                                case 9:
                                    ws.Cells[finCabecera, j].Value = item.total_pago;
                                    ws.Cells[finCabecera, j].Style.Numberformat.Format = "$#,##0.00";
                                    break;
                                default:
                                    ws.Cells[finCabecera, j].Value = "ERROR.";
                                    break;
                            }
                        }
                        finCabecera++;
                        numeracion++;
                    }

                    finCabecera++;

                    finCabecera++;
                    ws.Cells[finCabecera, 1].Value = "Total:";
                    ws.Cells[finCabecera, 1].Style.Font.Bold = true;
                    ws.Cells[finCabecera, 2].Value = matriz.Count;
                    ws.Cells[finCabecera, 7].Value = subtotal;
                    ws.Cells[finCabecera, 8].Value = iva;
                    ws.Cells[finCabecera, 9].Value = total;

                    ws.Cells[finCabecera, 1].Style.Border.Top.Style = ExcelBorderStyle.Hair;
                    ws.Cells[finCabecera, 1].Style.Border.Left.Style = ExcelBorderStyle.Hair;
                    ws.Cells[finCabecera, 1].Style.Border.Right.Style = ExcelBorderStyle.Hair;
                    ws.Cells[finCabecera, 1].Style.Border.Bottom.Style = ExcelBorderStyle.Hair;

                    ws.Cells[finCabecera, 2].Style.Border.Top.Style = ExcelBorderStyle.Hair;
                    ws.Cells[finCabecera, 2].Style.Border.Left.Style = ExcelBorderStyle.Hair;
                    ws.Cells[finCabecera, 2].Style.Border.Right.Style = ExcelBorderStyle.Hair;
                    ws.Cells[finCabecera, 2].Style.Border.Bottom.Style = ExcelBorderStyle.Hair;

                    ws.Cells[finCabecera, 7].Style.Numberformat.Format = "$#,##0.00";
                    ws.Cells[finCabecera, 8].Style.Numberformat.Format = "$#,##0.00";
                    ws.Cells[finCabecera, 9].Style.Numberformat.Format = "$#,##0.00";

                    ws.Cells[finCabecera, 7].Style.Font.Bold = true;
                    ws.Cells[finCabecera, 8].Style.Font.Bold = true;
                    ws.Cells[finCabecera, 9].Style.Font.Bold = true;

                }
                else
                {
                    var i = 1;
                    foreach (var item in columnas)
                    {

                        ws.Cells[finCabecera, i].Value = item;
                        ws.Cells[finCabecera, i].Style.Font.Bold = true;

                        ws.Cells[finCabecera, i].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                        ws.Cells[finCabecera, i].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                        ws.Cells[finCabecera, i].Style.Border.Top.Style = ExcelBorderStyle.Hair;
                        ws.Cells[finCabecera, i].Style.Border.Left.Style = ExcelBorderStyle.Hair;
                        ws.Cells[finCabecera, i].Style.Border.Right.Style = ExcelBorderStyle.Hair;
                        ws.Cells[finCabecera, i].Style.Border.Bottom.Style = ExcelBorderStyle.Hair;

                        //worksheet.Cells[1, i].Style.Fill.BackgroundColor. .SetColor(Color.FromArgb(23, 55, 93));
                        i++;
                    }
                    finCabecera++;


                    finCabecera++;
                    ws.Cells[finCabecera, 1].Value = "Total:";
                    ws.Cells[finCabecera, 1].Style.Font.Bold = true;
                    ws.Cells[finCabecera, 2].Value = matriz.Count;

                    ws.Cells[finCabecera, 1].Style.Border.Top.Style = ExcelBorderStyle.Hair;
                    ws.Cells[finCabecera, 1].Style.Border.Left.Style = ExcelBorderStyle.Hair;
                    ws.Cells[finCabecera, 1].Style.Border.Right.Style = ExcelBorderStyle.Hair;
                    ws.Cells[finCabecera, 1].Style.Border.Bottom.Style = ExcelBorderStyle.Hair;

                    ws.Cells[finCabecera, 2].Style.Border.Top.Style = ExcelBorderStyle.Hair;
                    ws.Cells[finCabecera, 2].Style.Border.Left.Style = ExcelBorderStyle.Hair;
                    ws.Cells[finCabecera, 2].Style.Border.Right.Style = ExcelBorderStyle.Hair;
                    ws.Cells[finCabecera, 2].Style.Border.Bottom.Style = ExcelBorderStyle.Hair;


                }

                #endregion

                ws.Column(2).Width = 22;
                ws.Column(3).Width = 18;

                ws.Column(4).Width = 18;
                ws.Column(5).Width = 20;

                ws.Column(6).Width = 39;
                ws.Column(7).Width = 15;
                ws.Column(8).Width = 15;
                ws.Column(9).Width = 15;

                string basePath = ConfigurationManager.AppSettings["RepositorioDocumentos"];
                string rutaArchivos = basePath + "\\GESTION_PPM\\PRESUPUESTOSFACTURADOS";


                var anioActual = DateTime.Now.Year.ToString();
                var almacenFisicoTemporal = Auxiliares.CrearCarpetasDirectorio(rutaArchivos, new List<string>() { anioActual, "Presupuestos-Facturas" });

                //var almacenFisicoOfficeToPDF = Auxiliares.CrearCarpetasDirectorio(Server.MapPath("~/OfficeToPDF/"), new List<string>());

                // Get the complete folder path and store the file inside it.    
                string pathExcel = Path.Combine(almacenFisicoTemporal, "ReportePresupuestosFacturados.xlsx");

                //Write the file to the disk
                FileInfo fi = new FileInfo(pathExcel);
                package.SaveAs(fi); 

                byte[] fileBytes = System.IO.File.ReadAllBytes(pathExcel);
                string fileName = Path.GetFileName(pathExcel);
                return File(fileBytes, System.Net.Mime.MediaTypeNames.Application.Octet, fileName);
            }
            catch (Exception ex)
            {
                string mensaje = "Error ({0})";
                ViewBag.Excepcion = string.Format(mensaje, ex.Message.ToString());
                return View("~/Views/Error/InternalServerError.cshtml");
                //throw;
                //return resultado;
            }
        }
        #endregion

        public IEnumerable<SelectListItem> ObtenerListadoEjecutivosCliente(int codigo, string seleccionado = null)
        {

            List<SelectListItem> listadoEjecutivos = new List<SelectListItem>();
            try
            {
                listadoEjecutivos = ContactoClienteEntity.ListarContactosCliente(codigo).ToList();

                if (listadoEjecutivos == null)
                {
                    return new List<SelectListItem>();
                }


                if (!string.IsNullOrEmpty(seleccionado))
                {
                    if (listadoEjecutivos.FirstOrDefault(s => s.Value == seleccionado.ToString()) != null)
                        listadoEjecutivos.FirstOrDefault(s => s.Value == seleccionado.ToString()).Selected = true;
                }

                return listadoEjecutivos;
            }
            catch (Exception ex)
            {
                return listadoEjecutivos;
            }
        }

    }
}
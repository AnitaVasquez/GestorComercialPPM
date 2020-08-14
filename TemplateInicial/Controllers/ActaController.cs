using GestionPPM.Entidades.Metodos;
using GestionPPM.Entidades.Modelo;
using GestionPPM.Repositorios;
using iTextSharp.text;
using iTextSharp.text.pdf;
using Newtonsoft.Json;
using NLog;
using NonFactors.Mvc.Grid;
using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Style;
using OfficeToPDF;
using Seguridad.Helper;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Configuration;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Web.Mvc;
//Test 2

namespace TemplateInicial.Controllers
{
    [Autenticado]
    public class ActaController : BaseAppController
    {
        public const string XlsxContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

        private static readonly Logger Log = LogManager.GetCurrentClassLogger();

        // GET: Acta
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
            
            ViewBag.NombreListado = Etiquetas.TituloGridActa;
            //Controlar permisos
            var user = ViewData["usuario"] = System.Web.HttpContext.Current.Session["usuario"];
            var usuario = int.Parse(user.ToString());
            string nombreControlador = ControllerContext.RouteData.Values["controller"].ToString();
            ViewBag.NombreControlador = nombreControlador;

            ViewBag.AccionesUsuario = ManejoPermisosEntity.ListadoAccionesCatalogoUsuario(usuario, nombreControlador);

            //Obtener Acciones del controlador
            ViewBag.AccionesControlador = GetMetodosControlador(nombreControlador);

           //Búsqueda
            var listado = ActaEntity.ListadoActa();

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

        public ActionResult _SeleccionTipoActa()
        {
            ViewBag.TituloModal = "Seleccionar el tipo de Acta";
            return PartialView();
        }

        // GET: Acta/Create
        //public ActionResult Create(int? tipoActa)
        //{
        //    //Obtener el codigo de usuario que se logeo
        //    var usuarioSesion = ViewData["usuario"] = System.Web.HttpContext.Current.Session["usuario"];
        //    ViewBag.UsuarioID = Convert.ToInt16(usuarioSesion);

        //    ViewBag.ActaTitulo = ActaEntity.GetNombreTipo(tipoActa.Value);
        //    ViewBag.CodigoActa = ActaEntity.GetCodigoTipo(tipoActa.Value);

        //    ViewBag.TipoActa = tipoActa.Value;

        //    return View();
        //}

        //// POST: Acta/Create
        //[HttpPost]
        //public ActionResult Create(Acta cabecera, DetallesActaParcial cuerpoActa, ActaInformacionAdicional piePaginaActa, int tipoActa)
        //{
        //    try
        //    {
        //        if (!ActaEntity.ValidacionRangosFechaInicioFin(cabecera.FechaInicio, cabecera.FechaFin))
        //            return Json(new { Resultado = new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeValidacionRangoFechasInicioFin } }, JsonRequestBehavior.AllowGet);
        //        if (tipoActa==1)
        //        {
        //            if (!ActaEntity.ValidacionHoraInicioFin(cabecera.HoraInicio, cabecera.HoraFin))
        //                return Json(new { Resultado = new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeValidacionRangoHoraInicioFin } }, JsonRequestBehavior.AllowGet);
        //        }

        //        RespuestaTransaccion resultado = ActaEntity.CrearActa(cabecera, cuerpoActa, piePaginaActa, tipoActa);
        //        return Json(new { Resultado = resultado }, JsonRequestBehavior.AllowGet);
        //    }
        //    catch (Exception ex)
        //    {
        //        return Json(new { Resultado = new RespuestaTransaccion { Estado = false, Respuesta = ex.Message } }, JsonRequestBehavior.AllowGet);
        //    }
        //}

        public ActionResult Create(int? tipoActa, int? idCliente, DateTime? fechaInicio, DateTime? fechaFin, int? ejecutivoID)
        {
            //Obtener el codigo de usuario que se logeo
            var usuarioSesion = ViewData["usuario"] = System.Web.HttpContext.Current.Session["usuario"];
            ViewBag.UsuarioID = Convert.ToInt16(usuarioSesion);

            ViewBag.ActaTitulo = ActaEntity.GetNombreTipo(tipoActa.Value);
            ViewBag.CodigoActa = ActaEntity.GetCodigoTipo(tipoActa.Value);

            ViewBag.idCliente = idCliente;
            ViewBag.fechaInicio = fechaInicio;
            ViewBag.fechaFin = fechaFin;
            if(tipoActa == 5)
            {
                ejecutivoID = null;
            }

            ViewBag.ListadoPrefacturas = idCliente.HasValue && fechaInicio.HasValue && fechaFin.HasValue ? CotizacionEntity.ListadoPrefacturaSAFI(idCliente, fechaInicio, fechaFin, ejecutivoID) : new List<PrefacturaSAFIInfo>();

            ViewBag.TipoActa = tipoActa.Value;

            return View();
        }

        // POST: Acta/Create
        [HttpPost]
        public ActionResult Create(Acta cabecera, DetallesActaParcial cuerpoActa, ActaInformacionAdicional piePaginaActa, int tipoActa)
        {
            RespuestaTransaccion resultado = new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeTransaccionFallida };
            try
            {
                string actaT = ActaEntity.GetCodigoTipo(tipoActa);

                switch (actaT)
                {
                    case "ARE":
                        cabecera.FechaFin = cabecera.FechaInicio;

                        if (!ActaEntity.ValidacionHoraInicioFin(cabecera.HoraInicio, cabecera.HoraFin))
                            return Json(new { Resultado = new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeValidacionRangoHoraInicioFin } }, JsonRequestBehavior.AllowGet);

                        string validacionAcuerdos = ActaEntity.ValidarFechaAcuerdos(cuerpoActa.Acuerdos, cabecera.FechaInicio);
                        if (!string.IsNullOrEmpty(validacionAcuerdos))
                            return Json(new { Resultado = new RespuestaTransaccion { Estado = false, Respuesta = validacionAcuerdos } }, JsonRequestBehavior.AllowGet);
                        if (!ActaEntity.ValidacionHoraInicioFinDiferente(cabecera.HoraInicio, cabecera.HoraFin))
                            return Json(new { Resultado = new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeValidacionRangoHoraInicioFinDiferente } }, JsonRequestBehavior.AllowGet);

                        resultado = ActaEntity.CrearActa(cabecera, cuerpoActa, piePaginaActa, tipoActa);
                        break;
                    case "AIP":
                    case "ACP":
                        if (!ActaEntity.ValidacionRangosFechaInicioFin(cabecera.FechaInicio, cabecera.FechaFin))
                            return Json(new { Resultado = new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeValidacionRangoFechasInicioFin } }, JsonRequestBehavior.AllowGet);
                        if (!ActaEntity.ValidacionRangosFechaFinEntrega(cabecera.FechaFin, cabecera.FechaEntrega.Value))
                            return Json(new { Resultado = new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeValidacionRangoFechasFinEntrega } }, JsonRequestBehavior.AllowGet);
                        if (cabecera.FechaEntrega.HasValue)
                            if (!ActaEntity.ValidacionFechaEntrega(cabecera.FechaInicio, cabecera.FechaEntrega.Value))
                                return Json(new { Resultado = new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeValidacionFechaEntrega } }, JsonRequestBehavior.AllowGet);

                        resultado = ActaEntity.CrearActa(cabecera, cuerpoActa, piePaginaActa, tipoActa);
                        break;
                    case "AECE":
                    case "AECF":
                        var verificarDuplicadosCliente = cuerpoActa.DetalleCliente.GroupBy(x => x.id_facturacion_safi).Any(g => g.Count() > 1);
                        var verificarDuplicadosContabilidad = cuerpoActa.DetalleContabilidad.GroupBy(x => x.id_facturacion_safi).Any(g => g.Count() > 1);

                        if (verificarDuplicadosCliente || verificarDuplicadosContabilidad)
                            return Json(new { Resultado = new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeErrorItemsRepetidos } }, JsonRequestBehavior.AllowGet);

                        resultado = ActaEntity.CrearActa(cabecera, cuerpoActa, piePaginaActa, tipoActa);
                        break;
                    default:
                        break;
                }

                //var archivosGenerados = ExportarExcel(Convert.ToInt32(resultado.EntidadID.Value));
                var archivosGenerados = GenerarArchivosActas(Convert.ToInt32(resultado.EntidadID.Value));

                return Json(new { Resultado = resultado }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                return Json(new { Resultado = new RespuestaTransaccion { Estado = false, Respuesta = ex.Message } }, JsonRequestBehavior.AllowGet);
            }
        }

        // GET: Acta/Edit/5
        public ActionResult Edit(int id)
        {
            var informacionCompleta = ActaEntity.ConsultarActaInformacionCompleta(id);
            ActaCompleta acta = new ActaCompleta(informacionCompleta);

            ViewBag.ActaTitulo = ActaEntity.GetNombreTipo(acta.Cabecera.TipoActaID);
            ViewBag.TipoActa = acta.Cabecera.TipoActaID;

            string CodigoActa = ActaEntity.GetCodigoTipo(acta.Cabecera.TipoActaID);
            ViewBag.CodigoActa = CodigoActa;

            var actas = acta.Cuerpo.DetalleCliente;

            int filtroClienteID = 0;
            DateTime filtroFechaInicio = DateTime.Now;
            DateTime filtroFechaFin = DateTime.Now;

            if (CodigoActa == "AECE" || CodigoActa == "AECF")
            {
                filtroClienteID = informacionCompleta.FirstOrDefault().ClienteID_ActaCliente.Value;
                filtroFechaInicio = CodigoActa == "AECE" ? informacionCompleta.OrderBy(t => t.fecha_prefactura_ActaCliente).First().fecha_prefactura_ActaCliente.Value : informacionCompleta.OrderBy(t => t.fecha_prefactura_ActaContabilidad).First().fecha_prefactura_ActaContabilidad.Value;
                filtroFechaFin = CodigoActa == "AECE" ? informacionCompleta.OrderByDescending(t => t.fecha_prefactura_ActaCliente).First().fecha_prefactura_ActaCliente.Value : informacionCompleta.OrderByDescending(t => t.fecha_prefactura_ActaContabilidad).First().fecha_prefactura_ActaContabilidad.Value;
            }

            ViewBag.ListadoPrefacturas = CotizacionEntity.ListadoPrefacturaSAFI(filtroClienteID, filtroFechaInicio, filtroFechaFin);

            return View(acta);
        }

        // POST: Acta/Edit/5
        [HttpPost]
        public ActionResult Edit(Acta cabecera, DetallesActaParcial cuerpoActa, ActaInformacionAdicional piePaginaActa)
        {
            RespuestaTransaccion resultado = new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeTransaccionFallida };
            try
            {
                var informacionCompleta = ActaEntity.ConsultarActaInformacionCompleta(cabecera.IDActa);
                ActaCompleta acta = new ActaCompleta(informacionCompleta);

                switch (acta.Cabecera.CodigoTipoActa)
                {
                    case "ARE":
                        string validacionAcuerdos = ActaEntity.ValidarFechaAcuerdos(cuerpoActa.Acuerdos, cabecera.FechaInicio);
                        if (!string.IsNullOrEmpty(validacionAcuerdos))
                            return Json(new { Resultado = new RespuestaTransaccion { Estado = false, Respuesta = validacionAcuerdos } }, JsonRequestBehavior.AllowGet);
                        if (!ActaEntity.ValidacionHoraInicioFinDiferente(cabecera.HoraInicio, cabecera.HoraFin))
                            return Json(new { Resultado = new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeValidacionRangoHoraInicioFinDiferente } }, JsonRequestBehavior.AllowGet);

                        cabecera.FechaFin = cabecera.FechaInicio;

                        if (!ActaEntity.ValidacionHoraInicioFin(cabecera.HoraInicio, cabecera.HoraFin))
                            return Json(new { Resultado = new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeValidacionRangoHoraInicioFin } }, JsonRequestBehavior.AllowGet);

                        resultado = ActaEntity.ActualizarActa(cabecera, cuerpoActa, piePaginaActa);
                        break;
                    case "AIP":
                    case "ACP":
                        if (cabecera.FechaEntrega.HasValue)
                            if (!ActaEntity.ValidacionFechaEntrega(cabecera.FechaInicio, cabecera.FechaEntrega.Value))
                                return Json(new { Resultado = new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeValidacionFechaEntrega } }, JsonRequestBehavior.AllowGet);
                        if (!ActaEntity.ValidacionRangosFechaInicioFin(cabecera.FechaInicio, cabecera.FechaFin))
                            return Json(new { Resultado = new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeValidacionRangoFechasInicioFin } }, JsonRequestBehavior.AllowGet);
                        if (!ActaEntity.ValidacionRangosFechaFinEntrega(cabecera.FechaFin, cabecera.FechaEntrega.Value))
                            return Json(new { Resultado = new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeValidacionRangoFechasFinEntrega } }, JsonRequestBehavior.AllowGet);

                        resultado = ActaEntity.ActualizarActa(cabecera, cuerpoActa, piePaginaActa);
                        break;
                    case "AECE":
                    case "AECF":
                        var verificarDuplicadosCliente = cuerpoActa.DetalleCliente.GroupBy(x => x.id_facturacion_safi).Any(g => g.Count() > 1);
                        var verificarDuplicadosContabilidad = cuerpoActa.DetalleContabilidad.GroupBy(x => x.id_facturacion_safi).Any(g => g.Count() > 1);

                        if (verificarDuplicadosCliente || verificarDuplicadosContabilidad)
                            return Json(new { Resultado = new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeErrorItemsRepetidos } }, JsonRequestBehavior.AllowGet);

                        resultado = ActaEntity.ActualizarActa(cabecera, cuerpoActa, piePaginaActa);
                        break;
                    default:
                        break;
                }

                //var archivosGenerados = ExportarExcel(Convert.ToInt32(cabecera.IDActa));
                var archivosGenerados = GenerarArchivosActas(Convert.ToInt32(cabecera.IDActa));

                return Json(new { Resultado = resultado }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                return Json(new { Resultado = new RespuestaTransaccion { Estado = false, Respuesta = ex.Message } }, JsonRequestBehavior.AllowGet);
            }
        }

        [HttpPost]
        public ActionResult Eliminar(long id)
        {
            if (id == 0)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            RespuestaTransaccion resultado = ActaEntity.EliminarActa(id);

            //return RedirectToAction("Index");
            return Json(new { Resultado = resultado }, JsonRequestBehavior.AllowGet);
        }

        public JsonResult _GetCodigosCotizacion(string busqueda)
        {
            List<AutoCompleteUI> items = new List<AutoCompleteUI>();
            busqueda = (busqueda ?? "").ToLower().Trim();

            var codigoCotizacion = CodigoCotizacionEntity.ListarCodigoCotizacion();

            items = codigoCotizacion.Where(o => o.codigo_cotizacion.ToLower().Contains(busqueda)).Select(o => new AutoCompleteUI(Convert.ToInt64(o.id_codigo_cotizacion), o.codigo_cotizacion, string.Empty, new Dictionary<string, CodigoCotizacionInfo>(){{ o.id_codigo_cotizacion.ToString(), new CodigoCotizacionInfo { nombre_comercial_cliente=o.nombre_comercial_cliente, nombre_proyecto=o.nombre_proyecto, descripcion_proyecto=o.descripcion_proyecto } }
            })).Take(10).ToList(); //new List<string>{ o.nombre_comercial_cliente, o.nombre_proyecto, o.descripcion_proyecto }
            return Json(new { results = items }, JsonRequestBehavior.AllowGet);
        }

        public JsonResult _GetClientes(string busqueda)
        {
            List<AutoCompleteUI> items = new List<AutoCompleteUI>();
            busqueda = (busqueda ?? "").ToLower().Trim();

            var clientes = ClienteEntity.ListarCliente();

            items = ClienteEntity.ListarCliente().Where(o => o.Nombre_Comercial.ToLower().Contains(busqueda)).Select(o => new AutoCompleteUI(Convert.ToInt64(o.Id), o.Nombre_Comercial, "")).Take(10).ToList();
            return Json(new { results = items }, JsonRequestBehavior.AllowGet);
        }

        public ActionResult _PrefacturaInformacion(int id)
        {
            var modelo = CotizacionEntity.ConsultarPrefacturaSAFI(id);
            //Titulo de la Pantalla
            ViewBag.TituloModal = Etiquetas.TituloPanelInformacionPrefactura + " " + (modelo.numero_prefactura ?? string.Empty);
            return PartialView(modelo);
        }

        public int Pixel2MTU(int pixels)
        {
            int mtus = pixels * 9525;
            return mtus;
        }

        #region REPORTE EXCEL
        private bool GenerarArchivosActas(int id)
        {
            bool resultado = false;
            try
            {

                var package = new ExcelPackage();

                var informacionCompleta = ActaEntity.ConsultarActaInformacionCompleta(id);
                ActaCompleta acta = new ActaCompleta(informacionCompleta);
                string actaTitulo = ActaEntity.GetNombreTipo(acta.Cabecera.TipoActaID);

                //PALETA DE COLORES PPM
                var colorGrisOscuroEstiloPPM = Color.FromArgb(60, 66, 87);
                var colorGrisClaroEstiloPPM = Color.FromArgb(240, 240, 240);
                var colorGrisClaro2EstiloPPM = Color.FromArgb(112, 117, 128);
                var colorGrisClaro3EstiloPPM = Color.FromArgb(225, 225, 225);
                var colorBlancoEstiloPPM = Color.FromArgb(255, 255, 255);
                var colorNegroEstiloPPM = Color.FromArgb(0, 0, 0);

                #region Cabecera

                var ws = package.Workbook.Worksheets.Add(acta.Cabecera.CodigoActa);

                int columnaFinalDocumentoExcel = 9;
                int columnaInicialDocumentoExcel = 1;

                ws.PrinterSettings.PaperSize = ePaperSize.A4;//ePaperSize.A3;
                ws.PrinterSettings.Orientation = acta.Cabecera.CodigoTipoActa == "AECE" || acta.Cabecera.CodigoTipoActa == "AECF" ? eOrientation.Landscape : eOrientation.Portrait;
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

                System.Drawing.Image img = System.Drawing.Image.FromFile(@pathUbicacion);
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

                // TITULO ACTA
                using (var range = ws.Cells[1, 5, 2, 8])
                {
                    range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                    range.Style.Fill.BackgroundColor.SetColor(colorGrisOscuroEstiloPPM);

                    range.Style.Font.Color.SetColor(colorBlancoEstiloPPM);

                    range.Value = actaTitulo;
                    //range.Style.Font.Bold = true;
                    range.Style.Font.Size = 18;
                    range.Style.Font.Name = "Raleway";
                    range.Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                    range.Merge = true;
                }

                // CODIGO ACTA
                using (var range = ws.Cells[1, 9, 2, 9])
                {
                    range.Style.HorizontalAlignment = ExcelHorizontalAlignment.CenterContinuous;
                    range.Style.Fill.BackgroundColor.SetColor(colorGrisClaroEstiloPPM);

                    range.Style.Font.Color.SetColor(colorGrisOscuroEstiloPPM);

                    range.Value = acta.Cabecera.CodigoActa;//actaTitulo + " " + acta.Cabecera.CodigoActa;
                    //range.Style.Font.Bold = true;
                    range.Style.Font.Size = 8;
                    range.Style.Font.Name = "Raleway";
                    range.Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                    range.Merge = true;
                }

                //FORMATO FUENTE TEXTO DE TODO EL DOCUMENTO

                int finCabecera = 4;

                if (acta.Cabecera.CodigoTipoActa == "AECE" || acta.Cabecera.CodigoTipoActa == "AECF")
                {
                    int totalRegistros = 0;

                    if (acta.Cabecera.CodigoTipoActa == "AECE")
                        totalRegistros = acta.Cuerpo.DetalleCliente.Count;
                    else
                        totalRegistros = acta.Cuerpo.DetalleContabilidad.Count;

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

                        ws.Cells[range.Start.Row, range.Columns + 1, range.End.Row, range.Columns + 4].Merge = true;

                        ws.Cells[range.Start.Row, range.Columns + 1, range.End.Row, range.Columns + 4].Style.Border.Top.Style = ExcelBorderStyle.Hair;
                        ws.Cells[range.Start.Row, range.Columns + 1, range.End.Row, range.Columns + 4].Style.Border.Left.Style = ExcelBorderStyle.Hair;
                        ws.Cells[range.Start.Row, range.Columns + 1, range.End.Row, range.Columns + 4].Style.Border.Right.Style = ExcelBorderStyle.Hair;
                        ws.Cells[range.Start.Row, range.Columns + 1, range.End.Row, range.Columns + 4].Style.Border.Bottom.Style = ExcelBorderStyle.Hair;

                        ws.Cells[range.Start.Row, range.Columns + 1].Value = acta.Cabecera.NombresElaboradoPor; // Valor campo
                        ws.Cells[range.Start.Row, range.Columns + 1].Style.Indent = 1;

                    }

                    using (var range = ws.Cells[finCabecera, 7, finCabecera, 8])
                    {
                        range.Style.Border.Top.Style = ExcelBorderStyle.Hair;
                        range.Style.Border.Left.Style = ExcelBorderStyle.Hair;
                        range.Style.Border.Right.Style = ExcelBorderStyle.Hair;
                        range.Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                        range.Merge = true;

                        range.Value = "Fecha de Creación:";
                        range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;


                        range.Style.Fill.BackgroundColor.SetColor(colorGrisOscuroEstiloPPM);
                        range.Style.Font.Color.SetColor(colorBlancoEstiloPPM);
                        range.Style.Font.Bold = true;

                        ws.Cells[range.Start.Row, range.End.Column + 1, range.End.Row, columnaFinalDocumentoExcel].Merge = true;

                        ws.Cells[range.Start.Row, range.End.Column + 1, range.End.Row, columnaFinalDocumentoExcel].Style.Border.Top.Style = ExcelBorderStyle.Hair;
                        ws.Cells[range.Start.Row, range.End.Column + 1, range.End.Row, columnaFinalDocumentoExcel].Style.Border.Left.Style = ExcelBorderStyle.Hair;
                        ws.Cells[range.Start.Row, range.End.Column + 1, range.End.Row, columnaFinalDocumentoExcel].Style.Border.Right.Style = ExcelBorderStyle.Hair;
                        ws.Cells[range.Start.Row, range.End.Column + 1, range.End.Row, columnaFinalDocumentoExcel].Style.Border.Bottom.Style = ExcelBorderStyle.Hair;

                        ws.Cells[range.Start.Row, range.End.Column + 1].Value = acta.Cabecera.FechaCreacion.ToString("yyyy/MM/dd"); // Valor campo
                        ws.Cells[range.Start.Row, range.End.Column + 1].Style.Indent = 1;
                    }

                    finCabecera++;

                    using (var range = ws.Cells[finCabecera, 1, finCabecera, 2])
                    {
                        range.Style.Border.Top.Style = ExcelBorderStyle.Hair;
                        range.Style.Border.Left.Style = ExcelBorderStyle.Hair;
                        range.Style.Border.Right.Style = ExcelBorderStyle.Hair;
                        range.Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                        range.Merge = true;

                        range.Value = "Código:";

                        range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;


                        range.Style.Fill.BackgroundColor.SetColor(colorGrisOscuroEstiloPPM);
                        range.Style.Font.Color.SetColor(colorBlancoEstiloPPM);
                        range.Style.Font.Bold = true;

                        ws.Cells[range.Start.Row, range.Columns + 1, range.End.Row, range.Columns + 4].Merge = true;

                        ws.Cells[range.Start.Row, range.Columns + 1, range.End.Row, range.Columns + 4].Style.Border.Top.Style = ExcelBorderStyle.Hair;
                        ws.Cells[range.Start.Row, range.Columns + 1, range.End.Row, range.Columns + 4].Style.Border.Left.Style = ExcelBorderStyle.Hair;
                        ws.Cells[range.Start.Row, range.Columns + 1, range.End.Row, range.Columns + 4].Style.Border.Right.Style = ExcelBorderStyle.Hair;
                        ws.Cells[range.Start.Row, range.Columns + 1, range.End.Row, range.Columns + 4].Style.Border.Bottom.Style = ExcelBorderStyle.Hair;

                        ws.Cells[range.Start.Row, range.Columns + 1].Value = "GCM-GCO-PRO-001-FOR-002"; // Valor campo
                        ws.Cells[range.Start.Row, range.Columns + 1].Style.Indent = 1;

                    }

                    using (var range = ws.Cells[finCabecera, 7, finCabecera, 8])
                    {
                        range.Style.Border.Top.Style = ExcelBorderStyle.Hair;
                        range.Style.Border.Left.Style = ExcelBorderStyle.Hair;
                        range.Style.Border.Right.Style = ExcelBorderStyle.Hair;
                        range.Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                        range.Merge = true;

                        range.Value = "N. Registros:";
                        range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;


                        range.Style.Fill.BackgroundColor.SetColor(colorGrisOscuroEstiloPPM);
                        range.Style.Font.Color.SetColor(colorBlancoEstiloPPM);
                        range.Style.Font.Bold = true;

                        ws.Cells[range.Start.Row, range.End.Column + 1, range.End.Row, columnaFinalDocumentoExcel].Merge = true;

                        ws.Cells[range.Start.Row, range.End.Column + 1, range.End.Row, columnaFinalDocumentoExcel].Style.Border.Top.Style = ExcelBorderStyle.Hair;
                        ws.Cells[range.Start.Row, range.End.Column + 1, range.End.Row, columnaFinalDocumentoExcel].Style.Border.Left.Style = ExcelBorderStyle.Hair;
                        ws.Cells[range.Start.Row, range.End.Column + 1, range.End.Row, columnaFinalDocumentoExcel].Style.Border.Right.Style = ExcelBorderStyle.Hair;
                        ws.Cells[range.Start.Row, range.End.Column + 1, range.End.Row, columnaFinalDocumentoExcel].Style.Border.Bottom.Style = ExcelBorderStyle.Hair;

                        ws.Cells[range.Start.Row, range.End.Column + 1].Value = totalRegistros; // Valor campo
                        ws.Cells[range.Start.Row, range.End.Column + 1].Style.Indent = 1;
                        ws.Cells[range.Start.Row, range.End.Column + 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    }





                    finCabecera += 2;
                    ws.Row(finCabecera).Height = 20.25;








                }
                else
                {
                    if (acta.Cabecera.CodigoTipoActa == "ARE")
                    {
                        // ws.Row(6).Height = MeasureTextHeight(ws.Cells[6, 3, 6, 6].Value.ToString(), ws.Cells[6, 3, 6, 6].Style.Font, 4);
                        //FECHA INICIO
                        using (var range = ws.Cells[5, 1, 5, 2])
                        {
                            range.Style.Border.Top.Style = ExcelBorderStyle.Hair;
                            range.Style.Border.Left.Style = ExcelBorderStyle.Hair;
                            range.Style.Border.Right.Style = ExcelBorderStyle.Hair;
                            range.Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                            range.Merge = true;

                            range.Value = "Fecha:";

                            range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                            range.Style.Fill.BackgroundColor.SetColor(colorGrisClaro3EstiloPPM);
                            range.Style.Font.Color.SetColor(colorGrisOscuroEstiloPPM);
                            range.Style.Font.Bold = true;

                            ws.Cells[range.Start.Row, range.Columns + 1, range.End.Row, range.Columns + 4].Merge = true;
                            ws.Cells[range.Start.Row, range.Columns + 1].Value = acta.Cabecera.FechaInicio.ToString("yyyy/MM/dd"); // Valor campo
                            ws.Cells[range.Start.Row, range.Columns + 1].Style.Indent = 1;

                        }

                        //Duracion
                        using (var range = ws.Cells[5, 7, 5, 8])
                        {
                            range.Style.Border.Top.Style = ExcelBorderStyle.Hair;
                            range.Style.Border.Left.Style = ExcelBorderStyle.Hair;
                            range.Style.Border.Right.Style = ExcelBorderStyle.Hair;
                            range.Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                            range.Merge = true;

                            range.Value = "Duración:";
                            range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

                            range.Style.Fill.BackgroundColor.SetColor(colorGrisClaro3EstiloPPM);
                            range.Style.Font.Color.SetColor(colorGrisOscuroEstiloPPM);
                            range.Style.Font.Bold = true;

                            ws.Cells[range.Start.Row, range.End.Column + 1, range.End.Row, columnaFinalDocumentoExcel].Merge = true;
                            ws.Cells[range.Start.Row, range.End.Column + 1].Value = ActaEntity.CalcularDuracion(acta.Cabecera.HoraInicio, acta.Cabecera.HoraFin, " " + "hora(s)");
                            ws.Cells[range.Start.Row, range.End.Column + 1].Style.Indent = 1;
                        }

                        //ELABORADO POR
                        using (var range = ws.Cells[6, 1, 6, 2])
                        {
                            range.Style.Border.Top.Style = ExcelBorderStyle.Hair;
                            range.Style.Border.Left.Style = ExcelBorderStyle.Hair;
                            range.Style.Border.Right.Style = ExcelBorderStyle.Hair;
                            range.Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                            range.Merge = true;

                            range.Value = "Lugar: ";
                            range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;


                            range.Style.Fill.BackgroundColor.SetColor(colorGrisOscuroEstiloPPM);
                            range.Style.Font.Color.SetColor(colorBlancoEstiloPPM);
                            range.Style.Font.Bold = true;

                            ws.Cells[range.Start.Row, range.Columns + 1, range.End.Row, range.Columns + 4].Merge = true;
                            ws.Cells[range.Start.Row, range.Columns + 1].Value = acta.Cabecera.Lugar; // Valor campo
                            ws.Cells[range.Start.Row, range.Columns + 1].Style.Indent = 1;
                        }
                    }
                    else
                    {

                        //FECHA INICIO
                        using (var range = ws.Cells[5, 1, 5, 2])
                        {
                            range.Style.Border.Top.Style = ExcelBorderStyle.Hair;
                            range.Style.Border.Left.Style = ExcelBorderStyle.Hair;
                            range.Style.Border.Right.Style = ExcelBorderStyle.Hair;
                            range.Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                            range.Merge = true;

                            range.Value = "Fecha de Creación:";

                            range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;


                            range.Style.Fill.BackgroundColor.SetColor(colorGrisOscuroEstiloPPM);
                            range.Style.Font.Color.SetColor(colorBlancoEstiloPPM);
                            range.Style.Font.Bold = true;

                            ws.Cells[range.Start.Row, range.Columns + 1, range.End.Row, range.Columns + 4].Merge = true;
                            ws.Cells[range.Start.Row, range.Columns + 1].Value = acta.Cabecera.FechaCreacion.ToString("yyyy/MM/dd"); // Valor campo
                            ws.Cells[range.Start.Row, range.Columns + 1].Style.Indent = 1;

                        }

                        //FECHA INICIO
                        using (var range = ws.Cells[5, 7, 5, 8])
                        {
                            range.Style.Border.Top.Style = ExcelBorderStyle.Hair;
                            range.Style.Border.Left.Style = ExcelBorderStyle.Hair;
                            range.Style.Border.Right.Style = ExcelBorderStyle.Hair;
                            range.Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                            range.Merge = true;

                            range.Value = "Fecha Inicio:";
                            range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;


                            range.Style.Fill.BackgroundColor.SetColor(colorGrisOscuroEstiloPPM);
                            range.Style.Font.Color.SetColor(colorBlancoEstiloPPM);
                            range.Style.Font.Bold = true;

                            ws.Cells[range.Start.Row, range.End.Column + 1, range.End.Row, columnaFinalDocumentoExcel].Merge = true;
                            ws.Cells[range.Start.Row, range.End.Column + 1].Value = acta.Cabecera.FechaInicio.ToString("yyyy/MM/dd"); // Valor campo
                            ws.Cells[range.Start.Row, range.End.Column + 1].Style.Indent = 1;
                        }

                        //ELABORADO POR
                        using (var range = ws.Cells[6, 1, 6, 2])
                        {
                            range.Style.Border.Top.Style = ExcelBorderStyle.Hair;
                            range.Style.Border.Left.Style = ExcelBorderStyle.Hair;
                            range.Style.Border.Right.Style = ExcelBorderStyle.Hair;
                            range.Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                            range.Merge = true;

                            range.Value = "Elaborado Por: ";

                            range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;


                            range.Style.Fill.BackgroundColor.SetColor(colorGrisClaro3EstiloPPM);
                            range.Style.Font.Color.SetColor(colorGrisOscuroEstiloPPM);
                            range.Style.Font.Bold = true;

                            ws.Cells[range.Start.Row, range.Columns + 1, range.End.Row, range.Columns + 4].Merge = true;
                            ws.Cells[range.Start.Row, range.Columns + 1].Value = acta.Cabecera.NombresElaboradoPor; // Valor campo
                            ws.Cells[range.Start.Row, range.Columns + 1].Style.Indent = 1;
                        }
                    }

                    int inicioParteCabeceraParte2 = 10;

                    switch (acta.Cabecera.CodigoTipoActa)
                    {
                        case "ARE":

                            //HoraInicio
                            using (var range = ws.Cells[6, 7, 6, 8])
                            {
                                range.Style.Border.Top.Style = ExcelBorderStyle.Hair;
                                range.Style.Border.Left.Style = ExcelBorderStyle.Hair;
                                range.Style.Border.Right.Style = ExcelBorderStyle.Hair;
                                range.Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                                range.Merge = true;

                                range.Value = "Hora de Inicio:";
                                range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

                                range.Style.Fill.BackgroundColor.SetColor(colorGrisOscuroEstiloPPM);
                                range.Style.Font.Color.SetColor(colorBlancoEstiloPPM);
                                range.Style.Font.Bold = true;

                                ws.Cells[range.Start.Row, range.End.Column + 1, range.End.Row, columnaFinalDocumentoExcel].Merge = true;
                                ws.Cells[range.Start.Row, range.End.Column + 1].Value = ActaEntity.CalcularHora(acta.Cabecera.HoraInicio); // Valor campo
                                ws.Cells[range.Start.Row, range.End.Column + 1].Style.Indent = 1;

                            }

                            //NÚMERO DE REUNIÓN
                            using (var range = ws.Cells[7, 1, 7, 2])
                            {
                                range.Style.Border.Top.Style = ExcelBorderStyle.Hair;
                                range.Style.Border.Left.Style = ExcelBorderStyle.Hair;
                                range.Style.Border.Right.Style = ExcelBorderStyle.Hair;
                                range.Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                                range.Merge = true;

                                range.Value = "N. Reunión: ";
                                range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

                                range.Style.Fill.BackgroundColor.SetColor(colorGrisClaro3EstiloPPM);
                                range.Style.Font.Color.SetColor(colorGrisOscuroEstiloPPM);
                                range.Style.Font.Bold = true;

                                ws.Cells[range.Start.Row, range.Columns + 1, range.End.Row, range.Columns + 4].Merge = true;
                                ws.Cells[range.Start.Row, range.Columns + 1].Value = acta.Cabecera.NumeroReunion;  // Valor campo
                                ws.Cells[range.Start.Row, range.Columns + 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;

                                ws.Cells[range.Start.Row, range.Columns + 1].Style.Indent = 1;
                            }

                            //HORA FIN
                            using (var range = ws.Cells[7, 7, 7, 8])
                            {
                                range.Style.Border.Top.Style = ExcelBorderStyle.Hair;
                                range.Style.Border.Left.Style = ExcelBorderStyle.Hair;
                                range.Style.Border.Right.Style = ExcelBorderStyle.Hair;
                                range.Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                                range.Merge = true;

                                range.Value = "Hora Fin:";
                                range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

                                range.Style.Fill.BackgroundColor.SetColor(colorGrisClaro3EstiloPPM);
                                range.Style.Font.Color.SetColor(colorGrisOscuroEstiloPPM);
                                range.Style.Font.Bold = true;

                                ws.Cells[range.Start.Row, range.End.Column + 1, range.End.Row, columnaFinalDocumentoExcel].Merge = true;
                                ws.Cells[range.Start.Row, range.End.Column + 1].Value = ActaEntity.CalcularHora(acta.Cabecera.HoraFin); // Valor campo
                                ws.Cells[range.Start.Row, range.End.Column + 1].Style.Indent = 1;
                            }


                            using (var range = ws.Cells[5, 3, 5, 8])
                            {
                                range.Style.Border.Top.Style = ExcelBorderStyle.Hair;
                                range.Style.Border.Left.Style = ExcelBorderStyle.Hair;
                                range.Style.Border.Right.Style = ExcelBorderStyle.Hair;
                                range.Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                            }
                            using (var range = ws.Cells[5, 9, 5, 9])
                            {
                                range.Style.Border.Top.Style = ExcelBorderStyle.Hair;
                                range.Style.Border.Left.Style = ExcelBorderStyle.Hair;
                                range.Style.Border.Right.Style = ExcelBorderStyle.Hair;
                                range.Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                            }

                            using (var range = ws.Cells[6, 5, 6, 8])
                            {
                                range.Style.Border.Top.Style = ExcelBorderStyle.Hair;
                                range.Style.Border.Left.Style = ExcelBorderStyle.Hair;
                                range.Style.Border.Right.Style = ExcelBorderStyle.Hair;
                                range.Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                            }

                            using (var range = ws.Cells[6, 9, 6, 9])
                            {
                                range.Style.Border.Top.Style = ExcelBorderStyle.Hair;
                                range.Style.Border.Left.Style = ExcelBorderStyle.Hair;
                                range.Style.Border.Right.Style = ExcelBorderStyle.Hair;
                                range.Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                            }
                            using (var range = ws.Cells[7, 1, 7, 2])
                            {
                                range.Style.Border.Top.Style = ExcelBorderStyle.Hair;
                                range.Style.Border.Left.Style = ExcelBorderStyle.Hair;
                                range.Style.Border.Right.Style = ExcelBorderStyle.Hair;
                                range.Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                            }
                            using (var range = ws.Cells[7, 3, 7, 9])
                            {
                                range.Style.Border.Top.Style = ExcelBorderStyle.Hair;
                                range.Style.Border.Left.Style = ExcelBorderStyle.Hair;
                                range.Style.Border.Right.Style = ExcelBorderStyle.Hair;
                                range.Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                            }
                            using (var range = ws.Cells[7, 9, 7, 9])
                            {
                                range.Style.Border.Top.Style = ExcelBorderStyle.Hair;
                                range.Style.Border.Left.Style = ExcelBorderStyle.Hair;
                                range.Style.Border.Right.Style = ExcelBorderStyle.Hair;
                                range.Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                            }
                            inicioParteCabeceraParte2 = 9;

                            ws.Row(3).Height = 4.12;
                            ws.Row(4).Height = 4.13;
                            ws.Row(5).Height = 20.25;
                            ws.Row(6).Height = 20.25;
                            ws.Row(7).Height = 20.25;
                            ws.Row(8).Height = 8.25;
                            ws.Row(9).Height = 20.25;
                            ws.Row(10).Height = 20.25;
                            ws.Row(16).Height = 8.25;
                            ws.Column(1).Width = 12;
                            break;

                        case "AIP":

                            using (var range = ws.Cells[6, 7, 6, 8])
                            {
                                range.Style.Border.Top.Style = ExcelBorderStyle.Hair;
                                range.Style.Border.Left.Style = ExcelBorderStyle.Hair;
                                range.Style.Border.Right.Style = ExcelBorderStyle.Hair;
                                range.Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                                range.Merge = true;

                                range.Value = "Fecha Fin:";
                                range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

                                range.Style.Fill.BackgroundColor.SetColor(colorGrisClaro3EstiloPPM);
                                range.Style.Font.Color.SetColor(colorGrisOscuroEstiloPPM);
                                range.Style.Font.Bold = true;

                                ws.Cells[range.Start.Row, range.End.Column + 1, range.End.Row, columnaFinalDocumentoExcel].Merge = true;
                                ws.Cells[range.Start.Row, range.End.Column + 1].Value = acta.Cabecera.FechaFin.ToString("yyyy/MM/dd"); // Valor campo
                                ws.Cells[range.Start.Row, range.End.Column + 1].Style.Indent = 1;

                            }

                            using (var range = ws.Cells[7, 1, 7, 2])
                            {
                                range.Style.Border.Top.Style = ExcelBorderStyle.Hair;
                                range.Style.Border.Left.Style = ExcelBorderStyle.Hair;
                                range.Style.Border.Right.Style = ExcelBorderStyle.Hair;
                                range.Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                                range.Merge = true;

                                range.Value = "Cargo: ";

                                range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;


                                range.Style.Fill.BackgroundColor.SetColor(colorGrisOscuroEstiloPPM);
                                range.Style.Font.Color.SetColor(colorBlancoEstiloPPM);
                                range.Style.Font.Bold = true;

                                ws.Cells[range.Start.Row, range.Columns + 1, range.End.Row, range.Columns + 4].Merge = true;
                                ws.Cells[range.Start.Row, range.Columns + 1].Value = acta.Cabecera.Cargo;  // Valor campo
                                ws.Cells[range.Start.Row, range.Columns + 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;

                                ws.Cells[range.Start.Row, range.Columns + 1].Style.Indent = 1;
                            }

                            //HORA Entrega
                            using (var range = ws.Cells[7, 7, 7, 8])
                            {
                                range.Style.Border.Top.Style = ExcelBorderStyle.Hair;
                                range.Style.Border.Left.Style = ExcelBorderStyle.Hair;
                                range.Style.Border.Right.Style = ExcelBorderStyle.Hair;
                                range.Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                                range.Merge = true;

                                range.Value = "Fecha Entrega:";

                                range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;


                                range.Style.Fill.BackgroundColor.SetColor(colorGrisOscuroEstiloPPM);
                                range.Style.Font.Color.SetColor(colorBlancoEstiloPPM);
                                range.Style.Font.Bold = true;

                                ws.Cells[range.Start.Row, range.End.Column + 1, range.End.Row, columnaFinalDocumentoExcel].Merge = true;
                                ws.Cells[range.Start.Row, range.End.Column + 1].Value = acta.Cabecera.FechaEntrega.Value.ToString("yyyy/MM/dd"); ; ; // Valor campo

                                ws.Cells[range.Start.Row, range.End.Column + 1].Style.Indent = 1;
                            }

                            using (var range = ws.Cells[5, 3, 5, 8])
                            {
                                range.Style.Border.Top.Style = ExcelBorderStyle.Hair;
                                range.Style.Border.Left.Style = ExcelBorderStyle.Hair;
                                range.Style.Border.Right.Style = ExcelBorderStyle.Hair;
                                range.Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                            }

                            using (var range = ws.Cells[5, 1, 5, 2])
                            {
                                range.Style.Border.Top.Style = ExcelBorderStyle.Hair;
                                range.Style.Border.Left.Style = ExcelBorderStyle.Hair;
                                range.Style.Border.Right.Style = ExcelBorderStyle.Hair;
                                range.Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                            }

                            using (var range = ws.Cells[6, 3, 6, 8])
                            {
                                range.Style.Border.Top.Style = ExcelBorderStyle.Hair;
                                range.Style.Border.Left.Style = ExcelBorderStyle.Hair;
                                range.Style.Border.Right.Style = ExcelBorderStyle.Hair;
                                range.Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                            }
                            using (var range = ws.Cells[6, 9, 6, 9])
                            {
                                range.Style.Border.Top.Style = ExcelBorderStyle.Hair;
                                range.Style.Border.Left.Style = ExcelBorderStyle.Hair;
                                range.Style.Border.Right.Style = ExcelBorderStyle.Hair;
                                range.Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                            }

                            using (var range = ws.Cells[5, 9, 5, 9])
                            {
                                range.Style.Border.Top.Style = ExcelBorderStyle.Hair;
                                range.Style.Border.Left.Style = ExcelBorderStyle.Hair;
                                range.Style.Border.Right.Style = ExcelBorderStyle.Hair;
                                range.Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                            }
                            using (var range = ws.Cells[7, 1, 7, 2])
                            {
                                range.Style.Border.Top.Style = ExcelBorderStyle.Hair;
                                range.Style.Border.Left.Style = ExcelBorderStyle.Hair;
                                range.Style.Border.Right.Style = ExcelBorderStyle.Hair;
                                range.Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                            }
                            using (var range = ws.Cells[7, 3, 7, 9])
                            {
                                range.Style.Border.Top.Style = ExcelBorderStyle.Hair;
                                range.Style.Border.Left.Style = ExcelBorderStyle.Hair;
                                range.Style.Border.Right.Style = ExcelBorderStyle.Hair;
                                range.Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                            }
                            using (var range = ws.Cells[7, 9, 7, 9])
                            {
                                range.Style.Border.Top.Style = ExcelBorderStyle.Hair;
                                range.Style.Border.Left.Style = ExcelBorderStyle.Hair;
                                range.Style.Border.Right.Style = ExcelBorderStyle.Hair;
                                range.Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                            }

                            inicioParteCabeceraParte2 = 9;

                            ws.Row(3).Height = 4.12;
                            ws.Row(4).Height = 4.13;
                            ws.Row(5).Height = 20.25;
                            ws.Row(6).Height = 20.25;
                            ws.Row(7).Height = 20.25;
                            ws.Row(8).Height = 8.25;
                            ws.Row(9).Height = 20.25;
                            ws.Row(10).Height = 20.25;
                            ws.Row(11).Height = 44.25;
                            ws.Row(12).Height = 44.25;
                            ws.Row(13).Height = 8.25;
                            ws.Row(15).Height = 8.25;
                            ws.Row(17).Height = 8.25;



                            break;
                        case "ACP":

                            //FECHA FIN
                            using (var range = ws.Cells[6, 7, 6, 8])
                            {
                                range.Style.Border.Top.Style = ExcelBorderStyle.Hair;
                                range.Style.Border.Left.Style = ExcelBorderStyle.Hair;
                                range.Style.Border.Right.Style = ExcelBorderStyle.Hair;
                                range.Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                                range.Merge = true;

                                range.Value = "Fecha Fin:";

                                range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

                                range.Style.Fill.BackgroundColor.SetColor(colorGrisClaro3EstiloPPM);
                                range.Style.Font.Color.SetColor(colorGrisOscuroEstiloPPM);
                                range.Style.Font.Bold = true;

                                ws.Cells[range.Start.Row, range.End.Column + 1, range.End.Row, columnaFinalDocumentoExcel].Merge = true;
                                ws.Cells[range.Start.Row, range.End.Column + 1].Value = acta.Cabecera.FechaFin.ToString("yyyy/MM/dd"); // Valor campo

                                ws.Cells[range.Start.Row, range.End.Column + 1].Style.Indent = 1;
                            }
                            using (var range = ws.Cells[7, 1, 7, 2])
                            {
                                range.Style.Border.Top.Style = ExcelBorderStyle.Hair;
                                range.Style.Border.Left.Style = ExcelBorderStyle.Hair;
                                range.Style.Border.Right.Style = ExcelBorderStyle.Hair;
                                range.Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                                range.Merge = true;

                                range.Value = "Cargo: ";
                                range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

                                range.Style.Fill.BackgroundColor.SetColor(colorGrisOscuroEstiloPPM);
                                range.Style.Font.Color.SetColor(colorBlancoEstiloPPM);
                                range.Style.Font.Bold = true;

                                ws.Cells[range.Start.Row, range.Columns + 1, range.End.Row, range.Columns + 4].Merge = true;
                                ws.Cells[range.Start.Row, range.Columns + 1].Value = acta.Cabecera.Cargo;  // Valor campo
                                ws.Cells[range.Start.Row, range.Columns + 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;

                                ws.Cells[range.Start.Row, range.Columns + 1].Style.Indent = 1;
                            }
                            ////HORA ENTREGA
                            using (var range = ws.Cells[7, 7, 7, 8])
                            {
                                range.Style.Border.Top.Style = ExcelBorderStyle.Hair;
                                range.Style.Border.Left.Style = ExcelBorderStyle.Hair;
                                range.Style.Border.Right.Style = ExcelBorderStyle.Hair;
                                range.Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                                range.Merge = true;
                                range.Value = "Fecha Entrega:";
                                range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                                range.Style.Fill.BackgroundColor.SetColor(colorGrisOscuroEstiloPPM);
                                range.Style.Font.Color.SetColor(colorBlancoEstiloPPM);
                                range.Style.Font.Bold = true;

                                ws.Cells[range.Start.Row, range.End.Column + 1, range.End.Row, columnaFinalDocumentoExcel].Merge = true;
                                ws.Cells[range.Start.Row, range.End.Column + 1].Value = acta.Cabecera.FechaEntrega.Value.ToString("yyyy/MM/dd"); // Valor campo
                                ws.Cells[range.Start.Row, range.End.Column + 1].Style.Indent = 1;
                            }


                            using (var range = ws.Cells[5, 3, 5, 8])
                            {
                                range.Style.Border.Top.Style = ExcelBorderStyle.Hair;
                                range.Style.Border.Left.Style = ExcelBorderStyle.Hair;
                                range.Style.Border.Right.Style = ExcelBorderStyle.Hair;
                                range.Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                            }

                            using (var range = ws.Cells[5, 1, 5, 2])
                            {
                                range.Style.Border.Top.Style = ExcelBorderStyle.Hair;
                                range.Style.Border.Left.Style = ExcelBorderStyle.Hair;
                                range.Style.Border.Right.Style = ExcelBorderStyle.Hair;
                                range.Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                            }

                            using (var range = ws.Cells[6, 3, 6, 8])
                            {
                                range.Style.Border.Top.Style = ExcelBorderStyle.Hair;
                                range.Style.Border.Left.Style = ExcelBorderStyle.Hair;
                                range.Style.Border.Right.Style = ExcelBorderStyle.Hair;
                                range.Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                            }
                            using (var range = ws.Cells[6, 9, 6, 9])
                            {
                                range.Style.Border.Top.Style = ExcelBorderStyle.Hair;
                                range.Style.Border.Left.Style = ExcelBorderStyle.Hair;
                                range.Style.Border.Right.Style = ExcelBorderStyle.Hair;
                                range.Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                            }

                            using (var range = ws.Cells[5, 9, 5, 9])
                            {
                                range.Style.Border.Top.Style = ExcelBorderStyle.Hair;
                                range.Style.Border.Left.Style = ExcelBorderStyle.Hair;
                                range.Style.Border.Right.Style = ExcelBorderStyle.Hair;
                                range.Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                            }
                            using (var range = ws.Cells[7, 1, 7, 2])
                            {
                                range.Style.Border.Top.Style = ExcelBorderStyle.Hair;
                                range.Style.Border.Left.Style = ExcelBorderStyle.Hair;
                                range.Style.Border.Right.Style = ExcelBorderStyle.Hair;
                                range.Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                            }
                            using (var range = ws.Cells[7, 3, 7, 9])
                            {
                                range.Style.Border.Top.Style = ExcelBorderStyle.Hair;
                                range.Style.Border.Left.Style = ExcelBorderStyle.Hair;
                                range.Style.Border.Right.Style = ExcelBorderStyle.Hair;
                                range.Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                            }
                            using (var range = ws.Cells[7, 9, 7, 9])
                            {
                                range.Style.Border.Top.Style = ExcelBorderStyle.Hair;
                                range.Style.Border.Left.Style = ExcelBorderStyle.Hair;
                                range.Style.Border.Right.Style = ExcelBorderStyle.Hair;
                                range.Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                            }

                            inicioParteCabeceraParte2 = 9;

                            ws.Row(3).Height = 4.12;
                            ws.Row(4).Height = 4.13;
                            ws.Row(5).Height = 20.25;
                            ws.Row(6).Height = 20.25;
                            ws.Row(7).Height = 20.25;
                            ws.Row(8).Height = 8.25;
                            ws.Row(9).Height = 20.25;
                            ws.Row(10).Height = 20.25;
                            ws.Row(11).Height = 44.25;
                            ws.Row(12).Height = 44.25;
                            ws.Row(13).Height = 8.25;
                            ws.Row(15).Height = 8.25;
                            ws.Row(17).Height = 8.25;

                            break;
                        default:
                            break;
                    }

                    //CodigoCotizacion
                    using (var range = ws.Cells[inicioParteCabeceraParte2, 1, inicioParteCabeceraParte2, 2])
                    {
                        range.Style.Border.Top.Style = ExcelBorderStyle.Hair;
                        range.Style.Border.Left.Style = ExcelBorderStyle.Hair;
                        range.Style.Border.Right.Style = ExcelBorderStyle.Hair;
                        range.Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                        range.Merge = true;

                        range.Value = "Código de Cotización: ";
                        range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

                        range.Style.Fill.BackgroundColor.SetColor(colorGrisClaro3EstiloPPM);
                        range.Style.Font.Color.SetColor(colorGrisOscuroEstiloPPM);
                        range.Style.Font.Bold = true;

                        ws.Cells[range.Start.Row, range.Columns + 1, range.End.Row, columnaFinalDocumentoExcel].Merge = true;
                        ws.Cells[range.Start.Row, range.Columns + 1].Value = acta.Cabecera.CodigoCotizacion;

                        ws.Cells[range.Start.Row, range.Columns + 1].Style.Indent = 1;

                        ws.Cells[range.Start.Row, range.Columns + 1, range.End.Row, columnaFinalDocumentoExcel].Style.Border.Top.Style = ExcelBorderStyle.Hair;
                        ws.Cells[range.Start.Row, range.Columns + 1, range.End.Row, columnaFinalDocumentoExcel].Style.Border.Left.Style = ExcelBorderStyle.Hair;
                        ws.Cells[range.Start.Row, range.Columns + 1, range.End.Row, columnaFinalDocumentoExcel].Style.Border.Right.Style = ExcelBorderStyle.Hair;
                        ws.Cells[range.Start.Row, range.Columns + 1, range.End.Row, columnaFinalDocumentoExcel].Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                    }

                    inicioParteCabeceraParte2++;

                    using (var range = ws.Cells[inicioParteCabeceraParte2, 1, inicioParteCabeceraParte2, 2])
                    {
                        range.Style.Border.Top.Style = ExcelBorderStyle.Hair;
                        range.Style.Border.Left.Style = ExcelBorderStyle.Hair;
                        range.Style.Border.Right.Style = ExcelBorderStyle.Hair;
                        range.Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                        range.Merge = true;

                        range.Value = "Cliente: ";
                        range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

                        range.Style.Fill.BackgroundColor.SetColor(colorGrisClaro3EstiloPPM);
                        range.Style.Font.Color.SetColor(colorGrisOscuroEstiloPPM);
                        range.Style.Font.Bold = true;

                        ws.Cells[range.Start.Row, range.Columns + 1, range.End.Row, columnaFinalDocumentoExcel].Merge = true;
                        ws.Cells[range.Start.Row, range.Columns + 1].Value = acta.Cabecera.Cliente;
                        ws.Cells[range.Start.Row, range.Columns + 1].Style.Indent = 1;

                        ws.Cells[range.Start.Row, range.Columns + 1, range.End.Row, columnaFinalDocumentoExcel].Style.Border.Top.Style = ExcelBorderStyle.Hair;
                        ws.Cells[range.Start.Row, range.Columns + 1, range.End.Row, columnaFinalDocumentoExcel].Style.Border.Left.Style = ExcelBorderStyle.Hair;
                        ws.Cells[range.Start.Row, range.Columns + 1, range.End.Row, columnaFinalDocumentoExcel].Style.Border.Right.Style = ExcelBorderStyle.Hair;
                        ws.Cells[range.Start.Row, range.Columns + 1, range.End.Row, columnaFinalDocumentoExcel].Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                    }

                    inicioParteCabeceraParte2++;

                    using (var range = ws.Cells[inicioParteCabeceraParte2, 1, inicioParteCabeceraParte2, 2])
                    {
                        range.Style.Border.Top.Style = ExcelBorderStyle.Hair;
                        range.Style.Border.Left.Style = ExcelBorderStyle.Hair;
                        range.Style.Border.Right.Style = ExcelBorderStyle.Hair;
                        range.Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                        range.Merge = true;

                        range.Value = "Nombre del Proyecto:";
                        range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

                        range.Style.Fill.BackgroundColor.SetColor(colorGrisClaro3EstiloPPM);
                        range.Style.Font.Color.SetColor(colorGrisOscuroEstiloPPM);
                        range.Style.Font.Bold = true;

                        ws.Cells[range.Start.Row, range.Columns + 1, range.End.Row, columnaFinalDocumentoExcel].Merge = true;
                        ws.Cells[range.Start.Row, range.Columns + 1].Value = acta.Cabecera.NombreProyecto;
                        ws.Cells[range.Start.Row, range.Columns + 1].Style.Indent = 1;

                        ws.Cells[range.Start.Row, range.Columns + 1, range.End.Row, columnaFinalDocumentoExcel].Style.Border.Top.Style = ExcelBorderStyle.Hair;
                        ws.Cells[range.Start.Row, range.Columns + 1, range.End.Row, columnaFinalDocumentoExcel].Style.Border.Left.Style = ExcelBorderStyle.Hair;
                        ws.Cells[range.Start.Row, range.Columns + 1, range.End.Row, columnaFinalDocumentoExcel].Style.Border.Right.Style = ExcelBorderStyle.Hair;
                        ws.Cells[range.Start.Row, range.Columns + 1, range.End.Row, columnaFinalDocumentoExcel].Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                        ws.Cells[range.Start.Row, range.Columns + 1, range.End.Row, columnaFinalDocumentoExcel].Style.WrapText = true;
                        var caracteres = acta.Cabecera.NombreProyecto.Count();

                        if (caracteres <= 104)
                        {
                            ws.Row(inicioParteCabeceraParte2).Height = 20.25;
                        }
                        else if (caracteres <= 208)
                        {
                            ws.Row(inicioParteCabeceraParte2).Height = 35.50;
                        }
                        else if (caracteres <= 312)
                        {
                            ws.Row(inicioParteCabeceraParte2).Height = 50.75;
                        }
                        else
                        {
                            ws.Row(inicioParteCabeceraParte2).Height = 66;
                        }
                    }

                    inicioParteCabeceraParte2++;

                    using (var range = ws.Cells[inicioParteCabeceraParte2, 1, inicioParteCabeceraParte2, 2])
                    {
                        range.Style.Border.Top.Style = ExcelBorderStyle.Hair;
                        range.Style.Border.Left.Style = ExcelBorderStyle.Hair;
                        range.Style.Border.Right.Style = ExcelBorderStyle.Hair;
                        range.Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                        range.Merge = true;

                        range.Value = "Descripción del Proyecto:";
                        range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

                        range.Style.Fill.BackgroundColor.SetColor(colorGrisClaro3EstiloPPM);
                        range.Style.Font.Color.SetColor(colorGrisOscuroEstiloPPM);
                        range.Style.Font.Bold = true;

                        ws.Cells[range.Start.Row, range.Columns + 1, range.End.Row, columnaFinalDocumentoExcel].Merge = true;
                        ws.Cells[range.Start.Row, range.Columns + 1].Value = acta.Cabecera.DescripcionProyecto;
                        ws.Cells[range.Start.Row, range.Columns + 1].Style.Indent = 1;

                        ws.Cells[range.Start.Row, range.Columns + 1, range.End.Row, columnaFinalDocumentoExcel].Style.Border.Top.Style = ExcelBorderStyle.Hair;
                        ws.Cells[range.Start.Row, range.Columns + 1, range.End.Row, columnaFinalDocumentoExcel].Style.Border.Left.Style = ExcelBorderStyle.Hair;
                        ws.Cells[range.Start.Row, range.Columns + 1, range.End.Row, columnaFinalDocumentoExcel].Style.Border.Right.Style = ExcelBorderStyle.Hair;
                        ws.Cells[range.Start.Row, range.Columns + 1, range.End.Row, columnaFinalDocumentoExcel].Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                        ws.Cells[range.Start.Row, range.Columns + 1, range.End.Row, columnaFinalDocumentoExcel].Style.WrapText = true;
                        var caracteres = acta.Cabecera.DescripcionProyecto.Count();
                        if (caracteres <= 104)
                        {
                            ws.Row(inicioParteCabeceraParte2).Height = 20.25;
                        }
                        else if (caracteres <= 208)
                        {
                            ws.Row(inicioParteCabeceraParte2).Height = 35.50;
                        }
                        else if (caracteres <= 312)
                        {
                            ws.Row(inicioParteCabeceraParte2).Height = 50.75;
                        }
                        else if (caracteres <= 416)
                        {
                            ws.Row(inicioParteCabeceraParte2).Height = 66;
                        }
                        else if (caracteres <= 500)
                        {
                            ws.Row(inicioParteCabeceraParte2).Height = 81.25;
                        }
                        else
                        {
                            ws.Row(inicioParteCabeceraParte2).Height = 96.50;
                        }


                    }

                    inicioParteCabeceraParte2++;

                    // Acta de cierre de proyecto o inicio de proyecto
                    if (acta.Cabecera.CodigoTipoActa == "AIP" || acta.Cabecera.CodigoTipoActa == "ACP")
                    {
                        inicioParteCabeceraParte2 += 1;
                        using (var range = ws.Cells[inicioParteCabeceraParte2, 1, inicioParteCabeceraParte2, 2])
                        {
                            range.Style.Border.Top.Style = ExcelBorderStyle.Hair;
                            range.Style.Border.Left.Style = ExcelBorderStyle.Hair;
                            range.Style.Border.Right.Style = ExcelBorderStyle.Hair;
                            range.Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                            range.Merge = true;

                            range.Value = "Referencia del Cliente:";
                            range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

                            range.Style.Fill.BackgroundColor.SetColor(colorGrisOscuroEstiloPPM);
                            range.Style.Font.Color.SetColor(colorBlancoEstiloPPM);
                            range.Style.Font.Bold = true;

                            ws.Cells[range.Start.Row, range.Columns + 1, range.End.Row, columnaFinalDocumentoExcel].Merge = true;
                            ws.Cells[range.Start.Row, range.Columns + 1].Value = acta.Cabecera.ReferenciaCliente;
                            ws.Cells[range.Start.Row, range.Columns + 1].Style.Indent = 1;

                            ws.Cells[range.Start.Row, range.Columns + 1, range.End.Row, columnaFinalDocumentoExcel].Style.Border.Top.Style = ExcelBorderStyle.Hair;
                            ws.Cells[range.Start.Row, range.Columns + 1, range.End.Row, columnaFinalDocumentoExcel].Style.Border.Left.Style = ExcelBorderStyle.Hair;
                            ws.Cells[range.Start.Row, range.Columns + 1, range.End.Row, columnaFinalDocumentoExcel].Style.Border.Right.Style = ExcelBorderStyle.Hair;
                            ws.Cells[range.Start.Row, range.Columns + 1, range.End.Row, columnaFinalDocumentoExcel].Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                            ws.Cells[range.Start.Row, range.Columns + 1, range.End.Row, columnaFinalDocumentoExcel].Style.WrapText = true;

                            var caracteres = (acta.Cabecera.ReferenciaCliente ?? "").Count();

                            if (caracteres <= 104)
                            {
                                ws.Row(inicioParteCabeceraParte2).Height = 20.25;
                            }
                            else if (caracteres <= 208)
                            {
                                ws.Row(inicioParteCabeceraParte2).Height = 35.50;
                            }

                            else
                            {
                                ws.Row(inicioParteCabeceraParte2).Height = 50.75;
                            }
                        }


                        inicioParteCabeceraParte2 += 2;
                        using (var range = ws.Cells[inicioParteCabeceraParte2, 1, inicioParteCabeceraParte2, 2])
                        {

                            range.Style.Border.Top.Style = ExcelBorderStyle.Hair;
                            range.Style.Border.Left.Style = ExcelBorderStyle.Hair;
                            range.Style.Border.Right.Style = ExcelBorderStyle.Hair;
                            range.Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                            range.Merge = true;

                            range.Value = "Objetivo o Alcance:";
                            range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

                            range.Style.Fill.BackgroundColor.SetColor(colorGrisOscuroEstiloPPM);
                            range.Style.Font.Color.SetColor(colorBlancoEstiloPPM);
                            range.Style.Font.Bold = true;

                            ws.Cells[range.Start.Row, range.Columns + 1, range.End.Row, columnaFinalDocumentoExcel].Merge = true;
                            ws.Cells[range.Start.Row, range.Columns + 1].Value = acta.Cabecera.AlcanceObjetivo;
                            ws.Cells[range.Start.Row, range.Columns + 1].Style.Indent = 1;

                            ws.Cells[range.Start.Row, range.Columns + 1, range.End.Row, columnaFinalDocumentoExcel].Style.Border.Top.Style = ExcelBorderStyle.Hair;
                            ws.Cells[range.Start.Row, range.Columns + 1, range.End.Row, columnaFinalDocumentoExcel].Style.Border.Left.Style = ExcelBorderStyle.Hair;
                            ws.Cells[range.Start.Row, range.Columns + 1, range.End.Row, columnaFinalDocumentoExcel].Style.Border.Right.Style = ExcelBorderStyle.Hair;
                            ws.Cells[range.Start.Row, range.Columns + 1, range.End.Row, columnaFinalDocumentoExcel].Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                            //  ws.Row(inicioParteCabeceraParte2).Height = 33.75;
                            ws.Cells[range.Start.Row, range.Columns + 1, range.End.Row, columnaFinalDocumentoExcel].Style.WrapText = true;

                            var caracteres = acta.Cabecera.AlcanceObjetivo.Count();

                            if (caracteres <= 104)
                            {
                                ws.Row(inicioParteCabeceraParte2).Height = 20.25;
                            }
                            else if (caracteres <= 208)
                            {
                                ws.Row(inicioParteCabeceraParte2).Height = 35.50;
                            }
                            else if (caracteres <= 312)
                            {
                                ws.Row(inicioParteCabeceraParte2).Height = 50.75;
                            }
                            else if (caracteres <= 416)
                            {
                                ws.Row(inicioParteCabeceraParte2).Height = 66;
                            }
                            else if (caracteres <= 500)
                            {
                                ws.Row(inicioParteCabeceraParte2).Height = 81.25;
                            }
                            else
                            {
                                ws.Row(inicioParteCabeceraParte2).Height = 96.50;
                            }
                        }
                    }
                    else
                    {
                        ws.Row(inicioParteCabeceraParte2).Height = 8.25;
                        inicioParteCabeceraParte2 += 1;
                        using (var range = ws.Cells[inicioParteCabeceraParte2, 1, inicioParteCabeceraParte2, 2])
                        {
                            range.Style.Border.Top.Style = ExcelBorderStyle.Hair;
                            range.Style.Border.Left.Style = ExcelBorderStyle.Hair;
                            range.Style.Border.Right.Style = ExcelBorderStyle.Hair;
                            range.Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                            range.Merge = true;

                            range.Value = "Facilitador o Moderador:";
                            range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

                            range.Style.Fill.BackgroundColor.SetColor(colorGrisOscuroEstiloPPM);
                            range.Style.Font.Color.SetColor(colorBlancoEstiloPPM);
                            range.Style.Font.Bold = true;

                            ws.Cells[range.Start.Row, range.Columns + 1, range.End.Row, columnaFinalDocumentoExcel].Merge = true;
                            ws.Cells[range.Start.Row, range.Columns + 1].Value = acta.Cabecera.FacilitadorModerador;
                            ws.Cells[range.Start.Row, range.Columns + 1].Style.Indent = 1;

                            ws.Cells[range.Start.Row, range.Columns + 1, range.End.Row, columnaFinalDocumentoExcel].Style.Border.Top.Style = ExcelBorderStyle.Hair;
                            ws.Cells[range.Start.Row, range.Columns + 1, range.End.Row, columnaFinalDocumentoExcel].Style.Border.Left.Style = ExcelBorderStyle.Hair;
                            ws.Cells[range.Start.Row, range.Columns + 1, range.End.Row, columnaFinalDocumentoExcel].Style.Border.Right.Style = ExcelBorderStyle.Hair;
                            ws.Cells[range.Start.Row, range.Columns + 1, range.End.Row, columnaFinalDocumentoExcel].Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                            ws.Cells[range.Start.Row, range.Columns + 1, range.End.Row, columnaFinalDocumentoExcel].Style.WrapText = true;
                        }
                        ws.Row(inicioParteCabeceraParte2).Height = 20.25;

                        inicioParteCabeceraParte2 += 1;

                        using (var range = ws.Cells[inicioParteCabeceraParte2, 1, inicioParteCabeceraParte2, 2])
                        {

                            range.Style.Border.Top.Style = ExcelBorderStyle.Hair;
                            range.Style.Border.Left.Style = ExcelBorderStyle.Hair;
                            range.Style.Border.Right.Style = ExcelBorderStyle.Hair;
                            range.Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                            range.Merge = true;

                            range.Value = "Objetivo o Alcance:";
                            range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

                            range.Style.Fill.BackgroundColor.SetColor(colorGrisOscuroEstiloPPM);
                            range.Style.Font.Color.SetColor(colorBlancoEstiloPPM);
                            range.Style.Font.Bold = true;

                            ws.Cells[range.Start.Row, range.Columns + 1, range.End.Row, columnaFinalDocumentoExcel].Merge = true;
                            ws.Cells[range.Start.Row, range.Columns + 1].Value = acta.Cabecera.AlcanceObjetivo;
                            ws.Cells[range.Start.Row, range.Columns + 1].Style.Indent = 1;

                            ws.Cells[range.Start.Row, range.Columns + 1, range.End.Row, columnaFinalDocumentoExcel].Style.Border.Top.Style = ExcelBorderStyle.Hair;
                            ws.Cells[range.Start.Row, range.Columns + 1, range.End.Row, columnaFinalDocumentoExcel].Style.Border.Left.Style = ExcelBorderStyle.Hair;
                            ws.Cells[range.Start.Row, range.Columns + 1, range.End.Row, columnaFinalDocumentoExcel].Style.Border.Right.Style = ExcelBorderStyle.Hair;
                            ws.Cells[range.Start.Row, range.Columns + 1, range.End.Row, columnaFinalDocumentoExcel].Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                            ws.Row(inicioParteCabeceraParte2).Height = 33.75;
                            ws.Cells[range.Start.Row, range.Columns + 1, range.End.Row, columnaFinalDocumentoExcel].Style.WrapText = true;


                            var caracteres = acta.Cabecera.AlcanceObjetivo.Count();

                            if (caracteres <= 104)
                            {
                                ws.Row(inicioParteCabeceraParte2).Height = 20.25;
                            }
                            else if (caracteres <= 208)
                            {
                                ws.Row(inicioParteCabeceraParte2).Height = 35.50;
                            }
                            else if (caracteres <= 312)
                            {
                                ws.Row(inicioParteCabeceraParte2).Height = 50.75;
                            }
                            else if (caracteres <= 416)
                            {
                                ws.Row(inicioParteCabeceraParte2).Height = 66;
                            }
                            else if (caracteres <= 500)
                            {
                                ws.Row(inicioParteCabeceraParte2).Height = 81.25;
                            }
                            else
                            {
                                ws.Row(inicioParteCabeceraParte2).Height = 96.50;
                            }
                        }
                    }

                    inicioParteCabeceraParte2 += 2;
                    using (ExcelRange row = ws.Cells["A5:A15"])
                    {
                        row.Style.Font.Bold = true;
                    }

                    //int finCabecera;

                    if (acta.Cabecera.CodigoTipoActa == "AIP" || acta.Cabecera.CodigoTipoActa == "ACP")
                    {

                        finCabecera = 18;
                        ws.Cells[finCabecera, 1].Value = "R  E  S  P  O  N  S  A  B  L  E  S   D  E  L   C  L  I  E  N  T  E";
                        ws.Cells[finCabecera, 1].Style.Fill.BackgroundColor.SetColor(colorGrisOscuroEstiloPPM);
                        ws.Cells[finCabecera, 1].Style.Font.Color.SetColor(colorBlancoEstiloPPM);
                        ws.Cells[finCabecera, 1, finCabecera, 9].Merge = true;
                        ws.Cells[finCabecera, 1, finCabecera, 9].Style.HorizontalAlignment = ExcelHorizontalAlignment.CenterContinuous;
                        ws.Cells[finCabecera, 1, finCabecera, 9].Style.Border.Top.Style = ExcelBorderStyle.Hair;
                        ws.Cells[finCabecera, 1, finCabecera, 9].Style.Border.Left.Style = ExcelBorderStyle.Hair;
                        ws.Cells[finCabecera, 1, finCabecera, 9].Style.Border.Right.Style = ExcelBorderStyle.Hair;
                        ws.Cells[finCabecera, 1, finCabecera, 9].Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                        ws.Row(finCabecera).Height = 20.25;

                    }
                    else
                    {
                        finCabecera = 16;

                    }

                    finCabecera++;
                    ws.Row(finCabecera).Height = 20.25;

                }

                #endregion

                Int32 col = 1;
                int contador = 1;

                #region Detalle Acta Cliente

                if (acta.Cuerpo.DetalleCliente.Any())
                {
                    var columnas = new List<string> { "N°", "CÓDIGO DE COTIZACIÓN", "CLIENTE", "INTERMEDIARIO", "DETALLE", "VALOR", "EJECUTIVO", "N. DOCUMENTO", "OBSERVACIONES" };

                    var i = 1;
                    foreach (var item in columnas)
                    {

                        ws.Cells[finCabecera, i].Value = item;
                        ws.Cells[finCabecera, i].Style.Font.Bold = true;

                        ws.Cells[finCabecera, i].Style.Border.Top.Style = ExcelBorderStyle.Hair;
                        ws.Cells[finCabecera, i].Style.Border.Left.Style = ExcelBorderStyle.Hair;
                        ws.Cells[finCabecera, i].Style.Border.Right.Style = ExcelBorderStyle.Hair;
                        ws.Cells[finCabecera, i].Style.Border.Bottom.Style = ExcelBorderStyle.Hair;

                        //worksheet.Cells[1, i].Style.Fill.BackgroundColor. .SetColor(Color.FromArgb(23, 55, 93));
                        i++;
                    }
                    finCabecera++;

                    int numeracion = 1;
                    foreach (var item in acta.Cuerpo.DetalleCliente)
                    {
                        var detalle = CotizacionEntity.ConsultarPrefacturaSAFI(item.id_facturacion_safi);
                        for (int j = 1; j <= 9; j++)
                        {
                            ws.Cells[finCabecera, j].Style.Border.Top.Style = ExcelBorderStyle.Hair;
                            ws.Cells[finCabecera, j].Style.Border.Left.Style = ExcelBorderStyle.Hair;
                            ws.Cells[finCabecera, j].Style.Border.Right.Style = ExcelBorderStyle.Hair;
                            ws.Cells[finCabecera, j].Style.Border.Bottom.Style = ExcelBorderStyle.Hair;

                            switch (j)
                            {
                                case 1:
                                    ws.Column(j).Width = 5;
                                    ws.Cells[finCabecera, j].Value = numeracion;
                                    break;
                                case 2:
                                    ws.Cells[finCabecera, j].Value = detalle.codigo_cotizacion;
                                    break;
                                case 3:
                                    ws.Cells[finCabecera, j].Value = detalle.nombre_comercial_cliente;
                                    ws.Cells[finCabecera, j].Style.WrapText = true;
                                    break;
                                case 4:
                                    ws.Cells[finCabecera, j].Value = detalle.TipoIntermediario;
                                    ws.Cells[finCabecera, j].Style.WrapText = true;
                                    break;
                                case 5:
                                    ws.Cells[finCabecera, j].Value = detalle.detalle_cotizacion;
                                    ws.Cells[finCabecera, j].Style.WrapText = true;
                                    break;
                                case 6:
                                    ws.Cells[finCabecera, j].Value = detalle.total_pago;
                                    ws.Cells[finCabecera, j].Style.Numberformat.Format = "$#,##0.00";
                                    break;
                                case 7:
                                    ws.Cells[finCabecera, j].Value = detalle.Ejecutivo;
                                    ws.Cells[finCabecera, j].Style.WrapText = true;
                                    break;
                                case 8:
                                    ws.Cells[finCabecera, j].Value = detalle.numero_prefactura;
                                    break;
                                case 9:
                                    ws.Cells[finCabecera, j].Value = detalle.ObservacionCotizacion;
                                    ws.Cells[finCabecera, j].Style.WrapText = true;
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
                }

                #endregion

                #region Detalle Acta Contabilidad
                if (acta.Cuerpo.DetalleContabilidad.Any())
                {
                    var columnas = new List<string> { "N°", "CÓDIGO DE COTIZACIÓN", "CLIENTE", "INTERMEDIARIO", "DETALLE", "VALOR", "EJECUTIVO", "N. DOCUMENTO", "OBSERVACIONES" };

                    var i = 1;
                    foreach (var item in columnas)
                    {

                        ws.Cells[finCabecera, i].Value = item;
                        ws.Cells[finCabecera, i].Style.Font.Bold = true;

                        ws.Cells[finCabecera, i].Style.Border.Top.Style = ExcelBorderStyle.Hair;
                        ws.Cells[finCabecera, i].Style.Border.Left.Style = ExcelBorderStyle.Hair;
                        ws.Cells[finCabecera, i].Style.Border.Right.Style = ExcelBorderStyle.Hair;
                        ws.Cells[finCabecera, i].Style.Border.Bottom.Style = ExcelBorderStyle.Hair;

                        //worksheet.Cells[1, i].Style.Fill.BackgroundColor. .SetColor(Color.FromArgb(23, 55, 93));
                        i++;
                    }
                    finCabecera++;

                    int numeracion = 1;
                    foreach (var item in acta.Cuerpo.DetalleContabilidad)
                    {
                        var detalle = CotizacionEntity.ConsultarPrefacturaSAFI(item.id_facturacion_safi);
                        for (int j = 1; j <= 9; j++)
                        {
                            ws.Cells[finCabecera, j].Style.Border.Top.Style = ExcelBorderStyle.Hair;
                            ws.Cells[finCabecera, j].Style.Border.Left.Style = ExcelBorderStyle.Hair;
                            ws.Cells[finCabecera, j].Style.Border.Right.Style = ExcelBorderStyle.Hair;
                            ws.Cells[finCabecera, j].Style.Border.Bottom.Style = ExcelBorderStyle.Hair;

                            switch (j)
                            {
                                case 1:
                                    ws.Column(j).Width = 5;
                                    ws.Cells[finCabecera, j].Value = numeracion;
                                    break;
                                case 2:
                                    ws.Cells[finCabecera, j].Value = detalle.codigo_cotizacion;
                                    break;
                                case 3:
                                    ws.Cells[finCabecera, j].Value = detalle.nombre_comercial_cliente;
                                    ws.Cells[finCabecera, j].Style.WrapText = true;
                                    break;
                                case 4:
                                    ws.Cells[finCabecera, j].Value = detalle.TipoIntermediario;
                                    ws.Cells[finCabecera, j].Style.WrapText = true;
                                    break;
                                case 5:
                                    ws.Cells[finCabecera, j].Value = detalle.detalle_cotizacion;
                                    ws.Cells[finCabecera, j].Style.WrapText = true;
                                    break;
                                case 6:
                                    ws.Cells[finCabecera, j].Value = detalle.total_pago;
                                    ws.Cells[finCabecera, j].Style.Numberformat.Format = "$#,##0.00";
                                    break;
                                case 7:
                                    ws.Cells[finCabecera, j].Value = detalle.Ejecutivo;
                                    ws.Cells[finCabecera, j].Style.WrapText = true;
                                    break;
                                case 8:
                                    ws.Cells[finCabecera, j].Value = detalle.numero_prefactura;
                                    break;
                                case 9:
                                    ws.Cells[finCabecera, j].Value = detalle.ObservacionCotizacion;
                                    ws.Cells[finCabecera, j].Style.WrapText = true;
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
                }
                #endregion

                #region DetalleResponsables
                if (acta.Cuerpo.Responsables.Any())
                {
                    int totalColumnasDetalleResponsables = typeof(DetalleActaResponsables).GetProperties().Length;

                    int contadorRegistrosResponsables = acta.Cuerpo.Responsables.Count;

                    if (contadorRegistrosResponsables > 30)
                    {
                        for (int i = 1; i <= 2; i++)
                        {
                            int totalMergeColumnasParticipantes = i == 1 ? 3 : 6;

                            using (var range = ws.Cells[finCabecera, col, finCabecera, totalMergeColumnasParticipantes])
                            {
                                range.Value = "NOMBRE";
                                range.Style.Fill.PatternType = ExcelFillStyle.Solid;
                                range.Style.Fill.BackgroundColor.SetColor(colorGrisClaro3EstiloPPM);
                                range.Style.Font.Color.SetColor(colorGrisOscuroEstiloPPM);
                                range.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                                range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                                // Assign borders
                                range.Style.Border.Top.Style = ExcelBorderStyle.Hair;
                                range.Style.Border.Left.Style = ExcelBorderStyle.Hair;
                                range.Style.Border.Right.Style = ExcelBorderStyle.Hair;
                                range.Style.Border.Bottom.Style = ExcelBorderStyle.Hair;

                                range.Merge = true;

                                if (col == 6)

                                    ws.Cells[finCabecera, 6, finCabecera, 7].Merge = true;
                                ws.Cells[finCabecera, 6, finCabecera, 7].Style.Border.Top.Style = ExcelBorderStyle.Hair;
                                ws.Cells[finCabecera, 6, finCabecera, 7].Style.Border.Left.Style = ExcelBorderStyle.Hair;
                                ws.Cells[finCabecera, 6, finCabecera, 7].Style.Border.Right.Style = ExcelBorderStyle.Hair;
                                ws.Cells[finCabecera, 6, finCabecera, 7].Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                            }

                            if (col == 6)
                            {
                                col += 2;
                            }
                            else
                            {
                                col += 3;
                            }

                            using (var range = ws.Cells[finCabecera, col, finCabecera, col])
                            {
                                range.Value = "EMPRESA";
                                range.Style.Fill.PatternType = ExcelFillStyle.Solid;
                                range.Style.Fill.BackgroundColor.SetColor(colorGrisClaro3EstiloPPM);
                                range.Style.Font.Color.SetColor(colorGrisOscuroEstiloPPM);
                                range.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                                range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                                // Assign borders
                                range.Style.Border.Top.Style = ExcelBorderStyle.Hair;
                                range.Style.Border.Left.Style = ExcelBorderStyle.Hair;
                                range.Style.Border.Right.Style = ExcelBorderStyle.Hair;
                                range.Style.Border.Bottom.Style = ExcelBorderStyle.Hair;

                                //range.Merge = true;
                            }
                            col += 1;
                            using (var range = ws.Cells[finCabecera, col, finCabecera, col])
                            {
                                range.Value = "ROL";
                                range.Style.Fill.PatternType = ExcelFillStyle.Solid;
                                range.Style.Fill.BackgroundColor.SetColor(colorGrisClaro3EstiloPPM);
                                range.Style.Font.Color.SetColor(colorGrisOscuroEstiloPPM);
                                range.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                                range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                                // Assign borders
                                range.Style.Border.Top.Style = ExcelBorderStyle.Hair;
                                range.Style.Border.Left.Style = ExcelBorderStyle.Hair;
                                range.Style.Border.Right.Style = ExcelBorderStyle.Hair;
                                range.Style.Border.Bottom.Style = ExcelBorderStyle.Hair;

                            }
                            col += 1;
                        }

                        finCabecera++;

                        contador = 1;

                        int inicioFilaResponsables = finCabecera;
                        int columnasDivision1 = contadorRegistrosResponsables / 2;
                        int columnasDivision2 = contadorRegistrosResponsables - (contadorRegistrosResponsables / 2);

                        foreach (var column in acta.Cuerpo.Responsables)
                        {
                            col = contador <= columnasDivision1 ? 1 : 6;

                            if (contador == columnasDivision2 + 1)
                            {
                                finCabecera = inicioFilaResponsables;
                            }

                            for (int i = 1; i <= totalColumnasDetalleResponsables; i++)
                            {
                                ws.Cells[finCabecera, col].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                                ws.Cells[finCabecera, col].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                                switch (i)
                                {
                                    case 1:
                                        if (col == 1)
                                        {
                                            ws.Cells[finCabecera, col, finCabecera, col + 2].Merge = true;
                                            ws.Cells[finCabecera, col, finCabecera, col + 2].Style.Border.Top.Style = ExcelBorderStyle.Hair;
                                            ws.Cells[finCabecera, col, finCabecera, col + 2].Style.Border.Left.Style = ExcelBorderStyle.Hair;
                                            ws.Cells[finCabecera, col, finCabecera, col + 2].Style.Border.Right.Style = ExcelBorderStyle.Hair;
                                            ws.Cells[finCabecera, col, finCabecera, col + 2].Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                                            ws.Cells[finCabecera, 6, finCabecera, 7].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                                            ws.Cells[finCabecera, col].Value = column.Nombres;
                                        }
                                        else
                                        {
                                            ws.Cells[finCabecera, 6, finCabecera, 7].Value = column.Nombres;
                                            ws.Cells[finCabecera, 6, finCabecera, 7].Merge = true;
                                            ws.Cells[finCabecera, 6, finCabecera, 7].Style.Border.Top.Style = ExcelBorderStyle.Hair;
                                            ws.Cells[finCabecera, 6, finCabecera, 7].Style.Border.Left.Style = ExcelBorderStyle.Hair;
                                            ws.Cells[finCabecera, 6, finCabecera, 7].Style.Border.Right.Style = ExcelBorderStyle.Hair;
                                            ws.Cells[finCabecera, 6, finCabecera, 7].Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                                            ws.Cells[finCabecera, 6, finCabecera, 7].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                                        }
                                        col += 3;
                                        break;

                                    case 2:
                                        if (col == 4)
                                        {
                                            ws.Cells[finCabecera, col].Merge = true;
                                            ws.Cells[finCabecera, col].Style.Border.Top.Style = ExcelBorderStyle.Hair;
                                            ws.Cells[finCabecera, col].Style.Border.Left.Style = ExcelBorderStyle.Hair;
                                            ws.Cells[finCabecera, col].Style.Border.Right.Style = ExcelBorderStyle.Hair;
                                            ws.Cells[finCabecera, col].Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                                            ws.Cells[finCabecera, col].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                                            ws.Cells[finCabecera, col].Value = column.Empresa;
                                            col += 1;
                                        }
                                        else
                                        {

                                            ws.Cells[finCabecera, 8, finCabecera, 8].Merge = true;
                                            ws.Cells[finCabecera, 8, finCabecera, 8].Style.Border.Top.Style = ExcelBorderStyle.Hair;
                                            ws.Cells[finCabecera, 8, finCabecera, 8].Style.Border.Left.Style = ExcelBorderStyle.Hair;
                                            ws.Cells[finCabecera, 8, finCabecera, 8].Style.Border.Right.Style = ExcelBorderStyle.Hair;
                                            ws.Cells[finCabecera, 8, finCabecera, 8].Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                                            ws.Cells[finCabecera, 8, finCabecera, 8].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                                            ws.Cells[finCabecera, 8, finCabecera, 8].Value = column.Empresa;

                                        }
                                        break;


                                    case 3:
                                        if (col == 5)
                                        {
                                            ws.Cells[finCabecera, col].Style.Border.Top.Style = ExcelBorderStyle.Hair;
                                            ws.Cells[finCabecera, col].Style.Border.Left.Style = ExcelBorderStyle.Hair;
                                            ws.Cells[finCabecera, col].Style.Border.Right.Style = ExcelBorderStyle.Hair;
                                            ws.Cells[finCabecera, col].Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                                            ws.Cells[finCabecera, col].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                                            ws.Cells[finCabecera, col].Value = column.Rol;
                                        }
                                        else
                                        {

                                            ws.Cells[finCabecera, 9, finCabecera, 9].Value = column.Rol;
                                            ws.Cells[finCabecera, 9, finCabecera, 9].Style.Border.Top.Style = ExcelBorderStyle.Hair;
                                            ws.Cells[finCabecera, 9, finCabecera, 9].Style.Border.Left.Style = ExcelBorderStyle.Hair;
                                            ws.Cells[finCabecera, 9, finCabecera, 9].Style.Border.Right.Style = ExcelBorderStyle.Hair;
                                            ws.Cells[finCabecera, 9, finCabecera, 9].Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                                            ws.Cells[finCabecera, 9, finCabecera, 9].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;

                                        }
                                        col += 1;

                                        break;


                                    default:
                                        ws.Cells[finCabecera, col++].Value = "";
                                        break;
                                }
                            }

                            ws.Row(finCabecera).Height = 20.25;
                            finCabecera++;
                            contador++;
                        }
                    }
                    else
                    {
                        for (int j = 1; j <= totalColumnasDetalleResponsables; j++)
                        {

                            ws.Cells[finCabecera, col].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            ws.Cells[finCabecera, col].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(31, 78, 120));
                            ws.Cells[finCabecera, col].Style.Font.Color.SetColor(Color.White);
                            ws.Cells[finCabecera, col].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                            ws.Cells[finCabecera, col].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            // Assign borders
                            ws.Cells[finCabecera, col].Style.Border.Top.Style = ExcelBorderStyle.Hair;
                            ws.Cells[finCabecera, col].Style.Border.Left.Style = ExcelBorderStyle.Hair;
                            ws.Cells[finCabecera, col].Style.Border.Right.Style = ExcelBorderStyle.Hair;
                            ws.Cells[finCabecera, col].Style.Border.Bottom.Style = ExcelBorderStyle.Hair;

                            switch (j)
                            {
                                case 1:
                                    using (var range = ws.Cells[finCabecera, col++, finCabecera, col + 2])
                                    {
                                        range.Style.Border.Top.Style = ExcelBorderStyle.Hair;
                                        range.Style.Border.Left.Style = ExcelBorderStyle.Hair;
                                        range.Style.Border.Right.Style = ExcelBorderStyle.Hair;
                                        range.Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                                        range.Merge = true;

                                        range.Value = "NOMBRE";
                                        range.Style.Fill.BackgroundColor.SetColor(colorGrisClaro3EstiloPPM);
                                        range.Style.Font.Color.SetColor(colorGrisOscuroEstiloPPM);
                                        range.Style.Font.Bold = true;

                                    }
                                    break;

                                case 2:
                                    using (var range = ws.Cells[finCabecera, col + 3, finCabecera, col + 5])
                                    {
                                        range.Style.Border.Top.Style = ExcelBorderStyle.Hair;
                                        range.Style.Border.Left.Style = ExcelBorderStyle.Hair;
                                        range.Style.Border.Right.Style = ExcelBorderStyle.Hair;
                                        range.Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                                        range.Merge = true;

                                        range.Value = "EMPRESA";
                                        range.Style.Fill.BackgroundColor.SetColor(colorGrisClaro3EstiloPPM);
                                        range.Style.Font.Color.SetColor(colorGrisOscuroEstiloPPM);
                                        range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                        range.Style.Font.Bold = true;

                                    }
                                    break;

                                case 3:
                                    using (var range = ws.Cells[finCabecera, col + 6, finCabecera, col + 7])
                                    {
                                        range.Style.Border.Top.Style = ExcelBorderStyle.Hair;
                                        range.Style.Border.Left.Style = ExcelBorderStyle.Hair;
                                        range.Style.Border.Right.Style = ExcelBorderStyle.Hair;
                                        range.Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                                        range.Merge = true;

                                        range.Value = "ROL";
                                        range.Style.Fill.BackgroundColor.SetColor(colorGrisClaro3EstiloPPM);
                                        range.Style.Font.Color.SetColor(colorGrisOscuroEstiloPPM);
                                        range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                        range.Style.Font.Bold = true;
                                    }
                                    break;
                                default:
                                    ws.Cells[finCabecera, col++].Value = "";
                                    break;
                            }
                        }

                        finCabecera++;
                        contador = 1;
                        foreach (var column in acta.Cuerpo.Responsables)
                        {
                            col = 1;
                            for (int i = 1; i <= totalColumnasDetalleResponsables; i++)
                            {
                                ws.Cells[finCabecera, col].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                                ws.Cells[finCabecera, col].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                // Assign borders
                                ws.Cells[finCabecera, col].Style.Border.Top.Style = ExcelBorderStyle.Hair;
                                ws.Cells[finCabecera, col].Style.Border.Left.Style = ExcelBorderStyle.Hair;
                                ws.Cells[finCabecera, col].Style.Border.Right.Style = ExcelBorderStyle.Hair;
                                ws.Cells[finCabecera, col].Style.Border.Bottom.Style = ExcelBorderStyle.Hair;

                                switch (i)
                                {
                                    case 1:

                                        using (var range = ws.Cells[finCabecera, col++, finCabecera, col + 2])
                                        {
                                            range.Value = column.Nombres;

                                            range.Style.Border.Top.Style = ExcelBorderStyle.Hair;
                                            range.Style.Border.Left.Style = ExcelBorderStyle.Hair;
                                            range.Style.Border.Right.Style = ExcelBorderStyle.Hair;
                                            range.Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                                            range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                                            range.Merge = true;

                                        }

                                        break;

                                    case 2:

                                        using (var range = ws.Cells[finCabecera, col + 3, finCabecera, col + 5])
                                        {
                                            range.Value = column.Empresa;

                                            range.Style.Border.Top.Style = ExcelBorderStyle.Hair;
                                            range.Style.Border.Left.Style = ExcelBorderStyle.Hair;
                                            range.Style.Border.Right.Style = ExcelBorderStyle.Hair;
                                            range.Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                                            range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                                            range.Merge = true;


                                        }
                                        break;

                                    case 3:

                                        using (var range = ws.Cells[finCabecera, col + 6, finCabecera, col + 7])
                                        {
                                            range.Value = column.Rol;

                                            range.Style.Border.Top.Style = ExcelBorderStyle.Hair;
                                            range.Style.Border.Left.Style = ExcelBorderStyle.Hair;
                                            range.Style.Border.Right.Style = ExcelBorderStyle.Hair;
                                            range.Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                                            range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                                            range.Merge = true;

                                        }

                                        break;

                                    default:
                                        ws.Cells[finCabecera, col++].Value = "";
                                        break;
                                }
                            }
                            ws.Row(finCabecera).Height = 20.25;
                            finCabecera++;
                            contador++;
                        }
                    }

                    ws.Column(2).AutoFit();
                    ws.Column(3).AutoFit();
                    ws.Column(4).AutoFit();
                    ws.Column(6).AutoFit();
                    ws.Column(7).AutoFit();
                    ws.Column(8).AutoFit();

                    ws.Row(finCabecera).Height = 8.25;
                    finCabecera += 1;
                }
                #endregion

                #region DetalleEntregable
                if (acta.Cuerpo.Entregables.Any())
                {

                    col = 1;
                    int totalColumnasDetalleEntregables = typeof(DetalleActaEntregables).GetProperties().Length;

                    for (int i = 1; i <= totalColumnasDetalleEntregables; i++)
                    {
                        ws.Cells[finCabecera, col].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                        ws.Cells[finCabecera, col].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        //Assign borders
                        ws.Cells[finCabecera, col].Style.Border.Top.Style = ExcelBorderStyle.Hair;
                        ws.Cells[finCabecera, col].Style.Border.Left.Style = ExcelBorderStyle.Hair;
                        ws.Cells[finCabecera, col].Style.Border.Right.Style = ExcelBorderStyle.Hair;
                        ws.Cells[finCabecera, col].Style.Border.Bottom.Style = ExcelBorderStyle.Hair;

                        switch (i)
                        {
                            case 1:
                                ws.Cells[finCabecera, col++].Value = "N.";
                                ws.Cells[finCabecera, 1, finCabecera, 1].Style.Border.Top.Style = ExcelBorderStyle.Hair;
                                ws.Cells[finCabecera, 1, finCabecera, 1].Style.Border.Left.Style = ExcelBorderStyle.Hair;
                                ws.Cells[finCabecera, 1, finCabecera, 1].Style.Border.Right.Style = ExcelBorderStyle.Hair;
                                ws.Cells[finCabecera, 1, finCabecera, 1].Style.Border.Bottom.Style = ExcelBorderStyle.Hair;

                                ws.Cells[finCabecera, 1, finCabecera, 1].Style.Fill.BackgroundColor.SetColor(colorGrisClaro3EstiloPPM);
                                ws.Cells[finCabecera, 1, finCabecera, 1].Style.Font.Color.SetColor(colorGrisOscuroEstiloPPM);
                                ws.Cells[finCabecera, 1, finCabecera, 1].Style.Font.Bold = true;

                                break;
                            case 2:
                                ws.Cells[finCabecera, col++].Value = "ENTREGABLES";
                                ws.Cells[finCabecera, 2, finCabecera, 7].Merge = true;
                                ws.Cells[finCabecera, 2, finCabecera, 7].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                                ws.Cells[finCabecera, 2, finCabecera, 7].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                // Assign borders
                                ws.Cells[finCabecera, 2, finCabecera, 7].Style.Border.Top.Style = ExcelBorderStyle.Hair;
                                ws.Cells[finCabecera, 2, finCabecera, 7].Style.Border.Left.Style = ExcelBorderStyle.Hair;
                                ws.Cells[finCabecera, 2, finCabecera, 7].Style.Border.Right.Style = ExcelBorderStyle.Hair;
                                ws.Cells[finCabecera, 2, finCabecera, 7].Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                                ws.Cells[finCabecera, 2, finCabecera, 7].Style.Font.Bold = true;

                                ws.Cells[finCabecera, 2, finCabecera, 7].Style.Fill.BackgroundColor.SetColor(colorGrisClaro3EstiloPPM);
                                ws.Cells[finCabecera, 2, finCabecera, 7].Style.Font.Color.SetColor(colorGrisOscuroEstiloPPM);

                                break;
                            case 3:
                                ws.Cells[finCabecera, 8].Value = "TIPO";
                                ws.Cells[finCabecera, 8, finCabecera, 9].Merge = true;
                                ws.Cells[finCabecera, 8, finCabecera, 9].Style.Fill.PatternType = ExcelFillStyle.Solid;

                                ws.Cells[finCabecera, 8, finCabecera, 9].Style.Fill.BackgroundColor.SetColor(colorGrisClaro3EstiloPPM);
                                ws.Cells[finCabecera, 8, finCabecera, 9].Style.Font.Color.SetColor(colorGrisOscuroEstiloPPM);

                                ws.Cells[finCabecera, 8, finCabecera, 9].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                                ws.Cells[finCabecera, 8, finCabecera, 9].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                // Assign border
                                ws.Cells[finCabecera, 8, finCabecera, 9].Style.Border.Top.Style = ExcelBorderStyle.Hair;
                                ws.Cells[finCabecera, 8, finCabecera, 9].Style.Border.Left.Style = ExcelBorderStyle.Hair;
                                ws.Cells[finCabecera, 8, finCabecera, 9].Style.Border.Right.Style = ExcelBorderStyle.Hair;
                                ws.Cells[finCabecera, 8, finCabecera, 9].Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                                ws.Cells[finCabecera, 8, finCabecera, 9].Style.Font.Bold = true;


                                break;
                            default:
                                ws.Cells[finCabecera, col++].Value = "";
                                break;
                        }
                        ws.Row(finCabecera).Height = 20.25;
                    }

                    finCabecera++;

                    contador = 1;
                    foreach (var column in acta.Cuerpo.Entregables)
                    {
                        col = 1;
                        for (int i = 1; i <= totalColumnasDetalleEntregables; i++)
                        {
                            ws.Cells[finCabecera, col].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                            ws.Cells[finCabecera, col].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            // Assign borders
                            ws.Cells[finCabecera, col].Style.Border.Top.Style = ExcelBorderStyle.Hair;
                            ws.Cells[finCabecera, col].Style.Border.Left.Style = ExcelBorderStyle.Hair;
                            ws.Cells[finCabecera, col].Style.Border.Right.Style = ExcelBorderStyle.Hair;
                            ws.Cells[finCabecera, col].Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                            switch (i)
                            {
                                case 1:
                                    ws.Cells[finCabecera, col++].Value = contador;
                                    ws.Cells[finCabecera, 1, finCabecera, 1].Style.Border.Top.Style = ExcelBorderStyle.Hair;
                                    ws.Cells[finCabecera, 1, finCabecera, 1].Style.Border.Left.Style = ExcelBorderStyle.Hair;
                                    ws.Cells[finCabecera, 1, finCabecera, 1].Style.Border.Right.Style = ExcelBorderStyle.Hair;
                                    ws.Cells[finCabecera, 1, finCabecera, 1].Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                                    break;
                                case 2:
                                    ws.Cells[finCabecera, col++].Value = column.Entregable;
                                    ws.Cells[finCabecera, 2, finCabecera, 7].Merge = true;
                                    ws.Cells[finCabecera, 2, finCabecera, 7].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                                    ws.Cells[finCabecera, 2, finCabecera, 7].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                                    // Assign borders
                                    ws.Cells[finCabecera, 2, finCabecera, 7].Style.Border.Top.Style = ExcelBorderStyle.Hair;
                                    ws.Cells[finCabecera, 2, finCabecera, 7].Style.Border.Left.Style = ExcelBorderStyle.Hair;
                                    ws.Cells[finCabecera, 2, finCabecera, 7].Style.Border.Right.Style = ExcelBorderStyle.Hair;
                                    ws.Cells[finCabecera, 2, finCabecera, 7].Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                                    ws.Cells[finCabecera, 2, finCabecera, 7].Style.WrapText = true;

                                    var caracteres = column.Entregable.Count();
                                    if (caracteres <= 70)
                                    {
                                        ws.Row(finCabecera).Height = 20.25;
                                    }
                                    else if (caracteres <= 140)
                                    {
                                        ws.Row(finCabecera).Height = 30.50;
                                    }
                                    else if (caracteres <= 210)
                                    {
                                        ws.Row(finCabecera).Height = 40.75;
                                    }
                                    else if (caracteres <= 280)
                                    {
                                        ws.Row(finCabecera).Height = 51;
                                    }
                                    else if (caracteres <= 350)
                                    {
                                        ws.Row(finCabecera).Height = 61.25;
                                    }
                                    else if (caracteres <= 420)
                                    {
                                        ws.Row(finCabecera).Height = 71.50;
                                    }
                                    else if (caracteres <= 500)
                                    {
                                        ws.Row(finCabecera).Height = 81.75;
                                    }
                                    else
                                    {
                                        ws.Row(finCabecera).Height = 92;
                                    }

                                    break;
                                case 3:
                                    ws.Cells[finCabecera, 8].Value = column.Tipo;
                                    ws.Cells[finCabecera, 8, finCabecera, 9].Merge = true;
                                    ws.Cells[finCabecera, 8, finCabecera, 9].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                                    ws.Cells[finCabecera, 8, finCabecera, 9].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                                    // Assign border
                                    ws.Cells[finCabecera, 8, finCabecera, 9].Style.Border.Top.Style = ExcelBorderStyle.Hair;
                                    ws.Cells[finCabecera, 8, finCabecera, 9].Style.Border.Left.Style = ExcelBorderStyle.Hair;
                                    ws.Cells[finCabecera, 8, finCabecera, 9].Style.Border.Right.Style = ExcelBorderStyle.Hair;
                                    ws.Cells[finCabecera, 8, finCabecera, 9].Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                                    ws.Cells[finCabecera, 8, finCabecera, 9].Style.WrapText = true;
                                    break;
                                default:
                                    ws.Cells[finCabecera, col++].Value = "";
                                    break;
                            }

                        }
                        finCabecera++;
                        contador++;
                    }

                    ws.Row(finCabecera).Height = 8.25;
                    finCabecera += 1;

                }
                #endregion

                #region Detalle Condiciones Generales
                if (acta.Cuerpo.CondicionesGenerales.Any())
                {
                    col = 1;
                    int totalColumnasDetalleCondiciones = typeof(DetalleActaCondicionesGenerales).GetProperties().Length;

                    for (int i = 1; i <= totalColumnasDetalleCondiciones; i++)
                    {

                        ws.Cells[finCabecera, col].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        ws.Cells[finCabecera, col].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                        ws.Cells[finCabecera, col].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        // Assign borders
                        ws.Cells[finCabecera, col].Style.Border.Top.Style = ExcelBorderStyle.Hair;
                        ws.Cells[finCabecera, col].Style.Border.Left.Style = ExcelBorderStyle.Hair;
                        ws.Cells[finCabecera, col].Style.Border.Right.Style = ExcelBorderStyle.Hair;
                        ws.Cells[finCabecera, col].Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                        switch (i)
                        {
                            case 1:
                                ws.Cells[finCabecera, col++].Value = "N.";
                                ws.Cells[finCabecera, 1, finCabecera, 1].Style.Border.Top.Style = ExcelBorderStyle.Hair;
                                ws.Cells[finCabecera, 1, finCabecera, 1].Style.Border.Left.Style = ExcelBorderStyle.Hair;
                                ws.Cells[finCabecera, 1, finCabecera, 1].Style.Border.Right.Style = ExcelBorderStyle.Hair;
                                ws.Cells[finCabecera, 1, finCabecera, 1].Style.Border.Bottom.Style = ExcelBorderStyle.Hair;

                                ws.Cells[finCabecera, 1, finCabecera, 1].Style.Font.Bold = true;

                                ws.Cells[finCabecera, 1, finCabecera, 1].Style.Fill.BackgroundColor.SetColor(colorGrisClaro3EstiloPPM);
                                ws.Cells[finCabecera, 1, finCabecera, 1].Style.Font.Color.SetColor(colorGrisOscuroEstiloPPM);
                                break;
                            case 2:
                                ws.Cells[finCabecera, col++].Value = "CONDICIONES GENERALES";
                                ws.Cells[finCabecera, 2, finCabecera, 9].Merge = true;

                                ws.Cells[finCabecera, col].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                ws.Cells[finCabecera, 2, finCabecera, 9].Style.Fill.BackgroundColor.SetColor(colorGrisClaro3EstiloPPM);
                                ws.Cells[finCabecera, 2, finCabecera, 9].Style.Font.Color.SetColor(colorGrisOscuroEstiloPPM);
                                ws.Cells[finCabecera, 2, finCabecera, 9].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                                ws.Cells[finCabecera, 2, finCabecera, 9].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                // Assign border
                                ws.Cells[finCabecera, 2, finCabecera, 9].Style.Border.Top.Style = ExcelBorderStyle.Hair;
                                ws.Cells[finCabecera, 2, finCabecera, 9].Style.Border.Left.Style = ExcelBorderStyle.Hair;
                                ws.Cells[finCabecera, 2, finCabecera, 9].Style.Border.Right.Style = ExcelBorderStyle.Hair;
                                ws.Cells[finCabecera, 2, finCabecera, 9].Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                                ws.Cells[finCabecera, 2, finCabecera, 9].Style.Font.Bold = true;
                                break;
                            default:
                                ws.Cells[finCabecera, col++].Value = "";
                                break;
                        }

                        ws.Row(finCabecera).Height = 20.25;
                    }

                    finCabecera++;

                    contador = 1;
                    foreach (var column in acta.Cuerpo.CondicionesGenerales)
                    {
                        col = 1;
                        for (int i = 1; i <= totalColumnasDetalleCondiciones; i++)
                        {
                            ws.Cells[finCabecera, col].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                            ws.Cells[finCabecera, col].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            // Assign borders
                            ws.Cells[finCabecera, col].Style.Border.Top.Style = ExcelBorderStyle.Hair;
                            ws.Cells[finCabecera, col].Style.Border.Left.Style = ExcelBorderStyle.Hair;
                            ws.Cells[finCabecera, col].Style.Border.Right.Style = ExcelBorderStyle.Hair;
                            ws.Cells[finCabecera, col].Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                            switch (i)
                            {
                                case 1:
                                    ws.Cells[finCabecera, col++].Value = contador;
                                    ws.Cells[finCabecera, 1, finCabecera, 1].Style.Border.Top.Style = ExcelBorderStyle.Hair;
                                    ws.Cells[finCabecera, 1, finCabecera, 1].Style.Border.Left.Style = ExcelBorderStyle.Hair;
                                    ws.Cells[finCabecera, 1, finCabecera, 1].Style.Border.Right.Style = ExcelBorderStyle.Hair;
                                    ws.Cells[finCabecera, 1, finCabecera, 1].Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                                    break;
                                case 2:
                                    ws.Cells[finCabecera, col++].Value = column.Condicion;
                                    ws.Cells[finCabecera, 2, finCabecera, 9].Merge = true;
                                    ws.Cells[finCabecera, 2, finCabecera, 9].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                                    ws.Cells[finCabecera, 2, finCabecera, 9].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                                    ws.Cells[finCabecera, 2, finCabecera, 9].Style.Border.Top.Style = ExcelBorderStyle.Hair;
                                    ws.Cells[finCabecera, 2, finCabecera, 9].Style.Border.Left.Style = ExcelBorderStyle.Hair;
                                    ws.Cells[finCabecera, 2, finCabecera, 9].Style.Border.Right.Style = ExcelBorderStyle.Hair;
                                    ws.Cells[finCabecera, 2, finCabecera, 9].Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                                    ws.Cells[finCabecera, 2, finCabecera, 9].Style.WrapText = true;

                                    var caracteres = column.Condicion.Count();
                                    if (caracteres <= 97)
                                    {
                                        ws.Row(finCabecera).Height = 20.25;
                                    }
                                    else if (caracteres <= 194)
                                    {
                                        ws.Row(finCabecera).Height = 30.50;
                                    }
                                    else if (caracteres <= 291)
                                    {
                                        ws.Row(finCabecera).Height = 40.75;
                                    }
                                    else if (caracteres <= 388)
                                    {
                                        ws.Row(finCabecera).Height = 51;
                                    }
                                    else if (caracteres <= 582)
                                    {
                                        ws.Row(finCabecera).Height = 61.25;
                                    }
                                    else
                                    {
                                        ws.Row(finCabecera).Height = 71.5;
                                    }

                                    break;
                                default:
                                    ws.Cells[finCabecera, col++].Value = "";
                                    break;
                            }

                        }
                        finCabecera++;
                        contador++;
                    }
                    ws.Row(finCabecera).Height = 8.25;
                    finCabecera += 1;
                }


                #endregion

                #region Detalle participantes
                if (acta.Cuerpo.Participantes.Any())
                {
                    col = 1;
                    int totalColumnasDetalleParticipantes = typeof(DetalleActaParticipantes).GetProperties().Length - 1;

                    int contadorRegistrosParticipantes = acta.Cuerpo.Participantes.Count;

                    if (contadorRegistrosParticipantes > 3)
                    {

                        for (int i = 1; i <= 2; i++)
                        {
                            int totalMergeColumnasParticipantes = i == 1 ? 3 : 8;

                            using (var range = ws.Cells[finCabecera, col, finCabecera, totalMergeColumnasParticipantes])
                            {
                                range.Value = "PARTICIPANTES";
                                range.Style.Fill.PatternType = ExcelFillStyle.Solid;
                                range.Style.Fill.BackgroundColor.SetColor(colorGrisClaro3EstiloPPM);
                                range.Style.Font.Color.SetColor(colorGrisOscuroEstiloPPM);
                                range.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                                range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                                // Assign borders
                                range.Style.Border.Top.Style = ExcelBorderStyle.Hair;
                                range.Style.Border.Left.Style = ExcelBorderStyle.Hair;
                                range.Style.Border.Right.Style = ExcelBorderStyle.Hair;
                                range.Style.Border.Bottom.Style = ExcelBorderStyle.Hair;

                                range.Merge = true;
                                range.Style.Font.Bold = true;
                            }

                            if (totalMergeColumnasParticipantes == 8)
                                col++;

                            col += 3;

                            using (var range = ws.Cells[finCabecera, col, finCabecera, col])
                            {
                                range.Value = "PRESENTE";
                                range.Style.Fill.PatternType = ExcelFillStyle.Solid;
                                range.Style.Fill.BackgroundColor.SetColor(colorGrisClaro3EstiloPPM);
                                range.Style.Font.Color.SetColor(colorGrisOscuroEstiloPPM);
                                range.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                                range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                                // Assign borders
                                range.Style.Border.Top.Style = ExcelBorderStyle.Hair;
                                range.Style.Border.Left.Style = ExcelBorderStyle.Hair;
                                range.Style.Border.Right.Style = ExcelBorderStyle.Hair;
                                range.Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                                range.Style.Font.Bold = true;
                            }
                            col += 1;
                        }

                        finCabecera++;

                        contador = 1;
                        int inicioFilaParticipantes = finCabecera;
                        int columnasDivision1 = contadorRegistrosParticipantes / 2;
                        int columnasDivision2 = contadorRegistrosParticipantes - (contadorRegistrosParticipantes / 2);

                        foreach (var column in acta.Cuerpo.Participantes)
                        {
                            col = contador <= columnasDivision1 ? 1 : 5;

                            if (contador == columnasDivision2 + 1)
                            {
                                finCabecera = inicioFilaParticipantes;
                            }

                            for (int i = 1; i <= totalColumnasDetalleParticipantes; i++)
                            {
                                ws.Cells[finCabecera, col].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                                ws.Cells[finCabecera, col].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                                // Assign borders
                                ws.Cells[finCabecera, col].Style.Border.Top.Style = ExcelBorderStyle.Hair;
                                ws.Cells[finCabecera, col].Style.Border.Left.Style = ExcelBorderStyle.Hair;
                                ws.Cells[finCabecera, col].Style.Border.Right.Style = ExcelBorderStyle.Hair;
                                ws.Cells[finCabecera, col].Style.Border.Bottom.Style = ExcelBorderStyle.Hair;

                                if (col == 8)
                                    col++;

                                switch (i)
                                {
                                    case 1:
                                        ws.Cells[finCabecera, col].Value = column.Nombres;
                                        ws.Cells[finCabecera, col, finCabecera, col + 2].Merge = true;
                                        ws.Cells[finCabecera, col, finCabecera, col + 2].Style.Border.Top.Style = ExcelBorderStyle.Hair;
                                        ws.Cells[finCabecera, col, finCabecera, col + 2].Style.Border.Left.Style = ExcelBorderStyle.Hair;
                                        ws.Cells[finCabecera, col, finCabecera, col + 2].Style.Border.Right.Style = ExcelBorderStyle.Hair;
                                        ws.Cells[finCabecera, col, finCabecera, col + 2].Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                                        ws.Cells[finCabecera, col, finCabecera, col + 2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                                        ws.Cells[finCabecera, col, finCabecera, col + 2].Style.WrapText = true;
                                        if (col == 5)

                                            ws.Cells[finCabecera, 5, finCabecera, 8].Merge = true;
                                        ws.Cells[finCabecera, 5, finCabecera, 8].Style.Border.Top.Style = ExcelBorderStyle.Hair;
                                        ws.Cells[finCabecera, 5, finCabecera, 8].Style.Border.Left.Style = ExcelBorderStyle.Hair;
                                        ws.Cells[finCabecera, 5, finCabecera, 8].Style.Border.Right.Style = ExcelBorderStyle.Hair;
                                        ws.Cells[finCabecera, 5, finCabecera, 8].Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                                        ws.Cells[finCabecera, 5, finCabecera, 8].Style.WrapText = true;
                                        col += 3;

                                        break;
                                    case 2:
                                        ws.Cells[finCabecera, col].Value = column.Presente ? "SI" : "NO";
                                        ws.Cells[finCabecera, col].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                                        ws.Cells[finCabecera, col].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                        ws.Cells[finCabecera, col].Style.Border.Top.Style = ExcelBorderStyle.Hair;
                                        ws.Cells[finCabecera, col].Style.Border.Left.Style = ExcelBorderStyle.Hair;
                                        ws.Cells[finCabecera, col].Style.Border.Right.Style = ExcelBorderStyle.Hair;
                                        ws.Cells[finCabecera, col].Style.Border.Bottom.Style = ExcelBorderStyle.Hair;

                                        break;
                                    default:
                                        ws.Cells[finCabecera, col++].Value = "";
                                        break;
                                }

                            }
                            ws.Cells[finCabecera, 1, finCabecera, 9].Style.WrapText = true;
                            ws.Row(finCabecera).Height = 20.25;
                            finCabecera++;
                            contador++;
                        }
                    }
                    else
                    {
                        for (int j = 1; j <= totalColumnasDetalleParticipantes; j++)
                        {
                            ws.Cells[finCabecera, col].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            ws.Cells[finCabecera, col].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                            ws.Cells[finCabecera, col].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                            // Assign borders
                            ws.Cells[finCabecera, col].Style.Border.Top.Style = ExcelBorderStyle.Hair;
                            ws.Cells[finCabecera, col].Style.Border.Left.Style = ExcelBorderStyle.Hair;
                            ws.Cells[finCabecera, col].Style.Border.Right.Style = ExcelBorderStyle.Hair;
                            ws.Cells[finCabecera, col].Style.Border.Bottom.Style = ExcelBorderStyle.Hair;

                            switch (j)
                            {
                                case 1:
                                    using (var range = ws.Cells[finCabecera, col++, finCabecera, col + 5])
                                    {
                                        range.Style.Border.Top.Style = ExcelBorderStyle.Hair;
                                        range.Style.Border.Left.Style = ExcelBorderStyle.Hair;
                                        range.Style.Border.Right.Style = ExcelBorderStyle.Hair;
                                        range.Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                                        range.Merge = true;

                                        range.Value = "PARTICIPANTES";
                                        range.Style.Fill.BackgroundColor.SetColor(colorGrisClaro3EstiloPPM);
                                        range.Style.Font.Color.SetColor(colorGrisOscuroEstiloPPM);
                                        range.Style.Font.Bold = true;

                                    }
                                    using (var range = ws.Cells[finCabecera, col + 6, finCabecera, col + 7])
                                    {
                                        range.Style.Border.Top.Style = ExcelBorderStyle.Hair;
                                        range.Style.Border.Left.Style = ExcelBorderStyle.Hair;
                                        range.Style.Border.Right.Style = ExcelBorderStyle.Hair;
                                        range.Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                                        range.Merge = true;

                                        range.Value = "PRESENTES";
                                        range.Style.Fill.BackgroundColor.SetColor(colorGrisClaro3EstiloPPM);
                                        range.Style.Font.Color.SetColor(colorGrisOscuroEstiloPPM);
                                        range.Style.HorizontalAlignment = ExcelHorizontalAlignment.CenterContinuous;
                                        range.Style.Font.Bold = true;

                                    }

                                    break;
                                case 2:
                                    //ws.Cells[finCabecera, col++].Value = "PRESENTES";
                                    break;
                                default:
                                    ws.Cells[finCabecera, col++].Value = "";
                                    break;
                            }
                        }

                        finCabecera++;
                        contador = 1;
                        foreach (var column in acta.Cuerpo.Participantes)
                        {
                            col = 1;
                            for (int i = 1; i <= totalColumnasDetalleParticipantes; i++)
                            {
                                ws.Cells[finCabecera, col].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                                ws.Cells[finCabecera, col].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;

                                // Assign borders
                                ws.Cells[finCabecera, col].Style.Border.Top.Style = ExcelBorderStyle.Hair;
                                ws.Cells[finCabecera, col].Style.Border.Left.Style = ExcelBorderStyle.Hair;
                                ws.Cells[finCabecera, col].Style.Border.Right.Style = ExcelBorderStyle.Hair;
                                ws.Cells[finCabecera, col].Style.Border.Bottom.Style = ExcelBorderStyle.Hair;

                                switch (i)
                                {

                                    case 1:

                                        using (var range = ws.Cells[finCabecera, col++, finCabecera, col + 5])
                                        {
                                            range.Value = column.Nombres;

                                            range.Style.Border.Top.Style = ExcelBorderStyle.Hair;
                                            range.Style.Border.Left.Style = ExcelBorderStyle.Hair;
                                            range.Style.Border.Right.Style = ExcelBorderStyle.Hair;
                                            range.Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                                            range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                                            range.Merge = true;
                                            range.Style.WrapText = true;

                                        }
                                        break;
                                    case 2:

                                        using (var range = ws.Cells[finCabecera, col + 6, finCabecera, col + 7])
                                        {
                                            range.Value = column.Presente ? "SI" : "NO";

                                            range.Style.Border.Top.Style = ExcelBorderStyle.Hair;
                                            range.Style.Border.Left.Style = ExcelBorderStyle.Hair;
                                            range.Style.Border.Right.Style = ExcelBorderStyle.Hair;
                                            range.Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                                            range.Style.HorizontalAlignment = ExcelHorizontalAlignment.CenterContinuous;
                                            range.Merge = true;
                                            range.Style.WrapText = true;

                                        }
                                        //ws.Cells[finCabecera, col++].Value = column.Presente ? "SI" : "NO";
                                        break;
                                    default:
                                        ws.Cells[finCabecera, col++].Value = "";

                                        break;
                                }
                            }

                            ws.Cells[finCabecera, 1, finCabecera, 9].Style.WrapText = true;
                            ws.Row(finCabecera).Height = 20.25;
                            finCabecera++;
                            contador++;

                        }

                    }

                    if (acta.Cuerpo.Participantes.Any())
                    {
                        if (acta.Cuerpo.Participantes.Count > 3 && acta.Cuerpo.Participantes.Count % 2 != 0)
                            finCabecera += 2;
                        else
                        {
                            ws.Row(finCabecera).Height = 8.25;
                            finCabecera++;
                        }
                    }
                    else
                    {
                        ws.Row(finCabecera).Height = 8.25;
                        finCabecera++;
                    }

                    int filaDatosParticipantes = finCabecera;
                    int inicio = filaDatosParticipantes;

                    int cantidadParticipantes = acta.Cuerpo.Participantes.Count;
                    int cantidadSI = acta.Cuerpo.Participantes.Where(s => s.Presente).Count();
                    int cantidadNO = acta.Cuerpo.Participantes.Where(s => !s.Presente).Count();

                    float porcentajeSI = ((float)cantidadSI / (float)cantidadParticipantes) * 100;
                    float porcentajeNO = ((float)cantidadNO / (float)cantidadParticipantes) * 100;

                    string resultadoSI = string.Format("SI % {0}", Math.Round(porcentajeSI, 2));
                    string resultadoNO = string.Format("NO % {0}", Math.Round(porcentajeNO, 2));
                    string resultadoFinal = porcentajeNO > porcentajeSI ? "SI" : "NO";

                    for (int i = 1; i <= 3; i++)
                    {
                        ws.Cells[filaDatosParticipantes, i].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        ws.Cells[filaDatosParticipantes, i].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(128, 128, 128));
                        ws.Cells[filaDatosParticipantes, i].Style.Font.Color.SetColor(Color.White);
                        ws.Cells[filaDatosParticipantes, i].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                        ws.Cells[filaDatosParticipantes, i].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                        ws.Cells[filaDatosParticipantes, i].Style.Border.Top.Style = ExcelBorderStyle.Hair;
                        ws.Cells[filaDatosParticipantes, i].Style.Border.Left.Style = ExcelBorderStyle.Hair;
                        ws.Cells[filaDatosParticipantes, i].Style.Border.Right.Style = ExcelBorderStyle.Hair;
                        ws.Cells[filaDatosParticipantes, i].Style.Border.Bottom.Style = ExcelBorderStyle.Hair;

                        switch (i)
                        {
                            case 1:
                                ws.Cells[filaDatosParticipantes, i].Value = "Asistencia";
                                ws.Cells[filaDatosParticipantes, i].Style.Font.Bold = true;

                                ws.Cells[filaDatosParticipantes, i].Style.Fill.BackgroundColor.SetColor(colorGrisOscuroEstiloPPM);
                                ws.Cells[filaDatosParticipantes, i].Style.Font.Color.SetColor(colorBlancoEstiloPPM);

                                filaDatosParticipantes++;

                                ws.Cells[filaDatosParticipantes, i].Value = resultadoSI;
                                ws.Cells[filaDatosParticipantes, i].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                                ws.Cells[filaDatosParticipantes, i].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                                // Assign borders
                                ws.Cells[filaDatosParticipantes, i].Style.Border.Top.Style = ExcelBorderStyle.Hair;
                                ws.Cells[filaDatosParticipantes, i].Style.Border.Left.Style = ExcelBorderStyle.Hair;
                                ws.Cells[filaDatosParticipantes, i].Style.Border.Right.Style = ExcelBorderStyle.Hair;
                                ws.Cells[filaDatosParticipantes, i].Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                                filaDatosParticipantes++;

                                ws.Cells[filaDatosParticipantes, i].Value = resultadoNO;
                                ws.Cells[filaDatosParticipantes, i].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                                ws.Cells[filaDatosParticipantes, i].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                // Assign borders
                                ws.Cells[filaDatosParticipantes, i].Style.Border.Top.Style = ExcelBorderStyle.Hair;
                                ws.Cells[filaDatosParticipantes, i].Style.Border.Left.Style = ExcelBorderStyle.Hair;
                                ws.Cells[filaDatosParticipantes, i].Style.Border.Right.Style = ExcelBorderStyle.Hair;
                                ws.Cells[filaDatosParticipantes, i].Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                                break;
                            case 2:

                                ws.Cells[filaDatosParticipantes, i].Value = "¿Se suspende?";
                                ws.Cells[filaDatosParticipantes, i].Style.Font.Bold = true;

                                ws.Cells[filaDatosParticipantes, i].Style.Fill.BackgroundColor.SetColor(colorGrisOscuroEstiloPPM);
                                ws.Cells[filaDatosParticipantes, i].Style.Font.Color.SetColor(colorBlancoEstiloPPM);

                                filaDatosParticipantes++;
                                ws.Cells[filaDatosParticipantes, i].Value = resultadoFinal;
                                ws.Cells[filaDatosParticipantes, i].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                                ws.Cells[filaDatosParticipantes, i].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                // Assign borders
                                ws.Cells[filaDatosParticipantes, i].Style.Border.Top.Style = ExcelBorderStyle.Hair;
                                ws.Cells[filaDatosParticipantes, i].Style.Border.Left.Style = ExcelBorderStyle.Hair;
                                ws.Cells[filaDatosParticipantes, i].Style.Border.Right.Style = ExcelBorderStyle.Hair;
                                ws.Cells[filaDatosParticipantes, i].Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                                break;
                            case 3:
                                ws.Cells[filaDatosParticipantes, 3, filaDatosParticipantes, 9].Value = "Observaciones";
                                ws.Cells[filaDatosParticipantes, 3, filaDatosParticipantes, 9].Merge = true;
                                ws.Cells[filaDatosParticipantes, 3, filaDatosParticipantes, 9].Style.Font.Bold = true;

                                ws.Cells[filaDatosParticipantes, 3, filaDatosParticipantes, 9].Style.Fill.BackgroundColor.SetColor(colorGrisOscuroEstiloPPM);
                                ws.Cells[filaDatosParticipantes, 3, filaDatosParticipantes, 9].Style.Font.Color.SetColor(colorBlancoEstiloPPM);

                                ws.Cells[filaDatosParticipantes, 3, filaDatosParticipantes, 9].Style.Border.Top.Style = ExcelBorderStyle.Hair;
                                ws.Cells[filaDatosParticipantes, 3, filaDatosParticipantes, 9].Style.Border.Left.Style = ExcelBorderStyle.Hair;
                                ws.Cells[filaDatosParticipantes, 3, filaDatosParticipantes, 9].Style.Border.Right.Style = ExcelBorderStyle.Hair;
                                ws.Cells[filaDatosParticipantes, 3, filaDatosParticipantes, 9].Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                                ws.Cells[filaDatosParticipantes, 3, filaDatosParticipantes, 9].Style.HorizontalAlignment = ExcelHorizontalAlignment.CenterContinuous;
                                filaDatosParticipantes++;
                                ws.Cells[filaDatosParticipantes, 3, filaDatosParticipantes, 9].Value = acta.Cabecera.Observaciones;
                                ws.Cells[filaDatosParticipantes, 3, filaDatosParticipantes, 9].Merge = true;
                                ws.Cells[filaDatosParticipantes, 3, filaDatosParticipantes, 9].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;

                                // Assign borders
                                ws.Cells[filaDatosParticipantes, 3, filaDatosParticipantes, 9].Style.Border.Top.Style = ExcelBorderStyle.Hair;
                                ws.Cells[filaDatosParticipantes, 3, filaDatosParticipantes, 9].Style.Border.Left.Style = ExcelBorderStyle.Hair;
                                ws.Cells[filaDatosParticipantes, 3, filaDatosParticipantes, 9].Style.Border.Right.Style = ExcelBorderStyle.Hair;
                                ws.Cells[filaDatosParticipantes, 3, filaDatosParticipantes, 9].Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                                ws.Cells[filaDatosParticipantes, 3, filaDatosParticipantes, 9].Style.WrapText = true;

                                var caracteres = acta.Cabecera.Observaciones.Count();

                                if (caracteres <= 104)
                                {
                                    ws.Row(filaDatosParticipantes).Height = 20.25;
                                }
                                else if (caracteres <= 208)
                                {
                                    ws.Row(filaDatosParticipantes).Height = 30;
                                }
                                else if (caracteres <= 312)
                                {
                                    ws.Row(filaDatosParticipantes).Height = 40;
                                }
                                else if (caracteres <= 416)
                                {
                                    ws.Row(filaDatosParticipantes).Height = 50;
                                }
                                else if (caracteres <= 500)
                                {
                                    ws.Row(filaDatosParticipantes).Height = 60;
                                }
                                else
                                {
                                    ws.Row(filaDatosParticipantes).Height = 70;
                                }


                                break;
                            default:
                                ws.Cells[filaDatosParticipantes, i].Value = ":";
                                ws.Cells[filaDatosParticipantes, i].Style.Font.Bold = true;
                                filaDatosParticipantes++;
                                ws.Cells[filaDatosParticipantes, i].Value = "";
                                break;

                        }
                        ws.Row(finCabecera).Height = 20.25;
                        finCabecera++;
                        filaDatosParticipantes = inicio;
                        ws.Column(5).AutoFit();
                        ws.Column(6).AutoFit();

                    }
                    ws.Row(finCabecera).Height = 8.25;
                    finCabecera = finCabecera + 1;
                    ws.Column(4).Width = 14.30;

                }
                #endregion

                #region Detalle Temas a tratar
                if (acta.Cuerpo.Temas.Any())
                {
                    col = 1;
                    int totalColumnasDetalleTemas = typeof(DetalleActaTemasTratar).GetProperties().Length;

                    for (int i = 1; i <= totalColumnasDetalleTemas; i++)
                    {

                        ws.Cells[finCabecera, col].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        ws.Cells[finCabecera, col].Style.Fill.BackgroundColor.SetColor(colorGrisClaro3EstiloPPM);
                        ws.Cells[finCabecera, col].Style.Font.Color.SetColor(colorGrisOscuroEstiloPPM);
                        ws.Cells[finCabecera, col].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                        ws.Cells[finCabecera, col].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        // Assign borders
                        ws.Cells[finCabecera, col].Style.Border.Top.Style = ExcelBorderStyle.Hair;
                        ws.Cells[finCabecera, col].Style.Border.Left.Style = ExcelBorderStyle.Hair;
                        ws.Cells[finCabecera, col].Style.Border.Right.Style = ExcelBorderStyle.Hair;
                        ws.Cells[finCabecera, col].Style.Border.Bottom.Style = ExcelBorderStyle.Hair;

                        switch (i)
                        {
                            case 1:
                                ws.Cells[finCabecera, col++].Value = "N.";
                                ws.Cells[finCabecera, 1, finCabecera, 1].Style.Border.Top.Style = ExcelBorderStyle.Hair;
                                ws.Cells[finCabecera, 1, finCabecera, 1].Style.Border.Left.Style = ExcelBorderStyle.Hair;
                                ws.Cells[finCabecera, 1, finCabecera, 1].Style.Border.Right.Style = ExcelBorderStyle.Hair;
                                ws.Cells[finCabecera, 1, finCabecera, 1].Style.Border.Bottom.Style = ExcelBorderStyle.Hair;

                                ws.Cells[finCabecera, 1, finCabecera, 1].Style.Fill.BackgroundColor.SetColor(colorGrisClaro3EstiloPPM);
                                ws.Cells[finCabecera, 1, finCabecera, 1].Style.Font.Color.SetColor(colorGrisOscuroEstiloPPM);
                                ws.Cells[finCabecera, 1, finCabecera, 1].Style.Font.Bold = true;
                                break;
                            case 2:
                                //ws.Column(col).Width2+ = 30;
                                ws.Cells[finCabecera, col++].Value = "TEMAS O PUNTOS A TRATAR";
                                ws.Cells[finCabecera, 2, finCabecera, 5].Merge = true;
                                ws.Cells[finCabecera, 2, finCabecera, 5].Style.Fill.PatternType = ExcelFillStyle.Solid;

                                ws.Cells[finCabecera, 2, finCabecera, 5].Style.Fill.BackgroundColor.SetColor(colorGrisClaro3EstiloPPM);
                                ws.Cells[finCabecera, 2, finCabecera, 5].Style.Font.Color.SetColor(colorGrisOscuroEstiloPPM);
                                ws.Cells[finCabecera, 2, finCabecera, 5].Style.Font.Bold = true;

                                ws.Cells[finCabecera, 2, finCabecera, 5].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                                ws.Cells[finCabecera, 2, finCabecera, 5].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                // Assign borders
                                ws.Cells[finCabecera, 2, finCabecera, 5].Style.Border.Top.Style = ExcelBorderStyle.Hair;
                                ws.Cells[finCabecera, 2, finCabecera, 5].Style.Border.Left.Style = ExcelBorderStyle.Hair;
                                ws.Cells[finCabecera, 2, finCabecera, 5].Style.Border.Right.Style = ExcelBorderStyle.Hair;
                                ws.Cells[finCabecera, 2, finCabecera, 5].Style.Border.Bottom.Style = ExcelBorderStyle.Hair;


                                break;
                            case 3:
                                ws.Cells[finCabecera, 6].Value = "RESPONSABLE(S)";
                                ws.Cells[finCabecera, 6, finCabecera, 9].Merge = true;
                                ws.Cells[finCabecera, 6, finCabecera, 9].Style.Fill.PatternType = ExcelFillStyle.Solid;

                                ws.Cells[finCabecera, 6, finCabecera, 9].Style.Fill.BackgroundColor.SetColor(colorGrisClaro3EstiloPPM);
                                ws.Cells[finCabecera, 6, finCabecera, 9].Style.Font.Color.SetColor(colorGrisOscuroEstiloPPM);
                                ws.Cells[finCabecera, 6, finCabecera, 9].Style.Font.Bold = true;

                                ws.Cells[finCabecera, 6, finCabecera, 9].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                                ws.Cells[finCabecera, 6, finCabecera, 9].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                // Assign borders
                                ws.Cells[finCabecera, 6, finCabecera, 9].Style.Border.Top.Style = ExcelBorderStyle.Hair;
                                ws.Cells[finCabecera, 6, finCabecera, 9].Style.Border.Left.Style = ExcelBorderStyle.Hair;
                                ws.Cells[finCabecera, 6, finCabecera, 9].Style.Border.Right.Style = ExcelBorderStyle.Hair;
                                ws.Cells[finCabecera, 6, finCabecera, 9].Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                                break;
                            default:
                                ws.Cells[finCabecera, col++].Value = "";
                                break;
                        }

                        ws.Row(finCabecera).Height = 20.25;
                    }

                    finCabecera++;

                    contador = 1;
                    foreach (var column in acta.Cuerpo.Temas)
                    {
                        col = 1;
                        for (int i = 1; i <= totalColumnasDetalleTemas; i++)
                        {
                            ws.Cells[finCabecera, col].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                            ws.Cells[finCabecera, col].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            // Assign borders
                            ws.Cells[finCabecera, col].Style.Border.Top.Style = ExcelBorderStyle.Hair;
                            ws.Cells[finCabecera, col].Style.Border.Left.Style = ExcelBorderStyle.Hair;
                            ws.Cells[finCabecera, col].Style.Border.Right.Style = ExcelBorderStyle.Hair;
                            ws.Cells[finCabecera, col].Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                            switch (i)
                            {
                                case 1:
                                    ws.Cells[finCabecera, col++].Value = contador;
                                    ws.Cells[finCabecera, 1, finCabecera, 1].Merge = true;
                                    ws.Cells[finCabecera, 1, finCabecera, 1].Style.Border.Top.Style = ExcelBorderStyle.Hair;
                                    ws.Cells[finCabecera, 1, finCabecera, 1].Style.Border.Left.Style = ExcelBorderStyle.Hair;
                                    ws.Cells[finCabecera, 1, finCabecera, 1].Style.Border.Right.Style = ExcelBorderStyle.Hair;
                                    ws.Cells[finCabecera, 1, finCabecera, 1].Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                                    break;
                                case 2:
                                    ws.Cells[finCabecera, col++].Value = column.Tema;
                                    ws.Cells[finCabecera, 2, finCabecera, 5].Merge = true;
                                    ws.Cells[finCabecera, 2, finCabecera, 5].Style.Border.Top.Style = ExcelBorderStyle.Hair;
                                    ws.Cells[finCabecera, 2, finCabecera, 5].Style.Border.Left.Style = ExcelBorderStyle.Hair;
                                    ws.Cells[finCabecera, 2, finCabecera, 5].Style.Border.Right.Style = ExcelBorderStyle.Hair;
                                    ws.Cells[finCabecera, 2, finCabecera, 5].Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                                    ws.Cells[finCabecera, 2, finCabecera, 5].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                                    ws.Cells[finCabecera, 2, finCabecera, 5].Style.WrapText = true;

                                    var caracteres = column.Tema.Count();

                                    if (caracteres <= 70)
                                    {
                                        ws.Row(finCabecera).Height = 20.25;
                                    }
                                    else if (caracteres <= 140)
                                    {
                                        ws.Row(finCabecera).Height = 40;
                                    }
                                    else if (caracteres <= 210)
                                    {
                                        ws.Row(finCabecera).Height = 55;
                                    }
                                    else if (caracteres <= 280)
                                    {
                                        ws.Row(finCabecera).Height = 70;
                                    }
                                    else if (caracteres <= 350)
                                    {
                                        ws.Row(finCabecera).Height = 85;
                                    }
                                    else if (caracteres <= 420)
                                    {
                                        ws.Row(finCabecera).Height = 100;
                                    }
                                    else if (caracteres <= 500)
                                    {
                                        ws.Row(finCabecera).Height = 115;
                                    }
                                    else
                                    {
                                        ws.Row(finCabecera).Height = 130;
                                    }
                                    break;
                                case 3:
                                    ws.Cells[finCabecera, 6].Value = column.Responsable;

                                    ws.Cells[finCabecera, 6, finCabecera, 9].Merge = true;
                                    ws.Cells[finCabecera, 6, finCabecera, 9].Style.Border.Top.Style = ExcelBorderStyle.Hair;
                                    ws.Cells[finCabecera, 6, finCabecera, 9].Style.Border.Left.Style = ExcelBorderStyle.Hair;
                                    ws.Cells[finCabecera, 6, finCabecera, 9].Style.Border.Right.Style = ExcelBorderStyle.Hair;
                                    ws.Cells[finCabecera, 6, finCabecera, 9].Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                                    ws.Cells[finCabecera, 6, finCabecera, 9].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                                    ws.Cells[finCabecera, 6, finCabecera, 9].Style.WrapText = true;
                                    break;
                                default:
                                    ws.Cells[finCabecera, col++].Value = "";
                                    break;
                            }
                        }

                        //ws.Row(finCabecera).Height = 39.5;
                        finCabecera++;
                        contador++;
                    }

                    ws.Row(finCabecera).Height = 8.25;
                    finCabecera += 1;


                }
                #endregion

                #region Detalle Acuerdos
                if (acta.Cuerpo.Acuerdos.Any())
                {
                    col = 1;
                    int totalColumnasDetalleAcuerdos = typeof(DetalleActaAcuerdos).GetProperties().Length;

                    for (int i = 1; i <= totalColumnasDetalleAcuerdos; i++)
                    {
                        //ws.Column(col).Width = 18;

                        ws.Cells[finCabecera, col].Style.Fill.PatternType = ExcelFillStyle.Solid;

                        ws.Cells[finCabecera, col].Style.Fill.BackgroundColor.SetColor(colorGrisClaro3EstiloPPM);
                        ws.Cells[finCabecera, col].Style.Font.Color.SetColor(colorGrisOscuroEstiloPPM);

                        ws.Cells[finCabecera, col].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                        ws.Cells[finCabecera, col].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        // Assign borders
                        ws.Cells[finCabecera, col].Style.Border.Top.Style = ExcelBorderStyle.Hair;
                        ws.Cells[finCabecera, col].Style.Border.Left.Style = ExcelBorderStyle.Hair;
                        ws.Cells[finCabecera, col].Style.Border.Right.Style = ExcelBorderStyle.Hair;
                        ws.Cells[finCabecera, col].Style.Border.Bottom.Style = ExcelBorderStyle.Hair;

                        switch (i)
                        {
                            case 1:
                                ws.Cells[finCabecera, col++].Value = "N.";
                                ws.Cells[finCabecera, 1, finCabecera, 1].Style.Border.Top.Style = ExcelBorderStyle.Hair;
                                ws.Cells[finCabecera, 1, finCabecera, 1].Style.Border.Left.Style = ExcelBorderStyle.Hair;
                                ws.Cells[finCabecera, 1, finCabecera, 1].Style.Border.Right.Style = ExcelBorderStyle.Hair;
                                ws.Cells[finCabecera, 1, finCabecera, 1].Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                                ws.Cells[finCabecera, 1, finCabecera, 1].Style.Font.Bold = true;

                                ws.Cells[finCabecera, 1, finCabecera, 1].Style.Fill.BackgroundColor.SetColor(colorGrisClaro3EstiloPPM);
                                ws.Cells[finCabecera, 1, finCabecera, 1].Style.Font.Color.SetColor(colorGrisOscuroEstiloPPM);
                                break;
                            case 2:
                                ws.Cells[finCabecera, col++].Value = "ACUERDOS";
                                ws.Cells[finCabecera, 2, finCabecera, 5].Merge = true;
                                ws.Cells[finCabecera, 2, finCabecera, 5].Style.Fill.PatternType = ExcelFillStyle.Solid;

                                ws.Cells[finCabecera, 2, finCabecera, 5].Style.Fill.BackgroundColor.SetColor(colorGrisClaro3EstiloPPM);
                                ws.Cells[finCabecera, 2, finCabecera, 5].Style.Font.Color.SetColor(colorGrisOscuroEstiloPPM);

                                ws.Cells[finCabecera, 2, finCabecera, 5].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                                ws.Cells[finCabecera, 2, finCabecera, 5].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                // Assign borders
                                ws.Cells[finCabecera, 2, finCabecera, 5].Style.Border.Top.Style = ExcelBorderStyle.Hair;
                                ws.Cells[finCabecera, 2, finCabecera, 5].Style.Border.Left.Style = ExcelBorderStyle.Hair;
                                ws.Cells[finCabecera, 2, finCabecera, 5].Style.Border.Right.Style = ExcelBorderStyle.Hair;
                                ws.Cells[finCabecera, 2, finCabecera, 5].Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                                ws.Cells[finCabecera, 2, finCabecera, 5].Style.Font.Bold = true;
                                break;

                            case 3:
                                ws.Cells[finCabecera, 6].Value = "RESPONSABLES";
                                ws.Cells[finCabecera, 6, finCabecera, 7].Merge = true;
                                ws.Cells[finCabecera, 6, finCabecera, 7].Style.Fill.PatternType = ExcelFillStyle.Solid;

                                ws.Cells[finCabecera, 6, finCabecera, 7].Style.Fill.BackgroundColor.SetColor(colorGrisClaro3EstiloPPM);
                                ws.Cells[finCabecera, 6, finCabecera, 7].Style.Font.Color.SetColor(colorGrisOscuroEstiloPPM);

                                ws.Cells[finCabecera, 6, finCabecera, 7].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                                ws.Cells[finCabecera, 6, finCabecera, 7].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                // Assign borders
                                ws.Cells[finCabecera, 6, finCabecera, 7].Style.Border.Top.Style = ExcelBorderStyle.Hair;
                                ws.Cells[finCabecera, 6, finCabecera, 7].Style.Border.Left.Style = ExcelBorderStyle.Hair;
                                ws.Cells[finCabecera, 6, finCabecera, 7].Style.Border.Right.Style = ExcelBorderStyle.Hair;
                                ws.Cells[finCabecera, 6, finCabecera, 7].Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                                ws.Cells[finCabecera, 6, finCabecera, 7].Style.Font.Bold = true;
                                break;
                            case 4:
                                ws.Cells[finCabecera, 8].Value = "FECHAS";
                                ws.Cells[finCabecera, 8, finCabecera, 9].Merge = true;
                                ws.Cells[finCabecera, 8, finCabecera, 9].Style.Fill.PatternType = ExcelFillStyle.Solid;

                                ws.Cells[finCabecera, 8, finCabecera, 9].Style.Fill.BackgroundColor.SetColor(colorGrisClaro3EstiloPPM);
                                ws.Cells[finCabecera, 8, finCabecera, 9].Style.Font.Color.SetColor(colorGrisOscuroEstiloPPM);

                                ws.Cells[finCabecera, 8, finCabecera, 9].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                                ws.Cells[finCabecera, 8, finCabecera, 9].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                // Assign borders
                                ws.Cells[finCabecera, 8, finCabecera, 9].Style.Border.Top.Style = ExcelBorderStyle.Hair;
                                ws.Cells[finCabecera, 8, finCabecera, 9].Style.Border.Left.Style = ExcelBorderStyle.Hair;
                                ws.Cells[finCabecera, 8, finCabecera, 9].Style.Border.Right.Style = ExcelBorderStyle.Hair;
                                ws.Cells[finCabecera, 8, finCabecera, 9].Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                                ws.Cells[finCabecera, 8, finCabecera, 9].Style.Font.Bold = true;

                                break;
                            default:
                                ws.Cells[finCabecera, col++].Value = "";
                                break;
                        }

                        ws.Row(finCabecera).Height = 20.25;
                    }

                    finCabecera++;

                    contador = 1;
                    foreach (var column in acta.Cuerpo.Acuerdos)
                    {
                        col = 1;
                        for (int i = 1; i <= totalColumnasDetalleAcuerdos; i++)
                        {
                            ws.Cells[finCabecera, col].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                            ws.Cells[finCabecera, col].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            // Assign borders
                            ws.Cells[finCabecera, col].Style.Border.Top.Style = ExcelBorderStyle.Hair;
                            ws.Cells[finCabecera, col].Style.Border.Left.Style = ExcelBorderStyle.Hair;
                            ws.Cells[finCabecera, col].Style.Border.Right.Style = ExcelBorderStyle.Hair;
                            ws.Cells[finCabecera, col].Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                            switch (i)
                            {
                                case 1:

                                    ws.Cells[finCabecera, col++].Value = contador;
                                    ws.Cells[finCabecera, 1, finCabecera, 1].Style.Border.Top.Style = ExcelBorderStyle.Hair;
                                    ws.Cells[finCabecera, 1, finCabecera, 1].Style.Border.Left.Style = ExcelBorderStyle.Hair;
                                    ws.Cells[finCabecera, 1, finCabecera, 1].Style.Border.Right.Style = ExcelBorderStyle.Hair;
                                    ws.Cells[finCabecera, 1, finCabecera, 1].Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                                    break;
                                case 2:
                                    ws.Cells[finCabecera, col++].Value = column.Acuerdo;
                                    ws.Cells[finCabecera, 2, finCabecera, 5].Merge = true;
                                    ws.Cells[finCabecera, 2, finCabecera, 5].Style.Border.Top.Style = ExcelBorderStyle.Hair;
                                    ws.Cells[finCabecera, 2, finCabecera, 5].Style.Border.Left.Style = ExcelBorderStyle.Hair;
                                    ws.Cells[finCabecera, 2, finCabecera, 5].Style.Border.Right.Style = ExcelBorderStyle.Hair;
                                    ws.Cells[finCabecera, 2, finCabecera, 5].Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                                    ws.Cells[finCabecera, 2, finCabecera, 5].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                                    ws.Cells[finCabecera, 2, finCabecera, 5].Style.WrapText = true;
                                    var caracteres = column.Acuerdo.Count();
                                    if (caracteres <= 70)
                                    {
                                        ws.Row(finCabecera).Height = 20.25;
                                    }
                                    else if (caracteres <= 140)
                                    {
                                        ws.Row(finCabecera).Height = 40;
                                    }
                                    else if (caracteres <= 210)
                                    {
                                        ws.Row(finCabecera).Height = 55;
                                    }
                                    else if (caracteres <= 280)
                                    {
                                        ws.Row(finCabecera).Height = 70;
                                    }
                                    else if (caracteres <= 350)
                                    {
                                        ws.Row(finCabecera).Height = 85;
                                    }
                                    else if (caracteres <= 420)
                                    {
                                        ws.Row(finCabecera).Height = 100;
                                    }
                                    else if (caracteres <= 500)
                                    {
                                        ws.Row(finCabecera).Height = 115;
                                    }
                                    else
                                    {
                                        ws.Row(finCabecera).Height = 130;
                                    }

                                    break;
                                case 3:
                                    ws.Cells[finCabecera, 6].Value = column.Responsable;
                                    ws.Cells[finCabecera, 6, finCabecera, 7].Merge = true;
                                    ws.Cells[finCabecera, 6, finCabecera, 7].Style.Border.Top.Style = ExcelBorderStyle.Hair;
                                    ws.Cells[finCabecera, 6, finCabecera, 7].Style.Border.Left.Style = ExcelBorderStyle.Hair;
                                    ws.Cells[finCabecera, 6, finCabecera, 7].Style.Border.Right.Style = ExcelBorderStyle.Hair;
                                    ws.Cells[finCabecera, 6, finCabecera, 7].Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                                    ws.Cells[finCabecera, 6, finCabecera, 7].Style.WrapText = true;
                                    break;
                                case 4:
                                    ws.Cells[finCabecera, 8].Value = column.Fecha.ToString("yyyy/MM/dd");
                                    ws.Cells[finCabecera, 8, finCabecera, 9].Merge = true;
                                    ws.Cells[finCabecera, 8, finCabecera, 9].Style.Border.Top.Style = ExcelBorderStyle.Hair;
                                    ws.Cells[finCabecera, 8, finCabecera, 9].Style.Border.Left.Style = ExcelBorderStyle.Hair;
                                    ws.Cells[finCabecera, 8, finCabecera, 9].Style.Border.Right.Style = ExcelBorderStyle.Hair;
                                    ws.Cells[finCabecera, 8, finCabecera, 9].Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                                    ws.Cells[finCabecera, 8, finCabecera, 9].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                    break;
                                default:
                                    ws.Cells[finCabecera, col++].Value = "";
                                    break;
                            }

                        }


                        finCabecera++;
                        contador++;

                    }

                    finCabecera += 2;
                }
                #endregion

                //Solo Acta de tipo Cliente
                if (acta.Cuerpo.DetalleCliente.Any() || acta.Cuerpo.DetalleContabilidad.Any())
                {
                    ws.Column(2).Width = 25;
                    ws.Column(3).Width = 15;

                    ws.Column(4).Width = 15;
                    ws.Column(5).Width = 30;

                    ws.Column(6).Width = 10;
                    ws.Column(7).Width = 25;
                    ws.Column(8).Width = 17;
                    ws.Column(9).Width = 20;

                    ws.Protection.IsProtected = true;
                }
                else
                {
                    ws.Column(2).Width = 18;
                    ws.Column(3).Width = 20;
                    ws.Column(6).Width = 15;
                    ws.Column(7).Width = 20;
                    ws.Column(8).Width = 12;
                    ws.Column(9).Width = 15;
                }

                #region Pie de Pagina 
                if (!string.IsNullOrEmpty(acta.PiePagina.AcuerdoConformidad) || acta.Cuerpo.DetalleCliente.Any() || acta.Cuerpo.DetalleContabilidad.Any())
                {

                    var firmas = JsonConvert.DeserializeObject<List<FirmasActaParcial>>(acta.PiePagina.Firmas);

                    ws.Cells[finCabecera, 1].Value = "A  C  U  E  R  D  O   D  E   C  O  N  F  O  R  M  I  D  A  D";

                    ws.Cells[finCabecera, 1].Style.Fill.BackgroundColor.SetColor(colorGrisOscuroEstiloPPM);
                    ws.Cells[finCabecera, 1].Style.Font.Color.SetColor(colorBlancoEstiloPPM);
                    ws.Cells[finCabecera, 1, finCabecera, 9].Merge = true;
                    ws.Cells[finCabecera, 1, finCabecera, 9].Style.HorizontalAlignment = ExcelHorizontalAlignment.CenterContinuous;
                    ws.Cells[finCabecera, 1, finCabecera, 9].Style.Border.Top.Style = ExcelBorderStyle.Hair;
                    ws.Cells[finCabecera, 1, finCabecera, 9].Style.Border.Left.Style = ExcelBorderStyle.Hair;
                    ws.Cells[finCabecera, 1, finCabecera, 9].Style.Border.Right.Style = ExcelBorderStyle.Hair;
                    ws.Cells[finCabecera, 1, finCabecera, 9].Style.Border.Bottom.Style = ExcelBorderStyle.Hair;

                    ws.Row(finCabecera).Height = 20.25;

                    int finBorder = finCabecera + 12;

                    ws.Cells[finCabecera, 1, finBorder, 9].Style.Border.BorderAround(ExcelBorderStyle.Hair);

                    finCabecera += 2;

                    // Solo para cualquier acta que no sea la de CLiente o de Contabilidad
                    if (!acta.Cuerpo.DetalleCliente.Any() && !acta.Cuerpo.DetalleContabilidad.Any())
                    {
                        ws.Cells[finCabecera, 1].Value = acta.PiePagina.AcuerdoConformidad;
                        ws.Cells[finCabecera, 1, finCabecera, 9].Merge = true;
                        ws.Cells[finCabecera, 1, finCabecera, 9].Style.HorizontalAlignment = ExcelHorizontalAlignment.CenterContinuous;
                    }

                    finCabecera += 7;

                    int inicioFilaFirma = finCabecera;
                    int auxiliar = inicioFilaFirma;


                    for (int i = 3; i <= 7; i++)
                    {
                        if (i == 3)
                        {
                            ws.Cells[inicioFilaFirma, i].Value = firmas.ElementAtOrDefault(0).nombre;
                            ws.Cells[inicioFilaFirma, i].Style.Border.Top.Style = ExcelBorderStyle.Hair;
                            ws.Cells[inicioFilaFirma, i].Style.HorizontalAlignment = ExcelHorizontalAlignment.CenterContinuous;
                            inicioFilaFirma++;


                            if (acta.Cuerpo.DetalleCliente.Any() || acta.Cuerpo.DetalleContabilidad.Any())
                            {
                                ws.Cells[inicioFilaFirma, i].Value = firmas.ElementAtOrDefault(0).empresa;
                                ws.Cells[inicioFilaFirma, i].Style.Font.Bold = true;
                                ws.Cells[inicioFilaFirma, i].Style.HorizontalAlignment = ExcelHorizontalAlignment.CenterContinuous;
                                inicioFilaFirma++;
                                ws.Cells[inicioFilaFirma, i].Value = firmas.ElementAtOrDefault(0).fecha;
                                ws.Cells[inicioFilaFirma, i].Style.Font.Bold = true;
                                ws.Cells[inicioFilaFirma, i].Style.HorizontalAlignment = ExcelHorizontalAlignment.CenterContinuous;
                            }
                            else
                            {
                                ws.Cells[inicioFilaFirma, i].Value = firmas.ElementAtOrDefault(0).cargo;
                                ws.Cells[inicioFilaFirma, i].Style.Font.Bold = true;
                                ws.Cells[inicioFilaFirma, i].Style.HorizontalAlignment = ExcelHorizontalAlignment.CenterContinuous;
                                inicioFilaFirma++;
                                ws.Cells[inicioFilaFirma, i].Value = firmas.ElementAtOrDefault(0).empresa;
                                ws.Cells[inicioFilaFirma, i].Style.Font.Bold = true;
                                ws.Cells[inicioFilaFirma, i].Style.HorizontalAlignment = ExcelHorizontalAlignment.CenterContinuous;
                            }
                        }
                        if (i == 7)
                        {
                            ws.Cells[inicioFilaFirma, i].Value = firmas.ElementAtOrDefault(1).usuarioNombre;
                            ws.Cells[inicioFilaFirma, i].Style.Border.Top.Style = ExcelBorderStyle.Hair;
                            ws.Cells[inicioFilaFirma, i].Style.HorizontalAlignment = ExcelHorizontalAlignment.CenterContinuous;
                            inicioFilaFirma++;

                            if (acta.Cuerpo.DetalleCliente.Any() || acta.Cuerpo.DetalleContabilidad.Any())
                            {
                                ws.Cells[inicioFilaFirma, i].Value = firmas.ElementAtOrDefault(1).usuarioEmpresa;
                                ws.Cells[inicioFilaFirma, i].Style.Font.Bold = true;
                                ws.Cells[inicioFilaFirma, i].Style.HorizontalAlignment = ExcelHorizontalAlignment.CenterContinuous;
                                inicioFilaFirma++;
                                ws.Cells[inicioFilaFirma, i].Value = firmas.ElementAtOrDefault(1).Usuariofecha;
                                ws.Cells[inicioFilaFirma, i].Style.Font.Bold = true;
                                ws.Cells[inicioFilaFirma, i].Style.HorizontalAlignment = ExcelHorizontalAlignment.CenterContinuous;
                            }
                            else
                            {
                                ws.Cells[inicioFilaFirma, i].Value = firmas.ElementAtOrDefault(1).usuarioCargo;
                                ws.Cells[inicioFilaFirma, i].Style.Font.Bold = true;
                                ws.Cells[inicioFilaFirma, i].Style.HorizontalAlignment = ExcelHorizontalAlignment.CenterContinuous;
                                inicioFilaFirma++;
                                ws.Cells[inicioFilaFirma, i].Value = firmas.ElementAtOrDefault(1).usuarioEmpresa;
                                ws.Cells[inicioFilaFirma, i].Style.Font.Bold = true;
                                ws.Cells[inicioFilaFirma, i].Style.HorizontalAlignment = ExcelHorizontalAlignment.CenterContinuous;
                            }

                        }
                        inicioFilaFirma = auxiliar;
                        finCabecera++;
                    }

                    if (acta.Cuerpo.DetalleCliente.Any() || acta.Cuerpo.DetalleContabilidad.Any())
                    {
                        ws.Column(8).Width = 17;
                        ws.Column(9).Width = 20;
                    }
                    else
                    {
                        ws.Column(8).Width = 15;
                        ws.Column(9).Width = 20;
                    }
                }
                //FORMATO FUENTE TEXTO DE TODO EL DOCUMENTO
                using (var range = ws.Cells[3, 1, finCabecera, columnaFinalDocumentoExcel])
                {
                    range.Style.Font.Size = 10;
                    range.Style.Font.Name = "Raleway";
                    range.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                }

                #endregion

                string basePath = ConfigurationManager.AppSettings["RepositorioDocumentos"];
                string rutaArchivos = basePath + "\\GESTION_PPM\\ACTAS";

                string codigoCotizacion = !string.IsNullOrEmpty(acta.Cabecera.CodigoCotizacion) ? acta.Cabecera.CodigoCotizacion : "SinCodigo" + id.ToString();

                var anioActual = DateTime.Now.Year.ToString();
                var almacenFisicoTemporal = Auxiliares.CrearCarpetasDirectorio(rutaArchivos, new List<string>() { anioActual, codigoCotizacion });

                var almacenFisicoOfficeToPDF = Auxiliares.CrearCarpetasDirectorio(Server.MapPath("~/OfficeToPDF/"), new List<string>());

                // Get the complete folder path and store the file inside it.    
                string pathExcel = Path.Combine(almacenFisicoTemporal, "ReporteActaExcel" + acta.Cabecera.CodigoActa + ".xlsx");

                //Write the file to the disk
                FileInfo fi = new FileInfo(pathExcel);
                package.SaveAs(fi);

                string pathExe = Path.Combine(almacenFisicoOfficeToPDF, "officetopdf.exe");
                string pathPDF = Path.Combine(almacenFisicoTemporal, "ReporteActaPDF" + acta.Cabecera.CodigoActa + ".pdf");

                object[] args = new object[] { "officetopdf.exe", pathExcel, pathPDF };

                string comando = String.Format("{0} {1} {2}", args);

                string comandoUbicarseEnRaiz = "cd " + Path.GetDirectoryName(pathExe);

                Log.Info(comandoUbicarseEnRaiz);
                Log.Info(comando);

                List<string> comandos = new List<string> { comandoUbicarseEnRaiz, comando }; // "OfficeToPDF.exe reporte.xlsx report.pdf"

                string archivoExe = Server.MapPath("~/OfficeToPDF/OfficeToPDF.exe");

                Auxiliares.EjecutarProcesosCMD(comandos, new List<string> { pathExcel, pathPDF }, archivoExe);

                resultado = true;

                return resultado;
            }
            catch (Exception ex)
            {
                return resultado;
            }
        }

        public ActionResult BuscarArchivoActa(int id)
        {
            try
            {
                var informacionCompleta = ActaEntity.ConsultarActaInformacionCompleta(id);
                ActaCompleta acta = new ActaCompleta(informacionCompleta);

                string basePath = ConfigurationManager.AppSettings["RepositorioDocumentos"];
                string rutaArchivos = basePath + "\\GESTION_PPM\\ACTAS";

                string codigoCotizacion = !string.IsNullOrEmpty(acta.Cabecera.CodigoCotizacion) ? acta.Cabecera.CodigoCotizacion : "SinCodigo" + id.ToString();
                //string codigoCotizacion = !string.IsNullOrEmpty(acta.Cabecera.CodigoCotizacion) ? acta.Cabecera.CodigoCotizacion : "SinCodigo" + Guid.NewGuid().ToString().Substring(0, 8);

                var anioActual = DateTime.Now.Year.ToString();
                var almacenFisicoTemporal = Auxiliares.CrearCarpetasDirectorio(rutaArchivos, new List<string>() { anioActual, codigoCotizacion });

                // Get the complete folder path and store the file inside it.    
                string pathExcel = Path.Combine(almacenFisicoTemporal, "ReporteActaExcel" + acta.Cabecera.CodigoActa + ".xlsx");
                string pathPDF = Path.Combine(almacenFisicoTemporal, "ReporteActaPDF" + acta.Cabecera.CodigoActa + ".pdf");

                bool rutaArchivoExcel = System.IO.File.Exists(pathExcel);
                bool rutaArchivoPDF = System.IO.File.Exists(pathExcel);

                if (!rutaArchivoExcel && !rutaArchivoPDF)
                {
                    return Json(new { Resultado = new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeArchivoNoExiste } }, JsonRequestBehavior.AllowGet);
                }

                List<string> archivosPath = new List<string> { pathExcel, pathPDF };

                return Json(new { Resultado = new RespuestaTransaccion { Estado = true, Respuesta = Mensajes.MensajeTransaccionExitosa }, PathsArchivos = archivosPath }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                return Json(new { Resultado = new RespuestaTransaccion { Estado = false, Respuesta = ex.Message }, PathsArchivos = new List<string> { } }, JsonRequestBehavior.AllowGet);
            }
        }

        public ActionResult ImpresionMasivaActas(string listadoIDs)
        {
            List<string> archivos = new List<string>();
            try
            {
                string basePath = ConfigurationManager.AppSettings["RepositorioDocumentos"];
                string rutaArchivos = basePath + "\\GESTION_PPM\\ACTAS";
                string masivas = "Actas_Masivas_" + DateTime.Now.ToString("yyyy-MM-dd");

                var anioActual = DateTime.Now.Year.ToString();
                var almacenFisicoTemporal = Auxiliares.CrearCarpetasDirectorio(rutaArchivos, new List<string>() { anioActual, masivas });

                // Get the complete folder path and store the file inside it.    
                string pathExcel = Path.Combine(almacenFisicoTemporal, "ReporteActasMasivasExcel" + Guid.NewGuid().ToString().Substring(0, 6) + ".xlsx");

                // Generar Excel
                var package = new ExcelPackage();
                package = ConsolidacionArchivoActasMasivas(listadoIDs);
                //Write the file to the disk
                FileInfo fi = new FileInfo(pathExcel);
                package.SaveAs(fi);

                string pathArchivosConsolidadoPrefacturas = Path.Combine(almacenFisicoTemporal, "ConsolidadoActas_" + Guid.NewGuid().ToString().Substring(0, 9) + ".pdf");

                var ids = !string.IsNullOrEmpty(listadoIDs) ? listadoIDs.Split(',').Select(int.Parse).ToList() : new List<int> { int.Parse(listadoIDs) };
                var pdfActas = GetPathsArchivosActas(ids);

                bool generacionCorrecta = MergeArhivosPDF(pdfActas, pathArchivosConsolidadoPrefacturas);

                List<string> archivosPath = new List<string> { pathExcel, pathArchivosConsolidadoPrefacturas };

                return Json(new { Resultado = new RespuestaTransaccion { Estado = true, Respuesta = Mensajes.MensajeTransaccionExitosa }, PathsArchivos = archivosPath }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                return Json(new { Resultado = new RespuestaTransaccion { Estado = true, Respuesta = Mensajes.MensajeTransaccionExitosa }, PathsArchivos = archivos }, JsonRequestBehavior.AllowGet);

            }

        }
         
        private ExcelPackage ConsolidacionArchivoActasMasivas(string listadoIDs)
        {
            var package = new ExcelPackage();
            try
            {
                var ids = !string.IsNullOrEmpty(listadoIDs) ? listadoIDs.Split(',').Select(int.Parse).ToList() : new List<int> { int.Parse(listadoIDs) };

                foreach (var id in ids)
                {
                    var informacionCompleta = ActaEntity.ConsultarActaInformacionCompleta(id);
                    ActaCompleta acta = new ActaCompleta(informacionCompleta);
                    string actaTitulo = ActaEntity.GetNombreTipo(acta.Cabecera.TipoActaID);

                    //PALETA DE COLORES PPM
                    var colorGrisOscuroEstiloPPM = Color.FromArgb(60, 66, 87);
                    var colorGrisClaroEstiloPPM = Color.FromArgb(240, 240, 240);
                    var colorGrisClaro2EstiloPPM = Color.FromArgb(112, 117, 128);
                    var colorGrisClaro3EstiloPPM = Color.FromArgb(225, 225, 225);
                    var colorBlancoEstiloPPM = Color.FromArgb(255, 255, 255);
                    var colorNegroEstiloPPM = Color.FromArgb(0, 0, 0);

                    #region Cabecera

                    var ws = package.Workbook.Worksheets.Add(acta.Cabecera.CodigoActa);

                    int columnaFinalDocumentoExcel = 9;
                    int columnaInicialDocumentoExcel = 1;

                    ws.PrinterSettings.PaperSize = ePaperSize.A4;//ePaperSize.A3;
                    ws.PrinterSettings.Orientation = acta.Cabecera.CodigoTipoActa == "AECE" || acta.Cabecera.CodigoTipoActa == "AECF" ? eOrientation.Landscape : eOrientation.Portrait;
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

                    System.Drawing.Image img = System.Drawing.Image.FromFile(@pathUbicacion);
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

                    // TITULO ACTA
                    using (var range = ws.Cells[1, 5, 2, 8])
                    {
                        range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        range.Style.Fill.BackgroundColor.SetColor(colorGrisOscuroEstiloPPM);

                        range.Style.Font.Color.SetColor(colorBlancoEstiloPPM);

                        range.Value = actaTitulo;
                        //range.Style.Font.Bold = true;
                        range.Style.Font.Size = 18;
                        range.Style.Font.Name = "Raleway";
                        range.Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                        range.Merge = true;
                    }

                    // CODIGO ACTA
                    using (var range = ws.Cells[1, 9, 2, 9])
                    {
                        range.Style.HorizontalAlignment = ExcelHorizontalAlignment.CenterContinuous;
                        range.Style.Fill.BackgroundColor.SetColor(colorGrisClaroEstiloPPM);

                        range.Style.Font.Color.SetColor(colorGrisOscuroEstiloPPM);

                        range.Value = acta.Cabecera.CodigoActa;//actaTitulo + " " + acta.Cabecera.CodigoActa;
                                                               //range.Style.Font.Bold = true;
                        range.Style.Font.Size = 8;
                        range.Style.Font.Name = "Raleway";
                        range.Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                        range.Merge = true;
                    }

                    //FORMATO FUENTE TEXTO DE TODO EL DOCUMENTO

                    int finCabecera = 4;

                    if (acta.Cabecera.CodigoTipoActa == "AECE" || acta.Cabecera.CodigoTipoActa == "AECF")
                    {
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

                            ws.Cells[range.Start.Row, range.Columns + 1, range.End.Row, range.Columns + 4].Merge = true;

                            ws.Cells[range.Start.Row, range.Columns + 1, range.End.Row, range.Columns + 4].Style.Border.Top.Style = ExcelBorderStyle.Hair;
                            ws.Cells[range.Start.Row, range.Columns + 1, range.End.Row, range.Columns + 4].Style.Border.Left.Style = ExcelBorderStyle.Hair;
                            ws.Cells[range.Start.Row, range.Columns + 1, range.End.Row, range.Columns + 4].Style.Border.Right.Style = ExcelBorderStyle.Hair;
                            ws.Cells[range.Start.Row, range.Columns + 1, range.End.Row, range.Columns + 4].Style.Border.Bottom.Style = ExcelBorderStyle.Hair;

                            ws.Cells[range.Start.Row, range.Columns + 1].Value = acta.Cabecera.NombresElaboradoPor; // Valor campo
                            ws.Cells[range.Start.Row, range.Columns + 1].Style.Indent = 1;

                        }

                        using (var range = ws.Cells[finCabecera, 7, finCabecera, 8])
                        {
                            range.Style.Border.Top.Style = ExcelBorderStyle.Hair;
                            range.Style.Border.Left.Style = ExcelBorderStyle.Hair;
                            range.Style.Border.Right.Style = ExcelBorderStyle.Hair;
                            range.Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                            range.Merge = true;

                            range.Value = "Fecha de Creación:";
                            range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;


                            range.Style.Fill.BackgroundColor.SetColor(colorGrisOscuroEstiloPPM);
                            range.Style.Font.Color.SetColor(colorBlancoEstiloPPM);
                            range.Style.Font.Bold = true;

                            ws.Cells[range.Start.Row, range.End.Column + 1, range.End.Row, columnaFinalDocumentoExcel].Merge = true;

                            ws.Cells[range.Start.Row, range.End.Column + 1, range.End.Row, columnaFinalDocumentoExcel].Style.Border.Top.Style = ExcelBorderStyle.Hair;
                            ws.Cells[range.Start.Row, range.End.Column + 1, range.End.Row, columnaFinalDocumentoExcel].Style.Border.Left.Style = ExcelBorderStyle.Hair;
                            ws.Cells[range.Start.Row, range.End.Column + 1, range.End.Row, columnaFinalDocumentoExcel].Style.Border.Right.Style = ExcelBorderStyle.Hair;
                            ws.Cells[range.Start.Row, range.End.Column + 1, range.End.Row, columnaFinalDocumentoExcel].Style.Border.Bottom.Style = ExcelBorderStyle.Hair;

                            ws.Cells[range.Start.Row, range.End.Column + 1].Value = acta.Cabecera.FechaCreacion.ToString("yyyy/MM/dd"); // Valor campo
                            ws.Cells[range.Start.Row, range.End.Column + 1].Style.Indent = 1;
                        }

                        finCabecera += 2;
                        ws.Row(finCabecera).Height = 20.25;
                    }
                    else
                    {
                        if (acta.Cabecera.CodigoTipoActa == "ARE")
                        {
                            // ws.Row(6).Height = MeasureTextHeight(ws.Cells[6, 3, 6, 6].Value.ToString(), ws.Cells[6, 3, 6, 6].Style.Font, 4);
                            //FECHA INICIO
                            using (var range = ws.Cells[5, 1, 5, 2])
                            {
                                range.Style.Border.Top.Style = ExcelBorderStyle.Hair;
                                range.Style.Border.Left.Style = ExcelBorderStyle.Hair;
                                range.Style.Border.Right.Style = ExcelBorderStyle.Hair;
                                range.Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                                range.Merge = true;

                                range.Value = "Fecha:";

                                range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                                range.Style.Fill.BackgroundColor.SetColor(colorGrisClaro3EstiloPPM);
                                range.Style.Font.Color.SetColor(colorGrisOscuroEstiloPPM);
                                range.Style.Font.Bold = true;

                                ws.Cells[range.Start.Row, range.Columns + 1, range.End.Row, range.Columns + 4].Merge = true;
                                ws.Cells[range.Start.Row, range.Columns + 1].Value = acta.Cabecera.FechaInicio.ToString("yyyy/MM/dd"); // Valor campo
                                ws.Cells[range.Start.Row, range.Columns + 1].Style.Indent = 1;

                            }

                            //Duracion
                            using (var range = ws.Cells[5, 7, 5, 8])
                            {
                                range.Style.Border.Top.Style = ExcelBorderStyle.Hair;
                                range.Style.Border.Left.Style = ExcelBorderStyle.Hair;
                                range.Style.Border.Right.Style = ExcelBorderStyle.Hair;
                                range.Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                                range.Merge = true;

                                range.Value = "Duración:";
                                range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

                                range.Style.Fill.BackgroundColor.SetColor(colorGrisClaro3EstiloPPM);
                                range.Style.Font.Color.SetColor(colorGrisOscuroEstiloPPM);
                                range.Style.Font.Bold = true;

                                ws.Cells[range.Start.Row, range.End.Column + 1, range.End.Row, columnaFinalDocumentoExcel].Merge = true;
                                ws.Cells[range.Start.Row, range.End.Column + 1].Value = ActaEntity.CalcularDuracion(acta.Cabecera.HoraInicio, acta.Cabecera.HoraFin, " " + "hora(s)");
                                ws.Cells[range.Start.Row, range.End.Column + 1].Style.Indent = 1;
                            }

                            //ELABORADO POR
                            using (var range = ws.Cells[6, 1, 6, 2])
                            {
                                range.Style.Border.Top.Style = ExcelBorderStyle.Hair;
                                range.Style.Border.Left.Style = ExcelBorderStyle.Hair;
                                range.Style.Border.Right.Style = ExcelBorderStyle.Hair;
                                range.Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                                range.Merge = true;

                                range.Value = "Lugar: ";
                                range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;


                                range.Style.Fill.BackgroundColor.SetColor(colorGrisOscuroEstiloPPM);
                                range.Style.Font.Color.SetColor(colorBlancoEstiloPPM);
                                range.Style.Font.Bold = true;

                                ws.Cells[range.Start.Row, range.Columns + 1, range.End.Row, range.Columns + 4].Merge = true;
                                ws.Cells[range.Start.Row, range.Columns + 1].Value = acta.Cabecera.Lugar; // Valor campo
                                ws.Cells[range.Start.Row, range.Columns + 1].Style.Indent = 1;
                            }
                        }
                        else
                        {

                            //FECHA INICIO
                            using (var range = ws.Cells[5, 1, 5, 2])
                            {
                                range.Style.Border.Top.Style = ExcelBorderStyle.Hair;
                                range.Style.Border.Left.Style = ExcelBorderStyle.Hair;
                                range.Style.Border.Right.Style = ExcelBorderStyle.Hair;
                                range.Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                                range.Merge = true;

                                range.Value = "Fecha de Creación:";

                                range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;


                                range.Style.Fill.BackgroundColor.SetColor(colorGrisOscuroEstiloPPM);
                                range.Style.Font.Color.SetColor(colorBlancoEstiloPPM);
                                range.Style.Font.Bold = true;

                                ws.Cells[range.Start.Row, range.Columns + 1, range.End.Row, range.Columns + 4].Merge = true;
                                ws.Cells[range.Start.Row, range.Columns + 1].Value = acta.Cabecera.FechaCreacion.ToString("yyyy/MM/dd"); // Valor campo
                                ws.Cells[range.Start.Row, range.Columns + 1].Style.Indent = 1;

                            }

                            //FECHA INICIO
                            using (var range = ws.Cells[5, 7, 5, 8])
                            {
                                range.Style.Border.Top.Style = ExcelBorderStyle.Hair;
                                range.Style.Border.Left.Style = ExcelBorderStyle.Hair;
                                range.Style.Border.Right.Style = ExcelBorderStyle.Hair;
                                range.Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                                range.Merge = true;

                                range.Value = "Fecha Inicio:";
                                range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;


                                range.Style.Fill.BackgroundColor.SetColor(colorGrisOscuroEstiloPPM);
                                range.Style.Font.Color.SetColor(colorBlancoEstiloPPM);
                                range.Style.Font.Bold = true;

                                ws.Cells[range.Start.Row, range.End.Column + 1, range.End.Row, columnaFinalDocumentoExcel].Merge = true;
                                ws.Cells[range.Start.Row, range.End.Column + 1].Value = acta.Cabecera.FechaInicio.ToString("yyyy/MM/dd"); // Valor campo
                                ws.Cells[range.Start.Row, range.End.Column + 1].Style.Indent = 1;
                            }

                            //ELABORADO POR
                            using (var range = ws.Cells[6, 1, 6, 2])
                            {
                                range.Style.Border.Top.Style = ExcelBorderStyle.Hair;
                                range.Style.Border.Left.Style = ExcelBorderStyle.Hair;
                                range.Style.Border.Right.Style = ExcelBorderStyle.Hair;
                                range.Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                                range.Merge = true;

                                range.Value = "Elaborado Por: ";

                                range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;


                                range.Style.Fill.BackgroundColor.SetColor(colorGrisClaro3EstiloPPM);
                                range.Style.Font.Color.SetColor(colorGrisOscuroEstiloPPM);
                                range.Style.Font.Bold = true;

                                ws.Cells[range.Start.Row, range.Columns + 1, range.End.Row, range.Columns + 4].Merge = true;
                                ws.Cells[range.Start.Row, range.Columns + 1].Value = acta.Cabecera.NombresElaboradoPor; // Valor campo
                                ws.Cells[range.Start.Row, range.Columns + 1].Style.Indent = 1;
                            }
                        }

                        int inicioParteCabeceraParte2 = 10;

                        switch (acta.Cabecera.CodigoTipoActa)
                        {
                            case "ARE":

                                //HoraInicio
                                using (var range = ws.Cells[6, 7, 6, 8])
                                {
                                    range.Style.Border.Top.Style = ExcelBorderStyle.Hair;
                                    range.Style.Border.Left.Style = ExcelBorderStyle.Hair;
                                    range.Style.Border.Right.Style = ExcelBorderStyle.Hair;
                                    range.Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                                    range.Merge = true;

                                    range.Value = "Hora de Inicio:";
                                    range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

                                    range.Style.Fill.BackgroundColor.SetColor(colorGrisOscuroEstiloPPM);
                                    range.Style.Font.Color.SetColor(colorBlancoEstiloPPM);
                                    range.Style.Font.Bold = true;

                                    ws.Cells[range.Start.Row, range.End.Column + 1, range.End.Row, columnaFinalDocumentoExcel].Merge = true;
                                    ws.Cells[range.Start.Row, range.End.Column + 1].Value = ActaEntity.CalcularHora(acta.Cabecera.HoraInicio); // Valor campo
                                    ws.Cells[range.Start.Row, range.End.Column + 1].Style.Indent = 1;

                                }

                                //NÚMERO DE REUNIÓN
                                using (var range = ws.Cells[7, 1, 7, 2])
                                {
                                    range.Style.Border.Top.Style = ExcelBorderStyle.Hair;
                                    range.Style.Border.Left.Style = ExcelBorderStyle.Hair;
                                    range.Style.Border.Right.Style = ExcelBorderStyle.Hair;
                                    range.Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                                    range.Merge = true;

                                    range.Value = "N. Reunión: ";
                                    range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

                                    range.Style.Fill.BackgroundColor.SetColor(colorGrisClaro3EstiloPPM);
                                    range.Style.Font.Color.SetColor(colorGrisOscuroEstiloPPM);
                                    range.Style.Font.Bold = true;

                                    ws.Cells[range.Start.Row, range.Columns + 1, range.End.Row, range.Columns + 4].Merge = true;
                                    ws.Cells[range.Start.Row, range.Columns + 1].Value = acta.Cabecera.NumeroReunion;  // Valor campo
                                    ws.Cells[range.Start.Row, range.Columns + 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;

                                    ws.Cells[range.Start.Row, range.Columns + 1].Style.Indent = 1;
                                }

                                //HORA FIN
                                using (var range = ws.Cells[7, 7, 7, 8])
                                {
                                    range.Style.Border.Top.Style = ExcelBorderStyle.Hair;
                                    range.Style.Border.Left.Style = ExcelBorderStyle.Hair;
                                    range.Style.Border.Right.Style = ExcelBorderStyle.Hair;
                                    range.Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                                    range.Merge = true;

                                    range.Value = "Hora Fin:";
                                    range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

                                    range.Style.Fill.BackgroundColor.SetColor(colorGrisClaro3EstiloPPM);
                                    range.Style.Font.Color.SetColor(colorGrisOscuroEstiloPPM);
                                    range.Style.Font.Bold = true;

                                    ws.Cells[range.Start.Row, range.End.Column + 1, range.End.Row, columnaFinalDocumentoExcel].Merge = true;
                                    ws.Cells[range.Start.Row, range.End.Column + 1].Value = ActaEntity.CalcularHora(acta.Cabecera.HoraFin); // Valor campo
                                    ws.Cells[range.Start.Row, range.End.Column + 1].Style.Indent = 1;
                                }


                                using (var range = ws.Cells[5, 3, 5, 8])
                                {
                                    range.Style.Border.Top.Style = ExcelBorderStyle.Hair;
                                    range.Style.Border.Left.Style = ExcelBorderStyle.Hair;
                                    range.Style.Border.Right.Style = ExcelBorderStyle.Hair;
                                    range.Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                                }
                                using (var range = ws.Cells[5, 9, 5, 9])
                                {
                                    range.Style.Border.Top.Style = ExcelBorderStyle.Hair;
                                    range.Style.Border.Left.Style = ExcelBorderStyle.Hair;
                                    range.Style.Border.Right.Style = ExcelBorderStyle.Hair;
                                    range.Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                                }

                                using (var range = ws.Cells[6, 5, 6, 8])
                                {
                                    range.Style.Border.Top.Style = ExcelBorderStyle.Hair;
                                    range.Style.Border.Left.Style = ExcelBorderStyle.Hair;
                                    range.Style.Border.Right.Style = ExcelBorderStyle.Hair;
                                    range.Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                                }

                                using (var range = ws.Cells[6, 9, 6, 9])
                                {
                                    range.Style.Border.Top.Style = ExcelBorderStyle.Hair;
                                    range.Style.Border.Left.Style = ExcelBorderStyle.Hair;
                                    range.Style.Border.Right.Style = ExcelBorderStyle.Hair;
                                    range.Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                                }
                                using (var range = ws.Cells[7, 1, 7, 2])
                                {
                                    range.Style.Border.Top.Style = ExcelBorderStyle.Hair;
                                    range.Style.Border.Left.Style = ExcelBorderStyle.Hair;
                                    range.Style.Border.Right.Style = ExcelBorderStyle.Hair;
                                    range.Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                                }
                                using (var range = ws.Cells[7, 3, 7, 9])
                                {
                                    range.Style.Border.Top.Style = ExcelBorderStyle.Hair;
                                    range.Style.Border.Left.Style = ExcelBorderStyle.Hair;
                                    range.Style.Border.Right.Style = ExcelBorderStyle.Hair;
                                    range.Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                                }
                                using (var range = ws.Cells[7, 9, 7, 9])
                                {
                                    range.Style.Border.Top.Style = ExcelBorderStyle.Hair;
                                    range.Style.Border.Left.Style = ExcelBorderStyle.Hair;
                                    range.Style.Border.Right.Style = ExcelBorderStyle.Hair;
                                    range.Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                                }
                                inicioParteCabeceraParte2 = 9;

                                ws.Row(3).Height = 4.12;
                                ws.Row(4).Height = 4.13;
                                ws.Row(5).Height = 20.25;
                                ws.Row(6).Height = 20.25;
                                ws.Row(7).Height = 20.25;
                                ws.Row(8).Height = 8.25;
                                ws.Row(9).Height = 20.25;
                                ws.Row(10).Height = 20.25;
                                ws.Row(16).Height = 8.25;
                                ws.Column(1).Width = 12;
                                break;

                            case "AIP":

                                using (var range = ws.Cells[6, 7, 6, 8])
                                {
                                    range.Style.Border.Top.Style = ExcelBorderStyle.Hair;
                                    range.Style.Border.Left.Style = ExcelBorderStyle.Hair;
                                    range.Style.Border.Right.Style = ExcelBorderStyle.Hair;
                                    range.Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                                    range.Merge = true;

                                    range.Value = "Fecha Fin:";
                                    range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

                                    range.Style.Fill.BackgroundColor.SetColor(colorGrisClaro3EstiloPPM);
                                    range.Style.Font.Color.SetColor(colorGrisOscuroEstiloPPM);
                                    range.Style.Font.Bold = true;

                                    ws.Cells[range.Start.Row, range.End.Column + 1, range.End.Row, columnaFinalDocumentoExcel].Merge = true;
                                    ws.Cells[range.Start.Row, range.End.Column + 1].Value = acta.Cabecera.FechaFin.ToString("yyyy/MM/dd"); // Valor campo
                                    ws.Cells[range.Start.Row, range.End.Column + 1].Style.Indent = 1;

                                }

                                using (var range = ws.Cells[7, 1, 7, 2])
                                {
                                    range.Style.Border.Top.Style = ExcelBorderStyle.Hair;
                                    range.Style.Border.Left.Style = ExcelBorderStyle.Hair;
                                    range.Style.Border.Right.Style = ExcelBorderStyle.Hair;
                                    range.Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                                    range.Merge = true;

                                    range.Value = "Cargo: ";

                                    range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;


                                    range.Style.Fill.BackgroundColor.SetColor(colorGrisOscuroEstiloPPM);
                                    range.Style.Font.Color.SetColor(colorBlancoEstiloPPM);
                                    range.Style.Font.Bold = true;

                                    ws.Cells[range.Start.Row, range.Columns + 1, range.End.Row, range.Columns + 4].Merge = true;
                                    ws.Cells[range.Start.Row, range.Columns + 1].Value = acta.Cabecera.Cargo;  // Valor campo
                                    ws.Cells[range.Start.Row, range.Columns + 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;

                                    ws.Cells[range.Start.Row, range.Columns + 1].Style.Indent = 1;
                                }

                                //HORA Entrega
                                using (var range = ws.Cells[7, 7, 7, 8])
                                {
                                    range.Style.Border.Top.Style = ExcelBorderStyle.Hair;
                                    range.Style.Border.Left.Style = ExcelBorderStyle.Hair;
                                    range.Style.Border.Right.Style = ExcelBorderStyle.Hair;
                                    range.Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                                    range.Merge = true;

                                    range.Value = "Fecha Entrega:";

                                    range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;


                                    range.Style.Fill.BackgroundColor.SetColor(colorGrisOscuroEstiloPPM);
                                    range.Style.Font.Color.SetColor(colorBlancoEstiloPPM);
                                    range.Style.Font.Bold = true;

                                    ws.Cells[range.Start.Row, range.End.Column + 1, range.End.Row, columnaFinalDocumentoExcel].Merge = true;
                                    ws.Cells[range.Start.Row, range.End.Column + 1].Value = acta.Cabecera.FechaEntrega.Value.ToString("yyyy/MM/dd"); ; ; // Valor campo

                                    ws.Cells[range.Start.Row, range.End.Column + 1].Style.Indent = 1;
                                }

                                using (var range = ws.Cells[5, 3, 5, 8])
                                {
                                    range.Style.Border.Top.Style = ExcelBorderStyle.Hair;
                                    range.Style.Border.Left.Style = ExcelBorderStyle.Hair;
                                    range.Style.Border.Right.Style = ExcelBorderStyle.Hair;
                                    range.Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                                }

                                using (var range = ws.Cells[5, 1, 5, 2])
                                {
                                    range.Style.Border.Top.Style = ExcelBorderStyle.Hair;
                                    range.Style.Border.Left.Style = ExcelBorderStyle.Hair;
                                    range.Style.Border.Right.Style = ExcelBorderStyle.Hair;
                                    range.Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                                }

                                using (var range = ws.Cells[6, 3, 6, 8])
                                {
                                    range.Style.Border.Top.Style = ExcelBorderStyle.Hair;
                                    range.Style.Border.Left.Style = ExcelBorderStyle.Hair;
                                    range.Style.Border.Right.Style = ExcelBorderStyle.Hair;
                                    range.Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                                }
                                using (var range = ws.Cells[6, 9, 6, 9])
                                {
                                    range.Style.Border.Top.Style = ExcelBorderStyle.Hair;
                                    range.Style.Border.Left.Style = ExcelBorderStyle.Hair;
                                    range.Style.Border.Right.Style = ExcelBorderStyle.Hair;
                                    range.Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                                }

                                using (var range = ws.Cells[5, 9, 5, 9])
                                {
                                    range.Style.Border.Top.Style = ExcelBorderStyle.Hair;
                                    range.Style.Border.Left.Style = ExcelBorderStyle.Hair;
                                    range.Style.Border.Right.Style = ExcelBorderStyle.Hair;
                                    range.Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                                }
                                using (var range = ws.Cells[7, 1, 7, 2])
                                {
                                    range.Style.Border.Top.Style = ExcelBorderStyle.Hair;
                                    range.Style.Border.Left.Style = ExcelBorderStyle.Hair;
                                    range.Style.Border.Right.Style = ExcelBorderStyle.Hair;
                                    range.Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                                }
                                using (var range = ws.Cells[7, 3, 7, 9])
                                {
                                    range.Style.Border.Top.Style = ExcelBorderStyle.Hair;
                                    range.Style.Border.Left.Style = ExcelBorderStyle.Hair;
                                    range.Style.Border.Right.Style = ExcelBorderStyle.Hair;
                                    range.Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                                }
                                using (var range = ws.Cells[7, 9, 7, 9])
                                {
                                    range.Style.Border.Top.Style = ExcelBorderStyle.Hair;
                                    range.Style.Border.Left.Style = ExcelBorderStyle.Hair;
                                    range.Style.Border.Right.Style = ExcelBorderStyle.Hair;
                                    range.Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                                }

                                inicioParteCabeceraParte2 = 9;

                                ws.Row(3).Height = 4.12;
                                ws.Row(4).Height = 4.13;
                                ws.Row(5).Height = 20.25;
                                ws.Row(6).Height = 20.25;
                                ws.Row(7).Height = 20.25;
                                ws.Row(8).Height = 8.25;
                                ws.Row(9).Height = 20.25;
                                ws.Row(10).Height = 20.25;
                                ws.Row(11).Height = 44.25;
                                ws.Row(12).Height = 44.25;
                                ws.Row(13).Height = 8.25;
                                ws.Row(15).Height = 8.25;
                                ws.Row(17).Height = 8.25;



                                break;
                            case "ACP":

                                //FECHA FIN
                                using (var range = ws.Cells[6, 7, 6, 8])
                                {
                                    range.Style.Border.Top.Style = ExcelBorderStyle.Hair;
                                    range.Style.Border.Left.Style = ExcelBorderStyle.Hair;
                                    range.Style.Border.Right.Style = ExcelBorderStyle.Hair;
                                    range.Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                                    range.Merge = true;

                                    range.Value = "Fecha Fin:";

                                    range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

                                    range.Style.Fill.BackgroundColor.SetColor(colorGrisClaro3EstiloPPM);
                                    range.Style.Font.Color.SetColor(colorGrisOscuroEstiloPPM);
                                    range.Style.Font.Bold = true;

                                    ws.Cells[range.Start.Row, range.End.Column + 1, range.End.Row, columnaFinalDocumentoExcel].Merge = true;
                                    ws.Cells[range.Start.Row, range.End.Column + 1].Value = acta.Cabecera.FechaFin.ToString("yyyy/MM/dd"); // Valor campo

                                    ws.Cells[range.Start.Row, range.End.Column + 1].Style.Indent = 1;
                                }
                                using (var range = ws.Cells[7, 1, 7, 2])
                                {
                                    range.Style.Border.Top.Style = ExcelBorderStyle.Hair;
                                    range.Style.Border.Left.Style = ExcelBorderStyle.Hair;
                                    range.Style.Border.Right.Style = ExcelBorderStyle.Hair;
                                    range.Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                                    range.Merge = true;

                                    range.Value = "Cargo: ";
                                    range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

                                    range.Style.Fill.BackgroundColor.SetColor(colorGrisOscuroEstiloPPM);
                                    range.Style.Font.Color.SetColor(colorBlancoEstiloPPM);
                                    range.Style.Font.Bold = true;

                                    ws.Cells[range.Start.Row, range.Columns + 1, range.End.Row, range.Columns + 4].Merge = true;
                                    ws.Cells[range.Start.Row, range.Columns + 1].Value = acta.Cabecera.Cargo;  // Valor campo
                                    ws.Cells[range.Start.Row, range.Columns + 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;

                                    ws.Cells[range.Start.Row, range.Columns + 1].Style.Indent = 1;
                                }
                                ////HORA ENTREGA
                                using (var range = ws.Cells[7, 7, 7, 8])
                                {
                                    range.Style.Border.Top.Style = ExcelBorderStyle.Hair;
                                    range.Style.Border.Left.Style = ExcelBorderStyle.Hair;
                                    range.Style.Border.Right.Style = ExcelBorderStyle.Hair;
                                    range.Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                                    range.Merge = true;
                                    range.Value = "Fecha Entrega:";
                                    range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                                    range.Style.Fill.BackgroundColor.SetColor(colorGrisOscuroEstiloPPM);
                                    range.Style.Font.Color.SetColor(colorBlancoEstiloPPM);
                                    range.Style.Font.Bold = true;

                                    ws.Cells[range.Start.Row, range.End.Column + 1, range.End.Row, columnaFinalDocumentoExcel].Merge = true;
                                    ws.Cells[range.Start.Row, range.End.Column + 1].Value = acta.Cabecera.FechaEntrega.Value.ToString("yyyy/MM/dd"); // Valor campo
                                    ws.Cells[range.Start.Row, range.End.Column + 1].Style.Indent = 1;
                                }


                                using (var range = ws.Cells[5, 3, 5, 8])
                                {
                                    range.Style.Border.Top.Style = ExcelBorderStyle.Hair;
                                    range.Style.Border.Left.Style = ExcelBorderStyle.Hair;
                                    range.Style.Border.Right.Style = ExcelBorderStyle.Hair;
                                    range.Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                                }

                                using (var range = ws.Cells[5, 1, 5, 2])
                                {
                                    range.Style.Border.Top.Style = ExcelBorderStyle.Hair;
                                    range.Style.Border.Left.Style = ExcelBorderStyle.Hair;
                                    range.Style.Border.Right.Style = ExcelBorderStyle.Hair;
                                    range.Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                                }

                                using (var range = ws.Cells[6, 3, 6, 8])
                                {
                                    range.Style.Border.Top.Style = ExcelBorderStyle.Hair;
                                    range.Style.Border.Left.Style = ExcelBorderStyle.Hair;
                                    range.Style.Border.Right.Style = ExcelBorderStyle.Hair;
                                    range.Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                                }
                                using (var range = ws.Cells[6, 9, 6, 9])
                                {
                                    range.Style.Border.Top.Style = ExcelBorderStyle.Hair;
                                    range.Style.Border.Left.Style = ExcelBorderStyle.Hair;
                                    range.Style.Border.Right.Style = ExcelBorderStyle.Hair;
                                    range.Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                                }

                                using (var range = ws.Cells[5, 9, 5, 9])
                                {
                                    range.Style.Border.Top.Style = ExcelBorderStyle.Hair;
                                    range.Style.Border.Left.Style = ExcelBorderStyle.Hair;
                                    range.Style.Border.Right.Style = ExcelBorderStyle.Hair;
                                    range.Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                                }
                                using (var range = ws.Cells[7, 1, 7, 2])
                                {
                                    range.Style.Border.Top.Style = ExcelBorderStyle.Hair;
                                    range.Style.Border.Left.Style = ExcelBorderStyle.Hair;
                                    range.Style.Border.Right.Style = ExcelBorderStyle.Hair;
                                    range.Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                                }
                                using (var range = ws.Cells[7, 3, 7, 9])
                                {
                                    range.Style.Border.Top.Style = ExcelBorderStyle.Hair;
                                    range.Style.Border.Left.Style = ExcelBorderStyle.Hair;
                                    range.Style.Border.Right.Style = ExcelBorderStyle.Hair;
                                    range.Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                                }
                                using (var range = ws.Cells[7, 9, 7, 9])
                                {
                                    range.Style.Border.Top.Style = ExcelBorderStyle.Hair;
                                    range.Style.Border.Left.Style = ExcelBorderStyle.Hair;
                                    range.Style.Border.Right.Style = ExcelBorderStyle.Hair;
                                    range.Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                                }

                                inicioParteCabeceraParte2 = 9;

                                ws.Row(3).Height = 4.12;
                                ws.Row(4).Height = 4.13;
                                ws.Row(5).Height = 20.25;
                                ws.Row(6).Height = 20.25;
                                ws.Row(7).Height = 20.25;
                                ws.Row(8).Height = 8.25;
                                ws.Row(9).Height = 20.25;
                                ws.Row(10).Height = 20.25;
                                ws.Row(11).Height = 44.25;
                                ws.Row(12).Height = 44.25;
                                ws.Row(13).Height = 8.25;
                                ws.Row(15).Height = 8.25;
                                ws.Row(17).Height = 8.25;

                                break;
                            default:
                                break;
                        }

                        //CodigoCotizacion
                        using (var range = ws.Cells[inicioParteCabeceraParte2, 1, inicioParteCabeceraParte2, 2])
                        {
                            range.Style.Border.Top.Style = ExcelBorderStyle.Hair;
                            range.Style.Border.Left.Style = ExcelBorderStyle.Hair;
                            range.Style.Border.Right.Style = ExcelBorderStyle.Hair;
                            range.Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                            range.Merge = true;

                            range.Value = "Código de Cotización: ";
                            range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

                            range.Style.Fill.BackgroundColor.SetColor(colorGrisClaro3EstiloPPM);
                            range.Style.Font.Color.SetColor(colorGrisOscuroEstiloPPM);
                            range.Style.Font.Bold = true;

                            ws.Cells[range.Start.Row, range.Columns + 1, range.End.Row, columnaFinalDocumentoExcel].Merge = true;
                            ws.Cells[range.Start.Row, range.Columns + 1].Value = acta.Cabecera.CodigoCotizacion;

                            ws.Cells[range.Start.Row, range.Columns + 1].Style.Indent = 1;

                            ws.Cells[range.Start.Row, range.Columns + 1, range.End.Row, columnaFinalDocumentoExcel].Style.Border.Top.Style = ExcelBorderStyle.Hair;
                            ws.Cells[range.Start.Row, range.Columns + 1, range.End.Row, columnaFinalDocumentoExcel].Style.Border.Left.Style = ExcelBorderStyle.Hair;
                            ws.Cells[range.Start.Row, range.Columns + 1, range.End.Row, columnaFinalDocumentoExcel].Style.Border.Right.Style = ExcelBorderStyle.Hair;
                            ws.Cells[range.Start.Row, range.Columns + 1, range.End.Row, columnaFinalDocumentoExcel].Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                        }

                        inicioParteCabeceraParte2++;

                        using (var range = ws.Cells[inicioParteCabeceraParte2, 1, inicioParteCabeceraParte2, 2])
                        {
                            range.Style.Border.Top.Style = ExcelBorderStyle.Hair;
                            range.Style.Border.Left.Style = ExcelBorderStyle.Hair;
                            range.Style.Border.Right.Style = ExcelBorderStyle.Hair;
                            range.Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                            range.Merge = true;

                            range.Value = "Cliente: ";
                            range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

                            range.Style.Fill.BackgroundColor.SetColor(colorGrisClaro3EstiloPPM);
                            range.Style.Font.Color.SetColor(colorGrisOscuroEstiloPPM);
                            range.Style.Font.Bold = true;

                            ws.Cells[range.Start.Row, range.Columns + 1, range.End.Row, columnaFinalDocumentoExcel].Merge = true;
                            ws.Cells[range.Start.Row, range.Columns + 1].Value = acta.Cabecera.Cliente;
                            ws.Cells[range.Start.Row, range.Columns + 1].Style.Indent = 1;

                            ws.Cells[range.Start.Row, range.Columns + 1, range.End.Row, columnaFinalDocumentoExcel].Style.Border.Top.Style = ExcelBorderStyle.Hair;
                            ws.Cells[range.Start.Row, range.Columns + 1, range.End.Row, columnaFinalDocumentoExcel].Style.Border.Left.Style = ExcelBorderStyle.Hair;
                            ws.Cells[range.Start.Row, range.Columns + 1, range.End.Row, columnaFinalDocumentoExcel].Style.Border.Right.Style = ExcelBorderStyle.Hair;
                            ws.Cells[range.Start.Row, range.Columns + 1, range.End.Row, columnaFinalDocumentoExcel].Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                        }

                        inicioParteCabeceraParte2++;

                        using (var range = ws.Cells[inicioParteCabeceraParte2, 1, inicioParteCabeceraParte2, 2])
                        {
                            range.Style.Border.Top.Style = ExcelBorderStyle.Hair;
                            range.Style.Border.Left.Style = ExcelBorderStyle.Hair;
                            range.Style.Border.Right.Style = ExcelBorderStyle.Hair;
                            range.Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                            range.Merge = true;

                            range.Value = "Nombre del Proyecto:";
                            range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

                            range.Style.Fill.BackgroundColor.SetColor(colorGrisClaro3EstiloPPM);
                            range.Style.Font.Color.SetColor(colorGrisOscuroEstiloPPM);
                            range.Style.Font.Bold = true;

                            ws.Cells[range.Start.Row, range.Columns + 1, range.End.Row, columnaFinalDocumentoExcel].Merge = true;
                            ws.Cells[range.Start.Row, range.Columns + 1].Value = acta.Cabecera.NombreProyecto;
                            ws.Cells[range.Start.Row, range.Columns + 1].Style.Indent = 1;

                            ws.Cells[range.Start.Row, range.Columns + 1, range.End.Row, columnaFinalDocumentoExcel].Style.Border.Top.Style = ExcelBorderStyle.Hair;
                            ws.Cells[range.Start.Row, range.Columns + 1, range.End.Row, columnaFinalDocumentoExcel].Style.Border.Left.Style = ExcelBorderStyle.Hair;
                            ws.Cells[range.Start.Row, range.Columns + 1, range.End.Row, columnaFinalDocumentoExcel].Style.Border.Right.Style = ExcelBorderStyle.Hair;
                            ws.Cells[range.Start.Row, range.Columns + 1, range.End.Row, columnaFinalDocumentoExcel].Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                            ws.Cells[range.Start.Row, range.Columns + 1, range.End.Row, columnaFinalDocumentoExcel].Style.WrapText = true;
                            var caracteres = acta.Cabecera.NombreProyecto.Count();

                            if (caracteres <= 104)
                            {
                                ws.Row(inicioParteCabeceraParte2).Height = 20.25;
                            }
                            else if (caracteres <= 208)
                            {
                                ws.Row(inicioParteCabeceraParte2).Height = 35.50;
                            }
                            else if (caracteres <= 312)
                            {
                                ws.Row(inicioParteCabeceraParte2).Height = 50.75;
                            }
                            else
                            {
                                ws.Row(inicioParteCabeceraParte2).Height = 66;
                            }
                        }

                        inicioParteCabeceraParte2++;

                        using (var range = ws.Cells[inicioParteCabeceraParte2, 1, inicioParteCabeceraParte2, 2])
                        {
                            range.Style.Border.Top.Style = ExcelBorderStyle.Hair;
                            range.Style.Border.Left.Style = ExcelBorderStyle.Hair;
                            range.Style.Border.Right.Style = ExcelBorderStyle.Hair;
                            range.Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                            range.Merge = true;

                            range.Value = "Descripción del Proyecto:";
                            range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

                            range.Style.Fill.BackgroundColor.SetColor(colorGrisClaro3EstiloPPM);
                            range.Style.Font.Color.SetColor(colorGrisOscuroEstiloPPM);
                            range.Style.Font.Bold = true;

                            ws.Cells[range.Start.Row, range.Columns + 1, range.End.Row, columnaFinalDocumentoExcel].Merge = true;
                            ws.Cells[range.Start.Row, range.Columns + 1].Value = acta.Cabecera.DescripcionProyecto;
                            ws.Cells[range.Start.Row, range.Columns + 1].Style.Indent = 1;

                            ws.Cells[range.Start.Row, range.Columns + 1, range.End.Row, columnaFinalDocumentoExcel].Style.Border.Top.Style = ExcelBorderStyle.Hair;
                            ws.Cells[range.Start.Row, range.Columns + 1, range.End.Row, columnaFinalDocumentoExcel].Style.Border.Left.Style = ExcelBorderStyle.Hair;
                            ws.Cells[range.Start.Row, range.Columns + 1, range.End.Row, columnaFinalDocumentoExcel].Style.Border.Right.Style = ExcelBorderStyle.Hair;
                            ws.Cells[range.Start.Row, range.Columns + 1, range.End.Row, columnaFinalDocumentoExcel].Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                            ws.Cells[range.Start.Row, range.Columns + 1, range.End.Row, columnaFinalDocumentoExcel].Style.WrapText = true;
                            var caracteres = acta.Cabecera.DescripcionProyecto.Count();
                            if (caracteres <= 104)
                            {
                                ws.Row(inicioParteCabeceraParte2).Height = 20.25;
                            }
                            else if (caracteres <= 208)
                            {
                                ws.Row(inicioParteCabeceraParte2).Height = 35.50;
                            }
                            else if (caracteres <= 312)
                            {
                                ws.Row(inicioParteCabeceraParte2).Height = 50.75;
                            }
                            else if (caracteres <= 416)
                            {
                                ws.Row(inicioParteCabeceraParte2).Height = 66;
                            }
                            else if (caracteres <= 500)
                            {
                                ws.Row(inicioParteCabeceraParte2).Height = 81.25;
                            }
                            else
                            {
                                ws.Row(inicioParteCabeceraParte2).Height = 96.50;
                            }


                        }

                        inicioParteCabeceraParte2++;

                        // Acta de cierre de proyecto o inicio de proyecto
                        if (acta.Cabecera.CodigoTipoActa == "AIP" || acta.Cabecera.CodigoTipoActa == "ACP")
                        {
                            inicioParteCabeceraParte2 += 1;
                            using (var range = ws.Cells[inicioParteCabeceraParte2, 1, inicioParteCabeceraParte2, 2])
                            {
                                range.Style.Border.Top.Style = ExcelBorderStyle.Hair;
                                range.Style.Border.Left.Style = ExcelBorderStyle.Hair;
                                range.Style.Border.Right.Style = ExcelBorderStyle.Hair;
                                range.Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                                range.Merge = true;

                                range.Value = "Referencia del Cliente:";
                                range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

                                range.Style.Fill.BackgroundColor.SetColor(colorGrisOscuroEstiloPPM);
                                range.Style.Font.Color.SetColor(colorBlancoEstiloPPM);
                                range.Style.Font.Bold = true;

                                ws.Cells[range.Start.Row, range.Columns + 1, range.End.Row, columnaFinalDocumentoExcel].Merge = true;
                                ws.Cells[range.Start.Row, range.Columns + 1].Value = acta.Cabecera.ReferenciaCliente;
                                ws.Cells[range.Start.Row, range.Columns + 1].Style.Indent = 1;

                                ws.Cells[range.Start.Row, range.Columns + 1, range.End.Row, columnaFinalDocumentoExcel].Style.Border.Top.Style = ExcelBorderStyle.Hair;
                                ws.Cells[range.Start.Row, range.Columns + 1, range.End.Row, columnaFinalDocumentoExcel].Style.Border.Left.Style = ExcelBorderStyle.Hair;
                                ws.Cells[range.Start.Row, range.Columns + 1, range.End.Row, columnaFinalDocumentoExcel].Style.Border.Right.Style = ExcelBorderStyle.Hair;
                                ws.Cells[range.Start.Row, range.Columns + 1, range.End.Row, columnaFinalDocumentoExcel].Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                                ws.Cells[range.Start.Row, range.Columns + 1, range.End.Row, columnaFinalDocumentoExcel].Style.WrapText = true;

                                var caracteres = (acta.Cabecera.ReferenciaCliente ?? "").Count();

                                if (caracteres <= 104)
                                {
                                    ws.Row(inicioParteCabeceraParte2).Height = 20.25;
                                }
                                else if (caracteres <= 208)
                                {
                                    ws.Row(inicioParteCabeceraParte2).Height = 35.50;
                                }

                                else
                                {
                                    ws.Row(inicioParteCabeceraParte2).Height = 50.75;
                                }
                            }


                            inicioParteCabeceraParte2 += 2;
                            using (var range = ws.Cells[inicioParteCabeceraParte2, 1, inicioParteCabeceraParte2, 2])
                            {

                                range.Style.Border.Top.Style = ExcelBorderStyle.Hair;
                                range.Style.Border.Left.Style = ExcelBorderStyle.Hair;
                                range.Style.Border.Right.Style = ExcelBorderStyle.Hair;
                                range.Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                                range.Merge = true;

                                range.Value = "Objetivo o Alcance:";
                                range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

                                range.Style.Fill.BackgroundColor.SetColor(colorGrisOscuroEstiloPPM);
                                range.Style.Font.Color.SetColor(colorBlancoEstiloPPM);
                                range.Style.Font.Bold = true;

                                ws.Cells[range.Start.Row, range.Columns + 1, range.End.Row, columnaFinalDocumentoExcel].Merge = true;
                                ws.Cells[range.Start.Row, range.Columns + 1].Value = acta.Cabecera.AlcanceObjetivo;
                                ws.Cells[range.Start.Row, range.Columns + 1].Style.Indent = 1;

                                ws.Cells[range.Start.Row, range.Columns + 1, range.End.Row, columnaFinalDocumentoExcel].Style.Border.Top.Style = ExcelBorderStyle.Hair;
                                ws.Cells[range.Start.Row, range.Columns + 1, range.End.Row, columnaFinalDocumentoExcel].Style.Border.Left.Style = ExcelBorderStyle.Hair;
                                ws.Cells[range.Start.Row, range.Columns + 1, range.End.Row, columnaFinalDocumentoExcel].Style.Border.Right.Style = ExcelBorderStyle.Hair;
                                ws.Cells[range.Start.Row, range.Columns + 1, range.End.Row, columnaFinalDocumentoExcel].Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                                //  ws.Row(inicioParteCabeceraParte2).Height = 33.75;
                                ws.Cells[range.Start.Row, range.Columns + 1, range.End.Row, columnaFinalDocumentoExcel].Style.WrapText = true;

                                var caracteres = acta.Cabecera.AlcanceObjetivo.Count();

                                if (caracteres <= 104)
                                {
                                    ws.Row(inicioParteCabeceraParte2).Height = 20.25;
                                }
                                else if (caracteres <= 208)
                                {
                                    ws.Row(inicioParteCabeceraParte2).Height = 35.50;
                                }
                                else if (caracteres <= 312)
                                {
                                    ws.Row(inicioParteCabeceraParte2).Height = 50.75;
                                }
                                else if (caracteres <= 416)
                                {
                                    ws.Row(inicioParteCabeceraParte2).Height = 66;
                                }
                                else if (caracteres <= 500)
                                {
                                    ws.Row(inicioParteCabeceraParte2).Height = 81.25;
                                }
                                else
                                {
                                    ws.Row(inicioParteCabeceraParte2).Height = 96.50;
                                }
                            }
                        }
                        else
                        {
                            ws.Row(inicioParteCabeceraParte2).Height = 8.25;
                            inicioParteCabeceraParte2 += 1;
                            using (var range = ws.Cells[inicioParteCabeceraParte2, 1, inicioParteCabeceraParte2, 2])
                            {
                                range.Style.Border.Top.Style = ExcelBorderStyle.Hair;
                                range.Style.Border.Left.Style = ExcelBorderStyle.Hair;
                                range.Style.Border.Right.Style = ExcelBorderStyle.Hair;
                                range.Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                                range.Merge = true;

                                range.Value = "Facilitador o Moderador:";
                                range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

                                range.Style.Fill.BackgroundColor.SetColor(colorGrisOscuroEstiloPPM);
                                range.Style.Font.Color.SetColor(colorBlancoEstiloPPM);
                                range.Style.Font.Bold = true;

                                ws.Cells[range.Start.Row, range.Columns + 1, range.End.Row, columnaFinalDocumentoExcel].Merge = true;
                                ws.Cells[range.Start.Row, range.Columns + 1].Value = acta.Cabecera.FacilitadorModerador;
                                ws.Cells[range.Start.Row, range.Columns + 1].Style.Indent = 1;

                                ws.Cells[range.Start.Row, range.Columns + 1, range.End.Row, columnaFinalDocumentoExcel].Style.Border.Top.Style = ExcelBorderStyle.Hair;
                                ws.Cells[range.Start.Row, range.Columns + 1, range.End.Row, columnaFinalDocumentoExcel].Style.Border.Left.Style = ExcelBorderStyle.Hair;
                                ws.Cells[range.Start.Row, range.Columns + 1, range.End.Row, columnaFinalDocumentoExcel].Style.Border.Right.Style = ExcelBorderStyle.Hair;
                                ws.Cells[range.Start.Row, range.Columns + 1, range.End.Row, columnaFinalDocumentoExcel].Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                                ws.Cells[range.Start.Row, range.Columns + 1, range.End.Row, columnaFinalDocumentoExcel].Style.WrapText = true;
                            }
                            ws.Row(inicioParteCabeceraParte2).Height = 20.25;

                            inicioParteCabeceraParte2 += 1;

                            using (var range = ws.Cells[inicioParteCabeceraParte2, 1, inicioParteCabeceraParte2, 2])
                            {

                                range.Style.Border.Top.Style = ExcelBorderStyle.Hair;
                                range.Style.Border.Left.Style = ExcelBorderStyle.Hair;
                                range.Style.Border.Right.Style = ExcelBorderStyle.Hair;
                                range.Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                                range.Merge = true;

                                range.Value = "Objetivo o Alcance:";
                                range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

                                range.Style.Fill.BackgroundColor.SetColor(colorGrisOscuroEstiloPPM);
                                range.Style.Font.Color.SetColor(colorBlancoEstiloPPM);
                                range.Style.Font.Bold = true;

                                ws.Cells[range.Start.Row, range.Columns + 1, range.End.Row, columnaFinalDocumentoExcel].Merge = true;
                                ws.Cells[range.Start.Row, range.Columns + 1].Value = acta.Cabecera.AlcanceObjetivo;
                                ws.Cells[range.Start.Row, range.Columns + 1].Style.Indent = 1;

                                ws.Cells[range.Start.Row, range.Columns + 1, range.End.Row, columnaFinalDocumentoExcel].Style.Border.Top.Style = ExcelBorderStyle.Hair;
                                ws.Cells[range.Start.Row, range.Columns + 1, range.End.Row, columnaFinalDocumentoExcel].Style.Border.Left.Style = ExcelBorderStyle.Hair;
                                ws.Cells[range.Start.Row, range.Columns + 1, range.End.Row, columnaFinalDocumentoExcel].Style.Border.Right.Style = ExcelBorderStyle.Hair;
                                ws.Cells[range.Start.Row, range.Columns + 1, range.End.Row, columnaFinalDocumentoExcel].Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                                ws.Row(inicioParteCabeceraParte2).Height = 33.75;
                                ws.Cells[range.Start.Row, range.Columns + 1, range.End.Row, columnaFinalDocumentoExcel].Style.WrapText = true;


                                var caracteres = acta.Cabecera.AlcanceObjetivo.Count();

                                if (caracteres <= 104)
                                {
                                    ws.Row(inicioParteCabeceraParte2).Height = 20.25;
                                }
                                else if (caracteres <= 208)
                                {
                                    ws.Row(inicioParteCabeceraParte2).Height = 35.50;
                                }
                                else if (caracteres <= 312)
                                {
                                    ws.Row(inicioParteCabeceraParte2).Height = 50.75;
                                }
                                else if (caracteres <= 416)
                                {
                                    ws.Row(inicioParteCabeceraParte2).Height = 66;
                                }
                                else if (caracteres <= 500)
                                {
                                    ws.Row(inicioParteCabeceraParte2).Height = 81.25;
                                }
                                else
                                {
                                    ws.Row(inicioParteCabeceraParte2).Height = 96.50;
                                }
                            }
                        }

                        inicioParteCabeceraParte2 += 2;
                        using (ExcelRange row = ws.Cells["A5:A15"])
                        {
                            row.Style.Font.Bold = true;
                        }

                        //int finCabecera;

                        if (acta.Cabecera.CodigoTipoActa == "AIP" || acta.Cabecera.CodigoTipoActa == "ACP")
                        {

                            finCabecera = 18;
                            ws.Cells[finCabecera, 1].Value = "R  E  S  P  O  N  S  A  B  L  E  S   D  E  L   C  L  I  E  N  T  E";
                            ws.Cells[finCabecera, 1].Style.Fill.BackgroundColor.SetColor(colorGrisOscuroEstiloPPM);
                            ws.Cells[finCabecera, 1].Style.Font.Color.SetColor(colorBlancoEstiloPPM);
                            ws.Cells[finCabecera, 1, finCabecera, 9].Merge = true;
                            ws.Cells[finCabecera, 1, finCabecera, 9].Style.HorizontalAlignment = ExcelHorizontalAlignment.CenterContinuous;
                            ws.Cells[finCabecera, 1, finCabecera, 9].Style.Border.Top.Style = ExcelBorderStyle.Hair;
                            ws.Cells[finCabecera, 1, finCabecera, 9].Style.Border.Left.Style = ExcelBorderStyle.Hair;
                            ws.Cells[finCabecera, 1, finCabecera, 9].Style.Border.Right.Style = ExcelBorderStyle.Hair;
                            ws.Cells[finCabecera, 1, finCabecera, 9].Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                            ws.Row(finCabecera).Height = 20.25;

                        }
                        else
                        {
                            finCabecera = 16;

                        }

                        finCabecera++;
                        ws.Row(finCabecera).Height = 20.25;

                    }

                    #endregion

                    Int32 col = 1;
                    int contador = 1;

                    #region Detalle Acta Cliente

                    if (acta.Cuerpo.DetalleCliente.Any())
                    {
                        var columnas = new List<string> { "N°", "CÓDIGO DE COTIZACIÓN", "CLIENTE", "INTERMEDIARIO", "DETALLE", "VALOR", "EJECUTIVO", "N. DOCUMENTO", "OBSERVACIONES" };

                        var i = 1;
                        foreach (var item in columnas)
                        {

                            ws.Cells[finCabecera, i].Value = item;
                            ws.Cells[finCabecera, i].Style.Font.Bold = true;

                            ws.Cells[finCabecera, i].Style.Border.Top.Style = ExcelBorderStyle.Hair;
                            ws.Cells[finCabecera, i].Style.Border.Left.Style = ExcelBorderStyle.Hair;
                            ws.Cells[finCabecera, i].Style.Border.Right.Style = ExcelBorderStyle.Hair;
                            ws.Cells[finCabecera, i].Style.Border.Bottom.Style = ExcelBorderStyle.Hair;

                            //worksheet.Cells[1, i].Style.Fill.BackgroundColor. .SetColor(Color.FromArgb(23, 55, 93));
                            i++;
                        }
                        finCabecera++;

                        int numeracion = 1;
                        foreach (var item in acta.Cuerpo.DetalleCliente)
                        {
                            var detalle = CotizacionEntity.ConsultarPrefacturaSAFI(item.id_facturacion_safi);
                            for (int j = 1; j <= 9; j++)
                            {
                                ws.Cells[finCabecera, j].Style.Border.Top.Style = ExcelBorderStyle.Hair;
                                ws.Cells[finCabecera, j].Style.Border.Left.Style = ExcelBorderStyle.Hair;
                                ws.Cells[finCabecera, j].Style.Border.Right.Style = ExcelBorderStyle.Hair;
                                ws.Cells[finCabecera, j].Style.Border.Bottom.Style = ExcelBorderStyle.Hair;

                                switch (j)
                                {
                                    case 1:
                                        ws.Column(j).Width = 5;
                                        ws.Cells[finCabecera, j].Value = numeracion;
                                        break;
                                    case 2:
                                        ws.Cells[finCabecera, j].Value = detalle.codigo_cotizacion;
                                        break;
                                    case 3:
                                        ws.Cells[finCabecera, j].Value = detalle.nombre_comercial_cliente;
                                        ws.Cells[finCabecera, j].Style.WrapText = true;
                                        break;
                                    case 4:
                                        ws.Cells[finCabecera, j].Value = detalle.TipoIntermediario;
                                        ws.Cells[finCabecera, j].Style.WrapText = true;
                                        break;
                                    case 5:
                                        ws.Cells[finCabecera, j].Value = detalle.detalle_cotizacion;
                                        ws.Cells[finCabecera, j].Style.WrapText = true;
                                        break;
                                    case 6:
                                        ws.Cells[finCabecera, j].Value = detalle.total_pago;
                                        ws.Cells[finCabecera, j].Style.Numberformat.Format = "$#,##0.00";
                                        break;
                                    case 7:
                                        ws.Cells[finCabecera, j].Value = detalle.Ejecutivo;
                                        ws.Cells[finCabecera, j].Style.WrapText = true;
                                        break;
                                    case 8:
                                        ws.Cells[finCabecera, j].Value = detalle.numero_prefactura;
                                        break;
                                    case 9:
                                        ws.Cells[finCabecera, j].Value = detalle.ObservacionCotizacion;
                                        ws.Cells[finCabecera, j].Style.WrapText = true;
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
                    }

                    #endregion

                    #region Detalle Acta Contabilidad
                    if (acta.Cuerpo.DetalleContabilidad.Any())
                    {
                        var columnas = new List<string> { "N°", "CÓDIGO DE COTIZACIÓN", "CLIENTE", "INTERMEDIARIO", "DETALLE", "VALOR", "EJECUTIVO", "N. DOCUMENTO", "OBSERVACIONES" };

                        var i = 1;
                        foreach (var item in columnas)
                        {

                            ws.Cells[finCabecera, i].Value = item;
                            ws.Cells[finCabecera, i].Style.Font.Bold = true;

                            ws.Cells[finCabecera, i].Style.Border.Top.Style = ExcelBorderStyle.Hair;
                            ws.Cells[finCabecera, i].Style.Border.Left.Style = ExcelBorderStyle.Hair;
                            ws.Cells[finCabecera, i].Style.Border.Right.Style = ExcelBorderStyle.Hair;
                            ws.Cells[finCabecera, i].Style.Border.Bottom.Style = ExcelBorderStyle.Hair;

                            //worksheet.Cells[1, i].Style.Fill.BackgroundColor. .SetColor(Color.FromArgb(23, 55, 93));
                            i++;
                        }
                        finCabecera++;

                        int numeracion = 1;
                        foreach (var item in acta.Cuerpo.DetalleContabilidad)
                        {
                            var detalle = CotizacionEntity.ConsultarPrefacturaSAFI(item.id_facturacion_safi);
                            for (int j = 1; j <= 9; j++)
                            {
                                ws.Cells[finCabecera, j].Style.Border.Top.Style = ExcelBorderStyle.Hair;
                                ws.Cells[finCabecera, j].Style.Border.Left.Style = ExcelBorderStyle.Hair;
                                ws.Cells[finCabecera, j].Style.Border.Right.Style = ExcelBorderStyle.Hair;
                                ws.Cells[finCabecera, j].Style.Border.Bottom.Style = ExcelBorderStyle.Hair;

                                switch (j)
                                {
                                    case 1:
                                        ws.Column(j).Width = 5;
                                        ws.Cells[finCabecera, j].Value = numeracion;
                                        break;
                                    case 2:
                                        ws.Cells[finCabecera, j].Value = detalle.codigo_cotizacion;
                                        break;
                                    case 3:
                                        ws.Cells[finCabecera, j].Value = detalle.nombre_comercial_cliente;
                                        ws.Cells[finCabecera, j].Style.WrapText = true;
                                        break;
                                    case 4:
                                        ws.Cells[finCabecera, j].Value = detalle.TipoIntermediario;
                                        ws.Cells[finCabecera, j].Style.WrapText = true;
                                        break;
                                    case 5:
                                        ws.Cells[finCabecera, j].Value = detalle.detalle_cotizacion;
                                        ws.Cells[finCabecera, j].Style.WrapText = true;
                                        break;
                                    case 6:
                                        ws.Cells[finCabecera, j].Value = detalle.total_pago;
                                        ws.Cells[finCabecera, j].Style.Numberformat.Format = "$#,##0.00";
                                        break;
                                    case 7:
                                        ws.Cells[finCabecera, j].Value = detalle.Ejecutivo;
                                        ws.Cells[finCabecera, j].Style.WrapText = true;
                                        break;
                                    case 8:
                                        ws.Cells[finCabecera, j].Value = detalle.numero_prefactura;
                                        break;
                                    case 9:
                                        ws.Cells[finCabecera, j].Value = detalle.ObservacionCotizacion;
                                        ws.Cells[finCabecera, j].Style.WrapText = true;
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
                    }
                    #endregion

                    #region DetalleResponsables
                    if (acta.Cuerpo.Responsables.Any())
                    {
                        int totalColumnasDetalleResponsables = typeof(DetalleActaResponsables).GetProperties().Length;

                        int contadorRegistrosResponsables = acta.Cuerpo.Responsables.Count;

                        if (contadorRegistrosResponsables > 30)
                        {
                            for (int i = 1; i <= 2; i++)
                            {
                                int totalMergeColumnasParticipantes = i == 1 ? 3 : 6;

                                using (var range = ws.Cells[finCabecera, col, finCabecera, totalMergeColumnasParticipantes])
                                {
                                    range.Value = "NOMBRE";
                                    range.Style.Fill.PatternType = ExcelFillStyle.Solid;
                                    range.Style.Fill.BackgroundColor.SetColor(colorGrisClaro3EstiloPPM);
                                    range.Style.Font.Color.SetColor(colorGrisOscuroEstiloPPM);
                                    range.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                                    range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                                    // Assign borders
                                    range.Style.Border.Top.Style = ExcelBorderStyle.Hair;
                                    range.Style.Border.Left.Style = ExcelBorderStyle.Hair;
                                    range.Style.Border.Right.Style = ExcelBorderStyle.Hair;
                                    range.Style.Border.Bottom.Style = ExcelBorderStyle.Hair;

                                    range.Merge = true;

                                    if (col == 6)

                                        ws.Cells[finCabecera, 6, finCabecera, 7].Merge = true;
                                    ws.Cells[finCabecera, 6, finCabecera, 7].Style.Border.Top.Style = ExcelBorderStyle.Hair;
                                    ws.Cells[finCabecera, 6, finCabecera, 7].Style.Border.Left.Style = ExcelBorderStyle.Hair;
                                    ws.Cells[finCabecera, 6, finCabecera, 7].Style.Border.Right.Style = ExcelBorderStyle.Hair;
                                    ws.Cells[finCabecera, 6, finCabecera, 7].Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                                }

                                if (col == 6)
                                {
                                    col += 2;
                                }
                                else
                                {
                                    col += 3;
                                }

                                using (var range = ws.Cells[finCabecera, col, finCabecera, col])
                                {
                                    range.Value = "EMPRESA";
                                    range.Style.Fill.PatternType = ExcelFillStyle.Solid;
                                    range.Style.Fill.BackgroundColor.SetColor(colorGrisClaro3EstiloPPM);
                                    range.Style.Font.Color.SetColor(colorGrisOscuroEstiloPPM);
                                    range.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                                    range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                                    // Assign borders
                                    range.Style.Border.Top.Style = ExcelBorderStyle.Hair;
                                    range.Style.Border.Left.Style = ExcelBorderStyle.Hair;
                                    range.Style.Border.Right.Style = ExcelBorderStyle.Hair;
                                    range.Style.Border.Bottom.Style = ExcelBorderStyle.Hair;

                                    //range.Merge = true;
                                }
                                col += 1;
                                using (var range = ws.Cells[finCabecera, col, finCabecera, col])
                                {
                                    range.Value = "ROL";
                                    range.Style.Fill.PatternType = ExcelFillStyle.Solid;
                                    range.Style.Fill.BackgroundColor.SetColor(colorGrisClaro3EstiloPPM);
                                    range.Style.Font.Color.SetColor(colorGrisOscuroEstiloPPM);
                                    range.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                                    range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                                    // Assign borders
                                    range.Style.Border.Top.Style = ExcelBorderStyle.Hair;
                                    range.Style.Border.Left.Style = ExcelBorderStyle.Hair;
                                    range.Style.Border.Right.Style = ExcelBorderStyle.Hair;
                                    range.Style.Border.Bottom.Style = ExcelBorderStyle.Hair;

                                }
                                col += 1;
                            }

                            finCabecera++;

                            contador = 1;

                            int inicioFilaResponsables = finCabecera;
                            int columnasDivision1 = contadorRegistrosResponsables / 2;
                            int columnasDivision2 = contadorRegistrosResponsables - (contadorRegistrosResponsables / 2);

                            foreach (var column in acta.Cuerpo.Responsables)
                            {
                                col = contador <= columnasDivision1 ? 1 : 6;

                                if (contador == columnasDivision2 + 1)
                                {
                                    finCabecera = inicioFilaResponsables;
                                }

                                for (int i = 1; i <= totalColumnasDetalleResponsables; i++)
                                {
                                    ws.Cells[finCabecera, col].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                                    ws.Cells[finCabecera, col].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                                    switch (i)
                                    {
                                        case 1:
                                            if (col == 1)
                                            {
                                                ws.Cells[finCabecera, col, finCabecera, col + 2].Merge = true;
                                                ws.Cells[finCabecera, col, finCabecera, col + 2].Style.Border.Top.Style = ExcelBorderStyle.Hair;
                                                ws.Cells[finCabecera, col, finCabecera, col + 2].Style.Border.Left.Style = ExcelBorderStyle.Hair;
                                                ws.Cells[finCabecera, col, finCabecera, col + 2].Style.Border.Right.Style = ExcelBorderStyle.Hair;
                                                ws.Cells[finCabecera, col, finCabecera, col + 2].Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                                                ws.Cells[finCabecera, 6, finCabecera, 7].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                                                ws.Cells[finCabecera, col].Value = column.Nombres;
                                            }
                                            else
                                            {
                                                ws.Cells[finCabecera, 6, finCabecera, 7].Value = column.Nombres;
                                                ws.Cells[finCabecera, 6, finCabecera, 7].Merge = true;
                                                ws.Cells[finCabecera, 6, finCabecera, 7].Style.Border.Top.Style = ExcelBorderStyle.Hair;
                                                ws.Cells[finCabecera, 6, finCabecera, 7].Style.Border.Left.Style = ExcelBorderStyle.Hair;
                                                ws.Cells[finCabecera, 6, finCabecera, 7].Style.Border.Right.Style = ExcelBorderStyle.Hair;
                                                ws.Cells[finCabecera, 6, finCabecera, 7].Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                                                ws.Cells[finCabecera, 6, finCabecera, 7].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                                            }
                                            col += 3;
                                            break;

                                        case 2:
                                            if (col == 4)
                                            {
                                                ws.Cells[finCabecera, col].Merge = true;
                                                ws.Cells[finCabecera, col].Style.Border.Top.Style = ExcelBorderStyle.Hair;
                                                ws.Cells[finCabecera, col].Style.Border.Left.Style = ExcelBorderStyle.Hair;
                                                ws.Cells[finCabecera, col].Style.Border.Right.Style = ExcelBorderStyle.Hair;
                                                ws.Cells[finCabecera, col].Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                                                ws.Cells[finCabecera, col].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                                                ws.Cells[finCabecera, col].Value = column.Empresa;
                                                col += 1;
                                            }
                                            else
                                            {

                                                ws.Cells[finCabecera, 8, finCabecera, 8].Merge = true;
                                                ws.Cells[finCabecera, 8, finCabecera, 8].Style.Border.Top.Style = ExcelBorderStyle.Hair;
                                                ws.Cells[finCabecera, 8, finCabecera, 8].Style.Border.Left.Style = ExcelBorderStyle.Hair;
                                                ws.Cells[finCabecera, 8, finCabecera, 8].Style.Border.Right.Style = ExcelBorderStyle.Hair;
                                                ws.Cells[finCabecera, 8, finCabecera, 8].Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                                                ws.Cells[finCabecera, 8, finCabecera, 8].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                                                ws.Cells[finCabecera, 8, finCabecera, 8].Value = column.Empresa;

                                            }
                                            break;


                                        case 3:
                                            if (col == 5)
                                            {
                                                ws.Cells[finCabecera, col].Style.Border.Top.Style = ExcelBorderStyle.Hair;
                                                ws.Cells[finCabecera, col].Style.Border.Left.Style = ExcelBorderStyle.Hair;
                                                ws.Cells[finCabecera, col].Style.Border.Right.Style = ExcelBorderStyle.Hair;
                                                ws.Cells[finCabecera, col].Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                                                ws.Cells[finCabecera, col].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                                                ws.Cells[finCabecera, col].Value = column.Rol;
                                            }
                                            else
                                            {

                                                ws.Cells[finCabecera, 9, finCabecera, 9].Value = column.Rol;
                                                ws.Cells[finCabecera, 9, finCabecera, 9].Style.Border.Top.Style = ExcelBorderStyle.Hair;
                                                ws.Cells[finCabecera, 9, finCabecera, 9].Style.Border.Left.Style = ExcelBorderStyle.Hair;
                                                ws.Cells[finCabecera, 9, finCabecera, 9].Style.Border.Right.Style = ExcelBorderStyle.Hair;
                                                ws.Cells[finCabecera, 9, finCabecera, 9].Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                                                ws.Cells[finCabecera, 9, finCabecera, 9].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;

                                            }
                                            col += 1;

                                            break;


                                        default:
                                            ws.Cells[finCabecera, col++].Value = "";
                                            break;
                                    }
                                }

                                ws.Row(finCabecera).Height = 20.25;
                                finCabecera++;
                                contador++;
                            }
                        }
                        else
                        {
                            for (int j = 1; j <= totalColumnasDetalleResponsables; j++)
                            {

                                ws.Cells[finCabecera, col].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                ws.Cells[finCabecera, col].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(31, 78, 120));
                                ws.Cells[finCabecera, col].Style.Font.Color.SetColor(Color.White);
                                ws.Cells[finCabecera, col].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                                ws.Cells[finCabecera, col].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                // Assign borders
                                ws.Cells[finCabecera, col].Style.Border.Top.Style = ExcelBorderStyle.Hair;
                                ws.Cells[finCabecera, col].Style.Border.Left.Style = ExcelBorderStyle.Hair;
                                ws.Cells[finCabecera, col].Style.Border.Right.Style = ExcelBorderStyle.Hair;
                                ws.Cells[finCabecera, col].Style.Border.Bottom.Style = ExcelBorderStyle.Hair;

                                switch (j)
                                {
                                    case 1:
                                        using (var range = ws.Cells[finCabecera, col++, finCabecera, col + 2])
                                        {
                                            range.Style.Border.Top.Style = ExcelBorderStyle.Hair;
                                            range.Style.Border.Left.Style = ExcelBorderStyle.Hair;
                                            range.Style.Border.Right.Style = ExcelBorderStyle.Hair;
                                            range.Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                                            range.Merge = true;

                                            range.Value = "NOMBRE";
                                            range.Style.Fill.BackgroundColor.SetColor(colorGrisClaro3EstiloPPM);
                                            range.Style.Font.Color.SetColor(colorGrisOscuroEstiloPPM);
                                            range.Style.Font.Bold = true;

                                        }
                                        break;

                                    case 2:
                                        using (var range = ws.Cells[finCabecera, col + 3, finCabecera, col + 5])
                                        {
                                            range.Style.Border.Top.Style = ExcelBorderStyle.Hair;
                                            range.Style.Border.Left.Style = ExcelBorderStyle.Hair;
                                            range.Style.Border.Right.Style = ExcelBorderStyle.Hair;
                                            range.Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                                            range.Merge = true;

                                            range.Value = "EMPRESA";
                                            range.Style.Fill.BackgroundColor.SetColor(colorGrisClaro3EstiloPPM);
                                            range.Style.Font.Color.SetColor(colorGrisOscuroEstiloPPM);
                                            range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                            range.Style.Font.Bold = true;

                                        }
                                        break;

                                    case 3:
                                        using (var range = ws.Cells[finCabecera, col + 6, finCabecera, col + 7])
                                        {
                                            range.Style.Border.Top.Style = ExcelBorderStyle.Hair;
                                            range.Style.Border.Left.Style = ExcelBorderStyle.Hair;
                                            range.Style.Border.Right.Style = ExcelBorderStyle.Hair;
                                            range.Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                                            range.Merge = true;

                                            range.Value = "ROL";
                                            range.Style.Fill.BackgroundColor.SetColor(colorGrisClaro3EstiloPPM);
                                            range.Style.Font.Color.SetColor(colorGrisOscuroEstiloPPM);
                                            range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                            range.Style.Font.Bold = true;
                                        }
                                        break;
                                    default:
                                        ws.Cells[finCabecera, col++].Value = "";
                                        break;
                                }
                            }

                            finCabecera++;
                            contador = 1;
                            foreach (var column in acta.Cuerpo.Responsables)
                            {
                                col = 1;
                                for (int i = 1; i <= totalColumnasDetalleResponsables; i++)
                                {
                                    ws.Cells[finCabecera, col].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                                    ws.Cells[finCabecera, col].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                    // Assign borders
                                    ws.Cells[finCabecera, col].Style.Border.Top.Style = ExcelBorderStyle.Hair;
                                    ws.Cells[finCabecera, col].Style.Border.Left.Style = ExcelBorderStyle.Hair;
                                    ws.Cells[finCabecera, col].Style.Border.Right.Style = ExcelBorderStyle.Hair;
                                    ws.Cells[finCabecera, col].Style.Border.Bottom.Style = ExcelBorderStyle.Hair;

                                    switch (i)
                                    {
                                        case 1:

                                            using (var range = ws.Cells[finCabecera, col++, finCabecera, col + 2])
                                            {
                                                range.Value = column.Nombres;

                                                range.Style.Border.Top.Style = ExcelBorderStyle.Hair;
                                                range.Style.Border.Left.Style = ExcelBorderStyle.Hair;
                                                range.Style.Border.Right.Style = ExcelBorderStyle.Hair;
                                                range.Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                                                range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                                                range.Merge = true;

                                            }

                                            break;

                                        case 2:

                                            using (var range = ws.Cells[finCabecera, col + 3, finCabecera, col + 5])
                                            {
                                                range.Value = column.Empresa;

                                                range.Style.Border.Top.Style = ExcelBorderStyle.Hair;
                                                range.Style.Border.Left.Style = ExcelBorderStyle.Hair;
                                                range.Style.Border.Right.Style = ExcelBorderStyle.Hair;
                                                range.Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                                                range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                                                range.Merge = true;


                                            }
                                            break;

                                        case 3:

                                            using (var range = ws.Cells[finCabecera, col + 6, finCabecera, col + 7])
                                            {
                                                range.Value = column.Rol;

                                                range.Style.Border.Top.Style = ExcelBorderStyle.Hair;
                                                range.Style.Border.Left.Style = ExcelBorderStyle.Hair;
                                                range.Style.Border.Right.Style = ExcelBorderStyle.Hair;
                                                range.Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                                                range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                                                range.Merge = true;

                                            }

                                            break;

                                        default:
                                            ws.Cells[finCabecera, col++].Value = "";
                                            break;
                                    }
                                }
                                ws.Row(finCabecera).Height = 20.25;
                                finCabecera++;
                                contador++;
                            }
                        }

                        ws.Column(2).AutoFit();
                        ws.Column(3).AutoFit();
                        ws.Column(4).AutoFit();
                        ws.Column(6).AutoFit();
                        ws.Column(7).AutoFit();
                        ws.Column(8).AutoFit();

                        ws.Row(finCabecera).Height = 8.25;
                        finCabecera += 1;
                    }
                    #endregion

                    #region DetalleEntregable
                    if (acta.Cuerpo.Entregables.Any())
                    {

                        col = 1;
                        int totalColumnasDetalleEntregables = typeof(DetalleActaEntregables).GetProperties().Length;

                        for (int i = 1; i <= totalColumnasDetalleEntregables; i++)
                        {
                            ws.Cells[finCabecera, col].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                            ws.Cells[finCabecera, col].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            //Assign borders
                            ws.Cells[finCabecera, col].Style.Border.Top.Style = ExcelBorderStyle.Hair;
                            ws.Cells[finCabecera, col].Style.Border.Left.Style = ExcelBorderStyle.Hair;
                            ws.Cells[finCabecera, col].Style.Border.Right.Style = ExcelBorderStyle.Hair;
                            ws.Cells[finCabecera, col].Style.Border.Bottom.Style = ExcelBorderStyle.Hair;

                            switch (i)
                            {
                                case 1:
                                    ws.Cells[finCabecera, col++].Value = "N.";
                                    ws.Cells[finCabecera, 1, finCabecera, 1].Style.Border.Top.Style = ExcelBorderStyle.Hair;
                                    ws.Cells[finCabecera, 1, finCabecera, 1].Style.Border.Left.Style = ExcelBorderStyle.Hair;
                                    ws.Cells[finCabecera, 1, finCabecera, 1].Style.Border.Right.Style = ExcelBorderStyle.Hair;
                                    ws.Cells[finCabecera, 1, finCabecera, 1].Style.Border.Bottom.Style = ExcelBorderStyle.Hair;

                                    ws.Cells[finCabecera, 1, finCabecera, 1].Style.Fill.BackgroundColor.SetColor(colorGrisClaro3EstiloPPM);
                                    ws.Cells[finCabecera, 1, finCabecera, 1].Style.Font.Color.SetColor(colorGrisOscuroEstiloPPM);
                                    ws.Cells[finCabecera, 1, finCabecera, 1].Style.Font.Bold = true;

                                    break;
                                case 2:
                                    ws.Cells[finCabecera, col++].Value = "ENTREGABLES";
                                    ws.Cells[finCabecera, 2, finCabecera, 7].Merge = true;
                                    ws.Cells[finCabecera, 2, finCabecera, 7].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                                    ws.Cells[finCabecera, 2, finCabecera, 7].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                    // Assign borders
                                    ws.Cells[finCabecera, 2, finCabecera, 7].Style.Border.Top.Style = ExcelBorderStyle.Hair;
                                    ws.Cells[finCabecera, 2, finCabecera, 7].Style.Border.Left.Style = ExcelBorderStyle.Hair;
                                    ws.Cells[finCabecera, 2, finCabecera, 7].Style.Border.Right.Style = ExcelBorderStyle.Hair;
                                    ws.Cells[finCabecera, 2, finCabecera, 7].Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                                    ws.Cells[finCabecera, 2, finCabecera, 7].Style.Font.Bold = true;

                                    ws.Cells[finCabecera, 2, finCabecera, 7].Style.Fill.BackgroundColor.SetColor(colorGrisClaro3EstiloPPM);
                                    ws.Cells[finCabecera, 2, finCabecera, 7].Style.Font.Color.SetColor(colorGrisOscuroEstiloPPM);

                                    break;
                                case 3:
                                    ws.Cells[finCabecera, 8].Value = "TIPO";
                                    ws.Cells[finCabecera, 8, finCabecera, 9].Merge = true;
                                    ws.Cells[finCabecera, 8, finCabecera, 9].Style.Fill.PatternType = ExcelFillStyle.Solid;

                                    ws.Cells[finCabecera, 8, finCabecera, 9].Style.Fill.BackgroundColor.SetColor(colorGrisClaro3EstiloPPM);
                                    ws.Cells[finCabecera, 8, finCabecera, 9].Style.Font.Color.SetColor(colorGrisOscuroEstiloPPM);

                                    ws.Cells[finCabecera, 8, finCabecera, 9].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                                    ws.Cells[finCabecera, 8, finCabecera, 9].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                    // Assign border
                                    ws.Cells[finCabecera, 8, finCabecera, 9].Style.Border.Top.Style = ExcelBorderStyle.Hair;
                                    ws.Cells[finCabecera, 8, finCabecera, 9].Style.Border.Left.Style = ExcelBorderStyle.Hair;
                                    ws.Cells[finCabecera, 8, finCabecera, 9].Style.Border.Right.Style = ExcelBorderStyle.Hair;
                                    ws.Cells[finCabecera, 8, finCabecera, 9].Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                                    ws.Cells[finCabecera, 8, finCabecera, 9].Style.Font.Bold = true;


                                    break;
                                default:
                                    ws.Cells[finCabecera, col++].Value = "";
                                    break;
                            }
                            ws.Row(finCabecera).Height = 20.25;
                        }

                        finCabecera++;

                        contador = 1;
                        foreach (var column in acta.Cuerpo.Entregables)
                        {
                            col = 1;
                            for (int i = 1; i <= totalColumnasDetalleEntregables; i++)
                            {
                                ws.Cells[finCabecera, col].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                                ws.Cells[finCabecera, col].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                // Assign borders
                                ws.Cells[finCabecera, col].Style.Border.Top.Style = ExcelBorderStyle.Hair;
                                ws.Cells[finCabecera, col].Style.Border.Left.Style = ExcelBorderStyle.Hair;
                                ws.Cells[finCabecera, col].Style.Border.Right.Style = ExcelBorderStyle.Hair;
                                ws.Cells[finCabecera, col].Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                                switch (i)
                                {
                                    case 1:
                                        ws.Cells[finCabecera, col++].Value = contador;
                                        ws.Cells[finCabecera, 1, finCabecera, 1].Style.Border.Top.Style = ExcelBorderStyle.Hair;
                                        ws.Cells[finCabecera, 1, finCabecera, 1].Style.Border.Left.Style = ExcelBorderStyle.Hair;
                                        ws.Cells[finCabecera, 1, finCabecera, 1].Style.Border.Right.Style = ExcelBorderStyle.Hair;
                                        ws.Cells[finCabecera, 1, finCabecera, 1].Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                                        break;
                                    case 2:
                                        ws.Cells[finCabecera, col++].Value = column.Entregable;
                                        ws.Cells[finCabecera, 2, finCabecera, 7].Merge = true;
                                        ws.Cells[finCabecera, 2, finCabecera, 7].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                                        ws.Cells[finCabecera, 2, finCabecera, 7].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                                        // Assign borders
                                        ws.Cells[finCabecera, 2, finCabecera, 7].Style.Border.Top.Style = ExcelBorderStyle.Hair;
                                        ws.Cells[finCabecera, 2, finCabecera, 7].Style.Border.Left.Style = ExcelBorderStyle.Hair;
                                        ws.Cells[finCabecera, 2, finCabecera, 7].Style.Border.Right.Style = ExcelBorderStyle.Hair;
                                        ws.Cells[finCabecera, 2, finCabecera, 7].Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                                        ws.Cells[finCabecera, 2, finCabecera, 7].Style.WrapText = true;

                                        var caracteres = column.Entregable.Count();
                                        if (caracteres <= 70)
                                        {
                                            ws.Row(finCabecera).Height = 20.25;
                                        }
                                        else if (caracteres <= 140)
                                        {
                                            ws.Row(finCabecera).Height = 30.50;
                                        }
                                        else if (caracteres <= 210)
                                        {
                                            ws.Row(finCabecera).Height = 40.75;
                                        }
                                        else if (caracteres <= 280)
                                        {
                                            ws.Row(finCabecera).Height = 51;
                                        }
                                        else if (caracteres <= 350)
                                        {
                                            ws.Row(finCabecera).Height = 61.25;
                                        }
                                        else if (caracteres <= 420)
                                        {
                                            ws.Row(finCabecera).Height = 71.50;
                                        }
                                        else if (caracteres <= 500)
                                        {
                                            ws.Row(finCabecera).Height = 81.75;
                                        }
                                        else
                                        {
                                            ws.Row(finCabecera).Height = 92;
                                        }

                                        break;
                                    case 3:
                                        ws.Cells[finCabecera, 8].Value = column.Tipo;
                                        ws.Cells[finCabecera, 8, finCabecera, 9].Merge = true;
                                        ws.Cells[finCabecera, 8, finCabecera, 9].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                                        ws.Cells[finCabecera, 8, finCabecera, 9].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                                        // Assign border
                                        ws.Cells[finCabecera, 8, finCabecera, 9].Style.Border.Top.Style = ExcelBorderStyle.Hair;
                                        ws.Cells[finCabecera, 8, finCabecera, 9].Style.Border.Left.Style = ExcelBorderStyle.Hair;
                                        ws.Cells[finCabecera, 8, finCabecera, 9].Style.Border.Right.Style = ExcelBorderStyle.Hair;
                                        ws.Cells[finCabecera, 8, finCabecera, 9].Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                                        ws.Cells[finCabecera, 8, finCabecera, 9].Style.WrapText = true;
                                        break;
                                    default:
                                        ws.Cells[finCabecera, col++].Value = "";
                                        break;
                                }

                            }
                            finCabecera++;
                            contador++;
                        }

                        ws.Row(finCabecera).Height = 8.25;
                        finCabecera += 1;

                    }
                    #endregion

                    #region Detalle Condiciones Generales
                    if (acta.Cuerpo.CondicionesGenerales.Any())
                    {
                        col = 1;
                        int totalColumnasDetalleCondiciones = typeof(DetalleActaCondicionesGenerales).GetProperties().Length;

                        for (int i = 1; i <= totalColumnasDetalleCondiciones; i++)
                        {

                            ws.Cells[finCabecera, col].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            ws.Cells[finCabecera, col].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                            ws.Cells[finCabecera, col].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            // Assign borders
                            ws.Cells[finCabecera, col].Style.Border.Top.Style = ExcelBorderStyle.Hair;
                            ws.Cells[finCabecera, col].Style.Border.Left.Style = ExcelBorderStyle.Hair;
                            ws.Cells[finCabecera, col].Style.Border.Right.Style = ExcelBorderStyle.Hair;
                            ws.Cells[finCabecera, col].Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                            switch (i)
                            {
                                case 1:
                                    ws.Cells[finCabecera, col++].Value = "N.";
                                    ws.Cells[finCabecera, 1, finCabecera, 1].Style.Border.Top.Style = ExcelBorderStyle.Hair;
                                    ws.Cells[finCabecera, 1, finCabecera, 1].Style.Border.Left.Style = ExcelBorderStyle.Hair;
                                    ws.Cells[finCabecera, 1, finCabecera, 1].Style.Border.Right.Style = ExcelBorderStyle.Hair;
                                    ws.Cells[finCabecera, 1, finCabecera, 1].Style.Border.Bottom.Style = ExcelBorderStyle.Hair;

                                    ws.Cells[finCabecera, 1, finCabecera, 1].Style.Font.Bold = true;

                                    ws.Cells[finCabecera, 1, finCabecera, 1].Style.Fill.BackgroundColor.SetColor(colorGrisClaro3EstiloPPM);
                                    ws.Cells[finCabecera, 1, finCabecera, 1].Style.Font.Color.SetColor(colorGrisOscuroEstiloPPM);
                                    break;
                                case 2:
                                    ws.Cells[finCabecera, col++].Value = "CONDICIONES GENERALES";
                                    ws.Cells[finCabecera, 2, finCabecera, 9].Merge = true;

                                    ws.Cells[finCabecera, col].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                    ws.Cells[finCabecera, 2, finCabecera, 9].Style.Fill.BackgroundColor.SetColor(colorGrisClaro3EstiloPPM);
                                    ws.Cells[finCabecera, 2, finCabecera, 9].Style.Font.Color.SetColor(colorGrisOscuroEstiloPPM);
                                    ws.Cells[finCabecera, 2, finCabecera, 9].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                                    ws.Cells[finCabecera, 2, finCabecera, 9].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                    // Assign border
                                    ws.Cells[finCabecera, 2, finCabecera, 9].Style.Border.Top.Style = ExcelBorderStyle.Hair;
                                    ws.Cells[finCabecera, 2, finCabecera, 9].Style.Border.Left.Style = ExcelBorderStyle.Hair;
                                    ws.Cells[finCabecera, 2, finCabecera, 9].Style.Border.Right.Style = ExcelBorderStyle.Hair;
                                    ws.Cells[finCabecera, 2, finCabecera, 9].Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                                    ws.Cells[finCabecera, 2, finCabecera, 9].Style.Font.Bold = true;
                                    break;
                                default:
                                    ws.Cells[finCabecera, col++].Value = "";
                                    break;
                            }

                            ws.Row(finCabecera).Height = 20.25;
                        }

                        finCabecera++;

                        contador = 1;
                        foreach (var column in acta.Cuerpo.CondicionesGenerales)
                        {
                            col = 1;
                            for (int i = 1; i <= totalColumnasDetalleCondiciones; i++)
                            {
                                ws.Cells[finCabecera, col].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                                ws.Cells[finCabecera, col].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                // Assign borders
                                ws.Cells[finCabecera, col].Style.Border.Top.Style = ExcelBorderStyle.Hair;
                                ws.Cells[finCabecera, col].Style.Border.Left.Style = ExcelBorderStyle.Hair;
                                ws.Cells[finCabecera, col].Style.Border.Right.Style = ExcelBorderStyle.Hair;
                                ws.Cells[finCabecera, col].Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                                switch (i)
                                {
                                    case 1:
                                        ws.Cells[finCabecera, col++].Value = contador;
                                        ws.Cells[finCabecera, 1, finCabecera, 1].Style.Border.Top.Style = ExcelBorderStyle.Hair;
                                        ws.Cells[finCabecera, 1, finCabecera, 1].Style.Border.Left.Style = ExcelBorderStyle.Hair;
                                        ws.Cells[finCabecera, 1, finCabecera, 1].Style.Border.Right.Style = ExcelBorderStyle.Hair;
                                        ws.Cells[finCabecera, 1, finCabecera, 1].Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                                        break;
                                    case 2:
                                        ws.Cells[finCabecera, col++].Value = column.Condicion;
                                        ws.Cells[finCabecera, 2, finCabecera, 9].Merge = true;
                                        ws.Cells[finCabecera, 2, finCabecera, 9].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                                        ws.Cells[finCabecera, 2, finCabecera, 9].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                                        ws.Cells[finCabecera, 2, finCabecera, 9].Style.Border.Top.Style = ExcelBorderStyle.Hair;
                                        ws.Cells[finCabecera, 2, finCabecera, 9].Style.Border.Left.Style = ExcelBorderStyle.Hair;
                                        ws.Cells[finCabecera, 2, finCabecera, 9].Style.Border.Right.Style = ExcelBorderStyle.Hair;
                                        ws.Cells[finCabecera, 2, finCabecera, 9].Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                                        ws.Cells[finCabecera, 2, finCabecera, 9].Style.WrapText = true;

                                        var caracteres = column.Condicion.Count();
                                        if (caracteres <= 97)
                                        {
                                            ws.Row(finCabecera).Height = 20.25;
                                        }
                                        else if (caracteres <= 194)
                                        {
                                            ws.Row(finCabecera).Height = 30.50;
                                        }
                                        else if (caracteres <= 291)
                                        {
                                            ws.Row(finCabecera).Height = 40.75;
                                        }
                                        else if (caracteres <= 388)
                                        {
                                            ws.Row(finCabecera).Height = 51;
                                        }
                                        else if (caracteres <= 582)
                                        {
                                            ws.Row(finCabecera).Height = 61.25;
                                        }
                                        else
                                        {
                                            ws.Row(finCabecera).Height = 71.5;
                                        }

                                        break;
                                    default:
                                        ws.Cells[finCabecera, col++].Value = "";
                                        break;
                                }

                            }
                            finCabecera++;
                            contador++;
                        }
                        ws.Row(finCabecera).Height = 8.25;
                        finCabecera += 1;
                    }


                    #endregion

                    #region Detalle participantes
                    if (acta.Cuerpo.Participantes.Any())
                    {
                        col = 1;
                        int totalColumnasDetalleParticipantes = typeof(DetalleActaParticipantes).GetProperties().Length - 1;

                        int contadorRegistrosParticipantes = acta.Cuerpo.Participantes.Count;

                        if (contadorRegistrosParticipantes > 3)
                        {

                            for (int i = 1; i <= 2; i++)
                            {
                                int totalMergeColumnasParticipantes = i == 1 ? 3 : 8;

                                using (var range = ws.Cells[finCabecera, col, finCabecera, totalMergeColumnasParticipantes])
                                {
                                    range.Value = "PARTICIPANTES";
                                    range.Style.Fill.PatternType = ExcelFillStyle.Solid;
                                    range.Style.Fill.BackgroundColor.SetColor(colorGrisClaro3EstiloPPM);
                                    range.Style.Font.Color.SetColor(colorGrisOscuroEstiloPPM);
                                    range.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                                    range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                                    // Assign borders
                                    range.Style.Border.Top.Style = ExcelBorderStyle.Hair;
                                    range.Style.Border.Left.Style = ExcelBorderStyle.Hair;
                                    range.Style.Border.Right.Style = ExcelBorderStyle.Hair;
                                    range.Style.Border.Bottom.Style = ExcelBorderStyle.Hair;

                                    range.Merge = true;
                                    range.Style.Font.Bold = true;
                                }

                                if (totalMergeColumnasParticipantes == 8)
                                    col++;

                                col += 3;

                                using (var range = ws.Cells[finCabecera, col, finCabecera, col])
                                {
                                    range.Value = "PRESENTE";
                                    range.Style.Fill.PatternType = ExcelFillStyle.Solid;
                                    range.Style.Fill.BackgroundColor.SetColor(colorGrisClaro3EstiloPPM);
                                    range.Style.Font.Color.SetColor(colorGrisOscuroEstiloPPM);
                                    range.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                                    range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                                    // Assign borders
                                    range.Style.Border.Top.Style = ExcelBorderStyle.Hair;
                                    range.Style.Border.Left.Style = ExcelBorderStyle.Hair;
                                    range.Style.Border.Right.Style = ExcelBorderStyle.Hair;
                                    range.Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                                    range.Style.Font.Bold = true;
                                }
                                col += 1;
                            }

                            finCabecera++;

                            contador = 1;
                            int inicioFilaParticipantes = finCabecera;
                            int columnasDivision1 = contadorRegistrosParticipantes / 2;
                            int columnasDivision2 = contadorRegistrosParticipantes - (contadorRegistrosParticipantes / 2);

                            foreach (var column in acta.Cuerpo.Participantes)
                            {
                                col = contador <= columnasDivision1 ? 1 : 5;

                                if (contador == columnasDivision2 + 1)
                                {
                                    finCabecera = inicioFilaParticipantes;
                                }

                                for (int i = 1; i <= totalColumnasDetalleParticipantes; i++)
                                {
                                    ws.Cells[finCabecera, col].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                                    ws.Cells[finCabecera, col].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                                    // Assign borders
                                    ws.Cells[finCabecera, col].Style.Border.Top.Style = ExcelBorderStyle.Hair;
                                    ws.Cells[finCabecera, col].Style.Border.Left.Style = ExcelBorderStyle.Hair;
                                    ws.Cells[finCabecera, col].Style.Border.Right.Style = ExcelBorderStyle.Hair;
                                    ws.Cells[finCabecera, col].Style.Border.Bottom.Style = ExcelBorderStyle.Hair;

                                    if (col == 8)
                                        col++;

                                    switch (i)
                                    {
                                        case 1:
                                            ws.Cells[finCabecera, col].Value = column.Nombres;
                                            ws.Cells[finCabecera, col, finCabecera, col + 2].Merge = true;
                                            ws.Cells[finCabecera, col, finCabecera, col + 2].Style.Border.Top.Style = ExcelBorderStyle.Hair;
                                            ws.Cells[finCabecera, col, finCabecera, col + 2].Style.Border.Left.Style = ExcelBorderStyle.Hair;
                                            ws.Cells[finCabecera, col, finCabecera, col + 2].Style.Border.Right.Style = ExcelBorderStyle.Hair;
                                            ws.Cells[finCabecera, col, finCabecera, col + 2].Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                                            ws.Cells[finCabecera, col, finCabecera, col + 2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                                            ws.Cells[finCabecera, col, finCabecera, col + 2].Style.WrapText = true;
                                            if (col == 5)

                                                ws.Cells[finCabecera, 5, finCabecera, 8].Merge = true;
                                            ws.Cells[finCabecera, 5, finCabecera, 8].Style.Border.Top.Style = ExcelBorderStyle.Hair;
                                            ws.Cells[finCabecera, 5, finCabecera, 8].Style.Border.Left.Style = ExcelBorderStyle.Hair;
                                            ws.Cells[finCabecera, 5, finCabecera, 8].Style.Border.Right.Style = ExcelBorderStyle.Hair;
                                            ws.Cells[finCabecera, 5, finCabecera, 8].Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                                            ws.Cells[finCabecera, 5, finCabecera, 8].Style.WrapText = true;
                                            col += 3;

                                            break;
                                        case 2:
                                            ws.Cells[finCabecera, col].Value = column.Presente ? "SI" : "NO";
                                            ws.Cells[finCabecera, col].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                                            ws.Cells[finCabecera, col].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                            ws.Cells[finCabecera, col].Style.Border.Top.Style = ExcelBorderStyle.Hair;
                                            ws.Cells[finCabecera, col].Style.Border.Left.Style = ExcelBorderStyle.Hair;
                                            ws.Cells[finCabecera, col].Style.Border.Right.Style = ExcelBorderStyle.Hair;
                                            ws.Cells[finCabecera, col].Style.Border.Bottom.Style = ExcelBorderStyle.Hair;

                                            break;
                                        default:
                                            ws.Cells[finCabecera, col++].Value = "";
                                            break;
                                    }

                                }
                                ws.Cells[finCabecera, 1, finCabecera, 9].Style.WrapText = true;
                                ws.Row(finCabecera).Height = 20.25;
                                finCabecera++;
                                contador++;
                            }
                        }
                        else
                        {
                            for (int j = 1; j <= totalColumnasDetalleParticipantes; j++)
                            {
                                ws.Cells[finCabecera, col].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                ws.Cells[finCabecera, col].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                                ws.Cells[finCabecera, col].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                                // Assign borders
                                ws.Cells[finCabecera, col].Style.Border.Top.Style = ExcelBorderStyle.Hair;
                                ws.Cells[finCabecera, col].Style.Border.Left.Style = ExcelBorderStyle.Hair;
                                ws.Cells[finCabecera, col].Style.Border.Right.Style = ExcelBorderStyle.Hair;
                                ws.Cells[finCabecera, col].Style.Border.Bottom.Style = ExcelBorderStyle.Hair;

                                switch (j)
                                {
                                    case 1:
                                        using (var range = ws.Cells[finCabecera, col++, finCabecera, col + 5])
                                        {
                                            range.Style.Border.Top.Style = ExcelBorderStyle.Hair;
                                            range.Style.Border.Left.Style = ExcelBorderStyle.Hair;
                                            range.Style.Border.Right.Style = ExcelBorderStyle.Hair;
                                            range.Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                                            range.Merge = true;

                                            range.Value = "PARTICIPANTES";
                                            range.Style.Fill.BackgroundColor.SetColor(colorGrisClaro3EstiloPPM);
                                            range.Style.Font.Color.SetColor(colorGrisOscuroEstiloPPM);
                                            range.Style.Font.Bold = true;

                                        }
                                        using (var range = ws.Cells[finCabecera, col + 6, finCabecera, col + 7])
                                        {
                                            range.Style.Border.Top.Style = ExcelBorderStyle.Hair;
                                            range.Style.Border.Left.Style = ExcelBorderStyle.Hair;
                                            range.Style.Border.Right.Style = ExcelBorderStyle.Hair;
                                            range.Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                                            range.Merge = true;

                                            range.Value = "PRESENTES";
                                            range.Style.Fill.BackgroundColor.SetColor(colorGrisClaro3EstiloPPM);
                                            range.Style.Font.Color.SetColor(colorGrisOscuroEstiloPPM);
                                            range.Style.HorizontalAlignment = ExcelHorizontalAlignment.CenterContinuous;
                                            range.Style.Font.Bold = true;

                                        }

                                        break;
                                    case 2:
                                        //ws.Cells[finCabecera, col++].Value = "PRESENTES";
                                        break;
                                    default:
                                        ws.Cells[finCabecera, col++].Value = "";
                                        break;
                                }
                            }

                            finCabecera++;
                            contador = 1;
                            foreach (var column in acta.Cuerpo.Participantes)
                            {
                                col = 1;
                                for (int i = 1; i <= totalColumnasDetalleParticipantes; i++)
                                {
                                    ws.Cells[finCabecera, col].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                                    ws.Cells[finCabecera, col].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;

                                    // Assign borders
                                    ws.Cells[finCabecera, col].Style.Border.Top.Style = ExcelBorderStyle.Hair;
                                    ws.Cells[finCabecera, col].Style.Border.Left.Style = ExcelBorderStyle.Hair;
                                    ws.Cells[finCabecera, col].Style.Border.Right.Style = ExcelBorderStyle.Hair;
                                    ws.Cells[finCabecera, col].Style.Border.Bottom.Style = ExcelBorderStyle.Hair;

                                    switch (i)
                                    {

                                        case 1:

                                            using (var range = ws.Cells[finCabecera, col++, finCabecera, col + 5])
                                            {
                                                range.Value = column.Nombres;

                                                range.Style.Border.Top.Style = ExcelBorderStyle.Hair;
                                                range.Style.Border.Left.Style = ExcelBorderStyle.Hair;
                                                range.Style.Border.Right.Style = ExcelBorderStyle.Hair;
                                                range.Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                                                range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                                                range.Merge = true;
                                                range.Style.WrapText = true;

                                            }
                                            break;
                                        case 2:

                                            using (var range = ws.Cells[finCabecera, col + 6, finCabecera, col + 7])
                                            {
                                                range.Value = column.Presente ? "SI" : "NO";

                                                range.Style.Border.Top.Style = ExcelBorderStyle.Hair;
                                                range.Style.Border.Left.Style = ExcelBorderStyle.Hair;
                                                range.Style.Border.Right.Style = ExcelBorderStyle.Hair;
                                                range.Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                                                range.Style.HorizontalAlignment = ExcelHorizontalAlignment.CenterContinuous;
                                                range.Merge = true;
                                                range.Style.WrapText = true;

                                            }
                                            //ws.Cells[finCabecera, col++].Value = column.Presente ? "SI" : "NO";
                                            break;
                                        default:
                                            ws.Cells[finCabecera, col++].Value = "";

                                            break;
                                    }
                                }

                                ws.Cells[finCabecera, 1, finCabecera, 9].Style.WrapText = true;
                                ws.Row(finCabecera).Height = 20.25;
                                finCabecera++;
                                contador++;

                            }

                        }

                        if (acta.Cuerpo.Participantes.Any())
                        {
                            if (acta.Cuerpo.Participantes.Count > 3 && acta.Cuerpo.Participantes.Count % 2 != 0)
                                finCabecera += 2;
                            else
                            {
                                ws.Row(finCabecera).Height = 8.25;
                                finCabecera++;
                            }
                        }
                        else
                        {
                            ws.Row(finCabecera).Height = 8.25;
                            finCabecera++;
                        }

                        int filaDatosParticipantes = finCabecera;
                        int inicio = filaDatosParticipantes;

                        int cantidadParticipantes = acta.Cuerpo.Participantes.Count;
                        int cantidadSI = acta.Cuerpo.Participantes.Where(s => s.Presente).Count();
                        int cantidadNO = acta.Cuerpo.Participantes.Where(s => !s.Presente).Count();

                        float porcentajeSI = ((float)cantidadSI / (float)cantidadParticipantes) * 100;
                        float porcentajeNO = ((float)cantidadNO / (float)cantidadParticipantes) * 100;

                        string resultadoSI = string.Format("SI % {0}", Math.Round(porcentajeSI, 2));
                        string resultadoNO = string.Format("NO % {0}", Math.Round(porcentajeNO, 2));
                        string resultadoFinal = porcentajeNO > porcentajeSI ? "SI" : "NO";

                        for (int i = 1; i <= 3; i++)
                        {
                            ws.Cells[filaDatosParticipantes, i].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            ws.Cells[filaDatosParticipantes, i].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(128, 128, 128));
                            ws.Cells[filaDatosParticipantes, i].Style.Font.Color.SetColor(Color.White);
                            ws.Cells[filaDatosParticipantes, i].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                            ws.Cells[filaDatosParticipantes, i].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                            ws.Cells[filaDatosParticipantes, i].Style.Border.Top.Style = ExcelBorderStyle.Hair;
                            ws.Cells[filaDatosParticipantes, i].Style.Border.Left.Style = ExcelBorderStyle.Hair;
                            ws.Cells[filaDatosParticipantes, i].Style.Border.Right.Style = ExcelBorderStyle.Hair;
                            ws.Cells[filaDatosParticipantes, i].Style.Border.Bottom.Style = ExcelBorderStyle.Hair;

                            switch (i)
                            {
                                case 1:
                                    ws.Cells[filaDatosParticipantes, i].Value = "Asistencia";
                                    ws.Cells[filaDatosParticipantes, i].Style.Font.Bold = true;

                                    ws.Cells[filaDatosParticipantes, i].Style.Fill.BackgroundColor.SetColor(colorGrisOscuroEstiloPPM);
                                    ws.Cells[filaDatosParticipantes, i].Style.Font.Color.SetColor(colorBlancoEstiloPPM);

                                    filaDatosParticipantes++;

                                    ws.Cells[filaDatosParticipantes, i].Value = resultadoSI;
                                    ws.Cells[filaDatosParticipantes, i].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                                    ws.Cells[filaDatosParticipantes, i].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                                    // Assign borders
                                    ws.Cells[filaDatosParticipantes, i].Style.Border.Top.Style = ExcelBorderStyle.Hair;
                                    ws.Cells[filaDatosParticipantes, i].Style.Border.Left.Style = ExcelBorderStyle.Hair;
                                    ws.Cells[filaDatosParticipantes, i].Style.Border.Right.Style = ExcelBorderStyle.Hair;
                                    ws.Cells[filaDatosParticipantes, i].Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                                    filaDatosParticipantes++;

                                    ws.Cells[filaDatosParticipantes, i].Value = resultadoNO;
                                    ws.Cells[filaDatosParticipantes, i].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                                    ws.Cells[filaDatosParticipantes, i].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                    // Assign borders
                                    ws.Cells[filaDatosParticipantes, i].Style.Border.Top.Style = ExcelBorderStyle.Hair;
                                    ws.Cells[filaDatosParticipantes, i].Style.Border.Left.Style = ExcelBorderStyle.Hair;
                                    ws.Cells[filaDatosParticipantes, i].Style.Border.Right.Style = ExcelBorderStyle.Hair;
                                    ws.Cells[filaDatosParticipantes, i].Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                                    break;
                                case 2:

                                    ws.Cells[filaDatosParticipantes, i].Value = "¿Se suspende?";
                                    ws.Cells[filaDatosParticipantes, i].Style.Font.Bold = true;

                                    ws.Cells[filaDatosParticipantes, i].Style.Fill.BackgroundColor.SetColor(colorGrisOscuroEstiloPPM);
                                    ws.Cells[filaDatosParticipantes, i].Style.Font.Color.SetColor(colorBlancoEstiloPPM);

                                    filaDatosParticipantes++;
                                    ws.Cells[filaDatosParticipantes, i].Value = resultadoFinal;
                                    ws.Cells[filaDatosParticipantes, i].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                                    ws.Cells[filaDatosParticipantes, i].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                    // Assign borders
                                    ws.Cells[filaDatosParticipantes, i].Style.Border.Top.Style = ExcelBorderStyle.Hair;
                                    ws.Cells[filaDatosParticipantes, i].Style.Border.Left.Style = ExcelBorderStyle.Hair;
                                    ws.Cells[filaDatosParticipantes, i].Style.Border.Right.Style = ExcelBorderStyle.Hair;
                                    ws.Cells[filaDatosParticipantes, i].Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                                    break;
                                case 3:
                                    ws.Cells[filaDatosParticipantes, 3, filaDatosParticipantes, 9].Value = "Observaciones";
                                    ws.Cells[filaDatosParticipantes, 3, filaDatosParticipantes, 9].Merge = true;
                                    ws.Cells[filaDatosParticipantes, 3, filaDatosParticipantes, 9].Style.Font.Bold = true;

                                    ws.Cells[filaDatosParticipantes, 3, filaDatosParticipantes, 9].Style.Fill.BackgroundColor.SetColor(colorGrisOscuroEstiloPPM);
                                    ws.Cells[filaDatosParticipantes, 3, filaDatosParticipantes, 9].Style.Font.Color.SetColor(colorBlancoEstiloPPM);

                                    ws.Cells[filaDatosParticipantes, 3, filaDatosParticipantes, 9].Style.Border.Top.Style = ExcelBorderStyle.Hair;
                                    ws.Cells[filaDatosParticipantes, 3, filaDatosParticipantes, 9].Style.Border.Left.Style = ExcelBorderStyle.Hair;
                                    ws.Cells[filaDatosParticipantes, 3, filaDatosParticipantes, 9].Style.Border.Right.Style = ExcelBorderStyle.Hair;
                                    ws.Cells[filaDatosParticipantes, 3, filaDatosParticipantes, 9].Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                                    ws.Cells[filaDatosParticipantes, 3, filaDatosParticipantes, 9].Style.HorizontalAlignment = ExcelHorizontalAlignment.CenterContinuous;
                                    filaDatosParticipantes++;
                                    ws.Cells[filaDatosParticipantes, 3, filaDatosParticipantes, 9].Value = acta.Cabecera.Observaciones;
                                    ws.Cells[filaDatosParticipantes, 3, filaDatosParticipantes, 9].Merge = true;
                                    ws.Cells[filaDatosParticipantes, 3, filaDatosParticipantes, 9].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;

                                    // Assign borders
                                    ws.Cells[filaDatosParticipantes, 3, filaDatosParticipantes, 9].Style.Border.Top.Style = ExcelBorderStyle.Hair;
                                    ws.Cells[filaDatosParticipantes, 3, filaDatosParticipantes, 9].Style.Border.Left.Style = ExcelBorderStyle.Hair;
                                    ws.Cells[filaDatosParticipantes, 3, filaDatosParticipantes, 9].Style.Border.Right.Style = ExcelBorderStyle.Hair;
                                    ws.Cells[filaDatosParticipantes, 3, filaDatosParticipantes, 9].Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                                    ws.Cells[filaDatosParticipantes, 3, filaDatosParticipantes, 9].Style.WrapText = true;

                                    var caracteres = acta.Cabecera.Observaciones.Count();

                                    if (caracteres <= 104)
                                    {
                                        ws.Row(filaDatosParticipantes).Height = 20.25;
                                    }
                                    else if (caracteres <= 208)
                                    {
                                        ws.Row(filaDatosParticipantes).Height = 30;
                                    }
                                    else if (caracteres <= 312)
                                    {
                                        ws.Row(filaDatosParticipantes).Height = 40;
                                    }
                                    else if (caracteres <= 416)
                                    {
                                        ws.Row(filaDatosParticipantes).Height = 50;
                                    }
                                    else if (caracteres <= 500)
                                    {
                                        ws.Row(filaDatosParticipantes).Height = 60;
                                    }
                                    else
                                    {
                                        ws.Row(filaDatosParticipantes).Height = 70;
                                    }


                                    break;
                                default:
                                    ws.Cells[filaDatosParticipantes, i].Value = ":";
                                    ws.Cells[filaDatosParticipantes, i].Style.Font.Bold = true;
                                    filaDatosParticipantes++;
                                    ws.Cells[filaDatosParticipantes, i].Value = "";
                                    break;

                            }
                            ws.Row(finCabecera).Height = 20.25;
                            finCabecera++;
                            filaDatosParticipantes = inicio;
                            ws.Column(5).AutoFit();
                            ws.Column(6).AutoFit();

                        }
                        ws.Row(finCabecera).Height = 8.25;
                        finCabecera = finCabecera + 1;
                        ws.Column(4).Width = 14.30;

                    }
                    #endregion

                    #region Detalle Temas a tratar
                    if (acta.Cuerpo.Temas.Any())
                    {
                        col = 1;
                        int totalColumnasDetalleTemas = typeof(DetalleActaTemasTratar).GetProperties().Length;

                        for (int i = 1; i <= totalColumnasDetalleTemas; i++)
                        {

                            ws.Cells[finCabecera, col].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            ws.Cells[finCabecera, col].Style.Fill.BackgroundColor.SetColor(colorGrisClaro3EstiloPPM);
                            ws.Cells[finCabecera, col].Style.Font.Color.SetColor(colorGrisOscuroEstiloPPM);
                            ws.Cells[finCabecera, col].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                            ws.Cells[finCabecera, col].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            // Assign borders
                            ws.Cells[finCabecera, col].Style.Border.Top.Style = ExcelBorderStyle.Hair;
                            ws.Cells[finCabecera, col].Style.Border.Left.Style = ExcelBorderStyle.Hair;
                            ws.Cells[finCabecera, col].Style.Border.Right.Style = ExcelBorderStyle.Hair;
                            ws.Cells[finCabecera, col].Style.Border.Bottom.Style = ExcelBorderStyle.Hair;

                            switch (i)
                            {
                                case 1:
                                    ws.Cells[finCabecera, col++].Value = "N.";
                                    ws.Cells[finCabecera, 1, finCabecera, 1].Style.Border.Top.Style = ExcelBorderStyle.Hair;
                                    ws.Cells[finCabecera, 1, finCabecera, 1].Style.Border.Left.Style = ExcelBorderStyle.Hair;
                                    ws.Cells[finCabecera, 1, finCabecera, 1].Style.Border.Right.Style = ExcelBorderStyle.Hair;
                                    ws.Cells[finCabecera, 1, finCabecera, 1].Style.Border.Bottom.Style = ExcelBorderStyle.Hair;

                                    ws.Cells[finCabecera, 1, finCabecera, 1].Style.Fill.BackgroundColor.SetColor(colorGrisClaro3EstiloPPM);
                                    ws.Cells[finCabecera, 1, finCabecera, 1].Style.Font.Color.SetColor(colorGrisOscuroEstiloPPM);
                                    ws.Cells[finCabecera, 1, finCabecera, 1].Style.Font.Bold = true;
                                    break;
                                case 2:
                                    //ws.Column(col).Width2+ = 30;
                                    ws.Cells[finCabecera, col++].Value = "TEMAS O PUNTOS A TRATAR";
                                    ws.Cells[finCabecera, 2, finCabecera, 5].Merge = true;
                                    ws.Cells[finCabecera, 2, finCabecera, 5].Style.Fill.PatternType = ExcelFillStyle.Solid;

                                    ws.Cells[finCabecera, 2, finCabecera, 5].Style.Fill.BackgroundColor.SetColor(colorGrisClaro3EstiloPPM);
                                    ws.Cells[finCabecera, 2, finCabecera, 5].Style.Font.Color.SetColor(colorGrisOscuroEstiloPPM);
                                    ws.Cells[finCabecera, 2, finCabecera, 5].Style.Font.Bold = true;

                                    ws.Cells[finCabecera, 2, finCabecera, 5].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                                    ws.Cells[finCabecera, 2, finCabecera, 5].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                    // Assign borders
                                    ws.Cells[finCabecera, 2, finCabecera, 5].Style.Border.Top.Style = ExcelBorderStyle.Hair;
                                    ws.Cells[finCabecera, 2, finCabecera, 5].Style.Border.Left.Style = ExcelBorderStyle.Hair;
                                    ws.Cells[finCabecera, 2, finCabecera, 5].Style.Border.Right.Style = ExcelBorderStyle.Hair;
                                    ws.Cells[finCabecera, 2, finCabecera, 5].Style.Border.Bottom.Style = ExcelBorderStyle.Hair;


                                    break;
                                case 3:
                                    ws.Cells[finCabecera, 6].Value = "RESPONSABLE(S)";
                                    ws.Cells[finCabecera, 6, finCabecera, 9].Merge = true;
                                    ws.Cells[finCabecera, 6, finCabecera, 9].Style.Fill.PatternType = ExcelFillStyle.Solid;

                                    ws.Cells[finCabecera, 6, finCabecera, 9].Style.Fill.BackgroundColor.SetColor(colorGrisClaro3EstiloPPM);
                                    ws.Cells[finCabecera, 6, finCabecera, 9].Style.Font.Color.SetColor(colorGrisOscuroEstiloPPM);
                                    ws.Cells[finCabecera, 6, finCabecera, 9].Style.Font.Bold = true;

                                    ws.Cells[finCabecera, 6, finCabecera, 9].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                                    ws.Cells[finCabecera, 6, finCabecera, 9].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                    // Assign borders
                                    ws.Cells[finCabecera, 6, finCabecera, 9].Style.Border.Top.Style = ExcelBorderStyle.Hair;
                                    ws.Cells[finCabecera, 6, finCabecera, 9].Style.Border.Left.Style = ExcelBorderStyle.Hair;
                                    ws.Cells[finCabecera, 6, finCabecera, 9].Style.Border.Right.Style = ExcelBorderStyle.Hair;
                                    ws.Cells[finCabecera, 6, finCabecera, 9].Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                                    break;
                                default:
                                    ws.Cells[finCabecera, col++].Value = "";
                                    break;
                            }

                            ws.Row(finCabecera).Height = 20.25;
                        }

                        finCabecera++;

                        contador = 1;
                        foreach (var column in acta.Cuerpo.Temas)
                        {
                            col = 1;
                            for (int i = 1; i <= totalColumnasDetalleTemas; i++)
                            {
                                ws.Cells[finCabecera, col].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                                ws.Cells[finCabecera, col].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                // Assign borders
                                ws.Cells[finCabecera, col].Style.Border.Top.Style = ExcelBorderStyle.Hair;
                                ws.Cells[finCabecera, col].Style.Border.Left.Style = ExcelBorderStyle.Hair;
                                ws.Cells[finCabecera, col].Style.Border.Right.Style = ExcelBorderStyle.Hair;
                                ws.Cells[finCabecera, col].Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                                switch (i)
                                {
                                    case 1:
                                        ws.Cells[finCabecera, col++].Value = contador;
                                        ws.Cells[finCabecera, 1, finCabecera, 1].Merge = true;
                                        ws.Cells[finCabecera, 1, finCabecera, 1].Style.Border.Top.Style = ExcelBorderStyle.Hair;
                                        ws.Cells[finCabecera, 1, finCabecera, 1].Style.Border.Left.Style = ExcelBorderStyle.Hair;
                                        ws.Cells[finCabecera, 1, finCabecera, 1].Style.Border.Right.Style = ExcelBorderStyle.Hair;
                                        ws.Cells[finCabecera, 1, finCabecera, 1].Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                                        break;
                                    case 2:
                                        ws.Cells[finCabecera, col++].Value = column.Tema;
                                        ws.Cells[finCabecera, 2, finCabecera, 5].Merge = true;
                                        ws.Cells[finCabecera, 2, finCabecera, 5].Style.Border.Top.Style = ExcelBorderStyle.Hair;
                                        ws.Cells[finCabecera, 2, finCabecera, 5].Style.Border.Left.Style = ExcelBorderStyle.Hair;
                                        ws.Cells[finCabecera, 2, finCabecera, 5].Style.Border.Right.Style = ExcelBorderStyle.Hair;
                                        ws.Cells[finCabecera, 2, finCabecera, 5].Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                                        ws.Cells[finCabecera, 2, finCabecera, 5].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                                        ws.Cells[finCabecera, 2, finCabecera, 5].Style.WrapText = true;

                                        var caracteres = column.Tema.Count();

                                        if (caracteres <= 70)
                                        {
                                            ws.Row(finCabecera).Height = 20.25;
                                        }
                                        else if (caracteres <= 140)
                                        {
                                            ws.Row(finCabecera).Height = 40;
                                        }
                                        else if (caracteres <= 210)
                                        {
                                            ws.Row(finCabecera).Height = 55;
                                        }
                                        else if (caracteres <= 280)
                                        {
                                            ws.Row(finCabecera).Height = 70;
                                        }
                                        else if (caracteres <= 350)
                                        {
                                            ws.Row(finCabecera).Height = 85;
                                        }
                                        else if (caracteres <= 420)
                                        {
                                            ws.Row(finCabecera).Height = 100;
                                        }
                                        else if (caracteres <= 500)
                                        {
                                            ws.Row(finCabecera).Height = 115;
                                        }
                                        else
                                        {
                                            ws.Row(finCabecera).Height = 130;
                                        }
                                        break;
                                    case 3:
                                        ws.Cells[finCabecera, 6].Value = column.Responsable;

                                        ws.Cells[finCabecera, 6, finCabecera, 9].Merge = true;
                                        ws.Cells[finCabecera, 6, finCabecera, 9].Style.Border.Top.Style = ExcelBorderStyle.Hair;
                                        ws.Cells[finCabecera, 6, finCabecera, 9].Style.Border.Left.Style = ExcelBorderStyle.Hair;
                                        ws.Cells[finCabecera, 6, finCabecera, 9].Style.Border.Right.Style = ExcelBorderStyle.Hair;
                                        ws.Cells[finCabecera, 6, finCabecera, 9].Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                                        ws.Cells[finCabecera, 6, finCabecera, 9].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                                        ws.Cells[finCabecera, 6, finCabecera, 9].Style.WrapText = true;
                                        break;
                                    default:
                                        ws.Cells[finCabecera, col++].Value = "";
                                        break;
                                }
                            }

                            //ws.Row(finCabecera).Height = 39.5;
                            finCabecera++;
                            contador++;
                        }

                        ws.Row(finCabecera).Height = 8.25;
                        finCabecera += 1;


                    }
                    #endregion

                    #region Detalle Acuerdos
                    if (acta.Cuerpo.Acuerdos.Any())
                    {
                        col = 1;
                        int totalColumnasDetalleAcuerdos = typeof(DetalleActaAcuerdos).GetProperties().Length;

                        for (int i = 1; i <= totalColumnasDetalleAcuerdos; i++)
                        {
                            //ws.Column(col).Width = 18;

                            ws.Cells[finCabecera, col].Style.Fill.PatternType = ExcelFillStyle.Solid;

                            ws.Cells[finCabecera, col].Style.Fill.BackgroundColor.SetColor(colorGrisClaro3EstiloPPM);
                            ws.Cells[finCabecera, col].Style.Font.Color.SetColor(colorGrisOscuroEstiloPPM);

                            ws.Cells[finCabecera, col].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                            ws.Cells[finCabecera, col].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            // Assign borders
                            ws.Cells[finCabecera, col].Style.Border.Top.Style = ExcelBorderStyle.Hair;
                            ws.Cells[finCabecera, col].Style.Border.Left.Style = ExcelBorderStyle.Hair;
                            ws.Cells[finCabecera, col].Style.Border.Right.Style = ExcelBorderStyle.Hair;
                            ws.Cells[finCabecera, col].Style.Border.Bottom.Style = ExcelBorderStyle.Hair;

                            switch (i)
                            {
                                case 1:
                                    ws.Cells[finCabecera, col++].Value = "N.";
                                    ws.Cells[finCabecera, 1, finCabecera, 1].Style.Border.Top.Style = ExcelBorderStyle.Hair;
                                    ws.Cells[finCabecera, 1, finCabecera, 1].Style.Border.Left.Style = ExcelBorderStyle.Hair;
                                    ws.Cells[finCabecera, 1, finCabecera, 1].Style.Border.Right.Style = ExcelBorderStyle.Hair;
                                    ws.Cells[finCabecera, 1, finCabecera, 1].Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                                    ws.Cells[finCabecera, 1, finCabecera, 1].Style.Font.Bold = true;

                                    ws.Cells[finCabecera, 1, finCabecera, 1].Style.Fill.BackgroundColor.SetColor(colorGrisClaro3EstiloPPM);
                                    ws.Cells[finCabecera, 1, finCabecera, 1].Style.Font.Color.SetColor(colorGrisOscuroEstiloPPM);
                                    break;
                                case 2:
                                    ws.Cells[finCabecera, col++].Value = "ACUERDOS";
                                    ws.Cells[finCabecera, 2, finCabecera, 5].Merge = true;
                                    ws.Cells[finCabecera, 2, finCabecera, 5].Style.Fill.PatternType = ExcelFillStyle.Solid;

                                    ws.Cells[finCabecera, 2, finCabecera, 5].Style.Fill.BackgroundColor.SetColor(colorGrisClaro3EstiloPPM);
                                    ws.Cells[finCabecera, 2, finCabecera, 5].Style.Font.Color.SetColor(colorGrisOscuroEstiloPPM);

                                    ws.Cells[finCabecera, 2, finCabecera, 5].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                                    ws.Cells[finCabecera, 2, finCabecera, 5].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                    // Assign borders
                                    ws.Cells[finCabecera, 2, finCabecera, 5].Style.Border.Top.Style = ExcelBorderStyle.Hair;
                                    ws.Cells[finCabecera, 2, finCabecera, 5].Style.Border.Left.Style = ExcelBorderStyle.Hair;
                                    ws.Cells[finCabecera, 2, finCabecera, 5].Style.Border.Right.Style = ExcelBorderStyle.Hair;
                                    ws.Cells[finCabecera, 2, finCabecera, 5].Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                                    ws.Cells[finCabecera, 2, finCabecera, 5].Style.Font.Bold = true;
                                    break;

                                case 3:
                                    ws.Cells[finCabecera, 6].Value = "RESPONSABLES";
                                    ws.Cells[finCabecera, 6, finCabecera, 7].Merge = true;
                                    ws.Cells[finCabecera, 6, finCabecera, 7].Style.Fill.PatternType = ExcelFillStyle.Solid;

                                    ws.Cells[finCabecera, 6, finCabecera, 7].Style.Fill.BackgroundColor.SetColor(colorGrisClaro3EstiloPPM);
                                    ws.Cells[finCabecera, 6, finCabecera, 7].Style.Font.Color.SetColor(colorGrisOscuroEstiloPPM);

                                    ws.Cells[finCabecera, 6, finCabecera, 7].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                                    ws.Cells[finCabecera, 6, finCabecera, 7].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                    // Assign borders
                                    ws.Cells[finCabecera, 6, finCabecera, 7].Style.Border.Top.Style = ExcelBorderStyle.Hair;
                                    ws.Cells[finCabecera, 6, finCabecera, 7].Style.Border.Left.Style = ExcelBorderStyle.Hair;
                                    ws.Cells[finCabecera, 6, finCabecera, 7].Style.Border.Right.Style = ExcelBorderStyle.Hair;
                                    ws.Cells[finCabecera, 6, finCabecera, 7].Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                                    ws.Cells[finCabecera, 6, finCabecera, 7].Style.Font.Bold = true;
                                    break;
                                case 4:
                                    ws.Cells[finCabecera, 8].Value = "FECHAS";
                                    ws.Cells[finCabecera, 8, finCabecera, 9].Merge = true;
                                    ws.Cells[finCabecera, 8, finCabecera, 9].Style.Fill.PatternType = ExcelFillStyle.Solid;

                                    ws.Cells[finCabecera, 8, finCabecera, 9].Style.Fill.BackgroundColor.SetColor(colorGrisClaro3EstiloPPM);
                                    ws.Cells[finCabecera, 8, finCabecera, 9].Style.Font.Color.SetColor(colorGrisOscuroEstiloPPM);

                                    ws.Cells[finCabecera, 8, finCabecera, 9].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                                    ws.Cells[finCabecera, 8, finCabecera, 9].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                    // Assign borders
                                    ws.Cells[finCabecera, 8, finCabecera, 9].Style.Border.Top.Style = ExcelBorderStyle.Hair;
                                    ws.Cells[finCabecera, 8, finCabecera, 9].Style.Border.Left.Style = ExcelBorderStyle.Hair;
                                    ws.Cells[finCabecera, 8, finCabecera, 9].Style.Border.Right.Style = ExcelBorderStyle.Hair;
                                    ws.Cells[finCabecera, 8, finCabecera, 9].Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                                    ws.Cells[finCabecera, 8, finCabecera, 9].Style.Font.Bold = true;

                                    break;
                                default:
                                    ws.Cells[finCabecera, col++].Value = "";
                                    break;
                            }

                            ws.Row(finCabecera).Height = 20.25;
                        }

                        finCabecera++;

                        contador = 1;
                        foreach (var column in acta.Cuerpo.Acuerdos)
                        {
                            col = 1;
                            for (int i = 1; i <= totalColumnasDetalleAcuerdos; i++)
                            {
                                ws.Cells[finCabecera, col].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                                ws.Cells[finCabecera, col].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                // Assign borders
                                ws.Cells[finCabecera, col].Style.Border.Top.Style = ExcelBorderStyle.Hair;
                                ws.Cells[finCabecera, col].Style.Border.Left.Style = ExcelBorderStyle.Hair;
                                ws.Cells[finCabecera, col].Style.Border.Right.Style = ExcelBorderStyle.Hair;
                                ws.Cells[finCabecera, col].Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                                switch (i)
                                {
                                    case 1:

                                        ws.Cells[finCabecera, col++].Value = contador;
                                        ws.Cells[finCabecera, 1, finCabecera, 1].Style.Border.Top.Style = ExcelBorderStyle.Hair;
                                        ws.Cells[finCabecera, 1, finCabecera, 1].Style.Border.Left.Style = ExcelBorderStyle.Hair;
                                        ws.Cells[finCabecera, 1, finCabecera, 1].Style.Border.Right.Style = ExcelBorderStyle.Hair;
                                        ws.Cells[finCabecera, 1, finCabecera, 1].Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                                        break;
                                    case 2:
                                        ws.Cells[finCabecera, col++].Value = column.Acuerdo;
                                        ws.Cells[finCabecera, 2, finCabecera, 5].Merge = true;
                                        ws.Cells[finCabecera, 2, finCabecera, 5].Style.Border.Top.Style = ExcelBorderStyle.Hair;
                                        ws.Cells[finCabecera, 2, finCabecera, 5].Style.Border.Left.Style = ExcelBorderStyle.Hair;
                                        ws.Cells[finCabecera, 2, finCabecera, 5].Style.Border.Right.Style = ExcelBorderStyle.Hair;
                                        ws.Cells[finCabecera, 2, finCabecera, 5].Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                                        ws.Cells[finCabecera, 2, finCabecera, 5].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                                        ws.Cells[finCabecera, 2, finCabecera, 5].Style.WrapText = true;
                                        var caracteres = column.Acuerdo.Count();
                                        if (caracteres <= 70)
                                        {
                                            ws.Row(finCabecera).Height = 20.25;
                                        }
                                        else if (caracteres <= 140)
                                        {
                                            ws.Row(finCabecera).Height = 40;
                                        }
                                        else if (caracteres <= 210)
                                        {
                                            ws.Row(finCabecera).Height = 55;
                                        }
                                        else if (caracteres <= 280)
                                        {
                                            ws.Row(finCabecera).Height = 70;
                                        }
                                        else if (caracteres <= 350)
                                        {
                                            ws.Row(finCabecera).Height = 85;
                                        }
                                        else if (caracteres <= 420)
                                        {
                                            ws.Row(finCabecera).Height = 100;
                                        }
                                        else if (caracteres <= 500)
                                        {
                                            ws.Row(finCabecera).Height = 115;
                                        }
                                        else
                                        {
                                            ws.Row(finCabecera).Height = 130;
                                        }

                                        break;
                                    case 3:
                                        ws.Cells[finCabecera, 6].Value = column.Responsable;
                                        ws.Cells[finCabecera, 6, finCabecera, 7].Merge = true;
                                        ws.Cells[finCabecera, 6, finCabecera, 7].Style.Border.Top.Style = ExcelBorderStyle.Hair;
                                        ws.Cells[finCabecera, 6, finCabecera, 7].Style.Border.Left.Style = ExcelBorderStyle.Hair;
                                        ws.Cells[finCabecera, 6, finCabecera, 7].Style.Border.Right.Style = ExcelBorderStyle.Hair;
                                        ws.Cells[finCabecera, 6, finCabecera, 7].Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                                        ws.Cells[finCabecera, 6, finCabecera, 7].Style.WrapText = true;
                                        break;
                                    case 4:
                                        ws.Cells[finCabecera, 8].Value = column.Fecha.ToString("yyyy/MM/dd");
                                        ws.Cells[finCabecera, 8, finCabecera, 9].Merge = true;
                                        ws.Cells[finCabecera, 8, finCabecera, 9].Style.Border.Top.Style = ExcelBorderStyle.Hair;
                                        ws.Cells[finCabecera, 8, finCabecera, 9].Style.Border.Left.Style = ExcelBorderStyle.Hair;
                                        ws.Cells[finCabecera, 8, finCabecera, 9].Style.Border.Right.Style = ExcelBorderStyle.Hair;
                                        ws.Cells[finCabecera, 8, finCabecera, 9].Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                                        ws.Cells[finCabecera, 8, finCabecera, 9].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                        break;
                                    default:
                                        ws.Cells[finCabecera, col++].Value = "";
                                        break;
                                }

                            }


                            finCabecera++;
                            contador++;

                        }

                        finCabecera += 2;
                    }
                    #endregion

                    //Solo Acta de tipo Cliente
                    if (acta.Cuerpo.DetalleCliente.Any() || acta.Cuerpo.DetalleContabilidad.Any())
                    {
                        ws.Column(2).Width = 25;
                        ws.Column(3).Width = 15;

                        ws.Column(4).Width = 15;
                        ws.Column(5).Width = 30;

                        ws.Column(6).Width = 10;
                        ws.Column(7).Width = 25;
                        ws.Column(8).Width = 17;
                        ws.Column(9).Width = 20;

                        ws.Protection.IsProtected = true;
                    }
                    else
                    {
                        ws.Column(2).Width = 18;
                        ws.Column(3).Width = 20;
                        ws.Column(6).Width = 15;
                        ws.Column(7).Width = 20;
                        ws.Column(8).Width = 12;
                        ws.Column(9).Width = 15;
                    }

                    #region Pie de Pagina 
                    if (!string.IsNullOrEmpty(acta.PiePagina.AcuerdoConformidad))
                    {

                        var firmas = JsonConvert.DeserializeObject<List<FirmasActaParcial>>(acta.PiePagina.Firmas);

                        ws.Cells[finCabecera, 1].Value = "A  C  U  E  R  D  O   D  E   C  O  N  F  O  R  M  I  D  A  D";

                        ws.Cells[finCabecera, 1].Style.Fill.BackgroundColor.SetColor(colorGrisOscuroEstiloPPM);
                        ws.Cells[finCabecera, 1].Style.Font.Color.SetColor(colorBlancoEstiloPPM);
                        ws.Cells[finCabecera, 1, finCabecera, 9].Merge = true;
                        ws.Cells[finCabecera, 1, finCabecera, 9].Style.HorizontalAlignment = ExcelHorizontalAlignment.CenterContinuous;
                        ws.Cells[finCabecera, 1, finCabecera, 9].Style.Border.Top.Style = ExcelBorderStyle.Hair;
                        ws.Cells[finCabecera, 1, finCabecera, 9].Style.Border.Left.Style = ExcelBorderStyle.Hair;
                        ws.Cells[finCabecera, 1, finCabecera, 9].Style.Border.Right.Style = ExcelBorderStyle.Hair;
                        ws.Cells[finCabecera, 1, finCabecera, 9].Style.Border.Bottom.Style = ExcelBorderStyle.Hair;

                        ws.Row(finCabecera).Height = 20.25;

                        int finBorder = finCabecera + 12;

                        ws.Cells[finCabecera, 1, finBorder, 9].Style.Border.BorderAround(ExcelBorderStyle.Hair);

                        finCabecera += 2;

                        // Solo para cualquier acta que no sea la de CLiente o de Contabilidad
                        if (!acta.Cuerpo.DetalleCliente.Any() && !acta.Cuerpo.DetalleContabilidad.Any())
                        {
                            ws.Cells[finCabecera, 1].Value = acta.PiePagina.AcuerdoConformidad;
                            ws.Cells[finCabecera, 1, finCabecera, 9].Merge = true;
                            ws.Cells[finCabecera, 1, finCabecera, 9].Style.HorizontalAlignment = ExcelHorizontalAlignment.CenterContinuous;
                        }

                        finCabecera += 7;

                        int inicioFilaFirma = finCabecera;
                        int auxiliar = inicioFilaFirma;


                        for (int i = 3; i <= 7; i++)
                        {
                            if (i == 3)
                            {
                                ws.Cells[inicioFilaFirma, i].Value = firmas.ElementAtOrDefault(0).nombre;
                                ws.Cells[inicioFilaFirma, i].Style.Border.Top.Style = ExcelBorderStyle.Hair;
                                ws.Cells[inicioFilaFirma, i].Style.HorizontalAlignment = ExcelHorizontalAlignment.CenterContinuous;
                                inicioFilaFirma++;
                                ws.Cells[inicioFilaFirma, i].Value = firmas.ElementAtOrDefault(0).cargo;
                                ws.Cells[inicioFilaFirma, i].Style.Font.Bold = true;
                                ws.Cells[inicioFilaFirma, i].Style.HorizontalAlignment = ExcelHorizontalAlignment.CenterContinuous;
                                inicioFilaFirma++;
                                ws.Cells[inicioFilaFirma, i].Value = firmas.ElementAtOrDefault(0).empresa;
                                ws.Cells[inicioFilaFirma, i].Style.Font.Bold = true;
                                ws.Cells[inicioFilaFirma, i].Style.HorizontalAlignment = ExcelHorizontalAlignment.CenterContinuous;
                            }
                            if (i == 7)
                            {
                                ws.Cells[inicioFilaFirma, i].Value = firmas.ElementAtOrDefault(1).usuarioNombre;
                                ws.Cells[inicioFilaFirma, i].Style.Border.Top.Style = ExcelBorderStyle.Hair;
                                ws.Cells[inicioFilaFirma, i].Style.HorizontalAlignment = ExcelHorizontalAlignment.CenterContinuous;
                                inicioFilaFirma++;
                                ws.Cells[inicioFilaFirma, i].Value = firmas.ElementAtOrDefault(1).usuarioCargo;
                                ws.Cells[inicioFilaFirma, i].Style.Font.Bold = true;
                                ws.Cells[inicioFilaFirma, i].Style.HorizontalAlignment = ExcelHorizontalAlignment.CenterContinuous;
                                inicioFilaFirma++;
                                ws.Cells[inicioFilaFirma, i].Value = firmas.ElementAtOrDefault(1).usuarioEmpresa;
                                ws.Cells[inicioFilaFirma, i].Style.Font.Bold = true;
                                ws.Cells[inicioFilaFirma, i].Style.HorizontalAlignment = ExcelHorizontalAlignment.CenterContinuous;
                            }
                            inicioFilaFirma = auxiliar;
                            finCabecera++;
                        }

                        if (acta.Cuerpo.DetalleCliente.Any() || acta.Cuerpo.DetalleContabilidad.Any())
                        {
                            ws.Column(8).Width = 17;
                            ws.Column(9).Width = 20;
                        }
                        else
                        {
                            ws.Column(8).Width = 15;
                            ws.Column(9).Width = 20;
                        }
                    }
                    //FORMATO FUENTE TEXTO DE TODO EL DOCUMENTO
                    using (var range = ws.Cells[3, 1, finCabecera, columnaFinalDocumentoExcel])
                    {
                        range.Style.Font.Size = 10;
                        range.Style.Font.Name = "Raleway";
                        range.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    }

                    #endregion

                    //string basePath = ConfigurationManager.AppSettings["RepositorioDocumentos"];
                    //string rutaArchivos = basePath + "\\GESTION_PPM\\ACTAS";

                    //string codigoCotizacion = !string.IsNullOrEmpty(acta.Cabecera.CodigoCotizacion) ? acta.Cabecera.CodigoCotizacion : "SinCodigo" + id.ToString();

                    //var anioActual = DateTime.Now.Year.ToString();
                    //var almacenFisicoTemporal = Auxiliares.CrearCarpetasDirectorio(rutaArchivos, new List<string>() { anioActual, codigoCotizacion });

                    //var almacenFisicoOfficeToPDF = Auxiliares.CrearCarpetasDirectorio(Server.MapPath("~/OfficeToPDF/"), new List<string>());

                    //// Get the complete folder path and store the file inside it.    
                    //string pathExcel = Path.Combine(almacenFisicoTemporal, "ReporteActaExcel" + acta.Cabecera.CodigoActa + ".xlsx");

                    ////Write the file to the disk
                    //FileInfo fi = new FileInfo(pathExcel);
                    //package.SaveAs(fi);

                    //string pathExe = Path.Combine(almacenFisicoOfficeToPDF, "officetopdf.exe");
                    //string pathPDF = Path.Combine(almacenFisicoTemporal, "ReporteActaPDF" + acta.Cabecera.CodigoActa + ".pdf");

                    //object[] args = new object[] { "officetopdf.exe", pathExcel, pathPDF };

                    //string comando = String.Format("{0} {1} {2}", args);

                    //string comandoUbicarseEnRaiz = "cd " + Path.GetDirectoryName(pathExe);

                    //Log.Info(comandoUbicarseEnRaiz);
                    //Log.Info(comando);

                    //List<string> comandos = new List<string> { comandoUbicarseEnRaiz, comando }; // "OfficeToPDF.exe reporte.xlsx report.pdf"

                    //string archivoExe = Server.MapPath("~/OfficeToPDF/OfficeToPDF.exe");

                    //Auxiliares.EjecutarProcesosCMD(comandos, new List<string> { pathExcel, pathPDF }, archivoExe);

                }

                return package;
            }
            catch (Exception ex)
            {
                return package;
            }
        }

        public List<string> GetPathsArchivosActas(List<int> ids)
        {
            List<string> archivosPaths = new List<string>();
            try
            {
                foreach (var id in ids)
                {
                    var informacionCompleta = ActaEntity.ConsultarActaInformacionCompleta(id);
                    ActaCompleta acta = new ActaCompleta(informacionCompleta);

                    string basePath = ConfigurationManager.AppSettings["RepositorioDocumentos"];
                    string rutaArchivos = basePath + "\\GESTION_PPM\\ACTAS";

                    string codigoCotizacion = !string.IsNullOrEmpty(acta.Cabecera.CodigoCotizacion) ? acta.Cabecera.CodigoCotizacion : "SinCodigo" + id.ToString();
                    //string codigoCotizacion = !string.IsNullOrEmpty(acta.Cabecera.CodigoCotizacion) ? acta.Cabecera.CodigoCotizacion : "SinCodigo" + Guid.NewGuid().ToString().Substring(0, 8);

                    var anioActual = DateTime.Now.Year.ToString();
                    var almacenFisicoTemporal = Auxiliares.CrearCarpetasDirectorio(rutaArchivos, new List<string>() { anioActual, codigoCotizacion });

                    // Get the complete folder path and store the file inside it.    
                    //string pathExcel = Path.Combine(almacenFisicoTemporal, "ReporteActaExcel" + acta.Cabecera.CodigoActa + ".xlsx");
                    string pathPDF = Path.Combine(almacenFisicoTemporal, "ReporteActaPDF" + acta.Cabecera.CodigoActa + ".pdf");

                    archivosPaths.Add(pathPDF);
                }
                return archivosPaths;
            }
            catch (Exception ex)
            {
                return archivosPaths;
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

        #endregion

        [HttpGet]
        public ActionResult DescargarReporteFormatoExcel()
        {
            // Using EPPlus from nuget
            using (ExcelPackage package = new ExcelPackage())
            {
                Int32 row = 2;
                Int32 col = 1;

                package.Workbook.Worksheets.Add("Data");
                IGrid<ActasInformacionGeneralInfo> grid = CreateExportableGrid();
                ExcelWorksheet sheet = package.Workbook.Worksheets["Data"];

                foreach (IGridColumn column in grid.Columns)
                {
                    sheet.Cells[1, col].Value = column.Title;
                    sheet.Column(col++).Width = 18;

                    column.IsEncoded = false;
                }

                foreach (IGridRow<ActasInformacionGeneralInfo> gridRow in grid.Rows)
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

                return File(package.GetAsByteArray(), "application/unknown", "ListadoActas.xlsx");
            }
        }

        public IGrid<ActasInformacionGeneralInfo> CreateExportableGrid()
        {
            IGrid<ActasInformacionGeneralInfo> grid = new Grid<ActasInformacionGeneralInfo>(ActaEntity.ListadoActa());
            grid.ViewContext = new ViewContext { HttpContext = HttpContext };
            grid.Query = Request.QueryString;

            grid.Columns.Add(model => model.NombreTipoActa).Titled("Tipo de Acta").AppendCss("celda-grande");
            grid.Columns.Add(model => model.CodigoActa).Titled("Código de Acta").AppendCss("celda-mediana");
            grid.Columns.Add(model => model.FechaCreacion).Titled("Fecha de Creación").Formatted("{0:d}").AppendCss("celda-grande");
            grid.Columns.Add(model => model.NombresElaboradoPor).Titled("Autor").AppendCss("celda-grande");
            grid.Columns.Add(model => model.CodigoCotizacion).Titled("Código Cotización").AppendCss("celda-grande");
            grid.Columns.Add(model => model.PreFacturasSAFI).Titled("Prefacturas");
            foreach (IGridColumn column in grid.Columns)
            {
                column.Filter.IsEnabled = true;
                column.Sort.IsEnabled = true;
            }

            return grid;
        }

        public ActionResult DescargarReporteFormatoCSV()
        {
            var comlumHeadrs = new string[]
            {

                "TIPO DE ACTA",
                "CÓDIGO DE ACTA",
                "FECHA DE CREACIÓN",
                "AUTOR",
                "CÓDIGO DE COTIZACIÓN",
                "PREFACTURAS",
            };

            var listado = (from item in ActaEntity.ListadoActa()
                           select new object[]
                           {
                                            item.NombreTipoActa,
                                            item.CodigoActa,
                                              item.FechaCreacion.ToString("yyyy-MM-dd"),
                                               item.NombresElaboradoPor,
                                                item.CodigoCotizacion,
                                                 item.PreFacturasSAFI,
                           }).ToList();

            // Build the file content
            var employeecsv = new StringBuilder();
            listado.ForEach(line =>
            {
                employeecsv.AppendLine(string.Join(",", line));
            });

            byte[] buffer = Encoding.Default.GetBytes($"{string.Join(",", comlumHeadrs)}\r\n{employeecsv.ToString()}");
            return File(buffer, "text/csv", $"ListadoActas.csv");
        }

        public ActionResult DescargarReporteFormatoPDF()
        {
            // Seleccionar las columnas a exportar
            var results = ActaEntity.ListadoActa();
            var list = Reportes.SerializeToJSON(results);
            return Content(list, "application/json");
        }

        public ActionResult GeneracionPrefactura(string listadoIDs, bool descargaDirecta = false)
        {

            try
            {
                return Json(new { Resultado = new RespuestaTransaccion { Estado = true, Respuesta = Mensajes.MensajeTransaccionExitosa } }, JsonRequestBehavior.AllowGet);

            }
            catch (Exception ex)
            {
                return Json(new { Resultado = new RespuestaTransaccion { Estado = false, Respuesta = ex.Message } }, JsonRequestBehavior.AllowGet);
            }


        }
    }
}

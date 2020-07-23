using GestionPPM.Entidades.Modelo;
using System.IO;
using iTextSharp.text;
using iTextSharp.text.pdf;
using System;
using System.Linq;
using System.Threading.Tasks;
using System.Web.Mvc;
using System.Collections.Generic;
using GestionPPM.Entidades.Metodos;
using GestionPPM.Repositorios;
using System.Net;
using TemplateInicial.Helper;
using System.Configuration;
using System.Text;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using Seguridad.Helper;

namespace TemplateInicial.Controllers
{
    [Autenticado]
    public class CotizacionController : BaseAppController
    {

        GestionPPMEntities db = new GestionPPMEntities();
        private const string XlsxContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
        private object basePath;

        // GET: Cotizacion
        public ActionResult Index(string codigoCotizacion)
        {
            var respuesta = System.Web.HttpContext.Current.Session["MensajeVersionCotizacionExistente"] as string;
            ViewBag.CodigoCotizacion = codigoCotizacion;
            ViewBag.Resultado = respuesta;

            Session["MensajeVersionCotizacionExistente"] = "";
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

        public ActionResult _DetalleCotizacion(int codigoEntregable, string entregable, string idTarifarioConGestion_modal, int? cantidadTarifarioConGestion_modal, decimal? valorUnitarioTarifarioConGestion_modal, 
            decimal? costoUnitarioTarifarioConGestion_modal, string idTarifarioSinGestion_modal, int? cantidadTarifarioSinGestion_modal, decimal? valorUnitarioTarifarioSinGestion_modal, 
            decimal? costoUnitarioTarifarioSinGestion_modal, bool soloVisualizacion)
        {
             
            ViewBag.TituloModal = "Detalles de Cotización";

            if(soloVisualizacion == true)
            {
                //codigo del entregable
                ViewBag.codigoEntregable_modal = codigoEntregable;

                ViewBag.entregable = entregable;

                ViewBag.idTarifarioConGestion_modal = string.IsNullOrEmpty(idTarifarioConGestion_modal) ? "" : idTarifarioConGestion_modal;
                ViewBag.idTarifarioSinGestion_modal = !string.IsNullOrEmpty(idTarifarioSinGestion_modal) ? idTarifarioSinGestion_modal : "";

                ViewBag.cantidadTarifarioConGestion_modal = cantidadTarifarioConGestion_modal.HasValue ? cantidadTarifarioConGestion_modal.Value : 0;
                ViewBag.cantidadTarifarioSinGestion_modal = cantidadTarifarioSinGestion_modal.HasValue ? cantidadTarifarioSinGestion_modal.Value : 0;

                ViewBag.valorUnitarioTarifarioConGestion_modal = valorUnitarioTarifarioConGestion_modal.HasValue ? valorUnitarioTarifarioConGestion_modal.Value.ToString() : "0";
                ViewBag.costoUnitarioTarifarioConGestion_modal = costoUnitarioTarifarioConGestion_modal.HasValue ? costoUnitarioTarifarioConGestion_modal.Value.ToString() : "0";

                ViewBag.valorUnitarioTarifarioSinGestion_modal = valorUnitarioTarifarioSinGestion_modal.HasValue ? valorUnitarioTarifarioSinGestion_modal.Value.ToString() : "0";
                ViewBag.costoUnitarioTarifarioSinGestion_modal = costoUnitarioTarifarioSinGestion_modal.HasValue ? costoUnitarioTarifarioSinGestion_modal.Value.ToString() : "0";
            }
            else
            {
                //codigo del entregable
                ViewBag.codigoEntregable_modal = 0;

                ViewBag.entregable = "";

                //inicializa listas desplegables
                ViewBag.idTarifarioConGestion_modal = "";
                ViewBag.idTarifarioSinGestion_modal = "";

                //Inicializa la cantidad
                ViewBag.cantidadTarifarioConGestion_modal = 0;
                ViewBag.cantidadTarifarioSinGestion_modal = 0;

                //Inicializa valores
                ViewBag.valorUnitarioTarifarioConGestion_modal = 0;
                ViewBag.costoUnitarioTarifarioConGestion_modal = 0;

                //Inicializa totales
                ViewBag.valorUnitarioTarifarioSinGestion_modal = 0;
                ViewBag.costoUnitarioTarifarioSinGestion_modal = 0;
            }

            ViewBag.visualizacion = soloVisualizacion;

            return PartialView();
        }


        [HttpGet]
        public async Task<PartialViewResult> IndexGrid(String search)
        {
            ViewBag.NombreListado = Etiquetas.TituloGridCotizacion;

            //Controlar permisos
            var user = ViewData["usuario"] = System.Web.HttpContext.Current.Session["usuario"];
            var usuario = int.Parse(user.ToString());
            string nombreControlador = ControllerContext.RouteData.Values["controller"].ToString();
            ViewBag.NombreControlador = nombreControlador;

            ViewBag.AccionesUsuario = ManejoPermisosEntity.ListadoAccionesCatalogoUsuario(usuario, nombreControlador);

            //Obtener Acciones del controlador
            ViewBag.AccionesControlador = GetMetodosControlador(nombreControlador);

            //Búsqueda
            var listado = CotizacionEntity.ListarCabecerasCotizacion().ToList();

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

        // GET: Cotizacion/Create
        public ActionResult Create(int? idCodigoCotizacion)
        {
            //ViewBag.CabeceraCotizacion = CotizacionEntity.ConsultarCotizadorCabecera(idCodigoCotizacion.Value);
            if (idCodigoCotizacion.HasValue)
            {
                var validacion = CotizacionEntity.ConsultarCotizadorCabeceraByCodigoCotizacion(idCodigoCotizacion.Value);

                if (validacion)
                {
                    //Almacenar en una variable de sesion
                    Session["MensajeVersionCotizacionExistente"] = "¡Ya existe una versión previa del Código de Cotización!";
                    //return View("Index");
                    return RedirectToAction("Index");
                }
                else
                { 
                    ViewBag.CabeceraCotizacion = CodigoCotizacionEntity.ConsultarCodigoCotizacionCompleto(idCodigoCotizacion.Value);

                    //Obtener el porcentaje por default 
                    var impuesto= ImpuestoEntity.ObtenerListadoImpuestos("1").FirstOrDefault();
                    var valorIva = ImpuestoEntity.ConsultarImpuesto(Convert.ToInt32(impuesto.Value));
                    ViewBag.porcentajeIva = valorIva.valor;

                    return View();
                }
            }
            else
            {
                return RedirectToAction("Index", "Home");
            }
        }

        // POST: Cotizacion/Create
        // To protect from overposting attacks, please enable the specific properties you want to bind to, for 
        // more details see https://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        public ActionResult Create(Cotizador cotizadorCabecera, List<DetalleCotizador> cotizadorDetalles, string aplicaContratoCodigoCotizacion, decimal? subtotal)
        {
            try
            {
                //if (!CodigoCotizacionEntity.ValidarPagosCodigoCotizacion(codigoCotizacion))
                //    return Json(new { Resultado = new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeValidacionPagosCodigoCotizacion } }, JsonRequestBehavior.AllowGet);
                bool aplicaContrato = Boolean.Parse(aplicaContratoCodigoCotizacion);

                cotizadorCabecera.subtotal = subtotal;

                RespuestaTransaccion resultado = CotizacionEntity.CrearCotizador(cotizadorCabecera, cotizadorDetalles, aplicaContrato);
                return Json(new { Resultado = resultado }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                return Json(new { Resultado = new RespuestaTransaccion { Estado = false, Respuesta = ex.Message } }, JsonRequestBehavior.AllowGet);
            }
        }

        // GET: Cotizacion/Edit/5
        public ActionResult Edit(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            var cotizador = CotizacionEntity.ConsultarCotizadorCabecera(id.Value);

            ViewBag.CabeceraCotizacion = CodigoCotizacionEntity.ConsultarCodigoCotizacionCompleto(cotizador.id_codigo_cotizacion.Value);

            ViewBag.listadoDetalleCotizador = CotizacionEntity.ConsultarCotizadorDetalle(id.Value);

            ViewBag.MaximoIDlistadoDetalleCotizador = CotizacionEntity.ConsultarCotizadorDetalle(id.Value).Max(s=> s.id_detalle_cotizador);

            //Obtener el porcentaje por default  
            ViewBag.porcentajeIva = cotizador.porcentaje_iva;

            if (cotizador == null)
            {
                return HttpNotFound();
            }
            return View(cotizador);
        }

        // POST: Cotizacion/Edit/5
        // To protect from overposting attacks, please enable the specific properties you want to bind to, for 
        // more details see https://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        public ActionResult Edit(Cotizador cotizadorCabecera, List<DetalleCotizador> cotizadorDetalles, string aplicaContratoCodigoCotizacion, decimal? subtotal)
        {
            try
            {
                RespuestaTransaccion resultado = new RespuestaTransaccion();

                bool aplicaContrato = Boolean.Parse(aplicaContratoCodigoCotizacion);

                cotizadorCabecera.subtotal = subtotal;

                //validar impuesto activo
                if(ImpuestoEntity.ValidarImpuestoActivo(cotizadorCabecera.id_impuesto.Value)==false)
                {
                    resultado = new RespuestaTransaccion { Estado = false, Respuesta = "El impuesto seleccionado no está vigente" };
                    return Json(new { Resultado = resultado }, JsonRequestBehavior.AllowGet);
                }
                else
                {
                    if (!cotizadorCabecera.estado_cotizador.Value)
                        resultado = CotizacionEntity.CrearCotizador(cotizadorCabecera, cotizadorDetalles, aplicaContrato);
                    else
                        resultado = CotizacionEntity.ActualizarCotizador(cotizadorCabecera, cotizadorDetalles, aplicaContrato);

                    return Json(new { Resultado = resultado }, JsonRequestBehavior.AllowGet);
                }                
            }
            catch (Exception ex)
            {
                return Json(new { Resultado = new RespuestaTransaccion { Estado = false, Respuesta = ex.Message } }, JsonRequestBehavior.AllowGet);
            }
        }

        public ActionResult GetTarifario(int id)
        {
            return Json(new { DataTarifario = TarifarioEntity.ConsultarTarifario(id)}, JsonRequestBehavior.AllowGet);
        }

        public ActionResult GeneraCotizadorPDF(int id)
        {  
            CotizadorPDF(id);

            CotizacionEntity.CambiarEstadoCotizador(id);

            //Cabecera 
            Cotizador CabeceraCotizador = db.Cotizador.Find(id);

            //Obtener el Codigo Cotizacion

            CodigoCotizacion codigoCotizacion = db.CodigoCotizacion.Find(CabeceraCotizador.id_codigo_cotizacion);

            string basePath = ConfigurationManager.AppSettings["RepositorioDocumentos"];
            string rutaArchivos = basePath + "\\GESTION_PPM";
             

            //ruta del archivo PDF que se va a crear
            string rutaDocumentos = rutaArchivos + "\\Cotización-" + codigoCotizacion.codigo_cotizacion + "-Versión" + CabeceraCotizador.version.ToString().Replace(",", ".") + ".pdf";

            var relativePath = "";
            var filename = "";

            filename = "Cotización-" + codigoCotizacion.codigo_cotizacion + "-Versión" + CabeceraCotizador.version.ToString().Replace(",", ".") + ".pdf";
            relativePath = Tools.CrearCaminos(basePath, new List<string>() { "GESTION_PPM" });
            if (!relativePath.EndsWith("\\"))
                relativePath += "\\";
            relativePath = relativePath + "\\" + filename;

            return File(relativePath, Tools.GetContentType(Path.GetExtension(filename)), filename);
             
        }

        public ActionResult DescargarPDF(int id)
        {   
            //Cabecera 
            Cotizador CabeceraCotizador = db.Cotizador.Find(id);

            //Obtener el Codigo Cotizacion

            CodigoCotizacion codigoCotizacion = db.CodigoCotizacion.Find(CabeceraCotizador.id_codigo_cotizacion);

            string basePath = ConfigurationManager.AppSettings["RepositorioDocumentos"];
            string rutaArchivos = basePath + "\\GESTION_PPM";

            //ruta del archivo PDF que se va a crear
            string rutaDocumentos = rutaArchivos + "\\Cotización-" + codigoCotizacion.codigo_cotizacion + "-Versión" + CabeceraCotizador.version.ToString().Replace(",", ".") + ".pdf";

            var relativePath = "";
            var filename = "";

            filename = "Cotización-" + codigoCotizacion.codigo_cotizacion + "-Versión" + CabeceraCotizador.version.ToString().Replace(",", ".") + ".pdf";
            relativePath = Tools.CrearCaminos(basePath, new List<string>() { "GESTION_PPM" });
            if (!relativePath.EndsWith("\\"))
                relativePath += "\\";
            relativePath = relativePath + "\\" + filename;

            return File(relativePath, Tools.GetContentType(Path.GetExtension(filename)), filename);

        }

        public void CotizadorPDF(int? id)
        {
            int par = 0;
            //Obtener los datos del Cotizador
            usp_b_datos_cotizador CabeceraInicialCotizador = db.usp_b_datos_cotizador(id).First();

            //Cabecera 
            Cotizador CabeceraCotizador = db.Cotizador.Find(id);

            //Obtener el Codigo Cotizacion
            CodigoCotizacion codigoCotizacion = db.CodigoCotizacion.Find(CabeceraCotizador.id_codigo_cotizacion);

            //Detalle
            List<usp_b_datos_detalle_cotizador> detalle = db.usp_b_datos_detalle_cotizador(CabeceraCotizador.id_cotizador).ToList();

            //Fonts para el PDF
            //Regular
            string fontpath = Server.MapPath("~/Content/fonts/ubuntu/Raleway-Regular.ttf");
            BaseFont customfont = BaseFont.CreateFont(fontpath, BaseFont.CP1252, BaseFont.EMBEDDED);

            //Bold
            string fontpathbold = Server.MapPath("~/Content/fonts/ubuntu/Raleway-Bold.ttf");
            BaseFont customfontbold = BaseFont.CreateFont(fontpathbold, BaseFont.CP1252, BaseFont.EMBEDDED);

            //ruta imagen del formulario
            String FilePath = Server.MapPath("~/Content/img/LogoPPMPDF.png");
            var physicalPath = Server.MapPath("~/Content/img/LogoPPMPDF.png");

            string basePath = ConfigurationManager.AppSettings["RepositorioDocumentos"];
            string rutaArchivos = basePath + "\\GESTION_PPM";

            //ruta del archivo PDF que se va a crear
            string rutaDocumentos = rutaArchivos + "\\Cotización-" + codigoCotizacion.codigo_cotizacion + "-Versión" + CabeceraCotizador.version.ToString().Replace(",", ".") + ".pdf";

            //Crear el formulario PDF
            FileStream fs = new FileStream(rutaDocumentos, FileMode.Create);
            Document document = new Document(iTextSharp.text.PageSize.A4, -20, -20, 0, 0);
            //Document document = new Document(iTextSharp.text.PageSize.A4);
            PdfWriter pw = PdfWriter.GetInstance(document, fs);

            //Abrir archivo para su edicion
            document.Open();

            //Salto de linea 
            document.Add(new Paragraph(" "));

            //Salto de linea 
            document.Add(new Paragraph(" "));

            //Cabecera del Cotizador         
            PdfPTable cabecera = new PdfPTable(7);
            cabecera.DefaultCell.Border = Rectangle.NO_BORDER;

            //Imagen
            iTextSharp.text.Image imagen = iTextSharp.text.Image.GetInstance(FilePath);
            imagen.ScalePercent(10);

            PdfPCell logocell = new PdfPCell(imagen, true);
            logocell.BackgroundColor = new BaseColor(60, 66, 82);
            logocell.Colspan = 2;
            logocell.Border = Rectangle.NO_BORDER;
            logocell.HorizontalAlignment = Element.ALIGN_LEFT;
            logocell.VerticalAlignment = Element.ALIGN_MIDDLE;

            //Titulo
            PdfPCell TituloCotizador = new PdfPCell(new Phrase("COTIZACIÓN N° " + codigoCotizacion.codigo_cotizacion + "      ", new Font(customfontbold, 9, 0, new BaseColor(255, 255, 255))));
            TituloCotizador.Rowspan = 2;
            TituloCotizador.Colspan = 4;
            TituloCotizador.Border = Rectangle.NO_BORDER;
            TituloCotizador.HorizontalAlignment = Element.ALIGN_RIGHT;
            TituloCotizador.VerticalAlignment = Element.ALIGN_MIDDLE;
            TituloCotizador.BackgroundColor = new BaseColor(60, 66, 82);
            TituloCotizador.FixedHeight = 50f;

            //Version
            PdfPCell VersionCotizador = new PdfPCell(new Phrase("Versión  " + CabeceraCotizador.version.ToString().Replace(",", "."), new Font(customfontbold, 7, 0, new BaseColor(60, 66, 82))));
            VersionCotizador.Rowspan = 2;
            VersionCotizador.Colspan = 1;
            VersionCotizador.Border = Rectangle.NO_BORDER;
            VersionCotizador.HorizontalAlignment = Element.ALIGN_CENTER;
            VersionCotizador.VerticalAlignment = Element.ALIGN_MIDDLE;
            VersionCotizador.BackgroundColor = new BaseColor(240, 240, 240);
            VersionCotizador.FixedHeight = 50f;

            cabecera.AddCell(logocell);
            cabecera.AddCell(TituloCotizador);
            cabecera.AddCell(VersionCotizador);

            document.Add(cabecera);

            //Salto Linea      
            PdfPTable saltoLinea1 = new PdfPTable(4);

            PdfPCell EtiquetaSaltoLinea = new PdfPCell(new Phrase(" ", new Font(customfontbold, 9)));
            EtiquetaSaltoLinea.Colspan = 4;
            EtiquetaSaltoLinea.Border = Rectangle.NO_BORDER;
            EtiquetaSaltoLinea.HorizontalAlignment = Element.ALIGN_CENTER;
            EtiquetaSaltoLinea.VerticalAlignment = Element.ALIGN_MIDDLE;
            EtiquetaSaltoLinea.FixedHeight = 8f;

            saltoLinea1.AddCell(EtiquetaSaltoLinea);

            document.Add(saltoLinea1);

            //Informacion del Cotizador         
            PdfPTable informacion = new PdfPTable(4);

            PdfPCell EtiquetaFechaCotizacion = new PdfPCell(new Phrase("Fecha de cotización: ", new Font(customfontbold, 7)));
            EtiquetaFechaCotizacion.Colspan = 1;
            EtiquetaFechaCotizacion.HorizontalAlignment = Element.ALIGN_RIGHT;
            EtiquetaFechaCotizacion.VerticalAlignment = Element.ALIGN_MIDDLE;
            EtiquetaFechaCotizacion.FixedHeight = 16f;
            EtiquetaFechaCotizacion.BackgroundColor = new BaseColor(225, 225, 225);
            EtiquetaFechaCotizacion.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell ValorFechaCotizacion = new PdfPCell(new Phrase("   " + CabeceraInicialCotizador.fechaCotizacion, new Font(customfont, 7)));
            ValorFechaCotizacion.Colspan = 1;
            ValorFechaCotizacion.HorizontalAlignment = Element.ALIGN_LEFT;
            ValorFechaCotizacion.VerticalAlignment = Element.ALIGN_MIDDLE;
            ValorFechaCotizacion.FixedHeight = 16f;
            ValorFechaCotizacion.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell EtiquetaEjecutivo = new PdfPCell(new Phrase("Ejecutiva comercial:  ", new Font(customfont, 7)));
            EtiquetaEjecutivo.Colspan = 1;
            EtiquetaEjecutivo.HorizontalAlignment = Element.ALIGN_RIGHT;
            EtiquetaEjecutivo.VerticalAlignment = Element.ALIGN_MIDDLE;
            EtiquetaEjecutivo.FixedHeight = 16f;
            EtiquetaEjecutivo.BackgroundColor = new BaseColor(225, 225, 225);
            EtiquetaEjecutivo.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell ValorEjecutivo = new PdfPCell(new Phrase("   " + CabeceraInicialCotizador.EjecutivoComercial, new Font(customfont, 7)));
            ValorEjecutivo.Colspan = 1;
            ValorEjecutivo.HorizontalAlignment = Element.ALIGN_LEFT;
            ValorEjecutivo.VerticalAlignment = Element.ALIGN_MIDDLE;
            ValorEjecutivo.FixedHeight = 16f;
            ValorEjecutivo.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell EtiquetaFechaVencimiento = new PdfPCell(new Phrase("Fecha de vencimiento: ", new Font(customfontbold, 7, 0, new BaseColor(255, 255, 255))));
            EtiquetaFechaVencimiento.Colspan = 1;
            EtiquetaFechaVencimiento.HorizontalAlignment = Element.ALIGN_RIGHT;
            EtiquetaFechaVencimiento.VerticalAlignment = Element.ALIGN_MIDDLE;
            EtiquetaFechaVencimiento.FixedHeight = 16f;
            EtiquetaFechaVencimiento.BackgroundColor = new BaseColor(60, 66, 82);
            EtiquetaFechaVencimiento.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell ValorFechaVencimiento = new PdfPCell(new Phrase("   " + CabeceraInicialCotizador.fechaVencimiento, new Font(customfontbold, 7, 0, new BaseColor(60, 66, 82))));
            ValorFechaVencimiento.Colspan = 1;
            ValorFechaVencimiento.HorizontalAlignment = Element.ALIGN_LEFT;
            ValorFechaVencimiento.VerticalAlignment = Element.ALIGN_MIDDLE;
            ValorFechaVencimiento.FixedHeight = 16f;
            ValorFechaVencimiento.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell EtiquetaTiempoIncioProyecto = new PdfPCell(new Phrase("Tiempo de inicio de proyecto: ", new Font(customfont, 7, 0, new BaseColor(255, 255, 255))));
            EtiquetaTiempoIncioProyecto.Colspan = 1;
            EtiquetaTiempoIncioProyecto.HorizontalAlignment = Element.ALIGN_RIGHT;
            EtiquetaTiempoIncioProyecto.VerticalAlignment = Element.ALIGN_MIDDLE;
            EtiquetaTiempoIncioProyecto.FixedHeight = 16f;
            EtiquetaTiempoIncioProyecto.BackgroundColor = new BaseColor(60, 66, 82);
            EtiquetaTiempoIncioProyecto.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell ValorFechaInicioProyecto = new PdfPCell(new Phrase("   " + CabeceraInicialCotizador.TiempoInicioActividades, new Font(customfontbold, 7, 0, new BaseColor(60, 66, 82))));
            ValorFechaInicioProyecto.Colspan = 1;
            ValorFechaInicioProyecto.HorizontalAlignment = Element.ALIGN_LEFT;
            ValorFechaInicioProyecto.VerticalAlignment = Element.ALIGN_MIDDLE;
            ValorFechaInicioProyecto.FixedHeight = 16f;
            ValorFechaInicioProyecto.BorderColor = new BaseColor(60, 66, 82);

            informacion.AddCell(EtiquetaFechaCotizacion);
            informacion.AddCell(ValorFechaCotizacion);
            informacion.AddCell(EtiquetaEjecutivo);
            informacion.AddCell(ValorEjecutivo);

            informacion.AddCell(EtiquetaFechaVencimiento);
            informacion.AddCell(ValorFechaVencimiento);
            informacion.AddCell(EtiquetaTiempoIncioProyecto);
            informacion.AddCell(ValorFechaInicioProyecto);

            document.Add(informacion);

            //Salto Linea       
            PdfPTable saltoLinea2 = new PdfPTable(4);

            PdfPCell EtiquetaSaltoLinea2 = new PdfPCell(new Phrase(" ", new Font(customfontbold, 9)));
            EtiquetaSaltoLinea2.Colspan = 4;
            EtiquetaSaltoLinea2.Border = Rectangle.NO_BORDER;
            EtiquetaSaltoLinea2.HorizontalAlignment = Element.ALIGN_CENTER;
            EtiquetaSaltoLinea2.VerticalAlignment = Element.ALIGN_MIDDLE;
            EtiquetaSaltoLinea2.FixedHeight = 8f;

            saltoLinea2.AddCell(EtiquetaSaltoLinea2);

            document.Add(saltoLinea2);

            //Informacion del Cotizador         
            PdfPTable cliente = new PdfPTable(6);

            PdfPCell EtiquetaCliente = new PdfPCell(new Phrase("Cliente: ", new Font(customfontbold, 7)));
            EtiquetaCliente.Colspan = 1;
            EtiquetaCliente.HorizontalAlignment = Element.ALIGN_RIGHT;
            EtiquetaCliente.VerticalAlignment = Element.ALIGN_MIDDLE;
            EtiquetaCliente.FixedHeight = 16f;
            EtiquetaCliente.BackgroundColor = new BaseColor(225, 225, 225);
            EtiquetaCliente.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell ValorCliente = new PdfPCell(new Phrase("   " + CabeceraInicialCotizador.Cliente, new Font(customfont, 7)));
            ValorCliente.Colspan = 5;
            ValorCliente.HorizontalAlignment = Element.ALIGN_LEFT;
            ValorCliente.VerticalAlignment = Element.ALIGN_MIDDLE;
            ValorCliente.FixedHeight = 16f;
            ValorCliente.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell EtiquetaContacto = new PdfPCell(new Phrase("Contacto: ", new Font(customfontbold, 7)));
            EtiquetaContacto.Colspan = 1;
            EtiquetaContacto.HorizontalAlignment = Element.ALIGN_RIGHT;
            EtiquetaContacto.VerticalAlignment = Element.ALIGN_MIDDLE;
            EtiquetaContacto.FixedHeight = 16f;
            EtiquetaContacto.BackgroundColor = new BaseColor(225, 225, 225);
            EtiquetaContacto.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell ValorContacto = new PdfPCell(new Phrase("   " + CabeceraInicialCotizador.Contacto, new Font(customfont, 7)));
            ValorContacto.Colspan = 5;
            ValorContacto.HorizontalAlignment = Element.ALIGN_LEFT;
            ValorContacto.VerticalAlignment = Element.ALIGN_MIDDLE;
            ValorContacto.FixedHeight = 16f;
            ValorContacto.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell EtiquetaNombreProyecto = new PdfPCell(new Phrase("Nombre del Proyecto: ", new Font(customfontbold, 7)));
            EtiquetaNombreProyecto.Colspan = 1;
            EtiquetaNombreProyecto.HorizontalAlignment = Element.ALIGN_RIGHT;
            EtiquetaNombreProyecto.VerticalAlignment = Element.ALIGN_MIDDLE;
            EtiquetaNombreProyecto.FixedHeight = 16f;
            EtiquetaNombreProyecto.BackgroundColor = new BaseColor(225, 225, 225);
            EtiquetaNombreProyecto.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell ValorNombreProyecto = new PdfPCell(new Phrase("   " + CabeceraInicialCotizador.NombreProyecto, new Font(customfont, 7)));
            ValorNombreProyecto.Colspan = 5;
            ValorNombreProyecto.HorizontalAlignment = Element.ALIGN_LEFT;
            ValorNombreProyecto.VerticalAlignment = Element.ALIGN_MIDDLE;
            ValorNombreProyecto.FixedHeight = 16f;
            ValorNombreProyecto.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell EtiquetaTipoProyecto = new PdfPCell(new Phrase("Tipo de proyecto: ", new Font(customfontbold, 7)));
            EtiquetaTipoProyecto.Colspan = 1;
            EtiquetaTipoProyecto.HorizontalAlignment = Element.ALIGN_RIGHT;
            EtiquetaTipoProyecto.VerticalAlignment = Element.ALIGN_MIDDLE;
            EtiquetaTipoProyecto.FixedHeight = 16f;
            EtiquetaTipoProyecto.BackgroundColor = new BaseColor(225, 225, 225);
            EtiquetaTipoProyecto.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell ValorTipoProyecto = new PdfPCell(new Phrase("   " + CabeceraInicialCotizador.TipoProyecto, new Font(customfontbold, 7)));
            ValorTipoProyecto.Colspan = 2;
            ValorTipoProyecto.HorizontalAlignment = Element.ALIGN_LEFT;
            ValorTipoProyecto.VerticalAlignment = Element.ALIGN_MIDDLE;
            ValorTipoProyecto.FixedHeight = 16f;
            ValorTipoProyecto.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell EtiquetaDimensionProyecto = new PdfPCell(new Phrase("Dimensión del proyecto: ", new Font(customfontbold, 7)));
            EtiquetaDimensionProyecto.Colspan = 2;
            EtiquetaDimensionProyecto.HorizontalAlignment = Element.ALIGN_RIGHT;
            EtiquetaDimensionProyecto.VerticalAlignment = Element.ALIGN_MIDDLE;
            EtiquetaDimensionProyecto.FixedHeight = 16f;
            EtiquetaDimensionProyecto.BackgroundColor = new BaseColor(225, 225, 225);
            EtiquetaDimensionProyecto.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell ValorDimensionProyecto = new PdfPCell(new Phrase("   " + CabeceraInicialCotizador.DimensionProyecto, new Font(customfont, 7)));
            ValorDimensionProyecto.Colspan = 1;
            ValorDimensionProyecto.HorizontalAlignment = Element.ALIGN_LEFT;
            ValorDimensionProyecto.VerticalAlignment = Element.ALIGN_MIDDLE;
            ValorDimensionProyecto.FixedHeight = 16f;
            ValorDimensionProyecto.BorderColor = new BaseColor(60, 66, 82);

            cliente.AddCell(EtiquetaCliente);
            cliente.AddCell(ValorCliente);

            cliente.AddCell(EtiquetaContacto);
            cliente.AddCell(ValorContacto);

            cliente.AddCell(EtiquetaNombreProyecto);
            cliente.AddCell(ValorNombreProyecto);

            cliente.AddCell(EtiquetaTipoProyecto);
            cliente.AddCell(ValorTipoProyecto);
            cliente.AddCell(EtiquetaDimensionProyecto);
            cliente.AddCell(ValorDimensionProyecto);

            document.Add(cliente);

            //Salto Linea       
            PdfPTable saltoLinea3 = new PdfPTable(4);

            PdfPCell EtiquetaSaltoLinea3 = new PdfPCell(new Phrase(" ", new Font(customfontbold, 9)));
            EtiquetaSaltoLinea3.Colspan = 4;
            EtiquetaSaltoLinea3.Border = Rectangle.NO_BORDER;
            EtiquetaSaltoLinea3.HorizontalAlignment = Element.ALIGN_CENTER;
            EtiquetaSaltoLinea3.VerticalAlignment = Element.ALIGN_MIDDLE;
            EtiquetaSaltoLinea3.FixedHeight = 8f;

            saltoLinea3.AddCell(EtiquetaSaltoLinea3);

            document.Add(saltoLinea3);

            //Entregables del Cotizador         
            PdfPTable entregables = new PdfPTable(6);

            PdfPCell EtiquetaEntregables = new PdfPCell(new Phrase("E   N   T   R   E   G   A   B   L   E", new Font(customfont, 9, 0, new BaseColor(255, 255, 255))));
            EtiquetaEntregables.Colspan = 5;
            EtiquetaEntregables.HorizontalAlignment = Element.ALIGN_CENTER;
            EtiquetaEntregables.VerticalAlignment = Element.ALIGN_MIDDLE;
            EtiquetaEntregables.FixedHeight = 16f;
            EtiquetaEntregables.BackgroundColor = new BaseColor(60, 66, 82);
            EtiquetaEntregables.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell EtiquetaCostoTotal = new PdfPCell(new Phrase("COSTO TOTAL", new Font(customfontbold, 6, 0, new BaseColor(60, 66, 82))));
            EtiquetaCostoTotal.Colspan = 1;
            EtiquetaCostoTotal.HorizontalAlignment = Element.ALIGN_CENTER;
            EtiquetaCostoTotal.VerticalAlignment = Element.ALIGN_MIDDLE;
            EtiquetaCostoTotal.FixedHeight = 16f;
            EtiquetaCostoTotal.BorderColor = new BaseColor(60, 66, 82);

            entregables.AddCell(EtiquetaEntregables);
            entregables.AddCell(EtiquetaCostoTotal);


            foreach (usp_b_datos_detalle_cotizador dtc in detalle)
            {
                PdfPCell EtiquetaDetalle = new PdfPCell(new Phrase("   " + dtc.Entregable, new Font(customfont, 7)));
                EtiquetaDetalle.Colspan = 5;
                EtiquetaDetalle.HorizontalAlignment = Element.ALIGN_LEFT;
                EtiquetaDetalle.VerticalAlignment = Element.ALIGN_MIDDLE;
                EtiquetaDetalle.FixedHeight = 23f;
                EtiquetaDetalle.BorderColor = new BaseColor(60, 66, 82);

                PdfPCell ValorDetalle = new PdfPCell(new Phrase((((String.Format("{0:n}", dtc.Total).Replace(",","-")).Replace(".",",")).Replace("-",".")), new Font(customfont, 7)));
                ValorDetalle.Colspan = 1;
                ValorDetalle.HorizontalAlignment = Element.ALIGN_RIGHT;
                ValorDetalle.VerticalAlignment = Element.ALIGN_MIDDLE;
                ValorDetalle.FixedHeight = 23f;
                ValorDetalle.BorderColor = new BaseColor(60, 66, 82);

                if (par == 1 || par == 3 || par == 5 || par == 7 || par == 9 || par == 1)
                {
                    ValorDetalle.BackgroundColor = new BaseColor(225, 225, 225);
                }

                entregables.AddCell(EtiquetaDetalle);
                entregables.AddCell(ValorDetalle);

                par = par + 1;
            }
             
            int registros = 10;

            if (codigoCotizacion.aplica_contrato == false)
            {
                registros = 13;
            }


            for (int i = par; i <= registros; i++)
            {
                PdfPCell EtiquetaDetalle = new PdfPCell(new Phrase("   ", new Font(customfont, 7)));
                EtiquetaDetalle.Colspan = 5;
                EtiquetaDetalle.HorizontalAlignment = Element.ALIGN_LEFT;
                EtiquetaDetalle.VerticalAlignment = Element.ALIGN_MIDDLE;
                EtiquetaDetalle.FixedHeight = 23f;
                EtiquetaDetalle.BorderColor = new BaseColor(60, 66, 82);

                PdfPCell ValorDetalle = new PdfPCell(new Phrase("        ", new Font(customfont, 7)));
                ValorDetalle.Colspan = 1;
                ValorDetalle.HorizontalAlignment = Element.ALIGN_RIGHT;
                ValorDetalle.VerticalAlignment = Element.ALIGN_MIDDLE;
                ValorDetalle.FixedHeight = 23f;
                ValorDetalle.BorderColor = new BaseColor(60, 66, 82);

                if (par == 1 || par == 3 || par == 5 || par == 7 || par == 9 || par == 11 || par == 13)
                {
                    ValorDetalle.BackgroundColor = new BaseColor(225, 225, 225);
                }

                entregables.AddCell(EtiquetaDetalle);
                entregables.AddCell(ValorDetalle);

                par = par + 1;
            }

            document.Add(entregables);

            //Salto Linea       
            PdfPTable saltoLinea4 = new PdfPTable(4);

            PdfPCell EtiquetaSaltoLinea4 = new PdfPCell(new Phrase(" ", new Font(customfontbold, 9)));
            EtiquetaSaltoLinea4.Colspan = 4;
            EtiquetaSaltoLinea4.Border = Rectangle.NO_BORDER;
            EtiquetaSaltoLinea4.HorizontalAlignment = Element.ALIGN_CENTER;
            EtiquetaSaltoLinea4.VerticalAlignment = Element.ALIGN_MIDDLE;
            EtiquetaSaltoLinea4.FixedHeight = 8f;

            saltoLinea4.AddCell(EtiquetaSaltoLinea3);

            document.Add(saltoLinea4);

            //Datos Finales        
            PdfPTable finales = new PdfPTable(6);

            PdfPCell EtiquetalObservaciones = new PdfPCell(new Phrase("Observaciones: ", new Font(customfontbold, 7, 0, new BaseColor(60, 66, 82))));
            EtiquetalObservaciones.Colspan = 1;
            EtiquetalObservaciones.Rowspan = 4;
            EtiquetalObservaciones.HorizontalAlignment = Element.ALIGN_RIGHT;
            EtiquetalObservaciones.VerticalAlignment = Element.ALIGN_MIDDLE;
            EtiquetalObservaciones.FixedHeight = 16f;
            EtiquetalObservaciones.BackgroundColor = new BaseColor(225, 225, 225);
            EtiquetalObservaciones.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell ValorObservaciones = new PdfPCell(new Phrase(CabeceraInicialCotizador.Observaciones, new Font(customfont, 7)));
            ValorObservaciones.Colspan = 3;
            ValorObservaciones.Rowspan = 4;
            ValorObservaciones.HorizontalAlignment = Element.ALIGN_CENTER;
            ValorObservaciones.VerticalAlignment = Element.ALIGN_MIDDLE;
            ValorObservaciones.FixedHeight = 16f;
            ValorObservaciones.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell EtiquetalSubtotal = new PdfPCell(new Phrase("Subtotal", new Font(customfont, 7)));
            EtiquetalSubtotal.Colspan = 1;
            EtiquetalSubtotal.HorizontalAlignment = Element.ALIGN_RIGHT;
            EtiquetalSubtotal.VerticalAlignment = Element.ALIGN_MIDDLE;
            EtiquetalSubtotal.FixedHeight = 16f;
            EtiquetalSubtotal.BackgroundColor = new BaseColor(225, 225, 225);
            EtiquetalSubtotal.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell ValorSubtotal = new PdfPCell(new Phrase("US $ " + (((String.Format("{0:n}", CabeceraInicialCotizador.Subtotal).Replace(",", "-")).Replace(".", ",")).Replace("-", ".")), new Font(customfont, 7)));
            ValorSubtotal.Colspan = 1;
            ValorSubtotal.HorizontalAlignment = Element.ALIGN_RIGHT;
            ValorSubtotal.VerticalAlignment = Element.ALIGN_MIDDLE;
            ValorSubtotal.FixedHeight = 16f;
            ValorSubtotal.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell EtiquetalDescuento = new PdfPCell(new Phrase(((CabeceraInicialCotizador.Descuento.ToString().Replace(".", ","))) + " % Descuento", new Font(customfont, 7)));
            EtiquetalDescuento.Colspan = 1;
            EtiquetalDescuento.HorizontalAlignment = Element.ALIGN_RIGHT;
            EtiquetalDescuento.VerticalAlignment = Element.ALIGN_MIDDLE;
            EtiquetalDescuento.FixedHeight = 16f;
            EtiquetalDescuento.BackgroundColor = new BaseColor(225, 225, 225);
            EtiquetalDescuento.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell ValorDescuento = new PdfPCell(new Phrase("US $ " + (((String.Format("{0:n}", CabeceraInicialCotizador.ValorDescuento).Replace(",", "-")).Replace(".", ",")).Replace("-", ".")), new Font(customfont, 7)));
            ValorDescuento.Colspan = 1;
            ValorDescuento.HorizontalAlignment = Element.ALIGN_RIGHT;
            ValorDescuento.VerticalAlignment = Element.ALIGN_MIDDLE;
            ValorDescuento.FixedHeight = 16f;
            ValorDescuento.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell EtiquetalIva = new PdfPCell(new Phrase(CabeceraInicialCotizador.Iva.ToString(), new Font(customfont, 7)));
            EtiquetalIva.Colspan = 1;
            EtiquetalIva.HorizontalAlignment = Element.ALIGN_RIGHT;
            EtiquetalIva.VerticalAlignment = Element.ALIGN_MIDDLE;
            EtiquetalIva.FixedHeight = 16f;
            EtiquetalIva.BackgroundColor = new BaseColor(225, 225, 225);
            EtiquetalIva.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell ValorIVA = new PdfPCell(new Phrase("US $ " + (((String.Format("{0:n}", CabeceraInicialCotizador.ValorIva).Replace(",", "-")).Replace(".", ",")).Replace("-", ".")), new Font(customfont, 7)));
            ValorIVA.Colspan = 1;
            ValorIVA.HorizontalAlignment = Element.ALIGN_RIGHT;
            ValorIVA.VerticalAlignment = Element.ALIGN_MIDDLE;
            ValorIVA.FixedHeight = 16f;
            ValorIVA.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell EtiquetalTotal = new PdfPCell(new Phrase("TOTAL", new Font(customfontbold, 7, 0, new BaseColor(255, 255, 255))));
            EtiquetalTotal.Colspan = 1;
            EtiquetalTotal.HorizontalAlignment = Element.ALIGN_RIGHT;
            EtiquetalTotal.VerticalAlignment = Element.ALIGN_MIDDLE;
            EtiquetalTotal.FixedHeight = 16f;
            EtiquetalTotal.BackgroundColor = new BaseColor(60, 66, 82);
            EtiquetalTotal.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell ValorTotal = new PdfPCell(new Phrase("US $ " + (((String.Format("{0:n}", CabeceraInicialCotizador.Total).Replace(",", "-")).Replace(".", ",")).Replace("-", ".")), new Font(customfontbold, 7, 0, new BaseColor(60, 66, 82))));
            ValorTotal.Colspan = 1;
            ValorTotal.HorizontalAlignment = Element.ALIGN_RIGHT;
            ValorTotal.VerticalAlignment = Element.ALIGN_MIDDLE;
            ValorTotal.FixedHeight = 16f;
            ValorTotal.BackgroundColor = new BaseColor(225, 225, 225);
            ValorTotal.BorderColor = new BaseColor(60, 66, 82);

            finales.AddCell(EtiquetalObservaciones);
            finales.AddCell(ValorObservaciones);
            finales.AddCell(EtiquetalSubtotal);
            finales.AddCell(ValorSubtotal);

            finales.AddCell(EtiquetalDescuento);
            finales.AddCell(ValorDescuento);

            finales.AddCell(EtiquetalIva);
            finales.AddCell(ValorIVA);

            finales.AddCell(EtiquetalTotal);
            finales.AddCell(ValorTotal);

            document.Add(finales);

            //Salto Linea       
            PdfPTable saltoLinea5 = new PdfPTable(4);

            PdfPCell EtiquetaSaltoLinea5 = new PdfPCell(new Phrase(" ", new Font(customfontbold, 9)));
            EtiquetaSaltoLinea5.Colspan = 4;
            EtiquetaSaltoLinea5.Border = Rectangle.NO_BORDER;
            EtiquetaSaltoLinea5.HorizontalAlignment = Element.ALIGN_CENTER;
            EtiquetaSaltoLinea5.VerticalAlignment = Element.ALIGN_MIDDLE;
            EtiquetaSaltoLinea5.FixedHeight = 8f;

            saltoLinea5.AddCell(EtiquetaSaltoLinea5);

            document.Add(saltoLinea5);

            //Datos Finales        
            PdfPTable terminosCondiciones = new PdfPTable(1);

            //Termino2
            PdfPTable Terminos = new PdfPTable(19);
            Terminos.DefaultCell.Border = Rectangle.NO_BORDER;

            //Encabecazado  
            PdfPCell EtiquetaTerminos = new PdfPCell(new Phrase("Términos y Condiciones Generales: ", new Font(customfontbold, 7, 0, new BaseColor(60, 66, 82))));
            EtiquetaTerminos.Colspan = 19;
            EtiquetaTerminos.HorizontalAlignment = Element.ALIGN_CENTER;
            EtiquetaTerminos.VerticalAlignment = Element.ALIGN_MIDDLE;
            EtiquetaTerminos.FixedHeight = 19f;
            EtiquetaTerminos.BackgroundColor = new BaseColor(225, 225, 225);
            EtiquetaTerminos.BorderColor = new BaseColor(255, 255, 255);

            PdfPCell Etiqueta1 = new PdfPCell(new Phrase("1. ", new Font(customfontbold, 7, 0, new BaseColor(60, 66, 82))));
            Etiqueta1.Colspan = 1;
            Etiqueta1.Rowspan = 2;
            Etiqueta1.FixedHeight = 24f;
            Etiqueta1.HorizontalAlignment = Element.ALIGN_RIGHT;
            Etiqueta1.VerticalAlignment = Element.ALIGN_MIDDLE;
            Etiqueta1.BorderColor = new BaseColor(255, 255, 255);

            PdfPCell EtiquetaTermino1 = new PdfPCell(new Phrase("Los valores fueron cotizados de acuerdo al alcance del trabajo determinado entre " + CabeceraInicialCotizador.Cliente
                + " y PUBLIPROMUEVE S.A.TRABAJOS ADICIONALES Y / O ESPECIALES Y/ O NO INCLUIDOS EN ESTA COTIZACIÓN SERÁN ACORDADOS PREVIAMENTE ENTRE LAS PARTES.", new Font(customfont, 7)));
            EtiquetaTermino1.Colspan = 18;
            EtiquetaTermino1.Rowspan = 2;
            EtiquetaTermino1.FixedHeight = 24f;
            EtiquetaTermino1.HorizontalAlignment = Element.ALIGN_LEFT;
            EtiquetaTermino1.VerticalAlignment = Element.ALIGN_MIDDLE;
            EtiquetaTermino1.BorderColor = new BaseColor(255, 255, 255);

            PdfPCell Etiqueta2 = new PdfPCell(new Phrase("2. ", new Font(customfontbold, 7, 0, new BaseColor(60, 66, 82))));
            Etiqueta2.Colspan = 1;
            Etiqueta2.HorizontalAlignment = Element.ALIGN_RIGHT;
            Etiqueta2.VerticalAlignment = Element.ALIGN_MIDDLE;
            Etiqueta2.FixedHeight = 24f;
            Etiqueta2.BorderColor = new BaseColor(255, 255, 255);

            PdfPCell EtiquetaTermino2 = new PdfPCell(new Phrase("Si "+CabeceraInicialCotizador.Cliente + " está de acuerdo con la cotización, debe enviar la aprobación al correo electrónico: " + CabeceraInicialCotizador.mail + ".", new Font(customfont, 7)));
            EtiquetaTermino2.Colspan = 18;
            EtiquetaTermino2.HorizontalAlignment = Element.ALIGN_LEFT;
            EtiquetaTermino2.VerticalAlignment = Element.ALIGN_MIDDLE;
            EtiquetaTermino2.FixedHeight = 24f;
            EtiquetaTermino2.BorderColor = new BaseColor(255, 255, 255);

            PdfPCell Etiqueta3 = new PdfPCell(new Phrase("3. ", new Font(customfontbold, 7, 0, new BaseColor(60, 66, 82))));
            Etiqueta3.Colspan = 1;
            Etiqueta3.HorizontalAlignment = Element.ALIGN_RIGHT;
            Etiqueta3.VerticalAlignment = Element.ALIGN_MIDDLE;
            Etiqueta3.FixedHeight = 24f;
            Etiqueta3.BorderColor = new BaseColor(255, 255, 255);

            PdfPCell EtiquetaTermino3 = new PdfPCell(new Phrase("UNA VEZ APROBADA LA PRESENTE COTIZACIÓN, " + CabeceraInicialCotizador.EjecutivoComercial
                + " LE COMUNICARÁ A " + CabeceraInicialCotizador.Cliente + " LOS TIEMPOS DE INICIO Y FINALIZACIÓN DE LOS TRABAJOS SOLICITADOS.", new Font(customfont, 7)));
            EtiquetaTermino3.Colspan = 18;
            EtiquetaTermino3.HorizontalAlignment = Element.ALIGN_LEFT;
            EtiquetaTermino3.VerticalAlignment = Element.ALIGN_MIDDLE;
            EtiquetaTermino3.FixedHeight = 24f;
            EtiquetaTermino3.BorderColor = new BaseColor(255, 255, 255);

            PdfPCell Etiqueta4 = new PdfPCell(new Phrase("4. ", new Font(customfontbold, 7, 0, new BaseColor(60, 66, 82))));
            Etiqueta4.Colspan = 1;
            Etiqueta4.HorizontalAlignment = Element.ALIGN_RIGHT;
            Etiqueta4.VerticalAlignment = Element.ALIGN_MIDDLE;
            Etiqueta4.FixedHeight = 24f;
            Etiqueta4.BorderColor = new BaseColor(255, 255, 255);

            PdfPCell EtiquetaTermino4 = new PdfPCell(new Phrase("Por medio de la presente, " + CabeceraInicialCotizador.Cliente
                + " formaliza una orden de pedido a PUBLIPROMUEVE S.A., que si es necesario por las partes se podrá formalizar en un contrato; la misma no exime a una firma de contrato.", new Font(customfont, 7)));
            EtiquetaTermino4.Colspan = 18;
            EtiquetaTermino4.HorizontalAlignment = Element.ALIGN_LEFT;
            EtiquetaTermino4.VerticalAlignment = Element.ALIGN_MIDDLE;
            EtiquetaTermino4.FixedHeight = 24f;
            EtiquetaTermino4.BorderColor = new BaseColor(255, 255, 255);

            PdfPCell Etiqueta5 = new PdfPCell(new Phrase("5. ", new Font(customfontbold, 7, 0, new BaseColor(60, 66, 82))));
            Etiqueta5.Colspan = 1;
            Etiqueta5.HorizontalAlignment = Element.ALIGN_RIGHT;
            Etiqueta5.VerticalAlignment = Element.ALIGN_MIDDLE;
            Etiqueta5.BorderColor = new BaseColor(255, 255, 255);

            PdfPCell EtiquetaTermino5 = new PdfPCell(new Phrase("En el caso de formalizarlo en un contrato, PUBLIPROMUEVE S.A. facturará de la siguiente manera: ", new Font(customfont, 7)));
            EtiquetaTermino5.Colspan = 18;
            EtiquetaTermino5.HorizontalAlignment = Element.ALIGN_LEFT;
            EtiquetaTermino5.VerticalAlignment = Element.ALIGN_MIDDLE;
            EtiquetaTermino5.BorderColor = new BaseColor(255, 255, 255);

            PdfPCell Etiqueta51 = new PdfPCell(new Phrase(" ", new Font(customfontbold, 7, 0, new BaseColor(60, 66, 82))));
            Etiqueta51.Colspan = 1;
            Etiqueta51.HorizontalAlignment = Element.ALIGN_RIGHT;
            Etiqueta51.VerticalAlignment = Element.ALIGN_MIDDLE;
            Etiqueta51.BorderColor = new BaseColor(255, 255, 255);

            PdfPCell EtiquetaTermino51 = new PdfPCell(new Phrase("- El 30% cuando se aprueba la cotización; ", new Font(customfont, 7)));
            EtiquetaTermino51.Colspan = 19;
            EtiquetaTermino51.HorizontalAlignment = Element.ALIGN_LEFT;
            EtiquetaTermino51.VerticalAlignment = Element.ALIGN_MIDDLE;
            EtiquetaTermino51.FixedHeight = 18f;
            EtiquetaTermino51.BorderColor = new BaseColor(255, 255, 255);

            PdfPCell Etiqueta52 = new PdfPCell(new Phrase(" ", new Font(customfontbold, 7, 0, new BaseColor(60, 66, 82))));
            Etiqueta52.Colspan = 1;
            Etiqueta52.HorizontalAlignment = Element.ALIGN_RIGHT;
            Etiqueta52.VerticalAlignment = Element.ALIGN_MIDDLE;
            Etiqueta52.BorderColor = new BaseColor(255, 255, 255);

            PdfPCell EtiquetaTermino52 = new PdfPCell(new Phrase("- El 20% a los 30 días contados a partir de la firma del contrato; y, ", new Font(customfont, 7)));
            EtiquetaTermino52.Colspan = 19;
            EtiquetaTermino52.HorizontalAlignment = Element.ALIGN_LEFT;
            EtiquetaTermino52.VerticalAlignment = Element.ALIGN_MIDDLE;
            EtiquetaTermino52.BorderColor = new BaseColor(255, 255, 255);

            PdfPCell Etiqueta53 = new PdfPCell(new Phrase(" ", new Font(customfontbold, 7, 0, new BaseColor(60, 66, 82))));
            Etiqueta53.Colspan = 1;
            Etiqueta53.HorizontalAlignment = Element.ALIGN_RIGHT;
            Etiqueta53.VerticalAlignment = Element.ALIGN_MIDDLE;
            Etiqueta53.BorderColor = new BaseColor(255, 255, 255);

            PdfPCell EtiquetaTermino53 = new PdfPCell(new Phrase("- El 50% restante cuando finalice el proyecto.", new Font(customfont, 7)));
            EtiquetaTermino53.Colspan = 19;
            EtiquetaTermino53.HorizontalAlignment = Element.ALIGN_LEFT;
            EtiquetaTermino53.VerticalAlignment = Element.ALIGN_MIDDLE;
            EtiquetaTermino53.BorderColor = new BaseColor(255, 255, 255);

            PdfPCell Etiqueta6 = new PdfPCell(new Phrase("6. ", new Font(customfontbold, 7, 0, new BaseColor(60, 66, 82))));
            Etiqueta6.Colspan = 1;
            Etiqueta6.HorizontalAlignment = Element.ALIGN_RIGHT;
            Etiqueta6.VerticalAlignment = Element.ALIGN_MIDDLE;
            Etiqueta6.FixedHeight = 24f;
            Etiqueta6.BorderColor = new BaseColor(255, 255, 255);

            PdfPCell EtiquetaTermino6 = new PdfPCell(new Phrase("Los valores serán pagaderos a 30 días previa presentación de las factura(s).", new Font(customfont, 7)));
            EtiquetaTermino6.Colspan = 18;
            EtiquetaTermino6.HorizontalAlignment = Element.ALIGN_LEFT;
            EtiquetaTermino6.VerticalAlignment = Element.ALIGN_MIDDLE;
            EtiquetaTermino6.FixedHeight = 24f;
            EtiquetaTermino6.BorderColor = new BaseColor(255, 255, 255);

            PdfPCell Etiqueta7 = new PdfPCell(new Phrase("7. ", new Font(customfontbold, 7, 0, new BaseColor(60, 66, 82))));
            Etiqueta7.Colspan = 1;
            Etiqueta7.HorizontalAlignment = Element.ALIGN_RIGHT;
            Etiqueta7.VerticalAlignment = Element.ALIGN_MIDDLE;
            Etiqueta7.FixedHeight = 24f;
            Etiqueta7.BorderColor = new BaseColor(255, 255, 255);

            PdfPCell EtiquetaTermino7 = new PdfPCell(new Phrase("Una vez que " + CabeceraInicialCotizador.Cliente
               + " apruebe esta cotización, las partes acordarán los plazos para la entrega de información de forma conjunta.", new Font(customfont, 7)));
            EtiquetaTermino7.Colspan = 18;
            EtiquetaTermino7.HorizontalAlignment = Element.ALIGN_LEFT;
            EtiquetaTermino7.VerticalAlignment = Element.ALIGN_MIDDLE;
            EtiquetaTermino7.FixedHeight = 24f;
            EtiquetaTermino7.BorderColor = new BaseColor(255, 255, 255);

            PdfPCell Etiqueta8 = new PdfPCell(new Phrase("8. ", new Font(customfontbold, 7, 0, new BaseColor(60, 66, 82))));
            Etiqueta8.Colspan = 1;
            Etiqueta8.HorizontalAlignment = Element.ALIGN_RIGHT;
            Etiqueta8.VerticalAlignment = Element.ALIGN_MIDDLE;
            Etiqueta8.FixedHeight = 24f;
            Etiqueta8.BorderColor = new BaseColor(255, 255, 255);

            PdfPCell EtiquetaTermino8 = new PdfPCell(new Phrase("PUBLIPROMUEVE S.A. iniciará los trabajos una vez que recibamos la información necesaria, la misma que será proporcionada en forma oportuna y completa por parte de " + CabeceraInicialCotizador.Cliente + ".", new Font(customfont, 7)));
            EtiquetaTermino8.Colspan = 18;
            EtiquetaTermino8.HorizontalAlignment = Element.ALIGN_LEFT;
            EtiquetaTermino8.VerticalAlignment = Element.ALIGN_MIDDLE;
            EtiquetaTermino8.FixedHeight = 24f;
            EtiquetaTermino8.BorderColor = new BaseColor(255, 255, 255);

            Terminos.AddCell(EtiquetaTerminos);

            Terminos.AddCell(Etiqueta1);
            Terminos.AddCell(EtiquetaTermino1);

            Terminos.AddCell(Etiqueta2);
            Terminos.AddCell(EtiquetaTermino2);

            Terminos.AddCell(Etiqueta3);
            Terminos.AddCell(EtiquetaTermino3);

            if(codigoCotizacion.aplica_contrato == true)
            {
                Terminos.AddCell(Etiqueta4);
                Terminos.AddCell(EtiquetaTermino4);

                Terminos.AddCell(Etiqueta5);
                Terminos.AddCell(EtiquetaTermino5);
                Terminos.AddCell(Etiqueta51);
                Terminos.AddCell(EtiquetaTermino51);
                Terminos.AddCell(Etiqueta52);
                Terminos.AddCell(EtiquetaTermino52);
                Terminos.AddCell(Etiqueta53);
                Terminos.AddCell(EtiquetaTermino53);
            }
             
            Terminos.AddCell(Etiqueta6);
            Terminos.AddCell(EtiquetaTermino6);

            Terminos.AddCell(Etiqueta7);
            Terminos.AddCell(EtiquetaTermino7);

            Terminos.AddCell(Etiqueta8);
            Terminos.AddCell(EtiquetaTermino8);

            terminosCondiciones.AddCell(Terminos);

            document.Add(terminosCondiciones);

            //Salto Linea      
            PdfPTable saltoLinea = new PdfPTable(4);

            PdfPCell EtiquetaSaltoLinea0 = new PdfPCell(new Phrase(" ", new Font(customfontbold, 9)));
            EtiquetaSaltoLinea0.Colspan = 4;
            EtiquetaSaltoLinea0.Border = Rectangle.NO_BORDER;
            EtiquetaSaltoLinea0.HorizontalAlignment = Element.ALIGN_CENTER;
            EtiquetaSaltoLinea0.VerticalAlignment = Element.ALIGN_MIDDLE;
            EtiquetaSaltoLinea0.FixedHeight = 8f;

            saltoLinea.AddCell(EtiquetaSaltoLinea0);

            document.Add(saltoLinea);

            //Head      
            PdfPTable head = new PdfPTable(7);
            head.DefaultCell.Border = Rectangle.NO_BORDER;

            //Titulo
            PdfPCell pagina = new PdfPCell(new Phrase("1 de 1", new Font(customfont, 5, 0, new BaseColor(60, 66, 82))));
            pagina.Colspan = 1;
            pagina.Border = Rectangle.NO_BORDER;
            pagina.HorizontalAlignment = Element.ALIGN_LEFT;
            pagina.VerticalAlignment = Element.ALIGN_MIDDLE;

            //Titulo
            PdfPCell head1 = new PdfPCell(new Phrase("GCM-GCO-PRO-001-FOR-004", new Font(customfont, 5, 0, new BaseColor(60, 66, 82))));
            head1.Colspan = 5;
            head1.Border = Rectangle.NO_BORDER;
            head1.HorizontalAlignment = Element.ALIGN_CENTER;
            head1.VerticalAlignment = Element.ALIGN_MIDDLE;

            //Titulo
            PdfPCell head2 = new PdfPCell(new Phrase("Versión 1.0", new Font(customfont, 5, 0, new BaseColor(60, 66, 82))));
            head2.Border = Rectangle.NO_BORDER;
            head2.HorizontalAlignment = Element.ALIGN_RIGHT;
            head2.VerticalAlignment = Element.ALIGN_MIDDLE;

            head.AddCell(pagina);
            head.AddCell(head1);
            head.AddCell(head2);

            document.Add(head);             

            //Cerrar Documento
            document.Close();
             
        }


        //public ActionResult DescargarReporteFormatoExcel()
        //{
        //    //Seleccionar las columnas a exportar
        //    var collection = CotizacionEntity.ListarCabecerasCotizacion();
        //    var package = new ExcelPackage();

        //    package = Reportes.ExportarExcel(collection.Cast<object>().ToList(), "Cotización"); //Reportes.ExportGenerico(collection.Cast<object>().ToList(), "Codigo Cotizacion");
        //    return File(package.GetAsByteArray(), XlsxContentType, "ListadoCotizacion.xlsx");
        //}

        #region Reportes Personalizados

        public ActionResult DescargarReporteFormatoExcel()
        {
            var collection = CotizacionEntity.ListarCabecerasCotizacion();

            ExcelPackage package = new ExcelPackage();
            var workSheet = package.Workbook.Worksheets.Add("Sheet1");
            workSheet.TabColor = System.Drawing.Color.Black;
            workSheet.DefaultRowHeight = 12;

            List<string> columnas = new List<string>() {
            "Código de Cotización",
                "Versión",
                "Estado Cotización",
                "Fecha de Cotización",
                "Fecha de Vencimiento",
                "Cliente",
                "Contacto",
                "Nombre del Proyecto",
                "Subtotal",
                "Descuento",
                "IVA",
                "Total",
                "Responsable",
                "Estado"};

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
                workSheet.Cells[recordIndex, 1].Value = item.Numero_Cotizacion;
                workSheet.Cells[recordIndex, 2].Value = item.Version;
                workSheet.Cells[recordIndex, 3].Value = item.EstadoCotizacion;
                workSheet.Cells[recordIndex, 4].Value = item.Fecha_Cotizacion.Value.ToString("yyyy/MM/dd");
                workSheet.Cells[recordIndex, 5].Value = item.Fecha_Vencimiento.Value.ToString("yyyy/MM/dd");
                workSheet.Cells[recordIndex, 6].Value = item.Cliente;
                workSheet.Cells[recordIndex, 7].Value = item.Ejecutivo;
                workSheet.Cells[recordIndex, 8].Value = item.NombreProyecto;
                workSheet.Cells[recordIndex, 9].Value = item.Subtotal;
                workSheet.Cells[recordIndex, 10].Value = item.Valor_Descuento;
                workSheet.Cells[recordIndex, 11].Value = (Convert.ToDouble(item.Subtotal) * 0.12);
                workSheet.Cells[recordIndex, 12].Value = item.Total;
                workSheet.Cells[recordIndex, 13].Value = item.Contacto;
                workSheet.Cells[recordIndex, 14].Value = item.Estado;
                recordIndex++;
            }

            for (int i = 1; i <= columnas.Count; i++)
            {
                workSheet.Column(i).AutoFit();
                workSheet.Column(i).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            }


            return File(package.GetAsByteArray(), XlsxContentType, "ListadoCotizacion.xlsx");
        }

        #endregion

        public ActionResult DescargarReporteFormatoPDF()
        {
            // Seleccionar las columnas a exportar
            var results = CotizacionEntity.ListarCabecerasCotizacionPDF();

            var list = Reportes.SerializeToJSON(results);
            return Content(list, "application/json");
        }

        public ActionResult DescargarReporteFormatoCSV()
        {
            var comlumHeadrs = new string[]
            {
                "Código de Cotización",
                "Versión",
                "Estado Cotización",
                "Fecha de Cotización",
                "Fecha de Vencimiento",
                "Cliente",
                "Contacto",
                "Nombre del Proyecto",
                "Subtotal",
                "Descuento",
                "IVA",
                "Total",
                "Responsable",
                "Estado",
            };

            var listado = (from item in CotizacionEntity.ListarCabecerasCotizacion()
                           select new object[]
                           {
                                            item.Numero_Cotizacion,
                                            $"{item.Version}",
                                            $"{item.EstadoCotizacion}",
                                             $"{item.Fecha_Cotizacion}",
                                              $"{item.Fecha_Vencimiento}",
                                               $"{item.Cliente}",
                                                $"{item.Ejecutivo}",
                                                 $"{item.NombreProyecto}",
                                                  $"{item.Subtotal}",
                                                  $"{item.Valor_Descuento}",
                                                 $"{item.Porcentaje_Iva}",
                                                 $"{item.Total}",
                                                 $"{item.Contacto}",
                                            $"\"{(item.Estado) }\"", //Escaping ","
                           }).ToList();

            // Build the file content
            var csv = new StringBuilder();
            listado.ForEach(line =>
            {
                csv.AppendLine(string.Join(",", line));
            });

            byte[] buffer = Encoding.Default.GetBytes($"{string.Join(",", comlumHeadrs)}\r\n{csv.ToString()}");
            return File(buffer, "text/csv", $"Cotizacion.csv");
        }

        public JsonResult ObtenerIVA(int id)
        {
            var impuestos = ImpuestoEntity.ConsultarImpuesto(id); 
            var data = impuestos.valor; 
            return Json(data, JsonRequestBehavior.AllowGet);
        }

        public ActionResult _EnviarCorreo(int? id)
        {
            ViewBag.TituloModal = "Enviar Cotización al Cliente";
            Cotizador cotizador = CotizacionEntity.ConsultarCotizadorCabecera(id.Value);
            return PartialView(cotizador);
        }

        [HttpPost]
        public ActionResult EnviarCorreo(Cotizador cotizacion, string cuerpo)
        {
            try
            {
                //Obtener el codigo de usuario que se logeo
                var usuarioSesion = ViewData["usuario"] = System.Web.HttpContext.Current.Session["usuario"];
                var idUsuario = Convert.ToInt16(usuarioSesion);
                var usuarioclave = ViewData["usuarioClave"] = System.Web.HttpContext.Current.Session["usuarioClave"];
                 
                RespuestaTransaccion resultado = CotizacionEntity.EnviarCorreo(cotizacion.id_cotizador, cuerpo, idUsuario, usuarioclave.ToString());
                return Json(new { Resultado = resultado }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                return Json(new { Resultado = new RespuestaTransaccion { Estado = false, Respuesta = ex.Message } }, JsonRequestBehavior.AllowGet);
            }
        }

    }
}

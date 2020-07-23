using DotNet.Highcharts;
using System.Drawing;
using DotNet.Highcharts.Enums;
using DotNet.Highcharts.Helpers;
using DotNet.Highcharts.Options;
using GestionPPM.Entidades.Metodos;
using GestionPPM.Entidades.Modelo;
using GestionPPM.Repositorios;
using Seguridad.Helper;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Web;
using System.Web.Mvc;
using Newtonsoft.Json;
using System.Text;
using NonFactors.Mvc.Grid;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using Newtonsoft.Json.Linq;
using System.Threading;
using System.Net;

namespace TemplateInicial.Controllers
{
    [Autenticado]
    public class SolicitudesClienteExternoController : BaseAppController
    {
        // GET: SolicitudesClienteExterno

        private static readonly GestionPPMEntities db = new GestionPPMEntities();
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
            ViewBag.NombreListado = Etiquetas.TituloPanelSolicitudRequerimientoClienteExterno;
            //var listado = SolicitudClienteExternoEntity.ListadoSolicitudClienteExterno();

            //var usuariosFiltro = OrganigramaEntity.GetEstructuraOrganigrama(2);

            // var filtros = JsonConvert.DeserializeObject<List<OrganigramaParcial>>(usuariosFiltro);


            var user = ViewData["usuario"] = System.Web.HttpContext.Current.Session["usuario"];

            var usuario = UsuarioEntity.ConsultarInformacionPrincipalUsuario(int.Parse(user.ToString()));

            bool uso_organigrama = usuario != null ? usuario.usar_organigrama : false;

            //filtros = filtros.Where(s => s.parent == user.ToString() || s.id == int.Parse(user.ToString())).ToList();

            //var usuarios = filtros.Select(s => s.id).ToList();

            var usuarios = OrganigramaEntity.GetHijosOrganigramaByUsuarioID(int.Parse(user.ToString()), 2);

            // El usuario principal puede ver todo
            //var listado = int.Parse(user.ToString()) != usuarioPrincipal ? SolicitudClienteExternoEntity.ListadoSolicitudClienteExterno().Where(s => usuarios.Contains(s.id_solicitante.Value)).ToList() : SolicitudClienteExternoEntity.ListadoSolicitudClienteExterno();
           
            //Controlar permisos
            var usuario2 = int.Parse(user.ToString());
            string nombreControlador = ControllerContext.RouteData.Values["controller"].ToString();
            ViewBag.NombreControlador = nombreControlador;


            ViewBag.AccionesUsuario = ManejoPermisosEntity.ListadoAccionesCatalogoUsuario(usuario2, nombreControlador);

            //Obtener Acciones del controlador
            ViewBag.AccionesControlador = GetMetodosControlador(nombreControlador);
            // El usuario principal puede ver todo
            var listado = new List<SolicitudClienteExternoInfo>();

            int usuarioPrincipal = int.Parse(ParametrosSistemaEntity.ConsultarParametros(2).valor);

            if (int.Parse(user.ToString()) == usuarioPrincipal || !uso_organigrama)
            {
                listado = SolicitudClienteExternoEntity.ListadoSolicitudClienteExterno();
            }
            else
            {
                //Primer filtro - Solicitudes del usuario propio
                //listado = SolicitudClienteExternoEntity.ListadoSolicitudClienteExterno().Where(s=> s.id_solicitante == int.Parse(user.ToString())).ToList();

                listado = SolicitudClienteExternoEntity.ListadoSolicitudClienteExterno().Where(s => usuarios.Contains(s.id_solicitante.Value) || s.id_solicitante == int.Parse(user.ToString())).ToList();

                // Si tiene dependientes filtrar también la de sus usuarios
                //if (usuarios.Count > 0) {
                //    listado = SolicitudClienteExternoEntity.ListadoSolicitudClienteExterno().Where(s => usuarios.Contains(s.id_solicitante.Value) || s.id_solicitante == int.Parse(user.ToString())).ToList();
                //}
            }

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


        //[ChildActionOnly]
        public async Task<PartialViewResult> _IndexGridComentarios(int? id, string search)
        {
            // trick to prevent deadlocks of calling async method 
            // and waiting for on a sync UI thread.
            //var syncContext = SynchronizationContext.Current;
            //SynchronizationContext.SetSynchronizationContext(null);

            ViewBag.NombreListado = Etiquetas.TituloGridActa;
            var listado = SolicitudClienteExternoEntity.ListadoComentarioSolicitud(id);

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

            // restore the context
            //SynchronizationContext.SetSynchronizationContext(syncContext);

            // Only grid query values will be available here.
            return PartialView(await Task.Run(() => listado));
        }

        private List<SolicitudClienteExternoInfo> GetListadoFiltrado()
        {
            List<SolicitudClienteExternoInfo> listado = new List<SolicitudClienteExternoInfo>();
            try
            {
                var user = ViewData["usuario"] = System.Web.HttpContext.Current.Session["usuario"];
                var usuarios = OrganigramaEntity.GetHijosOrganigramaByUsuarioID(int.Parse(user.ToString()), 2);
                int usuarioPrincipal = int.Parse(ParametrosSistemaEntity.ConsultarParametros(2).valor);

                if (int.Parse(user.ToString()) == usuarioPrincipal)
                    listado = SolicitudClienteExternoEntity.ListadoSolicitudClienteExterno();
                else
                    listado = SolicitudClienteExternoEntity.ListadoSolicitudClienteExterno().Where(s => usuarios.Contains(s.id_solicitante.Value) || s.id_solicitante == int.Parse(user.ToString())).ToList();

                return listado;
            }
            catch (Exception ex)
            {
                return listado;
            }
        }

        public ActionResult _Adjuntos(int? id)
        {
            ViewBag.TituloModal = "Repositorio de archivos adjuntos.";

            ViewBag.AdjuntosVacio = "No existen archivos adjuntos.";

            SolicitudClienteExternoInfo solicitud = SolicitudClienteExternoEntity.ConsultarSolicitudClienteExterno(id.Value);
            return PartialView(solicitud);
        }

        public ActionResult _Comentario(int? id)
        {
            SolicitudClienteExternoInfo solicitud = SolicitudClienteExternoEntity.ConsultarSolicitudClienteExterno(id.Value);
            ViewBag.TituloModal = "Comentarios - Solicitud " + solicitud.CodigoSolicitud;
            int maxcomentarios = int.Parse(ParametrosSistemaEntity.ConsultarParametros(9).valor);
            ViewBag.MaximoComentarios = " Límite de comentarios por solicitud ("+ maxcomentarios + ").";
            ViewBag.SolicitudID = solicitud.id_solicitud;

            //SolicitudClienteExternoInfo solicitud = SolicitudClienteExternoEntity.ConsultarSolicitudClienteExterno(id.Value);
            List<ComentariosSolicitudInfo> comentarios = SolicitudClienteExternoEntity.ListadoComentarioSolicitud(id);
            return PartialView(comentarios);
        }

        public ActionResult GuardarComentario(int? id)
        {
            string FileName = "";
            try
            {
                int maxcomentarios = int.Parse(ParametrosSistemaEntity.ConsultarParametros(9).valor);
                int? result = db.ConsultarCantidadComentariosSolictud(id).FirstOrDefault();
                if (result > maxcomentarios-1)
                {
                    return Json(new { Resultado = new RespuestaTransaccion { Estado = false, Respuesta = "Límite (" + maxcomentarios + ") de comentarios a la solicitud excedido." } }, JsonRequestBehavior.AllowGet);
                }

                var comentarioSolicitud = HttpContext.Request.Params.Get("Comentario");
                dynamic data = JObject.Parse(comentarioSolicitud);
                string comentario = Convert.ToString(data.Comentario);

                if (comentario == "" || comentario.Length == 0)
                {
                    return Json(new { Resultado = new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeDatosObligatorios } }, JsonRequestBehavior.AllowGet);
                }
                else
                {

                    ComentarioSolicitud comentariosSolicitudes = new ComentarioSolicitud();

                    //Obtener el codigo de usuario que se logeo
                    var usuarioSesion = ViewData["usuario"] = System.Web.HttpContext.Current.Session["usuario"];
                    var idUsuario = Convert.ToInt16(usuarioSesion);

                    HttpFileCollectionBase files = Request.Files;

                    if (files.Count == 0)
                    {
                        comentariosSolicitudes = new ComentarioSolicitud
                        {
                            Comentario = comentario,
                            Estado = true,
                            Fecha = DateTime.Now,
                            SolicitudID = id.Value,
                            NombreArchivo = null,
                            TipoArchivo = null,
                            UsuarioId = idUsuario
                        };
                    }
                    else
                    {
                        for (int i = 0; i < files.Count; i++)
                        {
                            HttpPostedFileBase file = files[i];
                            string path = string.Empty;

                            // Checking for Internet Explorer    
                            if (Request.Browser.Browser.ToUpper() == "IE" || Request.Browser.Browser.ToUpper() == "INTERNETEXPLORER")
                            {
                                string[] testfiles = file.FileName.Split(new char[] { '\\' });
                                path = testfiles[testfiles.Length - 1];
                            }
                            else
                            {
                                path = file.FileName;
                                FileName = file.FileName;
                            }

                            string tipoArchivo = file.ContentType;

                            comentariosSolicitudes = new ComentarioSolicitud
                            {
                                Comentario = comentario,
                                Estado = true,
                                Fecha = DateTime.Now,
                                SolicitudID = id.Value,
                                NombreArchivo = FileName,
                                TipoArchivo = tipoArchivo,
                                UsuarioId = idUsuario
                            };

                            using (var reader = new BinaryReader(file.InputStream))
                            {
                                comentariosSolicitudes.ArchivoAdjunto = reader.ReadBytes(file.ContentLength);
                            }
                        }
                    }
                    RespuestaTransaccion resultado = SolicitudClienteExternoEntity.CrearComentarioSolicitud(comentariosSolicitudes);

                    return Json(new { Resultado = resultado }, JsonRequestBehavior.AllowGet);
                }
            }
            catch (Exception ex)
            {
                return Json(new { Resultado = new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeAdjuntoFallido + ex.Message } }, JsonRequestBehavior.AllowGet);
            }
        }

        [HttpPost]
        public ActionResult EliminarComentario(int id)
        {
            if (id == 0)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            RespuestaTransaccion resultado = SolicitudClienteExternoEntity.EliminarComentarioSolicitud(id);

            //return RedirectToAction("Index");
            return Json(new { Resultado = resultado }, JsonRequestBehavior.AllowGet);
        }

        public ActionResult _AdjuntarArchivos(int? id)
        {
            ViewBag.TituloModal = "Adjuntar archivos";

            ViewBag.AdjuntosVacio = "No existen archivos adjuntos.";

            SolicitudClienteExternoInfo solicitud = SolicitudClienteExternoEntity.ConsultarSolicitudClienteExterno(id.Value);
            return PartialView(solicitud);
        }

        public ActionResult _AdjuntosActas(int? id)
        {
            ViewBag.TituloModal = "Archivos Adjuntos";

            ViewBag.AdjuntosVacio = "No existen archivos adjuntos.";

            SolicitudClienteExternoInfo solicitud = SolicitudClienteExternoEntity.ConsultarSolicitudClienteExterno(id.Value);
            return PartialView(solicitud);
        }

        public ActionResult _AdjuntosCotizaciones(int? id)
        {
            ViewBag.TituloModal = "Archivos Adjuntos";

            ViewBag.AdjuntosVacio = "No existen archivos adjuntos.";

            SolicitudClienteExternoInfo solicitud = SolicitudClienteExternoEntity.ConsultarSolicitudClienteExterno(id.Value);
            return PartialView(solicitud);
        }

        public ActionResult _AprobacionCotizacion(int? id)
        {
            AprobacionCotizacion aprobacion = CotizacionEntity.ConsultarCodigoCotizacion(id.Value);
            var codigoCotizacion = CodigoCotizacionEntity.ConsultarCodigoCotizacion(id.Value);
            ViewBag.TituloModal = "Aprobación - " + codigoCotizacion.codigo_cotizacion;

            ViewBag.codigoCotizacion = codigoCotizacion;
            return PartialView(aprobacion);
        }

        [HttpPost]
        public ActionResult AprobarCotizacion(AprobacionCotizacion codigoCotizacion)
        {
            try
            {
                RespuestaTransaccion resultado = CodigoCotizacionEntity.ActualizarStatusCodigoCotizacionClienteExterno(codigoCotizacion.CodigoCotizacionID, codigoCotizacion.estatus_codigo);

                resultado = CotizacionEntity.AprobarCotizacion(codigoCotizacion);

                return Json(new { Resultado = resultado }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                return Json(new { Resultado = new RespuestaTransaccion { Estado = false, Respuesta = ex.Message } }, JsonRequestBehavior.AllowGet);
            }
        }

        public ActionResult _TrackingProyecto(int? id, string nombreProyecto)
        {
            ViewBag.TituloModal = "Tracking Proyecto - " + nombreProyecto;

            ViewBag.InformacionVacia = "No existe información.";

            ViewBag.nombreProyecto = nombreProyecto;

            var historial = SolicitudClienteExternoEntity.ListadoDetalleHistorialAsignacionProyectos(id.Value);


            List<object> elementosY1 = new List<object>();

            foreach (var item in historial)
            {
                elementosY1.Add(item.porcentaje_avance);
            }

            List<object> elementosY2 = new List<object>();

            foreach (var item in historial)
            {
                elementosY2.Add(item.porcentaje_cumplimiento);
            }

            List<string> elementosX = new List<string>();

            int numeroAvance = 1;
            foreach (var item in historial)
            {
                //elementosX.Add(item.id_detalle_proyecto.ToString());
                elementosX.Add(numeroAvance.ToString());
                numeroAvance++;
            }

            var series1 = new Series { Color = Color.Green, Name = "Porcentaje Avance", Data = new Data(elementosY1.ToArray()) }; // Porcentaje avance
            var series2 = new Series { Color = Color.Green, Name = "Porcentaje Cumplimiento", Data = new Data(elementosY2.ToArray()) }; // Porcentaje Real cumplimiento

            //Highcharts chart = new Highcharts("chart")
            //    .InitChart(new Chart { DefaultSeriesType = ChartTypes.Column, ZoomType = ZoomTypes.Xy })
            //    .SetTitle(new Title { Text = nombreProyecto })
            //    .SetSubtitle(new Subtitle { Text = "Tracking" })
            //    .SetXAxis(new XAxis { Categories = elementosX.ToArray() })
            //    .SetYAxis(new YAxis
            //    {
            //        Min = 0,
            //        Max = 100,
            //        Title = new YAxisTitle { Text = "Porcentaje" }
            //    })
            //    .SetLegend(new Legend
            //    {
            //        Layout = Layouts.Vertical,
            //        Align = HorizontalAligns.Left,
            //        VerticalAlign = VerticalAligns.Top,
            //        X = 100,
            //        Y = 70,
            //        Floating = true,
            //        BackgroundColor = new BackColorOrGradient(ColorTranslator.FromHtml("#FFFFFF")),
            //        Shadow = true
            //    })
            //    .SetTooltip(new Tooltip { Formatter = @"function() { return 'Avance: '+ this.y +' %'; }" })//@"function() { return ''+ this.x +': '+ this.y +' %'; }" })
            //    .SetPlotOptions(new PlotOptions
            //    {
            //        Column = new PlotOptionsColumn
            //        {
            //            PointPadding = 0.2,
            //            BorderWidth = 0
            //        }
            //    })
            //    .SetCredits(new Credits { Enabled = false, /*Text = "https://www.ppm.com.ec/"*/ })
            //    .SetExporting(new Exporting { Enabled = true })
            //    .SetSeries(new[]
            //        {
            //        series1,
            //        //series2,
            //            //new Series { Color = Color.Tomato  , Name = "Tokyo", Data = new Data(new object[] { 49.9, 71.5, 10.4, 29.2, 44.0, 76.0, 35.6, 48.5, 16.4, 94.1, 95.6, 54.4 }) },
            //            //new Series { Name = "London", Data = new Data(new object[] { 48.9, 38.8, 39.3, 41.4, 47.0, 48.3, 59.0, 59.6, 52.4, 65.2, 59.3, 51.2 }) },
            //            //new Series { Name = "New York", Data = new Data(new object[] { 83.6, 78.8, 98.5, 93.4, 106.0, 84.5, 105.0, 104.3, 91.2, 83.5, 106.6, 92.3 }) },
            //            //new Series { Name = "Berlin", Data = new Data(new object[] { 42.4, 33.2, 34.5, 39.7, 52.6, 75.5, 57.4, 60.4, 47.6, 39.1, 46.8, 51.1 }) }
            //        }
            //    );
            //Highcharts chart = new Highcharts("chart")
            //    .InitChart(new Chart
            //    {
            //        Type = ChartTypes.Gauge,
            //        PlotBackgroundColor = null,
            //        PlotBackgroundImage = null,
            //        PlotBorderWidth = 0,
            //        PlotShadow = false
            //    })
            //    .SetTitle(new Title { Text = "Speedometer" })
            //    .SetPane(new Pane
            //    {
            //        StartAngle = -150,
            //        EndAngle = 150,
            //        Background = new[]
            //                {
            //                    new BackgroundObject
            //                        {
            //                            BackgroundColor = new BackColorOrGradient(new Gradient
            //                                {
            //                                    LinearGradient = new[] { 0, 0, 0, 1 },
            //                                    Stops = new object[,] { { 0, "#FFF" }, { 1, "#333" } }
            //                                }),
            //                            BorderWidth = new PercentageOrPixel(0),
            //                            OuterRadius = new PercentageOrPixel(109, true)
            //                        },
            //                    new BackgroundObject
            //                        {
            //                            BackgroundColor = new BackColorOrGradient(new Gradient
            //                                {
            //                                    LinearGradient = new[] { 0, 0, 0, 1 },
            //                                    Stops = new object[,] { { 0, "#333" }, { 1, "#FFF" } }
            //                                }),
            //                            BorderWidth = new PercentageOrPixel(1),
            //                            OuterRadius = new PercentageOrPixel(107, true)
            //                        },
            //                    new BackgroundObject(),
            //                    new BackgroundObject
            //                        {
            //                            BackgroundColor = new BackColorOrGradient(ColorTranslator.FromHtml("#DDD")),
            //                            BorderWidth = new PercentageOrPixel(0),
            //                            OuterRadius = new PercentageOrPixel(105, true),
            //                            InnerRadius = new PercentageOrPixel(103, true)
            //                        }
            //                }
            //    })
            //    .SetYAxis(new YAxis
            //    {
            //        Min = 0,
            //        Max = 200,

            //        //MinorTickInterval = "auto",
            //        MinorTickWidth = 1,
            //        MinorTickLength = 10,
            //        MinorTickPosition = TickPositions.Inside,
            //        MinorTickColor = ColorTranslator.FromHtml("#666"),
            //        TickPixelInterval = 30,
            //        TickWidth = 2,
            //        TickPosition = TickPositions.Inside,
            //        TickLength = 10,
            //        TickColor = ColorTranslator.FromHtml("#666"),
            //        Labels = new YAxisLabels
            //        {
            //            Step = 2,
            //            //Rotation = "auto"
            //        },
            //        Title = new YAxisTitle { Text = "km/h" },
            //        PlotBands = new[]
            //                {
            //                    new YAxisPlotBands { From = 0, To = 120, Color = ColorTranslator.FromHtml("#55BF3B") },
            //                    new YAxisPlotBands { From = 120, To = 160, Color = ColorTranslator.FromHtml("#DDDF0D") },
            //                    new YAxisPlotBands { From = 160, To = 200, Color = ColorTranslator.FromHtml("#DF5353") }
            //                }
            //    })
            //    .SetSeries(new Series
            //    {
            //        Name = "Speed",
            //        Data = new Data(new object[] { 80 })
            //    });
            Highcharts chart = new Highcharts("chart")
    .InitChart(new Chart
    {
        Type = ChartTypes.Gauge,
        AlignTicks = false,
        PlotBackgroundColor = null,
        PlotBackgroundImage = null,
        PlotBorderWidth = 0,
        PlotShadow = false
    })
    .SetTitle(new Title { Text = "Speedometer with dual axes" })
    .SetTooltip(new Tooltip { ValueSuffix = " km/h" })
    .SetPane(new Pane
    {
        StartAngle = -150,
        EndAngle = 150
    })
    .SetYAxis(new[]
        {
                        new YAxis
                            {
                                Min = 0,
                                Max = 200,
                                LineColor = ColorTranslator.FromHtml("#339"),
                                TickColor = ColorTranslator.FromHtml("#339"),
                                MinorTickColor = ColorTranslator.FromHtml("#339"),
                                Offset = -25,
                                LineWidth = 2,
                                TickLength = 5,
                                MinorTickLength = 5,
                                EndOnTick = false,
                                Labels = new YAxisLabels { Distance = -20 }
                            },
                        new YAxis
                            {
                                Min = 0,
                                Max = 124,
                                TickPosition = TickPositions.Outside,
                                LineColor = ColorTranslator.FromHtml("#933"),
                                LineWidth = 2,
                                MinorTickPosition = TickPositions.Outside,
                                TickColor = ColorTranslator.FromHtml("#933"),
                                MinorTickColor = ColorTranslator.FromHtml("#933"),
                                TickLength = 5,
                                MinorTickLength = 5,
                                Offset = -20,
                                EndOnTick = false,
                                Labels = new YAxisLabels { Distance = 12 }
                            }
        }
    )
    .SetPlotOptions(new PlotOptions
    {
        Gauge = new PlotOptionsGauge
        {
            DataLabels = new PlotOptionsGaugeDataLabels
            {
                Formatter = @"function() {
                                                        var kmh = this.y,
	                                                    mph = Math.round(kmh * 0.621);
	                                                    return '<span style=color:#339>'+ kmh + ' km/h</span><br/><span style=color:#933>' + mph + ' mph</span>';
                                                    }",
                BackgroundColor = new BackColorOrGradient(
                        new Gradient
                        {
                            LinearGradient = new[] { 0, 0, 0, 1 },
                            Stops = new object[,] { { 0, "#DDD" }, { 1, "#FFF" } }
                        })
            }
        }
    })
    .SetSeries(new Series
    {
        Name = "Speed",
        Data = new Data(new object[] { 80 }),
    });
            ViewBag.Grafico = chart;

            TrackingAsignacionProyectoInfo trackingProyecto = SolicitudClienteExternoEntity.ConsultarTrackingProyecto(id.Value);

            List<decimal> intervalos = new List<decimal> { 0 };

            if (trackingProyecto != null)
            {
                decimal horas = (trackingProyecto.numero_horas_real ?? 0) / 6;

                int totalIntervalos = 6;
                var intervalo = 0m; // decimal (money)

                for (int i = 0; i < totalIntervalos; i++)
                {
                    intervalo += horas;
                    intervalos.Add(Math.Round(intervalo, 2));
                }
            }

            ViewBag.IntervalosHoras = intervalos;

            return PartialView(trackingProyecto);
        }

        public ActionResult _HistorialDetalleAsignacionProyecto(int? id)
        {
            ViewBag.TituloModal = "Historial Detalle de Asignación de Proyecto";

            var historial = SolicitudClienteExternoEntity.ListadoDetalleHistorialAsignacionProyectos(id.Value);

            var codigoCotizacion = historial.FirstOrDefault();
            ViewBag.codigoCotizacion = codigoCotizacion != null ? codigoCotizacion.codigo_cotizacion : string.Empty;

            return PartialView(historial);
        }

        public JsonResult _GetArchivosAdjuntosCompletos(int? idSolicitud, string codigoCotizacion)
        {
            List<TreeViewJQueryUI> items = new List<TreeViewJQueryUI>();
            try
            {
                string basePath = ConfigurationManager.AppSettings["RepositorioDocumentos"];

                string rutaArchivosAdjuntosSolicitudes = basePath + "\\GESTION_PPM\\ADJUNTOS_SOLICITUDES";
                string rutaArchivosAdjuntosActas = basePath + "\\GESTION_PPM\\ACTAS";
                string rutaArchivosAdjuntosCotizaciones = basePath + "\\GESTION_PPM";

                List<string> rutas = new List<string> { rutaArchivosAdjuntosSolicitudes, rutaArchivosAdjuntosActas, rutaArchivosAdjuntosCotizaciones };

                foreach (var ruta in rutas)
                {
                    bool directorio = Directory.Exists(ruta);

                    if (!directorio)
                        Directory.CreateDirectory(ruta);

                    string carpeta = Path.GetFullPath(ruta);
                    string carpeta2 = Path.GetPathRoot(ruta);

                    //Solicitudes
                    if (carpeta.Contains("ADJUNTOS_SOLICITUDES"))
                        items.AddRange(getFolderNodes(ruta));

                    //Actas
                    if (carpeta.Contains("ACTAS"))
                        items.AddRange(getFolderNodesPDF(ruta));

                    //Cotizaciones
                    //if (carpeta.Contains("GESTION_PPM")) {
                    //    var files = Directory.GetFiles(ruta, "*.*", SearchOption.TopDirectoryOnly).Where(s => s.EndsWith(".pdf") && s.Contains(codigoCotizacion)).ToList();
                    //    items.AddRange(GetSoloArchivosEnDirectorio(files));
                    //}
                }

                foreach (var item in items)
                {
                    List<TreeViewJQueryUI> listaFiltrada = new List<TreeViewJQueryUI>();
                    foreach (var clienteFolder in item.children) // Clientes se encuentra en el segundo nivel
                    {

                        if (clienteFolder.text == codigoCotizacion.ToString())
                            listaFiltrada.Add(clienteFolder);
                    }
                    item.children = listaFiltrada;

                    item.children = item.children.Where(s => s.children.Any()).ToList();
                }

                items = items.Where(s => s.children.Any()).ToList(); // Todos los que contengan elementos.

                return Json(items, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                return Json(items, JsonRequestBehavior.AllowGet);
            }
        }

        public JsonResult _GetArchivosAdjuntos(string id)
        {
            List<TreeViewJQueryUI> items = new List<TreeViewJQueryUI>();
            if (!string.IsNullOrEmpty(id))
            {
                string basePath = ConfigurationManager.AppSettings["RepositorioDocumentos"];
                string rutaArchivos = basePath + "\\GESTION_PPM\\ADJUNTOS_SOLICITUDES";

                var RootDirectory = new DirectoryInfo(rutaArchivos);
                var directorio = RootDirectory.GetDirectories("*", SearchOption.AllDirectories).Where(s => s.Name.Equals(id)).FirstOrDefault();

                string pathCompletoDirectorio = directorio != null ? directorio.FullName : string.Empty;

                var files = !string.IsNullOrEmpty(pathCompletoDirectorio) ? Directory.GetFiles(pathCompletoDirectorio, "*.*", SearchOption.TopDirectoryOnly).ToList() : new List<string>();

                items = GetSoloArchivosEnDirectorio(files);

            }
            return Json(items, JsonRequestBehavior.AllowGet);
        }

        public JsonResult _GetArchivosAdjuntosActas(string id)
        {
            List<TreeViewJQueryUI> items = new List<TreeViewJQueryUI>();
            if (!string.IsNullOrEmpty(id))
            {

                string basePath = ConfigurationManager.AppSettings["RepositorioDocumentos"];
                string rutaArchivos = basePath + "\\GESTION_PPM\\ACTAS";

                var RootDirectory = new DirectoryInfo(rutaArchivos);
                var directorio = RootDirectory.GetDirectories("*", SearchOption.AllDirectories).Where(s => s.Name.Equals(id)).FirstOrDefault();

                string pathCompletoDirectorio = directorio != null ? directorio.FullName : string.Empty;

                var files = !string.IsNullOrEmpty(pathCompletoDirectorio) ? Directory.GetFiles(pathCompletoDirectorio, "*.*", SearchOption.TopDirectoryOnly).Where(s => s.EndsWith(".pdf")).ToList() : new List<string>();

                items = GetSoloArchivosEnDirectorio(files);
            }
            return Json(items, JsonRequestBehavior.AllowGet);
        }

        public JsonResult _GetArchivosAdjuntosCotizaciones(string id)
        {
            List<TreeViewJQueryUI> items = new List<TreeViewJQueryUI>();
            if (!string.IsNullOrEmpty(id))
            {

                string basePath = ConfigurationManager.AppSettings["RepositorioDocumentos"];
                string rutaArchivos = basePath + "\\GESTION_PPM";

                bool directorio = Directory.Exists(rutaArchivos);//Directory.Exists(Server.MapPath("~/ArchivosAdjuntosCotizaciones/"));
                // En caso de que no exista el directorio, crearlo.
                if (!directorio)
                    Directory.CreateDirectory(rutaArchivos);//Directory.CreateDirectory(Server.MapPath("~/ArchivosAdjuntosCotizaciones/"));

                //items = getFolderNodes(rutaArchivos, true);//items = getFolderNodes(Server.MapPath("~/ArchivosAdjuntosCotizaciones/"));

                var files = Directory.GetFiles(rutaArchivos, "*.*", SearchOption.TopDirectoryOnly).Where(s => s.EndsWith(".pdf") && s.Contains(id)).ToList();

                items = GetSoloArchivosEnDirectorio(files);


            }
            return Json(items, JsonRequestBehavior.AllowGet);
        }

        List<TreeViewJQueryUI> GetSoloArchivosEnDirectorio(List<string> dirs)
        {
            var nodes = new List<TreeViewJQueryUI>();
            foreach (string d in dirs)
            {
                var extensionArchivo = Path.GetExtension(d);
                var icono = Auxiliares.GetIconoExtension(extensionArchivo);

                DirectoryInfo di = new DirectoryInfo(d);
                TreeViewJQueryUI tn = new TreeViewJQueryUI(di.Name);
                tn.desc = di.FullName;
                tn.esCarpeta = false;
                tn.id = Guid.NewGuid();
                tn.children = null;
                tn.iconCls = icono;

                nodes.Add(tn);
            }
            return nodes;
        }

        List<TreeViewJQueryUI> getAllFolderNodes(string dir)
        {
            var dirs = Directory.GetDirectories(dir).ToArray();
            var nodes = new List<TreeViewJQueryUI>();
            foreach (string d in dirs)
            {
                DirectoryInfo di = new DirectoryInfo(d);
                TreeViewJQueryUI tn = new TreeViewJQueryUI(di.Name);
                tn.desc = di.FullName;
                tn.esCarpeta = true;
                tn.id = Guid.NewGuid();

                //tn.text = Path.GetFileName(di.FullName);// di.Name+ "." +di.Extension;
                int subCount = 0;
                try { subCount = Directory.GetDirectories(d).Count(); }
                catch { /* ignore accessdenied */  }
                if (subCount > 0)
                {
                    var subNodes = getAllFolderNodes(di.FullName);
                    tn.children.AddRange(subNodes.ToList());
                }
                nodes.Add(tn);
            }
            return nodes;
        }

        List<TreeViewJQueryUI> getFolderNodes(string dir, bool flag = false)
        {
            var dirs = Directory.GetDirectories(dir).ToArray();
            var nodes = new List<TreeViewJQueryUI>();
            foreach (string d in dirs)
            {
                DirectoryInfo di = new DirectoryInfo(d);
                TreeViewJQueryUI tn = new TreeViewJQueryUI(di.Name);
                tn.desc = di.FullName;
                tn.id = Guid.NewGuid();
                int subCount = 0;
                try { subCount = Directory.GetDirectories(d).Count(); }
                catch { /* ignore accessdenied */  }
                if (subCount > 0)
                {
                    var subnodes = getAllFolderNodes(di.FullName);
                    foreach (var item in subnodes)
                    {
                        string[] filePaths = Directory.GetFiles(item.desc, "*.*", SearchOption.AllDirectories);
                        foreach (var archivo in filePaths)
                        {
                            var extensionArchivo = Path.GetExtension(archivo);
                            var icono = Auxiliares.GetIconoExtension(extensionArchivo);

                            item.children.Add(new TreeViewJQueryUI
                            {
                                desc = archivo,
                                text = Path.GetFileName(archivo),
                                id = Guid.NewGuid(),
                                children = null,
                                iconCls = icono,
                                esCarpeta = false,
                            });



                        }
                    }
                    tn.children.AddRange(subnodes.ToList());
                }
                nodes.Add(tn);
            }
            return nodes;
        }

        List<TreeViewJQueryUI> getFolderNodesPDF(string dir)
        {
            var dirs = Directory.GetDirectories(dir).ToArray();
            var nodes = new List<TreeViewJQueryUI>();
            foreach (string d in dirs)
            {
                DirectoryInfo di = new DirectoryInfo(d);
                TreeViewJQueryUI tn = new TreeViewJQueryUI(di.Name);
                tn.desc = di.FullName;
                tn.id = Guid.NewGuid();
                int subCount = 0;
                try { subCount = Directory.GetDirectories(d).Count(); }
                catch { /* ignore accessdenied */  }
                if (subCount > 0)
                {
                    var subnodes = getAllFolderNodes(di.FullName);
                    foreach (var item in subnodes)
                    {
                        string[] filePaths = Directory.GetFiles(item.desc, "*.*", SearchOption.AllDirectories);
                        foreach (var archivo in filePaths)
                        {
                            var extensionArchivo = Path.GetExtension(archivo);
                            var icono = Auxiliares.GetIconoExtension(extensionArchivo);
                            if (extensionArchivo == ".pdf")
                            {
                                item.children.Add(new TreeViewJQueryUI
                                {
                                    desc = archivo,
                                    text = Path.GetFileName(archivo),
                                    id = Guid.NewGuid(),
                                    children = null,
                                    iconCls = icono,
                                    esCarpeta = false,
                                });
                            }
                        }
                    }
                    tn.children.AddRange(subnodes.ToList());
                }
                nodes.Add(tn);
            }
            return nodes;
        }

        public ActionResult DescargarArchivoAdjunto(string path)
        {
            byte[] fileBytes = System.IO.File.ReadAllBytes(path);
            string fileName = Path.GetFileName(path);
            return File(fileBytes, System.Net.Mime.MediaTypeNames.Application.Octet, fileName);
        }

        public ActionResult DescargarArchivo(int? id)
        {
            var comentario = SolicitudClienteExternoEntity.ConsultarComentarioSolicitud(id.Value);

            byte[] fileBytes = comentario.ArchivoAdjunto;//System.IO.File.ReadAllBytes(path);
            string fileName = comentario.NombreArchivo;
            return File(fileBytes, System.Net.Mime.MediaTypeNames.Application.Octet, fileName);
        }


        [HttpPost]
        public ActionResult GuardarRespuestComentario(int idComentario, string respuesta)
        {
            try
            {
                int contador = respuesta.Count();
                int maxcomentarios = int.Parse(ParametrosSistemaEntity.ConsultarParametros(10).valor);
                int? result = db.ConsultarCantidadRespuestaComentario(idComentario).FirstOrDefault();
                if (result > maxcomentarios - 1)
                {
                    return Json(new { Resultado = new RespuestaTransaccion { Estado = false, Respuesta = "Límite (" + maxcomentarios + ") de respuestas al comentario excedido." } }, JsonRequestBehavior.AllowGet);
                }
                if (respuesta == "")
                {
                    return Json(new { Resultado = new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeRespuestaRequerida } }, JsonRequestBehavior.AllowGet);
                }
                if (contador > 50)
                {
                    return Json(new { Resultado = new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeRespuestaLimiteCaracteres } }, JsonRequestBehavior.AllowGet);
                }

                var user = ViewData["usuario"] = System.Web.HttpContext.Current.Session["usuario"];

                RespuestaComentario comentario = new RespuestaComentario();
                comentario.id_comentario = idComentario;
                comentario.Respuesta = respuesta;
                comentario.id_usuario = Convert.ToInt32(user);
                comentario.Fecha = DateTime.Now;
                comentario.Estado = true;

                RespuestaTransaccion resultado = SolicitudClienteInternoEntity.CrearRespuestaComentarioSolicitud(comentario);

                return Json(new { Resultado = resultado }, JsonRequestBehavior.AllowGet);

                //return Json(new { Resultado = new RespuestaTransaccion { Estado = true, Respuesta = "Ok!" } }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                return Json(new { Resultado = new RespuestaTransaccion { Estado = false, Respuesta = ex.Message } }, JsonRequestBehavior.AllowGet);
            }
        }

        public ActionResult _ImprimirSolicitud(int id)
        {
            string basePath = ConfigurationManager.AppSettings["RepositorioDocumentos"];
            string rutaArchivos = basePath + "\\GESTION_PPM\\ADJUNTOS_SOLICITUDES";



            string codigoSolicitud = id.ToString().PadLeft(6, '0'); //Codigo de solicitud formado por 6 digitos



            string nombreArchivo = "Solicitud-{0}.pdf";



            nombreArchivo = string.Format(nombreArchivo, codigoSolicitud);



            rutaArchivos = Directory.GetFiles(rutaArchivos, nombreArchivo, SearchOption.AllDirectories).FirstOrDefault();



            //Mover archivos 
            //Obtener Ruta PDF
            string path = string.Empty;
            string controllerName = ControllerContext.RouteData.Values["controller"].ToString();
            var absolutePath = HttpContext.Server.MapPath(path);
            absolutePath = absolutePath.Replace(controllerName, "Solicitud\\" + nombreArchivo);



            System.IO.File.Copy(rutaArchivos, absolutePath, true);



            path = "../Solicitud/";
            ViewBag.Archivo = path + nombreArchivo;



            return PartialView();
        }

        #region Reportes
        [HttpGet]
        public ActionResult DescargarReporteFormatoExcel()
        {
            // Using EPPlus from nuget
            using (ExcelPackage package = new ExcelPackage())
            {
                Int32 row = 2;
                Int32 col = 1;

                package.Workbook.Worksheets.Add("Listado");
                IGrid<SolicitudClienteExternoInfo> grid = CreateExportableGrid();
                ExcelWorksheet sheet = package.Workbook.Worksheets["Listado"];

                foreach (IGridColumn column in grid.Columns)
                {
                    sheet.Cells[1, col].Value = column.Title;
                    sheet.Column(col++).Width = 18;

                    column.IsEncoded = false;
                }

                foreach (IGridRow<SolicitudClienteExternoInfo> gridRow in grid.Rows)
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

                #region Resumen Solicitantes

                package.Workbook.Worksheets.Add("Solicitantes");
                ExcelWorksheet Hoja2 = package.Workbook.Worksheets["Solicitantes"];

                List<string> columnas = new List<string> { "Solicitante", "Tipo", "Área Solicitante", "Total Solicitudes" };

                var ResumenSolicitudesSolicitante = SolicitudClienteExternoEntity.ListadoResumenSolicitudesSolicitantes();

                var i = 1;
                foreach (var item in columnas)
                {
                    Hoja2.Column(i).Width = 40;
                    Hoja2.Cells[1, i].Value = item;
                    Hoja2.Cells[1, i].Style.Font.Bold = true;
                    //Hoja2.Cells[1, i].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(23, 55, 93));

                    CambiarColorFila(Hoja2, 1, columnas.Count, Color.Orange);

                    i++;
                }

                int fila = 2;
                foreach (var item in ResumenSolicitudesSolicitante)
                {
                    var objeto = Auxiliares.GetValoresCamposObjeto(item);
                    int columna = 1;
                    foreach (var valor in objeto)
                    {
                        Hoja2.Cells[fila, columna].Value = valor;
                        columna++;
                    }
                    fila++;
                }

                #endregion 

                return File(package.GetAsByteArray(), "application/unknown", "ListadoClienteExterno.xlsx");
            }
        }

        private static void CambiarColorFila(ExcelWorksheet hoja, int fila, int totalColumnas, Color color)
        {
            for (int i = 1; i <= totalColumnas; i++)
            {
                using (ExcelRange rowRange = hoja.Cells[fila, i])
                {
                    rowRange.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    rowRange.Style.Fill.BackgroundColor.SetColor(color);
                }
            }
        }

        private IGrid<SolicitudClienteExternoInfo> CreateExportableGrid()
        {
            IGrid<SolicitudClienteExternoInfo> grid = new Grid<SolicitudClienteExternoInfo>(GetListadoFiltrado());
            grid.ViewContext = new ViewContext { HttpContext = HttpContext };
            grid.Query = Request.QueryString;

            grid.Columns.Add(model => model.fecha_hora_solicitud).Titled("Fecha").Formatted("{0:d}");
            grid.Columns.Add(model => model.TiempoTranscurrido).Titled("Tiempo Transcurrido(Horas)").Formatted("{0:d}");
            grid.Columns.Add(model => model.CodigoSolicitud).Titled("N.");
            grid.Columns.Add(model => model.NombresCompletosSolicitante).Titled("Ejecutivo/a");
            grid.Columns.Add(model => model.TextoCatalogoTipo).Titled("Tipo de Solicitud");
            grid.Columns.Add(model => model.TextoCatalogoSubTipo).Titled("Subtipo de Solicitud");
            grid.Columns.Add(model => model.mkt).Titled("Referencia Cliente");
            grid.Columns.Add(model => model.codigo_cotizacion).Titled("Código de Cotización");
            grid.Columns.Add(model => model.NombreProyectoSolicituCliente).Titled("Nombre de Proyecto");
            grid.Columns.Add(model => model.VersionCodigoCotizacion).Titled("Versión");
            grid.Columns.Add(model => model.EstatusCodigo).Titled("Estatus");
            grid.Columns.Add(model => model.TextoCatalogoMarca).Titled("Marca");
            grid.Columns.Add(model => model.cantidad).Titled("Cantidad");

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

        public ActionResult DescargarReporteFormatoCSV()
        {
            var comlumHeadrs = new string[]
            {

                "FECHA",
                "TIEMPO TRANSCURRIDO",
                "CODIGO DE SOLICITUD",
                "EJECUTIVO/A",
                "TIPO DE SOLICITUD",
                "SUBTIPO DE SOLICITUD",
                "REFERENCIA CLIENTE",
                "CODIGO DE COTIZACION",
                "NOMBRE DE PROYECTO",
                "VERSION",
                "ESTATUS",
                "MARCA",
                "CANTIDAD",
            };

            var listado = (from item in GetListadoFiltrado() //SolicitudClienteExternoEntity.ListadoSolicitudClienteExterno()
                           select new object[]
                           {
                                            item.fecha_hora_solicitud.HasValue ? item.fecha_hora_solicitud.Value.ToString("yyyy/MM/dd") : string.Empty,
                                            item.TiempoTranscurrido,
                                           item.CodigoSolicitud,
                                            item.NombresCompletosSolicitante,
                                            item.TextoCatalogoTipo,
                                            item.TextoCatalogoSubTipo,
                                            item.mkt,
                                            item.codigo_cotizacion,
                                            item.NombreProyectoCodigoCotizacion,
                                            item.VersionCodigoCotizacion,
                                            item.EstatusCodigo,
                                            item.TextoCatalogoMarca,
                                            item.cantidad,
                           }).ToList();

            // Build the file content
            var employeecsv = new StringBuilder();
            listado.ForEach(line =>
            {
                employeecsv.AppendLine(string.Join(",", line));
            });

            byte[] buffer = Encoding.Default.GetBytes($"{string.Join(",", comlumHeadrs)}\r\n{employeecsv.ToString()}");
            return File(buffer, "text/csv", $"ListadoSolicitudes.csv");
        }

        public ActionResult DescargarReporteFormatoPDF()
        {
            // Seleccionar las columnas a exportar
            var results = GetListadoFiltrado();//SolicitudClienteExternoEntity.ListadoSolicitudClienteExterno();

            var list = Reportes.SerializeToJSON(results);
            return Content(list, "application/json");
        }
        #endregion

    }
}
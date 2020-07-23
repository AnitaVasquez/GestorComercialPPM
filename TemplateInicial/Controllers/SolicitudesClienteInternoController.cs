using GestionPPM.Entidades.Metodos;
using GestionPPM.Entidades.Modelo;
using GestionPPM.Repositorios;
using iTextSharp.text;
using iTextSharp.text.pdf;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using NonFactors.Mvc.Grid;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using Seguridad.Helper;
using System;
using System.Collections.Generic;
using System.Configuration;
//using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using System.Web;
using System.Web.Mvc;

namespace TemplateInicial.Controllers
{
    [Autenticado]
    public class SolicitudesClienteInternoController : BaseAppController
    {
        GestionPPMEntities db = new GestionPPMEntities();

        // GET: SolicitudesClienteInterno
        public ActionResult Index()
        {
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
            var puedeAsignar = PuedeReasignarUsuarios(1, 1);
            ViewBag.PuedeReasigarUsuarios = puedeAsignar;//PuedeReasignarUsuarios(1, 1);

            ViewBag.NombreListado = Etiquetas.TituloGridSolicitudRequerimientoClienteInterno;

            var usuariosFiltro = OrganigramaEntity.GetEstructuraOrganigrama(1);

            //var filtros = JsonConvert.DeserializeObject<List<OrganigramaParcial>>(usuariosFiltro);

            var user = ViewData["usuario"] = System.Web.HttpContext.Current.Session["usuario"];

            var usuarios = OrganigramaEntity.GetHijosOrganigramaByUsuarioID(int.Parse(user.ToString()), 1);

            //filtros = filtros.Where(s => s.parent == user.ToString() || s.id == int.Parse(user.ToString()) ).ToList();

            //var usuarios = filtros.Select(s => s.id).ToList();

            //Controlar permisos
            var usuario = int.Parse(user.ToString());
            string nombreControlador = ControllerContext.RouteData.Values["controller"].ToString();
            ViewBag.NombreControlador = nombreControlador;


            ViewBag.AccionesUsuario = ManejoPermisosEntity.ListadoAccionesCatalogoUsuario(usuario, nombreControlador);

            //Obtener Acciones del controlador
            ViewBag.AccionesControlador = GetMetodosControlador(nombreControlador);

            //Asignaciones
            var solicitudesUsuarios = SolicitudClienteInternoEntity.ListarSolicitudesUsuariosByUsuarioID(int.Parse(user.ToString()));

            // El usuario principal puede ver todo
            var listado = new List<SolicitudClienteInternoInfo>();

            int usuarioPrincipal = int.Parse(ParametrosSistemaEntity.ConsultarParametros(2).valor);

            //Usuario 4 ve todas las solicitudes
            if (int.Parse(user.ToString()) == usuarioPrincipal || int.Parse(user.ToString()) == 4)
            {
                listado = SolicitudClienteInternoEntity.ListadoSolicitudClienteInterno();
            }
            else
            {
                listado = SolicitudClienteInternoEntity.ListadoSolicitudClienteInterno();

                //PRIMER FILTRO
                var idsSolicitudesAsignadas = SolicitudClienteInternoEntity.ListadoSolicitudClienteInterno().Where(s => solicitudesUsuarios.Contains(s.id_solicitud)).Select(s => s.id_solicitud).ToList();

                //SEGUNDO FILTRO
                var idsSolicitudesUsuariosDependientes = SolicitudClienteInternoEntity.ListadoSolicitudClienteInterno().Where(s => usuarios.Contains(s.id_solicitante.Value)).Select(s => s.id_solicitud).ToList();

                //TERCER FILTRO (USUARIOS QUE SE ENCUENTREN HASTA EL NIVEL DE PARAMETRIZACIÓN)
                //int nivelParametrizadoReasignaciones = int.Parse(ParametrosSistemaEntity.ConsultarParametros(4).valor);

                //var nivelUsuario = OrganigramaEntity.ConsultarNivelUsuario(1,1, int.Parse(user.ToString()));

                //var idsUsuariosDependientesNivelParametrizacion = OrganigramaEntity.ConsultarEstructuraOrganigramaUsuarioIDByRangoNivel(1, 1, int.Parse(user.ToString()), nivelUsuario, nivelParametrizadoReasignaciones).Select(s => s.id).ToList();
                //var idsSolicitudesUsuariosDependientesNivelParametrizacion = SolicitudClienteInternoEntity.ListadoSolicitudClienteInterno().Where(s => idsUsuariosDependientesNivelParametrizacion.Contains(s.id_solicitante.Value)).Select(s => s.id_solicitud).ToList();

                //listado = listado.Where(s => idsSolicitudesAsignadas.Contains(s.id_solicitud) || idsSolicitudesUsuariosDependientes.Contains(s.id_solicitud) || idsSolicitudesUsuariosDependientesNivelParametrizacion.Contains(s.id_solicitud)).ToList();

                var idsSolicitudesSinAsignar = SolicitudClienteInternoEntity.ConsultarSolicitudesSinAsignacion();

                var listado1 = listado.Where(s => idsSolicitudesSinAsignar.Contains(s.id_solicitud)).Select(s => s.id_solicitud).ToList();

                var listado2 = listado.Where(s => idsSolicitudesAsignadas.Contains(s.id_solicitud) || idsSolicitudesUsuariosDependientes.Contains(s.id_solicitud)).Select(s => s.id_solicitud).ToList();

                if (puedeAsignar)
                    listado = listado.Where(s => listado1.Contains(s.id_solicitud) || listado2.Contains(s.id_solicitud)).ToList();
                else
                    listado = listado.Where(s => listado2.Contains(s.id_solicitud)).ToList();

                //listado = listado.Where(s => idsSolicitudesAsignadas.Contains(s.id_solicitud) || idsSolicitudesUsuariosDependientes.Contains(s.id_solicitud)).ToList();
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

        private List<SolicitudClienteInternoInfo> GetListadoFiltrado()
        {
            List<SolicitudClienteInternoInfo> listado = new List<SolicitudClienteInternoInfo>();
            try
            {
                var puedeAsignar = PuedeReasignarUsuarios(1, 1);

                var user = ViewData["usuario"] = System.Web.HttpContext.Current.Session["usuario"];

                var usuarios = OrganigramaEntity.GetHijosOrganigramaByUsuarioID(int.Parse(user.ToString()), 1);

                //Asignaciones
                var solicitudesUsuarios = SolicitudClienteInternoEntity.ListarSolicitudesUsuariosByUsuarioID(int.Parse(user.ToString()));

                listado = SolicitudClienteInternoEntity.ListadoSolicitudClienteInterno();

                //PRIMER FILTRO
                var idsSolicitudesAsignadas = SolicitudClienteInternoEntity.ListadoSolicitudClienteInterno().Where(s => solicitudesUsuarios.Contains(s.id_solicitud)).Select(s => s.id_solicitud).ToList();

                //SEGUNDO FILTRO
                var idsSolicitudesUsuariosDependientes = SolicitudClienteInternoEntity.ListadoSolicitudClienteInterno().Where(s => usuarios.Contains(s.id_solicitante.Value)).Select(s => s.id_solicitud).ToList();

                var idsSolicitudesSinAsignar = SolicitudClienteInternoEntity.ConsultarSolicitudesSinAsignacion();

                var listado1 = listado.Where(s => idsSolicitudesSinAsignar.Contains(s.id_solicitud)).Select(s => s.id_solicitud).ToList();

                var listado2 = listado.Where(s => idsSolicitudesAsignadas.Contains(s.id_solicitud) || idsSolicitudesUsuariosDependientes.Contains(s.id_solicitud)).Select(s => s.id_solicitud).ToList();

                if (puedeAsignar)
                    listado = listado.Where(s => listado1.Contains(s.id_solicitud) || listado2.Contains(s.id_solicitud)).ToList();
                else
                    listado = listado.Where(s => listado2.Contains(s.id_solicitud)).ToList();

                return listado;
            }
            catch (Exception ex)
            {
                return listado;
            }
        }

        private static bool PuedeReasignarUsuarios(int idOrganigrama, int idEmpresa)
        {
            try
            {
                var usuarioID = Convert.ToInt32(System.Web.HttpContext.Current.Session["usuario"]);
                int nivelUsuario = 0;
                var usuario = OrganigramaEntity.ConsultarEstructuraOrganigrama(idOrganigrama, idEmpresa, usuarioID).FirstOrDefault();

                if (usuario != null)
                    nivelUsuario = usuario.nivel.Value;
                else
                    return false;

                int nivelParametrizadoReasignaciones = int.Parse(ParametrosSistemaEntity.ConsultarParametros(4).valor);

                if (nivelUsuario <= nivelParametrizadoReasignaciones)
                    return true;
                else
                    return false;
            }
            catch (Exception ex)
            {
                return false;
            }

        }

        public ActionResult _AdjuntarArchivos(int? id)
        {
            ViewBag.TituloModal = "Adjuntar archivos";

            ViewBag.AdjuntosVacio = "No existen archivos adjuntos.";

            SolicitudClienteInternoInfo solicitud = SolicitudClienteInternoEntity.ConsultarSolicitudClienteInterno(id.Value);
            return PartialView(solicitud);
        }

        public ActionResult _AsignacionUsuarios(int? id)
        {
            ViewBag.TituloModal = "Asignación Usuarios";
            //SolicitudClienteInternoInfo solicitud = SolicitudClienteInternoEntity.ConsultarSolicitudClienteInterno(id.Value);
            ViewBag.SolicitudID = id.Value;

            var user = ViewData["usuario"] = System.Web.HttpContext.Current.Session["usuario"];

            var solicitudesUsuarios = SolicitudClienteInternoEntity.ListarSolicitudesUsuariosBySolicitudID(id.Value);
            ViewBag.Modelo = string.Join(",", solicitudesUsuarios.Where(s => !solicitudesUsuarios.Contains(int.Parse(user.ToString()))));

            return PartialView();
        }

        public ActionResult _EditarSolicitud(int? id)
        {
            ViewBag.TituloModal = "Editar Solicitud";

            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }

            SolicitudCliente solicitud = SolicitudClienteEntity.ConsultarSolicitudCliente(id.Value);

            ViewBag.DatosUsuario = UsuarioEntity.ConsultarUsuario(solicitud.id_solicitante.Value);

            //Listado de Tipo 
            var Tipo = CatalogoEntity.ObtenerListadoCatalogosByCodigo("TSL-01");
            ViewBag.ListadoTipo = Tipo;

            //Validar los tipos
            var catalogo = CatalogoEntity.ConsultarCatalogo(solicitud.id_tipo);

            //Si es Eclub o Icare
            if (catalogo.nombre_catalgo == "ICARE" || catalogo.nombre_catalgo == "ECLUB")
            {
                var catalogos = new string[] { "ICARE", "ECLUB" };
                ViewBag.ListadoTipo = Tipo.Where(t => catalogos.Contains(t.Text)).ToList();
            }
            else
            {
                if (catalogo.nombre_catalgo == "HTML5" || catalogo.nombre_catalgo == "TAGS" || catalogo.nombre_catalgo == "HTML5 Y TAGS")
                {
                    var catalogosHtml = new string[] { "HTML5", "TAGS", "HTML5 Y TAGS" };
                    ViewBag.ListadoTipo = Tipo.Where(t => catalogosHtml.Contains(t.Text)).ToList();
                }
                else
                {
                    ViewBag.ListadoTipo = Tipo.Where(t => t.Text == catalogo.nombre_catalgo).ToList();
                }
            }

            //Listado Marca
            var Marca = CatalogoEntity.ObtenerListadoCatalogosByCodigo("MRC-01");
            ViewBag.ListadoMarca = Marca;

            //Listado Subtipo  
            var Subtipo = CatalogoEntity.ConsultarCatalogoPorPadre(solicitud.id_tipo.HasValue ? solicitud.id_tipo.Value : 0, "SUBTIPO");
            ViewBag.ListadoSubtipo = Subtipo;

            //Url de soporte
            List<UrlExternosSolicitud> ListadoUrlExterno = db.UrlExternosSolicitud(id).ToList();
            ViewBag.UrlExternosSolicitud = ListadoUrlExterno;

            if (solicitud == null)
            {
                return HttpNotFound();
            }
            return PartialView(solicitud);
        }

        [HttpPost]
        public ActionResult EditarSolicitud(SolicitudCliente solicitud, List<UrlExternoParcial> listadoUrlSoporte)
        {
            try
            {
                RespuestaTransaccion resultado = SolicitudClienteEntity.ActualizarSolicitud(solicitud, listadoUrlSoporte);

                //Listado de Tipo 
                var Tipo = CatalogoEntity.ObtenerListadoCatalogosByCodigo("TSL-01");
                ViewBag.ListadoTipo = Tipo;

                //Validacion de adjuntos
                var codigoGeneral = Tipo.Where(t => Convert.ToInt32(t.Value) == 758).FirstOrDefault().Value;
                int valorCodigoGeneral = Convert.ToInt16(codigoGeneral);

                //Validar si es icare o eclub y validar el tamaño de las imagenes y el formato de los archivos
                var codigoIcare = Tipo.Where(t => Convert.ToInt32(t.Value) == 735).FirstOrDefault().Value;
                int valorCodigoIcare = Convert.ToInt16(codigoIcare);

                //Validar si es icare o eclub y validar el tamaño de las imagenes y el formato de los archivos
                var codigoEclub = Tipo.Where(t => Convert.ToInt32(t.Value) == 734).FirstOrDefault().Value;
                int valorCodigoEclub = Convert.ToInt16(codigoEclub);

                //Validar si es icare o eclub y validar el tamaño de las imagenes y el formato de los archivos
                var codigoMantenimiento = Tipo.Where(t => Convert.ToInt32(t.Value) == 737).FirstOrDefault().Value;
                int valorCodigoMantenimiento = Convert.ToInt16(codigoMantenimiento);

                //Validar envios icare  
                var codigoEnviosIcare = Tipo.Where(t => Convert.ToInt32(t.Value) == 736).FirstOrDefault().Value;
                int valorcodigoEnviosIcare = Convert.ToInt16(codigoEnviosIcare);

                //Validar Html 
                var codigoHtml = Tipo.Where(t => Convert.ToInt32(t.Value) == 739).FirstOrDefault().Value;
                int valorcodigoHtml = Convert.ToInt16(codigoHtml);

                //Validar Html y Tags
                var codigoHtmlTag = Tipo.Where(t => Convert.ToInt32(t.Value) == 740).FirstOrDefault().Value;
                int valorcodigoHtmlTags = Convert.ToInt16(codigoHtmlTag);

                //Validar Tags
                var codigoTag = Tipo.Where(t => Convert.ToInt32(t.Value) == 738).FirstOrDefault().Value;
                int valorcodigoTags = Convert.ToInt16(codigoTag);

                //Si es una solicitud de tipo General
                if (solicitud.id_tipo == valorCodigoGeneral)
                {
                    GenerarPDFGeneral(resultado.SolicitudID);
                    return Json(new { Resultado = resultado }, JsonRequestBehavior.AllowGet);
                }
                else
                {
                    //Si es una solicitud de tipo Icare o Eclub
                    if (solicitud.id_tipo == valorCodigoIcare || solicitud.id_tipo == valorCodigoEclub)
                    {
                        GenerarPDFEclubIcare(solicitud.id_solicitud);
                        return Json(new { Resultado = resultado }, JsonRequestBehavior.AllowGet);
                    }
                    else
                    {
                        if (solicitud.id_tipo == valorCodigoMantenimiento)
                        {
                            GenerarPDFMantenimiento(solicitud.id_solicitud);
                            return Json(new { Resultado = resultado }, JsonRequestBehavior.AllowGet);
                        }
                        else
                        {
                            //Si es una solicitud de envios icare
                            if (solicitud.id_tipo == valorcodigoEnviosIcare)
                            {
                                GenerarPDFEnvioIcare(solicitud.id_solicitud);
                                return Json(new { Resultado = resultado }, JsonRequestBehavior.AllowGet);
                            }
                            else
                            {
                                //Si es de html5
                                if (solicitud.id_tipo == valorcodigoHtml)
                                {
                                    GenerarPDFHtml5(solicitud.id_solicitud);
                                    return Json(new { Resultado = resultado }, JsonRequestBehavior.AllowGet);
                                }
                                else
                                {
                                    //Si es de tags
                                    if (solicitud.id_tipo == valorcodigoTags)
                                    {
                                        GenerarPDFTags(solicitud.id_solicitud);
                                        return Json(new { Resultado = resultado }, JsonRequestBehavior.AllowGet);
                                    }
                                    else
                                    {
                                        //Si es de html5 y  tags
                                        if (solicitud.id_tipo == valorcodigoHtmlTags)
                                        {
                                            GenerarPDFHtmlTags(solicitud.id_solicitud);
                                            return Json(new { Resultado = resultado }, JsonRequestBehavior.AllowGet);
                                        }
                                        else
                                        {
                                            return Json(new { Resultado = resultado }, JsonRequestBehavior.AllowGet);
                                        }
                                    }

                                }
                            }

                        }
                    }
                }
            }
            catch (Exception ex)
            {
                return Json(new { Resultado = new RespuestaTransaccion { Estado = false, Respuesta = ex.Message } }, JsonRequestBehavior.AllowGet);
            }
        }

        #region Generacion PDF

        //Reporte Solicitud General
        public void GenerarPDFGeneral(int? id)
        {
            //Obtener los datos de la Solicitud
            DatosSolicitudCliente solicitud = db.DatosSolicitudCliente(id).First();

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

            //anio Actual
            var anioActual = System.DateTime.Now.Year.ToString();

            string basePath = ConfigurationManager.AppSettings["RepositorioDocumentos"];
            string rutaArchivos = basePath + "\\GESTION_PPM\\ADJUNTOS_SOLICITUDES\\" + anioActual + '\\' + solicitud.Id.ToString();

            // En caso de que no exista el directorio, crearlo.
            bool directorio = Directory.Exists(rutaArchivos);
            if (!directorio)
                Directory.CreateDirectory(rutaArchivos);

            //ruta del archivo PDF que se va a crear
            string rutaDocumentos = rutaArchivos + "\\Solicitud-" + solicitud.CodigoSolicitud.ToString() + ".pdf";

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
            PdfPCell TituloCotizador = new PdfPCell(new Phrase("SOLICITUD " + solicitud.Tipo.ToUpper() + " N° " + solicitud.CodigoSolicitud + "      ", new Font(customfontbold, 9, 0, new BaseColor(255, 255, 255))));
            TituloCotizador.Colspan = 5;
            TituloCotizador.Border = Rectangle.NO_BORDER;
            TituloCotizador.HorizontalAlignment = Element.ALIGN_RIGHT;
            TituloCotizador.VerticalAlignment = Element.ALIGN_MIDDLE;
            TituloCotizador.BackgroundColor = new BaseColor(60, 66, 82);
            TituloCotizador.FixedHeight = 50f;

            cabecera.AddCell(logocell);
            cabecera.AddCell(TituloCotizador);

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

            //Informacion de la solicitud         
            PdfPTable informacion = new PdfPTable(6);

            PdfPCell EtiquetaFechaCotizacion = new PdfPCell(new Phrase("Fecha y Hora: ", new Font(customfontbold, 7)));
            EtiquetaFechaCotizacion.Colspan = 1;
            EtiquetaFechaCotizacion.HorizontalAlignment = Element.ALIGN_RIGHT;
            EtiquetaFechaCotizacion.VerticalAlignment = Element.ALIGN_MIDDLE;
            EtiquetaFechaCotizacion.FixedHeight = 16f;
            EtiquetaFechaCotizacion.BackgroundColor = new BaseColor(225, 225, 225);
            EtiquetaFechaCotizacion.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell ValorFechaCotizacion = new PdfPCell(new Phrase("   " + solicitud.FechaHora, new Font(customfont, 7)));
            ValorFechaCotizacion.Colspan = 2;
            ValorFechaCotizacion.HorizontalAlignment = Element.ALIGN_LEFT;
            ValorFechaCotizacion.VerticalAlignment = Element.ALIGN_MIDDLE;
            ValorFechaCotizacion.FixedHeight = 16f;
            ValorFechaCotizacion.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell EtiquetaSolicitante = new PdfPCell(new Phrase("Solicitante:  ", new Font(customfontbold, 7)));
            EtiquetaSolicitante.Colspan = 1;
            EtiquetaSolicitante.HorizontalAlignment = Element.ALIGN_RIGHT;
            EtiquetaSolicitante.VerticalAlignment = Element.ALIGN_MIDDLE;
            EtiquetaSolicitante.FixedHeight = 16f;
            EtiquetaSolicitante.BackgroundColor = new BaseColor(225, 225, 225);
            EtiquetaSolicitante.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell ValorSolicitante = new PdfPCell(new Phrase("   " + solicitud.Solicitante, new Font(customfont, 7)));
            ValorSolicitante.Colspan = 2;
            ValorSolicitante.HorizontalAlignment = Element.ALIGN_LEFT;
            ValorSolicitante.VerticalAlignment = Element.ALIGN_MIDDLE;
            ValorSolicitante.FixedHeight = 16f;
            ValorSolicitante.BorderColor = new BaseColor(60, 66, 82);

            informacion.AddCell(EtiquetaFechaCotizacion);
            informacion.AddCell(ValorFechaCotizacion);
            informacion.AddCell(EtiquetaSolicitante);
            informacion.AddCell(ValorSolicitante);

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

            //Informacion de la Solicitud       
            PdfPTable datosGenerales = new PdfPTable(6);

            PdfPCell EtiquetaEmpresa = new PdfPCell(new Phrase("Empresa: ", new Font(customfontbold, 7)));
            EtiquetaEmpresa.Colspan = 1;
            EtiquetaEmpresa.HorizontalAlignment = Element.ALIGN_RIGHT;
            EtiquetaEmpresa.VerticalAlignment = Element.ALIGN_MIDDLE;
            EtiquetaEmpresa.FixedHeight = 16f;
            EtiquetaEmpresa.BackgroundColor = new BaseColor(225, 225, 225);
            EtiquetaEmpresa.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell ValorEmpresa = new PdfPCell(new Phrase("   " + solicitud.Empresa, new Font(customfont, 7)));
            ValorEmpresa.Colspan = 5;
            ValorEmpresa.HorizontalAlignment = Element.ALIGN_LEFT;
            ValorEmpresa.VerticalAlignment = Element.ALIGN_MIDDLE;
            ValorEmpresa.FixedHeight = 16f;
            ValorEmpresa.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell EtiquetaCargo = new PdfPCell(new Phrase("Cargo: ", new Font(customfontbold, 7)));
            EtiquetaCargo.Colspan = 1;
            EtiquetaCargo.HorizontalAlignment = Element.ALIGN_RIGHT;
            EtiquetaCargo.VerticalAlignment = Element.ALIGN_MIDDLE;
            EtiquetaCargo.FixedHeight = 16f;
            EtiquetaCargo.BackgroundColor = new BaseColor(225, 225, 225);
            EtiquetaCargo.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell ValorCargo = new PdfPCell(new Phrase("   " + solicitud.Cargo, new Font(customfont, 7)));
            ValorCargo.Colspan = 5;
            ValorCargo.HorizontalAlignment = Element.ALIGN_LEFT;
            ValorCargo.VerticalAlignment = Element.ALIGN_MIDDLE;
            ValorCargo.FixedHeight = 16f;
            ValorCargo.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell EtiquetaTelefono = new PdfPCell(new Phrase("Teléfono: ", new Font(customfontbold, 7)));
            EtiquetaTelefono.Colspan = 1;
            EtiquetaTelefono.HorizontalAlignment = Element.ALIGN_RIGHT;
            EtiquetaTelefono.VerticalAlignment = Element.ALIGN_MIDDLE;
            EtiquetaTelefono.FixedHeight = 16f;
            EtiquetaTelefono.BackgroundColor = new BaseColor(225, 225, 225);
            EtiquetaTelefono.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell ValorTelefono = new PdfPCell(new Phrase("   " + solicitud.Telefono, new Font(customfont, 7)));
            ValorTelefono.Colspan = 2;
            ValorTelefono.HorizontalAlignment = Element.ALIGN_LEFT;
            ValorTelefono.VerticalAlignment = Element.ALIGN_MIDDLE;
            ValorTelefono.FixedHeight = 16f;
            ValorTelefono.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell EtiquetaCelular = new PdfPCell(new Phrase("Celular: ", new Font(customfontbold, 7)));
            EtiquetaCelular.Colspan = 1;
            EtiquetaCelular.HorizontalAlignment = Element.ALIGN_RIGHT;
            EtiquetaCelular.VerticalAlignment = Element.ALIGN_MIDDLE;
            EtiquetaCelular.FixedHeight = 16f;
            EtiquetaCelular.BackgroundColor = new BaseColor(225, 225, 225);
            EtiquetaCelular.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell ValorCelular = new PdfPCell(new Phrase("   " + solicitud.Celular, new Font(customfont, 7)));
            ValorCelular.Colspan = 2;
            ValorCelular.HorizontalAlignment = Element.ALIGN_LEFT;
            ValorCelular.VerticalAlignment = Element.ALIGN_MIDDLE;
            ValorCelular.FixedHeight = 16f;
            ValorCelular.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell EtiquetaMail = new PdfPCell(new Phrase("Mail: ", new Font(customfontbold, 7)));
            EtiquetaMail.Colspan = 1;
            EtiquetaMail.HorizontalAlignment = Element.ALIGN_RIGHT;
            EtiquetaMail.VerticalAlignment = Element.ALIGN_MIDDLE;
            EtiquetaMail.FixedHeight = 16f;
            EtiquetaMail.BackgroundColor = new BaseColor(225, 225, 225);
            EtiquetaMail.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell ValorMail = new PdfPCell(new Phrase("   " + solicitud.Mail, new Font(customfont, 7)));
            ValorMail.Colspan = 5;
            ValorMail.HorizontalAlignment = Element.ALIGN_LEFT;
            ValorMail.VerticalAlignment = Element.ALIGN_MIDDLE;
            ValorMail.FixedHeight = 16f;
            ValorMail.BorderColor = new BaseColor(60, 66, 82);

            datosGenerales.AddCell(EtiquetaEmpresa);
            datosGenerales.AddCell(ValorEmpresa);

            datosGenerales.AddCell(EtiquetaCargo);
            datosGenerales.AddCell(ValorCargo);

            datosGenerales.AddCell(EtiquetaTelefono);
            datosGenerales.AddCell(ValorTelefono);
            datosGenerales.AddCell(EtiquetaCelular);
            datosGenerales.AddCell(ValorCelular);

            datosGenerales.AddCell(EtiquetaMail);
            datosGenerales.AddCell(ValorMail);

            document.Add(datosGenerales);

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

            //Descripcion del requermiento    
            PdfPTable entregables = new PdfPTable(6);

            PdfPCell EtiquetaDescripcion = new PdfPCell(new Phrase("DESCRIPCIÓN DEL REQUERIMIENTO", new Font(customfontbold, 9, 0, new BaseColor(255, 255, 255))));
            EtiquetaDescripcion.Colspan = 6;
            EtiquetaDescripcion.HorizontalAlignment = Element.ALIGN_CENTER;
            EtiquetaDescripcion.VerticalAlignment = Element.ALIGN_MIDDLE;
            EtiquetaDescripcion.FixedHeight = 16f;
            EtiquetaDescripcion.BackgroundColor = new BaseColor(60, 66, 82);
            EtiquetaDescripcion.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell ValorDescripcion = new PdfPCell(new Phrase(solicitud.Descripcion, new Font(customfont, 7)));
            ValorDescripcion.Colspan = 6;
            ValorDescripcion.Rowspan = 6;
            ValorDescripcion.VerticalAlignment = Element.ALIGN_MIDDLE;
            ValorDescripcion.BorderColor = new BaseColor(60, 66, 82);

            //validar el tamaño de las celdas
            var caracteres = solicitud.Descripcion.Count();
            if (caracteres <= 70)
            {
                ValorDescripcion.FixedHeight = 16f;
            }
            else if (caracteres <= 140)
            {
                ValorDescripcion.FixedHeight = 32f;
            }
            else if (caracteres <= 210)
            {
                ValorDescripcion.FixedHeight = 46f;
            }
            else if (caracteres <= 280)
            {
                ValorDescripcion.FixedHeight = 58f;
            }
            else if (caracteres <= 350)
            {
                ValorDescripcion.FixedHeight = 70f;
            }
            else if (caracteres <= 420)
            {
                ValorDescripcion.FixedHeight = 82f;
            }
            else if (caracteres <= 550)
            {
                ValorDescripcion.FixedHeight = 94f;
            }
            else if (caracteres <= 680)
            {
                ValorDescripcion.FixedHeight = 106f;
            }
            else if (caracteres <= 810)
            {
                ValorDescripcion.FixedHeight = 128f;
            }
            else
            {
                ValorDescripcion.FixedHeight = 140f;
            }

            entregables.AddCell(EtiquetaDescripcion);
            entregables.AddCell(ValorDescripcion);

            document.Add(entregables);
            document.Close();
        }

        //Reporte Solicitud Eclub o Icare
        public void GenerarPDFEclubIcare(int? id)
        {
            int camposPersonalizados = 1;
            int contadorAdjuntos = 1;

            //Obtener los datos de la Solicitud
            DatosSolicitudCliente solicitud = db.DatosSolicitudCliente(id).First();

            //Campos Personalizados
            List<CamposPersonalizadosSolicitud> ListadoCamposPersonalizados = db.CamposPersonalizadosSolicitud(id).ToList();

            //Url Externo 
            List<UrlExternosSolicitud> ListadoUrlExterno = db.UrlExternosSolicitud(id).ToList();

            //Adjuntos
            List<AdjuntosSolicitud> ListadoAdjuntos = db.AdjuntosSolicitud(id).ToList();

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

            //anio Actual
            var anioActual = System.DateTime.Now.Year.ToString();

            string basePath = ConfigurationManager.AppSettings["RepositorioDocumentos"];
            string rutaArchivos = basePath + "\\GESTION_PPM\\ADJUNTOS_SOLICITUDES\\" + anioActual + '\\' + solicitud.Id.ToString();

            // En caso de que no exista el directorio, crearlo.
            bool directorio = Directory.Exists(rutaArchivos);
            if (!directorio)
                Directory.CreateDirectory(rutaArchivos);

            //ruta del archivo PDF que se va a crear
            string rutaDocumentos = rutaArchivos + "\\Solicitud-" + solicitud.CodigoSolicitud.ToString() + ".pdf";

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
            PdfPCell TituloCotizador = new PdfPCell(new Phrase("SOLICITUD " + solicitud.Tipo.ToUpper() + " N° " + solicitud.CodigoSolicitud + "      ", new Font(customfontbold, 9, 0, new BaseColor(255, 255, 255))));
            TituloCotizador.Colspan = 5;
            TituloCotizador.Border = Rectangle.NO_BORDER;
            TituloCotizador.HorizontalAlignment = Element.ALIGN_RIGHT;
            TituloCotizador.VerticalAlignment = Element.ALIGN_MIDDLE;
            TituloCotizador.BackgroundColor = new BaseColor(60, 66, 82);
            TituloCotizador.FixedHeight = 50f;

            cabecera.AddCell(logocell);
            cabecera.AddCell(TituloCotizador);

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

            //Informacion de la solicitud         
            PdfPTable informacion = new PdfPTable(6);

            PdfPCell EtiquetaFechaCotizacion = new PdfPCell(new Phrase("Fecha y Hora: ", new Font(customfontbold, 7)));
            EtiquetaFechaCotizacion.Colspan = 1;
            EtiquetaFechaCotizacion.HorizontalAlignment = Element.ALIGN_RIGHT;
            EtiquetaFechaCotizacion.VerticalAlignment = Element.ALIGN_MIDDLE;
            EtiquetaFechaCotizacion.FixedHeight = 16f;
            EtiquetaFechaCotizacion.BackgroundColor = new BaseColor(225, 225, 225);
            EtiquetaFechaCotizacion.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell ValorFechaCotizacion = new PdfPCell(new Phrase("   " + solicitud.FechaHora, new Font(customfont, 7)));
            ValorFechaCotizacion.Colspan = 2;
            ValorFechaCotizacion.HorizontalAlignment = Element.ALIGN_LEFT;
            ValorFechaCotizacion.VerticalAlignment = Element.ALIGN_MIDDLE;
            ValorFechaCotizacion.FixedHeight = 16f;
            ValorFechaCotizacion.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell EtiquetaSolicitante = new PdfPCell(new Phrase("Solicitante:  ", new Font(customfontbold, 7)));
            EtiquetaSolicitante.Colspan = 1;
            EtiquetaSolicitante.HorizontalAlignment = Element.ALIGN_RIGHT;
            EtiquetaSolicitante.VerticalAlignment = Element.ALIGN_MIDDLE;
            EtiquetaSolicitante.FixedHeight = 16f;
            EtiquetaSolicitante.BackgroundColor = new BaseColor(225, 225, 225);
            EtiquetaSolicitante.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell ValorSolicitante = new PdfPCell(new Phrase("   " + solicitud.Solicitante, new Font(customfont, 7)));
            ValorSolicitante.Colspan = 2;
            ValorSolicitante.HorizontalAlignment = Element.ALIGN_LEFT;
            ValorSolicitante.VerticalAlignment = Element.ALIGN_MIDDLE;
            ValorSolicitante.FixedHeight = 16f;
            ValorSolicitante.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell EtiquetaArea1 = new PdfPCell(new Phrase("Área 1", new Font(customfontbold, 7)));
            EtiquetaArea1.Colspan = 1;
            EtiquetaArea1.HorizontalAlignment = Element.ALIGN_RIGHT;
            EtiquetaArea1.VerticalAlignment = Element.ALIGN_MIDDLE;
            EtiquetaArea1.FixedHeight = 16f;
            EtiquetaArea1.BackgroundColor = new BaseColor(225, 225, 225);
            EtiquetaArea1.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell ValorArea1 = new PdfPCell(new Phrase("   " + solicitud.Area1, new Font(customfont, 7)));
            ValorArea1.Colspan = 2;
            ValorArea1.HorizontalAlignment = Element.ALIGN_LEFT;
            ValorArea1.VerticalAlignment = Element.ALIGN_MIDDLE;
            ValorArea1.FixedHeight = 16f;
            ValorArea1.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell EtiquetaArea2 = new PdfPCell(new Phrase("Área 2:  ", new Font(customfontbold, 7)));
            EtiquetaArea2.Colspan = 1;
            EtiquetaArea2.HorizontalAlignment = Element.ALIGN_RIGHT;
            EtiquetaArea2.VerticalAlignment = Element.ALIGN_MIDDLE;
            EtiquetaArea2.FixedHeight = 16f;
            EtiquetaArea2.BackgroundColor = new BaseColor(225, 225, 225);
            EtiquetaArea2.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell ValorArea2 = new PdfPCell(new Phrase("   " + solicitud.Area2, new Font(customfont, 7)));
            ValorArea2.Colspan = 2;
            ValorArea2.HorizontalAlignment = Element.ALIGN_LEFT;
            ValorArea2.VerticalAlignment = Element.ALIGN_MIDDLE;
            ValorArea2.FixedHeight = 16f;
            ValorArea2.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell EtiquetaImplementacion = new PdfPCell(new Phrase("Implementación:  ", new Font(customfontbold, 7)));
            EtiquetaImplementacion.Colspan = 1;
            EtiquetaImplementacion.HorizontalAlignment = Element.ALIGN_RIGHT;
            EtiquetaImplementacion.VerticalAlignment = Element.ALIGN_MIDDLE;
            EtiquetaImplementacion.FixedHeight = 16f;
            EtiquetaImplementacion.BackgroundColor = new BaseColor(225, 225, 225);
            EtiquetaImplementacion.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell ValorImplementacion = new PdfPCell(new Phrase("   " + solicitud.Implementacion, new Font(customfont, 7)));
            ValorImplementacion.Colspan = 5;
            ValorImplementacion.HorizontalAlignment = Element.ALIGN_LEFT;
            ValorImplementacion.VerticalAlignment = Element.ALIGN_MIDDLE;
            ValorImplementacion.FixedHeight = 16f;
            ValorImplementacion.BorderColor = new BaseColor(60, 66, 82);

            informacion.AddCell(EtiquetaFechaCotizacion);
            informacion.AddCell(ValorFechaCotizacion);
            informacion.AddCell(EtiquetaSolicitante);
            informacion.AddCell(ValorSolicitante);

            informacion.AddCell(EtiquetaArea1);
            informacion.AddCell(ValorArea1);
            informacion.AddCell(EtiquetaArea2);
            informacion.AddCell(ValorArea2);

            informacion.AddCell(EtiquetaImplementacion);
            informacion.AddCell(ValorImplementacion);

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

            //Informacion de la Solicitud       
            PdfPTable datosGenerales = new PdfPTable(6);

            PdfPCell EtiquetaTipo = new PdfPCell(new Phrase("Tipo: ", new Font(customfontbold, 7)));
            EtiquetaTipo.Colspan = 1;
            EtiquetaTipo.HorizontalAlignment = Element.ALIGN_RIGHT;
            EtiquetaTipo.VerticalAlignment = Element.ALIGN_MIDDLE;
            EtiquetaTipo.FixedHeight = 16f;
            EtiquetaTipo.BackgroundColor = new BaseColor(225, 225, 225);
            EtiquetaTipo.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell ValorTipo = new PdfPCell(new Phrase("   " + solicitud.Tipo, new Font(customfont, 7)));
            ValorTipo.Colspan = 5;
            ValorTipo.HorizontalAlignment = Element.ALIGN_LEFT;
            ValorTipo.VerticalAlignment = Element.ALIGN_MIDDLE;
            ValorTipo.FixedHeight = 16f;
            ValorTipo.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell EtiquetaSubtipo = new PdfPCell(new Phrase("Subtipo: ", new Font(customfontbold, 7)));
            EtiquetaSubtipo.Colspan = 1;
            EtiquetaSubtipo.HorizontalAlignment = Element.ALIGN_RIGHT;
            EtiquetaSubtipo.VerticalAlignment = Element.ALIGN_MIDDLE;
            EtiquetaSubtipo.FixedHeight = 16f;
            EtiquetaSubtipo.BackgroundColor = new BaseColor(225, 225, 225);
            EtiquetaSubtipo.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell ValorSubtipo = new PdfPCell(new Phrase("   " + solicitud.Subtipo, new Font(customfont, 7)));
            ValorSubtipo.Colspan = 5;
            ValorSubtipo.HorizontalAlignment = Element.ALIGN_LEFT;
            ValorSubtipo.VerticalAlignment = Element.ALIGN_MIDDLE;
            ValorSubtipo.FixedHeight = 16f;
            ValorSubtipo.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell EtiquetaMarca = new PdfPCell(new Phrase("Marca: ", new Font(customfontbold, 7)));
            EtiquetaMarca.Colspan = 1;
            EtiquetaMarca.HorizontalAlignment = Element.ALIGN_RIGHT;
            EtiquetaMarca.VerticalAlignment = Element.ALIGN_MIDDLE;
            EtiquetaMarca.FixedHeight = 16f;
            EtiquetaMarca.BackgroundColor = new BaseColor(225, 225, 225);
            EtiquetaMarca.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell ValorMarca = new PdfPCell(new Phrase("   " + solicitud.Marca, new Font(customfont, 7)));
            ValorMarca.Colspan = 5;
            ValorMarca.HorizontalAlignment = Element.ALIGN_LEFT;
            ValorMarca.VerticalAlignment = Element.ALIGN_MIDDLE;
            ValorMarca.FixedHeight = 16f;
            ValorMarca.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell EtiquetaOP = new PdfPCell(new Phrase("OP: ", new Font(customfontbold, 7)));
            EtiquetaOP.Colspan = 1;
            EtiquetaOP.HorizontalAlignment = Element.ALIGN_RIGHT;
            EtiquetaOP.VerticalAlignment = Element.ALIGN_MIDDLE;
            EtiquetaOP.FixedHeight = 16f;
            EtiquetaOP.BackgroundColor = new BaseColor(225, 225, 225);
            EtiquetaOP.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell ValorOP = new PdfPCell(new Phrase("   " + solicitud.OP, new Font(customfont, 7)));
            ValorOP.Colspan = 2;
            ValorOP.HorizontalAlignment = Element.ALIGN_LEFT;
            ValorOP.VerticalAlignment = Element.ALIGN_MIDDLE;
            ValorOP.FixedHeight = 16f;
            ValorOP.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell EtiquetaMKT = new PdfPCell(new Phrase("MKT: ", new Font(customfontbold, 7)));
            EtiquetaMKT.Colspan = 1;
            EtiquetaMKT.HorizontalAlignment = Element.ALIGN_RIGHT;
            EtiquetaMKT.VerticalAlignment = Element.ALIGN_MIDDLE;
            EtiquetaMKT.FixedHeight = 16f;
            EtiquetaMKT.BackgroundColor = new BaseColor(225, 225, 225);
            EtiquetaMKT.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell ValorMKT = new PdfPCell(new Phrase("   " + solicitud.MKT, new Font(customfont, 7)));
            ValorMKT.Colspan = 2;
            ValorMKT.HorizontalAlignment = Element.ALIGN_LEFT;
            ValorMKT.VerticalAlignment = Element.ALIGN_MIDDLE;
            ValorMKT.FixedHeight = 16f;
            ValorMKT.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell EtiquetaNombre = new PdfPCell(new Phrase("Nombre: ", new Font(customfontbold, 7)));
            EtiquetaNombre.Colspan = 1;
            EtiquetaNombre.HorizontalAlignment = Element.ALIGN_RIGHT;
            EtiquetaNombre.VerticalAlignment = Element.ALIGN_MIDDLE;
            EtiquetaNombre.FixedHeight = 16f;
            EtiquetaNombre.BackgroundColor = new BaseColor(225, 225, 225);
            EtiquetaNombre.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell ValorNombre = new PdfPCell(new Phrase("   " + solicitud.Nombre, new Font(customfont, 7)));
            ValorNombre.Colspan = 5;
            ValorNombre.HorizontalAlignment = Element.ALIGN_LEFT;
            ValorNombre.VerticalAlignment = Element.ALIGN_MIDDLE;
            ValorNombre.FixedHeight = 16f;
            ValorNombre.BorderColor = new BaseColor(60, 66, 82);

            datosGenerales.AddCell(EtiquetaTipo);
            datosGenerales.AddCell(ValorTipo);

            datosGenerales.AddCell(EtiquetaSubtipo);
            datosGenerales.AddCell(ValorSubtipo);

            datosGenerales.AddCell(EtiquetaMarca);
            datosGenerales.AddCell(ValorMarca);

            datosGenerales.AddCell(EtiquetaOP);
            datosGenerales.AddCell(ValorOP);
            datosGenerales.AddCell(EtiquetaMKT);
            datosGenerales.AddCell(ValorMKT);

            datosGenerales.AddCell(EtiquetaNombre);
            datosGenerales.AddCell(ValorNombre);

            document.Add(datosGenerales);

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

            //Campos Personalizados
            PdfPTable camposDetallePersonalizados = new PdfPTable(6);

            PdfPCell EtiquetaCamposPersonalizados = new PdfPCell(new Phrase("CAMPOS PERSONALIZABLES", new Font(customfontbold, 9, 0, new BaseColor(255, 255, 255))));
            EtiquetaCamposPersonalizados.Colspan = 6;
            EtiquetaCamposPersonalizados.HorizontalAlignment = Element.ALIGN_CENTER;
            EtiquetaCamposPersonalizados.VerticalAlignment = Element.ALIGN_MIDDLE;
            EtiquetaCamposPersonalizados.FixedHeight = 16f;
            EtiquetaCamposPersonalizados.BackgroundColor = new BaseColor(60, 66, 82);
            EtiquetaCamposPersonalizados.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell EtiquetaCampo = new PdfPCell(new Phrase("Campo", new Font(customfontbold, 7)));
            EtiquetaCampo.Colspan = 1;
            EtiquetaCampo.HorizontalAlignment = Element.ALIGN_CENTER;
            EtiquetaCampo.VerticalAlignment = Element.ALIGN_MIDDLE;
            EtiquetaCampo.FixedHeight = 16f;
            EtiquetaCampo.BackgroundColor = new BaseColor(225, 225, 225);
            EtiquetaCampo.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell EtiquetaNombreCampo = new PdfPCell(new Phrase("Nombre", new Font(customfontbold, 7)));
            EtiquetaNombreCampo.Colspan = 5;
            EtiquetaNombreCampo.HorizontalAlignment = Element.ALIGN_LEFT;
            EtiquetaNombreCampo.VerticalAlignment = Element.ALIGN_MIDDLE;
            EtiquetaNombreCampo.FixedHeight = 16f;
            EtiquetaNombreCampo.BackgroundColor = new BaseColor(225, 225, 225);
            EtiquetaNombreCampo.BorderColor = new BaseColor(60, 66, 82);

            camposDetallePersonalizados.AddCell(EtiquetaCamposPersonalizados);

            camposDetallePersonalizados.AddCell(EtiquetaCampo);
            camposDetallePersonalizados.AddCell(EtiquetaNombreCampo);

            foreach (CamposPersonalizadosSolicitud campo in ListadoCamposPersonalizados)
            {
                PdfPCell EtiquetaCampoDetalle = new PdfPCell(new Phrase(camposPersonalizados.ToString(), new Font(customfont, 7)));
                EtiquetaCampoDetalle.Colspan = 1;
                EtiquetaCampoDetalle.HorizontalAlignment = Element.ALIGN_CENTER;
                EtiquetaCampoDetalle.VerticalAlignment = Element.ALIGN_MIDDLE;
                EtiquetaCampoDetalle.FixedHeight = 23f;
                EtiquetaCampoDetalle.BorderColor = new BaseColor(60, 66, 82);

                PdfPCell ValorDetalle = new PdfPCell(new Phrase(campo.nombre_campo, new Font(customfont, 7)));
                ValorDetalle.Colspan = 5;
                ValorDetalle.HorizontalAlignment = Element.ALIGN_LEFT;
                ValorDetalle.VerticalAlignment = Element.ALIGN_MIDDLE;
                ValorDetalle.FixedHeight = 23f;
                ValorDetalle.BorderColor = new BaseColor(60, 66, 82);

                camposDetallePersonalizados.AddCell(EtiquetaCampoDetalle);
                camposDetallePersonalizados.AddCell(ValorDetalle);

                camposPersonalizados = camposPersonalizados + 1;
            }

            if (ListadoCamposPersonalizados.Count > 0)
            {
                document.Add(camposDetallePersonalizados);

                //Salto Linea       
                PdfPTable saltoLinea4 = new PdfPTable(4);

                PdfPCell EtiquetaSaltoLinea4 = new PdfPCell(new Phrase(" ", new Font(customfontbold, 9)));
                EtiquetaSaltoLinea4.Colspan = 4;
                EtiquetaSaltoLinea4.Border = Rectangle.NO_BORDER;
                EtiquetaSaltoLinea4.HorizontalAlignment = Element.ALIGN_CENTER;
                EtiquetaSaltoLinea4.VerticalAlignment = Element.ALIGN_MIDDLE;
                EtiquetaSaltoLinea4.FixedHeight = 8f;

                saltoLinea4.AddCell(EtiquetaSaltoLinea4);

                document.Add(saltoLinea4);
            }

            //Url Externo  
            PdfPTable urlsExterno = new PdfPTable(7);

            PdfPCell EtiquetaCampoUrls = new PdfPCell(new Phrase("URL's Externos", new Font(customfontbold, 9, 0, new BaseColor(255, 255, 255))));
            EtiquetaCampoUrls.Colspan = 7;
            EtiquetaCampoUrls.HorizontalAlignment = Element.ALIGN_CENTER;
            EtiquetaCampoUrls.VerticalAlignment = Element.ALIGN_MIDDLE;
            EtiquetaCampoUrls.FixedHeight = 16f;
            EtiquetaCampoUrls.BackgroundColor = new BaseColor(60, 66, 82);
            EtiquetaCampoUrls.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell EtiquetaTipoUrl = new PdfPCell(new Phrase("Tipo", new Font(customfontbold, 7)));
            EtiquetaTipoUrl.Colspan = 2;
            EtiquetaTipoUrl.HorizontalAlignment = Element.ALIGN_CENTER;
            EtiquetaTipoUrl.VerticalAlignment = Element.ALIGN_MIDDLE;
            EtiquetaTipoUrl.FixedHeight = 16f;
            EtiquetaTipoUrl.BackgroundColor = new BaseColor(225, 225, 225);
            EtiquetaTipoUrl.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell EtiquetaDetalleUrl = new PdfPCell(new Phrase("Detalle", new Font(customfontbold, 7)));
            EtiquetaDetalleUrl.Colspan = 2;
            EtiquetaDetalleUrl.HorizontalAlignment = Element.ALIGN_CENTER;
            EtiquetaDetalleUrl.VerticalAlignment = Element.ALIGN_MIDDLE;
            EtiquetaDetalleUrl.FixedHeight = 16f;
            EtiquetaDetalleUrl.BackgroundColor = new BaseColor(225, 225, 225);
            EtiquetaDetalleUrl.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell EtiquetaUrlLink = new PdfPCell(new Phrase("Nombre", new Font(customfontbold, 7)));
            EtiquetaUrlLink.Colspan = 3;
            EtiquetaUrlLink.HorizontalAlignment = Element.ALIGN_CENTER;
            EtiquetaUrlLink.VerticalAlignment = Element.ALIGN_MIDDLE;
            EtiquetaUrlLink.FixedHeight = 16f;
            EtiquetaUrlLink.BackgroundColor = new BaseColor(225, 225, 225);
            EtiquetaUrlLink.BorderColor = new BaseColor(60, 66, 82);

            urlsExterno.AddCell(EtiquetaCampoUrls);

            urlsExterno.AddCell(EtiquetaTipoUrl);
            urlsExterno.AddCell(EtiquetaDetalleUrl);
            urlsExterno.AddCell(EtiquetaUrlLink);

            foreach (UrlExternosSolicitud urls in ListadoUrlExterno)
            {
                PdfPCell ValorTipoUrlDetalle = new PdfPCell(new Phrase(urls.tipo.ToString(), new Font(customfont, 7)));
                ValorTipoUrlDetalle.Colspan = 2;
                ValorTipoUrlDetalle.HorizontalAlignment = Element.ALIGN_CENTER;
                ValorTipoUrlDetalle.VerticalAlignment = Element.ALIGN_MIDDLE;
                ValorTipoUrlDetalle.FixedHeight = 23f;
                ValorTipoUrlDetalle.BorderColor = new BaseColor(60, 66, 82);

                PdfPCell ValorDetalleUrlDetalle = new PdfPCell(new Phrase(urls.detalle.ToString(), new Font(customfont, 7)));
                ValorDetalleUrlDetalle.Colspan = 2;
                ValorDetalleUrlDetalle.HorizontalAlignment = Element.ALIGN_CENTER;
                ValorDetalleUrlDetalle.VerticalAlignment = Element.ALIGN_MIDDLE;
                ValorDetalleUrlDetalle.FixedHeight = 23f;
                ValorDetalleUrlDetalle.BorderColor = new BaseColor(60, 66, 82);

                PdfPCell ValorUrlLinkDetalle = new PdfPCell(new Phrase(urls.url.ToString(), new Font(customfont, 7)));
                ValorUrlLinkDetalle.Colspan = 3;
                ValorUrlLinkDetalle.HorizontalAlignment = Element.ALIGN_CENTER;
                ValorUrlLinkDetalle.VerticalAlignment = Element.ALIGN_MIDDLE;
                ValorUrlLinkDetalle.FixedHeight = 23f;
                ValorUrlLinkDetalle.BorderColor = new BaseColor(60, 66, 82);

                urlsExterno.AddCell(ValorTipoUrlDetalle);
                urlsExterno.AddCell(ValorDetalleUrlDetalle);
                urlsExterno.AddCell(ValorUrlLinkDetalle);
            }

            if (ListadoUrlExterno.Count > 0)
            {
                document.Add(urlsExterno);

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
            }

            //Archivos Adjuntos
            PdfPTable camposAdjuntos = new PdfPTable(6);

            PdfPCell EtiquetaPiezaReferencia = new PdfPCell(new Phrase("PIEZA DE REFERENCIA", new Font(customfontbold, 9, 0, new BaseColor(255, 255, 255))));
            EtiquetaPiezaReferencia.Colspan = 6;
            EtiquetaPiezaReferencia.HorizontalAlignment = Element.ALIGN_CENTER;
            EtiquetaPiezaReferencia.VerticalAlignment = Element.ALIGN_MIDDLE;
            EtiquetaPiezaReferencia.FixedHeight = 16f;
            EtiquetaPiezaReferencia.BackgroundColor = new BaseColor(60, 66, 82);
            EtiquetaPiezaReferencia.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell EtiquetaNumero = new PdfPCell(new Phrase("# Adjunto", new Font(customfontbold, 7)));
            EtiquetaNumero.Colspan = 1;
            EtiquetaNumero.HorizontalAlignment = Element.ALIGN_CENTER;
            EtiquetaNumero.VerticalAlignment = Element.ALIGN_MIDDLE;
            EtiquetaNumero.FixedHeight = 16f;
            EtiquetaNumero.BackgroundColor = new BaseColor(225, 225, 225);
            EtiquetaNumero.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell EtiquetaNombreArchivo = new PdfPCell(new Phrase("Nombre Archivo", new Font(customfontbold, 7)));
            EtiquetaNombreArchivo.Colspan = 5;
            EtiquetaNombreArchivo.HorizontalAlignment = Element.ALIGN_LEFT;
            EtiquetaNombreArchivo.VerticalAlignment = Element.ALIGN_MIDDLE;
            EtiquetaNombreArchivo.FixedHeight = 16f;
            EtiquetaNombreArchivo.BackgroundColor = new BaseColor(225, 225, 225);
            EtiquetaNombreArchivo.BorderColor = new BaseColor(60, 66, 82);

            camposAdjuntos.AddCell(EtiquetaPiezaReferencia);

            camposAdjuntos.AddCell(EtiquetaNumero);
            camposAdjuntos.AddCell(EtiquetaNombreArchivo);

            foreach (AdjuntosSolicitud adjunto in ListadoAdjuntos)
            {
                PdfPCell EtiquetaCampoNumero = new PdfPCell(new Phrase(contadorAdjuntos.ToString(), new Font(customfont, 7)));
                EtiquetaCampoNumero.Colspan = 1;
                EtiquetaCampoNumero.HorizontalAlignment = Element.ALIGN_CENTER;
                EtiquetaCampoNumero.VerticalAlignment = Element.ALIGN_MIDDLE;
                EtiquetaCampoNumero.FixedHeight = 23f;
                EtiquetaCampoNumero.BorderColor = new BaseColor(60, 66, 82);

                PdfPCell ValorNombreAdjunto = new PdfPCell(new Phrase(adjunto.nombre_adjunto, new Font(customfont, 7)));
                ValorNombreAdjunto.Colspan = 5;
                ValorNombreAdjunto.HorizontalAlignment = Element.ALIGN_LEFT;
                ValorNombreAdjunto.VerticalAlignment = Element.ALIGN_MIDDLE;
                ValorNombreAdjunto.FixedHeight = 23f;
                ValorNombreAdjunto.BorderColor = new BaseColor(60, 66, 82);

                camposAdjuntos.AddCell(EtiquetaCampoNumero);
                camposAdjuntos.AddCell(ValorNombreAdjunto);

                contadorAdjuntos = contadorAdjuntos + 1;
            }

            if (ListadoAdjuntos.Count > 0)
            {
                document.Add(camposAdjuntos);

                //Salto Linea       
                PdfPTable saltoLinea6 = new PdfPTable(4);

                PdfPCell EtiquetaSaltoLinea6 = new PdfPCell(new Phrase(" ", new Font(customfontbold, 9)));
                EtiquetaSaltoLinea6.Colspan = 4;
                EtiquetaSaltoLinea6.Border = Rectangle.NO_BORDER;
                EtiquetaSaltoLinea6.HorizontalAlignment = Element.ALIGN_CENTER;
                EtiquetaSaltoLinea6.VerticalAlignment = Element.ALIGN_MIDDLE;
                EtiquetaSaltoLinea6.FixedHeight = 8f;

                saltoLinea6.AddCell(EtiquetaSaltoLinea6);

                document.Add(saltoLinea6);
            }

            //Cerrar Documento
            document.Close();

        }

        //Reporte Solicitud Mantenimiento
        public void GenerarPDFMantenimiento(int? id)
        {
            int contadorUrls = 1;
            int contadorAdjuntos = 1;

            //Obtener los datos de la Solicitud
            DatosSolicitudCliente solicitud = db.DatosSolicitudCliente(id).First();

            //Url Externo 
            List<UrlExternosSolicitud> ListadoUrlExterno = db.UrlExternosSolicitud(id).ToList();

            //Adjuntos
            List<AdjuntosSolicitud> ListadoAdjuntos = db.AdjuntosSolicitud(id).ToList();

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

            //anio Actual
            var anioActual = System.DateTime.Now.Year.ToString();

            string basePath = ConfigurationManager.AppSettings["RepositorioDocumentos"];
            string rutaArchivos = basePath + "\\GESTION_PPM\\ADJUNTOS_SOLICITUDES\\" + anioActual + '\\' + solicitud.Id.ToString();

            // En caso de que no exista el directorio, crearlo.
            bool directorio = Directory.Exists(rutaArchivos);
            if (!directorio)
                Directory.CreateDirectory(rutaArchivos);

            //ruta del archivo PDF que se va a crear
            string rutaDocumentos = rutaArchivos + "\\Solicitud-" + solicitud.CodigoSolicitud.ToString() + ".pdf";

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
            PdfPCell TituloCotizador = new PdfPCell(new Phrase("SOLICITUD " + solicitud.Tipo.ToUpper() + " N° " + solicitud.CodigoSolicitud + "      ", new Font(customfontbold, 9, 0, new BaseColor(255, 255, 255))));
            TituloCotizador.Colspan = 5;
            TituloCotizador.Border = Rectangle.NO_BORDER;
            TituloCotizador.HorizontalAlignment = Element.ALIGN_RIGHT;
            TituloCotizador.VerticalAlignment = Element.ALIGN_MIDDLE;
            TituloCotizador.BackgroundColor = new BaseColor(60, 66, 82);
            TituloCotizador.FixedHeight = 50f;

            cabecera.AddCell(logocell);
            cabecera.AddCell(TituloCotizador);

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

            //Informacion de la solicitud         
            PdfPTable informacion = new PdfPTable(6);

            PdfPCell EtiquetaFechaCotizacion = new PdfPCell(new Phrase("Fecha y Hora: ", new Font(customfontbold, 7)));
            EtiquetaFechaCotizacion.Colspan = 1;
            EtiquetaFechaCotizacion.HorizontalAlignment = Element.ALIGN_RIGHT;
            EtiquetaFechaCotizacion.VerticalAlignment = Element.ALIGN_MIDDLE;
            EtiquetaFechaCotizacion.FixedHeight = 16f;
            EtiquetaFechaCotizacion.BackgroundColor = new BaseColor(225, 225, 225);
            EtiquetaFechaCotizacion.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell ValorFechaCotizacion = new PdfPCell(new Phrase("   " + solicitud.FechaHora, new Font(customfont, 7)));
            ValorFechaCotizacion.Colspan = 2;
            ValorFechaCotizacion.HorizontalAlignment = Element.ALIGN_LEFT;
            ValorFechaCotizacion.VerticalAlignment = Element.ALIGN_MIDDLE;
            ValorFechaCotizacion.FixedHeight = 16f;
            ValorFechaCotizacion.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell EtiquetaSolicitante = new PdfPCell(new Phrase("Solicitante:  ", new Font(customfontbold, 7)));
            EtiquetaSolicitante.Colspan = 1;
            EtiquetaSolicitante.HorizontalAlignment = Element.ALIGN_RIGHT;
            EtiquetaSolicitante.VerticalAlignment = Element.ALIGN_MIDDLE;
            EtiquetaSolicitante.FixedHeight = 16f;
            EtiquetaSolicitante.BackgroundColor = new BaseColor(225, 225, 225);
            EtiquetaSolicitante.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell ValorSolicitante = new PdfPCell(new Phrase("   " + solicitud.Solicitante, new Font(customfont, 7)));
            ValorSolicitante.Colspan = 2;
            ValorSolicitante.HorizontalAlignment = Element.ALIGN_LEFT;
            ValorSolicitante.VerticalAlignment = Element.ALIGN_MIDDLE;
            ValorSolicitante.FixedHeight = 16f;
            ValorSolicitante.BorderColor = new BaseColor(60, 66, 82);

            informacion.AddCell(EtiquetaFechaCotizacion);
            informacion.AddCell(ValorFechaCotizacion);
            informacion.AddCell(EtiquetaSolicitante);
            informacion.AddCell(ValorSolicitante);

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

            //Informacion de la Solicitud       
            PdfPTable datosGenerales = new PdfPTable(6);

            PdfPCell EtiquetaTipo = new PdfPCell(new Phrase("Tipo: ", new Font(customfontbold, 7)));
            EtiquetaTipo.Colspan = 1;
            EtiquetaTipo.HorizontalAlignment = Element.ALIGN_RIGHT;
            EtiquetaTipo.VerticalAlignment = Element.ALIGN_MIDDLE;
            EtiquetaTipo.FixedHeight = 16f;
            EtiquetaTipo.BackgroundColor = new BaseColor(225, 225, 225);
            EtiquetaTipo.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell ValorTipo = new PdfPCell(new Phrase("   " + solicitud.Tipo, new Font(customfont, 7)));
            ValorTipo.Colspan = 5;
            ValorTipo.HorizontalAlignment = Element.ALIGN_LEFT;
            ValorTipo.VerticalAlignment = Element.ALIGN_MIDDLE;
            ValorTipo.FixedHeight = 16f;
            ValorTipo.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell EtiquetaSubtipo = new PdfPCell(new Phrase("Subtipo: ", new Font(customfontbold, 7)));
            EtiquetaSubtipo.Colspan = 1;
            EtiquetaSubtipo.HorizontalAlignment = Element.ALIGN_RIGHT;
            EtiquetaSubtipo.VerticalAlignment = Element.ALIGN_MIDDLE;
            EtiquetaSubtipo.FixedHeight = 16f;
            EtiquetaSubtipo.BackgroundColor = new BaseColor(225, 225, 225);
            EtiquetaSubtipo.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell ValorSubtipo = new PdfPCell(new Phrase("   " + solicitud.Subtipo, new Font(customfont, 7)));
            ValorSubtipo.Colspan = 5;
            ValorSubtipo.HorizontalAlignment = Element.ALIGN_LEFT;
            ValorSubtipo.VerticalAlignment = Element.ALIGN_MIDDLE;
            ValorSubtipo.FixedHeight = 16f;
            ValorSubtipo.BorderColor = new BaseColor(60, 66, 82);

            datosGenerales.AddCell(EtiquetaTipo);
            datosGenerales.AddCell(ValorTipo);

            datosGenerales.AddCell(EtiquetaSubtipo);
            datosGenerales.AddCell(ValorSubtipo);

            document.Add(datosGenerales);


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

            //Descripcion del requermiento    
            PdfPTable entregables = new PdfPTable(6);

            PdfPCell EtiquetaDescripcion = new PdfPCell(new Phrase("DESCRIPCIÓN DEL REQUERIMIENTO", new Font(customfontbold, 9, 0, new BaseColor(255, 255, 255))));
            EtiquetaDescripcion.Colspan = 6;
            EtiquetaDescripcion.HorizontalAlignment = Element.ALIGN_CENTER;
            EtiquetaDescripcion.VerticalAlignment = Element.ALIGN_MIDDLE;
            EtiquetaDescripcion.FixedHeight = 16f;
            EtiquetaDescripcion.BackgroundColor = new BaseColor(60, 66, 82);
            EtiquetaDescripcion.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell ValorDescripcion = new PdfPCell(new Phrase(solicitud.Descripcion, new Font(customfont, 7)));
            ValorDescripcion.Colspan = 6;
            ValorDescripcion.Rowspan = 6;
            ValorDescripcion.VerticalAlignment = Element.ALIGN_MIDDLE;
            ValorDescripcion.BorderColor = new BaseColor(60, 66, 82);

            //validar el tamaño de las celdas
            var caracteres = solicitud.Descripcion.Count();
            if (caracteres <= 70)
            {
                ValorDescripcion.FixedHeight = 16f;
            }
            else if (caracteres <= 140)
            {
                ValorDescripcion.FixedHeight = 32f;
            }
            else if (caracteres <= 210)
            {
                ValorDescripcion.FixedHeight = 46f;
            }
            else if (caracteres <= 280)
            {
                ValorDescripcion.FixedHeight = 58f;
            }
            else if (caracteres <= 350)
            {
                ValorDescripcion.FixedHeight = 70f;
            }
            else if (caracteres <= 420)
            {
                ValorDescripcion.FixedHeight = 82f;
            }
            else if (caracteres <= 550)
            {
                ValorDescripcion.FixedHeight = 94f;
            }
            else if (caracteres <= 680)
            {
                ValorDescripcion.FixedHeight = 106f;
            }
            else if (caracteres <= 810)
            {
                ValorDescripcion.FixedHeight = 128f;
            }
            else
            {
                ValorDescripcion.FixedHeight = 140f;
            }

            entregables.AddCell(EtiquetaDescripcion);
            entregables.AddCell(ValorDescripcion);

            document.Add(entregables);

            //Salto Linea       
            PdfPTable saltoLinea4 = new PdfPTable(4);

            PdfPCell EtiquetaSaltoLinea4 = new PdfPCell(new Phrase(" ", new Font(customfontbold, 9)));
            EtiquetaSaltoLinea4.Colspan = 4;
            EtiquetaSaltoLinea4.Border = Rectangle.NO_BORDER;
            EtiquetaSaltoLinea4.HorizontalAlignment = Element.ALIGN_CENTER;
            EtiquetaSaltoLinea4.VerticalAlignment = Element.ALIGN_MIDDLE;
            EtiquetaSaltoLinea4.FixedHeight = 8f;

            saltoLinea4.AddCell(EtiquetaSaltoLinea4);

            document.Add(saltoLinea4);

            //Url Externo  
            PdfPTable urlsExterno = new PdfPTable(6);

            PdfPCell EtiquetaCampoUrls = new PdfPCell(new Phrase("URL's de Soporte", new Font(customfontbold, 9, 0, new BaseColor(255, 255, 255))));
            EtiquetaCampoUrls.Colspan = 6;
            EtiquetaCampoUrls.HorizontalAlignment = Element.ALIGN_CENTER;
            EtiquetaCampoUrls.VerticalAlignment = Element.ALIGN_MIDDLE;
            EtiquetaCampoUrls.FixedHeight = 16f;
            EtiquetaCampoUrls.BackgroundColor = new BaseColor(60, 66, 82);
            EtiquetaCampoUrls.BorderColor = new BaseColor(60, 66, 82);

            urlsExterno.AddCell(EtiquetaCampoUrls);

            foreach (UrlExternosSolicitud urls in ListadoUrlExterno)
            {
                PdfPCell ValorTipoUrlDetalle = new PdfPCell(new Phrase(contadorUrls.ToString(), new Font(customfont, 7)));
                ValorTipoUrlDetalle.Colspan = 1;
                ValorTipoUrlDetalle.HorizontalAlignment = Element.ALIGN_CENTER;
                ValorTipoUrlDetalle.VerticalAlignment = Element.ALIGN_MIDDLE;
                ValorTipoUrlDetalle.FixedHeight = 23f;
                ValorTipoUrlDetalle.BorderColor = new BaseColor(60, 66, 82);

                PdfPCell ValorDetalleUrlDetalle = new PdfPCell(new Phrase(urls.url.ToString(), new Font(customfont, 7)));
                ValorDetalleUrlDetalle.Colspan = 5;
                ValorDetalleUrlDetalle.HorizontalAlignment = Element.ALIGN_LEFT;
                ValorDetalleUrlDetalle.VerticalAlignment = Element.ALIGN_MIDDLE;
                ValorDetalleUrlDetalle.FixedHeight = 23f;
                ValorDetalleUrlDetalle.BorderColor = new BaseColor(60, 66, 82);

                urlsExterno.AddCell(ValorTipoUrlDetalle);
                urlsExterno.AddCell(ValorDetalleUrlDetalle);
            }

            if (ListadoUrlExterno.Count > 0)
            {
                document.Add(urlsExterno);

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
            }
            //Archivos Adjuntos
            PdfPTable camposAdjuntos = new PdfPTable(6);

            PdfPCell EtiquetaPiezaReferencia = new PdfPCell(new Phrase("Archivos Adjunto Soporte", new Font(customfontbold, 9, 0, new BaseColor(255, 255, 255))));
            EtiquetaPiezaReferencia.Colspan = 6;
            EtiquetaPiezaReferencia.HorizontalAlignment = Element.ALIGN_CENTER;
            EtiquetaPiezaReferencia.VerticalAlignment = Element.ALIGN_MIDDLE;
            EtiquetaPiezaReferencia.FixedHeight = 16f;
            EtiquetaPiezaReferencia.BackgroundColor = new BaseColor(60, 66, 82);
            EtiquetaPiezaReferencia.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell EtiquetaNumero = new PdfPCell(new Phrase("# Adjunto", new Font(customfontbold, 7)));
            EtiquetaNumero.Colspan = 1;
            EtiquetaNumero.HorizontalAlignment = Element.ALIGN_CENTER;
            EtiquetaNumero.VerticalAlignment = Element.ALIGN_MIDDLE;
            EtiquetaNumero.FixedHeight = 16f;
            EtiquetaNumero.BackgroundColor = new BaseColor(225, 225, 225);
            EtiquetaNumero.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell EtiquetaNombreArchivo = new PdfPCell(new Phrase("Nombre Archivo", new Font(customfontbold, 7)));
            EtiquetaNombreArchivo.Colspan = 5;
            EtiquetaNombreArchivo.HorizontalAlignment = Element.ALIGN_LEFT;
            EtiquetaNombreArchivo.VerticalAlignment = Element.ALIGN_MIDDLE;
            EtiquetaNombreArchivo.FixedHeight = 16f;
            EtiquetaNombreArchivo.BackgroundColor = new BaseColor(225, 225, 225);
            EtiquetaNombreArchivo.BorderColor = new BaseColor(60, 66, 82);

            camposAdjuntos.AddCell(EtiquetaPiezaReferencia);

            camposAdjuntos.AddCell(EtiquetaNumero);
            camposAdjuntos.AddCell(EtiquetaNombreArchivo);

            foreach (AdjuntosSolicitud adjunto in ListadoAdjuntos)
            {
                PdfPCell EtiquetaCampoNumero = new PdfPCell(new Phrase(contadorAdjuntos.ToString(), new Font(customfont, 7)));
                EtiquetaCampoNumero.Colspan = 1;
                EtiquetaCampoNumero.HorizontalAlignment = Element.ALIGN_CENTER;
                EtiquetaCampoNumero.VerticalAlignment = Element.ALIGN_MIDDLE;
                EtiquetaCampoNumero.FixedHeight = 23f;
                EtiquetaCampoNumero.BorderColor = new BaseColor(60, 66, 82);

                PdfPCell ValorNombreAdjunto = new PdfPCell(new Phrase(adjunto.nombre_adjunto, new Font(customfont, 7)));
                ValorNombreAdjunto.Colspan = 5;
                ValorNombreAdjunto.HorizontalAlignment = Element.ALIGN_LEFT;
                ValorNombreAdjunto.VerticalAlignment = Element.ALIGN_MIDDLE;
                ValorNombreAdjunto.FixedHeight = 23f;
                ValorNombreAdjunto.BorderColor = new BaseColor(60, 66, 82);

                camposAdjuntos.AddCell(EtiquetaCampoNumero);
                camposAdjuntos.AddCell(ValorNombreAdjunto);

                contadorAdjuntos = contadorAdjuntos + 1;
            }

            if (ListadoAdjuntos.Count > 0)
            {
                document.Add(camposAdjuntos);

                //Salto Linea       
                PdfPTable saltoLinea6 = new PdfPTable(4);

                PdfPCell EtiquetaSaltoLinea6 = new PdfPCell(new Phrase(" ", new Font(customfontbold, 9)));
                EtiquetaSaltoLinea6.Colspan = 4;
                EtiquetaSaltoLinea6.Border = Rectangle.NO_BORDER;
                EtiquetaSaltoLinea6.HorizontalAlignment = Element.ALIGN_CENTER;
                EtiquetaSaltoLinea6.VerticalAlignment = Element.ALIGN_MIDDLE;
                EtiquetaSaltoLinea6.FixedHeight = 8f;

                saltoLinea6.AddCell(EtiquetaSaltoLinea6);
            }

            document.Close();
        }

        //Reporte Solicitud Envio Icare
        public void GenerarPDFEnvioIcare(int? id)
        {
            //Obtener los datos de la Solicitud
            DatosSolicitudCliente solicitud = db.DatosSolicitudCliente(id).First();

            //Adjuntos
            List<AdjuntosSolicitud> ListadoAdjuntos = db.AdjuntosSolicitud(id).ToList();

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

            //anio Actual
            var anioActual = System.DateTime.Now.Year.ToString();

            string basePath = ConfigurationManager.AppSettings["RepositorioDocumentos"];
            string rutaArchivos = basePath + "\\GESTION_PPM\\ADJUNTOS_SOLICITUDES\\" + anioActual + '\\' + solicitud.Id.ToString();

            // En caso de que no exista el directorio, crearlo.
            bool directorio = Directory.Exists(rutaArchivos);
            if (!directorio)
                Directory.CreateDirectory(rutaArchivos);

            //ruta del archivo PDF que se va a crear
            string rutaDocumentos = rutaArchivos + "\\Solicitud-" + solicitud.CodigoSolicitud.ToString() + ".pdf";

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
            PdfPCell TituloCotizador = new PdfPCell(new Phrase("SOLICITUD " + solicitud.Tipo.ToUpper() + " N° " + solicitud.CodigoSolicitud + "      ", new Font(customfontbold, 9, 0, new BaseColor(255, 255, 255))));
            TituloCotizador.Colspan = 5;
            TituloCotizador.Border = Rectangle.NO_BORDER;
            TituloCotizador.HorizontalAlignment = Element.ALIGN_RIGHT;
            TituloCotizador.VerticalAlignment = Element.ALIGN_MIDDLE;
            TituloCotizador.BackgroundColor = new BaseColor(60, 66, 82);
            TituloCotizador.FixedHeight = 50f;

            cabecera.AddCell(logocell);
            cabecera.AddCell(TituloCotizador);

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

            //Informacion de la solicitud         
            PdfPTable informacion = new PdfPTable(6);

            PdfPCell EtiquetaFechaCotizacion = new PdfPCell(new Phrase("Fecha y Hora: ", new Font(customfontbold, 7)));
            EtiquetaFechaCotizacion.Colspan = 1;
            EtiquetaFechaCotizacion.HorizontalAlignment = Element.ALIGN_RIGHT;
            EtiquetaFechaCotizacion.VerticalAlignment = Element.ALIGN_MIDDLE;
            EtiquetaFechaCotizacion.FixedHeight = 16f;
            EtiquetaFechaCotizacion.BackgroundColor = new BaseColor(225, 225, 225);
            EtiquetaFechaCotizacion.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell ValorFechaCotizacion = new PdfPCell(new Phrase("   " + solicitud.FechaHora, new Font(customfont, 7)));
            ValorFechaCotizacion.Colspan = 2;
            ValorFechaCotizacion.HorizontalAlignment = Element.ALIGN_LEFT;
            ValorFechaCotizacion.VerticalAlignment = Element.ALIGN_MIDDLE;
            ValorFechaCotizacion.FixedHeight = 16f;
            ValorFechaCotizacion.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell EtiquetaSolicitante = new PdfPCell(new Phrase("Solicitante:  ", new Font(customfontbold, 7)));
            EtiquetaSolicitante.Colspan = 1;
            EtiquetaSolicitante.HorizontalAlignment = Element.ALIGN_RIGHT;
            EtiquetaSolicitante.VerticalAlignment = Element.ALIGN_MIDDLE;
            EtiquetaSolicitante.FixedHeight = 16f;
            EtiquetaSolicitante.BackgroundColor = new BaseColor(225, 225, 225);
            EtiquetaSolicitante.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell ValorSolicitante = new PdfPCell(new Phrase("   " + solicitud.Solicitante, new Font(customfont, 7)));
            ValorSolicitante.Colspan = 2;
            ValorSolicitante.HorizontalAlignment = Element.ALIGN_LEFT;
            ValorSolicitante.VerticalAlignment = Element.ALIGN_MIDDLE;
            ValorSolicitante.FixedHeight = 16f;
            ValorSolicitante.BorderColor = new BaseColor(60, 66, 82);

            informacion.AddCell(EtiquetaFechaCotizacion);
            informacion.AddCell(ValorFechaCotizacion);
            informacion.AddCell(EtiquetaSolicitante);
            informacion.AddCell(ValorSolicitante);

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

            //Informacion de la Solicitud       
            PdfPTable datosGenerales = new PdfPTable(6);

            PdfPCell EtiquetaTipo = new PdfPCell(new Phrase("Tipo: ", new Font(customfontbold, 7)));
            EtiquetaTipo.Colspan = 1;
            EtiquetaTipo.HorizontalAlignment = Element.ALIGN_RIGHT;
            EtiquetaTipo.VerticalAlignment = Element.ALIGN_MIDDLE;
            EtiquetaTipo.FixedHeight = 16f;
            EtiquetaTipo.BackgroundColor = new BaseColor(225, 225, 225);
            EtiquetaTipo.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell ValorTipo = new PdfPCell(new Phrase("   " + solicitud.Tipo, new Font(customfont, 7)));
            ValorTipo.Colspan = 5;
            ValorTipo.HorizontalAlignment = Element.ALIGN_LEFT;
            ValorTipo.VerticalAlignment = Element.ALIGN_MIDDLE;
            ValorTipo.FixedHeight = 16f;
            ValorTipo.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell EtiquetaSubtipo = new PdfPCell(new Phrase("Subtipo: ", new Font(customfontbold, 7)));
            EtiquetaSubtipo.Colspan = 1;
            EtiquetaSubtipo.HorizontalAlignment = Element.ALIGN_RIGHT;
            EtiquetaSubtipo.VerticalAlignment = Element.ALIGN_MIDDLE;
            EtiquetaSubtipo.FixedHeight = 16f;
            EtiquetaSubtipo.BackgroundColor = new BaseColor(225, 225, 225);
            EtiquetaSubtipo.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell ValorSubtipo = new PdfPCell(new Phrase("   " + solicitud.Subtipo, new Font(customfont, 7)));
            ValorSubtipo.Colspan = 5;
            ValorSubtipo.HorizontalAlignment = Element.ALIGN_LEFT;
            ValorSubtipo.VerticalAlignment = Element.ALIGN_MIDDLE;
            ValorSubtipo.FixedHeight = 16f;
            ValorSubtipo.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell EtiquetaOP = new PdfPCell(new Phrase("OP: ", new Font(customfontbold, 7)));
            EtiquetaOP.Colspan = 1;
            EtiquetaOP.HorizontalAlignment = Element.ALIGN_RIGHT;
            EtiquetaOP.VerticalAlignment = Element.ALIGN_MIDDLE;
            EtiquetaOP.FixedHeight = 16f;
            EtiquetaOP.BackgroundColor = new BaseColor(225, 225, 225);
            EtiquetaOP.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell ValorOP = new PdfPCell(new Phrase("   " + solicitud.OP, new Font(customfont, 7)));
            ValorOP.Colspan = 2;
            ValorOP.HorizontalAlignment = Element.ALIGN_LEFT;
            ValorOP.VerticalAlignment = Element.ALIGN_MIDDLE;
            ValorOP.FixedHeight = 16f;
            ValorOP.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell EtiquetaMKT = new PdfPCell(new Phrase("MKT: ", new Font(customfontbold, 7)));
            EtiquetaMKT.Colspan = 1;
            EtiquetaMKT.HorizontalAlignment = Element.ALIGN_RIGHT;
            EtiquetaMKT.VerticalAlignment = Element.ALIGN_MIDDLE;
            EtiquetaMKT.FixedHeight = 16f;
            EtiquetaMKT.BackgroundColor = new BaseColor(225, 225, 225);
            EtiquetaMKT.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell ValorMKT = new PdfPCell(new Phrase("   " + solicitud.MKT, new Font(customfont, 7)));
            ValorMKT.Colspan = 2;
            ValorMKT.HorizontalAlignment = Element.ALIGN_LEFT;
            ValorMKT.VerticalAlignment = Element.ALIGN_MIDDLE;
            ValorMKT.FixedHeight = 16f;
            ValorMKT.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell EtiquetaNombre = new PdfPCell(new Phrase("Nombre Campaña: ", new Font(customfontbold, 7)));
            EtiquetaNombre.Colspan = 1;
            EtiquetaNombre.HorizontalAlignment = Element.ALIGN_RIGHT;
            EtiquetaNombre.VerticalAlignment = Element.ALIGN_MIDDLE;
            EtiquetaNombre.FixedHeight = 16f;
            EtiquetaNombre.BackgroundColor = new BaseColor(225, 225, 225);
            EtiquetaNombre.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell ValorNombre = new PdfPCell(new Phrase("   " + solicitud.Nombre, new Font(customfont, 7)));
            ValorNombre.Colspan = 5;
            ValorNombre.HorizontalAlignment = Element.ALIGN_LEFT;
            ValorNombre.VerticalAlignment = Element.ALIGN_MIDDLE;
            ValorNombre.FixedHeight = 16f;
            ValorNombre.BorderColor = new BaseColor(60, 66, 82);

            datosGenerales.AddCell(EtiquetaTipo);
            datosGenerales.AddCell(ValorTipo);

            datosGenerales.AddCell(EtiquetaSubtipo);
            datosGenerales.AddCell(ValorSubtipo);

            datosGenerales.AddCell(EtiquetaOP);
            datosGenerales.AddCell(ValorOP);
            datosGenerales.AddCell(EtiquetaMKT);
            datosGenerales.AddCell(ValorMKT);

            datosGenerales.AddCell(EtiquetaNombre);
            datosGenerales.AddCell(ValorNombre);

            document.Add(datosGenerales);

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

            //Archivos Adjuntos
            PdfPTable camposAdjuntos = new PdfPTable(6);

            PdfPCell EtiquetaPiezaReferencia = new PdfPCell(new Phrase("Archivos Arjuntos", new Font(customfontbold, 9, 0, new BaseColor(255, 255, 255))));
            EtiquetaPiezaReferencia.Colspan = 6;
            EtiquetaPiezaReferencia.HorizontalAlignment = Element.ALIGN_CENTER;
            EtiquetaPiezaReferencia.VerticalAlignment = Element.ALIGN_MIDDLE;
            EtiquetaPiezaReferencia.FixedHeight = 16f;
            EtiquetaPiezaReferencia.BackgroundColor = new BaseColor(60, 66, 82);
            EtiquetaPiezaReferencia.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell EtiquetaNumero = new PdfPCell(new Phrase("Tipo", new Font(customfontbold, 7)));
            EtiquetaNumero.Colspan = 2;
            EtiquetaNumero.HorizontalAlignment = Element.ALIGN_CENTER;
            EtiquetaNumero.VerticalAlignment = Element.ALIGN_MIDDLE;
            EtiquetaNumero.FixedHeight = 16f;
            EtiquetaNumero.BackgroundColor = new BaseColor(225, 225, 225);
            EtiquetaNumero.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell EtiquetaNombreArchivo = new PdfPCell(new Phrase("Nombre Archivo", new Font(customfontbold, 7)));
            EtiquetaNombreArchivo.Colspan = 4;
            EtiquetaNombreArchivo.HorizontalAlignment = Element.ALIGN_LEFT;
            EtiquetaNombreArchivo.VerticalAlignment = Element.ALIGN_MIDDLE;
            EtiquetaNombreArchivo.FixedHeight = 16f;
            EtiquetaNombreArchivo.BackgroundColor = new BaseColor(225, 225, 225);
            EtiquetaNombreArchivo.BorderColor = new BaseColor(60, 66, 82);

            camposAdjuntos.AddCell(EtiquetaPiezaReferencia);

            camposAdjuntos.AddCell(EtiquetaNumero);
            camposAdjuntos.AddCell(EtiquetaNombreArchivo);

            foreach (AdjuntosSolicitud adjunto in ListadoAdjuntos)
            {
                PdfPCell EtiquetaCampoNumero = new PdfPCell(new Phrase(adjunto.tipo.ToString(), new Font(customfont, 7)));
                EtiquetaCampoNumero.Colspan = 2;
                EtiquetaCampoNumero.HorizontalAlignment = Element.ALIGN_CENTER;
                EtiquetaCampoNumero.VerticalAlignment = Element.ALIGN_MIDDLE;
                EtiquetaCampoNumero.FixedHeight = 23f;
                EtiquetaCampoNumero.BorderColor = new BaseColor(60, 66, 82);

                PdfPCell ValorNombreAdjunto = new PdfPCell(new Phrase(adjunto.nombre_adjunto, new Font(customfont, 7)));
                ValorNombreAdjunto.Colspan = 4;
                ValorNombreAdjunto.HorizontalAlignment = Element.ALIGN_LEFT;
                ValorNombreAdjunto.VerticalAlignment = Element.ALIGN_MIDDLE;
                ValorNombreAdjunto.FixedHeight = 23f;
                ValorNombreAdjunto.BorderColor = new BaseColor(60, 66, 82);

                camposAdjuntos.AddCell(EtiquetaCampoNumero);
                camposAdjuntos.AddCell(ValorNombreAdjunto);

            }

            document.Add(camposAdjuntos);

            //Salto Linea       
            PdfPTable saltoLinea6 = new PdfPTable(4);

            PdfPCell EtiquetaSaltoLinea6 = new PdfPCell(new Phrase(" ", new Font(customfontbold, 9)));
            EtiquetaSaltoLinea6.Colspan = 4;
            EtiquetaSaltoLinea6.Border = Rectangle.NO_BORDER;
            EtiquetaSaltoLinea6.HorizontalAlignment = Element.ALIGN_CENTER;
            EtiquetaSaltoLinea6.VerticalAlignment = Element.ALIGN_MIDDLE;
            EtiquetaSaltoLinea6.FixedHeight = 8f;

            saltoLinea6.AddCell(EtiquetaSaltoLinea6);

            document.Add(saltoLinea6);

            //Cerrar Documento
            document.Close();

        }

        //Reporte Solicitud Html5
        public void GenerarPDFHtml5(int? id)
        {
            //Obtener los datos de la Solicitud
            DatosSolicitudCliente solicitud = db.DatosSolicitudCliente(id).First();

            //Adjuntos
            List<AdjuntosSolicitud> ListadoAdjuntos = db.AdjuntosSolicitud(id).ToList();

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

            //anio Actual
            var anioActual = System.DateTime.Now.Year.ToString();

            string basePath = ConfigurationManager.AppSettings["RepositorioDocumentos"];
            string rutaArchivos = basePath + "\\GESTION_PPM\\ADJUNTOS_SOLICITUDES\\" + anioActual + '\\' + solicitud.Id.ToString();

            // En caso de que no exista el directorio, crearlo.
            bool directorio = Directory.Exists(rutaArchivos);
            if (!directorio)
                Directory.CreateDirectory(rutaArchivos);

            //ruta del archivo PDF que se va a crear
            string rutaDocumentos = rutaArchivos + "\\Solicitud-" + solicitud.CodigoSolicitud.ToString() + ".pdf";

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
            PdfPCell TituloCotizador = new PdfPCell(new Phrase("SOLICITUD " + solicitud.Tipo.ToUpper() + " N° " + solicitud.CodigoSolicitud + "      ", new Font(customfontbold, 9, 0, new BaseColor(255, 255, 255))));
            TituloCotizador.Colspan = 5;
            TituloCotizador.Border = Rectangle.NO_BORDER;
            TituloCotizador.HorizontalAlignment = Element.ALIGN_RIGHT;
            TituloCotizador.VerticalAlignment = Element.ALIGN_MIDDLE;
            TituloCotizador.BackgroundColor = new BaseColor(60, 66, 82);
            TituloCotizador.FixedHeight = 50f;

            cabecera.AddCell(logocell);
            cabecera.AddCell(TituloCotizador);

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

            //Informacion de la solicitud         
            PdfPTable informacion = new PdfPTable(6);

            PdfPCell EtiquetaFechaCotizacion = new PdfPCell(new Phrase("Fecha y Hora: ", new Font(customfontbold, 7)));
            EtiquetaFechaCotizacion.Colspan = 1;
            EtiquetaFechaCotizacion.HorizontalAlignment = Element.ALIGN_RIGHT;
            EtiquetaFechaCotizacion.VerticalAlignment = Element.ALIGN_MIDDLE;
            EtiquetaFechaCotizacion.FixedHeight = 16f;
            EtiquetaFechaCotizacion.BackgroundColor = new BaseColor(225, 225, 225);
            EtiquetaFechaCotizacion.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell ValorFechaCotizacion = new PdfPCell(new Phrase("   " + solicitud.FechaHora, new Font(customfont, 7)));
            ValorFechaCotizacion.Colspan = 2;
            ValorFechaCotizacion.HorizontalAlignment = Element.ALIGN_LEFT;
            ValorFechaCotizacion.VerticalAlignment = Element.ALIGN_MIDDLE;
            ValorFechaCotizacion.FixedHeight = 16f;
            ValorFechaCotizacion.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell EtiquetaSolicitante = new PdfPCell(new Phrase("Solicitante:  ", new Font(customfontbold, 7)));
            EtiquetaSolicitante.Colspan = 1;
            EtiquetaSolicitante.HorizontalAlignment = Element.ALIGN_RIGHT;
            EtiquetaSolicitante.VerticalAlignment = Element.ALIGN_MIDDLE;
            EtiquetaSolicitante.FixedHeight = 16f;
            EtiquetaSolicitante.BackgroundColor = new BaseColor(225, 225, 225);
            EtiquetaSolicitante.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell ValorSolicitante = new PdfPCell(new Phrase("   " + solicitud.Solicitante, new Font(customfont, 7)));
            ValorSolicitante.Colspan = 2;
            ValorSolicitante.HorizontalAlignment = Element.ALIGN_LEFT;
            ValorSolicitante.VerticalAlignment = Element.ALIGN_MIDDLE;
            ValorSolicitante.FixedHeight = 16f;
            ValorSolicitante.BorderColor = new BaseColor(60, 66, 82);

            informacion.AddCell(EtiquetaFechaCotizacion);
            informacion.AddCell(ValorFechaCotizacion);
            informacion.AddCell(EtiquetaSolicitante);
            informacion.AddCell(ValorSolicitante);

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

            //Informacion de la Solicitud       
            PdfPTable datosGenerales = new PdfPTable(6);

            PdfPCell EtiquetaTipo = new PdfPCell(new Phrase("Tipo: ", new Font(customfontbold, 7)));
            EtiquetaTipo.Colspan = 1;
            EtiquetaTipo.HorizontalAlignment = Element.ALIGN_RIGHT;
            EtiquetaTipo.VerticalAlignment = Element.ALIGN_MIDDLE;
            EtiquetaTipo.FixedHeight = 16f;
            EtiquetaTipo.BackgroundColor = new BaseColor(225, 225, 225);
            EtiquetaTipo.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell ValorTipo = new PdfPCell(new Phrase("   " + solicitud.Tipo, new Font(customfont, 7)));
            ValorTipo.Colspan = 5;
            ValorTipo.HorizontalAlignment = Element.ALIGN_LEFT;
            ValorTipo.VerticalAlignment = Element.ALIGN_MIDDLE;
            ValorTipo.FixedHeight = 16f;
            ValorTipo.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell EtiquetaSubtipo = new PdfPCell(new Phrase("Subtipo: ", new Font(customfontbold, 7)));
            EtiquetaSubtipo.Colspan = 1;
            EtiquetaSubtipo.HorizontalAlignment = Element.ALIGN_RIGHT;
            EtiquetaSubtipo.VerticalAlignment = Element.ALIGN_MIDDLE;
            EtiquetaSubtipo.FixedHeight = 16f;
            EtiquetaSubtipo.BackgroundColor = new BaseColor(225, 225, 225);
            EtiquetaSubtipo.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell ValorSubtipo = new PdfPCell(new Phrase("   " + solicitud.Subtipo, new Font(customfont, 7)));
            ValorSubtipo.Colspan = 5;
            ValorSubtipo.HorizontalAlignment = Element.ALIGN_LEFT;
            ValorSubtipo.VerticalAlignment = Element.ALIGN_MIDDLE;
            ValorSubtipo.FixedHeight = 16f;
            ValorSubtipo.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell EtiquetaOP = new PdfPCell(new Phrase("OP: ", new Font(customfontbold, 7)));
            EtiquetaOP.Colspan = 1;
            EtiquetaOP.HorizontalAlignment = Element.ALIGN_RIGHT;
            EtiquetaOP.VerticalAlignment = Element.ALIGN_MIDDLE;
            EtiquetaOP.FixedHeight = 16f;
            EtiquetaOP.BackgroundColor = new BaseColor(225, 225, 225);
            EtiquetaOP.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell ValorOP = new PdfPCell(new Phrase("   " + solicitud.OP, new Font(customfont, 7)));
            ValorOP.Colspan = 2;
            ValorOP.HorizontalAlignment = Element.ALIGN_LEFT;
            ValorOP.VerticalAlignment = Element.ALIGN_MIDDLE;
            ValorOP.FixedHeight = 16f;
            ValorOP.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell EtiquetaMKT = new PdfPCell(new Phrase("MKT: ", new Font(customfontbold, 7)));
            EtiquetaMKT.Colspan = 1;
            EtiquetaMKT.HorizontalAlignment = Element.ALIGN_RIGHT;
            EtiquetaMKT.VerticalAlignment = Element.ALIGN_MIDDLE;
            EtiquetaMKT.FixedHeight = 16f;
            EtiquetaMKT.BackgroundColor = new BaseColor(225, 225, 225);
            EtiquetaMKT.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell ValorMKT = new PdfPCell(new Phrase("   " + solicitud.MKT, new Font(customfont, 7)));
            ValorMKT.Colspan = 2;
            ValorMKT.HorizontalAlignment = Element.ALIGN_LEFT;
            ValorMKT.VerticalAlignment = Element.ALIGN_MIDDLE;
            ValorMKT.FixedHeight = 16f;
            ValorMKT.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell EtiquetaNombre = new PdfPCell(new Phrase("Nombre Campaña: ", new Font(customfontbold, 7)));
            EtiquetaNombre.Colspan = 1;
            EtiquetaNombre.HorizontalAlignment = Element.ALIGN_RIGHT;
            EtiquetaNombre.VerticalAlignment = Element.ALIGN_MIDDLE;
            EtiquetaNombre.FixedHeight = 16f;
            EtiquetaNombre.BackgroundColor = new BaseColor(225, 225, 225);
            EtiquetaNombre.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell ValorNombre = new PdfPCell(new Phrase("   " + solicitud.Nombre, new Font(customfont, 7)));
            ValorNombre.Colspan = 5;
            ValorNombre.HorizontalAlignment = Element.ALIGN_LEFT;
            ValorNombre.VerticalAlignment = Element.ALIGN_MIDDLE;
            ValorNombre.FixedHeight = 16f;
            ValorNombre.BorderColor = new BaseColor(60, 66, 82);

            datosGenerales.AddCell(EtiquetaTipo);
            datosGenerales.AddCell(ValorTipo);

            datosGenerales.AddCell(EtiquetaSubtipo);
            datosGenerales.AddCell(ValorSubtipo);

            datosGenerales.AddCell(EtiquetaOP);
            datosGenerales.AddCell(ValorOP);
            datosGenerales.AddCell(EtiquetaMKT);
            datosGenerales.AddCell(ValorMKT);

            datosGenerales.AddCell(EtiquetaNombre);
            datosGenerales.AddCell(ValorNombre);

            document.Add(datosGenerales);

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

            //Archivos Adjuntos
            PdfPTable camposAdjuntos = new PdfPTable(6);

            PdfPCell EtiquetaPiezaReferencia = new PdfPCell(new Phrase("Archivos Arjuntos", new Font(customfontbold, 9, 0, new BaseColor(255, 255, 255))));
            EtiquetaPiezaReferencia.Colspan = 6;
            EtiquetaPiezaReferencia.HorizontalAlignment = Element.ALIGN_CENTER;
            EtiquetaPiezaReferencia.VerticalAlignment = Element.ALIGN_MIDDLE;
            EtiquetaPiezaReferencia.FixedHeight = 16f;
            EtiquetaPiezaReferencia.BackgroundColor = new BaseColor(60, 66, 82);
            EtiquetaPiezaReferencia.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell EtiquetaNumero = new PdfPCell(new Phrase("Tipo", new Font(customfontbold, 7)));
            EtiquetaNumero.Colspan = 2;
            EtiquetaNumero.HorizontalAlignment = Element.ALIGN_CENTER;
            EtiquetaNumero.VerticalAlignment = Element.ALIGN_MIDDLE;
            EtiquetaNumero.FixedHeight = 16f;
            EtiquetaNumero.BackgroundColor = new BaseColor(225, 225, 225);
            EtiquetaNumero.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell EtiquetaNombreArchivo = new PdfPCell(new Phrase("Nombre Archivo", new Font(customfontbold, 7)));
            EtiquetaNombreArchivo.Colspan = 4;
            EtiquetaNombreArchivo.HorizontalAlignment = Element.ALIGN_LEFT;
            EtiquetaNombreArchivo.VerticalAlignment = Element.ALIGN_MIDDLE;
            EtiquetaNombreArchivo.FixedHeight = 16f;
            EtiquetaNombreArchivo.BackgroundColor = new BaseColor(225, 225, 225);
            EtiquetaNombreArchivo.BorderColor = new BaseColor(60, 66, 82);

            camposAdjuntos.AddCell(EtiquetaPiezaReferencia);

            camposAdjuntos.AddCell(EtiquetaNumero);
            camposAdjuntos.AddCell(EtiquetaNombreArchivo);

            foreach (AdjuntosSolicitud adjunto in ListadoAdjuntos)
            {
                PdfPCell EtiquetaCampoNumero = new PdfPCell(new Phrase(adjunto.tipo.ToString(), new Font(customfont, 7)));
                EtiquetaCampoNumero.Colspan = 2;
                EtiquetaCampoNumero.HorizontalAlignment = Element.ALIGN_CENTER;
                EtiquetaCampoNumero.VerticalAlignment = Element.ALIGN_MIDDLE;
                EtiquetaCampoNumero.FixedHeight = 23f;
                EtiquetaCampoNumero.BorderColor = new BaseColor(60, 66, 82);

                PdfPCell ValorNombreAdjunto = new PdfPCell(new Phrase(adjunto.nombre_adjunto, new Font(customfont, 7)));
                ValorNombreAdjunto.Colspan = 4;
                ValorNombreAdjunto.HorizontalAlignment = Element.ALIGN_LEFT;
                ValorNombreAdjunto.VerticalAlignment = Element.ALIGN_MIDDLE;
                ValorNombreAdjunto.FixedHeight = 23f;
                ValorNombreAdjunto.BorderColor = new BaseColor(60, 66, 82);

                camposAdjuntos.AddCell(EtiquetaCampoNumero);
                camposAdjuntos.AddCell(ValorNombreAdjunto);

            }

            document.Add(camposAdjuntos);

            //Salto Linea       
            PdfPTable saltoLinea6 = new PdfPTable(4);

            PdfPCell EtiquetaSaltoLinea6 = new PdfPCell(new Phrase(" ", new Font(customfontbold, 9)));
            EtiquetaSaltoLinea6.Colspan = 4;
            EtiquetaSaltoLinea6.Border = Rectangle.NO_BORDER;
            EtiquetaSaltoLinea6.HorizontalAlignment = Element.ALIGN_CENTER;
            EtiquetaSaltoLinea6.VerticalAlignment = Element.ALIGN_MIDDLE;
            EtiquetaSaltoLinea6.FixedHeight = 8f;

            saltoLinea6.AddCell(EtiquetaSaltoLinea6);

            document.Add(saltoLinea6);

            //Cerrar Documento
            document.Close();

        }

        //Reporte Solicitud Tags
        public void GenerarPDFTags(int? id)
        {
            int contadorUrls = 1;

            //Obtener los datos de la Solicitud
            DatosSolicitudCliente solicitud = db.DatosSolicitudCliente(id).First();

            //Url Externo 
            List<UrlExternosSolicitud> ListadoUrlExterno = db.UrlExternosSolicitud(id).ToList();

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

            //anio Actual
            var anioActual = System.DateTime.Now.Year.ToString();

            string basePath = ConfigurationManager.AppSettings["RepositorioDocumentos"];
            string rutaArchivos = basePath + "\\GESTION_PPM\\ADJUNTOS_SOLICITUDES\\" + anioActual + '\\' + solicitud.Id.ToString();

            // En caso de que no exista el directorio, crearlo.
            bool directorio = Directory.Exists(rutaArchivos);
            if (!directorio)
                Directory.CreateDirectory(rutaArchivos);

            //ruta del archivo PDF que se va a crear
            string rutaDocumentos = rutaArchivos + "\\Solicitud-" + solicitud.CodigoSolicitud.ToString() + ".pdf";

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
            PdfPCell TituloCotizador = new PdfPCell(new Phrase("SOLICITUD " + solicitud.Tipo.ToUpper() + " N° " + solicitud.CodigoSolicitud + "      ", new Font(customfontbold, 9, 0, new BaseColor(255, 255, 255))));
            TituloCotizador.Colspan = 5;
            TituloCotizador.Border = Rectangle.NO_BORDER;
            TituloCotizador.HorizontalAlignment = Element.ALIGN_RIGHT;
            TituloCotizador.VerticalAlignment = Element.ALIGN_MIDDLE;
            TituloCotizador.BackgroundColor = new BaseColor(60, 66, 82);
            TituloCotizador.FixedHeight = 50f;

            cabecera.AddCell(logocell);
            cabecera.AddCell(TituloCotizador);

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

            //Informacion de la solicitud         
            PdfPTable informacion = new PdfPTable(6);

            PdfPCell EtiquetaFechaCotizacion = new PdfPCell(new Phrase("Fecha y Hora: ", new Font(customfontbold, 7)));
            EtiquetaFechaCotizacion.Colspan = 1;
            EtiquetaFechaCotizacion.HorizontalAlignment = Element.ALIGN_RIGHT;
            EtiquetaFechaCotizacion.VerticalAlignment = Element.ALIGN_MIDDLE;
            EtiquetaFechaCotizacion.FixedHeight = 16f;
            EtiquetaFechaCotizacion.BackgroundColor = new BaseColor(225, 225, 225);
            EtiquetaFechaCotizacion.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell ValorFechaCotizacion = new PdfPCell(new Phrase("   " + solicitud.FechaHora, new Font(customfont, 7)));
            ValorFechaCotizacion.Colspan = 2;
            ValorFechaCotizacion.HorizontalAlignment = Element.ALIGN_LEFT;
            ValorFechaCotizacion.VerticalAlignment = Element.ALIGN_MIDDLE;
            ValorFechaCotizacion.FixedHeight = 16f;
            ValorFechaCotizacion.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell EtiquetaSolicitante = new PdfPCell(new Phrase("Solicitante:  ", new Font(customfontbold, 7)));
            EtiquetaSolicitante.Colspan = 1;
            EtiquetaSolicitante.HorizontalAlignment = Element.ALIGN_RIGHT;
            EtiquetaSolicitante.VerticalAlignment = Element.ALIGN_MIDDLE;
            EtiquetaSolicitante.FixedHeight = 16f;
            EtiquetaSolicitante.BackgroundColor = new BaseColor(225, 225, 225);
            EtiquetaSolicitante.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell ValorSolicitante = new PdfPCell(new Phrase("   " + solicitud.Solicitante, new Font(customfont, 7)));
            ValorSolicitante.Colspan = 2;
            ValorSolicitante.HorizontalAlignment = Element.ALIGN_LEFT;
            ValorSolicitante.VerticalAlignment = Element.ALIGN_MIDDLE;
            ValorSolicitante.FixedHeight = 16f;
            ValorSolicitante.BorderColor = new BaseColor(60, 66, 82);

            informacion.AddCell(EtiquetaFechaCotizacion);
            informacion.AddCell(ValorFechaCotizacion);
            informacion.AddCell(EtiquetaSolicitante);
            informacion.AddCell(ValorSolicitante);

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

            //Informacion de la Solicitud       
            PdfPTable datosGenerales = new PdfPTable(6);

            PdfPCell EtiquetaTipo = new PdfPCell(new Phrase("Tipo: ", new Font(customfontbold, 7)));
            EtiquetaTipo.Colspan = 1;
            EtiquetaTipo.HorizontalAlignment = Element.ALIGN_RIGHT;
            EtiquetaTipo.VerticalAlignment = Element.ALIGN_MIDDLE;
            EtiquetaTipo.FixedHeight = 16f;
            EtiquetaTipo.BackgroundColor = new BaseColor(225, 225, 225);
            EtiquetaTipo.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell ValorTipo = new PdfPCell(new Phrase("   " + solicitud.Tipo, new Font(customfont, 7)));
            ValorTipo.Colspan = 5;
            ValorTipo.HorizontalAlignment = Element.ALIGN_LEFT;
            ValorTipo.VerticalAlignment = Element.ALIGN_MIDDLE;
            ValorTipo.FixedHeight = 16f;
            ValorTipo.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell EtiquetaSubtipo = new PdfPCell(new Phrase("Subtipo: ", new Font(customfontbold, 7)));
            EtiquetaSubtipo.Colspan = 1;
            EtiquetaSubtipo.HorizontalAlignment = Element.ALIGN_RIGHT;
            EtiquetaSubtipo.VerticalAlignment = Element.ALIGN_MIDDLE;
            EtiquetaSubtipo.FixedHeight = 16f;
            EtiquetaSubtipo.BackgroundColor = new BaseColor(225, 225, 225);
            EtiquetaSubtipo.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell ValorSubtipo = new PdfPCell(new Phrase("   " + solicitud.Subtipo, new Font(customfont, 7)));
            ValorSubtipo.Colspan = 5;
            ValorSubtipo.HorizontalAlignment = Element.ALIGN_LEFT;
            ValorSubtipo.VerticalAlignment = Element.ALIGN_MIDDLE;
            ValorSubtipo.FixedHeight = 16f;
            ValorSubtipo.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell EtiquetaOP = new PdfPCell(new Phrase("OP: ", new Font(customfontbold, 7)));
            EtiquetaOP.Colspan = 1;
            EtiquetaOP.HorizontalAlignment = Element.ALIGN_RIGHT;
            EtiquetaOP.VerticalAlignment = Element.ALIGN_MIDDLE;
            EtiquetaOP.FixedHeight = 16f;
            EtiquetaOP.BackgroundColor = new BaseColor(225, 225, 225);
            EtiquetaOP.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell ValorOP = new PdfPCell(new Phrase("   " + solicitud.OP, new Font(customfont, 7)));
            ValorOP.Colspan = 2;
            ValorOP.HorizontalAlignment = Element.ALIGN_LEFT;
            ValorOP.VerticalAlignment = Element.ALIGN_MIDDLE;
            ValorOP.FixedHeight = 16f;
            ValorOP.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell EtiquetaMKT = new PdfPCell(new Phrase("MKT: ", new Font(customfontbold, 7)));
            EtiquetaMKT.Colspan = 1;
            EtiquetaMKT.HorizontalAlignment = Element.ALIGN_RIGHT;
            EtiquetaMKT.VerticalAlignment = Element.ALIGN_MIDDLE;
            EtiquetaMKT.FixedHeight = 16f;
            EtiquetaMKT.BackgroundColor = new BaseColor(225, 225, 225);
            EtiquetaMKT.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell ValorMKT = new PdfPCell(new Phrase("   " + solicitud.MKT, new Font(customfont, 7)));
            ValorMKT.Colspan = 2;
            ValorMKT.HorizontalAlignment = Element.ALIGN_LEFT;
            ValorMKT.VerticalAlignment = Element.ALIGN_MIDDLE;
            ValorMKT.FixedHeight = 16f;
            ValorMKT.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell EtiquetaNombre = new PdfPCell(new Phrase("Nombre Campaña: ", new Font(customfontbold, 7)));
            EtiquetaNombre.Colspan = 1;
            EtiquetaNombre.HorizontalAlignment = Element.ALIGN_RIGHT;
            EtiquetaNombre.VerticalAlignment = Element.ALIGN_MIDDLE;
            EtiquetaNombre.FixedHeight = 16f;
            EtiquetaNombre.BackgroundColor = new BaseColor(225, 225, 225);
            EtiquetaNombre.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell ValorNombre = new PdfPCell(new Phrase("   " + solicitud.Nombre, new Font(customfont, 7)));
            ValorNombre.Colspan = 5;
            ValorNombre.HorizontalAlignment = Element.ALIGN_LEFT;
            ValorNombre.VerticalAlignment = Element.ALIGN_MIDDLE;
            ValorNombre.FixedHeight = 16f;
            ValorNombre.BorderColor = new BaseColor(60, 66, 82);

            datosGenerales.AddCell(EtiquetaTipo);
            datosGenerales.AddCell(ValorTipo);

            datosGenerales.AddCell(EtiquetaSubtipo);
            datosGenerales.AddCell(ValorSubtipo);

            datosGenerales.AddCell(EtiquetaOP);
            datosGenerales.AddCell(ValorOP);
            datosGenerales.AddCell(EtiquetaMKT);
            datosGenerales.AddCell(ValorMKT);

            datosGenerales.AddCell(EtiquetaNombre);
            datosGenerales.AddCell(ValorNombre);

            document.Add(datosGenerales);

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

            //Url Externo  
            PdfPTable urlsExterno = new PdfPTable(6);

            PdfPCell EtiquetaCampoUrls = new PdfPCell(new Phrase("URL's de Tipificación", new Font(customfontbold, 9, 0, new BaseColor(255, 255, 255))));
            EtiquetaCampoUrls.Colspan = 6;
            EtiquetaCampoUrls.HorizontalAlignment = Element.ALIGN_CENTER;
            EtiquetaCampoUrls.VerticalAlignment = Element.ALIGN_MIDDLE;
            EtiquetaCampoUrls.FixedHeight = 16f;
            EtiquetaCampoUrls.BackgroundColor = new BaseColor(60, 66, 82);
            EtiquetaCampoUrls.BorderColor = new BaseColor(60, 66, 82);

            urlsExterno.AddCell(EtiquetaCampoUrls);

            foreach (UrlExternosSolicitud urls in ListadoUrlExterno)
            {
                PdfPCell ValorTipoUrlDetalle = new PdfPCell(new Phrase(contadorUrls.ToString(), new Font(customfont, 7)));
                ValorTipoUrlDetalle.Colspan = 1;
                ValorTipoUrlDetalle.HorizontalAlignment = Element.ALIGN_CENTER;
                ValorTipoUrlDetalle.VerticalAlignment = Element.ALIGN_MIDDLE;
                ValorTipoUrlDetalle.FixedHeight = 23f;
                ValorTipoUrlDetalle.BorderColor = new BaseColor(60, 66, 82);

                PdfPCell ValorDetalleUrlDetalle = new PdfPCell(new Phrase(urls.url.ToString(), new Font(customfont, 7)));
                ValorDetalleUrlDetalle.Colspan = 5;
                ValorDetalleUrlDetalle.HorizontalAlignment = Element.ALIGN_LEFT;
                ValorDetalleUrlDetalle.VerticalAlignment = Element.ALIGN_MIDDLE;
                ValorDetalleUrlDetalle.FixedHeight = 23f;
                ValorDetalleUrlDetalle.BorderColor = new BaseColor(60, 66, 82);

                urlsExterno.AddCell(ValorTipoUrlDetalle);
                urlsExterno.AddCell(ValorDetalleUrlDetalle);
            }

            if (ListadoUrlExterno.Count > 0)
            {
                document.Add(urlsExterno);

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
            }

            document.Close();
        }

        //Reporte Solicitud Html y Tags
        public void GenerarPDFHtmlTags(int? id)
        {
            int contadorUrls = 1;

            //Obtener los datos de la Solicitud
            DatosSolicitudCliente solicitud = db.DatosSolicitudCliente(id).First();

            //Adjuntos
            List<AdjuntosSolicitud> ListadoAdjuntos = db.AdjuntosSolicitud(id).ToList();

            //Url Externo 
            List<UrlExternosSolicitud> ListadoUrlExterno = db.UrlExternosSolicitud(id).ToList();

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

            //anio Actual
            var anioActual = System.DateTime.Now.Year.ToString();

            string basePath = ConfigurationManager.AppSettings["RepositorioDocumentos"];
            string rutaArchivos = basePath + "\\GESTION_PPM\\ADJUNTOS_SOLICITUDES\\" + anioActual + '\\' + solicitud.Id.ToString();

            // En caso de que no exista el directorio, crearlo.
            bool directorio = Directory.Exists(rutaArchivos);
            if (!directorio)
                Directory.CreateDirectory(rutaArchivos);

            //ruta del archivo PDF que se va a crear
            string rutaDocumentos = rutaArchivos + "\\Solicitud-" + solicitud.CodigoSolicitud.ToString() + ".pdf";

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
            PdfPCell TituloCotizador = new PdfPCell(new Phrase("SOLICITUD " + solicitud.Tipo.ToUpper() + " N° " + solicitud.CodigoSolicitud + "      ", new Font(customfontbold, 9, 0, new BaseColor(255, 255, 255))));
            TituloCotizador.Colspan = 5;
            TituloCotizador.Border = Rectangle.NO_BORDER;
            TituloCotizador.HorizontalAlignment = Element.ALIGN_RIGHT;
            TituloCotizador.VerticalAlignment = Element.ALIGN_MIDDLE;
            TituloCotizador.BackgroundColor = new BaseColor(60, 66, 82);
            TituloCotizador.FixedHeight = 50f;

            cabecera.AddCell(logocell);
            cabecera.AddCell(TituloCotizador);

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

            //Informacion de la solicitud         
            PdfPTable informacion = new PdfPTable(6);

            PdfPCell EtiquetaFechaCotizacion = new PdfPCell(new Phrase("Fecha y Hora: ", new Font(customfontbold, 7)));
            EtiquetaFechaCotizacion.Colspan = 1;
            EtiquetaFechaCotizacion.HorizontalAlignment = Element.ALIGN_RIGHT;
            EtiquetaFechaCotizacion.VerticalAlignment = Element.ALIGN_MIDDLE;
            EtiquetaFechaCotizacion.FixedHeight = 16f;
            EtiquetaFechaCotizacion.BackgroundColor = new BaseColor(225, 225, 225);
            EtiquetaFechaCotizacion.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell ValorFechaCotizacion = new PdfPCell(new Phrase("   " + solicitud.FechaHora, new Font(customfont, 7)));
            ValorFechaCotizacion.Colspan = 2;
            ValorFechaCotizacion.HorizontalAlignment = Element.ALIGN_LEFT;
            ValorFechaCotizacion.VerticalAlignment = Element.ALIGN_MIDDLE;
            ValorFechaCotizacion.FixedHeight = 16f;
            ValorFechaCotizacion.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell EtiquetaSolicitante = new PdfPCell(new Phrase("Solicitante:  ", new Font(customfontbold, 7)));
            EtiquetaSolicitante.Colspan = 1;
            EtiquetaSolicitante.HorizontalAlignment = Element.ALIGN_RIGHT;
            EtiquetaSolicitante.VerticalAlignment = Element.ALIGN_MIDDLE;
            EtiquetaSolicitante.FixedHeight = 16f;
            EtiquetaSolicitante.BackgroundColor = new BaseColor(225, 225, 225);
            EtiquetaSolicitante.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell ValorSolicitante = new PdfPCell(new Phrase("   " + solicitud.Solicitante, new Font(customfont, 7)));
            ValorSolicitante.Colspan = 2;
            ValorSolicitante.HorizontalAlignment = Element.ALIGN_LEFT;
            ValorSolicitante.VerticalAlignment = Element.ALIGN_MIDDLE;
            ValorSolicitante.FixedHeight = 16f;
            ValorSolicitante.BorderColor = new BaseColor(60, 66, 82);

            informacion.AddCell(EtiquetaFechaCotizacion);
            informacion.AddCell(ValorFechaCotizacion);
            informacion.AddCell(EtiquetaSolicitante);
            informacion.AddCell(ValorSolicitante);

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

            //Informacion de la Solicitud       
            PdfPTable datosGenerales = new PdfPTable(6);

            PdfPCell EtiquetaTipo = new PdfPCell(new Phrase("Tipo: ", new Font(customfontbold, 7)));
            EtiquetaTipo.Colspan = 1;
            EtiquetaTipo.HorizontalAlignment = Element.ALIGN_RIGHT;
            EtiquetaTipo.VerticalAlignment = Element.ALIGN_MIDDLE;
            EtiquetaTipo.FixedHeight = 16f;
            EtiquetaTipo.BackgroundColor = new BaseColor(225, 225, 225);
            EtiquetaTipo.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell ValorTipo = new PdfPCell(new Phrase("   " + solicitud.Tipo, new Font(customfont, 7)));
            ValorTipo.Colspan = 5;
            ValorTipo.HorizontalAlignment = Element.ALIGN_LEFT;
            ValorTipo.VerticalAlignment = Element.ALIGN_MIDDLE;
            ValorTipo.FixedHeight = 16f;
            ValorTipo.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell EtiquetaSubtipo = new PdfPCell(new Phrase("Subtipo: ", new Font(customfontbold, 7)));
            EtiquetaSubtipo.Colspan = 1;
            EtiquetaSubtipo.HorizontalAlignment = Element.ALIGN_RIGHT;
            EtiquetaSubtipo.VerticalAlignment = Element.ALIGN_MIDDLE;
            EtiquetaSubtipo.FixedHeight = 16f;
            EtiquetaSubtipo.BackgroundColor = new BaseColor(225, 225, 225);
            EtiquetaSubtipo.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell ValorSubtipo = new PdfPCell(new Phrase("   " + solicitud.Subtipo, new Font(customfont, 7)));
            ValorSubtipo.Colspan = 5;
            ValorSubtipo.HorizontalAlignment = Element.ALIGN_LEFT;
            ValorSubtipo.VerticalAlignment = Element.ALIGN_MIDDLE;
            ValorSubtipo.FixedHeight = 16f;
            ValorSubtipo.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell EtiquetaOP = new PdfPCell(new Phrase("OP: ", new Font(customfontbold, 7)));
            EtiquetaOP.Colspan = 1;
            EtiquetaOP.HorizontalAlignment = Element.ALIGN_RIGHT;
            EtiquetaOP.VerticalAlignment = Element.ALIGN_MIDDLE;
            EtiquetaOP.FixedHeight = 16f;
            EtiquetaOP.BackgroundColor = new BaseColor(225, 225, 225);
            EtiquetaOP.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell ValorOP = new PdfPCell(new Phrase("   " + solicitud.OP, new Font(customfont, 7)));
            ValorOP.Colspan = 2;
            ValorOP.HorizontalAlignment = Element.ALIGN_LEFT;
            ValorOP.VerticalAlignment = Element.ALIGN_MIDDLE;
            ValorOP.FixedHeight = 16f;
            ValorOP.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell EtiquetaMKT = new PdfPCell(new Phrase("MKT: ", new Font(customfontbold, 7)));
            EtiquetaMKT.Colspan = 1;
            EtiquetaMKT.HorizontalAlignment = Element.ALIGN_RIGHT;
            EtiquetaMKT.VerticalAlignment = Element.ALIGN_MIDDLE;
            EtiquetaMKT.FixedHeight = 16f;
            EtiquetaMKT.BackgroundColor = new BaseColor(225, 225, 225);
            EtiquetaMKT.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell ValorMKT = new PdfPCell(new Phrase("   " + solicitud.MKT, new Font(customfont, 7)));
            ValorMKT.Colspan = 2;
            ValorMKT.HorizontalAlignment = Element.ALIGN_LEFT;
            ValorMKT.VerticalAlignment = Element.ALIGN_MIDDLE;
            ValorMKT.FixedHeight = 16f;
            ValorMKT.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell EtiquetaNombre = new PdfPCell(new Phrase("Nombre Campaña: ", new Font(customfontbold, 7)));
            EtiquetaNombre.Colspan = 1;
            EtiquetaNombre.HorizontalAlignment = Element.ALIGN_RIGHT;
            EtiquetaNombre.VerticalAlignment = Element.ALIGN_MIDDLE;
            EtiquetaNombre.FixedHeight = 16f;
            EtiquetaNombre.BackgroundColor = new BaseColor(225, 225, 225);
            EtiquetaNombre.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell ValorNombre = new PdfPCell(new Phrase("   " + solicitud.Nombre, new Font(customfont, 7)));
            ValorNombre.Colspan = 5;
            ValorNombre.HorizontalAlignment = Element.ALIGN_LEFT;
            ValorNombre.VerticalAlignment = Element.ALIGN_MIDDLE;
            ValorNombre.FixedHeight = 16f;
            ValorNombre.BorderColor = new BaseColor(60, 66, 82);

            datosGenerales.AddCell(EtiquetaTipo);
            datosGenerales.AddCell(ValorTipo);

            datosGenerales.AddCell(EtiquetaSubtipo);
            datosGenerales.AddCell(ValorSubtipo);

            datosGenerales.AddCell(EtiquetaOP);
            datosGenerales.AddCell(ValorOP);
            datosGenerales.AddCell(EtiquetaMKT);
            datosGenerales.AddCell(ValorMKT);

            datosGenerales.AddCell(EtiquetaNombre);
            datosGenerales.AddCell(ValorNombre);

            document.Add(datosGenerales);

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

            //Url Externo  
            PdfPTable urlsExterno = new PdfPTable(6);

            PdfPCell EtiquetaCampoUrls = new PdfPCell(new Phrase("URL's de Tipificación", new Font(customfontbold, 9, 0, new BaseColor(255, 255, 255))));
            EtiquetaCampoUrls.Colspan = 6;
            EtiquetaCampoUrls.HorizontalAlignment = Element.ALIGN_CENTER;
            EtiquetaCampoUrls.VerticalAlignment = Element.ALIGN_MIDDLE;
            EtiquetaCampoUrls.FixedHeight = 16f;
            EtiquetaCampoUrls.BackgroundColor = new BaseColor(60, 66, 82);
            EtiquetaCampoUrls.BorderColor = new BaseColor(60, 66, 82);

            urlsExterno.AddCell(EtiquetaCampoUrls);

            foreach (UrlExternosSolicitud urls in ListadoUrlExterno)
            {
                PdfPCell ValorTipoUrlDetalle = new PdfPCell(new Phrase(contadorUrls.ToString(), new Font(customfont, 7)));
                ValorTipoUrlDetalle.Colspan = 1;
                ValorTipoUrlDetalle.HorizontalAlignment = Element.ALIGN_CENTER;
                ValorTipoUrlDetalle.VerticalAlignment = Element.ALIGN_MIDDLE;
                ValorTipoUrlDetalle.FixedHeight = 23f;
                ValorTipoUrlDetalle.BorderColor = new BaseColor(60, 66, 82);

                PdfPCell ValorDetalleUrlDetalle = new PdfPCell(new Phrase(urls.url.ToString(), new Font(customfont, 7)));
                ValorDetalleUrlDetalle.Colspan = 5;
                ValorDetalleUrlDetalle.HorizontalAlignment = Element.ALIGN_LEFT;
                ValorDetalleUrlDetalle.VerticalAlignment = Element.ALIGN_MIDDLE;
                ValorDetalleUrlDetalle.FixedHeight = 23f;
                ValorDetalleUrlDetalle.BorderColor = new BaseColor(60, 66, 82);

                urlsExterno.AddCell(ValorTipoUrlDetalle);
                urlsExterno.AddCell(ValorDetalleUrlDetalle);
            }

            if (ListadoUrlExterno.Count > 0)
            {
                document.Add(urlsExterno);

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
            }

            //Archivos Adjuntos
            PdfPTable camposAdjuntos = new PdfPTable(6);

            PdfPCell EtiquetaPiezaReferencia = new PdfPCell(new Phrase("Archivos Arjuntos", new Font(customfontbold, 9, 0, new BaseColor(255, 255, 255))));
            EtiquetaPiezaReferencia.Colspan = 6;
            EtiquetaPiezaReferencia.HorizontalAlignment = Element.ALIGN_CENTER;
            EtiquetaPiezaReferencia.VerticalAlignment = Element.ALIGN_MIDDLE;
            EtiquetaPiezaReferencia.FixedHeight = 16f;
            EtiquetaPiezaReferencia.BackgroundColor = new BaseColor(60, 66, 82);
            EtiquetaPiezaReferencia.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell EtiquetaNumero = new PdfPCell(new Phrase("Tipo", new Font(customfontbold, 7)));
            EtiquetaNumero.Colspan = 2;
            EtiquetaNumero.HorizontalAlignment = Element.ALIGN_CENTER;
            EtiquetaNumero.VerticalAlignment = Element.ALIGN_MIDDLE;
            EtiquetaNumero.FixedHeight = 16f;
            EtiquetaNumero.BackgroundColor = new BaseColor(225, 225, 225);
            EtiquetaNumero.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell EtiquetaNombreArchivo = new PdfPCell(new Phrase("Nombre Archivo", new Font(customfontbold, 7)));
            EtiquetaNombreArchivo.Colspan = 4;
            EtiquetaNombreArchivo.HorizontalAlignment = Element.ALIGN_LEFT;
            EtiquetaNombreArchivo.VerticalAlignment = Element.ALIGN_MIDDLE;
            EtiquetaNombreArchivo.FixedHeight = 16f;
            EtiquetaNombreArchivo.BackgroundColor = new BaseColor(225, 225, 225);
            EtiquetaNombreArchivo.BorderColor = new BaseColor(60, 66, 82);

            camposAdjuntos.AddCell(EtiquetaPiezaReferencia);

            camposAdjuntos.AddCell(EtiquetaNumero);
            camposAdjuntos.AddCell(EtiquetaNombreArchivo);

            foreach (AdjuntosSolicitud adjunto in ListadoAdjuntos)
            {
                PdfPCell EtiquetaCampoNumero = new PdfPCell(new Phrase(adjunto.tipo.ToString(), new Font(customfont, 7)));
                EtiquetaCampoNumero.Colspan = 2;
                EtiquetaCampoNumero.HorizontalAlignment = Element.ALIGN_CENTER;
                EtiquetaCampoNumero.VerticalAlignment = Element.ALIGN_MIDDLE;
                EtiquetaCampoNumero.FixedHeight = 23f;
                EtiquetaCampoNumero.BorderColor = new BaseColor(60, 66, 82);

                PdfPCell ValorNombreAdjunto = new PdfPCell(new Phrase(adjunto.nombre_adjunto, new Font(customfont, 7)));
                ValorNombreAdjunto.Colspan = 4;
                ValorNombreAdjunto.HorizontalAlignment = Element.ALIGN_LEFT;
                ValorNombreAdjunto.VerticalAlignment = Element.ALIGN_MIDDLE;
                ValorNombreAdjunto.FixedHeight = 23f;
                ValorNombreAdjunto.BorderColor = new BaseColor(60, 66, 82);

                camposAdjuntos.AddCell(EtiquetaCampoNumero);
                camposAdjuntos.AddCell(ValorNombreAdjunto);

            }

            if (ListadoAdjuntos.Count > 0)
            {
                document.Add(camposAdjuntos);

                //Salto Linea       
                PdfPTable saltoLinea6 = new PdfPTable(4);

                PdfPCell EtiquetaSaltoLinea6 = new PdfPCell(new Phrase(" ", new Font(customfontbold, 9)));
                EtiquetaSaltoLinea6.Colspan = 4;
                EtiquetaSaltoLinea6.Border = Rectangle.NO_BORDER;
                EtiquetaSaltoLinea6.HorizontalAlignment = Element.ALIGN_CENTER;
                EtiquetaSaltoLinea6.VerticalAlignment = Element.ALIGN_MIDDLE;
                EtiquetaSaltoLinea6.FixedHeight = 8f;

                saltoLinea6.AddCell(EtiquetaSaltoLinea6);

                document.Add(saltoLinea6);
            }


            document.Close();
        }

        #endregion

        public JsonResult _GetUsuariosInternosAsignados()
        {
            List<MultiSelectJQueryUi> items = new List<MultiSelectJQueryUi>();
            try
            {
                var user = ViewData["usuario"] = System.Web.HttpContext.Current.Session["usuario"];
                int userID = int.Parse(user.ToString());

                var UsuariosDependientes = OrganigramaEntity.ConsultarHijosOrganigramaPorUsuarioId(1, 1, userID, true).ToList();

                items = UsuariosDependientes.Select(o => new MultiSelectJQueryUi(o.id.Value, o.usuario, o.codigo)).ToList();
                return Json(items, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                return Json(items, JsonRequestBehavior.AllowGet);
            }

            //var UsuariosDependientes = OrganigramaEntity.ConsultarHijosOrganigramaPorUsuarioId(1, 1, userID, true).Select(s => s.id).ToList();


            //var solicitudesUsuarios = SolicitudClienteInternoEntity.ListarUsuariosInternosAsignados();

            //if (UsuariosDependientes.Count > 0)
            //{
            //    // listado = listado.Where(s => idsSolicitudesAsignadas.Contains(s.id_solicitud) || idsSolicitudesUsuariosDependientes.Contains(s.id_solicitud)).ToList();
            //    //solicitudesUsuarios = solicitudesUsuarios.Where(s => UsuariosDependientes.Contains(s.id_usuario)).ToList();
            //    //items = solicitudesUsuarios.Select(o => new MultiSelectJQueryUi(o.id_usuario, o.nombre_usuario + " " + o.apellido_usuario, o.codigo_usuario)).ToList();
            //    items = UsuariosDependientes.Select(o => new MultiSelectJQueryUi(o.id.Value, o.usuario , o.codigo)).ToList();

            //}
            //else
            //{
            //    //items = SolicitudClienteInternoEntity.ListarUsuariosInternosAsignados().Where(s => UsuariosDependientes.Contains(s.id_usuario)).Select(o => new MultiSelectJQueryUi(o.id_usuario, o.nombre_usuario + " " + o.apellido_usuario, o.codigo_usuario)).ToList();
            //    items = UsuarioEntity.ListarUsuariosInternos().Select(o => new MultiSelectJQueryUi(o.Id, o.Nombres_Completos, o.Codigo)).ToList();
            //}

        }

        [HttpPost]
        public ActionResult CreateOrUpdate(int solicitudID, string usuariosAsignados)
        {
            RespuestaTransaccion resultado = new RespuestaTransaccion();
            try
            {
                var usuarios = usuariosAsignados.Split(',').Select(int.Parse).ToList();

                resultado = SolicitudClienteInternoEntity.CrearActualizarAsginacionesSolicitudesUsuarios(solicitudID, usuarios);

                return Json(new { Resultado = resultado }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                return Json(new { Resultado = new RespuestaTransaccion { Estado = false, Respuesta = ex.Message } }, JsonRequestBehavior.AllowGet);
            }
        }

        public JsonResult _GetArchivosAdjuntos(int? id)
        {
            List<TreeViewJQueryUI> items = new List<TreeViewJQueryUI>();
            if (id.HasValue)
            {

                string basePath = ConfigurationManager.AppSettings["RepositorioDocumentos"];
                string rutaArchivos = basePath + "\\GESTION_PPM\\ADJUNTOS_SOLICITUDES";

                bool directorio = Directory.Exists(rutaArchivos);//Directory.Exists(Server.MapPath("~/ArchivosAdjuntosCotizaciones/"));
                // En caso de que no exista el directorio, crearlo.
                if (!directorio)
                    Directory.CreateDirectory(rutaArchivos);//Directory.CreateDirectory(Server.MapPath("~/ArchivosAdjuntosCotizaciones/"));
                items = getFolderNodes(rutaArchivos);//items = getFolderNodes(Server.MapPath("~/ArchivosAdjuntosCotizaciones/"));

                if (id == null)
                {
                    items = new List<TreeViewJQueryUI>();
                }
                else
                {
                    foreach (var item in items)
                    {
                        List<TreeViewJQueryUI> listaFiltrada = new List<TreeViewJQueryUI>();
                        foreach (var clienteFolder in item.children) // Clientes se encuentra en el segundo nivel
                        {

                            if (clienteFolder.text == id.ToString())
                                listaFiltrada.Add(clienteFolder);
                        }
                        item.children = listaFiltrada;

                        item.children = item.children.Where(s => s.children.Any()).ToList();
                    }
                }

                items = items.Where(s => s.children.Any()).ToList(); // Todos los que contengan elementos.

            }
            return Json(items, JsonRequestBehavior.AllowGet);
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

        List<TreeViewJQueryUI> getFolderNodes(string dir)
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

        public ActionResult DescargarArchivoAdjunto(string path)
        {
            byte[] fileBytes = System.IO.File.ReadAllBytes(path);
            string fileName = Path.GetFileName(path);
            return File(fileBytes, System.Net.Mime.MediaTypeNames.Application.Octet, fileName);
        }

        public ActionResult AdjuntarArchivoSolicitud(int? idSolicitud)
        {
            string FileName = "";
            try
            {
                HttpFileCollectionBase files = Request.Files;
                for (int i = 0; i < files.Count; i++)
                {
                    //string path = AppDomain.CurrentDomain.BaseDirectory + "Uploads/";    
                    //string filename = Path.GetFileName(Request.Files[i].FileName);    

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

                    //anio Actual
                    var anioActual = System.DateTime.Now.Year.ToString();

                    string basePath = ConfigurationManager.AppSettings["RepositorioDocumentos"];
                    string rutaArchivos = basePath + "\\GESTION_PPM\\ADJUNTOS_SOLICITUDES\\" + anioActual + '\\' + idSolicitud.ToString();

                    // En caso de que no exista el directorio, crearlo.
                    bool directorio = Directory.Exists(rutaArchivos);
                    if (!directorio)
                        Directory.CreateDirectory(rutaArchivos);

                    path = Path.Combine(rutaArchivos, path);

                    //Guardar el archivo   
                    file.SaveAs(path);

                }
                return Json(new { Resultado = new RespuestaTransaccion { Estado = true, Respuesta = Mensajes.MensajeAdjuntoExitoso }, Archivo = FileName }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                return Json(new { Resultado = new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeAdjuntoFallido + ex.Message } }, JsonRequestBehavior.AllowGet);
            }
        }

        public ActionResult GetSubtipoDependiente(int id)
        {
            var subtipo = CatalogoEntity.ConsultarCatalogoPorPadre(id, "SUBTIPO").ToList();
            return Json(subtipo);
        }

        public ActionResult EliminarArchivoSolicitud(string pathArchivo)
        {
            var nombreArchivo = Path.GetFileName(pathArchivo);
            try
            {
                if (System.IO.File.Exists(pathArchivo))
                {
                    System.IO.File.Delete(pathArchivo);
                }

                DirectoryInfo di = new DirectoryInfo(pathArchivo);
                if (di.Exists)
                    di.Delete(true);

                return Json(new
                {
                    Resultado = new RespuestaTransaccion
                    {
                        Estado = true,
                        Respuesta = string.Format(Mensajes.MensajeAdjuntoEliminadoExitosamente, nombreArchivo)
                    },
                    Archivo = nombreArchivo
                }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                return Json(new { Resultado = new RespuestaTransaccion { Estado = false, Respuesta = string.Format(Mensajes.MensajeAdjuntoEliminadoFallido, nombreArchivo) + ex.Message } }, JsonRequestBehavior.AllowGet);
            }
        }

        public ActionResult _CambiarEstadoSolicitud(int? id)
        {
            SolicitudCliente solicitud = new SolicitudCliente();

            if (id != null)
            {
                solicitud = SolicitudClienteEntity.ConsultarSolicitudCliente(id.Value);
            }

            ViewBag.TituloModal = "Estado Solicitud  ";

            return PartialView(solicitud);
        }

        public ActionResult _AsignarSLA_ANS(int? id)
        {
            SolicitudCliente solicitud = new SolicitudCliente();

            if (id != null)
            {
                solicitud = SolicitudClienteEntity.ConsultarSolicitudCliente(id.Value);
            }

            ViewBag.TituloModal = "Tipo SLA - ANS";

            return PartialView(solicitud);
        }

        [HttpPost]
        public ActionResult CambiarEstadoSolicitud(SolicitudCliente solicitud)
        {
            try
            {
                RespuestaTransaccion resultado = SolicitudClienteEntity.ActualizarEstadoSolicitud(solicitud);

                return Json(new { Resultado = resultado }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                return Json(new { Resultado = new RespuestaTransaccion { Estado = false, Respuesta = ex.Message } }, JsonRequestBehavior.AllowGet);
            }
        }

        [HttpPost]
        public ActionResult CambiarANSCliente(SolicitudCliente solicitud)
        {
            try
            {
                RespuestaTransaccion resultado = SolicitudClienteEntity.ActualizarANSClienteSolicitud(solicitud);

                return Json(new { Resultado = resultado }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                return Json(new { Resultado = new RespuestaTransaccion { Estado = false, Respuesta = ex.Message } }, JsonRequestBehavior.AllowGet);
            }
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
                IGrid<SolicitudClienteInternoInfo> grid = CreateExportableGrid();
                ExcelWorksheet sheet = package.Workbook.Worksheets["Listado"];

                foreach (IGridColumn column in grid.Columns)
                {
                    sheet.Cells[1, col].Value = column.Title;
                    sheet.Column(col++).Width = 18;

                    column.IsEncoded = false;
                }

                foreach (IGridRow<SolicitudClienteInternoInfo> gridRow in grid.Rows)
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
                        rowRange.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Orange);
                    }
                }

                return File(package.GetAsByteArray(), "application/unknown", "ListadoPPM.xlsx");
            }
        }

        private IGrid<SolicitudClienteInternoInfo> CreateExportableGrid()
        {
            IGrid<SolicitudClienteInternoInfo> grid = new Grid<SolicitudClienteInternoInfo>(GetListadoFiltrado());
            grid.ViewContext = new ViewContext { HttpContext = HttpContext };
            grid.Query = Request.QueryString;

            grid.Columns.Add(model => model.fecha_hora_solicitud).Titled("Fecha").Formatted("{0:d}");
            grid.Columns.Add(model => model.fecha_hora_solicitud.HasValue ? model.fecha_hora_solicitud.Value.ToString("HH:mm:ss") : string.Empty).Titled("Hora");
            grid.Columns.Add(model => model.NombresCompletosSolicitante).Titled("Solicitante");
            grid.Columns.Add(model => model.TextoCatalogoTipo).Titled("Tipo de Solicitud");
            grid.Columns.Add(model => model.TextoCatalogoSubTipo).Titled("Subtipo de Solicitud");
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
                "HORA",
                "SOLICITANTE",
                "TIPO DE SOLICITUD",
                "SUBTIPO DE SOLICITUD",
                "MARCA",
                "CANTIDAD",
            };

            var listado = (from item in GetListadoFiltrado()//SolicitudClienteInternoEntity.ListadoSolicitudClienteInterno()
                           select new object[]
                           {
                                            item.fecha_hora_solicitud.HasValue ? item.fecha_hora_solicitud.Value.ToString("yyyy/MM/dd") : string.Empty,
                                            item.fecha_hora_solicitud.HasValue ? item.fecha_hora_solicitud.Value.ToString("HH:mm:ss") : string.Empty,
                                            item.NombresCompletosSolicitante,
                                            item.TextoCatalogoTipo,
                                            item.TextoCatalogoSubTipo,
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
            return File(buffer, "text/csv", $"ListadoPPM.csv");
        }

        public ActionResult DescargarReporteFormatoPDF()
        {
            // Seleccionar las columnas a exportar
            var results = GetListadoFiltrado();//SolicitudClienteInternoEntity.ListadoSolicitudClienteInterno();

            var list = Reportes.SerializeToJSON(results);
            return Content(list, "application/json");
        }
        #endregion

        public ActionResult _Comentario(int? id)
        {
            SolicitudClienteExternoInfo solicitud = SolicitudClienteExternoEntity.ConsultarSolicitudClienteExterno(id.Value);
            ViewBag.TituloModal = "Comentarios - Solicitud " + solicitud.CodigoSolicitud;

            ViewBag.SolicitudID = solicitud.id_solicitud;

            //SolicitudClienteExternoInfo solicitud = SolicitudClienteExternoEntity.ConsultarSolicitudClienteExterno(id.Value);
            List<ComentariosSolicitudInfo> comentarios = SolicitudClienteExternoEntity.ListadoComentarioSolicitud(id);
            return PartialView(comentarios);
        }

        public ActionResult DescargarArchivo(int? id)
        {
            var comentario = SolicitudClienteExternoEntity.ConsultarComentarioSolicitud(id.Value);

            byte[] fileBytes = comentario.ArchivoAdjunto;//System.IO.File.ReadAllBytes(path);
            string fileName = comentario.NombreArchivo;
            return File(fileBytes, System.Net.Mime.MediaTypeNames.Application.Octet, fileName);
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
        [HttpPost]
        public ActionResult Edit(SolicitudCliente solicitud)
        {
            try
            {
                RespuestaTransaccion resultado = SolicitudClienteEntity.ActualizarEstadoSolicitud(solicitud);

                return Json(new { Resultado = resultado }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                return Json(new { Resultado = new RespuestaTransaccion { Estado = false, Respuesta = ex.Message } }, JsonRequestBehavior.AllowGet);
            }
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
                    return Json(new { Resultado = new RespuestaTransaccion { Estado = false, Respuesta = "Límite ("+ maxcomentarios+") de respuestas al comentario excedido." } }, JsonRequestBehavior.AllowGet);
                }
                if (respuesta == "")
                {
                    return Json(new { Resultado = new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeRespuestaRequerida } }, JsonRequestBehavior.AllowGet);
                }
                if (contador > int.Parse(ParametrosSistemaEntity.ConsultarParametros(11).valor))
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

        public ActionResult GuardarComentario(int? id)
        {
            string FileName = "";
            try
            {
                int maxcomentarios = int.Parse(ParametrosSistemaEntity.ConsultarParametros(9).valor);
                int? result = db.ConsultarCantidadComentariosSolictud(id).FirstOrDefault();
                if (result > maxcomentarios - 1)
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

                    return Json(new { Resultado = new RespuestaTransaccion { Estado = true, Respuesta = Mensajes.MensajeAdjuntoExitoso }, Archivo = FileName }, JsonRequestBehavior.AllowGet);
                }
            }
            catch (Exception ex)
            {
                return Json(new { Resultado = new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeAdjuntoFallido + ex.Message } }, JsonRequestBehavior.AllowGet);
            }
        }
    } 
}
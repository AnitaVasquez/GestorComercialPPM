using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Net;
using System.Web.Mvc;
using GestionPPM.Entidades.Metodos;
using GestionPPM.Entidades.Modelo;
using GestionPPM.Repositorios;
using System.Threading.Tasks;
using OfficeOpenXml;
using System.Text;
using Omu.Awem.Helpers;
using OfficeOpenXml.Style;
using System.IO;
using Seguridad.Helper;
using System.Web;
using System.Configuration;
namespace TemplateInicial.Controllers
{
    [Autenticado]
    public class CodigoCotizacionGeneralController : BaseAppController
    {
        // GET: CodigoCotizacionGeneral

        private const string XlsxContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
        private int rolID = Convert.ToInt32(System.Web.HttpContext.Current.Session["rolID"]);
        public ActionResult Index()
        {
            rolID = Convert.ToInt32(System.Web.HttpContext.Current.Session["rolID"]);
            ViewBag.PerfilesUsuario = PerfilesEntity.ListarPerfilesPorRol(rolID);

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

        // Vista Detallada
        [HttpGet]
        public async Task<PartialViewResult> IndexGrid(String search)
        {
            ViewBag.NombreListado = Etiquetas.TituloGridCodigoCotizacion;

            //Controlar permisos
            var user = ViewData["usuario"] = System.Web.HttpContext.Current.Session["usuario"];
            var usuario = int.Parse(user.ToString());
            string nombreControlador = ControllerContext.RouteData.Values["controller"].ToString();
            ViewBag.NombreControlador = nombreControlador;


            ViewBag.AccionesUsuario = ManejoPermisosEntity.ListadoAccionesCatalogoUsuario(usuario, nombreControlador);

            //Obtener Acciones del controlador
            ViewBag.AccionesControlador = GetMetodosControlador(nombreControlador);

            //List<string> NavItems = new List<string>();

            //ReflectedControllerDescriptor controllerDesc = new ReflectedControllerDescriptor(this.GetType());

            //foreach (ActionDescriptor action in controllerDesc.GetCanonicalActions())
            //{
            //    bool validAction = true;

            //    object[] attributes = action.GetCustomAttributes(false);

            //    foreach (object filter in attributes)
            //    {

            //        if (filter is HttpPostAttribute || filter is ChildActionOnlyAttribute)
            //        {
            //            validAction = false;
            //            break;
            //        }
            //    }
            //    if (validAction)
            //        NavItems.Add(action.ActionName);
            //}

            //ViewBag.AccionesControlador = NavItems;

            //Búsqueda

            var listado = CodigoCotizacionEntity.ListarCodigoCotizacion();

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
        public ActionResult _CambiarStatus(int? id)
        {
            AprobacionCotizacion aprobacion = CotizacionEntity.ConsultarCodigoCotizacion(id.Value);
            var codigoCotizacion = CodigoCotizacionEntity.ConsultarCodigoCotizacion(id.Value);
            if (aprobacion.estatus_codigo == 0)
            {
                aprobacion.estatus_codigo = Convert.ToInt32(codigoCotizacion.estatus_codigo.ToString());
            }

            ViewBag.TituloModal = "Cambiar Status del Código de Cotización";

            ViewBag.codigoCotizacion = codigoCotizacion;
            return PartialView(aprobacion);
        }

        [HttpPost]
        public ActionResult EditarStatusCodigoCotizacion(AprobacionCotizacion codigoCotizacion)
        {
            try
            {
                RespuestaTransaccion resultado = CodigoCotizacionEntity.ActualizarStatusCodigoCotizacionClienteExterno(codigoCotizacion.CodigoCotizacionID, codigoCotizacion.estatus_codigo);
                //if (codigoCotizacion.CotizacionID>0 && codigoCotizacion.CodigoCotizacionID>0 && (codigoCotizacion.estatus_codigo==69 || codigoCotizacion.estatus_codigo == 68))
                if (codigoCotizacion.CotizacionID > 0 && codigoCotizacion.CodigoCotizacionID > 0)
                {
                    resultado = CotizacionEntity.AprobarCotizacion(codigoCotizacion);
                }

                return Json(new { Resultado = resultado }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                return Json(new { Resultado = new RespuestaTransaccion { Estado = false, Respuesta = ex.Message } }, JsonRequestBehavior.AllowGet);
            }
        }

        public ActionResult _AdjuntarArchivos(int? id)
        {
            ViewBag.TituloModal = "Adjuntar Archivos a la Solicitud";
            CodigoCotizacion codigoCotizacion = CodigoCotizacionEntity.ConsultarCodigoCotizacion(id.Value);
            return PartialView(codigoCotizacion);
        }

        public ActionResult _VerSolicitudesAdjuntas(int? id)
        {
            ViewBag.TituloModal = "Solicitudes Adjuntas a la Solicitud";

            ViewBag.AdjuntosVacio = "No existen archivos adjuntos.";

            var solicitud = SolicitudClienteInternoEntity.ConsultarSolicitudClienteByCodigoCotizacion(id.Value);
            return PartialView(solicitud);
        }

        // GET: CodigoCotizacion/Details/5
        public ActionResult Details(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            CodigoCotizacion codigoCotizacion = CodigoCotizacionEntity.ConsultarCodigoCotizacion(id.Value);
            if (codigoCotizacion == null)
            {
                return HttpNotFound();
            }

            return PartialView(codigoCotizacion);
        }

        // GET: CodigoCotizacion/Create
        public ActionResult Create()
        {
            //Obtener el codigo de usuario que se logeo
            var usuarioSesion = ViewData["usuario"] = System.Web.HttpContext.Current.Session["usuario"];
            var idUsuario = Convert.ToInt16(usuarioSesion);

            ViewBag.ValidacionRolCreacionSAFI = System.Web.HttpContext.Current.Session["ValidacionRolCreacionSAFI"];

            ViewBag.DatosUsuario = UsuarioEntity.ConsultarUsuario(idUsuario);
            //ViewBag.DatosUsuario = UsuarioEntity.ConsultarUsuario(5);
            return View();
        }

        // POST: CodigoCotizacion/Create
        // To protect from overposting attacks, please enable the specific properties you want to bind to, for 
        // more details see https://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        public ActionResult Create(CodigoCotizacion codigoCotizacion, int? codigoUsuario, List<SublineasNegocioCodigoCotizacionParcial> sublineasNegocio, List<int> idsContactosCotizacion, int? solicitudID)
        {
            try
            {
                if (!CodigoCotizacionEntity.ValidarPagosCodigoCotizacion(codigoCotizacion))
                    return Json(new { Resultado = new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeValidacionPagosCodigoCotizacion } }, JsonRequestBehavior.AllowGet);

                if (!CodigoCotizacionEntity.ValidarPagos1(codigoCotizacion))
                    return Json(new { Resultado = new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeValidacionPago1 } }, JsonRequestBehavior.AllowGet);

                RespuestaTransaccion resultado = CodigoCotizacionEntity.CrearCodigoCotizacion(codigoCotizacion, codigoUsuario.Value, sublineasNegocio, idsContactosCotizacion);

                //Verificar que cuando parta de una solicitud ambos procesos (crear codigo cotizacion y actualizar codigo cotizacion solicitud cliente interno se tiene que hacer en una solo transaccion)
                if (resultado.Estado && solicitudID.HasValue)
                {
                    SolicitudClienteInternoEntity.ActualizarCodigoCotizacionSolicitudClienteInterno(solicitudID.Value, resultado.CotizacionID.Value);
                }

                return Json(new { Resultado = resultado }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                return Json(new { Resultado = new RespuestaTransaccion { Estado = false, Respuesta = ex.Message } }, JsonRequestBehavior.AllowGet);
            }
        }

        // GET: CodigoCotizacion/Edit/5
        public ActionResult Edit(int? id, int? solicitudID)
        {

            CodigoCotizacion codigoCotizacion = new CodigoCotizacion();

            if (solicitudID.HasValue)
            {
                ViewBag.VieneSolicitud = true;

                ViewBag.solicitudID = solicitudID.Value;

                var solicitud = SolicitudClienteInternoEntity.ConsultarSolicitudClienteInterno(solicitudID.Value);

                var sublineaNegocio = CostoSublineaNegocioEntity.ConsultarCostosSublineaNegocioBySublineaNegocio(solicitud.id_subtipo.Value);

                ViewBag.SublineaNegocioCodigoCotizacion = sublineaNegocio;
                ViewBag.TotalSublineaNegocioCodigoCotizacion = sublineaNegocio.Sum(s => s.Valor);

                string nombreProyecto = solicitud.NombreProyectoSolicituCliente + "-" + solicitud.TextoCatalogoTipo + "-" + solicitud.TextoCatalogoSubTipo + "-" + solicitud.TextoCatalogoMarca + "-";
                string descripcionProyecto = solicitud.NombreProyectoSolicituCliente + "-" + solicitud.TextoCatalogoTipo + "-" + solicitud.TextoCatalogoSubTipo + "-" + solicitud.TextoCatalogoMarca + "-" + solicitud.op + "-" + solicitud.mkt;

                int? idCliente = UsuarioEntity.ConsultarUsuario(solicitud.id_solicitante.Value).cliente_asociado;

                codigoCotizacion = CodigoCotizacionEntity.ConsultarCodigoCotizacionValoresDefault(nombreProyecto, descripcionProyecto, idCliente);

                //Obtener el codigo de usuario que se logeo
                var usuarioSesion = ViewData["usuario"] = System.Web.HttpContext.Current.Session["usuario"];
                var idUsuario = Convert.ToInt16(usuarioSesion);

                codigoCotizacion.responsable = idUsuario;//solicitud.id_solicitante; // El solicitante se convierte en el responsable del código de cotización

                ViewBag.DatosUsuario = UsuarioEntity.ConsultarUsuario(idUsuario); //UsuarioEntity.ConsultarUsuario(solicitud.id_solicitante.Value);

                switch (solicitud.CodigoCatalogoTipo)
                {
                    case "TSL-ICLUB":
                    case "TSL-ICARE":
                        codigoCotizacion.estatus_codigo = 68;
                        codigoCotizacion.creacion_safi = 293;
                        codigoCotizacion.etapa_cliente = 183;
                        codigoCotizacion.tipo_intermediario = 26;
                        codigoCotizacion.tipo_cliente = 33;
                        //tipo proyecto no recurrente
                        codigoCotizacion.tipo_proyecto = 61;
                        break;

                    case "TSL-ENVIOS-ICARE":
                    case "TSL-TAGS":
                    case "TSL-HTML5":
                    case "TSL-HTML5-TAGS":
                        codigoCotizacion.estatus_codigo = 68;
                        codigoCotizacion.creacion_safi = 293;
                        codigoCotizacion.etapa_cliente = 183;
                        codigoCotizacion.tipo_intermediario = 26;
                        codigoCotizacion.tipo_cliente = 33;
                        break;

                    default:
                        codigoCotizacion.estatus_codigo = 0;
                        codigoCotizacion.creacion_safi = 0;
                        codigoCotizacion.etapa_cliente = 0;
                        break;
                }

                codigoCotizacion.codigo_cotizacion = "CODIGO DE COTIZACIÓN SIN GENERAR";

                ViewBag.ValidacionRolCreacionSAFI = System.Web.HttpContext.Current.Session["ValidacionRolCreacionSAFI"];

                ViewBag.idsContactosClientesFacturacion = "";

                return View(codigoCotizacion);
            }


            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }

            ViewBag.VieneSolicitud = false;

            codigoCotizacion = CodigoCotizacionEntity.ConsultarCodigoCotizacion(id.Value);
            ViewBag.DatosUsuario = UsuarioEntity.ConsultarUsuario(codigoCotizacion.responsable.Value);

            ViewBag.ValidacionRolCreacionSAFI = System.Web.HttpContext.Current.Session["ValidacionRolCreacionSAFI"];

            var contactosFacturacionCodigoCotizacion = ContactoClienteEntity.ListadIdsContactosCodigoCotizacion(id.Value);
            ViewBag.idsContactosClientesFacturacion = string.Join(",", contactosFacturacionCodigoCotizacion);

            ViewBag.SublineaNegocioCodigoCotizacion = CodigoCotizacionEntity.ListarSublineaNegocioCodigoCotizacion(id.Value);
            ViewBag.TotalSublineaNegocioCodigoCotizacion = CodigoCotizacionEntity.ListarSublineaNegocioCodigoCotizacion(id.Value).Sum(s => s.Valor);

            if (codigoCotizacion == null)
            {
                return HttpNotFound();
            }
            return View(codigoCotizacion);
        }

        // POST: CodigoCotizacion/Edit/5
        // To protect from overposting attacks, please enable the specific properties you want to bind to, for 
        // more details see https://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        public ActionResult Edit(CodigoCotizacion codigoCotizacion, List<SublineaNegocioCodigoCotizacionInfo> sublineasNegocio, List<int> idsContactosCotizacion)
        {
            try
            {
                if (!CodigoCotizacionEntity.ValidarPagosCodigoCotizacion(codigoCotizacion))
                    return Json(new { Resultado = new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeValidacionPagosCodigoCotizacion } }, JsonRequestBehavior.AllowGet);

                if (!CodigoCotizacionEntity.ValidarPagos1(codigoCotizacion))
                    return Json(new { Resultado = new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeValidacionPago1 } }, JsonRequestBehavior.AllowGet);


                RespuestaTransaccion resultado = CodigoCotizacionEntity.ActualizarCodigoCotizacion(codigoCotizacion, sublineasNegocio, idsContactosCotizacion);
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
            RespuestaTransaccion resultado = CodigoCotizacionEntity.EliminarCodigoCotizacion(id);// await db.Cabecera.FindAsync(id);

            //return RedirectToAction("Index");
            return Json(new { Resultado = resultado }, JsonRequestBehavior.AllowGet);

        }

        public JsonResult GetDependientesTipoLineaNegocio(int id)
        {
            var catalogo = CatalogoEntity.ObtenerListadoCatalogos(id).Select(o => new Oitem(o.Value, o.Text));
            return Json(catalogo);
        }


        public JsonResult GetDependientesClienteContactos(int? id)
        {
            var catalogo = ContactoClienteEntity.ListarContactosCliente(id);
            return Json(catalogo);

        }

        public JsonResult _GetContactosFacturacion(int? id)
        {
            List<MultiSelectJQueryUi> items = new List<MultiSelectJQueryUi>();
            if (id.HasValue)
            {
                items = ContactoClienteEntity.ListarContactosFacturacion().Where(s => s.idCliente == id.Value)
.Select(o => new MultiSelectJQueryUi(o.id_contacto, o.nombre_contacto + " " + o.apellido_contacto, o.TextoCatalogoContacto)).ToList();
            }
            return Json(items, JsonRequestBehavior.AllowGet);
        }

        public JsonResult _GetArchivosAdjuntosCodigoCotizacion(int? id)
        {
            List<TreeViewJQueryUI> items = new List<TreeViewJQueryUI>();
            if (id.HasValue)
            {

                string basePath = ConfigurationManager.AppSettings["RepositorioDocumentos"];
                string rutaArchivos = basePath + "\\GESTION_PPM\\ADJUNTOS_CODIGOS_COTIZACION";

                bool directorio = Directory.Exists(rutaArchivos);//Directory.Exists(Server.MapPath("~/ArchivosAdjuntosCotizaciones/"));
                // En caso de que no exista el directorio, crearlo.
                if (!directorio)
                    Directory.CreateDirectory(rutaArchivos);//Directory.CreateDirectory(Server.MapPath("~/ArchivosAdjuntosCotizaciones/"));
                items = getFolderNodes(rutaArchivos);//items = getFolderNodes(Server.MapPath("~/ArchivosAdjuntosCotizaciones/"));

                ClientesInfo cliente = new ClientesInfo();
                CodigoCotizacion codigoCotizacion = CodigoCotizacionEntity.ConsultarCodigoCotizacion(id.Value);
                if (codigoCotizacion != null)
                    cliente = ClienteEntity.ConsultarClienteInformacionCompleta(codigoCotizacion.id_cliente);

                if (cliente == null)
                {
                    items = new List<TreeViewJQueryUI>();
                }
                else
                {
                    List<TreeViewJQueryUI> listaFiltradaClientes = new List<TreeViewJQueryUI>();
                    foreach (var item in items)
                    {
                        foreach (var clienteFolder in item.children) // Clientes se encuentra en el segundo nivel
                        {
                            cliente.razon_social_cliente = cliente.razon_social_cliente.Replace(".", string.Empty);

                            if (clienteFolder.text.Contains(cliente.razon_social_cliente))
                                listaFiltradaClientes.Add(clienteFolder);
                        }
                        item.children = listaFiltradaClientes;
                    }
                }

                //bool flag = false;
                //foreach (var item in items)
                //{
                //    if (item.children.Count == 0) {
                //        item.children = null;
                //        flag = true;
                //    }
                //}

                //if(flag)
                //    items = new List<TreeViewJQueryUI>();

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

        public ActionResult AdjuntarArchivoCodigoCotizacion(int? idCodigoCotizacion)
        {
            string FileName = "";
            try
            {
                ClientesInfo cliente = new ClientesInfo();
                CodigoCotizacion codigoCotizacion = CodigoCotizacionEntity.ConsultarCodigoCotizacion(idCodigoCotizacion.Value);

                if (codigoCotizacion != null)
                    cliente = ClienteEntity.ConsultarClienteInformacionCompleta(codigoCotizacion.id_cliente);

                if (cliente == null)
                {
                    return Json(new { Resultado = new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeAdjuntoFallido + Mensajes.MensajeClienteInexistente } }, JsonRequestBehavior.AllowGet);
                }

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
                    var anioActual = System.DateTime.Now.Year.ToString();

                    string razonSocialCliente = cliente.razon_social_cliente.Replace(".", string.Empty);

                    string basePath = ConfigurationManager.AppSettings["RepositorioDocumentos"];
                    string rutaArchivos = basePath + "\\GESTION_PPM\\ADJUNTOS_CODIGOS_COTIZACION";

                    var almacenFisico = Auxiliares.CrearCarpetasDirectorio(rutaArchivos/*Server.MapPath("~/ArchivosAdjuntosCotizaciones/")*/, new List<string>() { anioActual, razonSocialCliente/*cliente.razon_social_cliente*/ });

                    // Get the complete folder path and store the file inside it.    
                    path = Path.Combine(almacenFisico, path);

                    file.SaveAs(path);

                }
                return Json(new { Resultado = new RespuestaTransaccion { Estado = true, Respuesta = Mensajes.MensajeAdjuntoExitoso }, Archivo = FileName }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                return Json(new { Resultado = new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeAdjuntoFallido + ex.Message } }, JsonRequestBehavior.AllowGet);
            }
        }

        public ActionResult EliminarArchivoCodigoCotizacion(string pathArchivo)
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

        public ActionResult DescargarArchivoAdjunto(string path)
        {
            byte[] fileBytes = System.IO.File.ReadAllBytes(path);
            string fileName = Path.GetFileName(path);
            return File(fileBytes, System.Net.Mime.MediaTypeNames.Application.Octet, fileName);
        }

        public ActionResult DescargarReporteFormatoExcel()
        {
            //Seleccionar las columnas a exportar
            var collection = CodigoCotizacionEntity.ListarCodigoCotizacion();
            var package = new ExcelPackage();

            package = Reportes.ExportarExcel(collection.Cast<object>().ToList(), "Codigo Cotizacion");
            return File(package.GetAsByteArray(), XlsxContentType, "ListadoCodigoCotizacion.xlsx");
        }

        public ActionResult DescargarReporteFormatoPDF()
        {
            // Seleccionar las columnas a exportar
            var results = CodigoCotizacionEntity.ListarCodigoCotizacion();

            var list = Reportes.SerializeToJSON(results);
            return Content(list, "application/json");
        }

        public ActionResult DescargarReporteResumidoFormatoCSV()
        {
            var comlumHeadrs = new string[]
            {
                "Id",
                "Fecha de Cotización",
                "Código de Cotización",
                "Estatus Código",
                "Responsable",
                "Cliente",
                "Nombre del Proyecto",
                "Ejecutivo",
                "Tipo FEE",
                "Tipo Proyecto",
                "Fases",
                "Etapa PTOP",
                "Total"
            };

            var listado = (from item in CodigoCotizacionEntity.ListarCodigoCotizacion()
                           select new object[]
                           {
                                            item.id_codigo_cotizacion,
                                            item.fecha_cotizacion,
                                            $"{item.codigo_cotizacion}",
                                            $"{item.EstatusCodigo}",
                                            $"{item.Responsable}",
                                            $"{item.nombre_comercial_cliente}",
                                            $"{item.nombre_proyecto}",
                                            $"{item.Ejecutivo}",
                                            $"{item.TipoFEE}",
                                            $"{item.TipoProyecto}",
                                            $"{item.EtapaCliente}",
                                            $"{item.TipoEtapaPTOP}",
                                            $"\"{(item.TotalSubLineaNegocio) }\"", //Escaping ","
                           }).ToList();

            // Build the file content
            var employeecsv = new StringBuilder();
            listado.ForEach(line =>
            {
                employeecsv.AppendLine(string.Join("¬", line));
            });

            byte[] buffer = Encoding.Default.GetBytes($"{string.Join("¬", comlumHeadrs)}\r\n{employeecsv.ToString()}");
            return File(buffer, "text/csv", $"CodigoCotizacion.csv");
        }

        public ActionResult DescargarReporteDetalladoFormatoCSV()
        {
            var comlumHeadrs = new string[]
            {
                "Id",
                "Fecha de Cotización",
                "Año",
                "Mes",
                "Código de Cotización",
                "Estatus Código",
                "Responsable",
                "Cliente",
                "Tipo de Cliente",
                "Nombre del Proyecto",
                "Descripción del Proyecto",
                "Ejecutivo",
                "Tipo de Requerimiento",
                "Tipo FEE",
                "Tipo Proyecto",
                "Fases",
                "Etapa PTOP",
                "Aplica Contrato",
                "Pagos Parciales",
                "Pago 1",
                "Pago 2",
                "Pago 3",
                "Pago 4",
                "Total",
                "Facturable",
                "Creación SAFI"
            };

            var listado = (from item in CodigoCotizacionEntity.ListarCodigoCotizacion()
                           select new object[]
                           {
                                            item.id_codigo_cotizacion,
                                            item.fecha_cotizacion,
                                            item.Anio_FechaCotizacion,
                                            item.Mes_FechaCotizacion,
                                            $"{item.codigo_cotizacion}",
                                            $"{item.EstatusCodigo}",
                                            $"{item.Responsable}",
                                            $"{item.nombre_comercial_cliente}",
                                            $"{item.TipoCliente}",
                                            $"{item.nombre_proyecto}",
                                            $"{item.descripcion_proyecto}",
                                            $"{item.Ejecutivo}",
                                            $"{item.TipoRequerido}",
                                            $"{item.TipoFEE}",
                                            $"{item.TipoProyecto}",
                                            $"{item.EtapaCliente}",
                                            $"{item.TipoEtapaPTOP}",
                                            $"{item.AplicaContrato}",
                                            $"{item.forma_pago}",
                                            $"{item.forma_pago_1}",
                                            $"{item.forma_pago_2}",
                                            $"{item.forma_pago_3}",
                                            $"{item.forma_pago_4}",
                                            $"{item.TotalSubLineaNegocio}",
                                            $"{item.Facturable}",
                                            $"\"{(item.CreacionSAFI) }\"", //Escaping ","
                           }).ToList();

            // Build the file content
            var employeecsv = new StringBuilder();
            listado.ForEach(line =>
            {
                employeecsv.AppendLine(string.Join("¬", line));
            });

            byte[] buffer = Encoding.Default.GetBytes($"{string.Join("¬", comlumHeadrs)}\r\n{employeecsv.ToString()}");

            return File(buffer, "text/csv", $"CodigoCotizacion.csv");
        }
        public JsonResult GetSublineaDependientesLineaNegocio(int? id)
        {
            var catalogo = CatalogoEntity.ObtenerListadoCatalogosSublineaNegocio(id.Value, null);
            return Json(catalogo);

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

        public ActionResult GetProductoComercial(int id)
        {
            var subtipo = ProductosGestorEntity.ConsultarProductosGestor(id).ToList();
            return Json(subtipo);
        }


        #region Reportes Personalizados

        public ActionResult DescargarReporteResumidoFormatoExcel()
        {
            var collection = CodigoCotizacionEntity.ListarCodigoCotizacion();

            ExcelPackage package = new ExcelPackage();
            var workSheet = package.Workbook.Worksheets.Add("Listado");
            workSheet.TabColor = System.Drawing.Color.Black;
            workSheet.DefaultRowHeight = 12;

            List<string> columnas = new List<string>() {
                "Id",
                "Fecha de Cotización",
                "Código de Cotización",
                "Estatus Código",
                "Responsable",
                "Cliente",
                "Nombre del Proyecto",
                "Ejecutivo",
                "Tipo FEE",
                "Tipo Proyecto",
                "Fases",
                "Etapa PTOP",
                "Total"};

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
                workSheet.Cells[recordIndex, 1].Value = item.id_codigo_cotizacion;
                workSheet.Cells[recordIndex, 2].Value = item.fecha_cotizacion.Value.ToString("yyyy/MM/dd");
                workSheet.Cells[recordIndex, 3].Value = item.codigo_cotizacion;
                workSheet.Cells[recordIndex, 4].Value = item.EstatusCodigo;
                workSheet.Cells[recordIndex, 5].Value = item.Responsable;
                workSheet.Cells[recordIndex, 6].Value = item.nombre_comercial_cliente;
                workSheet.Cells[recordIndex, 7].Value = item.nombre_proyecto;
                workSheet.Cells[recordIndex, 8].Value = item.Ejecutivo;
                workSheet.Cells[recordIndex, 9].Value = item.TipoFEE;
                workSheet.Cells[recordIndex, 10].Value = item.TipoProyecto;
                workSheet.Cells[recordIndex, 11].Value = item.EtapaCliente;
                workSheet.Cells[recordIndex, 12].Value = item.TipoEtapaPTOP;
                workSheet.Cells[recordIndex, 13].Value = item.TotalSubLineaNegocio;
                workSheet.Cells[recordIndex, 13].Style.Numberformat.Format = "###,##0.00";

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


            return File(package.GetAsByteArray(), XlsxContentType, "ListadoCodigoCotizacion.xlsx");
        }

        public ActionResult DescargarReporteDetalladoFormatoExcel()
        {
            var collection = CodigoCotizacionEntity.ListarCodigoCotizacion();

            //==========================================================================================
            //****************************DEFINICION DE LAS HOJAS*************************************//
            //==========================================================================================
            ExcelPackage ExcelPkg = new ExcelPackage();

            ExcelWorksheet Listado = ExcelPkg.Workbook.Worksheets.Add("Lista Códigos");
            Listado.TabColor = System.Drawing.Color.Black;
            Listado.DefaultRowHeight = 12;

            ExcelWorksheet MigracionCabecera = ExcelPkg.Workbook.Worksheets.Add("Códigos Completos");
            MigracionCabecera.TabColor = System.Drawing.Color.Black;
            MigracionCabecera.DefaultRowHeight = 12;

            ExcelWorksheet MigracionDetalle = ExcelPkg.Workbook.Worksheets.Add("Detalle Códigos");
            MigracionDetalle.TabColor = System.Drawing.Color.Black;
            MigracionDetalle.DefaultRowHeight = 12;


            List<string> columnas = new List<string>()
            {
                "Id",
                "Fecha de Cotización",
                "Año",
                "Mes",
                "Código de Cotización",
                "Estatus Código",
                "Responsable",
                "Cliente",
                "Tipo de Cliente",
                "Nombre del Proyecto",
                "Descripción del Proyecto",
                "Ejecutivo",
                "Tipo de Requerimiento",
                "Tipo FEE",
                "Tipo Proyecto",
                "Fases",
                "Etapa PTOP",
                "Aplica Contrato",
                "Pagos Parciales",
                "Pago 1",
                "Pago 2",
                "Pago 3",
                "Pago 4",
                "Total",
                "Facturable",
                "Creación SAFI"
            };

            Listado.Row(1).Height = 20;
            Listado.Row(1).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            Listado.Row(1).Style.Font.Bold = true;

            int contador = 0;
            for (int i = 1; i <= columnas.Count; i++)
            {
                Listado.Cells[1, i].Value = columnas.ElementAt(contador);
                contador++;
            }

            //Body of table  
            int recordIndex = 2;

            foreach (var item in collection)
            {
                Listado.Cells[recordIndex, 1].Value = item.id_codigo_cotizacion;
                Listado.Cells[recordIndex, 2].Value = item.fecha_cotizacion.Value.ToString("yyyy/MM/dd");
                Listado.Cells[recordIndex, 3].Value = item.Anio_FechaCotizacion;
                Listado.Cells[recordIndex, 4].Value = item.Mes_FechaCotizacion;
                Listado.Cells[recordIndex, 5].Value = item.codigo_cotizacion;
                Listado.Cells[recordIndex, 6].Value = item.EstatusCodigo;
                Listado.Cells[recordIndex, 7].Value = item.Responsable;
                Listado.Cells[recordIndex, 8].Value = item.nombre_comercial_cliente;
                Listado.Cells[recordIndex, 9].Value = item.TipoCliente;
                Listado.Cells[recordIndex, 10].Value = item.nombre_proyecto;
                Listado.Cells[recordIndex, 11].Value = item.descripcion_proyecto;
                Listado.Cells[recordIndex, 12].Value = item.Ejecutivo;
                Listado.Cells[recordIndex, 13].Value = item.TipoRequerido;
                Listado.Cells[recordIndex, 14].Value = item.TipoFEE;
                Listado.Cells[recordIndex, 15].Value = item.TipoProyecto;
                Listado.Cells[recordIndex, 16].Value = item.EtapaCliente;
                Listado.Cells[recordIndex, 17].Value = item.TipoEtapaPTOP;
                Listado.Cells[recordIndex, 18].Value = item.AplicaContrato;
                Listado.Cells[recordIndex, 19].Value = item.forma_pago;
                Listado.Cells[recordIndex, 20].Value = item.forma_pago_1;
                Listado.Cells[recordIndex, 20].Style.Numberformat.Format = "###,##0.00";
                Listado.Cells[recordIndex, 21].Value = item.forma_pago_2;
                Listado.Cells[recordIndex, 21].Style.Numberformat.Format = "###,##0.00";
                Listado.Cells[recordIndex, 22].Value = item.forma_pago_3;
                Listado.Cells[recordIndex, 22].Style.Numberformat.Format = "###,##0.00";
                Listado.Cells[recordIndex, 23].Value = item.forma_pago_4;
                Listado.Cells[recordIndex, 23].Style.Numberformat.Format = "###,##0.00";
                Listado.Cells[recordIndex, 24].Value = item.TotalSubLineaNegocio;
                Listado.Cells[recordIndex, 24].Style.Numberformat.Format = "###,##0.00";
                Listado.Cells[recordIndex, 25].Value = item.Facturable;
                Listado.Cells[recordIndex, 26].Value = item.CreacionSAFI;
                recordIndex++;
            }

            for (int i = 1; i <= columnas.Count; i++)
            {
                Listado.Column(i).AutoFit();
                Listado.Column(i).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                if (i == columnas.Count - 2)
                {
                    Listado.Column(i).Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                }
            }

            var collectionDetalle = CodigoCotizacionEntity.ListarCodigoCotizacion();

            List<string> columnasDetalle = new List<string>()
            {
                "Id",
                "Código Cotización",
                "Fecha Cotización",
                "Responsable",
                "Estatus Código",
                "Cliente",
                "Ejecutivo",
                "Tipo Requerimiento",
                "Tipo Intermediario",
                "Tipo Proyecto",
                "Dimension Proyecto",
                "Aplica Contrato",
                "Pagos_Parciales",
                "Pago 1",
                "Pago 2",
                "Pago 3",
                "Pago 4",
                "Etapa del Cliente",
                "Etapa General",
                "Estatus Detallado",
                "Estatus General",
                "Tipo Producto PtoP",
                "Tipo Plan",
                "Tipo Tarifa",
                "Tipo Migración",
                "Tipo Etapa PtoP",
                "Tipo Subsidio",
                "Nombre del Proyecto",
                "Descripción del Proyecto",
                "Tipo FEE",
                "Creación SAFI",
                "Facturable",
                "Area o Departamento",
                "País",
                "Ciudad",
                "Dirección",
                "Tipo Cliente",
                "Tipo CRM",
                "Referido",
                "Estado",
            };

            MigracionCabecera.Row(1).Height = 20;
            MigracionCabecera.Row(1).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            MigracionCabecera.Row(1).Style.Font.Bold = true;

            contador = 0;
            for (int i = 1; i <= columnasDetalle.Count; i++)
            {
                MigracionCabecera.Cells[1, i].Value = columnasDetalle.ElementAt(contador);
                contador++;
            }

            //Body of table  
            recordIndex = 2;

            foreach (var item in collectionDetalle)
            {
                MigracionCabecera.Cells[recordIndex, 1].Value = item.id_codigo_cotizacion;
                MigracionCabecera.Cells[recordIndex, 2].Value = item.codigo_cotizacion;
                MigracionCabecera.Cells[recordIndex, 3].Value = item.fecha_cotizacion.Value.ToString("yyyy/MM/dd");
                MigracionCabecera.Cells[recordIndex, 4].Value = item.Responsable;
                MigracionCabecera.Cells[recordIndex, 5].Value = item.EstatusCodigo;
                MigracionCabecera.Cells[recordIndex, 6].Value = item.razon_social_cliente;
                MigracionCabecera.Cells[recordIndex, 7].Value = item.Ejecutivo;
                MigracionCabecera.Cells[recordIndex, 8].Value = item.TipoRequerido;
                MigracionCabecera.Cells[recordIndex, 9].Value = item.TipoIntermediario;
                MigracionCabecera.Cells[recordIndex, 10].Value = item.TipoProyecto;
                MigracionCabecera.Cells[recordIndex, 11].Value = item.DimensionProyecto;
                MigracionCabecera.Cells[recordIndex, 12].Value = item.AplicaContrato;
                MigracionCabecera.Cells[recordIndex, 13].Value = item.forma_pago;
                MigracionCabecera.Cells[recordIndex, 14].Value = item.forma_pago_1;
                MigracionCabecera.Cells[recordIndex, 15].Value = item.forma_pago_2;
                MigracionCabecera.Cells[recordIndex, 16].Value = item.forma_pago_3;
                MigracionCabecera.Cells[recordIndex, 17].Value = item.forma_pago_4;
                MigracionCabecera.Cells[recordIndex, 18].Value = item.EtapaCliente;
                MigracionCabecera.Cells[recordIndex, 19].Value = item.EtapaGeneral;
                MigracionCabecera.Cells[recordIndex, 20].Value = item.EstatusDetallado;
                MigracionCabecera.Cells[recordIndex, 21].Value = item.EstatusGeneral;
                MigracionCabecera.Cells[recordIndex, 22].Value = item.TipoProductoPTOP;
                MigracionCabecera.Cells[recordIndex, 23].Value = item.TipoPlan;
                MigracionCabecera.Cells[recordIndex, 24].Value = item.TipoTarifa;
                MigracionCabecera.Cells[recordIndex, 25].Value = item.TipoMigracion;
                MigracionCabecera.Cells[recordIndex, 26].Value = item.TipoEtapaPTOP;
                MigracionCabecera.Cells[recordIndex, 27].Value = item.TipoSubsidio;
                MigracionCabecera.Cells[recordIndex, 28].Value = item.nombre_proyecto;
                MigracionCabecera.Cells[recordIndex, 29].Value = item.descripcion_proyecto;
                MigracionCabecera.Cells[recordIndex, 30].Value = item.TipoFEE;
                MigracionCabecera.Cells[recordIndex, 31].Value = item.CreacionSAFI;
                MigracionCabecera.Cells[recordIndex, 32].Value = item.Facturable;
                MigracionCabecera.Cells[recordIndex, 33].Value = item.AreaDepartamento;
                MigracionCabecera.Cells[recordIndex, 34].Value = item.Pais;
                MigracionCabecera.Cells[recordIndex, 35].Value = item.Ciudad;
                MigracionCabecera.Cells[recordIndex, 36].Value = item.direccion;
                MigracionCabecera.Cells[recordIndex, 37].Value = item.TipoCliente;
                MigracionCabecera.Cells[recordIndex, 38].Value = item.TipoZoho;
                MigracionCabecera.Cells[recordIndex, 39].Value = item.Referido;
                MigracionCabecera.Cells[recordIndex, 40].Value = item.Estado;

                recordIndex++;
            }

            for (int i = 1; i <= columnasDetalle.Count; i++)
            {
                MigracionCabecera.Column(i).AutoFit();
                MigracionCabecera.Column(i).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            }

            var collectionDetalleSublinea = CodigoCotizacionEntity.ListarSublineaCodigoCotizacion();

            List<string> columnasDetalleSublinea = new List<string>()
            {
                "Id",
                "Código Cotización",
                "Sublínea Negocio",
                "Valor"
            };

            MigracionDetalle.Row(1).Height = 20;
            MigracionDetalle.Row(1).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            MigracionDetalle.Row(1).Style.Font.Bold = true;

            contador = 0;
            for (int i = 1; i <= columnasDetalleSublinea.Count; i++)
            {
                MigracionDetalle.Cells[1, i].Value = columnasDetalleSublinea.ElementAt(contador);
                contador++;
            }

            //Body of table  
            recordIndex = 2;

            foreach (var item in collectionDetalleSublinea)
            {
                MigracionDetalle.Cells[recordIndex, 1].Value = item.IdSublineaNegocioCotizacion;
                MigracionDetalle.Cells[recordIndex, 2].Value = item.codigo_cotizacion;
                MigracionDetalle.Cells[recordIndex, 3].Value = item.TextoCatalogoSublineaNegocio;
                MigracionDetalle.Cells[recordIndex, 4].Value = item.Valor;
                MigracionDetalle.Cells[recordIndex, 4].Style.Numberformat.Format = "###,##0.00";

                recordIndex++;
            }

            for (int i = 1; i <= columnasDetalleSublinea.Count; i++)
            {
                MigracionDetalle.Column(i).AutoFit();
                MigracionDetalle.Column(i).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                if (i == columnasDetalleSublinea.Count)
                {
                    MigracionDetalle.Column(i).Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                }
            }

            return File(ExcelPkg.GetAsByteArray(), XlsxContentType, "ListadoCodigoCotizacion.xlsx");
        }
        #endregion


        #region Funcionalidad de carga Masiva
        public ActionResult CargarData()
        {
            List<CargaMasiva> errores = new List<CargaMasiva>();
            RespuestaTransaccion resultado = new RespuestaTransaccion { Estado = false };
            List<CodigoCotizacionExcel> listado = new List<CodigoCotizacionExcel>();
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

                    bool directorio = Directory.Exists(Server.MapPath("~/CargasMasivas/"));

                    // En caso de que no exista el directorio, crearlo.
                    if (!directorio)
                        Directory.CreateDirectory(Server.MapPath("~/CargasMasivas/"));

                    // Get the complete folder path and store the file inside it.    
                    path = Path.Combine(Server.MapPath("~/CargasMasivas/"), path);

                    file.SaveAs(path);

                    errores = VerificarEstructuraExcel(path); // Devuelve los errores en la estructura

                    if (errores.Count == 0)
                    {
                        listado = ToEntidadHojaExcelList(path); // Convierte el documento a un Listado tipo entidad

                        // Validacion de la columna N
                        if (listado == null)
                        {
                            errores.Add(new CargaMasiva { Fila = 0, Columna = 0, Valor = "Ninguno", Error = "Verificar la columna de numeración (N.-) del archivo." });
                            return Json(new { Resultado = new RespuestaTransaccion { Estado = resultado.Estado, Respuesta = Mensajes.MensajeCargaMasivaFallida }, Errores = errores, Archivo = FileName }, JsonRequestBehavior.AllowGet);
                        }

                        //Validacion TIPO FEE y TIPO DE PROYECTO
                        foreach (var item in listado)
                        {
                            if (item.tipo_fee == 149 || item.tipo_fee == 150)
                            {
                                if (item.tipo_proyecto != 60)
                                {
                                    errores.Add(new CargaMasiva { Fila = item.N, Columna = 0, Valor = "Tipo Fee : " + item.tipo_fee.ToString() + " ; Tipo de Proyecto: " + item.tipo_proyecto.ToString() + ".", Error = "El Tipo de FEE no aplica para el tipo de proyecto." });
                                    return Json(new { Resultado = new RespuestaTransaccion { Estado = resultado.Estado, Respuesta = Mensajes.MensajeCargaMasivaFallida }, Errores = errores, Archivo = FileName }, JsonRequestBehavior.AllowGet);
                                }
                            }

                            var registroValidar = new CodigoCotizacion
                            {
                                forma_pago = item.forma_pago,
                                forma_pago_1 = item.forma_pago_1,
                                forma_pago_2 = item.forma_pago_2,
                                forma_pago_3 = item.forma_pago_3,
                                forma_pago_4 = item.forma_pago_4
                            };

                            if (!CodigoCotizacionEntity.ValidarPagosCodigoCotizacion3(registroValidar))
                            {
                                errores.Add(new CargaMasiva { Fila = item.N, Columna = 0, Valor = "Ninguno", Error = "Error en la forma de pago. Fila: " + item.N });
                                return Json(new { Resultado = new RespuestaTransaccion { Estado = resultado.Estado, Respuesta = Mensajes.MensajeCargaMasivaFallida }, Errores = errores, Archivo = FileName }, JsonRequestBehavior.AllowGet);
                            }

                        }

                        //Validar que la columna de Numeración no se encuentre duplicada
                        var numeracionDuplicada = listado.GroupBy(x => x.N).Any(g => g.Count() > 1);

                        if (numeracionDuplicada)
                        {
                            errores.Add(new CargaMasiva { Fila = 0, Columna = 0, Valor = "Ninguno", Error = "Valores de la Columna de numeración 'N.-' no pueden ser duplicados." });
                            return Json(new { Resultado = new RespuestaTransaccion { Estado = resultado.Estado, Respuesta = Mensajes.MensajeCargaMasivaFallida }, Errores = errores, Archivo = FileName }, JsonRequestBehavior.AllowGet);
                        }

                        // Cuando existen contactos no asociados a clientes.
                        if (listado == null)
                        {
                            resultado = new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeValidacionCodigoCotizacionContactosCliente + Mensajes.MensajeCargaMasivaFallida };
                            errores.Add(new CargaMasiva { Fila = 0, Columna = 0, Valor = "Ninguno", Error = "Revisar que todos los clientes se encuentren asocidados a los contactos de la plantilla y que estos sean contactos de facturación." });
                            return Json(new { Resultado = resultado, Errores = errores, Archivo = FileName }, JsonRequestBehavior.AllowGet);
                        }

                        if (listado.Count > 0)
                            resultado = CodigoCotizacionEntity.CrearCodigoCotizacionCargaMasiva(listado); // Crea los Codigos de cotizacion
                        else
                        {
                            resultado = new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeCargaMasivaSinRegistros };
                            errores.Add(new CargaMasiva { Fila = 0, Columna = 0, Valor = "Ninguno", Error = "No se encontraron registros. Revisar la estructura del archivo." });
                        }
                    }
                }

                if (resultado.Estado)
                    return Json(new { Resultado = new RespuestaTransaccion { Estado = resultado.Estado, Respuesta = Mensajes.MensajeCargaMasivaExitosa }, Errores = errores, Archivo = FileName }, JsonRequestBehavior.AllowGet);
                else
                {
                    errores.Add(new CargaMasiva { Fila = 0, Columna = 0, Valor = "Ninguno", Error = resultado.Respuesta });
                    return Json(new { Resultado = new RespuestaTransaccion { Estado = resultado.Estado, Respuesta = resultado.Respuesta + Mensajes.MensajeCargaMasivaFallida }, Errores = errores, Archivo = FileName }, JsonRequestBehavior.AllowGet);
                }

            }
            catch (Exception ex)
            {
                return Json(new { Resultado = new RespuestaTransaccion { Estado = false, Respuesta = ex.Message } }, JsonRequestBehavior.AllowGet);
            }
        }

        private List<CodigoCotizacionExcel> ToEntidadHojaExcelList(string pathDelFicheroExcel)
        {
            List<CodigoCotizacionExcel> listadoCargaMasiva = new List<CodigoCotizacionExcel>();
            try
            {

                var user = ViewData["usuario"] = System.Web.HttpContext.Current.Session["usuario"];
                var usuarioID = Convert.ToInt16(user);
                var datosUsuario = UsuarioEntity.ConsultarUsuario(usuarioID);

                FileInfo existingFile = new FileInfo(pathDelFicheroExcel);

                using (ExcelPackage package = new ExcelPackage(existingFile))
                {
                    // get the first worksheet in the workbook
                    ExcelWorksheet hojaCodigoCotizacion = package.Workbook.Worksheets["CODIGO COTIZACION"];
                    ExcelWorksheet hojaSublineasNegocio = package.Workbook.Worksheets["SUBLINEAS"];
                    ExcelWorksheet hojaContactos = package.Workbook.Worksheets["CONTACTOS FACTURACION"];


                    if (hojaCodigoCotizacion != null)
                    {
                        int colCount = hojaCodigoCotizacion.Dimension.End.Column;  //get Column Count
                        int rowCount = hojaCodigoCotizacion.Dimension.End.Row;     //get row count
                        for (int row = 2; row <= rowCount; row++)
                        {

                            var aplicaContratoBool = Auxiliares.GetValorBooleano((hojaCodigoCotizacion.GetValue(row, 13) ?? "").ToString().Trim());
                            var pagosParcialesoBool = Auxiliares.GetValorBooleano((hojaCodigoCotizacion.GetValue(row, 14) ?? "").ToString().Trim());
                            var facturableBool = Auxiliares.GetValorBooleano((hojaCodigoCotizacion.GetValue(row, 15) ?? "").ToString().Trim());

                            listadoCargaMasiva.Add(new CodigoCotizacionExcel
                            {
                                N = !string.IsNullOrEmpty((hojaCodigoCotizacion.GetValue(row, 1) ?? "").ToString().Trim()) ? int.Parse((hojaCodigoCotizacion.GetValue(row, 1) ?? "").ToString().Trim()) : 0,
                                fecha_cotizacion = DateTime.Now,
                                responsable = usuarioID,//!string.IsNullOrEmpty((hojaCodigoCotizacion.GetValue(row, 2) ?? "").ToString().Trim()) ? int.Parse((hojaCodigoCotizacion.GetValue(row, 2) ?? "").ToString().Trim()) : 0,
                                estatus_codigo = !string.IsNullOrEmpty((hojaCodigoCotizacion.GetValue(row, 2) ?? "").ToString().Trim()) ? int.Parse((hojaCodigoCotizacion.GetValue(row, 2) ?? "").ToString().Trim()) : 0,
                                referido = !string.IsNullOrEmpty((hojaCodigoCotizacion.GetValue(row, 3) ?? "").ToString().Trim()) ? int.Parse((hojaCodigoCotizacion.GetValue(row, 3) ?? "").ToString().Trim()) : 0,
                                tipo_intermediario = !string.IsNullOrEmpty((hojaCodigoCotizacion.GetValue(row, 4) ?? "").ToString().Trim()) ? int.Parse((hojaCodigoCotizacion.GetValue(row, 4) ?? "").ToString().Trim()) : 0,
                                tipo_zoho = !string.IsNullOrEmpty((hojaCodigoCotizacion.GetValue(row, 5) ?? "").ToString().Trim()) ? int.Parse((hojaCodigoCotizacion.GetValue(row, 5) ?? "").ToString().Trim()) : 0,
                                id_cliente = !string.IsNullOrEmpty((hojaCodigoCotizacion.GetValue(row, 6) ?? "").ToString().Trim()) ? int.Parse((hojaCodigoCotizacion.GetValue(row, 6) ?? "").ToString().Trim()) : 0,
                                tipo_cliente = !string.IsNullOrEmpty((hojaCodigoCotizacion.GetValue(row, 7) ?? "").ToString().Trim()) ? int.Parse((hojaCodigoCotizacion.GetValue(row, 7) ?? "").ToString().Trim()) : 0,
                                ejecutivo = !string.IsNullOrEmpty((hojaCodigoCotizacion.GetValue(row, 8) ?? "").ToString().Trim()) ? int.Parse((hojaCodigoCotizacion.GetValue(row, 8) ?? "").ToString().Trim()) : 0,
                                tipo_requerido = !string.IsNullOrEmpty((hojaCodigoCotizacion.GetValue(row, 9) ?? "").ToString().Trim()) ? int.Parse((hojaCodigoCotizacion.GetValue(row, 9) ?? "").ToString().Trim()) : 0,
                                tipo_fee = !string.IsNullOrEmpty((hojaCodigoCotizacion.GetValue(row, 10) ?? "").ToString().Trim()) ? int.Parse((hojaCodigoCotizacion.GetValue(row, 10) ?? "").ToString().Trim()) : 0,
                                tipo_proyecto = !string.IsNullOrEmpty((hojaCodigoCotizacion.GetValue(row, 11) ?? "").ToString().Trim()) ? int.Parse((hojaCodigoCotizacion.GetValue(row, 11) ?? "").ToString().Trim()) : 0,
                                dimension_proyecto = !string.IsNullOrEmpty((hojaCodigoCotizacion.GetValue(row, 12) ?? "").ToString().Trim()) ? int.Parse((hojaCodigoCotizacion.GetValue(row, 12) ?? "").ToString().Trim()) : 0,

                                aplica_contrato = aplicaContratoBool,//!string.IsNullOrEmpty((hojaCodigoCotizacion.GetValue(row, 13) ?? "").ToString().Trim()) ? Convert.ToBoolean((hojaCodigoCotizacion.GetValue(row, 13) ?? "").ToString().Trim()) : false,
                                forma_pago = pagosParcialesoBool,//!string.IsNullOrEmpty((hojaCodigoCotizacion.GetValue(row, 14) ?? "").ToString().Trim()) ? Convert.ToBoolean((hojaCodigoCotizacion.GetValue(row, 14) ?? "").ToString().Trim()) : false,
                                facturable = facturableBool,//!string.IsNullOrEmpty((hojaCodigoCotizacion.GetValue(row, 15) ?? "").ToString().Trim()) ? Convert.ToBoolean((hojaCodigoCotizacion.GetValue(row, 15) ?? "").ToString().Trim()) : false,

                                creacion_safi = !string.IsNullOrEmpty((hojaCodigoCotizacion.GetValue(row, 16) ?? "").ToString().Trim()) ? int.Parse((hojaCodigoCotizacion.GetValue(row, 16) ?? "").ToString().Trim()) : 0,

                                forma_pago_1 = !string.IsNullOrEmpty((hojaCodigoCotizacion.GetValue(row, 17) ?? "").ToString().Trim()) ? decimal.Parse((hojaCodigoCotizacion.GetValue(row, 17) ?? "").ToString().Trim()) : 0,
                                forma_pago_2 = !string.IsNullOrEmpty((hojaCodigoCotizacion.GetValue(row, 18) ?? "").ToString().Trim()) ? decimal.Parse((hojaCodigoCotizacion.GetValue(row, 18) ?? "").ToString().Trim()) : 0,
                                forma_pago_3 = !string.IsNullOrEmpty((hojaCodigoCotizacion.GetValue(row, 19) ?? "").ToString().Trim()) ? decimal.Parse((hojaCodigoCotizacion.GetValue(row, 19) ?? "").ToString().Trim()) : 0,
                                forma_pago_4 = !string.IsNullOrEmpty((hojaCodigoCotizacion.GetValue(row, 20) ?? "").ToString().Trim()) ? decimal.Parse((hojaCodigoCotizacion.GetValue(row, 20) ?? "").ToString().Trim()) : 0,

                                etapa_cliente = !string.IsNullOrEmpty((hojaCodigoCotizacion.GetValue(row, 21) ?? "").ToString().Trim()) ? int.Parse((hojaCodigoCotizacion.GetValue(row, 21) ?? "").ToString().Trim()) : 0, // FASES
                                etapa_general = !string.IsNullOrEmpty((hojaCodigoCotizacion.GetValue(row, 22) ?? "").ToString().Trim()) ? int.Parse((hojaCodigoCotizacion.GetValue(row, 22) ?? "").ToString().Trim()) : 0,
                                estatus_detallado = !string.IsNullOrEmpty((hojaCodigoCotizacion.GetValue(row, 23) ?? "").ToString().Trim()) ? int.Parse((hojaCodigoCotizacion.GetValue(row, 23) ?? "").ToString().Trim()) : 0,
                                estatus_general = !string.IsNullOrEmpty((hojaCodigoCotizacion.GetValue(row, 24) ?? "").ToString().Trim()) ? int.Parse((hojaCodigoCotizacion.GetValue(row, 24) ?? "").ToString().Trim()) : 0,
                                tipo_producto_PtoP = !string.IsNullOrEmpty((hojaCodigoCotizacion.GetValue(row, 25) ?? "").ToString().Trim()) ? int.Parse((hojaCodigoCotizacion.GetValue(row, 25) ?? "").ToString().Trim()) : 0,

                                tipo_plan = !string.IsNullOrEmpty((hojaCodigoCotizacion.GetValue(row, 26) ?? "").ToString().Trim()) ? int.Parse((hojaCodigoCotizacion.GetValue(row, 26) ?? "").ToString().Trim()) : 0,
                                tipo_tarifa = !string.IsNullOrEmpty((hojaCodigoCotizacion.GetValue(row, 27) ?? "").ToString().Trim()) ? int.Parse((hojaCodigoCotizacion.GetValue(row, 27) ?? "").ToString().Trim()) : 0,
                                tipo_migracion = !string.IsNullOrEmpty((hojaCodigoCotizacion.GetValue(row, 28) ?? "").ToString().Trim()) ? int.Parse((hojaCodigoCotizacion.GetValue(row, 28) ?? "").ToString().Trim()) : 0,
                                tipo_subsidio = !string.IsNullOrEmpty((hojaCodigoCotizacion.GetValue(row, 29) ?? "").ToString().Trim()) ? int.Parse((hojaCodigoCotizacion.GetValue(row, 29) ?? "").ToString().Trim()) : 0,
                                tipo_etapa_PtoP = !string.IsNullOrEmpty((hojaCodigoCotizacion.GetValue(row, 30) ?? "").ToString().Trim()) ? int.Parse((hojaCodigoCotizacion.GetValue(row, 30) ?? "").ToString().Trim()) : 0,

                                nombre_proyecto = !string.IsNullOrEmpty((hojaCodigoCotizacion.GetValue(row, 31) ?? "").ToString().Trim()) ? (hojaCodigoCotizacion.GetValue(row, 31) ?? "").ToString().Trim() : string.Empty,
                                descripcion_proyecto = !string.IsNullOrEmpty((hojaCodigoCotizacion.GetValue(row, 32) ?? "").ToString().Trim()) ? (hojaCodigoCotizacion.GetValue(row, 32) ?? "").ToString().Trim() : string.Empty,

                                //Datos de cliente
                                estado_cliente = true,

                                //Datos de usuario
                                area_departamento_usuario = datosUsuario != null ? datosUsuario.area_departamento : string.Empty,
                                pais = datosUsuario != null ? datosUsuario.pais : 0,
                                ciudad = datosUsuario != null ? datosUsuario.ciudad : 0,
                                direccion = datosUsuario != null ? datosUsuario.direccion_usuario : string.Empty,
                            });
                        }
                    }

                    List<SublineaNegocioCodigoCotizacion> listadoSublineas = new List<SublineaNegocioCodigoCotizacion>();
                    if (hojaSublineasNegocio != null)
                    {
                        int colCount = hojaSublineasNegocio.Dimension.End.Column;  //get Column Count
                        int rowCount = hojaSublineasNegocio.Dimension.End.Row;     //get row count
                        for (int row = 2; row <= rowCount; row++)
                        {
                            listadoSublineas.Add(new SublineaNegocioCodigoCotizacion
                            {
                                IdCodigoCotizacion = !string.IsNullOrEmpty((hojaSublineasNegocio.GetValue(row, 1) ?? "").ToString().Trim()) ? int.Parse((hojaSublineasNegocio.GetValue(row, 1) ?? "").ToString().Trim()) : 0, // Columna N.-  --> en hoja de excel
                                CodigoCatalogoSublineaNegocio = !string.IsNullOrEmpty((hojaSublineasNegocio.GetValue(row, 2) ?? "").ToString().Trim()) ? int.Parse((hojaSublineasNegocio.GetValue(row, 2) ?? "").ToString().Trim()) : 0,
                                Valor = !string.IsNullOrEmpty((hojaSublineasNegocio.GetValue(row, 3) ?? "").ToString().Trim()) ? decimal.Parse((hojaSublineasNegocio.GetValue(row, 3) ?? "").ToString().Trim()) : 0,
                            });
                        }
                    }

                    List<ContactosCodigoCotizacion> listadoContactos = new List<ContactosCodigoCotizacion>();
                    if (hojaContactos != null)
                    {
                        int colCount = hojaContactos.Dimension.End.Column;  //get Column Count
                        int rowCount = hojaContactos.Dimension.End.Row;     //get row count
                        for (int row = 2; row <= rowCount; row++)
                        {
                            // Validando que los contactos se encuentren asociados al cliente
                            var idContacto = !string.IsNullOrEmpty((hojaContactos.GetValue(row, 2) ?? "").ToString().Trim()) ? int.Parse((hojaContactos.GetValue(row, 2) ?? "").ToString().Trim()) : 0;
                            foreach (var item in listadoCargaMasiva)
                            {
                                var contactosCliente = ContactoClienteEntity.ListarContactosFacturacion().Where(s => s.idCliente == item.id_cliente).ToList();
                                var validacionContactoEnCliente = contactosCliente.Where(s => s.id_contacto == idContacto).ToList();

                                if (!validacionContactoEnCliente.Any())
                                    return null;
                            }

                            listadoContactos.Add(new ContactosCodigoCotizacion
                            {
                                idCodigoCotizacion = !string.IsNullOrEmpty((hojaContactos.GetValue(row, 1) ?? "").ToString().Trim()) ? int.Parse((hojaContactos.GetValue(row, 1) ?? "").ToString().Trim()) : 0,
                                idContacto = idContacto,
                            });
                        }
                    }

                    var maximoHojaCodigo = listadoCargaMasiva.Select(s => s.N).Max();
                    var maximoHojaSublineas = listadoSublineas.Select(s => s.IdCodigoCotizacion).Max();
                    var maximoHojaContactos = listadoContactos.Select(s => s.idCodigoCotizacion).Max();

                    if (maximoHojaCodigo != maximoHojaSublineas || maximoHojaCodigo != maximoHojaContactos)
                    {
                        return null;
                    }

                    foreach (var item in listadoCargaMasiva)
                    {
                        item.SublineasNegocio = listadoSublineas.Where(s => s.IdCodigoCotizacion == item.N).ToList();
                        item.Contactos = listadoContactos.Where(s => s.idCodigoCotizacion == item.N).ToList();
                    }

                }

                return listadoCargaMasiva;
            }
            catch (Exception ex)
            {
                return new List<CodigoCotizacionExcel>();
            }
        }

        // Devuelve los errores que hayan en la plantilla
        private List<CargaMasiva> VerificarEstructuraExcel(string pathDelFicheroExcel)
        {
            List<CargaMasiva> errores = new List<CargaMasiva>();
            try
            {
                FileInfo existingFile = new FileInfo(pathDelFicheroExcel);

                using (ExcelPackage package = new ExcelPackage(existingFile))
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets.First();

                    if (worksheet != null)
                    {
                        int colCount = worksheet.Dimension.End.Column;  //get Column Count
                        int rowCount = worksheet.Dimension.End.Row;      //get row count - Cabecera
                        for (int row = 2; row <= rowCount; row++)
                        {
                            for (int col = 1; col <= colCount; col++)
                            {
                                var error = string.Empty;
                                string columna = (worksheet.Cells[1, col].Value ?? "").ToString().Trim(); // Nombre de la Columna
                                string valorColumna = (worksheet.Cells[row, col].Value ?? "").ToString().Trim();

                                error = ValidarCamposCargaMasiva(columna, valorColumna);

                                if (!string.IsNullOrEmpty(error))
                                {
                                    errores.Add(new CargaMasiva { Fila = row, Columna = col, Valor = valorColumna, Error = error });
                                }
                            }
                        }
                    }
                }

                return errores;
            }
            catch (Exception ex)
            {
                errores.Add(new CargaMasiva { Fila = 0, Columna = 0, Valor = "No se pudo verificar correctamente la estructura del excel.", Error = ex.Message.ToString() });
                return errores;
            }

        }

        // Valida todas las celdas del excel
        private static string ValidarCamposCargaMasiva(string columna, string valor)
        {
            var excepcion = "El campo {0} con el valor {1} , contiene errores. Error: {2}";
            try
            {
                valor = !string.IsNullOrEmpty(valor) ? valor.Trim() : "";

                bool esNumero;
                int longitudCaracteres = 0;
                var error = "El campo {0} con el valor {1} ,tiene una extensión de caracteres mayor a la permitida. Solo permite {2} caracteres";
                switch (columna)
                {
                    case "N.-":
                        esNumero = int.TryParse(valor, out int valorNumeracion);
                        if (!esNumero)
                            error = "Tipo de dato incorrecto para el campo " + columna + " , con el valor " + valor + " .Solo admite valores numéricos de tipo entero.";
                        else
                            error = null;

                        if (string.IsNullOrEmpty(valor))
                            error = "La numeración es obligatoria.";
                        else
                            error = null;
                        break;

                    case "ID CLIENTE":
                        esNumero = int.TryParse(valor, out int valorIDCliente);
                        if (!esNumero)
                            error = "Tipo de dato incorrecto para el campo " + columna + " , con el valor " + valor + " .Solo admite valores numéricos de tipo entero.";
                        else
                        {
                            var cliente = ClienteEntity.ConsultarCliente(int.Parse(valor));
                            if (cliente == null)
                                error = "No se pudo encontrar el cliente.";
                            else
                                error = null;
                        }
                        break;

                    case "APLICA CONTRATO":
                        error = null;
                        break;

                    case "FORMA DE PAGO":
                        error = null;
                        break;

                    case "CREACION SAFI":
                        error = null;
                        break;

                    case "FACTURABLE":
                        error = null;
                        break;
                    case "PAGOS PARCIALES":
                        error = null;
                        break;

                    case "ESTATUS CODIGO":

                    case "EJECUTIVO":
                    case "TIPO DE REQUERIMIENTO":
                    case "INTERMEDIARIO":
                    case "TIPO PROYECTO":
                    case "DIMENSION DE PROYECTO":
                    case "FASES":
                    case "ETAPA GENERAL":
                    case "ESTATUS DETALLADO":
                    case "ESTATUS GENERAL":
                    case "TIPO PRODUCTO PTOP":
                    case "TIPO PLAN":
                    case "TIPO TARIFA":
                    case "TIPO MIGRACION":
                    case "TIPO ETAPA PTOP":
                    case "TIPO SUBSIDIO":
                    case "TIPO FEE":
                    case "TIPO DE CLIENTE":
                    case "TIPO CRM":
                    case "REFERIDO":
                        esNumero = int.TryParse(valor, out int valorDimensionProyecto);
                        if (!esNumero)
                            error = "Tipo de dato incorrecto para el campo " + columna + " , con el valor " + valor + " .Solo admite valores numéricos de tipo entero.";
                        else
                        {
                            if (!CatalogoEntity.VerificarExistenciaCatalogo(int.Parse(valor)))
                                error = "El valor no es válido para el campo " + columna + " , con el valor " + valor + " ." + Mensajes.MensajeCatalogoNoDisponible;
                            else
                                error = null;
                        }

                        break;
                    case "FORMA PAGO 1":
                    case "FORMA PAGO 2":
                    case "FORMA PAGO 3":
                    case "FORMA PAGO 4":
                        var separadorIncorrecto = valor.IndexOf(".");

                        if (separadorIncorrecto != -1)
                            valor.Replace(".", ",");

                        esNumero = decimal.TryParse(valor, out decimal valorFormaPago);
                        if (!esNumero)
                            error = "Tipo de dato incorrecto para el campo " + columna + " , con el valor " + valor + " .Solo admite valores numéricos de tipo entero.";
                        else
                            error = null;
                        break;
                    case "NOMBRE DE PROYECTO":
                        longitudCaracteres = 250;
                        if (valor.Length > longitudCaracteres)
                        {
                            error = string.Format(error, columna, valor, longitudCaracteres);
                        }
                        else
                        {
                            error = null;
                        }
                        break;
                    case "DESCRIPCION DE PROYECTO":
                        longitudCaracteres = 300;
                        if (valor.Length > longitudCaracteres)
                        {
                            error = string.Format(error, columna, valor, longitudCaracteres);
                        }
                        else
                        {
                            error = null;
                        }
                        break;
                    default:
                        error = "No se pudo encontrar la columna " + columna + ".";
                        break;
                }

                return error;
            }
            catch (Exception ex)
            {
                return string.Format(excepcion, columna, valor, ex.Message.ToString());
            }
        }
        #endregion

    }
}

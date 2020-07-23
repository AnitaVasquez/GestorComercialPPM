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
using System.Net;
using System.Text;
using NonFactors.Mvc.Grid;
using Newtonsoft.Json.Linq;
using System.Web;
using System.Linq.Expressions;

//Extensión para Query string dinámico
using System.Linq.Dynamic;

namespace TemplateInicial.Controllers
{
    [Autenticado]
    public class DocumentoSAFIController : BaseAppController
    {
        // GET: DocumentoSAFI
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
        public async Task<PartialViewResult> IndexGrid(string search, string sort = "", string order = "", long? page = 1)
        {
            var listado = new List<PrefacturaSAFIInfo>();

            //Controlar permisos
            var user = ViewData["usuario"] = System.Web.HttpContext.Current.Session["usuario"];
            var usuario = int.Parse(user.ToString());
            string nombreControlador = ControllerContext.RouteData.Values["controller"].ToString();
            ViewBag.NombreControlador = nombreControlador;

            ViewBag.AccionesUsuario = ManejoPermisosEntity.ListadoAccionesCatalogoUsuario(usuario, nombreControlador);

            //Obtener Acciones del controlador
            ViewBag.AccionesControlador = GetMetodosControlador(nombreControlador);

            //Búsqueda
            ViewBag.NombreListado = Etiquetas.TituloGridDocumentosSAFI;
            page = page > 0 ? page - 1 : page;
            int totalPaginas = 1;
            try
            {
                var query = (HttpContext.Request.Params.Get("QUERY_STRING") ?? "").ToString();

                var dynamicQueryString = GetQueryString(query);
                var whereClause = BuildWhereDynamicClause(dynamicQueryString);

                //Siempre y cuando no haya filtros definidos en el Grid
                if (string.IsNullOrEmpty(whereClause))
                {
                    if (!string.IsNullOrEmpty(sort) && !string.IsNullOrEmpty(order))
                        listado = CotizacionEntity.ListadoGestionPrefacturaSAFI2(page.Value).OrderBy(sort + " " + order).ToList();
                    else
                        listado = CotizacionEntity.ListadoGestionPrefacturaSAFI2(page.Value).ToList();
                }

                search = !string.IsNullOrEmpty(search) ? search.Trim() : "";

                if (!string.IsNullOrEmpty(search))//filter
                {
                    listado = CotizacionEntity.ListadoGestionPrefacturaSAFI2(null, search);//CotizacionEntity.ListadoGestionPrefacturaSAFI();//
                }

                if (!string.IsNullOrEmpty(whereClause) && string.IsNullOrEmpty(search))
                {
                    if (!string.IsNullOrEmpty(sort) && !string.IsNullOrEmpty(order))
                        listado = CotizacionEntity.ListadoGestionPrefacturaSAFI2(null, null, whereClause).OrderBy(sort + " " + order).ToList();
                    else
                        listado = CotizacionEntity.ListadoGestionPrefacturaSAFI2(null, null, whereClause);

                    //totalPaginas = CotizacionEntity.ListadoGestionPrefacturaSAFI2(null, null, whereClause).Count();
                    //listado = CotizacionEntity.ListadoGestionPrefacturaSAFI2(page, null, whereClause);
                }
                else
                {

                    if (string.IsNullOrEmpty(search))
                        totalPaginas = CotizacionEntity.ObtenerTotalRegistrosListadoPrefacturaSAFI();

                }

                ViewBag.TotalPaginas = totalPaginas;
                // Only grid query values will be available here.
                return PartialView("_IndexGrid", await Task.Run(() => listado));
            }
            catch (Exception ex)
            {
                ViewBag.TotalPaginas = totalPaginas;
                // Only grid query values will be available here.
                return PartialView("_IndexGrid", await Task.Run(() => listado));
            }
        }

        private string BuildWhereDynamicClause(Dictionary<string, object> queryString)
        {
            string query = string.Empty;
            try
            {
                List<string> clausulas = new List<string>();

                string contains = "{0} LIKE '%{1}%'";
                string equals = "{0} = '{1}'";
                string NotEquals = "{0} != '{1}'";
                string StartsWith = "{0} LIKE '{1}%'";
                string EndWith = "{0} LIKE '%{1}'";

                string where = string.Empty;

                foreach (KeyValuePair<string, object> item in queryString)
                {
                    string columna = item.Key.Split('-').FirstOrDefault();

                    if (item.Key.Contains("contains"))
                    {
                        //query += string.Format(contains, columna, (item.Value ?? "").ToString());
                        clausulas.Add(string.Format(contains, columna, (item.Value ?? "").ToString().Trim()));
                    }
                    if (item.Key.Contains("equals"))
                    {
                        //query += string.Format(equals, columna, (item.Value ?? "").ToString());
                        clausulas.Add(string.Format(equals, columna, item.Value.ToString().Trim()));
                    }
                    if (item.Key.Contains("not-equals"))
                    {
                        //query += string.Format(NotEquals, columna, (item.Value ?? "").ToString());
                        clausulas.Add(string.Format(NotEquals, columna, (item.Value ?? "").ToString().Trim()));
                    }
                    if (item.Key.Contains("starts-with"))
                    {
                        //query += string.Format(StartsWith, columna, (item.Value ?? "").ToString());
                        clausulas.Add(string.Format(StartsWith, columna, (item.Value ?? "").ToString().Trim()));
                    }
                    if (item.Key.Contains("ends-with"))
                    {
                        //query += string.Format(EndWith, columna, (item.Value ?? "").ToString());
                        clausulas.Add(string.Format(EndWith, columna, (item.Value ?? "").ToString().Trim()));
                    }

                    where = " WHERE ";
                }

                query += where + string.Join(" AND ", clausulas.ToArray());

                return query;
            }
            catch (Exception ex)
            {
                return query;
            }
        }

        private Dictionary<string, object> GetQueryString(string queryString)
        {
            Dictionary<string, object> querystringDic = new Dictionary<string, object>();
            try
            {
                var parsed = HttpUtility.ParseQueryString(queryString);
                querystringDic = parsed.AllKeys.ToDictionary(k => k, k => (object)parsed[k]);

                querystringDic.Remove("_");

                //Parametros ya incluidos en el request del método IndexGrid
                querystringDic.Remove("search");
                querystringDic.Remove("page");
                querystringDic.Remove("order");
                querystringDic.Remove("sort");

                return querystringDic;
            }
            catch (Exception ex)
            {
                return querystringDic;
            }
        }

        public ActionResult _AdjuntarArchivos(int? id)
        {
            List<TreeViewJQueryUI> items = new List<TreeViewJQueryUI>();

            ViewBag.TituloModal = "Adjuntar archivos";
            PrefacturaSAFIInfo modelo = CotizacionEntity.ConsultarPrefacturaSAFI(id.Value);

            string basePath = ConfigurationManager.AppSettings["RepositorioDocumentos"];
            string rutaArchivos = basePath + "\\GESTION_PPM\\ADJUNTOS_PREFACTURA";

            bool directorio = Directory.Exists(rutaArchivos);//Directory.Exists(Server.MapPath("~/ArchivosAdjuntosCotizaciones/"));
                                                             // En caso de que no exista el directorio, crearlo.
            if (!directorio)
                Directory.CreateDirectory(rutaArchivos);//Directory.CreateDirectory(Server.MapPath("~/ArchivosAdjuntosCotizaciones/"));

            PrefacturaSAFIInfo prefactura = CotizacionEntity.ConsultarPrefacturaSAFI(id.Value);

            rutaArchivos = Path.Combine(rutaArchivos, prefactura.numero_prefactura);

            var directorioX = Directory.Exists(rutaArchivos) ? Directory.GetFiles(rutaArchivos, "*.*", SearchOption.TopDirectoryOnly) : null;
            var files = directorioX != null ? Directory.GetFiles(rutaArchivos, "*.*", SearchOption.TopDirectoryOnly).Where(s => s.EndsWith(".pdf")).ToList() : new List<string>();

            items = GetSoloArchivosEnDirectorio(files);

            var archivo = items.FirstOrDefault();

            ViewBag.RutaArchivo = archivo != null ? archivo.desc : string.Empty;

            var info = archivo != null ? new DirectoryInfo(archivo.desc) : null;

            ViewBag.NombreArchivo = info != null ? info.Name : string.Empty;

            return PartialView(modelo);
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

        public ActionResult DescargarArchivoAdjunto(string path)
        {
            byte[] fileBytes = System.IO.File.ReadAllBytes(path);
            string fileName = Path.GetFileName(path);
            return File(fileBytes, System.Net.Mime.MediaTypeNames.Application.Octet, fileName);
        }

        public JsonResult _GetArchivosAdjuntos(int? id)
        {
            List<TreeViewJQueryUI> items = new List<TreeViewJQueryUI>();
            try
            {
                if (id.HasValue)
                {

                    string basePath = ConfigurationManager.AppSettings["RepositorioDocumentos"];
                    string rutaArchivos = basePath + "\\GESTION_PPM\\ADJUNTOS_PREFACTURA";

                    bool directorio = Directory.Exists(rutaArchivos);//Directory.Exists(Server.MapPath("~/ArchivosAdjuntosCotizaciones/"));
                                                                     // En caso de que no exista el directorio, crearlo.
                    if (!directorio)
                        Directory.CreateDirectory(rutaArchivos);//Directory.CreateDirectory(Server.MapPath("~/ArchivosAdjuntosCotizaciones/"));

                    //items = getFolderNodes(rutaArchivos);//items = getFolderNodes(Server.MapPath("~/ArchivosAdjuntosCotizaciones/"));

                    PrefacturaSAFIInfo prefactura = CotizacionEntity.ConsultarPrefacturaSAFI(id.Value);

                    rutaArchivos = Path.Combine(rutaArchivos, prefactura.numero_prefactura);

                    //var files = Directory.GetFiles(rutaArchivos, "*.*", SearchOption.TopDirectoryOnly).Where(s => s.EndsWith(".pdf") && s.Contains(prefactura.numero_prefactura)).ToList();

                    var files = Directory.GetFiles(rutaArchivos, "*.*", SearchOption.TopDirectoryOnly).Where(s => s.EndsWith(".pdf")).ToList();

                    items = GetSoloArchivosEnDirectorio(files);
                }
                return Json(items, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                return Json(items, JsonRequestBehavior.AllowGet);
            }
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

        [HttpPost]
        public ActionResult Anular(int id)
        {
            if (id == 0)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            RespuestaTransaccion resultado = PrefacturasSAFIEntity.EliminarPrefactura(id);

            return Json(new { Resultado = resultado }, JsonRequestBehavior.AllowGet);
        }

        public ActionResult AdjuntarArchivo(int? id)
        {
            string FileName = "";
            try
            {
                var aprobacion_prefactura_ejecutivo = HttpContext.Request.Params.Get("aprobacion_prefactura_ejecutivo");
                dynamic data = JObject.Parse(aprobacion_prefactura_ejecutivo);
                bool aprobacionInicialEjecutivo = Convert.ToBoolean(data.aprobacion_prefactura_ejecutivo);

                var user = ViewData["usuario"] = System.Web.HttpContext.Current.Session["usuario"];
                var usuarioID = Convert.ToInt16(user);

                PrefacturaSAFIInfo prefactura = CotizacionEntity.ConsultarPrefacturaSAFI(id.Value);

                #region Verificar que el presupuesto o prefactura se encuentre en un Acta de Cliente
                if (!prefactura.UtilizadoEnActaCliente.Value)
                    return Json(new { Resultado = new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeValidacionPresupuestoActaCliente } }, JsonRequestBehavior.AllowGet);
                #endregion

                bool flag = true;

                if (aprobacionInicialEjecutivo)
                    //flag = CotizacionEntity.AprobacionInicialPrefactura(id.Value, usuarioID);

                if (!flag)
                {
                    string mensaje = "Error ({0})";
                    ViewBag.Excepcion = string.Format(mensaje, "No se pudo aprobar la prefactura.");
                    return View("~/Views/Error/InternalServerError.cshtml");
                }

                if (prefactura == null)
                    return Json(new { Resultado = new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeAdjuntoFallido + Mensajes.MensajeArchivoNoExiste } }, JsonRequestBehavior.AllowGet);

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

                    var anioActual = DateTime.Now.Year.ToString();
                    string numeroDocumento = prefactura.numero_prefactura;

                    //if(!FileName.Contains(numeroDocumento))
                    //    return Json(new { Resultado = new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeAdjuntoFallido } }, JsonRequestBehavior.AllowGet);

                    string basePath = ConfigurationManager.AppSettings["RepositorioDocumentos"];
                    string rutaArchivos = basePath + "\\GESTION_PPM\\ADJUNTOS_PREFACTURA";

                    //Eliminar archivos y reemplazarlos con el nuevo adjunto
                    var pathPreFactura = Path.Combine(rutaArchivos, numeroDocumento);
                    bool verificarDirectorio = Directory.Exists(pathPreFactura);
                    if (verificarDirectorio)
                    {
                        var archivos = Directory.GetFiles(Path.Combine(rutaArchivos, numeroDocumento), "*.*", SearchOption.TopDirectoryOnly).ToList();

                        foreach (var item in archivos)
                            DeleteFile(item);
                    }

                    var almacenFisico = Auxiliares.CrearCarpetasDirectorio(rutaArchivos/*Server.MapPath("~/ArchivosAdjuntosCotizaciones/")*/, new List<string>() { numeroDocumento/*anioActual*//*, numeroDocumento*//*cliente.razon_social_cliente*/ });

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

        private bool DeleteFile(string pathArchivo)
        {
            try
            {
                if (System.IO.File.Exists(pathArchivo))
                    System.IO.File.Delete(pathArchivo);

                DirectoryInfo di = new DirectoryInfo(pathArchivo);
                if (di.Exists)
                    di.Delete(true);

                return true;
            }
            catch (Exception ex)
            {
                return false;
            }
        }

        public IGrid<PrefacturaSAFIInfo> CreateExportableGrid()
        {
            IGrid<PrefacturaSAFIInfo> grid = new Grid<PrefacturaSAFIInfo>(CotizacionEntity.ListadoPrefacturaSAFI());
            grid.ViewContext = new ViewContext { HttpContext = HttpContext };
            grid.Query = Request.QueryString;

            grid.Columns.Add(model => model.codigo_cotizacion).Titled("Código de Cotización").AppendCss("celda-grande");

            grid.Columns.Add(model => model.AprobacionPrefacturaEjecutivo).Titled("Aprobada").AppendCss("celda-grande");
            grid.Columns.Add(model => model.fecha_aprobacion_prefactura_ejecutivo).Titled("Fecha de Aprobación").Formatted("{0:d}").AppendCss("celda-grande");
            grid.Columns.Add(model => model.PrefacturaConsolidada).Titled("Consolidada").AppendCss("celda-grande");

            grid.Columns.Add(model => model.numero_prefactura).Titled("Número PreFactura").AppendCss("celda-grande");
            grid.Columns.Add(model => model.nombre_comercial_cliente).Titled("Cliente").AppendCss("celda-grande");
            grid.Columns.Add(model => model.TipoDocumento).Titled("Tipo de Documento").AppendCss("celda-mediana");
            grid.Columns.Add(model => model.fecha_prefactura).Titled("Fecha").Formatted("{0:d}").AppendCss("celda-grande");
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
            var comlumHeadrs = new string[]
            {

                "CODIGO DE COTIZACION",
                "NÚMERO PREFACTURA",
                "APROBADA",
                "CONSOLIDADA",
                "CLIENTE",
                "TIPO DE DOCUMENTO",
                "FECHA",
                 "CANTIDAD",
                  "PRECIO UNITARIO",
                   "IVA",
                    "TOTAL",
            };

            var listado = (from item in CotizacionEntity.ListadoPrefacturaSAFI()
                           select new object[]
                           {
                                            item.codigo_cotizacion,
                                            item.numero_prefactura,
                                            item.AprobacionPrefacturaEjecutivo,
                                            item.PrefacturaConsolidada,
                                            item.nombre_comercial_cliente,
                                             item.TipoDocumento,
                                              item.fecha_prefactura.Value.ToString("yyyy-MM-dd"),
                                               item.cantidad,
                                                item.precio_unitario,
                                                 item.iva_pago,
                                                  item.total_pago,
                           }).ToList();

            // Build the file content
            var employeecsv = new System.Text.StringBuilder();
            listado.ForEach(line =>
            {
                employeecsv.AppendLine(string.Join(",", line));
            });

            byte[] buffer = Encoding.Default.GetBytes($"{string.Join(",", comlumHeadrs)}\r\n{employeecsv.ToString()}");
            return File(buffer, "text/csv", $"ListadoPrefacturas.csv");
        }

    }
}
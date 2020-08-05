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
    public class ReversarController : BaseAppController
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
            ViewBag.NombreListado = Etiquetas.TituloGridREversarConsolidacion;
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
            var listado = PrefacturasSAFIEntity.ListadoReversosPrefacturas();

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
        
        public ActionResult _Reversar(int? id)
        {
            ViewBag.TituloModal = "Reversar Consolidación";
            PrefacturaSAFIInfo modelo = CotizacionEntity.ConsultarPrefacturaSAFI(id.Value);
            return PartialView(modelo);
        }

        public ActionResult ReversarConsolidacion(string listadoIDs, bool ajax = false)
        {
            try
            {
                //obtener los ids
                var ids = !string.IsNullOrEmpty(listadoIDs) ? listadoIDs.Split(',').Select(int.Parse).ToList() : new List<int> { int.Parse(listadoIDs) };

                var user = ViewData["usuario"] = System.Web.HttpContext.Current.Session["usuario"];
                var usuarioID = Convert.ToInt16(user);
                bool aprobado = true;


                foreach (var item in ids)
                {
                    //validar que no tenga reveso
                    var reverso = SolicitudDeReversoEntity.ConsultarReversoConsolidacion(Convert.ToInt32(item.ToString()));

                    if (reverso != null && reverso.estado==true)
                    {
                        aprobado = SolicitudDeReversoEntity.ActualizarSolicitudReveso(ids);
                    }
                }        

                if (!aprobado)
                {
                    if (!ajax)
                    {
                        string mensaje = "Error ({0})";
                        ViewBag.Excepcion = string.Format(mensaje, "No se pudo aprobar la prefactura.");
                        return View("~/Views/Error/InternalServerError.cshtml");
                    }
                    else
                        return Json(new { Resultado = new RespuestaTransaccion { Estado = false, Respuesta = "No se pudo realizar la solicitud de reverso" } }, JsonRequestBehavior.AllowGet);
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
                    List<ConsultarPrefacturaSAFIDetalle> detallePrefactura = CotizacionEntity.ConsultarPrefacturaSAFIDetalle(id);

                    //detalle 1
                    int cantidad_1 = 0;
                    string codigo_1 = "";
                    string descripcion_1 = "";
                    decimal precion_unitario_1 = 0;
                    decimal precion_total_1 = 0;

                    //detalle 2
                    int cantidad_2 = 0;
                    string codigo_2 = "";
                    string descripcion_2 = "";
                    decimal precion_unitario_2 = 0;
                    decimal precion_total_2 = 0;

                    //detalle 3
                    int cantidad_3 = 0;
                    string codigo_3 = "";
                    string descripcion_3 = "";
                    decimal precion_unitario_3 = 0;
                    decimal precion_total_3 = 0;

                    //detalle 4
                    int cantidad_4 = 0;
                    string codigo_4 = "";
                    string descripcion_4 = "";
                    decimal precion_unitario_4 = 0;
                    decimal precion_total_4 = 0;

                    //detalle 5
                    int cantidad_5 = 0;
                    string codigo_5 = "";
                    string descripcion_5 = "";
                    decimal precion_unitario_5 = 0;
                    decimal precion_total_5 = 0;

                    //detalle 6
                    int cantidad_6 = 0;
                    string codigo_6 = "";
                    string descripcion_6 = "";
                    decimal precion_unitario_6 = 0;
                    decimal precion_total_6 = 0;

                    //detalle 7
                    int cantidad_7 = 0;
                    string codigo_7 = "";
                    string descripcion_7 = "";
                    decimal precion_unitario_7 = 0;
                    decimal precion_total_7 = 0;

                    //detalle 8
                    int cantidad_8 = 0;
                    string codigo_8 = "";
                    string descripcion_8 = "";
                    decimal precion_unitario_8 = 0;
                    decimal precion_total_8 = 0;

                    //detalle 9
                    int cantidad_9 = 0;
                    string codigo_9 = "";
                    string descripcion_9 = "";
                    decimal precion_unitario_9 = 0;
                    decimal precion_total_9 = 0;

                    //detalle 10
                    int cantidad_10 = 0;
                    string codigo_10 = "";
                    string descripcion_10 = "";
                    decimal precion_unitario_10 = 0;
                    decimal precion_total_10 = 0;

                    //detalle 11
                    int cantidad_11 = 0;
                    string codigo_11 = "";
                    string descripcion_11 = "";
                    decimal precion_unitario_11 = 0;
                    decimal precion_total_11 = 0;

                    //detalle 12
                    int cantidad_12 = 0;
                    string codigo_12 = "";
                    string descripcion_12 = "";
                    decimal precion_unitario_12 = 0;
                    decimal precion_total_12 = 0;

                    //detalle 13
                    int cantidad_13 = 0;
                    string codigo_13 = "";
                    string descripcion_13 = "";
                    decimal precion_unitario_13 = 0;
                    decimal precion_total_13 = 0;

                    //detalle 14
                    int cantidad_14 = 0;
                    string codigo_14 = "";
                    string descripcion_14 = "";
                    decimal precion_unitario_14 = 0;
                    decimal precion_total_14 = 0;

                    //detalle 15
                    int cantidad_15 = 0;
                    string codigo_15 = "";
                    string descripcion_15 = "";
                    decimal precion_unitario_15 = 0;
                    decimal precion_total_15 = 0;

                    int contadorDetalle = 1;

                    //barrido de los detalles
                    if (detallePrefactura.Any())
                    {
                        if (detallePrefactura.Count() > 0)
                        {
                            foreach (var detallado in detallePrefactura)
                            {
                                if (contadorDetalle == 1)
                                {
                                    cantidad_1 = detallado.cantidad.Value;
                                    codigo_1 = detallado.codigo_producto;
                                    descripcion_1 = detallado.nombre_producto;
                                    precion_unitario_1 = detallado.precio_unitario.Value;
                                    precion_total_1 = detallado.subtotal_pago.Value;
                                }

                                if (contadorDetalle == 2)
                                {
                                    cantidad_2 = detallado.cantidad.Value;
                                    codigo_2 = detallado.codigo_producto;
                                    descripcion_2 = detallado.nombre_producto;
                                    precion_unitario_2 = detallado.precio_unitario.Value;
                                    precion_total_2 = detallado.subtotal_pago.Value;
                                }

                                if (contadorDetalle == 3)
                                {
                                    cantidad_3 = detallado.cantidad.Value;
                                    codigo_3 = detallado.codigo_producto;
                                    descripcion_3 = detallado.nombre_producto;
                                    precion_unitario_3 = detallado.precio_unitario.Value;
                                    precion_total_3 = detallado.subtotal_pago.Value;
                                }

                                if (contadorDetalle == 4)
                                {
                                    cantidad_4 = detallado.cantidad.Value;
                                    codigo_4 = detallado.codigo_producto;
                                    descripcion_4 = detallado.nombre_producto;
                                    precion_unitario_4 = detallado.precio_unitario.Value;
                                    precion_total_4 = detallado.subtotal_pago.Value;
                                }

                                if (contadorDetalle == 5)
                                {
                                    cantidad_5 = detallado.cantidad.Value;
                                    codigo_5 = detallado.codigo_producto;
                                    descripcion_5 = detallado.nombre_producto;
                                    precion_unitario_5 = detallado.precio_unitario.Value;
                                    precion_total_5 = detallado.subtotal_pago.Value;
                                }

                                if (contadorDetalle == 6)
                                {
                                    cantidad_6 = detallado.cantidad.Value;
                                    codigo_6 = detallado.codigo_producto;
                                    descripcion_6 = detallado.nombre_producto;
                                    precion_unitario_6 = detallado.precio_unitario.Value;
                                    precion_total_6 = detallado.subtotal_pago.Value;
                                }

                                if (contadorDetalle == 7)
                                {
                                    cantidad_7 = detallado.cantidad.Value;
                                    codigo_7 = detallado.codigo_producto;
                                    descripcion_7 = detallado.nombre_producto;
                                    precion_unitario_7 = detallado.precio_unitario.Value;
                                    precion_total_7 = detallado.subtotal_pago.Value;
                                }

                                if (contadorDetalle == 8)
                                {
                                    cantidad_8 = detallado.cantidad.Value;
                                    codigo_8 = detallado.codigo_producto;
                                    descripcion_8 = detallado.nombre_producto;
                                    precion_unitario_8 = detallado.precio_unitario.Value;
                                    precion_total_8 = detallado.subtotal_pago.Value;
                                }

                                if (contadorDetalle == 9)
                                {
                                    cantidad_9 = detallado.cantidad.Value;
                                    codigo_9 = detallado.codigo_producto;
                                    descripcion_9 = detallado.nombre_producto;
                                    precion_unitario_9 = detallado.precio_unitario.Value;
                                    precion_total_9 = detallado.subtotal_pago.Value;
                                }

                                if (contadorDetalle == 10)
                                {
                                    cantidad_10 = detallado.cantidad.Value;
                                    codigo_10 = detallado.codigo_producto;
                                    descripcion_10 = detallado.nombre_producto;
                                    precion_unitario_10 = detallado.precio_unitario.Value;
                                    precion_total_10 = detallado.subtotal_pago.Value;
                                }

                                if (contadorDetalle == 11)
                                {
                                    cantidad_11 = detallado.cantidad.Value;
                                    codigo_11 = detallado.codigo_producto;
                                    descripcion_11 = detallado.nombre_producto;
                                    precion_unitario_11 = detallado.precio_unitario.Value;
                                    precion_total_11 = detallado.subtotal_pago.Value;
                                }

                                if (contadorDetalle == 12)
                                {
                                    cantidad_12 = detallado.cantidad.Value;
                                    codigo_12 = detallado.codigo_producto;
                                    descripcion_12 = detallado.nombre_producto;
                                    precion_unitario_12 = detallado.precio_unitario.Value;
                                    precion_total_12 = detallado.subtotal_pago.Value;
                                }

                                if (contadorDetalle == 13)
                                {
                                    cantidad_13 = detallado.cantidad.Value;
                                    codigo_13 = detallado.codigo_producto;
                                    descripcion_13 = detallado.nombre_producto;
                                    precion_unitario_13 = detallado.precio_unitario.Value;
                                    precion_total_13 = detallado.subtotal_pago.Value;
                                }

                                if (contadorDetalle == 14)
                                {
                                    cantidad_14 = detallado.cantidad.Value;
                                    codigo_14 = detallado.codigo_producto;
                                    descripcion_14 = detallado.nombre_producto;
                                    precion_unitario_14 = detallado.precio_unitario.Value;
                                    precion_total_14 = detallado.subtotal_pago.Value;
                                }

                                if (contadorDetalle == 15)
                                {
                                    cantidad_15 = detallado.cantidad.Value;
                                    codigo_15 = detallado.codigo_producto;
                                    descripcion_15 = detallado.nombre_producto;
                                    precion_unitario_15 = detallado.precio_unitario.Value;
                                    precion_total_15 = detallado.subtotal_pago.Value;
                                }

                                contadorDetalle += 1;

                            }
                        }

                    }

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

                                //Detalle 1
                                case "detalle_numero":
                                    form.SetField(fieldKey, "1");
                                    break;
                                case "detalle_codigo":
                                    form.SetField(fieldKey, codigo_1);
                                    break;
                                case "detalle_descripcion":
                                    form.SetField(fieldKey, descripcion_1);
                                    break;
                                case "detalle_cantidad":
                                    form.SetField(fieldKey, cantidad_1.ToString());
                                    break;
                                case "detalle_precio_unitario":
                                    form.SetField(fieldKey, precion_unitario_1.ToString());
                                    break;
                                case "detalle_total":
                                    form.SetField(fieldKey, precion_total_1.ToString());
                                    break;

                                //Detalle 2
                                case "detalle_numero_2":
                                    if (detallePrefactura.Count() >= 2)
                                    {
                                        form.SetField(fieldKey, "2");
                                    }
                                    break;
                                case "detalle_codigo_2":
                                    if (detallePrefactura.Count() >= 2)
                                    {
                                        form.SetField(fieldKey, codigo_2);
                                    }
                                    break;
                                case "detalle_descripcion_2":
                                    if (detallePrefactura.Count() >= 2)
                                    {
                                        form.SetField(fieldKey, descripcion_2);
                                    }
                                    break;
                                case "detalle_cantidad_2":
                                    if (detallePrefactura.Count() >= 2)
                                    {
                                        form.SetField(fieldKey, cantidad_2.ToString());
                                    }
                                    break;
                                case "detalle_precio_unitario_2":
                                    if (detallePrefactura.Count() >= 2)
                                    {
                                        form.SetField(fieldKey, precion_unitario_2.ToString());
                                    }
                                    break;
                                case "detalle_total_2":
                                    if (detallePrefactura.Count() >= 2)
                                    {
                                        form.SetField(fieldKey, precion_total_2.ToString());
                                    }
                                    break;

                                //Detalle 3
                                case "detalle_numero_3":
                                    if (detallePrefactura.Count() >= 3)
                                    {
                                        form.SetField(fieldKey, "3");
                                    }
                                    break;
                                case "detalle_codigo_3":
                                    if (detallePrefactura.Count() >= 3)
                                    {
                                        form.SetField(fieldKey, codigo_3);
                                    }
                                    break;
                                case "detalle_descripcion_3":
                                    if (detallePrefactura.Count() >= 3)
                                    {
                                        form.SetField(fieldKey, descripcion_3);
                                    }
                                    break;
                                case "detalle_cantidad_3":
                                    if (detallePrefactura.Count() >= 3)
                                    {
                                        form.SetField(fieldKey, cantidad_3.ToString());
                                    }
                                    break;
                                case "detalle_precio_unitario_3":
                                    if (detallePrefactura.Count() >= 3)
                                    {
                                        form.SetField(fieldKey, precion_unitario_3.ToString());
                                    }
                                    break;
                                case "detalle_total_3":
                                    if (detallePrefactura.Count() >= 3)
                                    {
                                        form.SetField(fieldKey, precion_total_3.ToString());
                                    }
                                    break;

                                //Detalle 4
                                case "detalle_numero_4":
                                    if (detallePrefactura.Count() >= 4)
                                    {
                                        form.SetField(fieldKey, "4");
                                    }
                                    break;
                                case "detalle_codigo_4":
                                    if (detallePrefactura.Count() >= 4)
                                    {
                                        form.SetField(fieldKey, codigo_4);
                                    }
                                    break;
                                case "detalle_descripcion_4":
                                    if (detallePrefactura.Count() >= 4)
                                    {
                                        form.SetField(fieldKey, descripcion_4);
                                    }
                                    break;
                                case "detalle_cantidad_4":
                                    if (detallePrefactura.Count() >= 4)
                                    {
                                        form.SetField(fieldKey, cantidad_4.ToString());
                                    }
                                    break;
                                case "detalle_precio_unitario_4":
                                    if (detallePrefactura.Count() >= 4)
                                    {
                                        form.SetField(fieldKey, precion_unitario_4.ToString());
                                    }
                                    break;
                                case "detalle_total_4":
                                    if (detallePrefactura.Count() >= 4)
                                    {
                                        form.SetField(fieldKey, precion_total_4.ToString());
                                    }
                                    break;

                                //Detalle 5
                                case "detalle_numero_5":
                                    if (detallePrefactura.Count() >= 5)
                                    {
                                        form.SetField(fieldKey, "5");
                                    }
                                    break;
                                case "detalle_codigo_5":
                                    if (detallePrefactura.Count() >= 5)
                                    {
                                        form.SetField(fieldKey, codigo_5);
                                    }
                                    break;
                                case "detalle_descripcion_5":
                                    if (detallePrefactura.Count() >= 5)
                                    {
                                        form.SetField(fieldKey, descripcion_5);
                                    }
                                    break;
                                case "detalle_cantidad_5":
                                    if (detallePrefactura.Count() >= 5)
                                    {
                                        form.SetField(fieldKey, cantidad_5.ToString());
                                    }
                                    break;
                                case "detalle_precio_unitario_5":
                                    if (detallePrefactura.Count() >= 5)
                                    {
                                        form.SetField(fieldKey, precion_unitario_5.ToString());
                                    }
                                    break;
                                case "detalle_total_5":
                                    if (detallePrefactura.Count() >= 5)
                                    {
                                        form.SetField(fieldKey, precion_total_5.ToString());
                                    }
                                    break;

                                //Detalle 6
                                case "detalle_numero_6":
                                    if (detallePrefactura.Count() >= 6)
                                    {
                                        form.SetField(fieldKey, "6");
                                    }
                                    break;
                                case "detalle_codigo_6":
                                    if (detallePrefactura.Count() >= 6)
                                    {
                                        form.SetField(fieldKey, codigo_6);
                                    }
                                    break;
                                case "detalle_descripcion_6":
                                    if (detallePrefactura.Count() >= 6)
                                    {
                                        form.SetField(fieldKey, descripcion_6);
                                    }
                                    break;
                                case "detalle_cantidad_6":
                                    if (detallePrefactura.Count() >= 6)
                                    {
                                        form.SetField(fieldKey, cantidad_6.ToString());
                                    }
                                    break;
                                case "detalle_precio_unitario_6":
                                    if (detallePrefactura.Count() >= 6)
                                    {
                                        form.SetField(fieldKey, precion_unitario_6.ToString());
                                    }
                                    break;
                                case "detalle_total_6":
                                    if (detallePrefactura.Count() >= 6)
                                    {
                                        form.SetField(fieldKey, precion_total_6.ToString());
                                    }
                                    break;

                                //Detalle 7
                                case "detalle_numero_7":
                                    if (detallePrefactura.Count() >= 7)
                                    {
                                        form.SetField(fieldKey, "7");
                                    }
                                    break;
                                case "detalle_codigo_7":
                                    if (detallePrefactura.Count() >= 7)
                                    {
                                        form.SetField(fieldKey, codigo_7);
                                    }
                                    break;
                                case "detalle_descripcion_7":
                                    if (detallePrefactura.Count() >= 7)
                                    {
                                        form.SetField(fieldKey, descripcion_7);
                                    }
                                    break;
                                case "detalle_cantidad_7":
                                    if (detallePrefactura.Count() >= 7)
                                    {
                                        form.SetField(fieldKey, cantidad_7.ToString());
                                    }
                                    break;
                                case "detalle_precio_unitario_7":
                                    if (detallePrefactura.Count() >= 7)
                                    {
                                        form.SetField(fieldKey, precion_unitario_6.ToString());
                                    }
                                    break;
                                case "detalle_total_7":
                                    if (detallePrefactura.Count() >= 7)
                                    {
                                        form.SetField(fieldKey, precion_total_7.ToString());
                                    }
                                    break;

                                //Detalle 8
                                case "detalle_numero_8":
                                    if (detallePrefactura.Count() >= 8)
                                    {
                                        form.SetField(fieldKey, "8");
                                    }
                                    break;
                                case "detalle_codigo_8":
                                    if (detallePrefactura.Count() >= 8)
                                    {
                                        form.SetField(fieldKey, codigo_8);
                                    }
                                    break;
                                case "detalle_descripcion_8":
                                    if (detallePrefactura.Count() >= 8)
                                    {
                                        form.SetField(fieldKey, descripcion_8);
                                    }
                                    break;
                                case "detalle_cantidad_8":
                                    if (detallePrefactura.Count() >= 8)
                                    {
                                        form.SetField(fieldKey, cantidad_8.ToString());
                                    }
                                    break;
                                case "detalle_precio_unitario_8":
                                    if (detallePrefactura.Count() >= 8)
                                    {
                                        form.SetField(fieldKey, precion_unitario_8.ToString());
                                    }
                                    break;
                                case "detalle_total_8":
                                    if (detallePrefactura.Count() >= 8)
                                    {
                                        form.SetField(fieldKey, precion_total_8.ToString());
                                    }
                                    break;

                                //Detalle 9
                                case "detalle_numero_9":
                                    if (detallePrefactura.Count() >= 9)
                                    {
                                        form.SetField(fieldKey, "9");
                                    }
                                    break;
                                case "detalle_codigo_9":
                                    if (detallePrefactura.Count() >= 9)
                                    {
                                        form.SetField(fieldKey, codigo_9);
                                    }
                                    break;
                                case "detalle_descripcion_9":
                                    if (detallePrefactura.Count() >= 9)
                                    {
                                        form.SetField(fieldKey, descripcion_9);
                                    }
                                    break;
                                case "detalle_cantidad_9":
                                    if (detallePrefactura.Count() >= 9)
                                    {
                                        form.SetField(fieldKey, cantidad_9.ToString());
                                    }
                                    break;
                                case "detalle_precio_unitario_9":
                                    if (detallePrefactura.Count() >= 9)
                                    {
                                        form.SetField(fieldKey, precion_unitario_9.ToString());
                                    }
                                    break;
                                case "detalle_total_9":
                                    if (detallePrefactura.Count() >= 9)
                                    {
                                        form.SetField(fieldKey, precion_total_9.ToString());
                                    }
                                    break;

                                //Detalle 10
                                case "detalle_numero_10":
                                    if (detallePrefactura.Count() >= 10)
                                    {
                                        form.SetField(fieldKey, "10");
                                    }
                                    break;
                                case "detalle_codigo_10":
                                    if (detallePrefactura.Count() >= 10)
                                    {
                                        form.SetField(fieldKey, codigo_6);
                                    }
                                    break;
                                case "detalle_descripcion_10":
                                    if (detallePrefactura.Count() >= 10)
                                    {
                                        form.SetField(fieldKey, descripcion_10);
                                    }
                                    break;
                                case "detalle_cantidad_10":
                                    if (detallePrefactura.Count() >= 10)
                                    {
                                        form.SetField(fieldKey, cantidad_10.ToString());
                                    }
                                    break;
                                case "detalle_precio_unitario_10":
                                    if (detallePrefactura.Count() >= 10)
                                    {
                                        form.SetField(fieldKey, precion_unitario_10.ToString());
                                    }
                                    break;
                                case "detalle_total_10":
                                    if (detallePrefactura.Count() >= 10)
                                    {

                                        form.SetField(fieldKey, precion_total_10.ToString());
                                    }
                                    break;

                                //Detalle 11
                                case "detalle_numero_11":
                                    if (detallePrefactura.Count() >= 11)
                                    {
                                        form.SetField(fieldKey, "11");
                                    }
                                    break;
                                case "detalle_codigo_11":
                                    if (detallePrefactura.Count() >= 11)
                                    {
                                        form.SetField(fieldKey, codigo_11);
                                    }
                                    break;
                                case "detalle_descripcion_11":
                                    if (detallePrefactura.Count() >= 11)
                                    {
                                        form.SetField(fieldKey, descripcion_11);
                                    }
                                    break;
                                case "detalle_cantidad_11":
                                    if (detallePrefactura.Count() >= 11)
                                    {
                                        form.SetField(fieldKey, cantidad_11.ToString());
                                    }
                                    break;
                                case "detalle_precio_unitario_11":
                                    if (detallePrefactura.Count() >= 11)
                                    {
                                        form.SetField(fieldKey, precion_unitario_11.ToString());
                                    }
                                    break;
                                case "detalle_total_11":
                                    if (detallePrefactura.Count() >= 11)
                                    {
                                        form.SetField(fieldKey, precion_total_11.ToString());
                                    }
                                    break;

                                //Detalle 12
                                case "detalle_numero_12":
                                    if (detallePrefactura.Count() >= 12)
                                    {
                                        form.SetField(fieldKey, "12");
                                    }
                                    break;
                                case "detalle_codigo_12":
                                    if (detallePrefactura.Count() >= 12)
                                    {
                                        form.SetField(fieldKey, codigo_12);
                                    }
                                    break;
                                case "detalle_descripcion_12":
                                    if (detallePrefactura.Count() >= 12)
                                    {
                                        form.SetField(fieldKey, descripcion_12);
                                    }
                                    break;
                                case "detalle_cantidad_12":
                                    if (detallePrefactura.Count() >= 12)
                                    {
                                        form.SetField(fieldKey, cantidad_12.ToString());
                                    }
                                    break;
                                case "detalle_precio_unitario_12":
                                    if (detallePrefactura.Count() >= 12)
                                    {
                                        form.SetField(fieldKey, precion_unitario_12.ToString());
                                    }
                                    break;
                                case "detalle_total_12":
                                    if (detallePrefactura.Count() >= 12)
                                    {
                                        form.SetField(fieldKey, precion_total_12.ToString());
                                    }
                                    break;

                                //Detalle 13
                                case "detalle_numero_13":
                                    if (detallePrefactura.Count() >= 13)
                                    {
                                        form.SetField(fieldKey, "13");
                                    }
                                    break;
                                case "detalle_codigo_13":
                                    if (detallePrefactura.Count() >= 13)
                                    {
                                        form.SetField(fieldKey, codigo_13);
                                    }
                                    break;
                                case "detalle_descripcion_13":
                                    if (detallePrefactura.Count() >= 13)
                                    {
                                        form.SetField(fieldKey, descripcion_13);
                                    }
                                    break;
                                case "detalle_cantidad_13":
                                    if (detallePrefactura.Count() >= 13)
                                    {
                                        form.SetField(fieldKey, cantidad_13.ToString());
                                    }
                                    break;
                                case "detalle_precio_unitario_13":
                                    if (detallePrefactura.Count() >= 13)
                                    {
                                        form.SetField(fieldKey, precion_unitario_13.ToString());
                                    }
                                    break;
                                case "detalle_total_13":
                                    if (detallePrefactura.Count() >= 13)
                                    {
                                        form.SetField(fieldKey, precion_total_13.ToString());
                                    }
                                    break;

                                //Detalle 14
                                case "detalle_numero_14":
                                    if (detallePrefactura.Count() >= 14)
                                    {
                                        form.SetField(fieldKey, "14");
                                    }
                                    break;
                                case "detalle_codigo_14":
                                    if (detallePrefactura.Count() >= 14)
                                    {
                                        form.SetField(fieldKey, codigo_14);
                                    }
                                    break;
                                case "detalle_descripcion_14":
                                    if (detallePrefactura.Count() >= 14)
                                    {
                                        form.SetField(fieldKey, descripcion_14);
                                    }
                                    break;
                                case "detalle_cantidad_14":
                                    if (detallePrefactura.Count() >= 14)
                                    {
                                        form.SetField(fieldKey, cantidad_14.ToString());
                                    }
                                    break;
                                case "detalle_precio_unitario_14":
                                    if (detallePrefactura.Count() >= 14)
                                    {
                                        form.SetField(fieldKey, precion_unitario_14.ToString());
                                    }
                                    break;
                                case "detalle_total_14":
                                    if (detallePrefactura.Count() >= 14)
                                    {
                                        form.SetField(fieldKey, precion_total_14.ToString());
                                    }
                                    break;

                                //Detalle 15
                                case "detalle_numero_15":
                                    if (detallePrefactura.Count() >= 15)
                                    {
                                        form.SetField(fieldKey, "15");
                                    }
                                    break;
                                case "detalle_codigo_15":
                                    if (detallePrefactura.Count() >= 15)
                                    {
                                        form.SetField(fieldKey, codigo_15);
                                    }
                                    break;
                                case "detalle_descripcion_15":
                                    if (detallePrefactura.Count() >= 15)
                                    {
                                        form.SetField(fieldKey, descripcion_15);
                                    }
                                    break;
                                case "detalle_cantidad_15":
                                    if (detallePrefactura.Count() >= 15)
                                    {
                                        form.SetField(fieldKey, cantidad_15.ToString());
                                    }
                                    break;
                                case "detalle_precio_unitario_15":
                                    if (detallePrefactura.Count() >= 15)
                                    {
                                        form.SetField(fieldKey, precion_unitario_15.ToString());
                                    }
                                    break;
                                case "detalle_total_15":
                                    if (detallePrefactura.Count() >= 15)
                                    {
                                        form.SetField(fieldKey, precion_total_15.ToString());
                                    }
                                    break;

                                //Final
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
                                    form.SetField(fieldKey, prefactura.subtotal_pago.ToString());
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

        [HttpGet]
        public ActionResult DescargarReporteFormatoExcel()
        {
            // Using EPPlus from nuget
            using (ExcelPackage package = new ExcelPackage())
            {
                Int32 row = 2;
                Int32 col = 1;

                package.Workbook.Worksheets.Add("Data");
                IGrid<PrefacturaSAFIInfo> grid = CreateExportableGrid();
                ExcelWorksheet sheet = package.Workbook.Worksheets["Data"];

                foreach (IGridColumn column in grid.Columns)
                {
                    sheet.Cells[1, col].Value = column.Title;
                    sheet.Column(col++).Width = 18;

                    column.IsEncoded = false;
                }

                foreach (IGridRow<PrefacturaSAFIInfo> gridRow in grid.Rows)
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

                return File(package.GetAsByteArray(), "application/unknown", "ListadoReversosPresupuestos.xlsx");
            }
        }

        public IGrid<PrefacturaSAFIInfo> CreateExportableGrid()
        {
            //Controlar permisos
            var user = ViewData["usuario"] = System.Web.HttpContext.Current.Session["usuario"];
            var usuario = int.Parse(user.ToString());

            IGrid<PrefacturaSAFIInfo> grid = new Grid<PrefacturaSAFIInfo>(PrefacturasSAFIEntity.ListadoReversosPrefacturas());
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

            var listado = (from item in PrefacturasSAFIEntity.ListadoReversosPrefacturas()
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
            return File(buffer, "text/csv", $"ListadoReversosPresupuestos.csv");
        }

        public ActionResult DescargarReporteFormatoPDF()
        {
            //Controlar permisos
            var user = ViewData["usuario"] = System.Web.HttpContext.Current.Session["usuario"];
            var usuario = int.Parse(user.ToString());

            // Seleccionar las columnas a exportar
            var results = PrefacturasSAFIEntity.ListadoReversosPrefacturas();

            var list = Reportes.SerializeToJSON(results);
            return Content(list, "application/json");
        }
    }
}
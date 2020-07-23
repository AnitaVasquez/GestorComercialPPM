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
    [Autenticado]
    public class ProductosGestorController : BaseAppController
    {
        private const string XlsxContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

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
            ViewBag.NombreListado = Etiquetas.TituloGridProductoGestor;

            //Controlar permisos
            var user = ViewData["usuario"] = System.Web.HttpContext.Current.Session["usuario"];
            var usuario = int.Parse(user.ToString());
            string nombreControlador = ControllerContext.RouteData.Values["controller"].ToString();
            ViewBag.NombreControlador = nombreControlador;

            ViewBag.AccionesUsuario = ManejoPermisosEntity.ListadoAccionesCatalogoUsuario(usuario, nombreControlador);

            //Obtener Acciones del controlador
            ViewBag.AccionesControlador = GetMetodosControlador(nombreControlador);

            var listado = ProductosGestorEntity.ListarProductosGestor();

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
            ProductosGestorCE productoCE = new ProductosGestorCE();
            try
            {
                ViewBag.TituloModal = Etiquetas.TituloPanelCreacionProductoGestor;

                if (id.HasValue)
                {
                    //Producto del Gestor
                    ProdutosGestor producto = ProductosGestorEntity.ConsultarProducto(id.Value);

                    //Producto Gestor General
                    var productoGeneral = ProductosGeneralGestorEntity.ConsultarProductoGeneral(producto.id_producto_general);

                    //Sublinea de Producto
                    var sublinea = SublineaNegocioEntity.ConsultarSublinea(productoGeneral.id_sublinea_negocio);

                    productoCE.nombre = producto.nombre;
                    productoCE.descripcon = producto.descripcon;
                    productoCE.estado = producto.estado.Value;
                    productoCE.id_producto_general = producto.id_producto_general;
                    productoCE.id_producto_gestor = producto.id_producto_gestor;
                    productoCE.id_sublinea_negocio = productoGeneral.id_sublinea_negocio;
                    productoCE.id_linea_negocio = sublinea.id_linea_negocio;
                    productoCE.id_tipo_producto = 907;

                    ViewBag.TipoProducto = 907;
                }

                return PartialView(productoCE);
            }
            catch (Exception ex)
            {
                string mensaje = "Un error ocurrió. {0}";
                ViewBag.Excepcion = string.Format(mensaje, ex.Message.ToString());
                return PartialView("~/Views/Error/_InternalServerError.cshtml");
            }
        }

        [HttpPost]
        public ActionResult CreateOrUpdate(ProductosGestorCE productoCE)
        {
            RespuestaTransaccion resultado = new RespuestaTransaccion();
            try
            {
                //validar el tipo de producto
                //linea de negocio
                if (productoCE.id_tipo_producto == 904)
                {
                    var varLineasNegocioIguales = LineaNegocioEntity.ListadoLineasNegocio().Where(s => (s.nombre ?? string.Empty).ToUpper().Trim() == productoCE.nombre.ToUpper() && s.id_linea_negocio != productoCE.id_linea_negocio).ToList();

                    if (varLineasNegocioIguales.Count > 0)
                    {
                        return Json(new { Resultado = new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeValidacionCodigoProductoGestorExistente } }, JsonRequestBehavior.AllowGet);
                    }
                    else
                    {
                        LineaNegocio lineaNegocio = new LineaNegocio();
                        lineaNegocio.nombre = productoCE.nombre;
                        lineaNegocio.descripcon = productoCE.descripcon;
                        lineaNegocio.estado = true;

                        resultado = LineaNegocioEntity.CrearLineaNegocio(lineaNegocio);
                    }

                }

                if (productoCE.id_tipo_producto == 905)
                {
                    var varSublineasNegocioIguales = SublineaNegocioEntity.ListadoSublineasNegocio().Where(s => (s.nombre ?? string.Empty).ToUpper().Trim() == productoCE.nombre.ToUpper() && s.id_linea_negocio == productoCE.id_linea_negocio && s.id_sublinea_negocio != productoCE.id_sublinea_negocio).ToList();

                    if (varSublineasNegocioIguales.Count > 0)
                    {
                        return Json(new { Resultado = new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeValidacionCodigoProductoGestorExistente } }, JsonRequestBehavior.AllowGet);
                    }
                    else
                    {
                        SublineaNegocio sublineaNegocio = new SublineaNegocio();
                        sublineaNegocio.id_linea_negocio = productoCE.id_linea_negocio;
                        sublineaNegocio.nombre = productoCE.nombre;
                        sublineaNegocio.descripcon = productoCE.descripcon;
                        sublineaNegocio.estado = true;

                        resultado = SublineaNegocioEntity.CrearSublineaNegocio(sublineaNegocio);
                    }
                }

                if (productoCE.id_tipo_producto == 906)
                {
                    var varProductosGeneralIguales = ProductosGeneralGestorEntity.ListadoProductoGeneral().Where(s => (s.nombre ?? string.Empty).ToUpper().Trim() == productoCE.nombre.ToUpper() && s.id_sublinea_negocio == productoCE.id_sublinea_negocio && s.id_producto_general != productoCE.id_producto_general).ToList();

                    if (varProductosGeneralIguales.Count > 0)
                    {
                        return Json(new { Resultado = new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeValidacionCodigoProductoGestorExistente } }, JsonRequestBehavior.AllowGet);
                    }
                    else
                    {
                        ProdutosGeneralGestor productoGeneral = new ProdutosGeneralGestor();
                        productoGeneral.id_sublinea_negocio = productoCE.id_sublinea_negocio;
                        productoGeneral.nombre = productoCE.nombre;
                        productoGeneral.descripcon = productoCE.descripcon;
                        productoGeneral.estado = true;

                        resultado = ProductosGeneralGestorEntity.CrearLineaNegocio(productoGeneral);
                    }
                }

                if (productoCE.id_tipo_producto == 907)
                {
                    if (productoCE.id_producto_gestor == null || productoCE.id_producto_gestor == 0)
                    {
                        var varProductosGestorIguales = ProductosGestorEntity.ListadoProductoGestor();
                        varProductosGestorIguales = varProductosGestorIguales.Where(s => (s.nombre ?? string.Empty).ToUpper().Trim() == productoCE.nombre.ToUpper() && s.id_producto_general == productoCE.id_producto_general && s.id_producto_gestor != productoCE.id_producto_gestor).ToList();

                        if (varProductosGestorIguales.Count > 0)
                        {
                            return Json(new { Resultado = new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeValidacionCodigoProductoGestorExistente } }, JsonRequestBehavior.AllowGet);
                        }
                        else
                        {
                            ProdutosGestor productoGestor = new ProdutosGestor();
                            productoGestor.id_producto_general = productoCE.id_producto_general;
                            productoGestor.nombre = productoCE.nombre;
                            productoGestor.descripcon = productoCE.descripcon;
                            productoGestor.estado = true;

                            resultado = ProductosGestorEntity.CrearProductosGestor(productoGestor);
                        }
                    }
                    else
                    {
                        ProdutosGestor productoGestor = new ProdutosGestor();
                        productoGestor.id_producto_general = productoCE.id_producto_general;
                        productoGestor.nombre = productoCE.nombre;
                        productoGestor.descripcon = productoCE.descripcon;
                        productoGestor.estado = true;
                        productoGestor.id_producto_gestor = productoCE.id_producto_gestor;

                        resultado = ProductosGestorEntity.ActualizarProductosGestor(productoGestor);
                    }
                }


                //if (productosIguales.Count > 0)
                //{
                //    return Json(new { Resultado = new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeValidacionCodigoProductoExistente } }, JsonRequestBehavior.AllowGet);
                //}
                //else
                //{
                //    if (CodigoProductoEntity.TipoCodigoProductoExistente(producto.id_bodega.Value, producto.id_catalogo.Value, producto.id_codigo_producto))
                //        return Json(new { Resultado = new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeValidacionCodigoProductoExistente } }, JsonRequestBehavior.AllowGet);

                //    if (producto.id_codigo_producto == 0)
                //        resultado = CodigoProductoEntity.CrearCodigoProducto(producto);
                //    else
                //        resultado = CodigoProductoEntity.ActualizarCodigoProducto(producto);

                return Json(new { Resultado = resultado }, JsonRequestBehavior.AllowGet);
                //}
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
            RespuestaTransaccion resultado = ProductosGestorEntity.EliminarProductosGestor(id);

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

        public ActionResult DescargarReporteFormatoExcel()
        {
            var collection = ProductosGestorEntity.ListarProductosGestor();

            ExcelPackage package = new ExcelPackage();
            var workSheet = package.Workbook.Worksheets.Add("Productos Gestor");
            workSheet.TabColor = System.Drawing.Color.Black;
            workSheet.DefaultRowHeight = 12;

            List<string> columnas = new List<string>() {
                "ID",
                "LINEA",
                "SEGMENTACION 1",
                "SEGMENTACION 2",
                "PRODUCTO COMERCIAL",
                "ESTADO"
            };

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
                workSheet.Cells[recordIndex, 1].Value = item.Id;
                workSheet.Cells[recordIndex, 2].Value = item.LineaNegocio;
                workSheet.Cells[recordIndex, 3].Value = item.SublineaNegocio;
                workSheet.Cells[recordIndex, 4].Value = item.ProductoGeneral;
                workSheet.Cells[recordIndex, 5].Value = item.ProductoComercial;
                workSheet.Cells[recordIndex, 6].Value = item.Estado;

                recordIndex++;
            }

            for (int i = 1; i <= columnas.Count; i++)
            {
                workSheet.Column(i).AutoFit();
                workSheet.Column(i).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            }
            return File(package.GetAsByteArray(), XlsxContentType, "ListadoProductosGestor.xlsx");
        }

        public ActionResult DescargarReporteFormatoPDF()
        {
            // Seleccionar las columnas a exportar
            var results = ProductosGestorEntity.ListarProductosGestor();

            var list = Reportes.SerializeToJSON(results);
            return Content(list, "application/json");
        }

        public ActionResult DescargarReporteFormatoCSV()
        {
            var comlumHeadrs = new string[]
            {
                "ID",
                "LINEA",
                "SEGMENTACION 1",
                "SEGMENTACION 2",
                "PRODUCTO COMERCIAL",
                "ESTADO"
            };

            var listado = (from item in ProductosGestorEntity.ListarProductosGestor()
                           select new object[]
{
                               item.Id,
                               item.LineaNegocio,
                               item.SublineaNegocio,
                               item.ProductoGeneral,
                               item.ProductoComercial,
                               item.Estado

}).ToList();

            // Build the file content
            var employeecsv = new StringBuilder();
            listado.ForEach(line =>
            {
                employeecsv.AppendLine(string.Join("¬", line));
            });

            byte[] buffer = Encoding.Default.GetBytes($"{string.Join("¬", comlumHeadrs)}\r\n{employeecsv.ToString()}");
            return File(buffer, "text/csv", $"ProductosGestor.csv");
        }

    }
}
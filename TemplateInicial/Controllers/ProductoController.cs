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
    public class ProductoController : BaseAppController
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
            ViewBag.NombreListado = Etiquetas.TituloGridProducto;

            //Controlar permisos
            var user = ViewData["usuario"] = System.Web.HttpContext.Current.Session["usuario"];
            var usuario = int.Parse(user.ToString());
            string nombreControlador = ControllerContext.RouteData.Values["controller"].ToString();
            ViewBag.NombreControlador = nombreControlador;

            ViewBag.AccionesUsuario = ManejoPermisosEntity.ListadoAccionesCatalogoUsuario(usuario, nombreControlador);

            //Obtener Acciones del controlador
            ViewBag.AccionesControlador = GetMetodosControlador(nombreControlador);

            var listado = CodigoProductoEntity.ListarProductos();

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
            CodigoProducto producto = new CodigoProducto();
            try
            {
                ViewBag.TituloModal = Etiquetas.TituloPanelCreacionProducto;

                if (id.HasValue)
                {
                    producto = CodigoProductoEntity.ConsultarProducto(id.Value); 
                }  

                return PartialView(producto);
            }
            catch (Exception ex)
            {
                string mensaje = "Un error ocurrió. {0}";
                ViewBag.Excepcion = string.Format(mensaje, ex.Message.ToString());
                return PartialView("~/Views/Error/_InternalServerError.cshtml");
            }
        }
          
        [HttpPost]
        public ActionResult CreateOrUpdate(CodigoProducto producto)
        {
            RespuestaTransaccion resultado = new RespuestaTransaccion();
            try
            {
                string codigoProducto = (producto.codigo_producto ?? string.Empty).ToLower().Trim();
                string nombreProduto = (producto.nombre_producto ?? string.Empty).ToLower().Trim();

                var productosIguales = CodigoProductoEntity.ListarProductos().Where(s => (s.CodigooProducto ?? string.Empty).ToLower().Trim() == codigoProducto && s.Id != producto.id_codigo_producto).ToList();

                if (productosIguales.Count > 0)
                {
                    return Json(new { Resultado = new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeValidacionCodigoProductoExistente } }, JsonRequestBehavior.AllowGet);
                }
                else
                {
                    if (CodigoProductoEntity.TipoCodigoProductoExistente(producto.id_bodega.Value, producto.id_catalogo.Value, producto.id_codigo_producto))
                        return Json(new { Resultado = new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeValidacionCodigoProductoExistente } }, JsonRequestBehavior.AllowGet);
                     
                    if (producto.id_codigo_producto == 0)
                        resultado = CodigoProductoEntity.CrearCodigoProducto(producto);
                    else
                        resultado = CodigoProductoEntity.ActualizarCodigoProducto(producto);

                    return Json(new { Resultado = resultado }, JsonRequestBehavior.AllowGet);
                }
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
            RespuestaTransaccion resultado = CodigoProductoEntity.EliminarCodigoProducto(id);

            return Json(new { Resultado = resultado }, JsonRequestBehavior.AllowGet);
        }

        public ActionResult DescargarReporteFormatoExcel()
        {
            var collection = CodigoProductoEntity.ListarProductos();

            ExcelPackage package = new ExcelPackage();
            var workSheet = package.Workbook.Worksheets.Add("CódigoProducto");
            workSheet.TabColor = System.Drawing.Color.Black;
            workSheet.DefaultRowHeight = 12;

            List<string> columnas = new List<string>() {
                "ID",
                "BODEGA",
                "CÓDIGO BODEGA",
                "SUBLINEA",
                "NOMBRE PRODUCTO",
                "CÓDIGO PRODUCTO",
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
                workSheet.Cells[recordIndex, 2].Value = item.Bodega;
                workSheet.Cells[recordIndex, 3].Value = item.CodigoBodega;
                workSheet.Cells[recordIndex, 4].Value = item.Tarifario;
                workSheet.Cells[recordIndex, 5].Value = item.NombreProducto;
                workSheet.Cells[recordIndex, 6].Value = item.CodigooProducto;
                workSheet.Cells[recordIndex, 7].Value = item.Estado;

                recordIndex++;
            }

            for (int i = 1; i <= columnas.Count; i++)
            {
                workSheet.Column(i).AutoFit();
                workSheet.Column(i).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            }
            return File(package.GetAsByteArray(), XlsxContentType, "ListadoProductos.xlsx");
        }

        public ActionResult DescargarReporteFormatoPDF()
        {
            // Seleccionar las columnas a exportar
            var results = CodigoProductoEntity.ListarProductos();

            var list = Reportes.SerializeToJSON(results);
            return Content(list, "application/json");
        }

        public ActionResult DescargarReporteFormatoCSV()
        {
            var comlumHeadrs = new string[]
            {
                "ID",
                "BODEGA",
                "CÓDIGO BODEGA",
                "SUBLINEA",
                "NOMBRE PRODUCTO",
                "CÓDIGO PRODUCTO",
                "ESTADO"
            };

            var listado = (from item in CodigoProductoEntity.ListarProductos() select new object[]
                                          {
                               item.Id,
                               item.Bodega,
                               item.CodigoBodega,
                               item.Tarifario,
                               item.NombreProducto,
                               item.CodigooProducto,
                               item.Estado

                                          }).ToList();

            // Build the file content
            var employeecsv = new StringBuilder();
            listado.ForEach(line =>
            {
                employeecsv.AppendLine(string.Join("¬", line));
            });

            byte[] buffer = Encoding.Default.GetBytes($"{string.Join("¬", comlumHeadrs)}\r\n{employeecsv.ToString()}");
            return File(buffer, "text/csv", $"CodigoProductos.csv");
        }

    }
}
using System;
using System.Collections.Generic;
using System.Data; 
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using System.Web;
using System.Web.Mvc;
using GestionPPM.Entidades.Metodos;
using GestionPPM.Entidades.Modelo;
using GestionPPM.Repositorios;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using Seguridad.Helper;

namespace TemplateInicial.Controllers
{
    [Autenticado]
    public class TablaPlanesController : BaseAppController
    { 
        private const string XlsxContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
        // GET: Tarifarios
        public ActionResult Index()
        {
            var respuesta = System.Web.HttpContext.Current.Session["Resultado"] as string;
            var estado = System.Web.HttpContext.Current.Session["Estado"] as string;

            ViewBag.Resultado = respuesta;
            ViewBag.Estado = estado;

            Session["Resultado"] = "";
            Session["Estado"] = "";
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
            ViewBag.NombreListado = Etiquetas.TituloGridTablaPlanes;

            //Controlar permisos
            var user = ViewData["usuario"] = System.Web.HttpContext.Current.Session["usuario"];
            var usuario = int.Parse(user.ToString());
            string nombreControlador = ControllerContext.RouteData.Values["controller"].ToString();
            ViewBag.NombreControlador = nombreControlador;

            ViewBag.AccionesUsuario = ManejoPermisosEntity.ListadoAccionesCatalogoUsuario(usuario, nombreControlador);

            //Obtener Acciones del controlador
            ViewBag.AccionesControlador = GetMetodosControlador(nombreControlador);

            //Búsqueda

            var listado = TablaPlanesEntity.ListarTablaPlanes(); 
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

        // GET: Tarifarios/Create
        public ActionResult Create()
        { 
            return View();
        }

        [HttpPost]
        public ActionResult Create(TablaPlanes planes)
        {
            try
            {
                string nombrePlan = (planes.nombre_plan ?? string.Empty).ToLower().Trim(); 
                 
                RespuestaTransaccion resultado = TablaPlanesEntity.CrearTablaPlanes(planes);

                return Json(new { Resultado = resultado }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                return Json(new { Resultado = new RespuestaTransaccion { Estado = false, Respuesta = ex.Message } }, JsonRequestBehavior.AllowGet);
            }
        }

        // GET: Tarifarios/Edit/5
        public ActionResult Edit(int? id)
        { 
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            var menu = TablaPlanesEntity.ConsultarTablaPlanes(id.Value);
            ViewBag.TransaccionMinima = menu.transaccion_minima.HasValue ? menu.transaccion_minima : 0;
            ViewBag.TransaccionMaxima = menu.transaccion_maxima.HasValue ? menu.transaccion_maxima : 0;
            ViewBag.ValorMinimo = menu.valor_minimo.HasValue ? menu.valor_minimo : 0;
            ViewBag.VamorMaximo = menu.valor_maximo.HasValue ? menu.valor_maximo : 0;
            ViewBag.CostoTransaccion = menu.costo_x_transaccion.HasValue ? menu.costo_x_transaccion : 0;
            ViewBag.CargoFijo = menu.valor_fijo.HasValue ? menu.valor_fijo : 0;

            if (menu == null)
            {
                return HttpNotFound();
            }
            return View(menu);
        }

        [HttpPost]
        public ActionResult Edit(TablaPlanes planes)
        {
            try
            {
                string nombrePlan = (planes.nombre_plan ?? string.Empty).ToLower().Trim();
                 
                RespuestaTransaccion resultado = TablaPlanesEntity.ActualizarTablaPlanes(planes);
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
            RespuestaTransaccion resultado = TablaPlanesEntity.EliminarTablaPlanes(id);// await db.Cabecera.FindAsync(id);

            //return RedirectToAction("Index");
            return Json(new { Resultado = resultado }, JsonRequestBehavior.AllowGet);

        }          

        public ActionResult DescargarReporteFormatoExcel()
        {
            var collection = TablaPlanesEntity.ListarTablaPlanes();

            ExcelPackage package = new ExcelPackage();
            var workSheet = package.Workbook.Worksheets.Add("Tabla Planes");
            workSheet.TabColor = System.Drawing.Color.Black;
            workSheet.DefaultRowHeight = 12;

            List<string> columnas = new List<string>() {
                "ID",
                "NOMBRE",
                "TRANSACCION MINIMA",
                "TRANSACCION MAXIMA",
                "VALOR MINIMO",
                "VALOR MAXIMO",
                "TIPO COBRO",
                "COSTO TRANSACCION",
                "CARGO FIJO",
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
                workSheet.Cells[recordIndex, 1].Value = item.Codigo;
                workSheet.Cells[recordIndex, 2].Value = item.Nombre_Plan;
                workSheet.Cells[recordIndex, 3].Value = item.Transaccion_Minima;
                workSheet.Cells[recordIndex, 3].Style.Numberformat.Format = "###,##0";
                workSheet.Cells[recordIndex, 4].Value = item.Transaccion_Maxima;
                workSheet.Cells[recordIndex, 4].Style.Numberformat.Format = "###,##0";
                workSheet.Cells[recordIndex, 5].Value = item.Valor_Minimo;
                workSheet.Cells[recordIndex, 5].Style.Numberformat.Format = "###,##0.00";
                workSheet.Cells[recordIndex, 6].Value = item.Valor_Maximo;
                workSheet.Cells[recordIndex, 6].Style.Numberformat.Format = "###,##0.00";
                workSheet.Cells[recordIndex, 7].Value = item.Tipo_Cobro;
                workSheet.Cells[recordIndex, 8].Value = item.Costo_Transaccion;
                workSheet.Cells[recordIndex, 8].Style.Numberformat.Format = "###,##0.00";
                workSheet.Cells[recordIndex, 9].Value = item.Valor_Pago_Mínimo;
                workSheet.Cells[recordIndex, 9].Style.Numberformat.Format = "###,##0.00";
                workSheet.Cells[recordIndex, 10].Value = item.Estado_Plan;                 

                recordIndex++;
            }

            for (int i = 1; i <= columnas.Count; i++)
            {
                workSheet.Column(i).AutoFit();
                workSheet.Column(i).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                if (i == 3)
                {
                    workSheet.Column(i).Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                }
            }

            return File(package.GetAsByteArray(), XlsxContentType, "ListadoTablaPlanes.xlsx"); 

        }

        public ActionResult DescargarReporteFormatoPDF()
        {
            // Seleccionar las columnas a exportar
            var results = TarifarioEntity.ListarTarifario();

            var list = Reportes.SerializeToJSON(results);
            return Content(list, "application/json");
        }

        public ActionResult DescargarReporteFormatoCSV()
        {
            var comlumHeadrs = new string[]
            {
                "ID",
                "NOMBRE",
                "TRANSACCION MINIMA",
                "TRANSACCION MAXIMA",
                "VALOR MINIMO",
                "VALOR MAXIMO",
                "TIPO COBRO",
                "COSTO TRANSACCION",
                "CARGO FIJO",
                "ESTADO"
            };

            var listado = (from item in TablaPlanesEntity.ListarTablaPlanes()
                           select new object[]
                           {
                               item.Codigo,
                               item.Nombre_Plan,
                               item.Transaccion_Maxima,
                               item.Transaccion_Minima,
                               item.Valor_Maximo,
                               item.Valor_Minimo,
                               item.Tipo_Cobro,
                               item.Costo_Transaccion,
                               item.Valor_Pago_Mínimo,
                               item.Estado_Plan

                           }).ToList();

            // Build the file content
            var employeecsv = new StringBuilder();
            listado.ForEach(line =>
            {
                employeecsv.AppendLine(string.Join(",", line));
            });

            byte[] buffer = Encoding.Default.GetBytes($"{string.Join(",", comlumHeadrs)}\r\n{employeecsv.ToString()}");
            return File(buffer, "text/csv", $"Tabla_Planes.csv");
        }

    }
}

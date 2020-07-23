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
    public class ImpuestoController : BaseAppController
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
            ViewBag.NombreListado = Etiquetas.TituloGridImpuesto;
            //Controlar permisos
            var user = ViewData["usuario"] = System.Web.HttpContext.Current.Session["usuario"];
            var usuario = int.Parse(user.ToString());
            string nombreControlador = ControllerContext.RouteData.Values["controller"].ToString();
            ViewBag.NombreControlador = nombreControlador;

            ViewBag.AccionesUsuario = ManejoPermisosEntity.ListadoAccionesCatalogoUsuario(usuario, nombreControlador);

            //Obtener Acciones del controlador
            ViewBag.AccionesControlador = GetMetodosControlador(nombreControlador);

            //Búsqueda
            var listado = ImpuestoEntity.ListarImpuestos();

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

        // GET: Create
        public ActionResult Create()
        { 
            return View();
        }

        [HttpPost]
        public ActionResult Create(Impuesto impuesto)
        {
            try
            {
                string nombreImpuesto = (impuesto.nombre_impuesto ?? string.Empty).ToLower().Trim();

                var impuestosIguales = ImpuestoEntity.ListarImpuestos().Where(s => (s.Nombre ?? string.Empty).ToLower().Trim() == nombreImpuesto).ToList();

                if (impuestosIguales.Count > 0)
                {
                    return Json(new { Resultado = new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeValidacionNombreImpuesto } }, JsonRequestBehavior.AllowGet);
                }
                else
                {
                    RespuestaTransaccion resultado = ImpuestoEntity.CrearImpuesto(impuesto);
                    return Json(new { Resultado = resultado }, JsonRequestBehavior.AllowGet);
                }             
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
            var impuesto = ImpuestoEntity.ConsultarImpuesto(id.Value);
            ViewBag.ValorUnitario = impuesto.valor.HasValue ? impuesto.valor : 0;

            if (impuesto == null)
            {
                return HttpNotFound();
            }
            return View(impuesto);
        }

        [HttpPost]
        public ActionResult Edit(Impuesto impuesto)
        {
            try
            {
                string nombreImpuesto = (impuesto.nombre_impuesto ?? string.Empty).ToLower().Trim();

                var impuestosIguales = ImpuestoEntity.ListarImpuestos().Where(s => (s.Nombre ?? string.Empty).ToLower().Trim() == nombreImpuesto && s.Codigo != impuesto.id_impuesto).ToList();                 

                if (impuestosIguales.Count > 0)
                {
                    return Json(new { Resultado = new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeValidacionNombreTarifario } }, JsonRequestBehavior.AllowGet);
                }
                else
                {
                    RespuestaTransaccion resultado = ImpuestoEntity.ActualizarImpuestos(impuesto);
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
            RespuestaTransaccion resultado = ImpuestoEntity.EliminarImpuesto(id);// await db.Cabecera.FindAsync(id);

            //return RedirectToAction("Index");
            return Json(new { Resultado = resultado }, JsonRequestBehavior.AllowGet);

        }          

        public ActionResult DescargarReporteFormatoExcel()
        {
            var collection = ImpuestoEntity.ListarImpuestos();

            ExcelPackage package = new ExcelPackage();
            var workSheet = package.Workbook.Worksheets.Add("Impuestos");
            workSheet.TabColor = System.Drawing.Color.Black;
            workSheet.DefaultRowHeight = 12;

            List<string> columnas = new List<string>() {
                "ID",
                "NOMBRE",
                "VALOR", 
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
                workSheet.Cells[recordIndex, 2].Value = item.Nombre;
                workSheet.Cells[recordIndex, 3].Value = item.Valor;
                workSheet.Cells[recordIndex, 3].Style.Numberformat.Format = "###,##0.00"; 
                workSheet.Cells[recordIndex, 4].Value = item.Estado;                 

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

            return File(package.GetAsByteArray(), XlsxContentType, "ListadoImpuestos.xlsx"); 

        }

        public ActionResult DescargarReporteFormatoPDF()
        {
            // Seleccionar las columnas a exportar
            var results = ImpuestoEntity.ListarImpuestos();

            var list = Reportes.SerializeToJSON(results);
            return Content(list, "application/json");
        }

        public ActionResult DescargarReporteFormatoCSV()
        {
            var comlumHeadrs = new string[]
            {
                "ID",
                "NOMBRE",
                "VALOR", 
                "ESTADO",
            };

            var listado = (from item in ImpuestoEntity.ListarImpuestos()
                           select new object[]
                           {
                               item.Codigo,
                               item.Nombre,
                               item.Valor, 
                               item.Estado

                           }).ToList();

            // Build the file content
            var employeecsv = new StringBuilder();
            listado.ForEach(line =>
            {
                employeecsv.AppendLine(string.Join("¬", line));
            });

            byte[] buffer = Encoding.Default.GetBytes($"{string.Join("¬", comlumHeadrs)}\r\n{employeecsv.ToString()}");
            return File(buffer, "text/csv", $"Impuestos.csv");
        }
    }
}

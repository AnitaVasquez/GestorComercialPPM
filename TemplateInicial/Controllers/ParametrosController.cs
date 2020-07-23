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
    public class ParametrosController : BaseAppController
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
            ViewBag.NombreListado = Etiquetas.TituloGridParametros;
            //Controlar permisos
            var user = ViewData["usuario"] = System.Web.HttpContext.Current.Session["usuario"];
            var usuario = int.Parse(user.ToString());
            string nombreControlador = ControllerContext.RouteData.Values["controller"].ToString();
            ViewBag.NombreControlador = nombreControlador;

            ViewBag.AccionesUsuario = ManejoPermisosEntity.ListadoAccionesCatalogoUsuario(usuario, nombreControlador);

            //Obtener Acciones del controlador
            ViewBag.AccionesControlador = GetMetodosControlador(nombreControlador);

            //Búsqueda

            var listado = ParametrosSistemaEntity.ListarParametros(); 
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
        public ActionResult Create(ParametrosSistema parametros)
        {
            try
            {
                string nombrePlan = (parametros.nombre ?? string.Empty).ToLower().Trim(); 
                 
                RespuestaTransaccion resultado = ParametrosSistemaEntity.CrearParametrosSistema(parametros);

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
            var parametro = ParametrosSistemaEntity.ConsultarParametros(id.Value); 

            if (parametro == null)
            {
                return HttpNotFound();
            }
            else
            {
                ViewBag.Valor = parametro.valor.ToString();
                ViewBag.Tipo = parametro.tipo.ToString();
                return View(parametro);
            }
        }

        [HttpPost]
        public ActionResult Edit(ParametrosSistema parametros)
        {
            try
            {
                string nombrePlan = (parametros.nombre ?? string.Empty).ToLower().Trim();
                 
                RespuestaTransaccion resultado = ParametrosSistemaEntity.ActualizarParametrosSistema(parametros);
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
            RespuestaTransaccion resultado = ParametrosSistemaEntity.EliminarParametrosSistema(id);// await db.Cabecera.FindAsync(id);
             
            return Json(new { Resultado = resultado }, JsonRequestBehavior.AllowGet);
        }          

        public ActionResult DescargarReporteFormatoExcel()
        {
            var collection = ParametrosSistemaEntity.ListarParametros();

            ExcelPackage package = new ExcelPackage();
            var workSheet = package.Workbook.Worksheets.Add("Tabla Planes");
            workSheet.TabColor = System.Drawing.Color.Black;
            workSheet.DefaultRowHeight = 12;

            List<string> columnas = new List<string>() {
                "ID",
                "NOMBRE",
                "DESCRIPCIÓN",
                "VALOR",
                "TIPO", 
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
                workSheet.Cells[recordIndex, 1].Value = item.id_parametro;
                workSheet.Cells[recordIndex, 2].Value = item.nombre;
                workSheet.Cells[recordIndex, 3].Value = item.descripcion; 
                workSheet.Cells[recordIndex, 4].Value = item.valor; 
                workSheet.Cells[recordIndex, 5].Value = item.tipo; 
                workSheet.Cells[recordIndex, 6].Value = item.estado;               

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

            return File(package.GetAsByteArray(), XlsxContentType, "ListadoParametros.xlsx"); 

        }

        public ActionResult DescargarReporteFormatoPDF()
        {
            // Seleccionar las columnas a exportar
            var results = ParametrosSistemaEntity.ListarParametros();

            var list = Reportes.SerializeToJSON(results);
            return Content(list, "application/json");
        }

        public ActionResult DescargarReporteFormatoCSV()
        {
            var comlumHeadrs = new string[]
            {
                "ID",
                "NOMBRE",
                "DESCRIPCIÓN",
                "VALOR",
                "TIPO",
                "ESTADO"
            };

            var listado = (from item in ParametrosSistemaEntity.ListarParametros()
            select new object[]
                           {
                               item.id_parametro,
                               item.nombre,
                               item.descripcion,
                               item.valor,
                               item.tipo,
                               item.estado 
                           }).ToList();

            // Build the file content
            var employeecsv = new StringBuilder();
            listado.ForEach(line =>
            {
                employeecsv.AppendLine(string.Join(",", line));
            });

            byte[] buffer = Encoding.Default.GetBytes($"{string.Join(",", comlumHeadrs)}\r\n{employeecsv.ToString()}");
            return File(buffer, "text/csv", $"ListadoParametros.csv");
        }

    }
}

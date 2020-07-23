using GestionPPM.Entidades.Metodos;
using GestionPPM.Entidades.Modelo;
using GestionPPM.Repositorios;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using Seguridad.Helper;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using System.Web;
using System.Web.Mvc;

namespace TemplateInicial.Controllers
{
    [Autenticado]
    public class BodegaController : BaseAppController
    {
        private const string XlsxContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

        // GET: Bodega
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
            ViewBag.NombreListado = Etiquetas.TituloGridBodega;

            //Controlar permisos
            var user = ViewData["usuario"] = System.Web.HttpContext.Current.Session["usuario"];
            var usuario = int.Parse(user.ToString());
            string nombreControlador = ControllerContext.RouteData.Values["controller"].ToString();
            ViewBag.NombreControlador = nombreControlador;

            ViewBag.AccionesUsuario = ManejoPermisosEntity.ListadoAccionesCatalogoUsuario(usuario, nombreControlador);

            //Obtener Acciones del controlador
            ViewBag.AccionesControlador = GetMetodosControlador(nombreControlador);

            //Búsqueda
            var listado = BodegaEntity.ListarBodegas();

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
        
        // GET: Bodega/Details/5
        public ActionResult Details(int id)
        {
            return View();
        }

        // GET: Bodega/Create
        public ActionResult Create()
        {
            return View();
        }

        // POST: Bodega/Create
        [HttpPost]
        public ActionResult Create(Bodega bodega)
        {
            try
            {
                string nombreBodega = (bodega.nombre_bodega ?? string.Empty).ToLower().Trim();
                string codigoBodega = (bodega.codigo_bodega ?? string.Empty).ToLower().Trim();

                var bodegaIguales = BodegaEntity.ListarBodegas().Where(s => (s.Codigo ?? string.Empty).ToLower().Trim() == codigoBodega).ToList();
                var nombresIguales = BodegaEntity.ListarBodegas().Where(s => (s.Nombre ?? string.Empty).ToLower().Trim() == nombreBodega).ToList();

                if (bodegaIguales.Count > 0 || nombresIguales.Count > 0)
                {
                    if (bodegaIguales.Count > 0)
                    {
                        return Json(new { Resultado = new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeValidacionCodigoBodega } }, JsonRequestBehavior.AllowGet);
                    }
                    else
                    {
                        return Json(new { Resultado = new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeValidacionNombreBodega } }, JsonRequestBehavior.AllowGet);
                    }
                }
                else
                {
                    RespuestaTransaccion resultado = BodegaEntity.CrearBodega(bodega);
                    return Json(new { Resultado = resultado }, JsonRequestBehavior.AllowGet);
                }
            }
            catch
            {
                return View();
            }
        }

        // GET: Bodega/Edit/5
        public ActionResult Edit(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            var bodega = BodegaEntity.ConsultarBodega(id.Value); 

            if (bodega == null)
            {
                return HttpNotFound();
            }
            return View(bodega); 
        }

        // POST: Bodega/Edit/5
        [HttpPost]
        public ActionResult Edit(Bodega bodega)
        {
            try
            {
                string nombreBodega = (bodega.nombre_bodega ?? string.Empty).ToLower().Trim();
                string codigoBodega = (bodega.codigo_bodega ?? string.Empty).ToLower().Trim();

                var bodegaIguales = BodegaEntity.ListarBodegas().Where(s => (s.Codigo ?? string.Empty).ToLower().Trim() == codigoBodega && s.Id != bodega.id_bodega).ToList();
                var nombresIguales = BodegaEntity.ListarBodegas().Where(s => (s.Nombre ?? string.Empty).ToLower().Trim() == nombreBodega && s.Id != bodega.id_bodega).ToList();

                if (bodegaIguales.Count > 0 || nombresIguales.Count > 0)
                {
                    if (bodegaIguales.Count > 0)
                    {
                        return Json(new { Resultado = new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeValidacionCodigoBodega } }, JsonRequestBehavior.AllowGet);
                    }
                    else
                    {
                        return Json(new { Resultado = new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeValidacionNombreBodega } }, JsonRequestBehavior.AllowGet);
                    }
                }
                else
                {
                    RespuestaTransaccion resultado = BodegaEntity.ActualizarBodega(bodega);
                    return Json(new { Resultado = resultado }, JsonRequestBehavior.AllowGet);
                } 
            }
            catch
            {
                return View();
            }
        }

        // GET: Bodega/Delete/5
        public ActionResult Delete(int id)
        {
            return View();
        }

        // POST: Bodega/Delete/5
        [HttpPost]
        public ActionResult Delete(int id, FormCollection collection)
        {
            try
            {
                // TODO: Add delete logic here

                return RedirectToAction("Index");
            }
            catch
            {
                return View();
            }
        }

        [HttpPost]
        public ActionResult Eliminar(int id)
        {
            if (id == 0)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            RespuestaTransaccion resultado = BodegaEntity.EliminarBodega(id);// await db.Cabecera.FindAsync(id);

            //return RedirectToAction("Index");
            return Json(new { Resultado = resultado }, JsonRequestBehavior.AllowGet);

        }

        public ActionResult DescargarReporteFormatoExcel()
        {
            var collection = BodegaEntity.ListarBodegas();

            ExcelPackage package = new ExcelPackage();
            var workSheet = package.Workbook.Worksheets.Add("Bodega");
            workSheet.TabColor = System.Drawing.Color.Black;
            workSheet.DefaultRowHeight = 12;

            List<string> columnas = new List<string>() {
                "ID",
                "NOMBRE",
                "DESCRIPCIÓN",
                "CÓDIGO",
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
                workSheet.Cells[recordIndex, 2].Value = item.Nombre;
                workSheet.Cells[recordIndex, 3].Value = item.Descripcion;
                workSheet.Cells[recordIndex, 4].Value = item.Codigo;
                workSheet.Cells[recordIndex, 5].Value = item.Estado;

                recordIndex++;
            }

            for (int i = 1; i <= columnas.Count; i++)
            {
                workSheet.Column(i).AutoFit();
                workSheet.Column(i).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;                 
            }
            return File(package.GetAsByteArray(), XlsxContentType, "ListadoBodegas.xlsx");
        }

        public ActionResult DescargarReporteFormatoPDF()
        {
            // Seleccionar las columnas a exportar
            var results = BodegaEntity.ListarBodegas();

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
                "CÓDIGO",
                "ESTADO"
            };

            var listado = (from item in BodegaEntity.ListarBodegas()
            select new object[]
                           {
                               item.Id,
                               item.Nombre,
                               item.Descripcion,
                               item.Codigo,
                               item.Estado

                           }).ToList();

            // Build the file content
            var employeecsv = new StringBuilder();
            listado.ForEach(line =>
            {
                employeecsv.AppendLine(string.Join("¬", line));
            });

            byte[] buffer = Encoding.Default.GetBytes($"{string.Join("¬", comlumHeadrs)}\r\n{employeecsv.ToString()}");
            return File(buffer, "text/csv", $"Bodegas.csv");
        }

    }
} 

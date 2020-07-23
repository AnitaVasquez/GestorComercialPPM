using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data; 
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using System.Web;
using System.Web.Mvc;
using GestionPPM.Entidades;
using GestionPPM.Entidades.Metodos;
using GestionPPM.Entidades.Modelo;
using GestionPPM.Repositorios;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using Omu.Awem.Helpers; 
using Seguridad.Helper;

namespace TemplateInicial.Controllers
{ 
    [Autenticado]
    public class FacturaComercialController : BaseAppController
    { 
        //Perfil del usuario
        private const string XlsxContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
        private int  rolID = Convert.ToInt32(System.Web.HttpContext.Current.Session["rolID"]);

        // GET: Cliente
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

        [HttpGet]
        public async Task<PartialViewResult> IndexGrid(String search)
        {
            ViewBag.NombreListado = Etiquetas.TituloGridFacturarSAFI;

            //Controlar permisos
            var user = ViewData["usuario"] = System.Web.HttpContext.Current.Session["usuario"];
            var usuario = int.Parse(user.ToString());
            string nombreControlador = ControllerContext.RouteData.Values["controller"].ToString();
            ViewBag.NombreControlador = nombreControlador;

            ViewBag.AccionesUsuario = ManejoPermisosEntity.ListadoAccionesCatalogoUsuario(usuario, nombreControlador);

            //Obtener Acciones del controlador
            ViewBag.AccionesControlador = GetMetodosControlador(nombreControlador);

            //Búsqueda

            var listado = FacturasComercialEntity.ListadoPrefacturasFacturar();

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
            SAFIGeneral prefactura = new SAFIGeneral();
            try
            {
                ViewBag.TituloModal = Etiquetas.TituloPanelFacturarSAFI;

                if (id.HasValue)
                {
                    prefactura = FacturasComercialEntity.ConsultarPrefactura(id.Value);
                }

                return PartialView(prefactura);
            }
            catch (Exception ex)
            {
                string mensaje = "Un error ocurrió. {0}";
                ViewBag.Excepcion = string.Format(mensaje, ex.Message.ToString());
                return PartialView("~/Views/Error/_InternalServerError.cshtml");
            }
        }

        [HttpPost]
        public ActionResult CreateOrUpdate(SAFIGeneral prefactura)
        {
            RespuestaTransaccion resultado = new RespuestaTransaccion();
            try
            {
                SAFIGeneral prefacturaActual = new SAFIGeneral();
                prefacturaActual = FacturasComercialEntity.ConsultarPrefactura(prefactura.id_facturacion_safi);

                //cambiar el concepto
                prefacturaActual.detalle_cotizacion = prefactura.detalle_cotizacion;
                prefacturaActual.cuenta_contable = prefactura.cuenta_contable;
                prefacturaActual.correos_facturacion = prefactura.correos_facturacion;

                //actualizar el registro
                resultado = FacturasComercialEntity.ActualizarDetallePrefactura(prefacturaActual);
                return Json(new { Resultado = resultado }, JsonRequestBehavior.AllowGet);

            }
            catch (Exception ex)
            {
                return Json(new { Resultado = new RespuestaTransaccion { Estado = false, Respuesta = ex.Message } }, JsonRequestBehavior.AllowGet);
            }
        }


        [HttpPost]
        public ActionResult FacturaIndividual(int id)
        {
            if (id == 0)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            RespuestaTransaccion resultado = FacturasComercialEntity.FacturarSAFI(id);// await db.Cabecera.FindAsync(id);

            //return RedirectToAction("Index");
            return Json(new { Resultado = resultado }, JsonRequestBehavior.AllowGet);

        }

        [HttpPost] 
        public ActionResult EnvioMasivoSistemaContable(List<int> codigosPresupuestos)
        { 
            RespuestaTransaccion resultado = new RespuestaTransaccion();

            foreach (var item in codigosPresupuestos)
            {
                resultado = FacturasComercialEntity.FacturarSAFI(item);
            }
             
            return Json(new { Resultado = resultado }, JsonRequestBehavior.AllowGet); 
        }

        public ActionResult DescargarReporteFormatoExcel()
        {
            var collection = FacturasComercialEntity.ListadoPrefacturasFacturar();

            ExcelPackage package = new ExcelPackage();
            var workSheet = package.Workbook.Worksheets.Add("Listado Presupuestos a SAFI");
            workSheet.TabColor = System.Drawing.Color.Black;
            workSheet.DefaultRowHeight = 12;

            List<string> columnas = new List<string>() {
                "CLIENTE",
                "IDENTIFICACION",
                "DIRECCION",
                "PRESUPUESTOS",
                "DETALLE",
                "CORREOS CONTACTOS",
                "FECHA PRESUPUESTO",
                "CANTIDAD", 
                "SUBTOTAL",
                "IVA",
                "DESCUENTO",
                "TOTAL"                  
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
                workSheet.Cells[recordIndex, 1].Value = item.RazonSocial;
                workSheet.Cells[recordIndex, 2].Value = item.RUC;
                workSheet.Cells[recordIndex, 3].Value = item.Direccion;
                workSheet.Cells[recordIndex, 4].Value = item.Presupuesto;
                workSheet.Cells[recordIndex, 5].Value = item.detalle_cotizacion;
                workSheet.Cells[recordIndex, 6].Value = item.Correos;
                workSheet.Cells[recordIndex, 7].Value = item.FechaPresupuesto;
                workSheet.Cells[recordIndex, 8].Value = item.Cantidad;
                workSheet.Cells[recordIndex, 9].Value = item.Subtotal;
                workSheet.Cells[recordIndex, 9].Style.Numberformat.Format = "###,##0.00";
                workSheet.Cells[recordIndex, 10].Value = item.Iva;
                workSheet.Cells[recordIndex, 10].Style.Numberformat.Format = "###,##0.00";
                workSheet.Cells[recordIndex, 11].Value = item.Descuento;
                workSheet.Cells[recordIndex, 11].Style.Numberformat.Format = "###,##0.00";
                workSheet.Cells[recordIndex, 12].Value = item.Total;
                workSheet.Cells[recordIndex, 12].Style.Numberformat.Format = "###,##0.00"; 
                recordIndex++;
            }

            for (int i = 1; i <= columnas.Count; i++)
            {
                workSheet.Column(i).AutoFit();
                workSheet.Column(i).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            }


            return File(package.GetAsByteArray(), XlsxContentType, "ListadoPresupuestosSAFI.xlsx");
        }

        public ActionResult DescargarReporteFormatoPDF()
        {
            // Seleccionar las columnas a exportar
            var results = FacturasComercialEntity.ListadoPrefacturasFacturar();

            var list = Reportes.SerializeToJSON(results);
            return Content(list, "application/json");
        }

        public ActionResult DescargarReporteFormatoCSV()
        {
            var comlumHeadrs = new string[]
            {
                "CLIENTE",
                "IDENTIFICACION",
                "DIRECCION",
                "PRESUPUESTOS",
                "DETALLE",
                "CORREOS CONTACTOS",
                "FECHA PRESUPUESTO",
                "CANTIDAD",
                "SUBTOTAL",
                "IVA",
                "DESCUENTO",
                "TOTAL"
            };

            var listado = (from item in FacturasComercialEntity.ListadoPrefacturasFacturar()
            select new object[]
                           {
                                item.RazonSocial, 
                                item.RUC,
                                item.Direccion,
                                item.Presupuesto,
                                item.detalle_cotizacion,
                                item.Correos,
                                item.FechaPresupuesto,
                                item.Cantidad,
                                item.Subtotal,
                                item.Iva,
                                item.Descuento, 
                                item.Total 
                           }).ToList();

            // Build the file content
            var employeecsv = new StringBuilder();
            listado.ForEach(line =>
            {
                employeecsv.AppendLine(string.Join(",", line));
            });

            byte[] buffer = Encoding.Default.GetBytes($"{string.Join(",", comlumHeadrs)}\r\n{employeecsv.ToString()}");
            return File(buffer, "text/csv", $"ListadoPresupuestosSAFI.csv");
        }
           
    }
}

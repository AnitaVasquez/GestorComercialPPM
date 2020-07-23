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
    public class PresupuestosSAFIController : BaseAppController
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
            ViewBag.NombreListado = Etiquetas.TituloPanelPrefacturaSAFI;

            //Controlar permisos
            var user = ViewData["usuario"] = System.Web.HttpContext.Current.Session["usuario"];
            var usuario = int.Parse(user.ToString());
            string nombreControlador = ControllerContext.RouteData.Values["controller"].ToString();
            ViewBag.NombreControlador = nombreControlador;

            ViewBag.AccionesUsuario = ManejoPermisosEntity.ListadoAccionesCatalogoUsuario(usuario, nombreControlador);

            //Obtener Acciones del controlador
            ViewBag.AccionesControlador = GetMetodosControlador(nombreControlador);

            var listado = PrefacturasSAFIEntity.ListadoCodigosPrefacturar();

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
            
        [HttpPost]
        public ActionResult PresupuestoIndividual(int id)
        {
            if (id == 0)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            RespuestaTransaccion resultado = PrefacturasSAFIEntity.PrefacturarSAFI(id);// await db.Cabecera.FindAsync(id);

            //return RedirectToAction("Index");
            return Json(new { Resultado = resultado }, JsonRequestBehavior.AllowGet);

        }

        [HttpPost] 
        public ActionResult EnvioMasivoSistemaContable(List<int> codigoCotizacion)
        { 
            RespuestaTransaccion resultado = new RespuestaTransaccion();
            bool estadoMensaje = true;
            var mensaje = "";
            var codigos = "";

            //armar una lista de codigos a realizar presupuestos
            foreach (var item in codigoCotizacion)
            {
                //Armar consolidado
                if (codigos == "")
                {
                    codigos = item.ToString();
                }
                else
                {
                    codigos += "," + item.ToString();
                }
            }

            foreach (var item in codigoCotizacion)
            {
                //Armar consolidado
                if (codigos == "")
                {
                    codigos = item.ToString();
                }
                else
                {
                    codigos += "," + item.ToString();
                }

                //generar prefactura
                resultado = PrefacturasSAFIEntity.PrefacturarSAFI(item);

                if (resultado.Estado == false)
                {
                    estadoMensaje = false;
                    mensaje = resultado.Respuesta.ToString();
                }
            }

            //enviar correo de prefacturas generadas


            if(mensaje != "")
            { 
                return Json(new { Resultado = new RespuestaTransaccion { Estado = false, Respuesta = mensaje } }, JsonRequestBehavior.AllowGet);
            }
            else
            {
                return Json(new { Resultado = new RespuestaTransaccion { Estado = true, Respuesta = Mensajes.MensajeTransaccionExitosa } }, JsonRequestBehavior.AllowGet);
            }           
            
        }

        public ActionResult DescargarReporteFormatoExcel()
        {
            var collection = PrefacturasSAFIEntity.ListadoCodigosPrefacturar();

            ExcelPackage package = new ExcelPackage();
            var workSheet = package.Workbook.Worksheets.Add("Listado Cotizaciones Presupuestos");
            workSheet.TabColor = System.Drawing.Color.Black;
            workSheet.DefaultRowHeight = 12;

            List<string> columnas = new List<string>() {
                "CODIGO DE COTIZACION",
                "DETALLE",
                "RUC CLIENTE",
                "CLIENTE",
                "CORREO",
                "EJECUTIVO",
                "SUBTOTAL" 
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
                workSheet.Cells[recordIndex, 1].Value = item.CodigoCotizacion;
                workSheet.Cells[recordIndex, 2].Value = item.Detalle;
                workSheet.Cells[recordIndex, 3].Value = item.RUC;
                workSheet.Cells[recordIndex, 4].Value = item.Cliente;
                workSheet.Cells[recordIndex, 5].Value = item.Correo;
                workSheet.Cells[recordIndex, 6].Value = item.Ejecutivo; 
                workSheet.Cells[recordIndex, 7].Value = item.Valor;
                workSheet.Cells[recordIndex, 7].Style.Numberformat.Format = "###,##0.00";                
                recordIndex++;
            }

            for (int i = 1; i <= columnas.Count; i++)
            {
                workSheet.Column(i).AutoFit();
                workSheet.Column(i).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            }


            return File(package.GetAsByteArray(), XlsxContentType, "ListadoCotizacionesPresupuestos.xlsx");
        }

        public ActionResult DescargarReporteFormatoPDF()
        {
            // Seleccionar las columnas a exportar
            var results = PrefacturasSAFIEntity.ListadoCodigosPrefacturar();

            var list = Reportes.SerializeToJSON(results);
            return Content(list, "application/json");
        }

        public ActionResult DescargarReporteFormatoCSV()
        {
            var comlumHeadrs = new string[]
            {
                "CODIGODECOTIZACION",
                "DETALLE",
                "RUCCLIENTE",
                "CLIENTE",
                "CORREO",
                "EJECUTIVO",
                "SUBTOTAL"
            };

            var listado = (from item in PrefacturasSAFIEntity.ListadoCodigosPrefacturar()
            select new object[]
                           {
                                item.CodigoCotizacion, 
                                item.Detalle,
                                item.RUC,
                                item.Correo,
                                item.Ejecutivo,
                                item.Valor 
                           }).ToList();

            // Build the file content
            var employeecsv = new StringBuilder();
            listado.ForEach(line =>
            {
                employeecsv.AppendLine(string.Join(",", line));
            });

            byte[] buffer = Encoding.Default.GetBytes($"{string.Join(",", comlumHeadrs)}\r\n{employeecsv.ToString()}");
            return File(buffer, "text/csv", $"ListadoCotizacionesPresupuestos.csv");
        }
           
    }
}

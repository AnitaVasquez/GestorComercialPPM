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
    public class FacturaPTOPController : BaseAppController
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
            ViewBag.NombreListado = Etiquetas.TituloGridFacturaPTOP;
            //Controlar permisos
            var user = ViewData["usuario"] = System.Web.HttpContext.Current.Session["usuario"];
            var usuario = int.Parse(user.ToString());
            string nombreControlador = ControllerContext.RouteData.Values["controller"].ToString();
            ViewBag.NombreControlador = nombreControlador;

            ViewBag.AccionesUsuario = ManejoPermisosEntity.ListadoAccionesCatalogoUsuario(usuario, nombreControlador);

            //Obtener Acciones del controlador
            ViewBag.AccionesControlador = GetMetodosControlador(nombreControlador);

            //Búsqueda
            var listado = FacturaPTOPEntity.ListarListadoPTOPSAFI();

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
         
        public ActionResult DescargarReporteFormatoExcel()
        {
            var collection = FacturaPTOPEntity.ListarListadoPTOPSAFI();

            ExcelPackage package = new ExcelPackage();
            var workSheet = package.Workbook.Worksheets.Add("Listado Facturas");
            workSheet.TabColor = System.Drawing.Color.Black;
            workSheet.DefaultRowHeight = 12;

            List<string> columnas = new List<string>() {
                "ID",
                "CLIENTE",
                "IDENTIFICACIÓN",
                "DIRECCIÓN",
                "TELÉFONO",
                "CORREO",
                "AÑO",
                "MES",
                "CANTIDAD",
                "PRECIO UNITARIO",
                "DESCUENTO",
                "SUBTOTAL",
                "IVA",
                "TOTAL",
                "FECHA FACTURACIÓN",
                "DETALLE",
                "# FACTURA",
                "# NOTA CRÉDITO",
                "ESTADO FACTURA",
                "FACTURADO SAFI"
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
                workSheet.Cells[recordIndex, 1].Value = item.ID;
                workSheet.Cells[recordIndex, 2].Value = item.Cliente;
                workSheet.Cells[recordIndex, 3].Value = item.Identificacion;
                workSheet.Cells[recordIndex, 4].Value = item.Direccion;
                workSheet.Cells[recordIndex, 5].Value = item.Telefono;
                workSheet.Cells[recordIndex, 6].Value = item.Correos;
                workSheet.Cells[recordIndex, 7].Value = item.anio;
                workSheet.Cells[recordIndex, 8].Value = item.mes;
                workSheet.Cells[recordIndex, 9].Value = item.cantidad;
                workSheet.Cells[recordIndex, 10].Value = item.precio_unitario;
                workSheet.Cells[recordIndex, 10].Style.Numberformat.Format = "###,##0.00";
                workSheet.Cells[recordIndex, 11].Value = item.descuento;
                workSheet.Cells[recordIndex, 11].Style.Numberformat.Format = "###,##0.00";
                workSheet.Cells[recordIndex, 12].Value = item.subtotal;
                workSheet.Cells[recordIndex, 12].Style.Numberformat.Format = "###,##0.00";
                workSheet.Cells[recordIndex, 13].Value = item.iva;
                workSheet.Cells[recordIndex, 13].Style.Numberformat.Format = "###,##0.00";
                workSheet.Cells[recordIndex, 14].Value = item.total;
                workSheet.Cells[recordIndex, 14].Style.Numberformat.Format = "###,##0.00";
                workSheet.Cells[recordIndex, 15].Value = item.fecha_factura.Value.ToString("yyyy/MM/dd HH:mm:ss"); 
                workSheet.Cells[recordIndex, 16].Value = item.detalle;
                workSheet.Cells[recordIndex, 17].Value = item.numero_factura;
                workSheet.Cells[recordIndex, 18].Value = item.numero_nota_credito;
                workSheet.Cells[recordIndex, 19].Value = item.EstadoFactura;
                workSheet.Cells[recordIndex, 20].Value = item.FacturadoSAFI;
                recordIndex++;
            }

            for (int i = 1; i <= columnas.Count; i++)
            {
                workSheet.Column(i).AutoFit();
                workSheet.Column(i).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            }


            return File(package.GetAsByteArray(), XlsxContentType, "ListadoFacturas.xlsx");
        }

        public ActionResult DescargarReporteFormatoPDF()
        {
            // Seleccionar las columnas a exportar
            var results = FacturaPTOPEntity.ListarListadoPTOPSAFI();

            var list = Reportes.SerializeToJSON(results);
            return Content(list, "application/json");
        }

        public ActionResult DescargarReporteFormatoCSV()
        {
            var comlumHeadrs = new string[]
            {
                "ID",
                "CLIENTE",
                "IDENTIFICACIÓN",
                "DIRECCIÓN",
                "TELÉFONO",
                "CORREO",
                "AÑO",
                "MES",
                "CANTIDAD",
                "PRECIO UNITARIO",
                "DESCUENTO",
                "SUBTOTAL",
                "IVA",
                "TOTAL",
                "FECHA FACTURACIÓN",
                "DETALLE",
                "# FACTURA",
                "# NOTA CRÉDITO",
                "ESTADO FACTURA",
                "FACTURADO SAFI"
            };

            var listado = (from item in FacturaPTOPEntity.ListarListadoPTOPSAFI()
            select new object[]
                           {
                                item.ID,
                                item.Cliente,
                                item.Identificacion,
                                item.Direccion,
                                item.Telefono,
                                item.Correos,
                                item.anio,
                                item.mes,
                                item.cantidad,
                                item.precio_unitario,
                                item.descuento,
                                item.subtotal,
                                item.iva,
                                item.total,
                                item.fecha_factura.Value.ToString("yyyy/MM/dd HH:mm:ss"),
                                item.detalle,
                                item.numero_factura,
                                item.numero_factura,
                                item.EstadoFactura,
                                item.FacturadoSAFI
                           }).ToList();

            // Build the file content
            var employeecsv = new StringBuilder();
            listado.ForEach(line =>
            {
                employeecsv.AppendLine(string.Join(",", line));
            });

            byte[] buffer = Encoding.Default.GetBytes($"{string.Join(",", comlumHeadrs)}\r\n{employeecsv.ToString()}");
            return File(buffer, "text/csv", $"ListadoFacturas.csv");
        }

        public ActionResult _RegistrarFactura(int? id)
        {
            ViewBag.TituloModal = "Registro de Facturas";

            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }

            FacturaPTOPSAFI factura = FacturaPTOPEntity.ConsultarFactura(id.Value).FirstOrDefault();

            if (factura == null)
            {
                return HttpNotFound();
            }

            return PartialView(factura);
        }

        public ActionResult _AnularDocumento(int? id)
        {
            ViewBag.TituloModal = "Anular Documento";

            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }

            FacturaPTOPSAFI factura = FacturaPTOPEntity.ConsultarFactura(id.Value).FirstOrDefault();

            if (factura == null)
            {
                return HttpNotFound();
            }

            return PartialView(factura);
        }

        [HttpPost]
        public ActionResult RegistrarFactura(FacturaPTOPSAFI factura)
        {
            try
            {
                RespuestaTransaccion resultado = FacturaPTOPEntity.ActualizarFactura(factura);
                return Json(new { Resultado = resultado }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                return Json(new { Resultado = new RespuestaTransaccion { Estado = false, Respuesta = ex.Message } }, JsonRequestBehavior.AllowGet);
            }
        }

        [HttpPost]
        public ActionResult AnularDocumento(FacturaPTOPSAFI factura, int tipo_documento, int metodo, int accion, string motivo)
        {
            try
            {
                RespuestaTransaccion resultado = new RespuestaTransaccion();

                //anulacion de factura
                if (tipo_documento  == 1)
                {
                    //anulacion normal
                    if (metodo == 1)
                    {
                        resultado = FacturaPTOPEntity.AnularFacturaNormal(factura, accion); 
                    }

                    //anulacion nota credito
                    if (metodo == 2)
                    {
                        //Anulacio de las facturas por accion
                        resultado = FacturaPTOPEntity.AnularFacturaNormal(factura, accion);

                        //Generar la nota de Credito
                        resultado = NotaCreditoPTOPEntity.CrearNotaCredito(factura.id_factura_PTOP, motivo);
                    }
                }

                //anulacion de nota de credito
                if (tipo_documento == 2)
                {
                    resultado = NotaCreditoPTOPEntity.InactivarNotaCreditoPTOP(factura.id_factura_PTOP);
                }

                //RespuestaTransaccion resultado = FacturaPTOPEntity.AnularFacturaNormal(factura, accion);
                return Json(new { Resultado = resultado }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                return Json(new { Resultado = new RespuestaTransaccion { Estado = false, Respuesta = ex.Message } }, JsonRequestBehavior.AllowGet);
            }
        }
    }
}

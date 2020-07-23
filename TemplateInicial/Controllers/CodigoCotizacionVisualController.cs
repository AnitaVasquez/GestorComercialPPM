using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Net;
using System.Web.Mvc;
using GestionPPM.Entidades.Metodos;
using GestionPPM.Entidades.Modelo;
using GestionPPM.Repositorios;
using System.Threading.Tasks;
using OfficeOpenXml;
using System.Text;
using Omu.Awem.Helpers;
using OfficeOpenXml.Style;
using System.IO;
using Seguridad.Helper;
using System.Web;
using System.Configuration;

namespace TemplateInicial.Controllers
{
    [Autenticado]
    public class CodigoCotizacionVisualController : BaseAppController
    {
       
        // GET: CodigoCotizaciónVisual
        private const string XlsxContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
        private int rolID = Convert.ToInt32(System.Web.HttpContext.Current.Session["rolID"]);

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
            ViewBag.NombreListado = Etiquetas.TituloGridCodigoCotizacion;

            //Controlar permisos
            var user = ViewData["usuario"] = System.Web.HttpContext.Current.Session["usuario"];
            var usuario = int.Parse(user.ToString());
            string nombreControlador = ControllerContext.RouteData.Values["controller"].ToString();
            ViewBag.NombreControlador = nombreControlador;


            ViewBag.AccionesUsuario = ManejoPermisosEntity.ListadoAccionesCatalogoUsuario(usuario, nombreControlador);

            //Obtener Acciones del controlador

            ViewBag.AccionesControlador = GetMetodosControlador(nombreControlador);
            //List<string> NavItems = new List<string>();

            //ReflectedControllerDescriptor controllerDesc = new ReflectedControllerDescriptor(this.GetType());

            //foreach (ActionDescriptor action in controllerDesc.GetCanonicalActions())
            //{
            //    bool validAction = true;

            //    object[] attributes = action.GetCustomAttributes(false);

            //    foreach (object filter in attributes)
            //    {

            //        if (filter is HttpPostAttribute || filter is ChildActionOnlyAttribute)
            //        {
            //            validAction = false;
            //            break;
            //        }
            //    }
            //    if (validAction)
            //        NavItems.Add(action.ActionName);
            //}

            //ViewBag.AccionesControlador = NavItems;

            //Búsqueda

            var listado = CodigoCotizacionEntity.ListarCodigoCotizacion();

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
            //Seleccionar las columnas a exportar
            var collection = CodigoCotizacionEntity.ListarCodigoCotizacion();
            var package = new ExcelPackage();

            package = Reportes.ExportarExcel(collection.Cast<object>().ToList(), "Codigo Cotizacion");
            return File(package.GetAsByteArray(), XlsxContentType, "ListadoCodigoCotizacion.xlsx");
        }

        public ActionResult DescargarReporteFormatoPDF()
        {
            // Seleccionar las columnas a exportar
            var results = CodigoCotizacionEntity.ListarCodigoCotizacion();

            var list = Reportes.SerializeToJSON(results);
            return Content(list, "application/json");
        }

        public ActionResult DescargarReporteResumidoFormatoCSV()
        {
            var comlumHeadrs = new string[]
            {
                "Id",
                "Fecha de Cotización",
                "Código de Cotización",
                "Estatus Código",
                "Responsable",
                "Cliente",
                "Nombre del Proyecto",
                "Ejecutivo",
                "Tipo FEE",
                "Tipo Proyecto",
                "Fases",
                "Etapa PTOP",
                "Total"
            };

            var listado = (from item in CodigoCotizacionEntity.ListarCodigoCotizacion()
                           select new object[]
                           {
                                            item.id_codigo_cotizacion,
                                            item.fecha_cotizacion,
                                            $"{item.codigo_cotizacion}",
                                            $"{item.EstatusCodigo}",
                                            $"{item.Responsable}",
                                            $"{item.nombre_comercial_cliente}",
                                            $"{item.nombre_proyecto}",
                                            $"{item.Ejecutivo}",
                                            $"{item.TipoFEE}",
                                            $"{item.TipoProyecto}",
                                            $"{item.EtapaCliente}",
                                            $"{item.TipoEtapaPTOP}",
                                            $"\"{(item.TotalSubLineaNegocio) }\"", //Escaping ","
                           }).ToList();

            // Build the file content
            var employeecsv = new StringBuilder();
            listado.ForEach(line =>
            {
                employeecsv.AppendLine(string.Join("¬", line));
            });

            byte[] buffer = Encoding.Default.GetBytes($"{string.Join("¬", comlumHeadrs)}\r\n{employeecsv.ToString()}");
            return File(buffer, "text/csv", $"CodigoCotizacion.csv");
        }

        public ActionResult DescargarReporteDetalladoFormatoCSV()
        {
            var comlumHeadrs = new string[]
            {
                "Id",
                "Fecha de Cotización",
                "Año",
                "Mes",
                "Código de Cotización",
                "Estatus Código",
                "Responsable",
                "Cliente",
                "Tipo de Cliente",
                "Nombre del Proyecto",
                "Descripción del Proyecto",
                "Ejecutivo",
                "Tipo de Requerimiento",
                "Tipo FEE",
                "Tipo Proyecto",
                "Fases",
                "Etapa PTOP",
                "Aplica Contrato",
                "Pagos Parciales",
                "Pago 1",
                "Pago 2",
                "Pago 3",
                "Pago 4",
                "Total",
                "Facturable",
                "Creación SAFI"
            };

            var listado = (from item in CodigoCotizacionEntity.ListarCodigoCotizacion()
                           select new object[]
                           {
                                            item.id_codigo_cotizacion,
                                            item.fecha_cotizacion,
                                            item.Anio_FechaCotizacion,
                                            item.Mes_FechaCotizacion,
                                            $"{item.codigo_cotizacion}",
                                            $"{item.EstatusCodigo}",
                                            $"{item.Responsable}",
                                            $"{item.nombre_comercial_cliente}",
                                            $"{item.TipoCliente}",
                                            $"{item.nombre_proyecto}",
                                            $"{item.descripcion_proyecto}",
                                            $"{item.Ejecutivo}",
                                            $"{item.TipoRequerido}",
                                            $"{item.TipoFEE}",
                                            $"{item.TipoProyecto}",
                                            $"{item.EtapaCliente}",
                                            $"{item.TipoEtapaPTOP}",
                                            $"{item.AplicaContrato}",
                                            $"{item.forma_pago}",
                                            $"{item.forma_pago_1}",
                                            $"{item.forma_pago_2}",
                                            $"{item.forma_pago_3}",
                                            $"{item.forma_pago_4}",
                                            $"{item.TotalSubLineaNegocio}",
                                            $"{item.Facturable}",
                                            $"\"{(item.CreacionSAFI) }\"", //Escaping ","
                           }).ToList();

            // Build the file content
            var employeecsv = new StringBuilder();
            listado.ForEach(line =>
            {
                employeecsv.AppendLine(string.Join("¬", line));
            });

            byte[] buffer = Encoding.Default.GetBytes($"{string.Join("¬", comlumHeadrs)}\r\n{employeecsv.ToString()}");

            return File(buffer, "text/csv", $"CodigoCotizacion.csv");
        }



        #region Reportes Personalizados

        public ActionResult DescargarReporteResumidoFormatoExcel()
        {
            var collection = CodigoCotizacionEntity.ListarCodigoCotizacion();

            ExcelPackage package = new ExcelPackage();
            var workSheet = package.Workbook.Worksheets.Add("Listado");
            workSheet.TabColor = System.Drawing.Color.Black;
            workSheet.DefaultRowHeight = 12;

            List<string> columnas = new List<string>() {
                "Id",
                "Fecha de Cotización",
                "Código de Cotización",
                "Estatus Código",
                "Responsable",
                "Cliente",
                "Nombre del Proyecto",
                "Ejecutivo",
                "Tipo FEE",
                "Tipo Proyecto",
                "Fases",
                "Etapa PTOP",
                "Total"};

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
                workSheet.Cells[recordIndex, 1].Value = item.id_codigo_cotizacion;
                workSheet.Cells[recordIndex, 2].Value = item.fecha_cotizacion.Value.ToString("yyyy/MM/dd");
                workSheet.Cells[recordIndex, 3].Value = item.codigo_cotizacion;
                workSheet.Cells[recordIndex, 4].Value = item.EstatusCodigo;
                workSheet.Cells[recordIndex, 5].Value = item.Responsable;
                workSheet.Cells[recordIndex, 6].Value = item.nombre_comercial_cliente;
                workSheet.Cells[recordIndex, 7].Value = item.nombre_proyecto;
                workSheet.Cells[recordIndex, 8].Value = item.Ejecutivo;
                workSheet.Cells[recordIndex, 9].Value = item.TipoFEE;
                workSheet.Cells[recordIndex, 10].Value = item.TipoProyecto;
                workSheet.Cells[recordIndex, 11].Value = item.EtapaCliente;
                workSheet.Cells[recordIndex, 12].Value = item.TipoEtapaPTOP;
                workSheet.Cells[recordIndex, 13].Value = item.TotalSubLineaNegocio;
                workSheet.Cells[recordIndex, 13].Style.Numberformat.Format = "###,##0.00";

                recordIndex++;
            }

            for (int i = 1; i <= columnas.Count; i++)
            {
                workSheet.Column(i).AutoFit();
                workSheet.Column(i).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                if (i == columnas.Count)
                {
                    workSheet.Column(i).Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                }
            }


            return File(package.GetAsByteArray(), XlsxContentType, "ListadoCodigoCotizacion.xlsx");
        }

        public ActionResult DescargarReporteDetalladoFormatoExcel()
        {
            var collection = CodigoCotizacionEntity.ListarCodigoCotizacion();

            //==========================================================================================
            //****************************DEFINICION DE LAS HOJAS*************************************//
            //==========================================================================================
            ExcelPackage ExcelPkg = new ExcelPackage();

            ExcelWorksheet Listado = ExcelPkg.Workbook.Worksheets.Add("Lista Códigos");
            Listado.TabColor = System.Drawing.Color.Black;
            Listado.DefaultRowHeight = 12;

            ExcelWorksheet MigracionCabecera = ExcelPkg.Workbook.Worksheets.Add("Códigos Completos");
            MigracionCabecera.TabColor = System.Drawing.Color.Black;
            MigracionCabecera.DefaultRowHeight = 12;

            ExcelWorksheet MigracionDetalle = ExcelPkg.Workbook.Worksheets.Add("Detalle Códigos");
            MigracionDetalle.TabColor = System.Drawing.Color.Black;
            MigracionDetalle.DefaultRowHeight = 12;


            List<string> columnas = new List<string>()
            {
                "Id",
                "Fecha de Cotización",
                "Año",
                "Mes",
                "Código de Cotización",
                "Estatus Código",
                "Responsable",
                "Cliente",
                "Tipo de Cliente",
                "Nombre del Proyecto",
                "Descripción del Proyecto",
                "Ejecutivo",
                "Tipo de Requerimiento",
                "Tipo FEE",
                "Tipo Proyecto",
                "Fases",
                "Etapa PTOP",
                "Aplica Contrato",
                "Pagos Parciales",
                "Pago 1",
                "Pago 2",
                "Pago 3",
                "Pago 4",
                "Total",
                "Facturable",
                "Creación SAFI"
            };

            Listado.Row(1).Height = 20;
            Listado.Row(1).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            Listado.Row(1).Style.Font.Bold = true;

            int contador = 0;
            for (int i = 1; i <= columnas.Count; i++)
            {
                Listado.Cells[1, i].Value = columnas.ElementAt(contador);
                contador++;
            }

            //Body of table  
            int recordIndex = 2;

            foreach (var item in collection)
            {
                Listado.Cells[recordIndex, 1].Value = item.id_codigo_cotizacion;
                Listado.Cells[recordIndex, 2].Value = item.fecha_cotizacion.Value.ToString("yyyy/MM/dd");
                Listado.Cells[recordIndex, 3].Value = item.Anio_FechaCotizacion;
                Listado.Cells[recordIndex, 4].Value = item.Mes_FechaCotizacion;
                Listado.Cells[recordIndex, 5].Value = item.codigo_cotizacion;
                Listado.Cells[recordIndex, 6].Value = item.EstatusCodigo;
                Listado.Cells[recordIndex, 7].Value = item.Responsable;
                Listado.Cells[recordIndex, 8].Value = item.nombre_comercial_cliente;
                Listado.Cells[recordIndex, 9].Value = item.TipoCliente;
                Listado.Cells[recordIndex, 10].Value = item.nombre_proyecto;
                Listado.Cells[recordIndex, 11].Value = item.descripcion_proyecto;
                Listado.Cells[recordIndex, 12].Value = item.Ejecutivo;
                Listado.Cells[recordIndex, 13].Value = item.TipoRequerido;
                Listado.Cells[recordIndex, 14].Value = item.TipoFEE;
                Listado.Cells[recordIndex, 15].Value = item.TipoProyecto;
                Listado.Cells[recordIndex, 16].Value = item.EtapaCliente;
                Listado.Cells[recordIndex, 17].Value = item.TipoEtapaPTOP;
                Listado.Cells[recordIndex, 18].Value = item.AplicaContrato;
                Listado.Cells[recordIndex, 19].Value = item.forma_pago;
                Listado.Cells[recordIndex, 20].Value = item.forma_pago_1;
                Listado.Cells[recordIndex, 20].Style.Numberformat.Format = "###,##0.00";
                Listado.Cells[recordIndex, 21].Value = item.forma_pago_2;
                Listado.Cells[recordIndex, 21].Style.Numberformat.Format = "###,##0.00";
                Listado.Cells[recordIndex, 22].Value = item.forma_pago_3;
                Listado.Cells[recordIndex, 22].Style.Numberformat.Format = "###,##0.00";
                Listado.Cells[recordIndex, 23].Value = item.forma_pago_4;
                Listado.Cells[recordIndex, 23].Style.Numberformat.Format = "###,##0.00";
                Listado.Cells[recordIndex, 24].Value = item.TotalSubLineaNegocio;
                Listado.Cells[recordIndex, 24].Style.Numberformat.Format = "###,##0.00";
                Listado.Cells[recordIndex, 25].Value = item.Facturable;
                Listado.Cells[recordIndex, 26].Value = item.CreacionSAFI;
                recordIndex++;
            }

            for (int i = 1; i <= columnas.Count; i++)
            {
                Listado.Column(i).AutoFit();
                Listado.Column(i).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                if (i == columnas.Count - 2)
                {
                    Listado.Column(i).Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                }
            }

            var collectionDetalle = CodigoCotizacionEntity.ListarCodigoCotizacion();

            List<string> columnasDetalle = new List<string>()
            {
                "Id",
                "Código Cotización",
                "Fecha Cotización",
                "Responsable",
                "Estatus Código",
                "Cliente",
                "Ejecutivo",
                "Tipo Requerimiento",
                "Tipo Intermediario",
                "Tipo Proyecto",
                "Dimension Proyecto",
                "Aplica Contrato",
                "Pagos_Parciales",
                "Pago 1",
                "Pago 2",
                "Pago 3",
                "Pago 4",
                "Etapa del Cliente",
                "Etapa General",
                "Estatus Detallado",
                "Estatus General",
                "Tipo Producto PtoP",
                "Tipo Plan",
                "Tipo Tarifa",
                "Tipo Migración",
                "Tipo Etapa PtoP",
                "Tipo Subsidio",
                "Nombre del Proyecto",
                "Descripción del Proyecto",
                "Tipo FEE",
                "Creación SAFI",
                "Facturable",
                "Area o Departamento",
                "País",
                "Ciudad",
                "Dirección",
                "Tipo Cliente",
                "Tipo CRM",
                "Referido",
                "Estado",
            };

            MigracionCabecera.Row(1).Height = 20;
            MigracionCabecera.Row(1).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            MigracionCabecera.Row(1).Style.Font.Bold = true;

            contador = 0;
            for (int i = 1; i <= columnasDetalle.Count; i++)
            {
                MigracionCabecera.Cells[1, i].Value = columnasDetalle.ElementAt(contador);
                contador++;
            }

            //Body of table  
            recordIndex = 2;

            foreach (var item in collectionDetalle)
            {
                MigracionCabecera.Cells[recordIndex, 1].Value = item.id_codigo_cotizacion;
                MigracionCabecera.Cells[recordIndex, 2].Value = item.codigo_cotizacion;
                MigracionCabecera.Cells[recordIndex, 3].Value = item.fecha_cotizacion.Value.ToString("yyyy/MM/dd");
                MigracionCabecera.Cells[recordIndex, 4].Value = item.Responsable;
                MigracionCabecera.Cells[recordIndex, 5].Value = item.EstatusCodigo;
                MigracionCabecera.Cells[recordIndex, 6].Value = item.razon_social_cliente;
                MigracionCabecera.Cells[recordIndex, 7].Value = item.Ejecutivo;
                MigracionCabecera.Cells[recordIndex, 8].Value = item.TipoRequerido;
                MigracionCabecera.Cells[recordIndex, 9].Value = item.TipoIntermediario;
                MigracionCabecera.Cells[recordIndex, 10].Value = item.TipoProyecto;
                MigracionCabecera.Cells[recordIndex, 11].Value = item.DimensionProyecto;
                MigracionCabecera.Cells[recordIndex, 12].Value = item.AplicaContrato;
                MigracionCabecera.Cells[recordIndex, 13].Value = item.forma_pago;
                MigracionCabecera.Cells[recordIndex, 14].Value = item.forma_pago_1;
                MigracionCabecera.Cells[recordIndex, 15].Value = item.forma_pago_2;
                MigracionCabecera.Cells[recordIndex, 16].Value = item.forma_pago_3;
                MigracionCabecera.Cells[recordIndex, 17].Value = item.forma_pago_4;
                MigracionCabecera.Cells[recordIndex, 18].Value = item.EtapaCliente;
                MigracionCabecera.Cells[recordIndex, 19].Value = item.EtapaGeneral;
                MigracionCabecera.Cells[recordIndex, 20].Value = item.EstatusDetallado;
                MigracionCabecera.Cells[recordIndex, 21].Value = item.EstatusGeneral;
                MigracionCabecera.Cells[recordIndex, 22].Value = item.TipoProductoPTOP;
                MigracionCabecera.Cells[recordIndex, 23].Value = item.TipoPlan;
                MigracionCabecera.Cells[recordIndex, 24].Value = item.TipoTarifa;
                MigracionCabecera.Cells[recordIndex, 25].Value = item.TipoMigracion;
                MigracionCabecera.Cells[recordIndex, 26].Value = item.TipoEtapaPTOP;
                MigracionCabecera.Cells[recordIndex, 27].Value = item.TipoSubsidio;
                MigracionCabecera.Cells[recordIndex, 28].Value = item.nombre_proyecto;
                MigracionCabecera.Cells[recordIndex, 29].Value = item.descripcion_proyecto;
                MigracionCabecera.Cells[recordIndex, 30].Value = item.TipoFEE;
                MigracionCabecera.Cells[recordIndex, 31].Value = item.CreacionSAFI;
                MigracionCabecera.Cells[recordIndex, 32].Value = item.Facturable;
                MigracionCabecera.Cells[recordIndex, 33].Value = item.AreaDepartamento;
                MigracionCabecera.Cells[recordIndex, 34].Value = item.Pais;
                MigracionCabecera.Cells[recordIndex, 35].Value = item.Ciudad;
                MigracionCabecera.Cells[recordIndex, 36].Value = item.direccion;
                MigracionCabecera.Cells[recordIndex, 37].Value = item.TipoCliente;
                MigracionCabecera.Cells[recordIndex, 38].Value = item.TipoZoho;
                MigracionCabecera.Cells[recordIndex, 39].Value = item.Referido;
                MigracionCabecera.Cells[recordIndex, 40].Value = item.Estado;

                recordIndex++;
            }

            for (int i = 1; i <= columnasDetalle.Count; i++)
            {
                MigracionCabecera.Column(i).AutoFit();
                MigracionCabecera.Column(i).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            }

            var collectionDetalleSublinea = CodigoCotizacionEntity.ListarSublineaCodigoCotizacion();

            List<string> columnasDetalleSublinea = new List<string>()
            {
                "Id",
                "Código Cotización",
                "Sublínea Negocio",
                "Valor"
            };

            MigracionDetalle.Row(1).Height = 20;
            MigracionDetalle.Row(1).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            MigracionDetalle.Row(1).Style.Font.Bold = true;

            contador = 0;
            for (int i = 1; i <= columnasDetalleSublinea.Count; i++)
            {
                MigracionDetalle.Cells[1, i].Value = columnasDetalleSublinea.ElementAt(contador);
                contador++;
            }

            //Body of table  
            recordIndex = 2;

            foreach (var item in collectionDetalleSublinea)
            {
                MigracionDetalle.Cells[recordIndex, 1].Value = item.IdSublineaNegocioCotizacion;
                MigracionDetalle.Cells[recordIndex, 2].Value = item.codigo_cotizacion;
                MigracionDetalle.Cells[recordIndex, 3].Value = item.TextoCatalogoSublineaNegocio;
                MigracionDetalle.Cells[recordIndex, 4].Value = item.Valor;
                MigracionDetalle.Cells[recordIndex, 4].Style.Numberformat.Format = "###,##0.00";

                recordIndex++;
            }

            for (int i = 1; i <= columnasDetalleSublinea.Count; i++)
            {
                MigracionDetalle.Column(i).AutoFit();
                MigracionDetalle.Column(i).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                if (i == columnasDetalleSublinea.Count)
                {
                    MigracionDetalle.Column(i).Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                }
            }

            return File(ExcelPkg.GetAsByteArray(), XlsxContentType, "ListadoCodigoCotizacion.xlsx");
        }
        #endregion
    }

}
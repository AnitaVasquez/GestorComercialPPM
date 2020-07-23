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
    public class CargaPTOPController : BaseAppController
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
            ViewBag.NombreListado = Etiquetas.TituloGridCargaPTOP;
            //Controlar permisos
            var user = ViewData["usuario"] = System.Web.HttpContext.Current.Session["usuario"];
            var usuario = int.Parse(user.ToString());
            string nombreControlador = ControllerContext.RouteData.Values["controller"].ToString();
            ViewBag.NombreControlador = nombreControlador;

            ViewBag.AccionesUsuario = ManejoPermisosEntity.ListadoAccionesCatalogoUsuario(usuario, nombreControlador);

            //Obtener Acciones del controlador
            ViewBag.AccionesControlador = GetMetodosControlador(nombreControlador);

            //Búsqueda

            var listado = CargaPTOPEntity.ListarCargaPTOP();

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
         

        // GET: Edit/5
        public ActionResult Edit(int? id)
        {
            //Listado Comercios
            var Comercios = ComercioEntity.ListarComerciosPTOP();
            ViewBag.ListadoComercios = Comercios;

            //Listado Catalogo Certificacion 
            var Certificacion = CatalogoEntity.ObtenerListadoCatalogosByCodigo("FCR-01");
            ViewBag.ListadoCertificacion = Certificacion;

            //Listado Catalogo Facturacion  Mensual
            var Mensual = CatalogoEntity.ObtenerListadoCatalogosByCodigo("FMS-01");
            ViewBag.ListadoMensual = Mensual;
             

            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }

            CargaPTOP carga = CargaPTOPEntity.ConsultarCargaPTOP(id.Value);  

            if (carga == null)
            {
                return HttpNotFound();
            }
            
            //Valor Certificacion
            ViewBag.Certificacion = carga.valor_certificacion.HasValue ? carga.valor_certificacion : 0;

            return View(carga);
        } 
        
        [HttpPost]
        public ActionResult Edit(CargaPTOP carga)
        {
            try
            {
                 
                RespuestaTransaccion resultado = CargaPTOPEntity.ActualizarCargaPTOP(carga);
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
            RespuestaTransaccion resultado = CargaPTOPEntity.EliminarCargaPTOP(id);// await db.Cabecera.FindAsync(id);

            //return RedirectToAction("Index");
            return Json(new { Resultado = resultado }, JsonRequestBehavior.AllowGet);

        } 
         
        public ActionResult DescargarReporteFormatoExcel()
        {
            var collection = CargaPTOPEntity.ListarCargaPTOPExcel();

            ExcelPackage package = new ExcelPackage();
            var workSheet = package.Workbook.Worksheets.Add("Listado Carga PTOP");
            workSheet.TabColor = System.Drawing.Color.Black;
            workSheet.DefaultRowHeight = 12;

            List<string> columnas = new List<string>() {
                "ID",
                "COMERCIO",
                "RUC CLIENTE",
                "CORREOS",
                "ESTADO",
                "FECHA SALIDA A PRODUCCIÓN",
                "FECHA INACTIVO",
                "PLAN",
                "AGRUPACIÓN",
                "ESTATUS CONTRATO",
                "FACTURABLE CERTIFICACIÓN",
                "FACTURABLE MENSUAL",
                "MES",
                "AÑO",
                "DETALLE",
                "VALOR CERTIFICACIÓN",
                "TRX APROBADAS",
                "TRX RECHAZADAS",
                "MONTO APROBADO",
                "MONTO RECHAZADO",
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
                workSheet.Cells[recordIndex, 1].Value = item.Codigo;
                workSheet.Cells[recordIndex, 2].Value = item.NombreComercio;
                workSheet.Cells[recordIndex, 3].Value = item.RucCliente;
                workSheet.Cells[recordIndex, 4].Value = item.Correos;
                workSheet.Cells[recordIndex, 5].Value = item.estado_comercio;
                workSheet.Cells[recordIndex, 6].Value = item.fecha_salida_produccion;
                workSheet.Cells[recordIndex, 7].Value = item.fecha_inactivo;
                workSheet.Cells[recordIndex, 8].Value = item.NombrePlan;
                workSheet.Cells[recordIndex, 9].Value = item.Agrupacion;
                workSheet.Cells[recordIndex, 10].Value = item.EstatusContrato;
                workSheet.Cells[recordIndex, 11].Value = item.FacturableCertificacion;
                workSheet.Cells[recordIndex, 12].Value = item.FacturableMensual;
                workSheet.Cells[recordIndex, 13].Value = item.Mes;
                workSheet.Cells[recordIndex, 14].Value = item.Anio;
                workSheet.Cells[recordIndex, 15].Value = item.Detalle;
                workSheet.Cells[recordIndex, 16].Value = item.ValorCertificacion;
                workSheet.Cells[recordIndex, 16].Style.Numberformat.Format = "###,##0.00";
                workSheet.Cells[recordIndex, 17].Value = item.TransaccionesAprobadas;
                workSheet.Cells[recordIndex, 18].Value = item.TransaccionesRechazadas;
                workSheet.Cells[recordIndex, 19].Value = item.MontoVendido;
                workSheet.Cells[recordIndex, 19].Style.Numberformat.Format = "###,##0.00";
                workSheet.Cells[recordIndex, 20].Value = item.MontoRechazado;
                workSheet.Cells[recordIndex, 20].Style.Numberformat.Format = "###,##0.00";
                workSheet.Cells[recordIndex, 21].Value = item.Cantidad;
                workSheet.Cells[recordIndex, 22].Value = item.PrecioUnitario;
                workSheet.Cells[recordIndex, 22].Style.Numberformat.Format = "###,##0.00";
                workSheet.Cells[recordIndex, 23].Value = item.Descuento_Factura;
                workSheet.Cells[recordIndex, 23].Style.Numberformat.Format = "###,##0.00";
                workSheet.Cells[recordIndex, 24].Value = item.Subtotal;
                workSheet.Cells[recordIndex, 24].Style.Numberformat.Format = "###,##0.00";
                workSheet.Cells[recordIndex, 25].Value = item.Iva;
                workSheet.Cells[recordIndex, 25].Style.Numberformat.Format = "###,##0.00";
                workSheet.Cells[recordIndex, 26].Value = item.Total;
                workSheet.Cells[recordIndex, 26].Style.Numberformat.Format = "###,##0.00";
                workSheet.Cells[recordIndex, 27].Value = item.FechaFactura.ToString();
                workSheet.Cells[recordIndex, 28].Value = item.Detalle;
                workSheet.Cells[recordIndex, 29].Value = item.numero_factura;
                workSheet.Cells[recordIndex, 30].Value = item.numero_nota_credito;
                workSheet.Cells[recordIndex, 31].Value = item.EstadoFactura;
                workSheet.Cells[recordIndex, 32].Value = item.FacturadoSAFI;
                recordIndex++;
            }

            for (int i = 1; i <= columnas.Count; i++)
            {
                workSheet.Column(i).AutoFit();
                workSheet.Column(i).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            }


            return File(package.GetAsByteArray(), XlsxContentType, "ListadoCargaPTOP.xlsx");
        }

        public ActionResult DescargarReporteFormatoPDF()
        {
            // Seleccionar las columnas a exportar
            var results = CargaPTOPEntity.ListarCargaPTOPExcel();

            var list = Reportes.SerializeToJSON(results);
            return Content(list, "application/json");
        }

        public ActionResult DescargarReporteFormatoCSV()
        {
            var comlumHeadrs = new string[]
            {
                "ID",
                "COMERCIO",
                "RUC CLIENTE",
                "CORREOS",
                "ESTADO",
                "FECHA SALIDA A PRODUCCIÓN",
                "FECHA INACTIVO",
                "PLAN",
                "AGRUPACIÓN",
                "ESTATUS CONTRATO",
                "FACTURABLE CERTIFICACIÓN",
                "FACTURABLE MENSUAL",
                "MES",
                "AÑO",
                "DETALLE",
                "VALOR CERTIFICACIÓN",
                "TRX APROBADAS",
                "TRX RECHAZADAS",
                "MONTO APROBADO",
                "MONTO RECHAZADO",
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

            var listado = (from item in CargaPTOPEntity.ListarCargaPTOPExcel()
                           select new object[]
                           {
                                item.Codigo, 
                                item.NombreComercio,
                                item.RucCliente,
                                item.fecha_salida_produccion,
                                item.fecha_inactivo,
                                item.NombrePlan,
                                item.EstatusContrato,
                                item.FacturableCertificacion,
                                item.FacturableMensual,
                                item.Mes,
                                item.Anio, 
                                item.ValorCertificacion,
                                item.TransaccionesAprobadas,
                                item.TransaccionesRechazadas,
                                item.MontoVendido,
                                item.MontoRechazado
                           }).ToList();

            // Build the file content
            var employeecsv = new StringBuilder();
            listado.ForEach(line =>
            {
                employeecsv.AppendLine(string.Join(",", line));
            });

            byte[] buffer = Encoding.Default.GetBytes($"{string.Join(",", comlumHeadrs)}\r\n{employeecsv.ToString()}");
            return File(buffer, "text/csv", $"ListadoCargaPTOP.csv");
        }
         

    #region Funcionalidad de carga Masiva
        public ActionResult CargarData()
        {
            List<CargaMasiva> errores = new List<CargaMasiva>();
            RespuestaTransaccion resultado = new RespuestaTransaccion { Estado = false };
            List<CargaPTOPExcel> listado = new List<CargaPTOPExcel>();
            string FileName = "";

            try
            {
                HttpFileCollectionBase files = Request.Files;
                for (int i = 0; i < files.Count; i++)
                {  
                    HttpPostedFileBase file = files[i];
                    string path = string.Empty;

                    // Checking for Internet Explorer    
                    if (Request.Browser.Browser.ToUpper() == "IE" || Request.Browser.Browser.ToUpper() == "INTERNETEXPLORER")
                    {
                        string[] testfiles = file.FileName.Split(new char[] { '\\' });
                        path = testfiles[testfiles.Length - 1];
                    }
                    else
                    {
                        path = file.FileName;
                        FileName = file.FileName;
                    }

                    bool directorio = Directory.Exists(Server.MapPath("~/CargasMasivas/"));

                    // En caso de que no exista el directorio, crearlo.
                    if (!directorio)
                        Directory.CreateDirectory(Server.MapPath("~/CargasMasivas/"));

                    // Get the complete folder path and store the file inside it.    
                    path = Path.Combine(Server.MapPath("~/CargasMasivas/"), path);

                    string extension = Path.GetExtension(path);
                    if (extension != ".xls" && extension != ".xlsx")
                    {
                        errores.Add(new CargaMasiva { Fila = 0, Columna = 0, Valor = "Ninguno", Error = "Formato inválido." });
                        return Json(new { Resultado = new RespuestaTransaccion { Estado = false, Respuesta = "Formato inválido." }, Errores = errores }, JsonRequestBehavior.AllowGet);
                    }

                    file.SaveAs(path);

                    errores = VerificarEstructuraExcel(path); // Devuelve los errores en la estructura

                    if (errores.Count == 0)
                    {
                        listado = ToEntidadHojaExcelList(path); // Convierte el documento a un Listado tipo entidad

                        //Validar que la columna de Numeración no se encuentre duplicada
                        var numeracionDuplicada = listado.GroupBy(x => x.id).Any(g => g.Count() > 1);

                        //Validar que en la data no haya incosistencia de pago mensual o certificacion  
                        var dataIncorrectaFactable = listado.Where( l => l.facturable_certificacion.ToUpper() == "SI" && l.facturable_mensual.ToUpper() =="SI");
                        var dataIncorrectaValorCertificacion = listado.Where(l => l.facturable_certificacion.ToUpper() == "SI" && l.valor_certificacion <= 0);
                        var dataIncorrectaValorNoCertificacion = listado.Where(l => l.facturable_certificacion.ToUpper() == "NO" && l.valor_certificacion > 0);
                        var dataIncorrectaValorMensual = listado.Where(l => l.facturable_mensual.ToUpper() == "SI" && l.valor_certificacion > 0);
                        var dataIncorrectaValorSiMensual = listado.Where(l => l.facturable_certificacion.ToUpper() == "SI" && l.facturable_mensual.ToUpper() != "NO");

                        if (dataIncorrectaFactable.Count() > 0 || dataIncorrectaValorSiMensual.Count() > 0)
                        {
                            errores.Add(new CargaMasiva { Fila = 0, Columna = 0, Valor = "Ninguno", Error = "Los datos cargados para facturar son incosistentes en facturable certificación y facturable mensual" });
                            return Json(new { Resultado = new RespuestaTransaccion { Estado = resultado.Estado, Respuesta = Mensajes.MensajeCargaMasivaFallida }, Errores = errores, Archivo = FileName }, JsonRequestBehavior.AllowGet);
                        }

                        if (dataIncorrectaValorCertificacion.Count() > 0)
                        {
                            errores.Add(new CargaMasiva { Fila = 0, Columna = 0, Valor = "Ninguno", Error = "No se puede emitir una factura por certificación con valor cero" });
                            return Json(new { Resultado = new RespuestaTransaccion { Estado = resultado.Estado, Respuesta = Mensajes.MensajeCargaMasivaFallida }, Errores = errores, Archivo = FileName }, JsonRequestBehavior.AllowGet);
                        }

                        if (dataIncorrectaValorNoCertificacion.Count() > 0 || dataIncorrectaValorMensual.Count() > 0)
                        {
                            errores.Add(new CargaMasiva { Fila = 0, Columna = 0, Valor = "Ninguno", Error = "No se puede emitir una factura mensual con valor de certificación" });
                            return Json(new { Resultado = new RespuestaTransaccion { Estado = resultado.Estado, Respuesta = Mensajes.MensajeCargaMasivaFallida }, Errores = errores, Archivo = FileName }, JsonRequestBehavior.AllowGet);
                        } 

                        if (numeracionDuplicada) {
                            errores.Add(new CargaMasiva { Fila = 0, Columna = 0, Valor = "Ninguno", Error = "Valores de la Columna de numeración 'ID' no pueden ser duplicados." });
                            return Json(new { Resultado = new RespuestaTransaccion { Estado = resultado.Estado, Respuesta = Mensajes.MensajeCargaMasivaFallida }, Errores = errores, Archivo = FileName }, JsonRequestBehavior.AllowGet);
                        }

                        if (listado.Count > 0)
                        {
                            resultado = CargaPTOPEntity.CrearActualizarRegistrosPTOPCargaMasiva(listado); // Crea los clientes

                            //Si el resultado es Ok se continua con la generacion de factutas
                            if (resultado.Estado)
                            {
                            //    BussinesLogTran modelLog = new BussinesLogTran();
                            //    BussinesFactura modelFact = new BussinesFactura();
                            //    //BussinesTipoDocumento modelTDoc = new BussinesTipoDocumento();

                            }

                        }
                        else
                        {
                            resultado = new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeCargaMasivaSinRegistros };
                            errores.Add(new CargaMasiva { Fila = 0, Columna = 0, Valor = "Ninguno", Error = "No se encontraron registros. Revisar la estructura del archivo." });
                        }
                    }
                }

                if (resultado.Estado)
                    return Json(new { Resultado = new RespuestaTransaccion { Estado = resultado.Estado, Respuesta = Mensajes.MensajeCargaMasivaExitosa }, Errores = errores, Archivo = FileName }, JsonRequestBehavior.AllowGet);
                else
                {
                    string mensaje = !string.IsNullOrEmpty(resultado.Respuesta) ? resultado.Respuesta  + " " +  Mensajes.MensajeCargaMasivaFallida : "" + Mensajes.MensajeCargaMasivaFallida ;
                    return Json(new { Resultado = new RespuestaTransaccion { Estado = resultado.Estado, Respuesta = mensaje }, Errores = errores, Archivo = FileName }, JsonRequestBehavior.AllowGet);
                }

            }
            catch (Exception ex)
            {
                return Json(new { Resultado = new RespuestaTransaccion { Estado = false, Respuesta = ex.Message } }, JsonRequestBehavior.AllowGet);
            }
        }

        //Cargar la data en un arreglo temporal
        private List<CargaPTOPExcel> ToEntidadHojaExcelList(string pathDelFicheroExcel)
        {
            List<CargaPTOPExcel> listadoCargaMasiva = new List<CargaPTOPExcel>();
            try
            {

                var user = ViewData["usuario"] = System.Web.HttpContext.Current.Session["usuario"];
                var usuarioID = Convert.ToInt16(user);

                FileInfo existingFile = new FileInfo(pathDelFicheroExcel);

                using (ExcelPackage package = new ExcelPackage(existingFile))
                {
                    // get the first worksheet in the workbook
                    ExcelWorksheet worksheet = package.Workbook.Worksheets["DETALLE MENSUAL O CORRECCIONES"]; 

                    if (worksheet != null)
                    {
                        int colCount = worksheet.Dimension.End.Column;  //get Column Count
                        int rowCount = worksheet.Dimension.End.Row;     //get row count
                        for (int row = 2; row <= rowCount; row++)
                        {
                            listadoCargaMasiva.Add(new CargaPTOPExcel
                            {
                                id = !string.IsNullOrEmpty((worksheet.GetValue(row, 1) ?? "").ToString().Trim()) ? int.Parse((worksheet.GetValue(row, 1) ?? "").ToString().Trim()) : 0,
                                id_comercio = !string.IsNullOrEmpty((worksheet.GetValue(row, 2) ?? "").ToString().Trim()) ? (worksheet.GetValue(row, 2) ?? "").ToString().Trim() : string.Empty,
                                plan = !string.IsNullOrEmpty((worksheet.GetValue(row, 3) ?? "").ToString().Trim()) ? (worksheet.GetValue(row, 3) ?? "").ToString().Trim() : string.Empty,
                                facturable_certificacion = !string.IsNullOrEmpty((worksheet.GetValue(row, 4) ?? "").ToString().Trim()) ? (worksheet.GetValue(row, 4) ?? "").ToString().Trim() : string.Empty,
                                facturable_mensual = !string.IsNullOrEmpty((worksheet.GetValue(row, 5) ?? "").ToString().Trim()) ? (worksheet.GetValue(row, 5) ?? "").ToString().Trim() : string.Empty,
                                mes = !string.IsNullOrEmpty((worksheet.GetValue(row, 6) ?? "").ToString().Trim()) ? int.Parse((worksheet.GetValue(row, 6) ?? "").ToString().Trim()) : 0,
                                anio = !string.IsNullOrEmpty((worksheet.GetValue(row, 7) ?? "").ToString().Trim()) ? int.Parse((worksheet.GetValue(row, 7) ?? "").ToString().Trim()) : 0,
                                detalle = !string.IsNullOrEmpty((worksheet.GetValue(row, 8) ?? "").ToString().Trim()) ? (worksheet.GetValue(row, 8) ?? "").ToString().Trim() : string.Empty,
                                valor_certificacion = !string.IsNullOrEmpty((worksheet.GetValue(row, 9) ?? "").ToString().Trim()) ? decimal.Parse((worksheet.GetValue(row, 9) ?? "").ToString().Trim()) : 0,
                                transacciones_aprobadas = !string.IsNullOrEmpty((worksheet.GetValue(row, 10) ?? "").ToString().Trim()) ? int.Parse((worksheet.GetValue(row, 10) ?? "").ToString().Trim()) : 0,
                                transacciones_rechazadas = !string.IsNullOrEmpty((worksheet.GetValue(row, 11) ?? "").ToString().Trim()) ? int.Parse((worksheet.GetValue(row, 11) ?? "").ToString().Trim()) : 0,
                                monto_aprobado = !string.IsNullOrEmpty((worksheet.GetValue(row, 12) ?? "").ToString().Trim()) ? decimal.Parse((worksheet.GetValue(row, 12) ?? "").ToString().Trim()) : 0,
                                monto_rechazado = !string.IsNullOrEmpty((worksheet.GetValue(row, 13) ?? "").ToString().Trim()) ? decimal.Parse((worksheet.GetValue(row, 13) ?? "").ToString().Trim()) : 0,
                                observaciones = !string.IsNullOrEmpty((worksheet.GetValue(row, 14) ?? "").ToString().Trim()) ? (worksheet.GetValue(row, 14) ?? "").ToString().Trim() : "",
                                email = !string.IsNullOrEmpty((worksheet.GetValue(row, 15) ?? "").ToString().Trim()) ? (worksheet.GetValue(row, 15) ?? "").ToString().Trim() : string.Empty   
                            });    
                        }
                    } 
                }

                return listadoCargaMasiva;
            }
            catch (Exception ex)
            {
                return new List<CargaPTOPExcel>();
            }
        }

        // Devuelve los errores que hayan en la plantilla
        private List<CargaMasiva> VerificarEstructuraExcel(string pathDelFicheroExcel)
        {
            List<CargaMasiva> errores = new List<CargaMasiva>();
            try
            {
                FileInfo existingFile = new FileInfo(pathDelFicheroExcel);

                using (ExcelPackage package = new ExcelPackage(existingFile))
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets.First();

                    if (worksheet != null)
                    {
                        int colCount = worksheet.Dimension.End.Column;  //get Column Count
                        int rowCount = worksheet.Dimension.End.Row;      //get row count - Cabecera
                        for (int row = 2; row <= rowCount; row++)
                        {
                            for (int col = 1; col <= colCount; col++)
                            {
                                var error = string.Empty;

                                string columna = (worksheet.Cells[1, col].Value ?? "").ToString().Trim(); // Nombre de la Columna 
                                string valorColumna = (worksheet.Cells[row, col].Value ?? "").ToString().Trim();

                                error = ValidarCamposCargaMasivaCargarPTOP(columna, valorColumna);

                                if (!string.IsNullOrEmpty(error))
                                {
                                    errores.Add(new CargaMasiva { Fila = row, Columna = col, Valor = valorColumna, Error = error });
                                }
                            }

                            //validar que los registros sean consistentes
                            string nombrePlan = (worksheet.Cells[row, 3].Value ?? "").ToString().Trim();
                            int numeroTransacciones = 0;
                            decimal valorTransaccionado = 0;
                            try
                            {
                                numeroTransacciones = Convert.ToInt32(worksheet.Cells[row, 10].Value ?? "");
                            }
                            catch (Exception ex)
                            {
                                string error = "Error con el numero de transacciones de la fila " + row + ", " + worksheet.Cells[row, 10].Value;

                                if (!string.IsNullOrEmpty(error))
                                {
                                    errores.Add(new CargaMasiva { Fila = row, Columna = 3, Valor = "", Error = error });
                                } 
                            }

                            try
                            {
                                valorTransaccionado = Convert.ToDecimal(worksheet.Cells[row, 12].Value ?? "");
                            }
                            catch (Exception ex)
                            {
                                string error = "Error con el valor de transacciones de la fila " + row + ", " + worksheet.Cells[row, 12].Value;

                                if (!string.IsNullOrEmpty(error))
                                {
                                    errores.Add(new CargaMasiva { Fila = row, Columna = 3, Valor = "", Error = error });
                                }
                            }
                            decimal descuento = 0;
                            bool esCertificacion;
                            if ((worksheet.GetValue(row, 4) ?? "").ToString().Trim().ToUpper() == "SI")
                            {
                                esCertificacion = true;
                            }
                            else
                            {
                                esCertificacion = false;
                            }
                            decimal porcentajeIva = 0;
                            try
                            { 
                                decimal valorCertificacion = Convert.ToDecimal(worksheet.Cells[row, 9].Value ?? "");
                                var calculo = CargaPTOPEntity.CalcularValoresPTOP(nombrePlan, numeroTransacciones, valorTransaccionado, esCertificacion, porcentajeIva, valorCertificacion, descuento).FirstOrDefault();
                                //Verificar que no tenga errores
                                if (calculo.Correcto == false)
                                {
                                    string error = "Error al encontrar el plan de la fila " + row + ", " + calculo.Error;

                                    if (!string.IsNullOrEmpty(error))
                                    {
                                        errores.Add(new CargaMasiva { Fila = row, Columna = 3, Valor = "", Error = error });
                                    }
                                }
                            }
                            catch
                            {
                                string error = "Error en el valor de certificación" + row + ", " + worksheet.Cells[row, 9].Value;
                                if (!string.IsNullOrEmpty(error))
                                {
                                    errores.Add(new CargaMasiva { Fila = row, Columna = 3, Valor = "", Error = error });
                                }
                            }
  
                        }
                    }
                }

                return errores;
            }
            catch (Exception ex)
            {
                errores.Add(new CargaMasiva { Fila = 0, Columna = 0, Valor = "No se pudo verificar correctamente la estructura del excel.", Error = ex.Message.ToString() });
                return errores;
            }

        }

        // Valida todas las celdas del excel
        private static string ValidarCamposCargaMasivaCargarPTOP(string columna, string valor)
        {
            var excepcion = "El campo {0} con el valor {1} ,tiene una extensión de caracteres mayor a la permitida. Error: {2}";
            try
            {
                valor = !string.IsNullOrEmpty(valor) ? valor.Trim() : "";

                bool esNumero; 
                int longitudCaracteres = 0;
                var error = "El campo {0} con el valor {1} ,tiene una extensión de caracteres mayor a la permitida. Solo permite {2} caracteres";

                switch (columna)
                {
                    case "ID":
                        esNumero = int.TryParse(valor, out int valorNumeracion);
                        if (!esNumero)
                            error = "Tipo de dato incorrecto para el campo " + columna + ", con el valor " + valor + " . Solo admite valores numéricos de tipo entero.";
                        else
                            error = null;
                        break;

                    case "ID COMERCIO":
                        longitudCaracteres = 50;
                        if (valor.Length > longitudCaracteres)
                        {
                            error = string.Format(error, columna, valor, longitudCaracteres);
                        }
                        else
                        {
                            if (valor.Length == 0)
                            {
                                error = "El campo " + columna + " , no puede ser vacío" + valor;
                            }
                            else
                            {
                                //validar que el comercio exista
                                var comercio = ComercioEntity.ListarComercios().FirstOrDefault(c => c.ID_Comercio == valor.ToUpper());  
                                if (comercio == null)
                                {
                                    error = "El campo " + columna + ", con el valor " + valor + ", no se encuentra activo o registrado en el Sistema.";
                                }
                                else
                                {
                                    error = null;
                                }
                            }
                        }
                        break;

                    case "PLAN":
                        //VALIDAR QUE EXISTA EL PLAN                         
                        var plan = TablaPlanesEntity.ListarTablaPlanes().FirstOrDefault(c => c.Nombre_Plan == valor.ToUpper());

                        if (plan == null)
                        {
                            error = "El campo " + columna + ", con el valor " + valor + ", no se encuentra activo o registrado en el Sistema.";
                        }
                        else
                            error = null;
                        break;

                    case "FACTURABLE CERTIFICACIÓN":
                        //VALIDAR QUE EXISTA EL CATALOGO                          
                        var certificacion = CatalogoEntity.ListadoCatalogosPorCodigo("FCR-01").FirstOrDefault(c => c.Text == valor.ToUpper());

                        if (certificacion == null)
                        {
                            error = "El campo " + columna + ", con el valor " + valor + ", no se encuentra activo o registrado en Catálogos del Sistema.";
                        }
                        else
                            error = null;
                        break;

                    case "FACTURABLE MENSUAL":
                        //VALIDAR QUE EXISTA EL CATALOGO                          
                        var mensual = CatalogoEntity.ListadoCatalogosPorCodigo("FMS-01").FirstOrDefault(c => c.Text == valor.ToUpper());

                        if (mensual == null)
                        {
                            error = "El campo " + columna + ", con el valor " + valor + ", no se encuentra activo o registrado en Catálogos del Sistema.";
                        }
                        else
                            error = null;
                        break;

                    case "MES":
                    case "AÑO":
                    case "TRX APROBADAS":
                    case "TRX RECHAZADAS":
                        esNumero = int.TryParse(valor, out int valorMes);
                        if (!esNumero)
                            error = "Tipo de dato incorrecto para el campo " + columna + ", con el valor " + valor + " . Solo admite valores numéricos de tipo entero.";
                        else
                            error = null;
                        break; 

                    case "DETALLE":
                    case "EMAIL":
                        longitudCaracteres = 500;
                        if (valor.Length > longitudCaracteres)
                        {
                            error = string.Format(error, columna, valor, longitudCaracteres);
                        }
                        else
                        {
                            if (valor.Length == 0)
                            {
                                error = "El campo " + columna + ", no puede ser vacío" + valor;
                            }
                            else
                                error = null;
                        }
                        break;
                         
                    case "VALOR A COBRAR POR CERTIFICACIÓN":
                    case "MONTO VENDIDO APROBADO":
                    case "MONTO VENDIDO RECHAZADO":
                        if (valor.Length == 0)
                        {
                            error = "El campo " + columna + ", no puede ser vacío" + valor;
                        }
                        else
                        {
                            var separadorIncorrecto = valor.IndexOf(".");

                            if (separadorIncorrecto != -1)
                                valor.Replace(".", ",");

                            esNumero = decimal.TryParse(valor, out decimal valorFormaPago);
                            if (!esNumero)
                                error = "Tipo de dato incorrecto para el campo " + columna + " , con el valor " + valor + " .Solo admite valores numéricos de tipo entero.";
                            else
                                error = null;
                        }
                        break;

                    case "OBSERVACIÓN":
                        error = null; 
                        break;
                default:
                        error = "No se pudo encontrar la columna " + columna + ".";
                        break;
                } 
                return error;
            }
            catch (Exception ex)
            {
                return string.Format(excepcion, columna, valor, ex.Message.ToString());
            }
        }

    #endregion 

        public ActionResult _GestionComercios(int? id)
        {
            ViewBag.TituloModal = "Gestión Contratos Comercios";

            //Listado Estatus Contrato
            var EstatusContrato = CatalogoEntity.ObtenerListadoCatalogosByCodigo("ECP-01");
            ViewBag.ListadoEstatusContrato = EstatusContrato;

            //Listado Canal de Comunicación
            var CanalComunicacion = CatalogoEntity.ObtenerListadoCatalogosByCodigo("CCP-01");
            ViewBag.ListadoCanalComunicacion = CanalComunicacion;

            //Listado Tipificación
            var Tipificacion = CatalogoEntity.ObtenerListadoCatalogosByCodigo("TFP-01");
            ViewBag.ListadoTipificacion = Tipificacion;

            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }

            Comercio comercio = ComercioEntity.ConsultarComercio(id.Value);

            GestionContratosComercios gestion = new GestionContratosComercios();
            gestion.id_comercio = comercio.id_comercio;
            gestion.id_estatus_contrato_actual = comercio.id_estatus_contrato;
            gestion.id_estatus_contrato_anterior = comercio.id_estatus_contrato;

            if (gestion == null)
            {
                return HttpNotFound();
            }
            return PartialView(gestion); 
        }

        [HttpPost]
        public ActionResult GestionComercios(GestionContratosComercios gestion)
        {
            try
            {
                //Obtener el codigo de usuario que se logeo
                var usuarioSesion = ViewData["usuario"] = System.Web.HttpContext.Current.Session["usuario"];
                var idUsuario = Convert.ToInt16(usuarioSesion);
                var usuarioclave = ViewData["usuarioClave"] = System.Web.HttpContext.Current.Session["usuarioClave"];

                gestion.id_usuario = idUsuario;
                gestion.fecha_comunicacion = DateTime.Now;

                RespuestaTransaccion resultado = GestionContratosComerciosEntity.CrearGestionContratosComercios(gestion);
                return Json(new { Resultado = resultado }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                return Json(new { Resultado = new RespuestaTransaccion { Estado = false, Respuesta = ex.Message } }, JsonRequestBehavior.AllowGet);
            }
        }
    }
}

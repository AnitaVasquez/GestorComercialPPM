using System;
using System.Collections.Generic;
using System.Data; 
using System.IO;
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
using Omu.Awem.Helpers; 
using Seguridad.Helper;

namespace TemplateInicial.Controllers
{ 
    [Autenticado]
    public class ComercioController : BaseAppController
    {
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
            ViewBag.NombreListado = Etiquetas.TituloGridComercio;

            //Controlar permisos
            var user = ViewData["usuario"] = System.Web.HttpContext.Current.Session["usuario"];
            var usuario = int.Parse(user.ToString());
            string nombreControlador = ControllerContext.RouteData.Values["controller"].ToString();
            ViewBag.NombreControlador = nombreControlador;

            ViewBag.AccionesUsuario = ManejoPermisosEntity.ListadoAccionesCatalogoUsuario(usuario, nombreControlador);

            //Obtener Acciones del controlador
            ViewBag.AccionesControlador = GetMetodosControlador(nombreControlador);

            //Búsqueda

            var listado = ComercioEntity.ListarComercios();

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

        // GET: Cliente/Create
        public ActionResult Create()
        {
            //Listado Estatus Contrato
            var EstatusContrato = CatalogoEntity.ObtenerListadoCatalogosByCodigo("ECP-01");
            ViewBag.ListadoEstatusContrato = EstatusContrato;

            //Listado Tipo Subsidio
            var TipoSubsidio = CatalogoEntity.ObtenerListadoCatalogosByCodigo("TSP-01");
            ViewBag.ListadoTipoSubsidio = TipoSubsidio;

            //Listado Tipo Agrupación
            var TipoAgrupacion = CatalogoEntity.ObtenerListadoCatalogosByCodigo("AGP-01");
            ViewBag.ListadoTipoAgrupacion = TipoAgrupacion;
             
            //Listado Clientes
            var ClienteLineaNegocio = ClienteEntity.ListarClientePorLineaNegocio("PLACE TO PAY");
            ViewBag.ListadoCliente = ClienteLineaNegocio;
              
            return View();
        }

        [HttpPost]
        public ActionResult Create(Comercio comercio)
        {
            try
            {
                string nombreComercio = (comercio.nombre_comercio ?? string.Empty).ToLower().Trim();
                string idComercio = (comercio.id_comercio_predictive ?? string.Empty).ToLower().Trim();
                 
                var ComercioIguales = ComercioEntity.ListarComercios().Where(s => (s.ID_Comercio ?? string.Empty).ToLower().Trim() == idComercio).ToList();

                if (ComercioIguales.Count > 0)
                {
                    return Json(new { Resultado = new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeResgistroExistente } }, JsonRequestBehavior.AllowGet);
                }
                else
                {
                    RespuestaTransaccion resultado = ComercioEntity.CrearComercio(comercio);
                    return Json(new { Resultado = resultado }, JsonRequestBehavior.AllowGet);
                }
            }
            catch (Exception ex)
            {
                return Json(new { Resultado = new RespuestaTransaccion { Estado = false, Respuesta = ex.Message } }, JsonRequestBehavior.AllowGet);
            }
        }

        // GET: Cliente/Edit/5
        public ActionResult Edit(int? id)
        {
            //Listado Estatus Contrato
            var EstatusContrato = CatalogoEntity.ObtenerListadoCatalogosByCodigo("ECP-01");
            ViewBag.ListadoEstatusContrato = EstatusContrato;

            //Listado Tipo Subsidio
            var TipoSubsidio = CatalogoEntity.ObtenerListadoCatalogosByCodigo("TSP-01");
            ViewBag.ListadoTipoSubsidio = TipoSubsidio;

            //Listado Tipo Agrupación
            var TipoAgrupacion = CatalogoEntity.ObtenerListadoCatalogosByCodigo("AGP-01");
            ViewBag.ListadoTipoAgrupacion = TipoAgrupacion;

            //Listado Clientes
            var ClienteLineaNegocio = ClienteEntity.ListarClientePorLineaNegocio("PLACE TO PAY");
            ViewBag.ListadoCliente = ClienteLineaNegocio;

            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }

            Comercio comercio = ComercioEntity.ConsultarComercio(id.Value);

            

            if (comercio == null)
            {
                return HttpNotFound();
            }
            else
            {
                ViewBag.fechaProduccion = Convert.ToDateTime(comercio.fecha_salida_produccion).ToString("yyyy-MM-dd");
                if (comercio.fecha_inactivo != null)
                {
                    ViewBag.fechainactiva = Convert.ToDateTime(comercio.fecha_inactivo).ToString("yyyy-MM-dd");
                }
                 
                ViewBag.Descuento = comercio.descuento.HasValue ? comercio.descuento : 0;
                ViewBag.Iva = comercio.porcentaje_iva.HasValue ? comercio.porcentaje_iva : 0;
            }
            return View(comercio);
        } 
        
        [HttpPost]
        public ActionResult Edit(Comercio comercio)
        {
            try
            {
                string nombreComercio = (comercio.nombre_comercio ?? string.Empty).ToLower().Trim();
                string idComercio = (comercio.id_comercio_predictive ?? string.Empty).ToLower().Trim();

                var ComercioIguales = ComercioEntity.ListarComercios().Where(s => (s.ID_Comercio ?? string.Empty).ToLower().Trim() == idComercio && s.Codigo != comercio.id_comercio).ToList();

                if (ComercioIguales.Count > 0)
                {
                    return Json(new { Resultado = new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeResgistroExistente } }, JsonRequestBehavior.AllowGet);
                }
                else
                {
                    RespuestaTransaccion resultado = ComercioEntity.ActualizarComercio(comercio);
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
            RespuestaTransaccion resultado = ClienteEntity.EliminarCliente(id);// await db.Cabecera.FindAsync(id);

            //return RedirectToAction("Index");
            return Json(new { Resultado = resultado }, JsonRequestBehavior.AllowGet);

        } 
         
        public ActionResult DescargarReporteFormatoExcel()
        {
            var collection = ComercioEntity.ListarComercios();

            ExcelPackage package = new ExcelPackage();
            var workSheet = package.Workbook.Worksheets.Add("Listado Comercios");
            workSheet.TabColor = System.Drawing.Color.Black;
            workSheet.DefaultRowHeight = 12;

            List<string> columnas = new List<string>() {
                "ID",
                "ID COMERCIO",
                "COMERCIO",
                "RUC CLIENTE",
                "FECHA SALIDA A PRODUCCIÓN",
                "FECHA INACTIVO",
                "MESES ACTIVOS",
                "MESES INACTIVOS",
                "ESTATUS CONTRATO",
                "TIPO SUBSIDIADO",
                "AGRUPACIÓN",
                "COMPARTIDO CON PTOP",
                "TIPO_DESCUENTO",
                "DESCUENTO",
                "% IVA",
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
                workSheet.Cells[recordIndex, 2].Value = item.ID_Comercio;
                workSheet.Cells[recordIndex, 3].Value = item.Comercio;
                workSheet.Cells[recordIndex, 4].Value = item.RazonSocial;
                workSheet.Cells[recordIndex, 5].Value = item.fecha_salida_produccion;
                workSheet.Cells[recordIndex, 6].Value = item.fecha_inactivo;
                workSheet.Cells[recordIndex, 7].Value = item.meses_activos;
                workSheet.Cells[recordIndex, 8].Value = item.meses_inactivos;
                workSheet.Cells[recordIndex, 9].Value = item.EstatusContrato;
                workSheet.Cells[recordIndex, 10].Value = item.TipoSubsidio;
                workSheet.Cells[recordIndex, 11].Value = item.Agrupacion;
                workSheet.Cells[recordIndex, 12].Value = item.CompartidoPTOP.ToUpper(); 
                workSheet.Cells[recordIndex, 13].Value = item.cobro_porcentaje;
                workSheet.Cells[recordIndex, 14].Value = item.Descuento;
                workSheet.Cells[recordIndex, 14].Style.Numberformat.Format = "###,##0.00";
                workSheet.Cells[recordIndex, 15].Value = item.PorcentajeIva;
                workSheet.Cells[recordIndex, 15].Style.Numberformat.Format = "###,##0.00";
                workSheet.Cells[recordIndex, 16].Value = item.Estado_comercio.ToUpper();
                recordIndex++;
            }

            for (int i = 1; i <= columnas.Count; i++)
            {
                workSheet.Column(i).AutoFit();
                workSheet.Column(i).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            }


            return File(package.GetAsByteArray(), XlsxContentType, "ListadoComercio.xlsx");
        }

        public ActionResult DescargarReporteFormatoPDF()
        {
            // Seleccionar las columnas a exportar
            var results = ComercioEntity.ListarComercios();

            var list = Reportes.SerializeToJSON(results);
            return Content(list, "application/json");
        }

        public ActionResult DescargarReporteFormatoCSV()
        {
            var comlumHeadrs = new string[]
            {
                "ID",
                "ID COMERCIO",
                "COMERCIO",
                "RUC CLIENTE",
                "FECHA SALIDA A PRODUCCIÓN",
                "FECHA INACTIVO",
                "MESES ACTIVOS",
                "MESES INACTIVOS",
                "ESTATUS CONTRATO",
                "TIPO SUBSIDIADO",
                "AGRUPACIÓN",
                "COMPARTIDO CON PTOP",
                "TIPO_DESCUENTO",
                "DESCUENTO",
                "% IVA",
                "ESTADO"
            };

            var listado = (from item in ComercioEntity.ListarComercios()
                           select new object[]
                           {
                                item.Codigo,
                                item.ID_Comercio,
                                item.Comercio,
                                item.RazonSocial,
                                item.fecha_salida_produccion,
                                item.fecha_inactivo,
                                item.meses_activos,
                                item.meses_inactivos,
                                item.EstatusContrato,
                                item.TipoSubsidio,
                                item.Agrupacion,
                                item.CompartidoPTOP.ToUpper(), 
                                item.cobro_porcentaje,
                                item.Descuento,
                                item.PorcentajeIva,
                                item.Estado_comercio.ToUpper()
                           }).ToList();

            // Build the file content
            var employeecsv = new StringBuilder();
            listado.ForEach(line =>
            {
                employeecsv.AppendLine(string.Join(",", line));
            });

            byte[] buffer = Encoding.Default.GetBytes($"{string.Join(",", comlumHeadrs)}\r\n{employeecsv.ToString()}");
            return File(buffer, "text/csv", $"Comercios.csv");
        }
         
        #region Funcionalidad de carga Masiva
        public ActionResult CargarData()
        {
            List<CargaMasiva> errores = new List<CargaMasiva>();
            RespuestaTransaccion resultado = new RespuestaTransaccion { Estado = false };
            List<ComercioExcel> listado = new List<ComercioExcel>();
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
                        
                        if (numeracionDuplicada) {
                            errores.Add(new CargaMasiva { Fila = 0, Columna = 0, Valor = "Ninguno", Error = "Valores de la Columna de numeración 'ID' no pueden ser duplicados." });
                            return Json(new { Resultado = new RespuestaTransaccion { Estado = resultado.Estado, Respuesta = Mensajes.MensajeCargaMasivaFallida }, Errores = errores, Archivo = FileName }, JsonRequestBehavior.AllowGet);
                        }

                        if (listado.Count > 0)
                            resultado = ComercioEntity.CrearActualizarComerciosCargaMasiva(listado); // Crea los clientes
                        else {
                            resultado = new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeCargaMasivaSinRegistros };
                            errores.Add(new CargaMasiva { Fila = 0, Columna = 0, Valor = "Ninguno", Error = "No se encontraron registros. Revisar la estructura del archivo." });
                        }
                    }
                }

                if (resultado.Estado)
                    return Json(new { Resultado = new RespuestaTransaccion { Estado = resultado.Estado, Respuesta = Mensajes.MensajeCargaMasivaExitosa }, Errores = errores, Archivo = FileName }, JsonRequestBehavior.AllowGet);
                else
                {
                    string mensaje = !string.IsNullOrEmpty(resultado.Respuesta) ? "" + Mensajes.MensajeCargaMasivaFallida : resultado.Respuesta + Mensajes.MensajeCargaMasivaFallida;
                    return Json(new { Resultado = new RespuestaTransaccion { Estado = resultado.Estado, Respuesta = mensaje }, Errores = errores, Archivo = FileName }, JsonRequestBehavior.AllowGet);
                }
            }
            catch (Exception ex)
            {
                return Json(new { Resultado = new RespuestaTransaccion { Estado = false, Respuesta = ex.Message } }, JsonRequestBehavior.AllowGet);
            }
        }

        //Cargar la data en un arreglo temporal
        private List<ComercioExcel> ToEntidadHojaExcelList(string pathDelFicheroExcel)
        {
            List<ComercioExcel> listadoCargaMasiva = new List<ComercioExcel>();
            try
            {

                var user = ViewData["usuario"] = System.Web.HttpContext.Current.Session["usuario"];
                var usuarioID = Convert.ToInt16(user);

                FileInfo existingFile = new FileInfo(pathDelFicheroExcel);

                using (ExcelPackage package = new ExcelPackage(existingFile))
                {
                    // get the first worksheet in the workbook
                    ExcelWorksheet worksheet = package.Workbook.Worksheets["COMERCIOS"]; 

                    if (worksheet != null)
                    {
                        int colCount = worksheet.Dimension.End.Column;  //get Column Count
                        int rowCount = worksheet.Dimension.End.Row;     //get row count
                        for (int row = 2; row <= rowCount; row++)
                        {  
                            listadoCargaMasiva.Add(new ComercioExcel
                            {
                                id = !string.IsNullOrEmpty((worksheet.GetValue(row, 1) ?? "").ToString().Trim()) ? int.Parse((worksheet.GetValue(row, 1) ?? "").ToString().Trim()) : 0,
                                id_comercio = !string.IsNullOrEmpty((worksheet.GetValue(row, 2) ?? "").ToString().Trim()) ? (worksheet.GetValue(row, 2) ?? "").ToString().Trim() : string.Empty,
                                comercio = !string.IsNullOrEmpty((worksheet.GetValue(row, 3) ?? "").ToString().Trim()) ? (worksheet.GetValue(row, 3) ?? "").ToString().Trim() : string.Empty,
                                ruc_cliente = !string.IsNullOrEmpty((worksheet.GetValue(row, 4) ?? "").ToString().Trim()) ? (worksheet.GetValue(row, 4) ?? "").ToString().Trim() : string.Empty,
                                fecha_salida_produccion = !string.IsNullOrEmpty((worksheet.GetValue(row, 5) ?? "").ToString().Trim()) ? Convert.ToDateTime((worksheet.GetValue(row, 5) ?? "")).ToString("yyyy/MM/dd"): null,
                                fecha_inactivo = !string.IsNullOrEmpty((worksheet.GetValue(row, 6) ?? "").ToString().Trim()) ? Convert.ToDateTime((worksheet.GetValue(row, 6) ?? "")).ToString("yyyy/MM/dd") : null,
                                estatus_contrato = !string.IsNullOrEmpty((worksheet.GetValue(row, 7) ?? "").ToString().Trim()) ? (worksheet.GetValue(row, 7) ?? "").ToString().Trim() : "",
                                tipo_subsidio = !string.IsNullOrEmpty((worksheet.GetValue(row, 8) ?? "").ToString().Trim()) ? (worksheet.GetValue(row, 8) ?? "").ToString().Trim() : "",
                                agrupacion = !string.IsNullOrEmpty((worksheet.GetValue(row, 9) ?? "").ToString().Trim()) ? (worksheet.GetValue(row, 9) ?? "").ToString().Trim() : string.Empty, 
                                compartido_ptop = !string.IsNullOrEmpty((worksheet.GetValue(row, 10) ?? "").ToString().Trim()) ? (worksheet.GetValue(row, 10) ?? "").ToString().Trim() : string.Empty,
                                descuento = !string.IsNullOrEmpty((worksheet.GetValue(row, 11) ?? "").ToString().Trim()) ? decimal.Parse((worksheet.GetValue(row, 11) ?? "").ToString().Trim()) : 0,
                                cobro_porcentaje = !string.IsNullOrEmpty((worksheet.GetValue(row, 12) ?? "").ToString().Trim()) ? (worksheet.GetValue(row, 12) ?? "").ToString().Trim() : string.Empty,
                                porcentaje_iva = !string.IsNullOrEmpty((worksheet.GetValue(row, 13) ?? "").ToString().Trim()) ? decimal.Parse((worksheet.GetValue(row, 13) ?? "").ToString().Trim()) : 0,

                            });   
                        }
                    } 
                }

                return listadoCargaMasiva;
            }
            catch (Exception ex)
            {
                return new List<ComercioExcel>();
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

                                error = ValidarCamposCargaMasivaComercios(columna, valorColumna);

                                if (!string.IsNullOrEmpty(error))
                                {
                                    errores.Add(new CargaMasiva { Fila = row, Columna = col, Valor = valorColumna, Error = error });
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
        private static string ValidarCamposCargaMasivaComercios(string columna, string valor)
        {
            var excepcion = "El campo {0} con el valor {1} ,tiene una extensión de caracteres mayor a la permitida. Error: {2}";
            try
            {
                valor = !string.IsNullOrEmpty(valor) ? valor.Trim() : "";

                bool esNumero;
                bool esCedulaRuc; 
                int longitudCaracteres = 0;
                var error = "El campo {0} con el valor {1} ,tiene una extensión de caracteres mayor a la permitida. Solo permite {2} caracteres";

                switch (columna)
                {
                    case "ID":
                        esNumero = int.TryParse(valor, out int valorNumeracion);
                        if (!esNumero)
                            error = "Tipo de dato incorrecto para el campo " + columna + " , con el valor " + valor + " . Solo admite valores numéricos de tipo entero.";
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
                                error = null;
                        }
                        break;

                    case "COMERCIO":
                        longitudCaracteres = 500;
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
                                error = null;
                        }
                        break;

                    case "RUC CLIENTE":
                        //validar el ruc sea valido en caracteres
                        if(valor.Length == 9 || valor.Length == 12)
                        {
                            error = "El valor del " + columna + " , con valor " + valor + ".  Es una Cédula/RUC incorrecta .";
                        }
                        else
                        {
                            if (valor.Length == 0)
                            {
                                error = "El campo " + columna + " , no puede ser vacío" + valor;
                            }
                            else
                            {
                                //Aplicar el validador de cedula o ruc
                                esCedulaRuc = Validaciones.VerificaIdentificacion(valor);
                                if (!esCedulaRuc)
                                    error = "El valor del " + columna + " , con valor " + valor + ".  Es una Cédula/RUC incorrecta .";
                                else
                                {
                                    //validar que el cliente exista
                                    var cliente = ClienteEntity.ConsultarClienteLineaNegocio("Place to Pay",valor);
                                    if (cliente == null)
                                    {
                                        error = "El campo " + columna + ", con el valor " + valor + ", no se encuentra activo o registrado en el Sistema.";
                                    }
                                    else
                                        error = null; 
                                }
                            }                          

                        } 
                        break; 

                    case "FECHA SALIDA A PRODUCCIÓN":

                        if (valor.Length == 0)
                        {
                            error = "El campo " + columna + " , no puede ser vacío" + valor;
                        }
                        else
                        {
                            if (valor.Length != 10)
                            {
                                error = "El campo " + columna + " , con el valor" + valor + ". Solo permite fechas con formato (yyyy-MM-dd)";
                            }
                            else
                            {
                                esNumero = int.TryParse(valor.Replace("-", ""), out int fecha);
                                if (!esNumero)
                                    error = "Tipo de dato incorrecto para el campo " + columna + " , con el valor " + valor + " . Solo permite fechas con formato (yyyy-MM-dd)";
                                else
                                    error = null;
                            }                            
                        }                          
                        break;

                    case "FECHA INACTIVO":

                        if (valor.Length != 0)
                        {
                            if (valor.Length != 10)
                            {
                                error = "El campo " + columna + " , con el valor" + valor + ". Solo permite fechas con formato (yyyy-MM-dd)";
                            }
                            else
                            {
                                esNumero = int.TryParse(valor.Replace("-", ""), out int fecha);
                                if (!esNumero)
                                    error = "Tipo de dato incorrecto para el campo " + columna + " , con el valor " + valor + " . Solo permite fechas con formato (yyyy-MM-dd)";
                                else
                                    error = null;
                            }
                        }
                        else
                            error = null;
                        break;

                    case "ESTATUS CONTRATO":
                        //VALIDAR QUE EXISTA EL CATALOGO                          
                        var estatus = CatalogoEntity.ListadoCatalogosPorCodigo("ECP-01").FirstOrDefault(c => c.Text == valor.ToUpper());

                        if (estatus == null)
                        { 
                            error = "El campo " + columna + ", con el valor " + valor + ", no se encuentra activo o regsitrado en Catálogos del Sistema.";
                        }
                        else 
                            error = null;
                        break;

                    case "TIPO":
                        //VALIDAR QUE EXISTA EL CATALOGO                          
                        var tipo = CatalogoEntity.ListadoCatalogosPorCodigo("TSP-01").FirstOrDefault(c => c.Text == valor.ToUpper());

                        if (tipo == null)
                        {
                            error = "El campo " + columna + ", con el valor " + valor + ", no se encuentra activo o regsitrado en Catálogos del Sistema.";
                        }
                        else
                            error = null;
                        break;

                    case "AGRUPACIÓN":
                        //VALIDAR QUE EXISTA EL CATALOGO                          
                        var agrupacion = CatalogoEntity.ListadoCatalogosPorCodigo("AGP-01").FirstOrDefault(c => c.Text == valor.ToUpper());

                        if (agrupacion == null)
                        {
                            error = "El campo " + columna + ", con el valor " + valor + ", no se encuentra activo o regsitrado en Catálogos del Sistema.";
                        }
                        else
                            error = null;
                        break;

                    case "COMPARTIDO CON PTOP":
                       if(valor.ToUpper() == "SI" || valor.ToUpper() =="NO")
                            error = null;
                        else
                        {
                            error = "El campo " + columna + ", con el valor " + valor + ". Solo admite SI o NO";
                        }
                        break;

                    case "DESCUENTO":
                    case "PORCENTAJE IVA": 
                        var separadorIncorrecto = valor.IndexOf(".");

                        if (separadorIncorrecto != -1)
                            valor.Replace(".", ",");

                        esNumero = decimal.TryParse(valor, out decimal valorFormaPago);
                        if (!esNumero)
                            error = "Tipo de dato incorrecto para el campo " + columna + " , con el valor " + valor + " .Solo admite valores numéricos de tipo entero.";
                        else
                            error = null;
                        break;

                    case "DESCUENTO PORCENTAJE":
                        if (valor.ToUpper() == "SI" || valor.ToUpper() == "NO")
                            error = null;
                        else
                        {
                            error = "El campo " + columna + ", con el valor " + valor + ". Solo admite SI o NO";
                        }
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

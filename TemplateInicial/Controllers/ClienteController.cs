using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Entity;
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
using Newtonsoft.Json;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using Omu.Awem.Helpers;
using Seguridad.Helper;

namespace TemplateInicial.Controllers
{
    public partial class ClienteContactos
    {
        public int id_cliente { get; set; }
        public int? usuario_id { get; set; }
        public int? referido { get; set; }
        public int? intermediario { get; set; }
        public int? tipo_zoho { get; set; }
        public int? tipo_cliente { get; set; }
        public int? tamanio_empresa { get; set; }
        public int? etapa_cliente { get; set; }
        public int? potencial_crecimiento { get; set; }
        public int? categorizacion_cliente { get; set; }
        public string ruc_ci_cliente { get; set; }
        public string razon_social_cliente { get; set; }
        public string nombre_comercial_cliente { get; set; }
        public decimal? ingresos_anuales_cliente { get; set; }
        public int? sector { get; set; }
        public int? pais { get; set; }
        public int? ciudad { get; set; }
        public string direccion_cliente { get; set; }
        public bool? estado_cliente { get; set; }
        public List<int> idsContactosClientes_guardar { get; set; }
        public List<int> idsContactosClientes_editar { get; set; }
    }

    [Autenticado]
    public class ClienteController : BaseAppController
    {
        private const string XlsxContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
        private int rolID = Convert.ToInt32(System.Web.HttpContext.Current.Session["rolID"]);

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
            ViewBag.NombreListado = Etiquetas.TituloGridCliente;

            //Controlar permisos
            var user = ViewData["usuario"] = System.Web.HttpContext.Current.Session["usuario"];
            var usuario = int.Parse(user.ToString());
            string nombreControlador = ControllerContext.RouteData.Values["controller"].ToString();
            ViewBag.NombreControlador = nombreControlador;

            ViewBag.AccionesUsuario = ManejoPermisosEntity.ListadoAccionesCatalogoUsuario(usuario, nombreControlador);

            //Obtener Acciones del controlador
            ViewBag.AccionesControlador = GetMetodosControlador(nombreControlador);

            //Búsqueda
            var listado = ClienteEntity.ListarCliente();
            //var listado = TarifarioEntity.ListadoTarifario();
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
            //Listado Referido
            var Referido = CatalogoEntity.ObtenerListadoCatalogosByCodigo("REF-01");
            ViewBag.ListadoReferido = Referido;

            //Listado Intermediario
            var Intermediario = CatalogoEntity.ObtenerListadoCatalogosByCodigo("INT-01");
            ViewBag.ListadoIntermediario = Intermediario;

            //Listado TipoZoho
            var TipoZoho = CatalogoEntity.ObtenerListadoCatalogosByCodigo("TZO-01");
            ViewBag.ListadoTipoZoho = TipoZoho;

            //Listado TipoCliente
            var TipoCliente = CatalogoEntity.ObtenerListadoCatalogosByCodigo("TCL-01");
            ViewBag.ListadoTipoCliente = TipoCliente;

            //Listado TamanioEmpresa
            var TamanioEmpresa = CatalogoEntity.ObtenerListadoCatalogosByCodigo("TEP-01");
            ViewBag.ListadoTamanioEmpresa = TamanioEmpresa;

            //Listado Sector
            var Sector = CatalogoEntity.ObtenerListadoCatalogosByCodigo("SCT-01");
            ViewBag.ListadoSector = Sector;

            //Listado Pais
            var Pais = CatalogoEntity.ObtenerListadoCatalogosByCodigo("PAI-01");
            ViewBag.ListadoPais = Pais;

            //Listado Ciudad
            var Ciudad = CatalogoEntity.ObtenerListadoCatalogosByCodigo("CIU-01");
            ViewBag.ListadoCiudad = Ciudad;

            //Listado PotenciaCrecimiento 
            var PotenciaCrecimiento = CatalogoEntity.ObtenerListadoCatalogosByCodigo("PTC-01");
            ViewBag.ListadoPotenciaCrecimiento = PotenciaCrecimiento;

            //Listado EstatusCliente
            var EstatusCliente = CatalogoEntity.ObtenerListadoCatalogosByCodigo("ECL-02");
            ViewBag.ListadoEstatusCliente = EstatusCliente;

            //Listado CategorizacionCliente
            var CategorizacionCliente = CatalogoEntity.ObtenerListadoCatalogosByCodigo("CCL-01");
            ViewBag.ListadoCategorizacionCliente = CategorizacionCliente;

            //Listado EtapaCliente
            var EtapaCliente = CatalogoEntity.ObtenerListadoCatalogosByCodigo("ETC-02");
            ViewBag.ListadoEtapaCliente = EtapaCliente;

            Cliente cliente = new Cliente();
            cliente.estado_cliente = true;

            return View(cliente);
        }

        [HttpPost]
        public ActionResult Create(ClienteContactos cliente, List<int> contactosCliente, List<int> lineasNegocioCliente, bool? usar_organigrama)
        {
            try
            {
                //Validacion solo para el país de Ecuador.
                if (!Validaciones.VerificaIdentificacion(cliente.ruc_ci_cliente) && cliente.pais == 233)
                    return Json(new { Resultado = new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeCedulaRucIncorrecto } }, JsonRequestBehavior.AllowGet);

                if (!ClienteEntity.VerificarRUCCedulaExistente(cliente.ruc_ci_cliente))
                    return Json(new { Resultado = new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeValidacionCedulaRucRegistrado } }, JsonRequestBehavior.AllowGet);

                var user = ViewData["usuario"] = System.Web.HttpContext.Current.Session["usuario"];

                cliente.usuario_id = Convert.ToInt16(user);

                RespuestaTransaccion resultado = ClienteEntity.CrearCliente(new Cliente
                {
                    referido = cliente.referido,
                    intermediario = cliente.intermediario,
                    tipo_zoho = cliente.tipo_zoho,
                    tipo_cliente = cliente.tipo_cliente,
                    tamanio_empresa = cliente.tamanio_empresa,
                    etapa_cliente = cliente.etapa_cliente,
                    potencial_crecimiento = cliente.potencial_crecimiento,
                    categorizacion_cliente = cliente.categorizacion_cliente,
                    ruc_ci_cliente = cliente.ruc_ci_cliente,
                    razon_social_cliente = cliente.razon_social_cliente,
                    nombre_comercial_cliente = cliente.nombre_comercial_cliente,
                    ingresos_anuales_cliente = cliente.ingresos_anuales_cliente,
                    sector = cliente.sector,
                    pais = cliente.pais,
                    ciudad = cliente.ciudad,
                    direccion_cliente = cliente.direccion_cliente,
                    estado_cliente = cliente.estado_cliente,
                    usar_organigrama = usar_organigrama.Value,
                    usuario_id = Convert.ToInt16(user)

                }, contactosCliente, lineasNegocioCliente);

                return Json(new { Resultado = resultado }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                return Json(new { Resultado = new RespuestaTransaccion { Estado = false, Respuesta = ex.Message } }, JsonRequestBehavior.AllowGet);
            }
        }

        // GET: Cliente/Edit/5
        public ActionResult Edit(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            Cliente cliente = ClienteEntity.ConsultarCliente(id.Value); //db.Cliente.Find(id);

            #region Catalogos
            //Listado Referido
            var Referido = CatalogoEntity.ObtenerListadoCatalogosByCodigo("REF-01");
            ViewBag.ListadoReferido = Referido;

            //Listado Intermediario
            var Intermediario = CatalogoEntity.ObtenerListadoCatalogosByCodigo("INT-01");
            ViewBag.ListadoIntermediario = Intermediario;

            //Listado TipoZoho
            var TipoZoho = CatalogoEntity.ObtenerListadoCatalogosByCodigo("TZO-01");
            ViewBag.ListadoTipoZoho = TipoZoho;

            //Listado TipoCliente
            var TipoCliente = CatalogoEntity.ObtenerListadoCatalogosByCodigo("TCL-01");
            ViewBag.ListadoTipoCliente = TipoCliente;

            //Listado TamanioEmpresa
            var TamanioEmpresa = CatalogoEntity.ObtenerListadoCatalogosByCodigo("TEP-01");
            ViewBag.ListadoTamanioEmpresa = TamanioEmpresa;

            //Listado Sector
            var Sector = CatalogoEntity.ObtenerListadoCatalogosByCodigo("SCT-01");
            ViewBag.ListadoSector = Sector;

            //Listado Pais
            var Pais = CatalogoEntity.ObtenerListadoCatalogosByCodigo("PAI-01");
            ViewBag.ListadoPais = Pais;

            //Listado Ciudad
            var Ciudad = CatalogoEntity.ConsultarCatalogoPorPadre(cliente.pais.HasValue ? cliente.pais.Value : 0, "CIUDAD");//CatalogoEntity.ObtenerListadoCatalogosByCodigo("CIU-01");
            ViewBag.ListadoCiudad = Ciudad;

            //Listado PotenciaCrecimiento 
            var PotenciaCrecimiento = CatalogoEntity.ObtenerListadoCatalogosByCodigo("PTC-01");
            ViewBag.ListadoPotenciaCrecimiento = PotenciaCrecimiento;

            //Listado EstatusCliente
            var EstatusCliente = CatalogoEntity.ObtenerListadoCatalogosByCodigo("ECL-02");
            ViewBag.ListadoEstatusCliente = EstatusCliente;

            //Listado CategorizacionCliente
            var CategorizacionCliente = CatalogoEntity.ObtenerListadoCatalogosByCodigo("CCL-01");
            ViewBag.ListadoCategorizacionCliente = CategorizacionCliente;

            //Listado EtapaCliente
            var EtapaCliente = CatalogoEntity.ObtenerListadoCatalogosByCodigo("ETC-02");
            ViewBag.ListadoEtapaCliente = EtapaCliente;
            #endregion

            ViewBag.IngresosAnuales = cliente.ingresos_anuales_cliente.HasValue ? cliente.ingresos_anuales_cliente : 0;

            var contactosCliente = ContactoClienteEntity.ListadIdsContactosClientesByCliente(id.Value);
            ViewBag.idsContactosClientes = string.Join(",", contactosCliente);  //ContactoClienteEntity.ListadIdsContactosClientesByCliente(id.Value);

            var lineasNegocioCliente = ClienteEntity.ListarIdsLineasNegocioCliente(id.Value);
            ViewBag.LineasNegocioCliente = string.Join(",", lineasNegocioCliente);  //ContactoClienteEntity.ListadIdsContactosClientesByCliente(id.Value);

            ViewBag.IngresosAnuales = cliente.ingresos_anuales_cliente;

            if (cliente == null)
            {
                return HttpNotFound();
            }
            return View(cliente);
        }

        public ActionResult _Categorizacion()
        {
            //Titulo de la Pantalla
            ViewBag.TituloModal = Etiquetas.TituloPanelCategorizacion;
            return PartialView();
        }

        public ActionResult GetDependientesPais(int id)
        {
            return Json(new { DataPaises = CatalogoEntity.ConsultarCatalogoPorPadre(id, "CIUDAD"), Prefijo = CatalogoEntity.ConsultarCatalogoPorPadre(id, "PREFIJO") }, JsonRequestBehavior.AllowGet);
        }

        public JsonResult GetDependientesEtapaClienteEtapaGeneral(int id)
        {
            var catalogo = CatalogoEntity.ListadoCatalogosEtapaCliente(id).Where(s => s.codigo_catalogo == "ETAPA-GENERAL").Select(o => new Oitem(o.id_catalogo, o.nombre_catalgo));
            return Json(catalogo);
        }

        public JsonResult GetDependientesEtapaClienteEstatusDetallado(int id)
        {
            var catalogo = CatalogoEntity.ListadoCatalogosEtapaCliente(id).Where(s => s.codigo_catalogo == "ESTATUS-DETALLADO").Select(o => new Oitem(o.id_catalogo, o.nombre_catalgo)); ;
            return Json(catalogo);
        }

        public JsonResult GetDependientesEtapaClienteEstatusGeneral(int id)
        {
            var catalogo = CatalogoEntity.ListadoCatalogosEtapaCliente(id).Where(s => s.codigo_catalogo == "ESTATUS-GENERAL").Select(o => new Oitem(o.id_catalogo, o.nombre_catalgo)); ;
            return Json(catalogo);
        }

        public JsonResult GetDependientesEtapaCliente(int? id)
        {
            CatalogoEntity.ListadoCatalogosEtapaCliente(id.Value);
            return Json(CatalogoEntity.ListadoCatalogosEtapaCliente(id.Value));
        }

        public JsonResult GetContactosCliente(int? id)
        {
            if (id.HasValue)
            {
                var items = ContactoClienteEntity.ListarContactosClientesByCliente(id.Value)
    .Select(o => new Oitem(o.id_contacto, o.nombre_contacto + " " + o.apellido_contacto));

                return Json(items);
            }
            else
                return Json(new List<Oitem>());

        }

        public JsonResult GetDependientesTipoZoho(int id)
        {
            CatalogoEntity.ListadoCatalogosEtapaClienteTipoZoho(id);
            return Json(CatalogoEntity.ListadoCatalogosEtapaClienteTipoZoho(id));
        }

        [HttpPost]
        public ActionResult Edit(ClienteContactos cliente, List<int> contactosCliente, List<int> lineasNegocioCliente, bool? usar_organigrama)
        {
            try
            {
                //Validacion solo para país Ecuador
                if (!Validaciones.VerificaIdentificacion(cliente.ruc_ci_cliente) && cliente.pais == 233)
                    return Json(new { Resultado = new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeCedulaRucIncorrecto } }, JsonRequestBehavior.AllowGet);

                if (!ClienteEntity.VerificarRUCCedulaExistente(cliente.ruc_ci_cliente, cliente.id_cliente))
                    return Json(new { Resultado = new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeValidacionCedulaRucRegistrado } }, JsonRequestBehavior.AllowGet);

                RespuestaTransaccion resultado = ClienteEntity.ActualizarCliente(new Cliente
                {
                    id_cliente = cliente.id_cliente,
                    referido = cliente.referido,
                    intermediario = cliente.intermediario,
                    tipo_zoho = cliente.tipo_zoho,
                    tipo_cliente = cliente.tipo_cliente,
                    tamanio_empresa = cliente.tamanio_empresa,
                    etapa_cliente = cliente.etapa_cliente,
                    potencial_crecimiento = cliente.potencial_crecimiento,
                    categorizacion_cliente = cliente.categorizacion_cliente,
                    ruc_ci_cliente = cliente.ruc_ci_cliente,
                    razon_social_cliente = cliente.razon_social_cliente,
                    nombre_comercial_cliente = cliente.nombre_comercial_cliente,
                    ingresos_anuales_cliente = cliente.ingresos_anuales_cliente,
                    sector = cliente.sector,
                    pais = cliente.pais,
                    ciudad = cliente.ciudad,
                    direccion_cliente = cliente.direccion_cliente,
                    estado_cliente = cliente.estado_cliente,
                    usar_organigrama = usar_organigrama.Value,
                    usuario_id = cliente.usuario_id
                }, contactosCliente, lineasNegocioCliente);


                return Json(new { Resultado = resultado }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                return Json(new { Resultado = new RespuestaTransaccion { Estado = false, Respuesta = ex.Message } }, JsonRequestBehavior.AllowGet);
            }
        }

        public JsonResult GetContactos(int? id)
        {
            var items = ContactoClienteEntity.ListarContactos()
                .Select(o => new Oitem(o.id_contacto, o.nombre_contacto + " " + o.apellido_contacto + " - " + o.TextoCatalogoContacto));

            return Json(items);
        }

        public JsonResult GetContactosFacturacion(int? id)
        {
            List<Oitem> items = new List<Oitem>();
            if (id.HasValue)
            {
                items = ContactoClienteEntity.ListarContactosFacturacion().Where(s => s.idCliente == id.Value)
    .Select(o => new Oitem(o.id_contacto, o.nombre_contacto + " " + o.apellido_contacto)).ToList();
            }
            return Json(items);
        }


        public JsonResult _GetContactosFacturacion()
        {
            List<MultiSelectJQueryUi> items = new List<MultiSelectJQueryUi>();

            items = ContactoClienteEntity.ListarContactos()
.Select(o => new MultiSelectJQueryUi(o.id_contacto, o.nombre_contacto + " " + o.apellido_contacto, o.TextoCatalogoContacto)).ToList();
            return Json(items, JsonRequestBehavior.AllowGet);
        }

        public JsonResult _GetCatalogoLineasNegocioCliente()
        {
            List<MultiSelectJQueryUi> items = new List<MultiSelectJQueryUi>();

            items = CatalogoEntity.ListadoCatalogosPorCodigo("LNG-CLI-01")
.Select(o => new MultiSelectJQueryUi(Convert.ToInt64(o.Value), o.Text, "")).ToList();
            return Json(items, JsonRequestBehavior.AllowGet);
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

        //public ActionResult DescargarReporteFormatoExcel()
        //{
        //    //Seleccionar las columnas a exportar
        //    var collection = ClienteEntity.ListarCliente();
        //    var package = new ExcelPackage();

        //    package = Reportes.ExportarExcel(collection.Cast<object>().ToList(),"Clientes");
        //    return File(package.GetAsByteArray(), XlsxContentType, "ListadoCliente.xlsx");
        //}

        public ActionResult DescargarReporteFormatoExcel()
        {
            var collection = ClienteEntity.ListarCliente();

            ExcelPackage package = new ExcelPackage();
            var workSheet = package.Workbook.Worksheets.Add("Listado");
            workSheet.TabColor = System.Drawing.Color.Black;
            workSheet.DefaultRowHeight = 12;

            List<string> columnas = new List<string>() {
                "ID",
                "USUARIO",
                "RUC",
                "RAZON SOCIAL",
                "NOMBRE COMERCIAL",
                "TAMAÑO EMPRESA",
                "SECTOR",
                "PAIS",
                "CIUDAD",
                "DIRECCION",
                "POTENCIAL CRECIMIENTO",
                "CATEGORIZACIÓN CLIENTE",
                "ETAPA CLIENTE",
                "TIPO CLIENTE",
                "TIPO CRM",
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
                workSheet.Cells[recordIndex, 2].Value = item.Usuario;
                workSheet.Cells[recordIndex, 3].Value = item.RUC;
                workSheet.Cells[recordIndex, 4].Value = item.Razon_Social;
                workSheet.Cells[recordIndex, 5].Value = item.Nombre_Comercial;
                workSheet.Cells[recordIndex, 6].Value = item.Tamanio_Empresa;
                workSheet.Cells[recordIndex, 7].Value = item.sector;
                workSheet.Cells[recordIndex, 8].Value = item.Pais;
                workSheet.Cells[recordIndex, 9].Value = item.ciudad;
                workSheet.Cells[recordIndex, 10].Value = item.Direccion;
                workSheet.Cells[recordIndex, 11].Value = item.PotencialCrecimiento;
                workSheet.Cells[recordIndex, 12].Value = item.CategorizacionCliente;
                workSheet.Cells[recordIndex, 13].Value = item.EtapaCliente;
                workSheet.Cells[recordIndex, 14].Value = item.Tipo_Cliente;
                workSheet.Cells[recordIndex, 15].Value = item.Tipo_Zoho;
                workSheet.Cells[recordIndex, 16].Value = item.Estado;
                recordIndex++;
            }

            for (int i = 1; i <= columnas.Count; i++)
            {
                workSheet.Column(i).AutoFit();
                workSheet.Column(i).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            }


            return File(package.GetAsByteArray(), XlsxContentType, "ListadoCliente.xlsx");
        }

        public ActionResult DescargarReporteFormatoPDF()
        {
            // Seleccionar las columnas a exportar
            var results = ClienteEntity.ListarCliente();

            var list = Reportes.SerializeToJSON(results);
            return Content(list, "application/json");
        }

        public ActionResult DescargarReporteFormatoCSV()
        {
            var comlumHeadrs = new string[]
            {
                "ID",
                "USUARIO",
                "RUC",
                "RAZON SOCIAL",
                "NOMBRE COMERCIAL",
                "TAMAÑO EMPRESA",
                "SECTOR",
                "PAIS",
                "CIUDAD",
                "DIRECCION",
                "POTENCIAL CRECIMIENTO",
                "CATEGORIZACIÓN CLIENTE",
                "ETAPA CLIENTE",
                "TIPO CLIENTE",
                "TIPO CRM",
                "ESTADO"
            };

            var listado = (from item in ClienteEntity.ListarCliente()
                           select new object[]
                           {
                                item.Id,
                                item.Usuario,
                                item.RUC,
                                item.Razon_Social,
                                item.Nombre_Comercial,
                                item.Tamanio_Empresa,
                                item.sector,
                                item.Pais,
                                item.ciudad,
                                item.Direccion,
                                item.PotencialCrecimiento,
                                item.CategorizacionCliente,
                                item.EtapaCliente,
                                item.Tipo_Cliente,
                                item.Tipo_Zoho,
                                item.Estado
                           }).ToList();

            // Build the file content
            var employeecsv = new StringBuilder();
            listado.ForEach(line =>
            {
                employeecsv.AppendLine(string.Join(",", line));
            });

            byte[] buffer = Encoding.Default.GetBytes($"{string.Join(",", comlumHeadrs)}\r\n{employeecsv.ToString()}");
            return File(buffer, "text/csv", $"Clientes.csv");
        }

        #region Funcionalidad de carga Masiva
        public ActionResult CargarData()
        {
            List<CargaMasiva> errores = new List<CargaMasiva>();
            RespuestaTransaccion resultado = new RespuestaTransaccion { Estado = false };
            List<ClienteExcel> listado = new List<ClienteExcel>();
            string FileName = "";

            try
            {
                HttpFileCollectionBase files = Request.Files;
                for (int i = 0; i < files.Count; i++)
                {
                    //string path = AppDomain.CurrentDomain.BaseDirectory + "Uploads/";    
                    //string filename = Path.GetFileName(Request.Files[i].FileName);    

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

                    file.SaveAs(path);

                    errores = VerificarEstructuraExcel(path); // Devuelve los errores en la estructura

                    if (errores.Count == 0)
                    {
                        listado = ToEntidadHojaExcelList(path); // Convierte el documento a un Listado tipo entidad

                        // Validacion de la columna N
                        if (listado == null)
                        {
                            errores.Add(new CargaMasiva { Fila = 0, Columna = 0, Valor = "Ninguno", Error = "Verificar la columna de numeración (N.-) del archivo." });
                            return Json(new { Resultado = new RespuestaTransaccion { Estado = resultado.Estado, Respuesta = Mensajes.MensajeCargaMasivaFallida }, Errores = errores, Archivo = FileName }, JsonRequestBehavior.AllowGet);
                        }

                        //Validar que la columna de Numeración no se encuentre duplicada
                        var numeracionDuplicada = listado.GroupBy(x => x.N).Any(g => g.Count() > 1);

                        if (numeracionDuplicada)
                        {
                            errores.Add(new CargaMasiva { Fila = 0, Columna = 0, Valor = "Ninguno", Error = "Valores de la Columna de numeración 'N.-' no pueden ser duplicados." });
                            return Json(new { Resultado = new RespuestaTransaccion { Estado = resultado.Estado, Respuesta = Mensajes.MensajeCargaMasivaFallida }, Errores = errores, Archivo = FileName }, JsonRequestBehavior.AllowGet);
                        }

                        if (listado.Count > 0)
                            resultado = ClienteEntity.CrearActualizarClienteCargaMasiva(listado); // Crea los clientes
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
                    return Json(new { Resultado = new RespuestaTransaccion { Estado = resultado.Estado, Respuesta = resultado.Respuesta + Mensajes.MensajeCargaMasivaFallida }, Errores = errores, Archivo = FileName }, JsonRequestBehavior.AllowGet);

            }
            catch (Exception ex)
            {
                return Json(new { Resultado = new RespuestaTransaccion { Estado = false, Respuesta = ex.Message } }, JsonRequestBehavior.AllowGet);
            }
        }

        private List<ClienteExcel> ToEntidadHojaExcelList(string pathDelFicheroExcel)
        {
            List<ClienteExcel> listadoCargaMasiva = new List<ClienteExcel>();
            try
            {

                var user = ViewData["usuario"] = System.Web.HttpContext.Current.Session["usuario"];
                var usuarioID = Convert.ToInt16(user);

                FileInfo existingFile = new FileInfo(pathDelFicheroExcel);

                using (ExcelPackage package = new ExcelPackage(existingFile))
                {
                    // get the first worksheet in the workbook
                    ExcelWorksheet worksheet = package.Workbook.Worksheets["CLIENTES"];
                    ExcelWorksheet hojaContactos = package.Workbook.Worksheets["CONTACTOS"];
                    ExcelWorksheet hojaLineasNegocio = package.Workbook.Worksheets["LINEAS NEGOCIO"];

                    if (worksheet != null)
                    {
                        int colCount = worksheet.Dimension.End.Column;  //get Column Count
                        int rowCount = worksheet.Dimension.End.Row;     //get row count
                        for (int row = 2; row <= rowCount; row++)
                        {
                            listadoCargaMasiva.Add(new ClienteExcel
                            {
                                N = !string.IsNullOrEmpty((worksheet.GetValue(row, 1) ?? "").ToString().Trim()) ? int.Parse((worksheet.GetValue(row, 1) ?? "").ToString().Trim()) : 0,
                                referido = !string.IsNullOrEmpty((worksheet.GetValue(row, 2) ?? "").ToString().Trim()) ? int.Parse((worksheet.GetValue(row, 2) ?? "").ToString().Trim()) : 0,
                                intermediario = !string.IsNullOrEmpty((worksheet.GetValue(row, 3) ?? "").ToString().Trim()) ? int.Parse((worksheet.GetValue(row, 3) ?? "").ToString().Trim()) : 0,
                                tipo_cliente = !string.IsNullOrEmpty((worksheet.GetValue(row, 4) ?? "").ToString().Trim()) ? int.Parse((worksheet.GetValue(row, 4) ?? "").ToString().Trim()) : 0,
                                tipo_zoho = !string.IsNullOrEmpty((worksheet.GetValue(row, 5) ?? "").ToString().Trim()) ? int.Parse((worksheet.GetValue(row, 5) ?? "").ToString().Trim()) : 0,
                                potencial_crecimiento = !string.IsNullOrEmpty((worksheet.GetValue(row, 6) ?? "").ToString().Trim()) ? int.Parse((worksheet.GetValue(row, 6) ?? "").ToString().Trim()) : 0,
                                categorizacion_cliente = !string.IsNullOrEmpty((worksheet.GetValue(row, 7) ?? "").ToString().Trim()) ? int.Parse((worksheet.GetValue(row, 7) ?? "").ToString().Trim()) : 0,
                                etapa_cliente = !string.IsNullOrEmpty((worksheet.GetValue(row, 8) ?? "").ToString().Trim()) ? int.Parse((worksheet.GetValue(row, 8) ?? "").ToString().Trim()) : 0,
                                ruc_ci_cliente = !string.IsNullOrEmpty((worksheet.GetValue(row, 9) ?? "").ToString().Trim()) ? (worksheet.GetValue(row, 9) ?? "").ToString().Trim() : string.Empty,//(worksheet.GetValue(row, 9) ?? "").ToString().Trim(),
                                razon_social_cliente = !string.IsNullOrEmpty((worksheet.GetValue(row, 10) ?? "").ToString().Trim()) ? (worksheet.GetValue(row, 10) ?? "").ToString().Trim() : string.Empty,// (worksheet.GetValue(row, 10) ?? "").ToString().Trim(),
                                nombre_comercial_cliente = !string.IsNullOrEmpty((worksheet.GetValue(row, 11) ?? "").ToString().Trim()) ? (worksheet.GetValue(row, 11) ?? "").ToString().Trim() : string.Empty,//(worksheet.GetValue(row, 11) ?? "").ToString().Trim(),
                                tamanio_empresa = !string.IsNullOrEmpty((worksheet.GetValue(row, 12) ?? "").ToString().Trim()) ? int.Parse((worksheet.GetValue(row, 12) ?? "").ToString().Trim()) : 0,
                                sector = !string.IsNullOrEmpty((worksheet.GetValue(row, 13) ?? "").ToString().Trim()) ? int.Parse((worksheet.GetValue(row, 13) ?? "").ToString().Trim()) : 0,
                                ingresos_anuales_cliente = !string.IsNullOrEmpty((worksheet.GetValue(row, 14) ?? "").ToString().Trim()) ? decimal.Parse((worksheet.GetValue(row, 14) ?? "").ToString().Trim()) : 0,
                                pais = !string.IsNullOrEmpty((worksheet.GetValue(row, 15) ?? "").ToString().Trim()) ? int.Parse((worksheet.GetValue(row, 15) ?? "").ToString().Trim()) : 0,
                                ciudad = !string.IsNullOrEmpty((worksheet.GetValue(row, 16) ?? "").ToString().Trim()) ? int.Parse((worksheet.GetValue(row, 16) ?? "").ToString().Trim()) : 0,
                                direccion_cliente = !string.IsNullOrEmpty((worksheet.GetValue(row, 17) ?? "").ToString().Trim()) ? (worksheet.GetValue(row, 17) ?? "").ToString().Trim() : string.Empty,//(worksheet.GetValue(row, 17) ?? "").ToString().Trim(),

                                usuario_id = usuarioID
                            });
                        }
                    }

                    List<ContactosClientes> listadoContactos = new List<ContactosClientes>();
                    if (hojaContactos != null)
                    {
                        int colCount = hojaContactos.Dimension.End.Column;  //get Column Count
                        int rowCount = hojaContactos.Dimension.End.Row;     //get row count
                        for (int row = 2; row <= rowCount; row++)
                        {
                            listadoContactos.Add(new ContactosClientes
                            {
                                idCliente = !string.IsNullOrEmpty((hojaContactos.GetValue(row, 1) ?? "").ToString().Trim()) ? int.Parse((hojaContactos.GetValue(row, 1) ?? "").ToString().Trim()) : 0,
                                idContacto = !string.IsNullOrEmpty((hojaContactos.GetValue(row, 2) ?? "").ToString().Trim()) ? int.Parse((hojaContactos.GetValue(row, 2) ?? "").ToString().Trim()) : 0,
                            });
                        }
                    }

                    foreach (var item in listadoCargaMasiva)
                    {
                        item.contactos = listadoContactos.Where(s => s.idCliente == item.N).ToList();
                    }


                    List<ClienteLineaNegocio> listadoLineasNegocio = new List<ClienteLineaNegocio>();
                    if (hojaLineasNegocio != null)
                    {
                        int colCount = hojaLineasNegocio.Dimension.End.Column;  //get Column Count
                        int rowCount = hojaLineasNegocio.Dimension.End.Row;     //get row count
                        for (int row = 2; row <= rowCount; row++)
                        {
                            listadoLineasNegocio.Add(new ClienteLineaNegocio
                            {
                                ClienteID = !string.IsNullOrEmpty((hojaLineasNegocio.GetValue(row, 1) ?? "").ToString().Trim()) ? int.Parse((hojaLineasNegocio.GetValue(row, 1) ?? "").ToString().Trim()) : 0,
                                CatalogoLineaNegocioID = !string.IsNullOrEmpty((hojaLineasNegocio.GetValue(row, 2) ?? "").ToString().Trim()) ? int.Parse((hojaLineasNegocio.GetValue(row, 2) ?? "").ToString().Trim()) : 0,
                            });
                        }
                    }

                    var maximoHojaCodigo = listadoCargaMasiva.Select(s => s.N).Max();
                    var maximoHojaLineasNegocio = listadoLineasNegocio.Select(s => s.ClienteID).Max();

                    if (maximoHojaCodigo != maximoHojaLineasNegocio)
                    {
                        return null;
                    }

                    foreach (var item in listadoCargaMasiva)
                    {
                        item.lineasNegocio = listadoLineasNegocio.Where(s => s.ClienteID == item.N).ToList();
                    }

                }

                return listadoCargaMasiva;
            }
            catch (Exception ex)
            {
                return new List<ClienteExcel>();
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
                    ExcelWorksheet hojaLineasNegocio = package.Workbook.Worksheets["LINEAS NEGOCIO"];


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

                                int pais = int.Parse((worksheet.Cells[row, 15].Value ?? "").ToString().Trim()); // El tipo de contacto se encuentra en la ultima columna

                                error = ValidarCamposCargaMasivaCliente(columna, valorColumna, pais);//columna == "PAIS" ? ValidarCamposCargaMasivaCliente(columna, valorColumna, pais) : ValidarCamposCargaMasivaCliente(columna, valorColumna); // Solo para validar el MAIL se pasa el tipo de Contacto

                                //error = ValidarCamposCargaMasivaCliente(columna, valorColumna);

                                if (!string.IsNullOrEmpty(error))
                                {
                                    errores.Add(new CargaMasiva { Fila = row, Columna = col, Valor = valorColumna, Error = error });
                                }
                            }
                        }
                    }

                    if (hojaLineasNegocio != null)
                    {
                        int colCount = hojaLineasNegocio.Dimension.End.Column;  //get Column Count
                        int rowCount = hojaLineasNegocio.Dimension.End.Row;      //get row count - Cabecera
                        for (int row = 2; row <= rowCount; row++)
                        {
                            for (int col = 1; col <= colCount; col++)
                            {
                                var error = string.Empty;
                                string columna = (hojaLineasNegocio.Cells[1, col].Value ?? "").ToString().Trim(); // Nombre de la Columna
                                string valorColumna = (hojaLineasNegocio.Cells[row, col].Value ?? "").ToString().Trim();

                                error = ValidarCamposCargaMasivaCliente(columna, valorColumna);

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
        private static string ValidarCamposCargaMasivaCliente(string columna, string valor, int? pais = null)
        {
            var excepcion = "El campo {0} con el valor {1} , contiene errores. Error: {2}";
            try
            {
                valor = !string.IsNullOrEmpty(valor) ? valor.Trim() : "´Vacío.´";

                bool esNumero;
                int longitudCaracteres = 0;
                var error = "El campo {0} con el valor {1} ,tiene una extensión de caracteres mayor a la permitida. Solo permite {2} caracteres";

                switch (columna)
                {
                    case "N.-":
                        esNumero = int.TryParse(valor, out int valorNumeracion);
                        if (!esNumero)
                            error = "Tipo de dato incorrecto para el campo " + columna + " , con el valor " + valor + " .Solo admite valores numéricos de tipo entero.";
                        else
                            error = null;
                        break;
                    case "REFERIDO":
                        esNumero = int.TryParse(valor, out int valorReferido);
                        if (!esNumero)
                            error = "Tipo de dato incorrecto para el campo " + columna + " , con el valor " + valor + " .Solo admite valores numéricos de tipo entero.";
                        else
                        {
                            if (!CatalogoEntity.VerificarExistenciaCatalogo(int.Parse(valor)))
                                error = "El valor no es válido para el campo " + columna + " , con el valor " + valor + " . Verificar validez catálogo.";
                            else
                                error = null;
                        }

                        break;
                    case "INTERMEDIARIO":
                        esNumero = int.TryParse(valor, out int valorIntermediario);
                        if (!esNumero)
                            error = "Tipo de dato incorrecto para el campo " + columna + " , con el valor " + valor + " .Solo admite valores numéricos de tipo entero.";
                        else
                        {
                            if (!CatalogoEntity.VerificarExistenciaCatalogo(int.Parse(valor)))
                                error = "El valor no es válido para el campo " + columna + " , con el valor " + valor + " . Verificar validez catálogo.";
                            else
                                error = null;
                        }
                        break;
                    case "TIPO DE CLIENTE":
                        esNumero = int.TryParse(valor, out int valorTipoCliente);
                        if (!esNumero)
                            error = "Tipo de dato incorrecto para el campo " + columna + " , con el valor " + valor + " .Solo admite valores numéricos de tipo entero.";
                        else
                        {
                            if (!CatalogoEntity.VerificarExistenciaCatalogo(int.Parse(valor)))
                                error = "El valor no es válido para el campo " + columna + " , con el valor " + valor + " . Verificar validez catálogo.";
                            else
                                error = null;
                        }

                        break;
                    case "TIPO CRM":
                        esNumero = int.TryParse(valor, out int valorTipoZoho);
                        if (!esNumero)
                            error = "Tipo de dato incorrecto para el campo " + columna + " , con el valor " + valor + " .Solo admite valores numéricos de tipo entero.";
                        else
                        {
                            if (!CatalogoEntity.VerificarExistenciaCatalogo(int.Parse(valor)))
                                error = "El valor no es válido para el campo " + columna + " , con el valor " + valor + " . Verificar validez catálogo.";
                            else
                                error = null;
                        }
                        break;
                    case "POTENCIAL DE CRECIMIENTO":
                        esNumero = int.TryParse(valor, out int valorPotencialCrecimiento);
                        if (!esNumero)
                            error = "Tipo de dato incorrecto para el campo " + columna + " , con el valor " + valor + " .Solo admite valores numéricos de tipo entero.";
                        else
                        {
                            if (!CatalogoEntity.VerificarExistenciaCatalogo(int.Parse(valor)))
                                error = "El valor no es válido para el campo " + columna + " , con el valor " + valor + " . Verificar validez catálogo.";
                            else
                                error = null;
                        }
                        break;
                    case "CATEGORIZACION DE CLIENTE":
                        esNumero = int.TryParse(valor, out int valorCategorizacionCliente);
                        if (!esNumero)
                            error = "Tipo de dato incorrecto para el campo " + columna + " , con el valor " + valor + " .Solo admite valores numéricos de tipo entero.";
                        else
                        {
                            if (!CatalogoEntity.VerificarExistenciaCatalogo(int.Parse(valor)))
                                error = "El valor no es válido para el campo " + columna + " , con el valor " + valor + " . Verificar validez catálogo.";
                            else
                                error = null;
                        }
                        break;
                    case "ETAPA DE CLIENTE":
                        esNumero = int.TryParse(valor, out int valorEtapaCliente);
                        if (!esNumero)
                            error = "Tipo de dato incorrecto para el campo " + columna + " , con el valor " + valor + " .Solo admite valores numéricos de tipo entero.";
                        else
                        {
                            if (!CatalogoEntity.VerificarExistenciaCatalogo(int.Parse(valor)))
                                error = "El valor no es válido para el campo " + columna + " , con el valor " + valor + " . Verificar validez catálogo.";
                            else
                                error = null;
                        }
                        break;
                    case "IDENTIFICACION DE CLIENTE":
                        if (pais.Value == 233) // VERIFICA SIEMPRE Y CUANDO SEA ECUADOR EL PAIS -- CAMBIAR ID POR CODIGO DE CATALOGO DEL PAIS ECUADOR 
                        {
                            if (!Validaciones.VerificaIdentificacion(valor))
                                error = "El valor no es válido para el campo " + columna + " , con el valor " + valor + " . " + Mensajes.MensajeCedulaRucIncorrecto;
                            else
                                error = null;
                        }
                        else
                        {
                            error = null;
                        }
                        break;
                    case "RAZON SOCIAL DE CLIENTE":
                        longitudCaracteres = 250;
                        if (valor.Length > longitudCaracteres)
                        {
                            error = string.Format(error, columna, valor, longitudCaracteres);
                        }
                        else
                        {
                            error = null;
                        }
                        break;
                    case "NOMBRE COMERCIAL DE CLIENTE":
                        longitudCaracteres = 250;
                        if (valor.Length > longitudCaracteres)
                        {
                            error = string.Format(error, columna, valor, longitudCaracteres);
                        }
                        else
                        {
                            error = null;
                        }
                        break;
                    case "TAMAÑO DE EMPRESA":
                        esNumero = int.TryParse(valor, out int valorTamanioEmpresa);
                        if (!esNumero)
                            error = "Tipo de dato incorrecto para el campo " + columna + " , con el valor " + valor + " .Solo admite valores numéricos de tipo entero.";
                        else
                        {
                            if (!CatalogoEntity.VerificarExistenciaCatalogo(int.Parse(valor)))
                                error = "El valor no es válido para el campo " + columna + " , con el valor " + valor + " . Verificar validez catálogo.";
                            else
                                error = null;
                        }
                        break;
                    case "SECTOR":
                        esNumero = int.TryParse(valor, out int valorSector);
                        if (!esNumero)
                            error = "Tipo de dato incorrecto para el campo " + columna + " , con el valor " + valor + " .Solo admite valores numéricos de tipo entero.";
                        else
                        {
                            if (!CatalogoEntity.VerificarExistenciaCatalogo(int.Parse(valor)))
                                error = "El valor no es válido para el campo " + columna + " , con el valor " + valor + " . Verificar validez catálogo.";
                            else
                                error = null;
                        }
                        break;
                    case "INGRESOS ANUALES DE CLIENTE":
                        //var separadorIncorrecto = valor.IndexOf(".");

                        //if (separadorIncorrecto != -1)
                        //    valor.Replace(".", ",");

                        esNumero = decimal.TryParse(valor, out decimal valorIngresosAnualesCliente);
                        if (!esNumero)
                            error = "Tipo de dato incorrecto para el campo " + columna + " , con el valor " + valor + " .Solo admite valores numéricos de tipo entero.";
                        else
                            error = null;
                        break;

                    case "PAIS":
                        esNumero = int.TryParse(valor, out int valorPais);
                        if (!esNumero)
                            error = "Tipo de dato incorrecto para el campo " + columna + " , con el valor " + valor + " .Solo admite valores numéricos de tipo entero.";
                        else
                        {
                            if (!CatalogoEntity.VerificarExistenciaCatalogo(int.Parse(valor)))
                                error = "El valor no es válido para el campo " + columna + " , con el valor " + valor + " . Verificar validez catálogo.";
                            else
                                error = null;
                        }
                        break;

                    case "LINEA DE NEGOCIO":
                        esNumero = int.TryParse(valor, out int valorLineaNegocio);
                        if (!esNumero)
                            error = "Tipo de dato incorrecto para el campo " + columna + " , con el valor " + valor + " .Solo admite valores numéricos de tipo entero.";
                        else
                        {
                            if (!CatalogoEntity.VerificarExistenciaCatalogoLineaNegocio(int.Parse(valor)))
                                error = "El valor no es válido para el campo " + columna + " , con el valor " + valor + " . Verificar validez catálogo.";
                            else
                                error = null;
                        }
                        break;

                    case "CIUDAD":
                        esNumero = int.TryParse(valor, out int valorCiudad);
                        if (!esNumero)
                            error = "Tipo de dato incorrecto para el campo " + columna + " , con el valor " + valor + " .Solo admite valores numéricos de tipo entero.";
                        else
                        {
                            if (!CatalogoEntity.VerificarExistenciaCatalogo(int.Parse(valor)))
                                error = "El valor no es válido para el campo " + columna + " , con el valor " + valor + " . Verificar validez catálogo.";
                            else
                                error = null;
                        }
                        break;
                    case "DIRECCION DE CLIENTE":
                        longitudCaracteres = 500;
                        if (valor.Length > longitudCaracteres)
                        {
                            error = string.Format(error, columna, valor, longitudCaracteres);
                        }
                        else
                        {
                            error = null;
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

    }
}

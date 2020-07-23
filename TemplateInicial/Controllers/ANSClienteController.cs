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
    public class ANSClienteController : BaseAppController
    {
        private const string XlsxContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

        // GET: ANSCliente
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
            ViewBag.NombreListado = Etiquetas.TituloGridANSCliente;

            //Controlar permisos
            var user = ViewData["usuario"] = System.Web.HttpContext.Current.Session["usuario"];
            var usuario = int.Parse(user.ToString());
            string nombreControlador = ControllerContext.RouteData.Values["controller"].ToString();
            ViewBag.NombreControlador = nombreControlador;

            ViewBag.AccionesUsuario = ManejoPermisosEntity.ListadoAccionesCatalogoUsuario(usuario, nombreControlador);

            //Obtener Acciones del controlador
            ViewBag.AccionesControlador = GetMetodosControlador(nombreControlador);

            //Búsqueda
            var listado = ANSClienteEntity.ListarANSCliente();

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

        // GET: ANSCliente/Details/5
        public ActionResult Details(int id)
        {
            return View();
        }

        // GET: ANSCliente/Create
        public ActionResult Create()
        {
            //Listado Clientes
            var listadoClientes = ClienteEntity.ObtenerListadoCliente();
            ViewBag.ListadoCliente = listadoClientes;

            //Listado de Tipo de Solicitud
            var Tipo = CatalogoEntity.ObtenerListadoCatalogosByCodigo("TSL-01");
            ViewBag.ListadoTipo = Tipo;

            return View();
        }

        // POST: ANSCliente/Create
        [HttpPost]
        public ActionResult Create(ANSCliente ansCliente)
        {
            //Listado Clientes
            var listadoClientes = ClienteEntity.ObtenerListadoCliente();
            ViewBag.ListadoCliente = listadoClientes;

            //Listado de Tipo de Solicitud
            var Tipo = CatalogoEntity.ObtenerListadoCatalogosByCodigo("TSL-01");
            ViewBag.ListadoTipo = Tipo;

            try
            {
                string tipoRequerimiento = (ansCliente.tipo_requerimiento ?? string.Empty).ToLower().Trim();

                var tipoRequerimientoCliente = ANSClienteEntity.ListarANSCliente().Where(s => (s.TipoRequerimiento ?? string.Empty).ToLower().Trim() == ansCliente.tipo_requerimiento && s.idCliente == ansCliente .id_cliente ).ToList();

                if (tipoRequerimientoCliente.Count > 0)
                    return Json(new { Resultado = new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeResgistroExistente } }, JsonRequestBehavior.AllowGet);
                 
                RespuestaTransaccion resultado = ANSClienteEntity.CrearANSCliente(ansCliente);

                return Json(new { Resultado = resultado }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                return Json(new { Resultado = new RespuestaTransaccion { Estado = false, Respuesta = ex.Message } }, JsonRequestBehavior.AllowGet);
            }
        }

        // GET: ANSCliente/Edit/5
        public ActionResult Edit(int id)
        {
            //Listado Clientes
            var listadoClientes = ClienteEntity.ObtenerListadoCliente();
            ViewBag.ListadoCliente = listadoClientes;

            //Listado de Tipo de Solicitud
            var Tipo = CatalogoEntity.ObtenerListadoCatalogosByCodigo("TSL-01");
            ViewBag.ListadoTipo = Tipo;

            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            var ansCliente = ANSClienteEntity.ConsultarANSCliente(id);
            ViewBag.TiempoRespuesta = ansCliente.tiempo_atencion_min.HasValue ? ansCliente.tiempo_atencion_min : 0;

            if (ansCliente == null)
            {
                return HttpNotFound();
            }
            return View(ansCliente);  
        }

        // POST: ANSCliente/Edit/5
        [HttpPost]
        public ActionResult Edit(ANSCliente ansCliente)
        {
            //Listado Clientes
            var listadoClientes = ClienteEntity.ObtenerListadoCliente();
            ViewBag.ListadoCliente = listadoClientes;

            //Listado de Tipo de Solicitud
            var Tipo = CatalogoEntity.ObtenerListadoCatalogosByCodigo("TSL-01");
            ViewBag.ListadoTipo = Tipo;

            try
            {
                string tipoRequerimiento = (ansCliente.tipo_requerimiento ?? string.Empty).ToLower().Trim();

                var tipoRequerimientoCliente = ANSClienteEntity.ListarANSCliente().Where(s => (s.TipoRequerimiento ?? string.Empty).ToLower().Trim() == ansCliente.tipo_requerimiento && s.idCliente == ansCliente.id_cliente && s.Codigo != ansCliente.id_ans_sla).ToList();
                 
                if (tipoRequerimientoCliente.Count > 0)
                    return Json(new { Resultado = new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeValidacionNombreTarifario } }, JsonRequestBehavior.AllowGet);
                 
                RespuestaTransaccion resultado = ANSClienteEntity.ActualizarANSCliente(ansCliente);

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
            RespuestaTransaccion resultado = ANSClienteEntity.EliminarANSCliente(id);// await db.Cabecera.FindAsync(id);

            //return RedirectToAction("Index");
            return Json(new { Resultado = resultado }, JsonRequestBehavior.AllowGet);

        }

        public ActionResult DescargarReporteFormatoExcel()
        {
            var collection = ANSClienteEntity.ListarANSCliente();

            ExcelPackage package = new ExcelPackage();
            var workSheet = package.Workbook.Worksheets.Add("SLA/ANS Cliente");
            workSheet.TabColor = System.Drawing.Color.Black;
            workSheet.DefaultRowHeight = 12;

            List<string> columnas = new List<string>() {
                "ID",
                "CLIENTE",
                "RUC",
                "TIPO SOLICITUD",
                "TIPO REQUERIMENTO",
                "DETALLE",
                "TIEMPO (MIN)",
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
                workSheet.Cells[recordIndex, 2].Value = item.Cliente;
                workSheet.Cells[recordIndex, 3].Value = item.RUC;
                workSheet.Cells[recordIndex, 4].Value = item.TipoSolicitud;
                workSheet.Cells[recordIndex, 5].Value = item.TipoRequerimiento;
                workSheet.Cells[recordIndex, 6].Value = item.Detalle;
                workSheet.Cells[recordIndex, 7].Value = item.Tiempo;
                workSheet.Cells[recordIndex, 8].Value = item.EstadoANS;

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

            return File(package.GetAsByteArray(), XlsxContentType, "ListadoSLACliente.xlsx");

        }

        public ActionResult DescargarReporteFormatoPDF()
        {
            // Seleccionar las columnas a exportar
            var results = ANSClienteEntity.ListarANSCliente();

            var list = Reportes.SerializeToJSON(results);
            return Content(list, "application/json");
        }

        public ActionResult DescargarReporteFormatoCSV()
        {
            var comlumHeadrs = new string[]
            {
                "ID",
                "CLIENTE",
                "RUC",
                "TIPO SOLICITUD",
                "TIPO REQUERIMENTO",
                "DETALLE",
                "TIEMPO (MIN)",
                "ESTADO"
            };

            var listado = (from item in ANSClienteEntity.ListarANSCliente()
                           select new object[]
                           {
                               item.Codigo,
                               item.Cliente,
                               item.RUC,
                               item.TipoSolicitud,
                               item.TipoRequerimiento,
                               item.Detalle,
                               item.Tiempo,
                               item.EstadoANS

                           }).ToList();

            // Build the file content
            var employeecsv = new StringBuilder();
            listado.ForEach(line =>
            {
                employeecsv.AppendLine(string.Join(",", line));
            });

            byte[] buffer = Encoding.Default.GetBytes($"{string.Join(",", comlumHeadrs)}\r\n{employeecsv.ToString()}");
            return File(buffer, "text/csv", $"SLA-ANS-Cliente.csv");
        }
    }
}

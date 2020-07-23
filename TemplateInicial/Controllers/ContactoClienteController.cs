using GestionPPM.Entidades.Metodos;
using GestionPPM.Entidades.Modelo;
using GestionPPM.Repositorios;
using OfficeOpenXml;
using Omu.Awem.Helpers;
using Omu.AwesomeMvc;
using Seguridad.Helper;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using System.Web;
using System.Web.Mvc;

namespace TemplateInicial.Controllers
{
    [Autenticado]
    public class ContactoClienteController : BaseAppController
    {
        private GestionPPMEntities db = new GestionPPMEntities();
        private const string XlsxContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
        private int rolID = Convert.ToInt32(System.Web.HttpContext.Current.Session["rolID"]);

        // GET: ContactoCliente
        public ActionResult Index()
        {
            rolID = Convert.ToInt32(System.Web.HttpContext.Current.Session["rolID"]);
            ViewBag.PerfilesUsuario = PerfilesEntity.ListarPerfilesPorRol(rolID);

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

        // GET: ContactoCliente/Details/5
        public ActionResult Details(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            Contactos contactosClientes = db.Contactos.Find(id);
            if (contactosClientes == null)
            {
                return HttpNotFound();
            }
            return View(contactosClientes);
        }

        [HttpGet]
        public async Task<PartialViewResult> IndexGrid(String search)
        {
            ViewBag.NombreListado = Etiquetas.TituloGridContacto;
            //Controlar permisos
            var user = ViewData["usuario"] = System.Web.HttpContext.Current.Session["usuario"];
            var usuario = int.Parse(user.ToString());
            string nombreControlador = ControllerContext.RouteData.Values["controller"].ToString();
            ViewBag.NombreControlador = nombreControlador;

            ViewBag.AccionesUsuario = ManejoPermisosEntity.ListadoAccionesCatalogoUsuario(usuario, nombreControlador);

            //Obtener Acciones del controlador
            ViewBag.AccionesControlador = GetMetodosControlador(nombreControlador);

            //Búsqueda
            var listado = ContactoClienteEntity.ListadoContactosClientes();

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

        // GET: ContactoCliente/Create
        public ActionResult Create()
        {
            //Listado tipoContacto
            var tipoContacto = CatalogoEntity.ObtenerListadoCatalogosByCodigo("TCT-01");
            ViewBag.ListadoTipoContacto = tipoContacto;

            return View();
        }

        public ActionResult _Create(bool? flag, string prefijo, int? idCliente)
        {
            ViewBag.TituloModal = Etiquetas.TituloPanelCreacionContacto;
            //Listado tipoContacto
            var tipoContacto = CatalogoEntity.ListarCatalogo();
            if (flag.HasValue)
            {
                ViewBag.ListadoTipoContacto = flag.Value ? tipoContacto.Where(s => s.codigo_catalogo == "TCCLI-01").OrderBy(c => c.nombre_catalgo).Select(c => new SelectListItem
                {
                    Text = c.nombre_catalgo,
                    Value = c.id_catalogo.ToString(),
                }).ToList() : tipoContacto.Where(s => s.codigo_catalogo == "TCFACT-01").OrderBy(c => c.nombre_catalgo).Select(c => new SelectListItem
                {
                    Text = c.nombre_catalgo,
                    Value = c.id_catalogo.ToString(),
                }).ToList();
            }
            else
            {
                ViewBag.ListadoTipoContacto = tipoContacto.Where(s => s.codigo_catalogo == "TCFACT-01" || s.codigo_catalogo == "TCCLI-01").OrderBy(c => c.nombre_catalgo).Select(c => new SelectListItem
                {
                    Text = c.nombre_catalgo,
                    Value = c.id_catalogo.ToString(),
                }).ToList();
            }


            ViewBag.Prefijo = prefijo;

            if (idCliente.HasValue)
                ViewBag.Cliente = ClienteEntity.ConsultarClienteInformacionCompleta(idCliente.Value);

            return PartialView();
        }

        // POST: ContactoCliente/Create
        // To protect from overposting attacks, please enable the specific properties you want to bind to, for 
        // more details see https://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Create([Bind(Include = "id_contacto,id_cliente,nombre_contacto,apellido_contacto,cargo_contacto,mail_contacto,telefono_contacto,extension_contacto,prefijo_pais,celular_contacto,tipo_contacto, CodigoCatalogoCargoContacto, CodigoCatalogoTipoContacto")] ContactosInfo contactosClientes)
        {
            try
            {
                if (!Validaciones.ValidarMail(contactosClientes.mail_contacto)) {
                    ViewBag.Resultado = Mensajes.MensajeCorreoIncorrecto;
                    ViewBag.Estado = "false";
                    return View(contactosClientes);
                }
                    //return Json(new { Resultado = new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeCorreoIncorrecto } }, JsonRequestBehavior.AllowGet);

                if (UsuarioEntity.VerificarCorreoUsuarioExistente(contactosClientes.mail_contacto))
                {
                    ViewBag.Resultado = Mensajes.MensajeEmailExistenteAsociadoUsuario;
                    ViewBag.Estado = "false";
                    return View(contactosClientes);
                }
                    //return Json(new { Resultado = new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeEmailExistenteAsociadoUsuario } }, JsonRequestBehavior.AllowGet);

                if (ContactoClienteEntity.VerificarCorreoContactoExistente(contactosClientes.mail_contacto, contactosClientes.CodigoCatalogoTipoContacto.Value))
                {
                    ViewBag.Resultado = Mensajes.MensajeEmailExistenteAsociadoContacto;
                    ViewBag.Estado = "false";
                    return View(contactosClientes);
                }
                    //return Json(new { Resultado = new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeEmailExistenteAsociadoContacto } }, JsonRequestBehavior.AllowGet);

                if (ModelState.IsValid)
                {
                    //Listado tipoContacto
                    var tipoContacto = CatalogoEntity.ObtenerListadoCatalogosByCodigo("TCT-01");
                    ViewBag.ListadoTipoContacto = tipoContacto;

                    RespuestaTransaccion resultado = ContactoClienteEntity.CrearContactosClientes(contactosClientes);
                    //Almacenar en una variable de sesion
                    Session["Resultado"] = resultado.Respuesta;
                    Session["Estado"] = resultado.Estado.ToString();

                    if (resultado.Estado)
                    {
                        return RedirectToAction("Index");
                    }
                    else
                    {
                        ViewBag.Resultado = resultado.Respuesta;
                        ViewBag.Estado = resultado.Estado.ToString();
                        return View(contactosClientes);
                    }
                }

                return View(contactosClientes);
            }
            catch (Exception ex)
            {
                List<RespuestaTransaccion> validaciones = new List<RespuestaTransaccion>
                {
                    new RespuestaTransaccion { Estado = false, Respuesta = ex.Message }
                };
                ViewBag.Resultado = validaciones;
                return View(contactosClientes);
            }

        }

        [HttpPost]
        public ActionResult CreateAjax(ContactosInfo contactosClientes, int? idCliente, string prefijo)
        {
            try
            {
                contactosClientes.prefijo_pais = prefijo;

                if (!Validaciones.ValidarMail(contactosClientes.mail_contacto))
                    return Json(new { Resultado = new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeCorreoIncorrecto } }, JsonRequestBehavior.AllowGet);

                if (UsuarioEntity.VerificarCorreoUsuarioExistente(contactosClientes.mail_contacto))
                    return Json(new { Resultado = new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeEmailExistenteAsociadoUsuario } }, JsonRequestBehavior.AllowGet);

                if (ContactoClienteEntity.VerificarCorreoContactoExistente(contactosClientes.mail_contacto, contactosClientes.CodigoCatalogoTipoContacto.Value))
                    return Json(new { Resultado = new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeEmailExistenteAsociadoContacto } }, JsonRequestBehavior.AllowGet);

                RespuestaTransaccion resultado = ContactoClienteEntity.CrearContactosClientes(contactosClientes, idCliente);
                return Json(new { Resultado = resultado }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                return Json(new { Resultado = new RespuestaTransaccion { Estado = false, Respuesta = ex.Message } }, JsonRequestBehavior.AllowGet);
            }
        }


        // GET: ContactoCliente/Edit/5
        public ActionResult Edit(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            var contactosClientes = ContactoClienteEntity.ConsultarContactoCliente(id.Value);
            //Listado tipoContacto
            var tipoContacto = CatalogoEntity.ObtenerListadoCatalogosByCodigo("TCT-01");
            tipoContacto.Where(s => s.Value == contactosClientes.tipo_contacto.ToString()).FirstOrDefault().Selected = true;

            ViewBag.ListadoTipoContacto = tipoContacto;

            bool success = Int32.TryParse(contactosClientes.cargo_contacto, out int number);

            ViewBag.CodigoCatalogo = contactosClientes.cargo_contacto;
            ViewBag.TextoCatalogo = success ? CatalogoEntity.ConsultarCatalogo(Convert.ToInt32(contactosClientes.cargo_contacto)).nombre_catalgo : contactosClientes.cargo_contacto;

            if (contactosClientes == null)
            {
                return HttpNotFound();
            }
            return View(contactosClientes);
        }

        // POST: ContactoCliente/Edit/5
        // To protect from overposting attacks, please enable the specific properties you want to bind to, for 
        // more details see https://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Edit([Bind(Include = "id_contacto,id_cliente,nombre_contacto,apellido_contacto,cargo_contacto,mail_contacto,telefono_contacto,extension_contacto,prefijo_pais,celular_contacto,tipo_contacto, CodigoCatalogoCargoContacto, CodigoCatalogoTipoContacto")] ContactosInfo contactosClientes)
        {
            try
            {
                bool success = Int32.TryParse(contactosClientes.cargo_contacto, out int number);
                ViewBag.CodigoCatalogo = contactosClientes.cargo_contacto;
                ViewBag.TextoCatalogo = success ? CatalogoEntity.ConsultarCatalogo(Convert.ToInt32(contactosClientes.cargo_contacto)).nombre_catalgo : contactosClientes.cargo_contacto;

                if (!Validaciones.ValidarMail(contactosClientes.mail_contacto))
                {
                    ViewBag.Resultado = Mensajes.MensajeCorreoIncorrecto;
                    ViewBag.Estado = "false";
                    return View(new Contactos
                    {
                        id_contacto = contactosClientes.id_contacto,
                        nombre_contacto = contactosClientes.nombre_contacto,
                        apellido_contacto = contactosClientes.apellido_contacto,
                        cargo_contacto = contactosClientes.CodigoCatalogoCargoContacto == null ? contactosClientes.cargo_contacto : contactosClientes.CodigoCatalogoCargoContacto,
                        mail_contacto = contactosClientes.mail_contacto,
                        telefono_contacto = contactosClientes.telefono_contacto,
                        extension_contacto = contactosClientes.extension_contacto,
                        prefijo_pais = contactosClientes.prefijo_pais,
                        celular_contacto = contactosClientes.celular_contacto,
                        tipo_contacto = contactosClientes.CodigoCatalogoTipoContacto,
                        estado_contacto = true,
                    });
                }
                //return Json(new { Resultado = new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeCorreoIncorrecto } }, JsonRequestBehavior.AllowGet);

                if (UsuarioEntity.VerificarCorreoUsuarioExistente(contactosClientes.mail_contacto))
                {
                    ViewBag.Resultado = Mensajes.MensajeEmailExistenteAsociadoUsuario;
                    ViewBag.Estado = "false";
                    return View(new Contactos
                    {
                        id_contacto = contactosClientes.id_contacto,
                        nombre_contacto = contactosClientes.nombre_contacto,
                        apellido_contacto = contactosClientes.apellido_contacto,
                        cargo_contacto = contactosClientes.CodigoCatalogoCargoContacto == null ? contactosClientes.cargo_contacto : contactosClientes.CodigoCatalogoCargoContacto,
                        mail_contacto = contactosClientes.mail_contacto,
                        telefono_contacto = contactosClientes.telefono_contacto,
                        extension_contacto = contactosClientes.extension_contacto,
                        prefijo_pais = contactosClientes.prefijo_pais,
                        celular_contacto = contactosClientes.celular_contacto,
                        tipo_contacto = contactosClientes.CodigoCatalogoTipoContacto,
                        estado_contacto = true,
                    });
                }
                //return Json(new { Resultado = new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeEmailExistenteAsociadoUsuario } }, JsonRequestBehavior.AllowGet);

                if (ContactoClienteEntity.VerificarCorreoContactoExistente(contactosClientes.mail_contacto, contactosClientes.CodigoCatalogoTipoContacto.Value))
                {
                    ViewBag.Resultado = Mensajes.MensajeEmailExistenteAsociadoContacto;
                    ViewBag.Estado = "false";
                    return View(new Contactos
                    {
                        id_contacto = contactosClientes.id_contacto,
                        nombre_contacto = contactosClientes.nombre_contacto,
                        apellido_contacto = contactosClientes.apellido_contacto,
                        cargo_contacto = contactosClientes.CodigoCatalogoCargoContacto == null ? contactosClientes.cargo_contacto : contactosClientes.CodigoCatalogoCargoContacto,
                        mail_contacto = contactosClientes.mail_contacto,
                        telefono_contacto = contactosClientes.telefono_contacto,
                        extension_contacto = contactosClientes.extension_contacto,
                        prefijo_pais = contactosClientes.prefijo_pais,
                        celular_contacto = contactosClientes.celular_contacto,
                        tipo_contacto = contactosClientes.CodigoCatalogoTipoContacto,
                        estado_contacto = true,
                    });
                } 

                if (ModelState.IsValid)
                {

                    Contactos contacto = new Contactos
                    {
                        id_contacto = contactosClientes.id_contacto,
                        nombre_contacto = contactosClientes.nombre_contacto,
                        apellido_contacto = contactosClientes.apellido_contacto,
                        cargo_contacto = contactosClientes.CodigoCatalogoCargoContacto == null ? contactosClientes.cargo_contacto : contactosClientes.CodigoCatalogoCargoContacto,
                        mail_contacto = contactosClientes.mail_contacto,
                        telefono_contacto = contactosClientes.telefono_contacto,
                        extension_contacto = contactosClientes.extension_contacto,
                        prefijo_pais = contactosClientes.prefijo_pais,
                        celular_contacto = contactosClientes.celular_contacto,
                        tipo_contacto = contactosClientes.CodigoCatalogoTipoContacto,
                        estado_contacto = true,
                    };

                    RespuestaTransaccion resultado = ContactoClienteEntity.ActualizarContactosClientes(contacto);
                    //Almacenar en una variable de sesion
                    Session["Resultado"] = resultado.Respuesta;
                    Session["Estado"] = resultado.Estado.ToString();

                    if (resultado.Estado)
                    {
                        return RedirectToAction("Index");
                    }
                    else
                    {
                        ViewBag.Resultado = resultado.Respuesta;
                        ViewBag.Estado = resultado.Estado.ToString();
                        return View(contactosClientes);
                    }
                }
                return View(contactosClientes);
            }
            catch (Exception ex)
            {
                List<RespuestaTransaccion> validaciones = new List<RespuestaTransaccion>
                {
                    new RespuestaTransaccion { Estado = false, Respuesta = ex.Message }
                };
                ViewBag.Resultado = validaciones;
                return View(contactosClientes);
            }

        }

        [HttpPost]
        public ActionResult Eliminar(int id)
        {
            if (id == 0)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            RespuestaTransaccion resultado = ContactoClienteEntity.EliminarContactosClientes(id);// await db.Cabecera.FindAsync(id);

            //return RedirectToAction("Index");
            return Json(new { Resultado = resultado }, JsonRequestBehavior.AllowGet);

        }

        //Autocompleto
        public ActionResult GetItemsCargos(string busqueda)// v is the entered text
        {
            busqueda = (busqueda ?? "").ToLower().Trim();
            var items = CatalogoEntity.ObtenerListadoCatalogosByCodigo("CRG-01").Where(o => o.Text.ToLower().Contains(busqueda));
            return Json(items.Take(10).Select(o => new KeyContent(o.Value, o.Text)));
        }


        public JsonResult _GetItemsCargos(string busqueda)
        {
            List<AutoCompleteUI> items = new List<AutoCompleteUI>();
            busqueda = (busqueda ?? "").ToLower().Trim();

            items = CatalogoEntity.ObtenerListadoCatalogosByCodigo("CRG-01").Where(o => o.Text.ToLower().Contains(busqueda)).Select(o => new AutoCompleteUI(Convert.ToInt64(o.Value), o.Text, "")).ToList();
            return Json(new { results = items } , JsonRequestBehavior.AllowGet);
        }


        public ActionResult DescargarReporteFormatoExcel()
        {
            //Seleccionar las columnas a exportar
            var collection = ContactoClienteEntity.ListadoContactosClientes();
            var package = new ExcelPackage();

            package = Reportes.ExportarExcel(collection.Cast<object>().ToList(), "Contacto Cliente");
            return File(package.GetAsByteArray(), XlsxContentType, "ListadoContactosCliente.xlsx");
        }

        public ActionResult DescargarReporteFormatoPDF()
        {
            // Seleccionar las columnas a exportar
            var results = ContactoClienteEntity.ListadoContactosClientes();

            var list = Reportes.SerializeToJSON(results);
            return Content(list, "application/json");
        }

        public ActionResult DescargarReporteFormatoCSV()
        {
            var comlumHeadrs = new string[]
            {
                "ID",
                "USUARIO",
                "CLIENTE",
                "TIPOCONTACTO",
                "NOMBRESCOMPLETOS",
                "CARGO",
                "MAIL",
                "PREFIJO",
                "TELEFONO",
                "EXTENSION",
                "CELULAR",
                "ESTADO"
            };

            var listado = (from item in ContactoClienteEntity.ListadoContactosClientes()
                           select new object[]
                           {
                    item.Id,
                    item.Usuario,
                    item.Cliente,
                    item.TipoContacto,
                    item.NombresCompletos,
                    item.Cargo,
                    item.Mail,
                    item.Prefijo,
                    item.Telefono,
                    item.Extension,
                    item.Celular,
                    item.Estado

                           }).ToList();

            // Build the file content
            var employeecsv = new StringBuilder();
            listado.ForEach(line =>
            {
                employeecsv.AppendLine(string.Join(",", line));
            });

            byte[] buffer = Encoding.ASCII.GetBytes($"{string.Join(",", comlumHeadrs)}\r\n{employeecsv.ToString()}");
            return File(buffer, "text/csv", $"ContactoCliente.csv");
        }


        #region Funcionalidad de carga Masiva
        public ActionResult CargarData()
        {
            List<CargaMasiva> errores = new List<CargaMasiva>();
            RespuestaTransaccion resultado = new RespuestaTransaccion { Estado = false };
            List<Contactos> listado = new List<Contactos>();
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

                        if (listado.Count > 0)
                            resultado = ContactoClienteEntity.CrearContactosCargaMasiva(listado); // Crea los Contactos
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

        private List<Contactos> ToEntidadHojaExcelList(string pathDelFicheroExcel)
        {
            List<Contactos> listadoCargaMasiva = new List<Contactos>();
            try
            {

                var user = ViewData["usuario"] = System.Web.HttpContext.Current.Session["usuario"];
                var usuarioID = Convert.ToInt16(user);

                FileInfo existingFile = new FileInfo(pathDelFicheroExcel);

                using (ExcelPackage package = new ExcelPackage(existingFile))
                {
                    // get the first worksheet in the workbook
                    ExcelWorksheet worksheet = package.Workbook.Worksheets.First();

                    if (worksheet != null)
                    {
                        int colCount = worksheet.Dimension.End.Column;  //get Column Count
                        int rowCount = worksheet.Dimension.End.Row;     //get row count
                        for (int row = 2; row <= rowCount; row++)
                        {
                            listadoCargaMasiva.Add(new Contactos
                            {
                                nombre_contacto = !string.IsNullOrEmpty((worksheet.GetValue(row, 1) ?? "").ToString().Trim()) ? (worksheet.GetValue(row, 1) ?? "").ToString().Trim() : string.Empty,
                                apellido_contacto = !string.IsNullOrEmpty((worksheet.GetValue(row, 2) ?? "").ToString().Trim()) ? (worksheet.GetValue(row, 2) ?? "").ToString().Trim() : string.Empty,
                                cargo_contacto = !string.IsNullOrEmpty((worksheet.GetValue(row, 3) ?? "").ToString().Trim()) ? (worksheet.GetValue(row, 3) ?? "").ToString().Trim() : string.Empty,
                                mail_contacto = !string.IsNullOrEmpty((worksheet.GetValue(row, 4) ?? "").ToString().Trim()) ? (worksheet.GetValue(row, 4) ?? "").ToString().Trim() : string.Empty,
                                telefono_contacto = !string.IsNullOrEmpty((worksheet.GetValue(row, 5) ?? "").ToString().Trim()) ? (worksheet.GetValue(row, 5) ?? "").ToString().Trim() : string.Empty,
                                extension_contacto = !string.IsNullOrEmpty((worksheet.GetValue(row, 6) ?? "").ToString().Trim()) ? (worksheet.GetValue(row, 6) ?? "").ToString().Trim() : string.Empty,
                                prefijo_pais = !string.IsNullOrEmpty((worksheet.GetValue(row, 7) ?? "").ToString().Trim()) ? (worksheet.GetValue(row, 7) ?? "").ToString().Trim() : string.Empty,
                                celular_contacto = !string.IsNullOrEmpty((worksheet.GetValue(row, 8) ?? "").ToString().Trim()) ? (worksheet.GetValue(row, 8) ?? "").ToString().Trim() : string.Empty,
                                tipo_contacto = !string.IsNullOrEmpty((worksheet.GetValue(row, 9) ?? "").ToString().Trim()) ? int.Parse((worksheet.GetValue(row, 9) ?? "").ToString().Trim()) : 0,
                                
                            });
                        }
                    }
                }

                return listadoCargaMasiva;
            }
            catch (Exception ex)
            {
                return new List<Contactos>();
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

                                int tipoContacto = int.Parse((worksheet.Cells[row, colCount].Value ?? "").ToString().Trim()); // El tipo de contacto se encuentra en la ultima columna

                                error = columna == "MAIL" ?  ValidarCamposCargaMasiva(columna, valorColumna, tipoContacto) : ValidarCamposCargaMasiva(columna, valorColumna); // Solo para validar el MAIL se pasa el tipo de Contacto

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
        private static string ValidarCamposCargaMasiva(string columna, string valor , int? tipoContacto = null)
        {
            var excepcion = "El campo {0} con el valor {1} , contiene errores. Error: {2}";
            try
            {
                valor = !string.IsNullOrEmpty(valor) ? valor.Trim() : "";

                bool esNumero;
                int longitudCaracteres = 0;
                var error = "El campo {0} con el valor {1} ,tiene una extensión de caracteres mayor a la permitida. Solo permite {2} caracteres";
                switch (columna)
                {
                    case "TIPO":
                        esNumero = int.TryParse(valor, out int valorTipo);
                        if (!esNumero)
                            error = "Tipo de dato incorrecto para el campo " + columna + " , con el valor " + valor + " .Solo admite valores numéricos de tipo entero.";
                        else {
                            if (!CatalogoEntity.VerificarExistenciaCatalogo(int.Parse(valor)))
                                error = "El valor no es válido para el campo " + columna + " , con el valor " + valor + " ." + Mensajes.MensajeCatalogoNoDisponible;
                            else
                                error = null;
                        }

                        break;
                    case "NOMBRES":
                        longitudCaracteres = 150;
                        if (valor.Length > longitudCaracteres)
                        {
                            error = string.Format(error, columna, valor, longitudCaracteres);
                        }
                        else
                        {
                            error = null;
                        }
                        break;
                    case "APELLIDOS":
                        longitudCaracteres = 200;
                        if (valor.Length > longitudCaracteres)
                        {
                            error = string.Format(error, columna, valor, longitudCaracteres);
                        }
                        else
                        {
                            error = null;
                        }
                        break;
                    case "CARGO":
                        longitudCaracteres = 200;
                        if (valor.Length > longitudCaracteres)
                        {
                            error = string.Format(error, columna, valor, longitudCaracteres);
                        }
                        else
                        {
                            error = null;
                        }
                        break;
                    case "EXTENSION":
                        longitudCaracteres = 5;
                        if (valor.Length > longitudCaracteres)
                            error = string.Format(error, columna, valor, longitudCaracteres);
                        else
                            error = null;
                        break;
                    case "PREFIJO PAIS":

                        longitudCaracteres = 5;
                        if (valor.Length > longitudCaracteres)
                            error = string.Format(error, columna, valor, longitudCaracteres);
                        else
                            error = null;

                        var busquedaPrefijo = CatalogoEntity.ValidarExistenciaPrefijo(valor);

                        if(!busquedaPrefijo)
                            error = "El valor no es válido para el campo " + columna + " , con el valor " + valor + " ." + Mensajes.MensajeValidacionExistenciaPrefijo;
                        else
                            error = null;

                        break;
                    case "MAIL":

                        longitudCaracteres = 300;
                        if (valor.Length > longitudCaracteres)
                            error = string.Format(error, columna, valor, longitudCaracteres);
                        else
                            error = null;

                        if (!Validaciones.ValidarMail(valor))
                            error = "El valor no es válido para el campo " + columna + " , con el valor " + valor + " ." + Mensajes.MensajeCorreoIncorrecto;
                        else
                            error = null;

                        if (UsuarioEntity.VerificarCorreoUsuarioExistente(valor))
                            error = "El valor no es válido para el campo " + columna + " , con el valor " + valor + " ." + Mensajes.MensajeEmailExistenteAsociadoUsuario;
                        else
                            error = null;

                        if (tipoContacto.HasValue)
                        {
                            if (ContactoClienteEntity.VerificarCorreoContactoExistente(valor, tipoContacto.Value))
                                error = "El valor no es válido para el campo " + columna + " , con el valor " + valor + " ." + Mensajes.MensajeEmailExistenteAsociadoContacto;
                            else
                                error = null;
                        }
                        else {
                            error = "El valor no es válido para el campo " + columna + " , con el valor " + valor + " ." + Mensajes.MensajeValidacionCorreoExistenteContactos;
                        }

                        break;
                    case "CELULAR":
                    case "TELEFONO":
                        longitudCaracteres = 15;
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
                        error = "No se pudo encontrar la columna " + columna +".";
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

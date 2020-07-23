using System;
using System.Collections.Generic;
using System.Data; 
using System.Data.Entity.Migrations;
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
using Seguridad.Helper;

namespace TemplateInicial.Controllers
{
    [Autenticado]
    public class MenuController : BaseAppController
    {
        private GestionPPMEntities db = new GestionPPMEntities();
        private const string XlsxContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
        // GET: Menu
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
            ViewBag.NombreListado = Etiquetas.TituloGridMenu;
            //Controlar permisos
            var user = ViewData["usuario"] = System.Web.HttpContext.Current.Session["usuario"];
            var usuario = int.Parse(user.ToString());
            string nombreControlador = ControllerContext.RouteData.Values["controller"].ToString();
            ViewBag.NombreControlador = nombreControlador;

            ViewBag.AccionesUsuario = ManejoPermisosEntity.ListadoAccionesCatalogoUsuario(usuario, nombreControlador);

            //Obtener Acciones del controlador
            ViewBag.AccionesControlador = GetMetodosControlador(nombreControlador);

            //Búsqueda
            var listado = MenuEntity.ListarMenu();

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

        // GET: Menu/Details/5
        public ActionResult Details(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            Menu menu = db.Menu.Find(id);
            if (menu == null)
            {
                return HttpNotFound();
            }
            return View(menu);
        }

        // GET: Menu/Create
        public ActionResult Create()
        {
            var listado = MenuEntity.ListarMenu().Select(s => new Menu { id_menu = s.Id, nombre_menu = s.Opcion_Menu + " ( " + s.Ruta_Acceso + " )" }).AsEnumerable();
            ViewBag.listadoMenu = new SelectList(listado, "id_menu", "nombre_menu");

            return View();
        }

        // POST: Menu/Create
        // To protect from overposting attacks, please enable the specific properties you want to bind to, for 
        // more details see https://go.microsoft.com/fwlink/?LinkId=317598.
        //[HttpPost]
        //[ValidateAntiForgeryToken]
        //public ActionResult Create([Bind(Include = "id_menu,nombre_menu,nombre_pagina_menu,estado_menu,id_menu_padre")] Menu menu)
        //{
        //    if (ModelState.IsValid)
        //    {
        //        var listado = MenuEntity.ListarMenu().Select(s => new Menu { id_menu = s.id_menu, nombre_menu = s.nombre_menu + " ( " + s.nombre_pagina_menu + " )" }).AsEnumerable();
        //        ViewBag.listadoMenu = new SelectList(listado, "id_menu", "nombre_menu");

        //        string nombreMenu = (menu.nombre_menu ?? string.Empty).ToLower().Trim();

        //        var validacionNombreRolUnico = MenuEntity.ListarMenu().Where(s => (s.nombre_menu ?? string.Empty).ToLower().Trim() == nombreMenu).ToList();

        //        if (validacionNombreRolUnico.Count > 0)
        //        {

        //            List<RespuestaTransaccion> validacionesRespuesta = new List<RespuestaTransaccion>{
        //            new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeResgistroExistente }
        //            };

        //            //Listado de validaciones especificas de la entidad
        //            ViewBag.Resultado = validacionesRespuesta;

        //            return View(menu);
        //        }

        //        RespuestaTransaccion resultado = MenuEntity.CrearMenu(menu);
        //        //return Json(new { Resultado = resultado }, JsonRequestBehavior.AllowGet);

        //        //Almacenar en una variable de sesion
        //        Session["Resultado"] = resultado.Respuesta;
        //        Session["Estado"] = resultado.Estado.ToString();

        //        if (resultado.Estado.ToString() == "True")
        //        {
        //            return RedirectToAction("Index");
        //        }
        //        else
        //        {
        //            ViewBag.Resultado = resultado.Respuesta;
        //            ViewBag.Estado = resultado.Estado.ToString();
        //            return View(menu);
        //        }

        //        //ViewBag.Resultado = resultado;
        //        //return View("Index");
        //    }

        //    return View(menu);
        //}
        [HttpPost]
        public ActionResult Create(Menu menu)
        {
            try
            {
                var listado = MenuEntity.ListarMenu().Select(s => new Menu { id_menu = s.Id, nombre_menu = s.Opcion_Menu + " ( " + s.Ruta_Acceso + " )" }).AsEnumerable();
                ViewBag.listadoMenu = new SelectList(listado, "id_menu", "nombre_menu");

                string nombreMenu = (menu.nombre_menu ?? string.Empty).ToLower().Trim();

                var validacionNombreRolUnico = MenuEntity.ListarMenu().Where(s => (s.Opcion_Menu ?? string.Empty).ToLower().Trim() == nombreMenu).ToList();

                if (validacionNombreRolUnico.Count > 0)
                    return Json(new { Resultado = new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeResgistroExistente } }, JsonRequestBehavior.AllowGet);

                RespuestaTransaccion resultado = MenuEntity.CrearMenu(menu);

                return Json(new { Resultado = resultado }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                return Json(new { Resultado = new RespuestaTransaccion { Estado = false, Respuesta = ex.Message } }, JsonRequestBehavior.AllowGet);
            }
        }
        // GET: Menu/Edit/5
        public ActionResult Edit(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            var menu = MenuEntity.ConsultarMenu(id.Value);

            var listado = MenuEntity.ListarMenu().Select(s => new Menu { id_menu = s.Id, nombre_menu = s.Opcion_Menu + " ( " + s.Ruta_Acceso + " )" }).AsEnumerable();
            ViewBag.listadoMenu = new SelectList(listado, "id_menu", "nombre_menu", id.Value);

            if (menu == null)
            {
                return HttpNotFound();
            }
            return View(menu);
        }

        // POST: Menu/Edit/5
        // To protect from overposting attacks, please enable the specific properties you want to bind to, for 
        // more details see https://go.microsoft.com/fwlink/?LinkId=317598.
        //[HttpPost]
        //[ValidateAntiForgeryToken]
        //public ActionResult Edit([Bind(Include = "id_menu,nombre_menu,nombre_pagina_menu,estado_menu,id_menu_padre")] Menu menu)
        //{
        //    if (ModelState.IsValid)
        //    {
        //        var listado = MenuEntity.ListarMenu().Select(s => new Menu { id_menu = s.id_menu, nombre_menu = s.nombre_menu + " ( " + s.nombre_pagina_menu + " )" }).AsEnumerable();
        //        ViewBag.listadoMenu = new SelectList(listado, "id_menu", "nombre_menu");

        //        string nombreMenu = (menu.nombre_menu ?? string.Empty).ToLower().Trim();

        //        var validacionNombreRolUnico = MenuEntity.ListarMenu().Where(s => (s.nombre_menu ?? string.Empty).ToLower().Trim() == nombreMenu && s.id_menu != menu.id_menu).ToList();

        //        if (validacionNombreRolUnico.Count > 0)
        //        {

        //            List<RespuestaTransaccion> validacionesRespuesta = new List<RespuestaTransaccion>{
        //            new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeResgistroExistente }
        //            };

        //            //Listado de validaciones especificas de la entidad
        //            ViewBag.Resultado = validacionesRespuesta;

        //            return View(menu);
        //        }

        //        RespuestaTransaccion resultado = MenuEntity.ActualizarMenu(menu);

        //        //Almacenar en una variable de sesion
        //        Session["Resultado"] = resultado.Respuesta;
        //        Session["Estado"] = resultado.Estado.ToString();

        //        if (resultado.Estado.ToString() == "True")
        //        {
        //            return RedirectToAction("Index");
        //        }
        //        else
        //        {
        //            ViewBag.Resultado = resultado.Respuesta;
        //            ViewBag.Estado = resultado.Estado.ToString();
        //            return View(menu);
        //        }
        //        //return Json(new { Resultado = resultado }, JsonRequestBehavior.AllowGet);
        //    }

        //    return View(menu);
        //}
        [HttpPost]
        public ActionResult Edit(Menu menu)
        {
            try
            {
                var listado = MenuEntity.ListarMenu().Select(s => new Menu { id_menu = s.Id, nombre_menu = s.Opcion_Menu + " ( " + s.Ruta_Acceso + " )" }).AsEnumerable();
                ViewBag.listadoMenu = new SelectList(listado, "id_menu", "nombre_menu");

                string nombreMenu = (menu.nombre_menu ?? string.Empty).ToLower().Trim();

                var validacionNombreRolUnico = MenuEntity.ListarMenu().Where(s => (s.Opcion_Menu ?? string.Empty).ToLower().Trim() == nombreMenu && s.Id != menu.id_menu).ToList();

                if (validacionNombreRolUnico.Count > 0)
                    return Json(new { Resultado = new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeResgistroExistente } }, JsonRequestBehavior.AllowGet);

                RespuestaTransaccion resultado = MenuEntity.ActualizarMenu(menu);

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
            RespuestaTransaccion resultado = MenuEntity.EliminarMenu(id);// await db.Cabecera.FindAsync(id);

            //return RedirectToAction("Index");
            return Json(new { Resultado = resultado }, JsonRequestBehavior.AllowGet);

        }

        // GET: Menu/Delete/5
        public ActionResult Delete(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            Menu menu = db.Menu.Find(id);
            if (menu == null)
            {
                return HttpNotFound();
            }
            return View(menu);
        }

        // POST: Menu/Delete/5
        [HttpPost, ActionName("Delete")]
        [ValidateAntiForgeryToken]
        public ActionResult DeleteConfirmed(int id)
        {
            Menu menu = db.Menu.Find(id);
            db.Menu.Remove(menu);
            db.SaveChanges();
            return RedirectToAction("Index");
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                db.Dispose();
            }
            base.Dispose(disposing);
        }

        public ActionResult DescargarReporteFormatoExcel()
        {
            //Seleccionar las columnas a exportar
            var collection = MenuEntity.ListarMenu();
            var package = new ExcelPackage();

            package = Reportes.ExportarExcel(collection.Cast<object>().ToList(),"Menu");
            return File(package.GetAsByteArray(), XlsxContentType, "ListadoOpcionesMenu.xlsx");
        }

        public ActionResult DescargarReporteFormatoPDF()
        {
            // Seleccionar las columnas a exportar
            var results = MenuEntity.ListarMenu();

            var list = Reportes.SerializeToJSON(results);
            return Content(list, "application/json");
        }

        public ActionResult DescargarReporteFormatoCSV()
        {
            var comlumHeadrs = new string[]
            {
                "menu Id",
                "nombre_pagina_menu",
                "estado_menu",
            };

            var listado = (from item in MenuEntity.ListarMenu()
                           select new object[]
                           {
                                            item.Opcion_Menu,
                                            item.Ruta_Acceso,
                                            item.Estado 
                           }).ToList();

            // Build the file content
            var employeecsv = new StringBuilder();
            listado.ForEach(line =>
            {
                employeecsv.AppendLine(string.Join(",", line));
            });

            byte[] buffer = Encoding.ASCII.GetBytes($"{string.Join(",", comlumHeadrs)}\r\n{employeecsv.ToString()}");
            return File(buffer, "text/csv", $"OpcionesMenu.csv");
        }

        public ActionResult OrdenMenu()
        {  
            var item = db.ListadoOrdenMenu().ToList();
            var item2 =  item.OrderBy(x => x.orden_menu);

            return View(item2);
        }

        public ActionResult ActualizarOrdenPadre(string itemIds)
        {
            int count = 1;
            List<int> itemIdList = new List<int>();
            itemIdList = itemIds.Split(",".ToCharArray(), StringSplitOptions.RemoveEmptyEntries).Select(int.Parse).ToList();
            foreach (var itemId in itemIdList)
            {
                try
                {
                    Menu item = db.Menu.Where(x => x.id_menu == itemId).FirstOrDefault();
                    item.orden_menu = count;
                    db.Menu.AddOrUpdate(item);
                    db.SaveChanges();
                }
                catch (Exception)
                {
                    continue;
                }
                count++;

            }
            return Json(true, JsonRequestBehavior.AllowGet);
        }

        public ActionResult ActualizarOrdenHijo(string itemIds)
        {
            int count = 1;
            List<int> itemIdList = new List<int>();
            itemIdList = itemIds.Split(",".ToCharArray(), StringSplitOptions.RemoveEmptyEntries).Select(int.Parse).ToList();
            foreach (var itemId in itemIdList)
            {
                try
                {
                    Menu item = db.Menu.Where(x => x.id_menu == itemId).FirstOrDefault();
                    item.orden_menu = count;
                    db.Menu.AddOrUpdate(item);
                    db.SaveChanges();
                }
                catch (Exception)
                {
                    continue;
                }
                count++;

            }
            return Json(true, JsonRequestBehavior.AllowGet);
        }

        //ADJUNTAR ARCHIVOS
        public ActionResult _AdjuntarArchivos(int? id)
        {

            ViewBag.TituloModal = "Adjuntar Manual de Usuario";
            Menu idMenu = MenuEntity.ConsultarMenu(id.Value);
            return PartialView(idMenu);
        }
         
        public ActionResult AdjuntarArchivoManualUsuario(int? idMenu)
        {
            string FileName = "";
            try
            {

                Menu menu = MenuEntity.ConsultarMenu(idMenu.Value);

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
                    string extension = Path.GetExtension(path);
                    extension = extension.ToLower();

                    string[] nombreController = menu.nombre_pagina_menu.Split(new char[] { '/' });

                    var nombrearchivo = nombreController[0] + ".pdf";

                    bool directorio = Directory.Exists(Server.MapPath("~/AdjuntosManual/"));

                    // En caso de que no exista el directorio, crearlo.
                    if (!directorio)
                        Directory.CreateDirectory(Server.MapPath("~/AdjuntosManual/"));

                    // Get the complete folder path and store the file inside it.    
                    path = Path.Combine(Server.MapPath("~/AdjuntosManual/"), nombrearchivo);

                    if (extension != ".pdf")
                    {
                        return Json(new { Resultado = new RespuestaTransaccion { Estado = false, Respuesta = "La extensión de formato de carga está incorrecto" } }, JsonRequestBehavior.AllowGet);
                    }
                    else

                        file.SaveAs(path);

                }
                return Json(new { Resultado = new RespuestaTransaccion { Estado = true, Respuesta = Mensajes.MensajeAdjuntoExitoso }, Archivo = FileName }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                return Json(new { Resultado = new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeAdjuntoFallido + ex.Message } }, JsonRequestBehavior.AllowGet);
            }
        }
    }
}
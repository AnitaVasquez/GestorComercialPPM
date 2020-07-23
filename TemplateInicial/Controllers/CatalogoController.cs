using GestionPPM.Entidades.Metodos;
using GestionPPM.Entidades.Modelo;
using GestionPPM.Repositorios;
using NLog;
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
    public class CatalogoController : BaseAppController
    {

        private const string XlsxContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
        private static readonly Logger Log = LogManager.GetCurrentClassLogger();

        // GET: Catalogo
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
            ViewBag.NombreListado = Etiquetas.TituloGridCatalogo;

            //Controlar permisos
            var user = ViewData["usuario"] = System.Web.HttpContext.Current.Session["usuario"];
            var usuario = int.Parse(user.ToString());
            string nombreControlador = ControllerContext.RouteData.Values["controller"].ToString();
            ViewBag.NombreControlador = nombreControlador;


            ViewBag.AccionesUsuario = ManejoPermisosEntity.ListadoAccionesCatalogoUsuario(usuario, nombreControlador);

            //Obtener Acciones del controlador
            ViewBag.AccionesControlador = GetMetodosControlador(nombreControlador);

            //Búsqueda

            var listado = CatalogoEntity.ListarCatalogos();

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

        // GET: Catalogo/Create
        public ActionResult Create()
        {
            try
            {
                return View();
            }
            catch (Exception ex)
            {
                Log.Error(ex, "Unhandled application exception");
                return View();
            }

        }

        // POST: Catalogo/Create
        [HttpPost]
        public ActionResult Create(Catalogo catalogo)
        {
            try
            {
                if (ModelState.IsValid)
                {
                    RespuestaTransaccion resultado = CatalogoEntity.CrearCatalogo(catalogo);

                    //Almacenar en una variable de sesion
                    Session["Resultado"] = resultado.Respuesta;
                    Session["Estado"] = resultado.Estado.ToString();

                    if (resultado.Estado.ToString() == "True")
                    {
                        return RedirectToAction("Index");
                    }
                    else
                    {
                        ViewBag.Resultado = resultado.Respuesta;
                        ViewBag.Estado = resultado.Estado.ToString();
                        Session["Resultado"] = "";
                        Session["Estado"] = "";
                        return View(catalogo);
                    }
                }
                return View(catalogo);
            }
            catch
            {
                return View(catalogo);
            }
        }

        // GET: Catalogo/Edit/5
        public ActionResult Edit(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            else
            {
                var catalogo = CatalogoEntity.ConsultarCatalogo(id.Value);

                if (catalogo == null)
                {
                    return HttpNotFound();
                }
                else
                {
                    return View(catalogo);
                }
            }
        }

        // POST: Catalogo/Edit/5
        [HttpPost]
        public ActionResult Edit(Catalogo catalogo)
        {
            try
            {
                if (ModelState.IsValid)
                {
                    RespuestaTransaccion resultado = CatalogoEntity.ActualizarCatalogo(catalogo);

                    //Almacenar en una variable de sesion
                    Session["Resultado"] = resultado.Respuesta;
                    Session["Estado"] = resultado.Estado.ToString();

                    if (resultado.Estado.ToString() == "True")
                    {
                        return RedirectToAction("Index");
                    }
                    else
                    {
                        ViewBag.Resultado = resultado.Respuesta;
                        ViewBag.Estado = resultado.Estado.ToString();
                        Session["Resultado"] = "";
                        Session["Estado"] = "";
                        return View(catalogo);
                    }
                }

                return View(catalogo);
            }
            catch
            {
                return View();
            }
        }

        [HttpPost]
        public ActionResult Eliminar(int id)
        {
            if (id == 0)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }

            RespuestaTransaccion resultado = CatalogoEntity.EliminarCatalogo(id);

            //Almacenar en una variable de sesion
            Session["Resultado"] = resultado.Respuesta;
            Session["Estado"] = resultado.Estado.ToString();

            //return RedirectToAction("Index");
            return Json(new { Resultado = resultado }, JsonRequestBehavior.AllowGet);
        }

        // GET: Subcatalogo/
        public ActionResult IndexSubcatalogo(int id)
        {
            var respuesta = System.Web.HttpContext.Current.Session["Resultado"] as string;
            var estado = System.Web.HttpContext.Current.Session["Estado"] as string;

            ViewBag.Resultado = respuesta;
            ViewBag.Estado = estado;

            Session["Resultado"] = "";
            Session["Estado"] = "";

            System.Web.HttpContext.Current.Session["id_catalogo"] = id.ToString();

            //Obtener Listado de Hijos del Catalogo 
            var numeroHijos = CatalogoEntity.ObtenerNumeroHijosCatalgo(id);
            ViewBag.numeroHijos = numeroHijos;

            //Obtener listado de hijos del id seleccionado
            var ListadoHijosPadre = CatalogoEntity.ListadoHijosoCatalogoPorIdPadre(id);
            ViewBag.ListadoHijosPadre = ListadoHijosPadre;

            //Obtener listado de hijos 
            var ListadoCatalogoPadre = CatalogoEntity.ListadoCatalogosPorIdSinOrdenar(id);
            ViewBag.ListadoCatalogoPadre = ListadoCatalogoPadre;

            ViewBag.EtapaGeneral = new Catalogo();
            ViewBag.EstatusDetallado = new Catalogo();
            ViewBag.EstatusGeneral = new Catalogo();

            ViewBag.IdCatalogo = id;
            var codigoCatalogo = CatalogoEntity.ConsultarCodigocatalogo(id);
            ViewBag.CodigoCatalago = codigoCatalogo;

            return View();
        }

        [HttpGet]
        public async Task<PartialViewResult> IndexGridSubcatalogo(String search, int? tipo, int? subcatalogo, string filtro)
        {

            //Controlar permisos
            var user = ViewData["usuario"] = System.Web.HttpContext.Current.Session["usuario"];
            var usuario = int.Parse(user.ToString());
            string nombreControlador = ControllerContext.RouteData.Values["controller"].ToString();
            ViewBag.NombreControlador = nombreControlador;


            ViewBag.AccionesUsuario = ManejoPermisosEntity.ListadoAccionesCatalogoUsuario(usuario, nombreControlador);

            //Obtener Acciones del controlador
            ViewBag.AccionesControlador = GetMetodosControlador(nombreControlador);
            
            //Búsqueda

            //Cuando no selecciona tipo
            if (tipo == null && subcatalogo == null)
            {
                //Obtener le id del catalogo
                var id_catalogo = ViewData["id_catalogo"] = System.Web.HttpContext.Current.Session["id_catalogo"] as String;

                var nombre = CatalogoEntity.ConsultarNombreCatalogo(Convert.ToInt32(id_catalogo));

                ViewBag.NombreListado = "Catálogo -" + nombre;
                var listado = CatalogoEntity.ListarCatalogosPorId(Convert.ToInt32(id_catalogo), filtro);

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
                return PartialView("_IndexGridSubcatalogo", await Task.Run(() => listado));
            }
            else
            {
                //Cuando no selecciona tipo
                if (tipo != null && subcatalogo != null)
                {
                    //Obtener le id del catalogo
                    var id_catalogo = ViewData["id_catalogo"] = System.Web.HttpContext.Current.Session["id_catalogo"] as String;

                    var nombre = CatalogoEntity.ConsultarNombreCatalogo(Convert.ToInt32(subcatalogo));

                    ViewBag.NombreListado = "Catálogo -" + nombre;
                    var listado = CatalogoEntity.ListarCatalogosPorId(subcatalogo.Value, filtro);

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
                    return PartialView("_IndexGridSubcatalogo", await Task.Run(() => listado));
                }
                else
                {
                    //Obtener le id del catalogo
                    var id_catalogo = ViewData["id_catalogo"] = System.Web.HttpContext.Current.Session["id_catalogo"] as String;

                    var nombre = CatalogoEntity.ConsultarNombreCatalogo(Convert.ToInt32(tipo));

                    ViewBag.NombreListado = "Catálogo -" + nombre;
                    var listado = CatalogoEntity.ListarCatalogosPorId(Convert.ToInt32(tipo), filtro);

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
                    return PartialView("_IndexGridSubcatalogo", await Task.Run(() => listado));
                }
            }
        }

        //Accion para crear subcatalogo
        [HttpPost]
        public ActionResult IndexSubcatalogo(Catalogo catalogo, string general, string detallado, string statusGeneral)
        {
            //Variable para las respuestas
            RespuestaTransaccion resultado = new RespuestaTransaccion();
            int idCatalogo = 0;

            //Validar campos llenos
            if ((catalogo.nombre_catalgo == null) && (general == null || detallado == null || statusGeneral == null))
            {
                resultado = new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeDatosObligatorios };
            }
            else
            {
                //Recuperado de la vista
                idCatalogo = catalogo.id_catalogo;

                //validar si el tipo es diferente catalogo
                if (catalogo.id_catalogo == 0)
                {
                    idCatalogo = CatalogoEntity.ObtenerIdPadre(catalogo.id_catalogo_padre.Value);
                }

                //Validar que no sea de tipo etapa cliente

                if (catalogo.id_catalogo == 0 && idCatalogo == 167)
                {
                    //Se realiza mas de un registro en linea
                    resultado = CatalogoEntity.CrearSubcatalogoEtapaCliente(catalogo, general, detallado, statusGeneral);
                }
                else
                {
                    //Crear la subcategoria
                    resultado = CatalogoEntity.CrearSubcatalogo(catalogo);
                }
            }

            //Almacenar en una variable de sesion
            Session["Resultado"] = resultado.Respuesta;
            Session["Estado"] = resultado.Estado.ToString();

            //Obtener Listado de Hijos del Catalogo 
            var numeroHijos = CatalogoEntity.ObtenerNumeroHijosCatalgo(Convert.ToInt32(idCatalogo));
            ViewBag.numeroHijos = numeroHijos;

            //Obtener listado de hijos del id seleccionado
            var ListadoHijosPadre = CatalogoEntity.ListadoHijosoCatalogoPorIdPadre(Convert.ToInt32(idCatalogo));
            ViewBag.ListadoHijosPadre = ListadoHijosPadre;

            //Obtener listado de hijos 
            var ListadoCatalogoPadre = CatalogoEntity.ListadoCatalogosPorIdSinOrdenar(Convert.ToInt32(idCatalogo));
            ViewBag.ListadoCatalogoPadre = ListadoCatalogoPadre;

            ViewBag.EtapaGeneral = new Catalogo();
            ViewBag.EstatusDetallado = new Catalogo();
            ViewBag.EstatusGeneral = new Catalogo();

            ViewBag.IdCatalogo = idCatalogo;

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

            if (resultado.Estado.ToString() == "True")
            {
                ViewBag.Resultado = resultado.Respuesta;
                ViewBag.Estado = resultado.Estado.ToString();

                return RedirectToAction("IndexSubcatalogo/" + idCatalogo, "Catalogo");
            }
            else
            {
                ViewBag.Resultado = resultado.Respuesta;
                ViewBag.Estado = resultado.Estado.ToString();

                Session["Resultado"] = "";
                Session["Estado"] = "";

                return View(catalogo);
            }
        }

        [HttpPost]
        public ActionResult CreateSubCatalogo(Catalogo catalogo)
        {
            RespuestaTransaccion resultado = new RespuestaTransaccion();
            try
            {
                resultado = CatalogoEntity.CrearSubcatalogo(catalogo);

                return Json(new { Resultado = resultado }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                return Json(new { Resultado = new RespuestaTransaccion { Estado = false, Respuesta = ex.Message } }, JsonRequestBehavior.AllowGet);
            }
        }

        //[HttpPost]
        //public ActionResult IndexSubcatalogo(Catalogo catalogo, string general, string detallado, string statusGeneral)
        //{
        //    try
        //    {
        //        RespuestaTransaccion resultado = new RespuestaTransaccion();

        //        int idCatalogo = 0;
        //        //Validar campos llenos
        //        if ((catalogo.nombre_catalgo == null) && (general == null || detallado == null || statusGeneral == null))
        //        {
        //            return Json(new { Resultado = new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeDatosObligatorios } }, JsonRequestBehavior.AllowGet);
        //        }
        //        else
        //        {
        //            //Recuperado de la vista
        //            idCatalogo = catalogo.id_catalogo;

        //            //validar si el tipo es diferente catalogo
        //            if (catalogo.id_catalogo == 0)
        //            {
        //                idCatalogo = CatalogoEntity.ObtenerIdPadre(catalogo.id_catalogo_padre.Value);
        //            }

        //            //Validar que no sea de tipo etapa cliente

        //            if (catalogo.id_catalogo == 0 && idCatalogo == 167)
        //            {
        //                //Se realiza mas de un registro en linea
        //                resultado = CatalogoEntity.CrearSubcatalogoEtapaCliente(catalogo, general, detallado, statusGeneral);
        //            }
        //            else
        //            {
        //                //Crear la subcategoria
        //                resultado = CatalogoEntity.CrearSubcatalogo(catalogo);
        //            }
        //        }

        //        //Obtener Listado de Hijos del Catalogo 
        //        var numeroHijos = CatalogoEntity.ObtenerNumeroHijosCatalgo(Convert.ToInt32(idCatalogo));
        //        ViewBag.numeroHijos = numeroHijos;

        //        //Obtener listado de hijos del id seleccionado
        //        var ListadoHijosPadre = CatalogoEntity.ListadoHijosoCatalogoPorIdPadre(Convert.ToInt32(idCatalogo));
        //        ViewBag.ListadoHijosPadre = ListadoHijosPadre;

        //        //Obtener listado de hijos 
        //        var ListadoCatalogoPadre = CatalogoEntity.ListadoCatalogosPorIdSinOrdenar(Convert.ToInt32(idCatalogo));
        //        ViewBag.ListadoCatalogoPadre = ListadoCatalogoPadre;

        //        ViewBag.EtapaGeneral = new Catalogo();
        //        ViewBag.EstatusDetallado = new Catalogo();
        //        ViewBag.EstatusGeneral = new Catalogo();

        //        ViewBag.IdCatalogo = idCatalogo;

        //        return Json(new { Resultado = resultado }, JsonRequestBehavior.AllowGet);
        //    }
        //    catch (Exception ex)
        //    {
        //        return Json(new { Resultado = new RespuestaTransaccion { Estado = false, Respuesta = ex.Message } }, JsonRequestBehavior.AllowGet);
        //    }
        //}

        // GET: Subcatalogo/Create
        public ActionResult _EditSubcatalogo(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            else
            {
                //Titulo de la Pantalla
                ViewBag.TituloModal = Etiquetas.TituloPanelCreacionCatalogo;

                var catalogo = CatalogoEntity.ConsultarCatalogo(id);

                if (catalogo == null)
                {
                    return HttpNotFound();
                }
                else
                {
                    return PartialView(catalogo);
                }
            }

        }

        [HttpPost]
        public ActionResult _EditSubcatalogo(Catalogo catalogo)
        {
            //Titulo de la Pantalla
            ViewBag.TituloModal = Etiquetas.TituloPanelCreacionCatalogo;

            if (ModelState.IsValid)
            {
                RespuestaTransaccion resultado = CatalogoEntity.ActualizarSubCatalogo(catalogo);

                //Almacenar en una variable de sesion
                Session["Resultado"] = resultado.Respuesta;
                Session["Estado"] = resultado.Estado.ToString();

                if (resultado.Estado.ToString() == "True")
                {
                    var id_catalogo = ViewData["id_catalogo"] = System.Web.HttpContext.Current.Session["id_catalogo"] as String;
                    return RedirectToAction("IndexSubcatalogo/" + id_catalogo, "Catalogo");
                }
                else
                {
                    ViewBag.Resultado = resultado.Respuesta;
                    ViewBag.Estado = resultado.Estado.ToString();
                    Session["Resultado"] = "";
                    Session["Estado"] = "";
                    return PartialView(catalogo);
                }
            }
            else
            {
                return View(catalogo);
            }
        }
        // hh555
        public ActionResult DescargarReporteFormatoExcel()
        {
            var collection = CatalogoEntity.ListarCatalogos();

            //==========================================================================================
            //****************************DEFINICION DE LAS HOJAS*************************************//
            //==========================================================================================
            ExcelPackage ExcelPkg = new ExcelPackage();

            ExcelWorksheet Listado = ExcelPkg.Workbook.Worksheets.Add("Listado Catálogos");
            Listado.TabColor = System.Drawing.Color.Black;
            Listado.DefaultRowHeight = 12;

            ExcelWorksheet MigracionCabecera = ExcelPkg.Workbook.Worksheets.Add("Listado Subcatálogos");
            MigracionCabecera.TabColor = System.Drawing.Color.Black;
            MigracionCabecera.DefaultRowHeight = 12;


            List<string> columnas = new List<string>()
            {
                "ID",
                "NOMBRE",
                "DESCRIPCIÓN",
                "CÓDIGO CATÁLOGO",
                "ESTADO"
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
                Listado.Cells[recordIndex, 1].Value = item.Id;
                Listado.Cells[recordIndex, 2].Value = item.Nombre;
                Listado.Cells[recordIndex, 3].Value = item.Descripcion;
                Listado.Cells[recordIndex, 4].Value = item.Codigo_catalogo;
                Listado.Cells[recordIndex, 5].Value = item.Estado;
                recordIndex++;
            }

            for (int i = 1; i <= columnas.Count; i++)
            {
                Listado.Column(i).AutoFit();
                Listado.Column(i).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;

            }

            var collectionDetalle = CatalogoEntity.ListarSubcatalogos();

            List<string> columnasDetalle = new List<string>()
            {

                "ID",
                "PADRE",
                "NOMBRE",
                "DESCRIPCIÓN",
                "CÓDIGO CATÁLOGO",
                "ESTADO"
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
                MigracionCabecera.Cells[recordIndex, 1].Value = item.Id;
                MigracionCabecera.Cells[recordIndex, 2].Value = item.Padre;
                MigracionCabecera.Cells[recordIndex, 3].Value = item.Nombre;
                MigracionCabecera.Cells[recordIndex, 4].Value = item.Descripcion;
                MigracionCabecera.Cells[recordIndex, 5].Value = item.codigo_catalogo;
                MigracionCabecera.Cells[recordIndex, 6].Value = item.Estado;

                recordIndex++;
            }

            for (int i = 1; i <= columnasDetalle.Count; i++)
            {
                MigracionCabecera.Column(i).AutoFit();
                MigracionCabecera.Column(i).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
            }

            return File(ExcelPkg.GetAsByteArray(), XlsxContentType, "ListadoCatálogos.xlsx");
        }

        public ActionResult DescargarReporteFormatoPDF()
        {
            // Seleccionar las columnas a exportar
            var results = CatalogoEntity.ListarCatalogos();

            var list = Reportes.SerializeToJSON(results);
            return Content(list, "application/json");
        }

        public ActionResult DescargarReporteFormatoCSV()
        {
            var comlumHeadrs = new string[]
            {
                "ID",
                "NOMBRE",
                "DESCRIPCION",
                "CODIGO_CATALOGO",
                "ESTADO",
            };

            var listado = (from item in CatalogoEntity.ListarCatalogos()
                           select new object[]
                                          {
                               item.Id,
                               item.Nombre,
                               item.Descripcion,
                               item.Codigo_catalogo,
                               item.Estado
                                          }).ToList();

            // Build the file content
            var employeecsv = new StringBuilder();
            listado.ForEach(line =>
            {
                employeecsv.AppendLine(string.Join(",", line));
            });

            byte[] buffer = Encoding.Default.GetBytes($"{string.Join(",", comlumHeadrs)}\r\n{employeecsv.ToString()}");
            return File(buffer, "text/csv", $"Listado_Catálogos.csv");
        }

        [ActionName("BorrarSubcatalogo")]
        public ActionResult BorrarSubcatalogo(int? id)
        {
            if (id == 0)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }

            RespuestaTransaccion resultado = CatalogoEntity.EliminarSubcatalogo(id.Value);

            //Almacenar en una variable de sesion
            Session["Resultado"] = resultado.Respuesta;
            Session["Estado"] = resultado.Estado.ToString();

            //return RedirectToAction("Index");
            return Json(new { Resultado = resultado }, JsonRequestBehavior.AllowGet);

        }

        public JsonResult GetSubcatalogos(int id, string tipo)
        {
            var datos = CatalogoEntity.ListarCatalogosPorId(Convert.ToInt32(id), "");
            //datos = datos.Where(d => d.codigo_catalogo == tipo).ToList();
            ViewBag.ListadoCatalogos = datos;
            return Json(datos, JsonRequestBehavior.AllowGet);
        }
    }
}

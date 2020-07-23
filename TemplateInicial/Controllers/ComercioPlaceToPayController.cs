using System;
using System.Collections.Generic;
using System.Linq;
using System.Web.Mvc;
using System.Threading.Tasks;

//Extensión para Query string dinámico
using System.Linq.Dynamic;
using System.Web;
using GestionPPM.Entidades.Metodos;
using GestionPPM.Entidades.Modelo.PlaceToPay;
using GestionPPM.Repositorios;
using Seguridad.Helper;
using GestionPPM.Entidades.Modelo;
using System.IO;

namespace TemplateInicial.Controllers
{
    [Autenticado]
    public class ComercioPlaceToPayController : BaseAppController
    {
        private List<string> columnasReportesBasicos = new List<string> { "FECHA DE SOLICITUD", "USUARIO", "EQUIPOS", "HERRAMIENTAS ADICIONALES", "EMPRESA", "CARGO", "DEPARTAMENTO" };

        // GET: ComercioPlaceToPay
        public ActionResult Index()
        {
            return View();
        }

        [HttpGet]
        public async Task<PartialViewResult> _IndexGrid(string search, string sort = "", string order = "", long? page = 1)
        {
            page = page > 0 ? page - 1 : page;
            int totalPaginas = 1;
            var listado = new List<ComercioPlaceToPayInfo>();


            ViewBag.NombreListado = Etiquetas.TituloGridComercioPlaceToPay;
            //Controlar permisos
            var user = ViewData["usuario"] = System.Web.HttpContext.Current.Session["usuario"];
            var usuario = int.Parse(user.ToString());
            string nombreControlador = ControllerContext.RouteData.Values["controller"].ToString();
            ViewBag.NombreControlador = nombreControlador;

            ViewBag.AccionesUsuario = ManejoPermisosEntity.ListadoAccionesCatalogoUsuario(usuario, nombreControlador);

            //Obtener Acciones del controlador
            ViewBag.AccionesControlador = GetMetodosControlador(nombreControlador);

            try
            {
                var query = (HttpContext.Request.Params.Get("QUERY_STRING") ?? "").ToString();

                var dynamicQueryString = GetQueryString(query);
                var whereClause = BuildWhereDynamicClause(dynamicQueryString);

                //Siempre y cuando no haya filtros definidos en el Grid
                if (string.IsNullOrEmpty(whereClause))
                {
                    if (!string.IsNullOrEmpty(sort) && !string.IsNullOrEmpty(order))
                        listado = ComercioPlaceToPayEntity.ListadoComercioPlaceToPay(page.Value).OrderBy(sort + " " + order).ToList();
                    else
                        listado = ComercioPlaceToPayEntity.ListadoComercioPlaceToPay(page.Value).ToList();
                }

                search = !string.IsNullOrEmpty(search) ? search.Trim() : "";

                if (!string.IsNullOrEmpty(search))//filter
                {
                    listado = ComercioPlaceToPayEntity.ListadoComercioPlaceToPay(null, search);
                }

                if (!string.IsNullOrEmpty(whereClause) && string.IsNullOrEmpty(search))
                {
                    if (!string.IsNullOrEmpty(sort) && !string.IsNullOrEmpty(order))
                        listado = ComercioPlaceToPayEntity.ListadoComercioPlaceToPay(null, null, whereClause).OrderBy(sort + " " + order).ToList();
                    else
                        listado = ComercioPlaceToPayEntity.ListadoComercioPlaceToPay(null, null, whereClause);
                }
                else
                {

                    if (string.IsNullOrEmpty(search))
                        totalPaginas = ComercioPlaceToPayEntity.ObtenerTotalRegistrosListadoComercioPlaceToPay();
                }

                ViewBag.TotalPaginas = totalPaginas;

                // Only grid query values will be available here.
                return PartialView(await Task.Run(() => listado));
            }
            catch (Exception ex)
            {
                ViewBag.TotalPaginas = totalPaginas;
                // Only grid query values will be available here.
                return PartialView(await Task.Run(() => listado));
            }
        }

        #region Funcionalidades Genéricas para Grids
        public string BuildWhereDynamicClause(Dictionary<string, object> queryString)
        {
            string query = string.Empty;
            try
            {
                List<string> clausulas = new List<string>();

                string contains = "{0} LIKE '%{1}%'";
                string equals = "{0} = '{1}'";
                string NotEquals = "{0} != '{1}'";
                string StartsWith = "{0} LIKE '{1}%'";
                string EndWith = "{0} LIKE '%{1}'";

                string where = string.Empty;

                foreach (KeyValuePair<string, object> item in queryString)
                {
                    string columna = item.Key.Split('-').FirstOrDefault();

                    if (item.Key.Contains("contains"))
                    {
                        //query += string.Format(contains, columna, (item.Value ?? "").ToString());
                        clausulas.Add(string.Format(contains, columna, (item.Value ?? "").ToString().Trim()));
                    }
                    if (item.Key.Contains("equals"))
                    {
                        //query += string.Format(equals, columna, (item.Value ?? "").ToString());
                        clausulas.Add(string.Format(equals, columna, item.Value.ToString().Trim()));
                    }
                    if (item.Key.Contains("not-equals"))
                    {
                        //query += string.Format(NotEquals, columna, (item.Value ?? "").ToString());
                        clausulas.Add(string.Format(NotEquals, columna, (item.Value ?? "").ToString().Trim()));
                    }
                    if (item.Key.Contains("starts-with"))
                    {
                        //query += string.Format(StartsWith, columna, (item.Value ?? "").ToString());
                        clausulas.Add(string.Format(StartsWith, columna, (item.Value ?? "").ToString().Trim()));
                    }
                    if (item.Key.Contains("ends-with"))
                    {
                        //query += string.Format(EndWith, columna, (item.Value ?? "").ToString());
                        clausulas.Add(string.Format(EndWith, columna, (item.Value ?? "").ToString().Trim()));
                    }

                    where = " WHERE ";
                }

                query += clausulas.Any() ? where + string.Join(" AND ", clausulas.ToArray()) : string.Empty;

                return query;
            }
            catch (Exception ex)
            {
                return query;
            }
        }

        public Dictionary<string, object> GetQueryString(string queryString)
        {
            Dictionary<string, object> querystringDic = new Dictionary<string, object>();
            try
            {
                var parsed = HttpUtility.ParseQueryString(queryString);
                querystringDic = parsed.AllKeys.ToDictionary(k => k, k => (object)parsed[k]);

                querystringDic.Remove("_");

                //Parametros ya incluidos en el request del método IndexGrid
                querystringDic.Remove("search");
                querystringDic.Remove("page");
                querystringDic.Remove("order");
                querystringDic.Remove("sort");

                return querystringDic;
            }
            catch (Exception ex)
            {
                return querystringDic;
            }
        }
        #endregion

        public async Task<ActionResult> Formulario(int? id)
        {
            try
            {
                ComercioPlaceToPay model = new ComercioPlaceToPay { FechaAfiliacion = DateTime.Now, Fecha = DateTime.Now };
                var usuario = UsuarioEntity.ConsultarUsuario(GetCurrentUser());

                ViewBag.Usuario = usuario;

                if (id.HasValue)
                    model = await ComercioPlaceToPayEntity.GetComercioPlaceToPayAsync(id.Value);

                return View(model);
            }
            catch (Exception ex)
            {
                return View(new ComercioPlaceToPay { FechaAfiliacion = DateTime.Now, Fecha = DateTime.Now });
            }
        }

        //Busqueda por RUC
        public  JsonResult _GetInformacionPrincipalComercio(string busqueda)
        {
            List<AutoCompleteUI> items = new List<AutoCompleteUI>();
            busqueda = (busqueda ?? "").ToLower().Trim();

            var results = ComercioPlaceToPayEntity.ConsultarComercioPlaceToPayPorRUC(busqueda).GroupBy(p => p.RUC).Select(g => g.First()).ToList();

            items = results.Where(o => o.RUC.ToLower().Contains(busqueda)).Select(o => new AutoCompleteUI(o.IDComercioPlaceToPay, o.RUC, string.Empty, new Dictionary<string, ComercioPlaceToPayInfo>(){
                { o.IDComercioPlaceToPay.ToString(), new ComercioPlaceToPayInfo {
                    IDComercioPlaceToPay = o.IDComercioPlaceToPay,
                    TipoCodigo = o.TipoCodigo,
                    TextoCatalogoTipoCodigo = o.TextoCatalogoTipoCodigo,
                    CodigoUnico = o.CodigoUnico,
                    FechaAfiliacion = o.FechaAfiliacion,
                    Establecimiento = o.Establecimiento,
                    MID = o.MID,
                    Especialidad = o.Especialidad,
                    NombreRepresentanteLegal = o.NombreRepresentanteLegal,
                    RazonSocial = o.RazonSocial,
                    DireccionComercio = o.DireccionComercio,
                    Mail = o.Mail,
                    Marca = o.Marca,
                    Prefijo1 = o.Prefijo1,
                    Prefijo2 = o.Prefijo2,
                    Prefijo3 = o.Prefijo3,
                    Prefijo4 = o.Prefijo4,
                    Telefono1 = o.Telefono1,
                    Telefono2 = o.Telefono2,
                    Telefono3 = o.Telefono3,
                    Telefono4 = o.Telefono4,
                    TelefonoCompleto1 = o.TelefonoCompleto1
                } }
            })).Take(10).ToList();
            return Json(new { results = items }, JsonRequestBehavior.AllowGet);
        }

        [HttpPost]
        public async Task<ActionResult> Create(ComercioPlaceToPay formulario, List<string> archivos)
        {
            try
            {
                if (!Validaciones.VerificaIdentificacion(formulario.IdentificacionRepresentanteLegal))
                    return Json(new { Resultado = new RespuestaTransaccion { Estado = false, Respuesta = string.Format(Mensajes.MensajeIdentificacionIncorrecto, formulario.IdentificacionRepresentanteLegal) } }, JsonRequestBehavior.AllowGet);

                if (!Validaciones.VerificaIdentificacion(formulario.IdentificacionAdministrador))
                    return Json(new { Resultado = new RespuestaTransaccion { Estado = false, Respuesta = string.Format(Mensajes.MensajeIdentificacionIncorrecto, formulario.IdentificacionAdministrador) } }, JsonRequestBehavior.AllowGet);

                if (!Validaciones.ValidarMail(formulario.CorreoElectronicoEjecutivo))
                    return Json(new { Resultado = new RespuestaTransaccion { Estado = false, Respuesta = string.Format(Mensajes.MensajeEmailIncorrecto, formulario.CorreoElectronicoEjecutivo) } }, JsonRequestBehavior.AllowGet);

                if (!Validaciones.ValidarMail(formulario.CorreoElectronicoLiderProyectoComercio))
                    return Json(new { Resultado = new RespuestaTransaccion { Estado = false, Respuesta = string.Format(Mensajes.MensajeEmailIncorrecto, formulario.CorreoElectronicoLiderProyectoComercio) } }, JsonRequestBehavior.AllowGet);

                #region Guardar archivos adjuntos

                string mensajeAdvertenciaAdjuntoNoGenerado = string.Empty;
                bool ok = false;

                if (archivos != null)
                {
                    string rutaBase = basePathRepositorioDocumentos + "\\GESTION_PPM\\PLACETOPAY\\PropuestasAdjuntas";
                    bool existeRutaDisco = Directory.Exists(rutaBase); // VERIFICAR SI ESA RUTA EXISTE

                    if (!existeRutaDisco)
                        Directory.CreateDirectory(rutaBase);

                    string adjuntoDetalle = archivos.ElementAt(0);
                    string identificadorArchivo = !string.IsNullOrEmpty(formulario.CodigoUnico) ? formulario.CodigoUnico : Guid.NewGuid().ToString().Substring(0, 10);

                    string pathFinal = Path.Combine(rutaBase, "PropuestaComercio " + identificadorArchivo + ".pdf");

                    //Decodificar y guardar en ruta.
                    ok = Auxiliares.Base64Decode(adjuntoDetalle, pathFinal);

                    //Solo si el archivo se logra decodificar correctamente
                    if (ok)
                        formulario.PropuestaAdjuntada = pathFinal;
                    else
                        mensajeAdvertenciaAdjuntoNoGenerado = Mensajes.MensajeAdjuntoFallido;
                }
                else {
                    return Json(new { Resultado = new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeArchivoNoExiste } }, JsonRequestBehavior.AllowGet);
                }

                #endregion

                formulario.CreatedAt = DateTime.Now;
                formulario.CreatedBy = GetCurrentUser();

                var Resultado = await ComercioPlaceToPayEntity.CrearComercioPlaceToPay(formulario);

                //Agregar mensaje de advertencia si el archivo no fue generado correctamente
                if (!ok)
                    Resultado.Respuesta = Resultado.Respuesta + mensajeAdvertenciaAdjuntoNoGenerado;

                return Json(new { Resultado = Resultado }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                return Json(new { Resultado = new RespuestaTransaccion { Estado = false, Respuesta = ex.Message + " ; " + ex.InnerException.Message } }, JsonRequestBehavior.AllowGet);
            }
        }

        [HttpPost]
        public async Task<ActionResult> Edit(ComercioPlaceToPay formulario, List<string> archivos)
        {
            try
            {
                if (!Validaciones.VerificaIdentificacion(formulario.IdentificacionRepresentanteLegal))
                    return Json(new { Resultado = new RespuestaTransaccion { Estado = false, Respuesta = string.Format(Mensajes.MensajeIdentificacionIncorrecto, formulario.IdentificacionRepresentanteLegal) } }, JsonRequestBehavior.AllowGet);

                if (!Validaciones.VerificaIdentificacion(formulario.IdentificacionAdministrador))
                    return Json(new { Resultado = new RespuestaTransaccion { Estado = false, Respuesta = string.Format(Mensajes.MensajeIdentificacionIncorrecto, formulario.IdentificacionAdministrador) } }, JsonRequestBehavior.AllowGet);

                if (!Validaciones.ValidarMail(formulario.CorreoElectronicoEjecutivo))
                    return Json(new { Resultado = new RespuestaTransaccion { Estado = false, Respuesta = string.Format(Mensajes.MensajeEmailIncorrecto, formulario.CorreoElectronicoEjecutivo) } }, JsonRequestBehavior.AllowGet);

                if (!Validaciones.ValidarMail(formulario.CorreoElectronicoLiderProyectoComercio))
                    return Json(new { Resultado = new RespuestaTransaccion { Estado = false, Respuesta = string.Format(Mensajes.MensajeEmailIncorrecto, formulario.CorreoElectronicoLiderProyectoComercio) } }, JsonRequestBehavior.AllowGet);

                #region Guardar archivos adjuntos

                string mensajeAdvertenciaAdjuntoNoGenerado = string.Empty;
                bool ok = false;
                if (archivos != null) {
                    string rutaBase = basePathRepositorioDocumentos + "\\GESTION_PPM\\PLACETOPAY\\PropuestasAdjuntas";
                    bool existeRutaDisco = Directory.Exists(rutaBase); // VERIFICAR SI ESA RUTA EXISTE

                    if (!existeRutaDisco)
                        Directory.CreateDirectory(rutaBase);

                    string adjuntoDetalle = archivos.ElementAt(0);
                    string identificadorArchivo = !string.IsNullOrEmpty(formulario.CodigoUnico) ? formulario.CodigoUnico : Guid.NewGuid().ToString().Substring(0, 10);

                    string pathFinal = Path.Combine(rutaBase, "PropuestaComercio " + identificadorArchivo + ".pdf");

                    //Decodificar y guardar en ruta.
                    ok = Auxiliares.Base64Decode(adjuntoDetalle, pathFinal);

                    //Solo si el archivo se logra decodificar correctamente
                    if (ok)
                        formulario.PropuestaAdjuntada = pathFinal;
                    else
                        mensajeAdvertenciaAdjuntoNoGenerado = Mensajes.MensajeAdjuntoFallido;
                }

                #endregion

                formulario.UpdatedAt = DateTime.Now;
                formulario.UpdatedBy = GetCurrentUser();

                var Resultado = await ComercioPlaceToPayEntity.ActualizarComercioPlaceToPay(formulario);

                return Json(new { Resultado = Resultado }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                return Json(new { Resultado = new RespuestaTransaccion { Estado = false, Respuesta = ex.Message  } }, JsonRequestBehavior.AllowGet);
            }
        }


        #region Metodos sin uso para que funcionen los permisos
        public ActionResult IndexGrid()
        {
            return View();
        }
        #endregion

        #region REPORTES BASICOS
        public ActionResult DescargarReporteFormatoExcel()
        {
            var collection = ComercioPlaceToPayEntity.ListadoComercioPlaceToPay();
            var package = GetEXCEL(columnasReportesBasicos, collection.Cast<object>().ToList());
            return File(package.GetAsByteArray(), XlsxContentType, "ListadoReporte.xlsx");
        }

        public ActionResult DescargarReporteFormatoCSV()
        {
            var collection = ComercioPlaceToPayEntity.ListadoComercioPlaceToPay();
            byte[] buffer = GetCSV(columnasReportesBasicos, collection.Cast<object>().ToList());
            return File(buffer, CSVContentType, $"ListadoReporte.csv");
        }
        #endregion



    }
}
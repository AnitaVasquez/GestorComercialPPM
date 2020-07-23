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
using iTextSharp.text.pdf;
using iTextSharp.text;
using Newtonsoft.Json;

namespace TemplateInicial.Controllers
{
    [Autenticado]
    public class SolicitudController : Controller
    {
        GestionPPMEntities db = new GestionPPMEntities();
        private const string XlsxContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
        private int rolID = Convert.ToInt32(System.Web.HttpContext.Current.Session["rolID"]);

        public ActionResult Create()
        {
            //Listado de Tipo 
            var Tipo = CatalogoEntity.ObtenerListadoCatalogosByCodigo("TSL-01");
            ViewBag.ListadoTipo = Tipo;

            //Listado Marca
            var Marca = CatalogoEntity.ObtenerListadoCatalogosByCodigo("MRC-01");
            ViewBag.ListadoMarca = Marca;

            //Listado Area Encargada
            var AreaEncargada = CatalogoEntity.ObtenerListadoCatalogosByCodigo("AEN-01");
            ViewBag.AreaEncargada = AreaEncargada;

            //Listado Solcitante DCE
            var solicitanteDCE = CatalogoEntity.ObtenerListadoCatalogosByCodigo("SDCE-01");
            ViewBag.SolicitanteDCE = solicitanteDCE;

            //Obtener el codigo de usuario que se logeo
            var usuarioSesion = ViewData["usuario"] = System.Web.HttpContext.Current.Session["usuario"];
            var idUsuario = Convert.ToInt16(usuarioSesion);

            ViewBag.DatosUsuario = UsuarioEntity.ConsultarUsuario(idUsuario);

            //Obtener una inicializacion de solicitud
            SolicitudCliente solicitud = new SolicitudCliente();

            //Obtener el codigo de general
            var codigoGeneral = Tipo.Where(t => Convert.ToInt32(t.Value) == 737).FirstOrDefault().Value;
            solicitud.id_tipo = Convert.ToInt16(codigoGeneral);

            //Listado Subtipo  
            var Subtipo = CatalogoEntity.ConsultarCatalogoPorPadre(solicitud.id_tipo.HasValue ? solicitud.id_tipo.Value : 0, "SUBTIPO");
            ViewBag.ListadoSubtipo = Subtipo;

            //Obtener el codigo de subtipo n/a
            var codigoSubtipo = Subtipo.Where(t => t.Text.ToUpper() == "N/A").FirstOrDefault().Value;
            solicitud.id_subtipo = Convert.ToInt16(codigoSubtipo);

            //Obtener el codigo de marca n/a
            var codigoMarca = Marca.Where(t => t.Text.ToUpper() == "N/A").FirstOrDefault().Value;
            solicitud.id_marca = Convert.ToInt16(codigoMarca);




            return View(solicitud);
        }

        [HttpPost]
        public ActionResult Create(int? codigo)
        {
            try
            {
                //variable de guardado
                bool guardarSolicitud = false;
                bool sftp = false;
                bool excel = false;
                bool zip = false;
                bool matrizexcel = false;

                //Listado de Tipo 
                var Tipo = CatalogoEntity.ObtenerListadoCatalogosByCodigo("TSL-01");
                ViewBag.ListadoTipo = Tipo;

                //Listado Marca
                var Marca = CatalogoEntity.ObtenerListadoCatalogosByCodigo("MRC-01");
                ViewBag.ListadoMarca = Marca;

                //Obtener el codigo de usuario que se logeo
                var usuarioSesion = ViewData["usuario"] = System.Web.HttpContext.Current.Session["usuario"];
                var idUsuario = Convert.ToInt16(usuarioSesion);
                ViewBag.DatosUsuario = UsuarioEntity.ConsultarUsuario(idUsuario);

                //Obtener los datos del usuario
                int codigoUsuario = Convert.ToInt32(HttpContext.Request.Params.Get("codigoUsuario"));

                //Obtener si es un mantenimiento urgente    
                var urgente = HttpContext.Request.Params.Get("urgente");

                //Obtener los datos de la solicitud
                var JsonSolicitud = HttpContext.Request.Params.Get("solicitud");
                SolicitudCliente solicitud = JsonConvert.DeserializeObject<SolicitudCliente>(JsonSolicitud);

                //Listado Subtipo  
                var Subtipo = CatalogoEntity.ConsultarCatalogoPorPadre(solicitud.id_tipo.HasValue ? solicitud.id_tipo.Value : 0, "SUBTIPO");
                ViewBag.ListadoSubtipo = Subtipo;

                //Obtener los datos de campos personalizados 
                var campos = HttpContext.Request.Params.Get("camposPersonalizados");
                List<CamposPersonalizadosParcial> camposPersonalizados = JsonConvert.DeserializeObject<List<CamposPersonalizadosParcial>>(campos);

                //Obtener los datos de url externo
                var urls = HttpContext.Request.Params.Get("urlExterno");
                List<UrlExternoParcial> urlExterno = JsonConvert.DeserializeObject<List<UrlExternoParcial>>(urls);

                //Obtener los datos de url soporte
                var urlsSoporte = HttpContext.Request.Params.Get("listadoUrlSoporte");
                List<UrlExternoParcial> urlSoporte = JsonConvert.DeserializeObject<List<UrlExternoParcial>>(urlsSoporte);

                //Obtener los datos de portales
                List<int> portales = new List<int>();

                try
                {

                    var idPortales = HttpContext.Request.Params.Get("portales");

                    portales = JsonConvert.DeserializeObject<List<int>>(idPortales);

                }

                catch

                {

                    portales = new List<int>();

                }

                //Obtener datos Adjuntos para la Solicitud
                HttpFileCollectionBase files = Request.Files;
                List<AdjutoSolicitudParcial> adjuntos = new List<AdjutoSolicitudParcial>();

                //Validacion de adjuntos
                var codigoGeneral = Tipo.Where(t => Convert.ToInt32(t.Value) == 758).FirstOrDefault().Value;
                int valorCodigoGeneral = Convert.ToInt16(codigoGeneral);

                //Validar si es icare o eclub y validar el tamaño de las imagenes y el formato de los archivos
                var codigoIcare = Tipo.Where(t => Convert.ToInt32(t.Value) == 735).FirstOrDefault().Value;
                int valorCodigoIcare = Convert.ToInt16(codigoIcare);

                //Validar si es icare o eclub y validar el tamaño de las imagenes y el formato de los archivos
                var codigoEclub = Tipo.Where(t => Convert.ToInt32(t.Value) == 734).FirstOrDefault().Value;
                int valorCodigoEclub = Convert.ToInt16(codigoEclub);

                //Validar si es icare o eclub y validar el tamaño de las imagenes y el formato de los archivos
                var codigoMantenimiento = Tipo.Where(t => Convert.ToInt32(t.Value) == 737).FirstOrDefault().Value;
                int valorCodigoMantenimiento = Convert.ToInt16(codigoMantenimiento);

                //Validar envios icare  
                var codigoEnviosIcare = Tipo.Where(t => Convert.ToInt32(t.Value) == 736).FirstOrDefault().Value;
                int valorcodigoEnviosIcare = Convert.ToInt16(codigoEnviosIcare);

                //Validar Html 
                var codigoHtml = Tipo.Where(t => Convert.ToInt32(t.Value) == 739).FirstOrDefault().Value;
                int valorcodigoHtml = Convert.ToInt16(codigoHtml);

                //Validar Html y Tags
                var codigoHtmlTag = Tipo.Where(t => Convert.ToInt32(t.Value) == 740).FirstOrDefault().Value;
                int valorcodigoHtmlTags = Convert.ToInt16(codigoHtmlTag);

                //Validar Tags
                var codigoTag = Tipo.Where(t => Convert.ToInt32(t.Value) == 738).FirstOrDefault().Value;
                int valorcodigoTags = Convert.ToInt16(codigoTag);

                //Si es una solicitud de tipo General
                if (solicitud.id_tipo == valorCodigoGeneral)
                {
                    for (int i = 0; i < files.Count; i++)
                    {
                        HttpPostedFileBase file = files[i];
                        string path = string.Empty;
                        string FileName = "";

                        path = file.FileName;
                        FileName = file.FileName;

                        //agregar a mi lista de adjuntos
                        adjuntos.Add(new AdjutoSolicitudParcial
                        {
                            tipoAdjunto = "Adjunto Soporte",
                            nombreAdjunto = path
                        });
                    }

                    guardarSolicitud = true;
                }
                else
                {
                    //validar que no tenga mas archivos adjuntos que los que define en su cantidad
                    int numeroSolicitudes = solicitud.cantidad.Value;

                    //Si es una solicitud de tipo Icare o Eclub
                    if (solicitud.id_tipo == valorCodigoIcare || solicitud.id_tipo == valorCodigoEclub)
                    {
                        //validar los adjuntos
                        if (files.Count == 0)
                        {
                            return Json(new { Resultado = new RespuestaTransaccion { Estado = false, Respuesta = "Debe adjuntar la pieza(s) de referencia" } }, JsonRequestBehavior.AllowGet);
                        }
                        else
                        {
                            int adjuntosGifs = 0;

                            //Contar Archivos solo wepb
                            for (int i = 0; i < files.Count; i++)
                            {
                                HttpPostedFileBase file = files[i];

                                string path = string.Empty;
                                path = file.FileName;

                                string extension = Path.GetExtension(path);
                                extension = extension.ToLower();

                                if (extension == ".webp")
                                {
                                    adjuntosGifs += 1;
                                }
                            }

                            //Si los adjuntos imagenes - gifs no coinciden 
                            if ((files.Count - adjuntosGifs) > numeroSolicitudes)
                            {
                                return Json(new { Resultado = new RespuestaTransaccion { Estado = false, Respuesta = "Las piezas de referencia adjuntadas no pueden ser mayor a la cantidad de solicitudes" } }, JsonRequestBehavior.AllowGet);
                            }
                            else
                            {
                                if ((files.Count - adjuntosGifs) == numeroSolicitudes)
                                {
                                    //validar la extension y el tamaño del adjunto
                                    //Recorrer todos los adjuntos
                                    for (int i = 0; i < files.Count; i++)
                                    {
                                        HttpPostedFileBase file = files[i];
                                        string path = string.Empty;
                                        string FileName = "";

                                        path = file.FileName;
                                        FileName = file.FileName;

                                        string extension = Path.GetExtension(path);
                                        extension = extension.ToLower();
                                        var tamaño = (files[i].ContentLength) / 1024;

                                        if (extension != ".jpg" && extension != ".jpeg" && extension != ".webp")
                                        {
                                            guardarSolicitud = false;
                                            return Json(new { Resultado = new RespuestaTransaccion { Estado = false, Respuesta = "La extensión de formato de carga está incorrecto" } }, JsonRequestBehavior.AllowGet);
                                        }
                                        else
                                        {
                                            //pero para imagenes
                                            if (extension == ".jpg" || extension == ".jpeg")
                                            {
                                                //agregar a mi lista de adjuntos
                                                adjuntos.Add(new AdjutoSolicitudParcial
                                                {
                                                    tipoAdjunto = "Pieza de Referencia",
                                                    nombreAdjunto = path
                                                });

                                                if (tamaño > 550)
                                                {
                                                    guardarSolicitud = false;
                                                    return Json(new { Resultado = new RespuestaTransaccion { Estado = false, Respuesta = "Los archivos no deben superar los 550KB" } }, JsonRequestBehavior.AllowGet);
                                                }
                                                else
                                                {
                                                    guardarSolicitud = true;
                                                }
                                            }

                                            //pero los gifs
                                            if (extension == ".webp")
                                            {
                                                //agregar a mi lista de adjuntos
                                                adjuntos.Add(new AdjutoSolicitudParcial
                                                {
                                                    tipoAdjunto = "Gif de Referencia",
                                                    nombreAdjunto = path
                                                });

                                                if (tamaño > 1000)
                                                {
                                                    guardarSolicitud = false;
                                                    return Json(new { Resultado = new RespuestaTransaccion { Estado = false, Respuesta = "Los archivos no deben superar 1MB" } }, JsonRequestBehavior.AllowGet);
                                                }
                                                else
                                                {
                                                    guardarSolicitud = true;
                                                }
                                            }
                                        }
                                    }

                                }
                                else
                                {
                                    if (files.Count < numeroSolicitudes)
                                    {
                                        return Json(new { Resultado = new RespuestaTransaccion { Estado = false, Respuesta = "Las piezas de referencia adjuntadas no pueden ser menor a la cantidad de solicitudes" } }, JsonRequestBehavior.AllowGet);
                                    }
                                }
                            }
                        }
                    }

                    //Si es de mantenimiento
                    if (solicitud.id_tipo == valorCodigoMantenimiento)
                    {
                        //Validar que tenga al menos un detalle de url de soporte
                        if (urlSoporte == null || urlSoporte.Count == 0)
                        {
                            guardarSolicitud = false;
                            return Json(new { Resultado = new RespuestaTransaccion { Estado = false, Respuesta = "Debe registrar al menos una url para el soporte" } }, JsonRequestBehavior.AllowGet);
                        }
                        else
                        {
                            for (int i = 0; i < files.Count; i++)
                            {
                                HttpPostedFileBase file = files[i];
                                string path = string.Empty;
                                string FileName = "";

                                path = file.FileName;
                                FileName = file.FileName;

                                //agregar a mi lista de adjuntos
                                adjuntos.Add(new AdjutoSolicitudParcial
                                {
                                    tipoAdjunto = "Adjunto Soporte",
                                    nombreAdjunto = path
                                });
                            }
                            guardarSolicitud = true;
                        }
                    }

                    //Si es una solicitud de envios icare
                    if (solicitud.id_tipo == valorcodigoEnviosIcare)
                    {
                        //validar los adjuntos
                        if (files.Count == 0)
                        {
                            return Json(new { Resultado = new RespuestaTransaccion { Estado = false, Respuesta = "Debe adjuntar la imagen SFTP y el archivo Excel con formato de carga" } }, JsonRequestBehavior.AllowGet);
                        }
                        else
                        {
                            if (files.Count != 2)
                            {
                                return Json(new { Resultado = new RespuestaTransaccion { Estado = false, Respuesta = "Debe adjuntar solo 2 archivos la imagen SFTP y el archivo Excel con formato de carga" } }, JsonRequestBehavior.AllowGet);
                            }
                            else
                            {
                                //Recorrer todos los adjuntos
                                for (int i = 0; i < files.Count; i++)
                                {
                                    HttpPostedFileBase file = files[i];
                                    string path = string.Empty;
                                    string FileName = "";

                                    path = file.FileName;
                                    FileName = file.FileName;

                                    string extension = Path.GetExtension(path);
                                    extension = extension.ToLower();

                                    if (extension != ".jpg" && extension != ".jpeg" && extension != ".xls" && extension != ".xlsm" && extension != ".xlsx")
                                    {
                                        return Json(new { Resultado = new RespuestaTransaccion { Estado = false, Respuesta = "La extensión de formato de carga está incorrecto" } }, JsonRequestBehavior.AllowGet);
                                    }
                                    else
                                    {
                                        if (extension == ".jpg" || extension == ".jpeg")
                                        {
                                            //agregar a mi lista de adjuntos
                                            adjuntos.Add(new AdjutoSolicitudParcial
                                            {
                                                tipoAdjunto = "SFTP",
                                                nombreAdjunto = path
                                            });
                                            sftp = true;
                                        }
                                        if (extension == ".xls" || extension == ".xlsm" || extension == ".xlsx")
                                        {
                                            //agregar a mi lista de adjuntos
                                            adjuntos.Add(new AdjutoSolicitudParcial
                                            {
                                                tipoAdjunto = "Formato de Carga",
                                                nombreAdjunto = path
                                            });
                                            excel = true;
                                        }
                                    }
                                }

                                //si tengo los dos archivos paso
                                if (sftp == true && excel == true)
                                {
                                    guardarSolicitud = true;
                                }
                                else
                                {
                                    guardarSolicitud = false;
                                    return Json(new { Resultado = new RespuestaTransaccion { Estado = false, Respuesta = "Debe adjuntar solo 2 archivos la imagen SFTP y el archivo Excel con formato de carga" } }, JsonRequestBehavior.AllowGet);
                                }
                            }
                        }
                    }

                    //Si es de html5
                    if (solicitud.id_tipo == valorcodigoHtml)
                    {
                        if (files.Count != 2)
                        {
                            return Json(new { Resultado = new RespuestaTransaccion { Estado = false, Respuesta = "Debe arjuntar 2 archivos, el de matriz y materiales" } }, JsonRequestBehavior.AllowGet);
                        }
                        else
                        {
                            for (int i = 0; i < files.Count; i++)
                            {
                                HttpPostedFileBase file = files[i];
                                string path = string.Empty;
                                string FileName = "";

                                path = file.FileName;
                                FileName = file.FileName;

                                string extension = Path.GetExtension(path);
                                extension = extension.ToLower();

                                if (extension != ".zip" && extension != ".rar" && extension != ".xls" && extension != ".xlsm" && extension != ".xlsx")
                                {
                                    return Json(new { Resultado = new RespuestaTransaccion { Estado = false, Respuesta = "La extensión de formato de carga está incorrecto" } }, JsonRequestBehavior.AllowGet);
                                }
                                else
                                {
                                    if (extension == ".zip" || extension == ".rar")
                                    {
                                        //agregar a mi lista de adjuntos
                                        adjuntos.Add(new AdjutoSolicitudParcial
                                        {
                                            tipoAdjunto = "Materiales",
                                            nombreAdjunto = path
                                        });
                                        zip = true;
                                    }
                                    if (extension == ".xls" || extension == ".xlsm" || extension == ".xlsx")
                                    {
                                        //agregar a mi lista de adjuntos
                                        adjuntos.Add(new AdjutoSolicitudParcial
                                        {
                                            tipoAdjunto = "Matriz",
                                            nombreAdjunto = path
                                        });
                                        matrizexcel = true;
                                    }
                                }
                            }

                            //si tengo los dos archivos paso
                            if (zip == true && matrizexcel == true)
                            {
                                guardarSolicitud = true;
                            }
                            else
                            {
                                guardarSolicitud = false;
                                return Json(new { Resultado = new RespuestaTransaccion { Estado = false, Respuesta = "Debe adjuntar solo 2 archivos, el de matriz y materiales" } }, JsonRequestBehavior.AllowGet);
                            }
                        }
                    }

                    //Si es de tags
                    if (solicitud.id_tipo == valorcodigoTags)
                    {
                        //Validar que tenga al menos un detalle de url de soporte
                        if (urlSoporte == null || urlSoporte.Count == 0)
                        {
                            guardarSolicitud = false;
                            return Json(new { Resultado = new RespuestaTransaccion { Estado = false, Respuesta = "Debe registrar el url de tipificación" } }, JsonRequestBehavior.AllowGet);
                        }
                        else
                        {
                            guardarSolicitud = true;
                        }
                    }

                    //Si es de html5 y  tags
                    if (solicitud.id_tipo == valorcodigoHtmlTags)
                    {
                        if (files.Count != 2)
                        {
                            return Json(new { Resultado = new RespuestaTransaccion { Estado = false, Respuesta = "Debe arjuntar 2 archivos, el de matriz y materiales" } }, JsonRequestBehavior.AllowGet);
                        }
                        else
                        {
                            //Validar que tenga al menos un detalle de url de soporte
                            if (urlSoporte == null || urlSoporte.Count == 0)
                            {
                                guardarSolicitud = false;
                                return Json(new { Resultado = new RespuestaTransaccion { Estado = false, Respuesta = "Debe registrar el url de tipificación" } }, JsonRequestBehavior.AllowGet);
                            }
                            else
                            {
                                for (int i = 0; i < files.Count; i++)
                                {
                                    HttpPostedFileBase file = files[i];
                                    string path = string.Empty;
                                    string FileName = "";

                                    path = file.FileName;
                                    FileName = file.FileName;

                                    string extension = Path.GetExtension(path);
                                    extension = extension.ToLower();

                                    if (extension != ".zip" && extension != ".rar" && extension != ".xls" && extension != ".xlsm" && extension != ".xlsx")
                                    {
                                        return Json(new { Resultado = new RespuestaTransaccion { Estado = false, Respuesta = "La extensión de formato de carga está incorrecto" } }, JsonRequestBehavior.AllowGet);
                                    }
                                    else
                                    {
                                        if (extension == ".zip" || extension == ".rar")
                                        {
                                            //agregar a mi lista de adjuntos
                                            adjuntos.Add(new AdjutoSolicitudParcial
                                            {
                                                tipoAdjunto = "Materiales",
                                                nombreAdjunto = path
                                            });
                                            zip = true;
                                        }
                                        if (extension == ".xls" || extension == ".xlsm" || extension == ".xlsx")
                                        {
                                            //agregar a mi lista de adjuntos
                                            adjuntos.Add(new AdjutoSolicitudParcial
                                            {
                                                tipoAdjunto = "Matriz",
                                                nombreAdjunto = path
                                            });
                                            matrizexcel = true;
                                        }
                                    }
                                }

                                //si tengo los dos archivos paso
                                if (zip == true && matrizexcel == true)
                                {
                                    guardarSolicitud = true;
                                }
                                else
                                {
                                    guardarSolicitud = false;
                                    return Json(new { Resultado = new RespuestaTransaccion { Estado = false, Respuesta = "Debe adjuntar solo 2 archivos, el de matriz y materiales" } }, JsonRequestBehavior.AllowGet);
                                }
                            }
                        }
                    }
                }

                if (guardarSolicitud == true)
                {
                    //Crear la solicitud con datos generales
                    if (urgente == "true")
                    {
                        solicitud.solicitud_urgente = true;
                    }
                    else
                    {
                        solicitud.solicitud_urgente = false;
                    }

                    RespuestaTransaccion resultado = SolicitudClienteEntity.CrearSolicitud(solicitud, idUsuario, camposPersonalizados, urlExterno, adjuntos, urlSoporte, portales);

                    //Validar los adjuntos del eclub o icare
                    if (files.Count > 0)
                    {
                        //anio Actual
                        var anioActual = System.DateTime.Now.Year.ToString();

                        string basePath = ConfigurationManager.AppSettings["RepositorioDocumentos"];
                        string rutaArchivos = basePath + "\\GESTION_PPM\\ADJUNTOS_SOLICITUDES\\" + anioActual + '\\' + resultado.SolicitudID.ToString();

                        // En caso de que no exista el directorio, crearlo.
                        bool directorio = Directory.Exists(rutaArchivos);
                        if (!directorio)
                            Directory.CreateDirectory(rutaArchivos);

                        //Recorrer todos los adjuntos
                        for (int i = 0; i < files.Count; i++)
                        {
                            HttpPostedFileBase file = files[i];
                            string path = string.Empty;
                            string FileName = "";

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

                            path = Path.Combine(rutaArchivos, path);

                            file.SaveAs(path);
                        }
                    }

                    //Mandar a generar el pdf si todo es ok
                    if (resultado.Estado == true)
                    {
                        //Si es una solicitud de tipo General
                        if (solicitud.id_tipo == valorCodigoGeneral)
                        {
                            GenerarPDFGeneral(resultado.SolicitudID);
                            return Json(new { Resultado = resultado }, JsonRequestBehavior.AllowGet);
                        }
                        else
                        {
                            //Si es una solicitud de tipo Icare o Eclub
                            if (solicitud.id_tipo == valorCodigoIcare || solicitud.id_tipo == valorCodigoEclub)
                            {
                                GenerarPDFEclubIcare(resultado.SolicitudID);
                                return Json(new { Resultado = resultado }, JsonRequestBehavior.AllowGet);
                            }
                            else
                            {
                                if (solicitud.id_tipo == valorCodigoMantenimiento)
                                {
                                    GenerarPDFMantenimiento(resultado.SolicitudID);
                                    return Json(new { Resultado = resultado }, JsonRequestBehavior.AllowGet);
                                }
                                else
                                {
                                    //Si es una solicitud de envios icare
                                    if (solicitud.id_tipo == valorcodigoEnviosIcare)
                                    {
                                        GenerarPDFEnvioIcare(resultado.SolicitudID);
                                        return Json(new { Resultado = resultado }, JsonRequestBehavior.AllowGet);
                                    }
                                    else
                                    {
                                        //Si es de html5
                                        if (solicitud.id_tipo == valorcodigoHtml)
                                        {
                                            GenerarPDFHtml5(resultado.SolicitudID);
                                            return Json(new { Resultado = resultado }, JsonRequestBehavior.AllowGet);
                                        }
                                        else
                                        {
                                            //Si es de tags
                                            if (solicitud.id_tipo == valorcodigoTags)
                                            {
                                                GenerarPDFTags(resultado.SolicitudID);
                                                return Json(new { Resultado = resultado }, JsonRequestBehavior.AllowGet);
                                            }
                                            else
                                            {
                                                //Si es de html5 y  tags
                                                if (solicitud.id_tipo == valorcodigoHtmlTags)
                                                {
                                                    GenerarPDFHtmlTags(resultado.SolicitudID);
                                                    return Json(new { Resultado = resultado }, JsonRequestBehavior.AllowGet);
                                                }
                                                else
                                                {
                                                    return Json(new { Resultado = resultado }, JsonRequestBehavior.AllowGet);
                                                }
                                            }

                                        }
                                    }

                                }
                            }
                        }
                    }
                    else
                    {
                        return Json(new { Resultado = resultado }, JsonRequestBehavior.AllowGet);
                    }
                }

                else
                {
                    return Json(new { Resultado = "Hay campos obligatorios" }, JsonRequestBehavior.AllowGet);
                }
            }
            catch (Exception ex)
            {
                return Json(new { Resultado = new RespuestaTransaccion { Estado = false, Respuesta = ex.Message } }, JsonRequestBehavior.AllowGet);
            }
        }

        // GET: CodigoCotizacion/Edit/5
        public ActionResult Edit(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            SolicitudCliente solicitud = SolicitudClienteEntity.ConsultarSolicitudCliente(id.Value);

            ViewBag.DatosUsuario = UsuarioEntity.ConsultarUsuario(solicitud.id_solicitante.Value);

            //Listado de Tipo 
            var Tipo = CatalogoEntity.ObtenerListadoCatalogosByCodigo("TSL-01");
            ViewBag.ListadoTipo = Tipo;

            //Listado Marca
            var Marca = CatalogoEntity.ObtenerListadoCatalogosByCodigo("MRC-01");
            ViewBag.ListadoMarca = Marca;

            //Listado Subtipo  
            var Subtipo = CatalogoEntity.ConsultarCatalogoPorPadre(solicitud.id_tipo.HasValue ? solicitud.id_tipo.Value : 0, "SUBTIPO");
            ViewBag.ListadoSubtipo = Subtipo;

            ViewBag.SublineaNegocioCodigoCotizacion = CodigoCotizacionEntity.ListarSublineaNegocioCodigoCotizacion(id.Value);
            ViewBag.TotalSublineaNegocioCodigoCotizacion = CodigoCotizacionEntity.ListarSublineaNegocioCodigoCotizacion(id.Value).Sum(s => s.Valor);

            if (solicitud == null)
            {
                return HttpNotFound();
            }
            return View(solicitud);
        }

        // POST: CodigoCotizacion/Edit/5
        // To protect from overposting attacks, please enable the specific properties you want to bind to, for 
        // more details see https://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        public ActionResult Edit(int? codigo, string archivo)
        {
            try
            {
                //Listado de Tipo 
                var Tipo = CatalogoEntity.ObtenerListadoCatalogosByCodigo("TSL-01");
                ViewBag.ListadoTipo = Tipo;

                //Listado Marca
                var Marca = CatalogoEntity.ObtenerListadoCatalogosByCodigo("MRC-01");
                ViewBag.ListadoMarca = Marca;

                //Obtener el codigo de usuario que se logeo
                var usuarioSesion = ViewData["usuario"] = System.Web.HttpContext.Current.Session["usuario"];
                var idUsuario = Convert.ToInt16(usuarioSesion);

                ViewBag.DatosUsuario = UsuarioEntity.ConsultarUsuario(idUsuario);

                //Obtener una inicializacion de solicitud
                SolicitudCliente solicitud = new SolicitudCliente();

                //Obtener el codigo de general
                var codigoGeneral = Tipo.Where(t => Convert.ToInt32(t.Value) == 758).FirstOrDefault().Value;
                solicitud.id_tipo = Convert.ToInt16(codigoGeneral);

                //Listado Subtipo  
                var Subtipo = CatalogoEntity.ConsultarCatalogoPorPadre(solicitud.id_tipo.HasValue ? solicitud.id_tipo.Value : 0, "SUBTIPO");
                ViewBag.ListadoSubtipo = Subtipo;

                return Json(new { Resultado = "" }, JsonRequestBehavior.AllowGet);

            }
            catch (Exception ex)
            {
                return Json(new { Resultado = new RespuestaTransaccion { Estado = false, Respuesta = ex.Message } }, JsonRequestBehavior.AllowGet);
            }
        }

        [HttpPost]
        public ActionResult EditarStatusCodigoCotizacion(CodigoCotizacion codigoCotizacion)
        {
            try
            {
                RespuestaTransaccion resultado = CodigoCotizacionEntity.ActualizarStatusCodigoCotizacion(codigoCotizacion);
                return Json(new { Resultado = resultado }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                return Json(new { Resultado = new RespuestaTransaccion { Estado = false, Respuesta = ex.Message } }, JsonRequestBehavior.AllowGet);
            }
        }

        public ActionResult GetSubtipoDependiente(int id)
        {
            var subtipo = CatalogoEntity.ConsultarCatalogoPorPadre(id, "SUBTIPO").ToList();
            return Json(subtipo);
        }

        #region Generacion PDF

        //Reporte Solicitud General
        public void GenerarPDFGeneral(int? id)
        {
            //Obtener los datos de la Solicitud
            DatosSolicitudCliente solicitud = db.DatosSolicitudCliente(id).First();

            //Fonts para el PDF
            //Regular
            string fontpath = Server.MapPath("~/Content/fonts/ubuntu/Raleway-Regular.ttf");
            BaseFont customfont = BaseFont.CreateFont(fontpath, BaseFont.CP1252, BaseFont.EMBEDDED);

            //Bold
            string fontpathbold = Server.MapPath("~/Content/fonts/ubuntu/Raleway-Bold.ttf");
            BaseFont customfontbold = BaseFont.CreateFont(fontpathbold, BaseFont.CP1252, BaseFont.EMBEDDED);

            //ruta imagen del formulario
            String FilePath = Server.MapPath("~/Content/img/LogoPPMPDF.png");
            var physicalPath = Server.MapPath("~/Content/img/LogoPPMPDF.png");

            //anio Actual
            var anioActual = System.DateTime.Now.Year.ToString();

            string basePath = ConfigurationManager.AppSettings["RepositorioDocumentos"];
            string rutaArchivos = basePath + "\\GESTION_PPM\\ADJUNTOS_SOLICITUDES\\" + anioActual + '\\' + solicitud.Id.ToString();

            // En caso de que no exista el directorio, crearlo.
            bool directorio = Directory.Exists(rutaArchivos);
            if (!directorio)
                Directory.CreateDirectory(rutaArchivos);

            //ruta del archivo PDF que se va a crear
            string rutaDocumentos = rutaArchivos + "\\Solicitud-" + solicitud.CodigoSolicitud.ToString() + ".pdf";

            //Crear el formulario PDF
            FileStream fs = new FileStream(rutaDocumentos, FileMode.Create);
            Document document = new Document(iTextSharp.text.PageSize.A4, -20, -20, 0, 0);

            //Document document = new Document(iTextSharp.text.PageSize.A4);
            PdfWriter pw = PdfWriter.GetInstance(document, fs);

            //Abrir archivo para su edicion
            document.Open();

            //Salto de linea 
            document.Add(new Paragraph(" "));

            //Salto de linea 
            document.Add(new Paragraph(" "));

            //Cabecera del Cotizador         
            PdfPTable cabecera = new PdfPTable(7);
            cabecera.DefaultCell.Border = Rectangle.NO_BORDER;

            //Imagen
            iTextSharp.text.Image imagen = iTextSharp.text.Image.GetInstance(FilePath);
            imagen.ScalePercent(10);

            PdfPCell logocell = new PdfPCell(imagen, true);
            logocell.BackgroundColor = new BaseColor(60, 66, 82);
            logocell.Colspan = 2;
            logocell.Border = Rectangle.NO_BORDER;
            logocell.HorizontalAlignment = Element.ALIGN_LEFT;
            logocell.VerticalAlignment = Element.ALIGN_MIDDLE;

            //Titulo
            PdfPCell TituloCotizador = new PdfPCell(new Phrase("SOLICITUD " + solicitud.Tipo.ToUpper() + " N° " + solicitud.CodigoSolicitud + "      ", new Font(customfontbold, 9, 0, new BaseColor(255, 255, 255))));
            TituloCotizador.Colspan = 5;
            TituloCotizador.Border = Rectangle.NO_BORDER;
            TituloCotizador.HorizontalAlignment = Element.ALIGN_RIGHT;
            TituloCotizador.VerticalAlignment = Element.ALIGN_MIDDLE;
            TituloCotizador.BackgroundColor = new BaseColor(60, 66, 82);
            TituloCotizador.FixedHeight = 50f;

            cabecera.AddCell(logocell);
            cabecera.AddCell(TituloCotizador);

            document.Add(cabecera);

            //Salto Linea      
            PdfPTable saltoLinea1 = new PdfPTable(4);

            PdfPCell EtiquetaSaltoLinea = new PdfPCell(new Phrase(" ", new Font(customfontbold, 9)));
            EtiquetaSaltoLinea.Colspan = 4;
            EtiquetaSaltoLinea.Border = Rectangle.NO_BORDER;
            EtiquetaSaltoLinea.HorizontalAlignment = Element.ALIGN_CENTER;
            EtiquetaSaltoLinea.VerticalAlignment = Element.ALIGN_MIDDLE;
            EtiquetaSaltoLinea.FixedHeight = 8f;

            saltoLinea1.AddCell(EtiquetaSaltoLinea);

            document.Add(saltoLinea1);

            //Informacion de la solicitud         
            PdfPTable informacion = new PdfPTable(6);

            PdfPCell EtiquetaFechaCotizacion = new PdfPCell(new Phrase("Fecha y Hora: ", new Font(customfontbold, 7)));
            EtiquetaFechaCotizacion.Colspan = 1;
            EtiquetaFechaCotizacion.HorizontalAlignment = Element.ALIGN_RIGHT;
            EtiquetaFechaCotizacion.VerticalAlignment = Element.ALIGN_MIDDLE;
            EtiquetaFechaCotizacion.FixedHeight = 16f;
            EtiquetaFechaCotizacion.BackgroundColor = new BaseColor(225, 225, 225);
            EtiquetaFechaCotizacion.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell ValorFechaCotizacion = new PdfPCell(new Phrase("   " + solicitud.FechaHora, new Font(customfont, 7)));
            ValorFechaCotizacion.Colspan = 2;
            ValorFechaCotizacion.HorizontalAlignment = Element.ALIGN_LEFT;
            ValorFechaCotizacion.VerticalAlignment = Element.ALIGN_MIDDLE;
            ValorFechaCotizacion.FixedHeight = 16f;
            ValorFechaCotizacion.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell EtiquetaSolicitante = new PdfPCell(new Phrase("Solicita:  ", new Font(customfontbold, 7)));
            EtiquetaSolicitante.Colspan = 1;
            EtiquetaSolicitante.HorizontalAlignment = Element.ALIGN_RIGHT;
            EtiquetaSolicitante.VerticalAlignment = Element.ALIGN_MIDDLE;
            EtiquetaSolicitante.FixedHeight = 16f;
            EtiquetaSolicitante.BackgroundColor = new BaseColor(225, 225, 225);
            EtiquetaSolicitante.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell ValorSolicitante = new PdfPCell(new Phrase("   " + solicitud.Solicitante, new Font(customfont, 7)));
            ValorSolicitante.Colspan = 2;
            ValorSolicitante.HorizontalAlignment = Element.ALIGN_LEFT;
            ValorSolicitante.VerticalAlignment = Element.ALIGN_MIDDLE;
            ValorSolicitante.FixedHeight = 16f;
            ValorSolicitante.BorderColor = new BaseColor(60, 66, 82);

            informacion.AddCell(EtiquetaFechaCotizacion);
            informacion.AddCell(ValorFechaCotizacion);
            informacion.AddCell(EtiquetaSolicitante);
            informacion.AddCell(ValorSolicitante);

            document.Add(informacion);

            //Salto Linea       
            PdfPTable saltoLinea2 = new PdfPTable(4);

            PdfPCell EtiquetaSaltoLinea2 = new PdfPCell(new Phrase(" ", new Font(customfontbold, 9)));
            EtiquetaSaltoLinea2.Colspan = 4;
            EtiquetaSaltoLinea2.Border = Rectangle.NO_BORDER;
            EtiquetaSaltoLinea2.HorizontalAlignment = Element.ALIGN_CENTER;
            EtiquetaSaltoLinea2.VerticalAlignment = Element.ALIGN_MIDDLE;
            EtiquetaSaltoLinea2.FixedHeight = 8f;

            saltoLinea2.AddCell(EtiquetaSaltoLinea2);

            document.Add(saltoLinea2);

            //Informacion de la Solicitud       
            PdfPTable datosGenerales = new PdfPTable(6);

            PdfPCell EtiquetaEmpresa = new PdfPCell(new Phrase("Empresa: ", new Font(customfontbold, 7)));
            EtiquetaEmpresa.Colspan = 1;
            EtiquetaEmpresa.HorizontalAlignment = Element.ALIGN_RIGHT;
            EtiquetaEmpresa.VerticalAlignment = Element.ALIGN_MIDDLE;
            EtiquetaEmpresa.FixedHeight = 16f;
            EtiquetaEmpresa.BackgroundColor = new BaseColor(225, 225, 225);
            EtiquetaEmpresa.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell ValorEmpresa = new PdfPCell(new Phrase("   " + solicitud.Empresa, new Font(customfont, 7)));
            ValorEmpresa.Colspan = 5;
            ValorEmpresa.HorizontalAlignment = Element.ALIGN_LEFT;
            ValorEmpresa.VerticalAlignment = Element.ALIGN_MIDDLE;
            ValorEmpresa.FixedHeight = 16f;
            ValorEmpresa.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell EtiquetaCargo = new PdfPCell(new Phrase("Cargo: ", new Font(customfontbold, 7)));
            EtiquetaCargo.Colspan = 1;
            EtiquetaCargo.HorizontalAlignment = Element.ALIGN_RIGHT;
            EtiquetaCargo.VerticalAlignment = Element.ALIGN_MIDDLE;
            EtiquetaCargo.FixedHeight = 16f;
            EtiquetaCargo.BackgroundColor = new BaseColor(225, 225, 225);
            EtiquetaCargo.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell ValorCargo = new PdfPCell(new Phrase("   " + solicitud.Cargo, new Font(customfont, 7)));
            ValorCargo.Colspan = 5;
            ValorCargo.HorizontalAlignment = Element.ALIGN_LEFT;
            ValorCargo.VerticalAlignment = Element.ALIGN_MIDDLE;
            ValorCargo.FixedHeight = 16f;
            ValorCargo.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell EtiquetaTelefono = new PdfPCell(new Phrase("Teléfono: ", new Font(customfontbold, 7)));
            EtiquetaTelefono.Colspan = 1;
            EtiquetaTelefono.HorizontalAlignment = Element.ALIGN_RIGHT;
            EtiquetaTelefono.VerticalAlignment = Element.ALIGN_MIDDLE;
            EtiquetaTelefono.FixedHeight = 16f;
            EtiquetaTelefono.BackgroundColor = new BaseColor(225, 225, 225);
            EtiquetaTelefono.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell ValorTelefono = new PdfPCell(new Phrase("   " + solicitud.Telefono, new Font(customfont, 7)));
            ValorTelefono.Colspan = 2;
            ValorTelefono.HorizontalAlignment = Element.ALIGN_LEFT;
            ValorTelefono.VerticalAlignment = Element.ALIGN_MIDDLE;
            ValorTelefono.FixedHeight = 16f;
            ValorTelefono.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell EtiquetaCelular = new PdfPCell(new Phrase("Celular: ", new Font(customfontbold, 7)));
            EtiquetaCelular.Colspan = 1;
            EtiquetaCelular.HorizontalAlignment = Element.ALIGN_RIGHT;
            EtiquetaCelular.VerticalAlignment = Element.ALIGN_MIDDLE;
            EtiquetaCelular.FixedHeight = 16f;
            EtiquetaCelular.BackgroundColor = new BaseColor(225, 225, 225);
            EtiquetaCelular.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell ValorCelular = new PdfPCell(new Phrase("   " + solicitud.Celular, new Font(customfont, 7)));
            ValorCelular.Colspan = 2;
            ValorCelular.HorizontalAlignment = Element.ALIGN_LEFT;
            ValorCelular.VerticalAlignment = Element.ALIGN_MIDDLE;
            ValorCelular.FixedHeight = 16f;
            ValorCelular.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell EtiquetaMail = new PdfPCell(new Phrase("Mail: ", new Font(customfontbold, 7)));
            EtiquetaMail.Colspan = 1;
            EtiquetaMail.HorizontalAlignment = Element.ALIGN_RIGHT;
            EtiquetaMail.VerticalAlignment = Element.ALIGN_MIDDLE;
            EtiquetaMail.FixedHeight = 16f;
            EtiquetaMail.BackgroundColor = new BaseColor(225, 225, 225);
            EtiquetaMail.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell ValorMail = new PdfPCell(new Phrase("   " + solicitud.Mail, new Font(customfont, 7)));
            ValorMail.Colspan = 5;
            ValorMail.HorizontalAlignment = Element.ALIGN_LEFT;
            ValorMail.VerticalAlignment = Element.ALIGN_MIDDLE;
            ValorMail.FixedHeight = 16f;
            ValorMail.BorderColor = new BaseColor(60, 66, 82);

            datosGenerales.AddCell(EtiquetaEmpresa);
            datosGenerales.AddCell(ValorEmpresa);

            datosGenerales.AddCell(EtiquetaCargo);
            datosGenerales.AddCell(ValorCargo);

            datosGenerales.AddCell(EtiquetaTelefono);
            datosGenerales.AddCell(ValorTelefono);
            datosGenerales.AddCell(EtiquetaCelular);
            datosGenerales.AddCell(ValorCelular);

            datosGenerales.AddCell(EtiquetaMail);
            datosGenerales.AddCell(ValorMail);

            document.Add(datosGenerales);

            //Salto Linea       
            PdfPTable saltoLinea3 = new PdfPTable(4);

            PdfPCell EtiquetaSaltoLinea3 = new PdfPCell(new Phrase(" ", new Font(customfontbold, 9)));
            EtiquetaSaltoLinea3.Colspan = 4;
            EtiquetaSaltoLinea3.Border = Rectangle.NO_BORDER;
            EtiquetaSaltoLinea3.HorizontalAlignment = Element.ALIGN_CENTER;
            EtiquetaSaltoLinea3.VerticalAlignment = Element.ALIGN_MIDDLE;
            EtiquetaSaltoLinea3.FixedHeight = 8f;

            saltoLinea3.AddCell(EtiquetaSaltoLinea3);

            document.Add(saltoLinea3);

            //Descripcion del requermiento    
            PdfPTable entregables = new PdfPTable(6);

            PdfPCell EtiquetaDescripcion = new PdfPCell(new Phrase("DESCRIPCIÓN DEL REQUERIMIENTO", new Font(customfontbold, 9, 0, new BaseColor(255, 255, 255))));
            EtiquetaDescripcion.Colspan = 6;
            EtiquetaDescripcion.HorizontalAlignment = Element.ALIGN_CENTER;
            EtiquetaDescripcion.VerticalAlignment = Element.ALIGN_MIDDLE;
            EtiquetaDescripcion.FixedHeight = 16f;
            EtiquetaDescripcion.BackgroundColor = new BaseColor(60, 66, 82);
            EtiquetaDescripcion.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell ValorDescripcion = new PdfPCell(new Phrase(solicitud.Descripcion, new Font(customfont, 7)));
            ValorDescripcion.Colspan = 6;
            ValorDescripcion.Rowspan = 6;
            ValorDescripcion.VerticalAlignment = Element.ALIGN_MIDDLE;
            ValorDescripcion.BorderColor = new BaseColor(60, 66, 82);

            //validar el tamaño de las celdas
            var caracteres = solicitud.Descripcion.Count();
            if (caracteres <= 70)
            {
                ValorDescripcion.FixedHeight = 16f;
            }
            else if (caracteres <= 140)
            {
                ValorDescripcion.FixedHeight = 32f;
            }
            else if (caracteres <= 210)
            {
                ValorDescripcion.FixedHeight = 46f;
            }
            else if (caracteres <= 280)
            {
                ValorDescripcion.FixedHeight = 58f;
            }
            else if (caracteres <= 350)
            {
                ValorDescripcion.FixedHeight = 70f;
            }
            else if (caracteres <= 420)
            {
                ValorDescripcion.FixedHeight = 82f;
            }
            else if (caracteres <= 550)
            {
                ValorDescripcion.FixedHeight = 94f;
            }
            else if (caracteres <= 680)
            {
                ValorDescripcion.FixedHeight = 106f;
            }
            else if (caracteres <= 810)
            {
                ValorDescripcion.FixedHeight = 128f;
            }
            else
            {
                ValorDescripcion.FixedHeight = 140f;
            }

            entregables.AddCell(EtiquetaDescripcion);
            entregables.AddCell(ValorDescripcion);

            document.Add(entregables);
            document.Close();
        }

        //Reporte Solicitud Eclub o Icare
        public void GenerarPDFEclubIcare(int? id)
        {
            int camposPersonalizados = 1;
            int contadorAdjuntos = 1;

            //Obtener los datos de la Solicitud
            DatosSolicitudCliente solicitud = db.DatosSolicitudCliente(id).First();

            //Campos Personalizados
            List<CamposPersonalizadosSolicitud> ListadoCamposPersonalizados = db.CamposPersonalizadosSolicitud(id).ToList();

            //Url Externo 
            List<UrlExternosSolicitud> ListadoUrlExterno = db.UrlExternosSolicitud(id).ToList();

            //Adjuntos
            List<AdjuntosSolicitud> ListadoAdjuntos = db.AdjuntosSolicitud(id).ToList();

            //Fonts para el PDF
            //Regular
            string fontpath = Server.MapPath("~/Content/fonts/ubuntu/Raleway-Regular.ttf");
            BaseFont customfont = BaseFont.CreateFont(fontpath, BaseFont.CP1252, BaseFont.EMBEDDED);

            //Bold
            string fontpathbold = Server.MapPath("~/Content/fonts/ubuntu/Raleway-Bold.ttf");
            BaseFont customfontbold = BaseFont.CreateFont(fontpathbold, BaseFont.CP1252, BaseFont.EMBEDDED);

            //ruta imagen del formulario
            String FilePath = Server.MapPath("~/Content/img/LogoPPMPDF.png");
            var physicalPath = Server.MapPath("~/Content/img/LogoPPMPDF.png");

            //anio Actual
            var anioActual = System.DateTime.Now.Year.ToString();

            string basePath = ConfigurationManager.AppSettings["RepositorioDocumentos"];
            string rutaArchivos = basePath + "\\GESTION_PPM\\ADJUNTOS_SOLICITUDES\\" + anioActual + '\\' + solicitud.Id.ToString();

            // En caso de que no exista el directorio, crearlo.
            bool directorio = Directory.Exists(rutaArchivos);
            if (!directorio)
                Directory.CreateDirectory(rutaArchivos);

            //ruta del archivo PDF que se va a crear
            string rutaDocumentos = rutaArchivos + "\\Solicitud-" + solicitud.CodigoSolicitud.ToString() + ".pdf";

            //Crear el formulario PDF
            FileStream fs = new FileStream(rutaDocumentos, FileMode.Create);
            Document document = new Document(iTextSharp.text.PageSize.A4, -20, -20, 0, 0);

            //Document document = new Document(iTextSharp.text.PageSize.A4);
            PdfWriter pw = PdfWriter.GetInstance(document, fs);

            //Abrir archivo para su edicion
            document.Open();

            //Salto de linea 
            document.Add(new Paragraph(" "));

            //Salto de linea 
            document.Add(new Paragraph(" "));

            //Cabecera del Cotizador         
            PdfPTable cabecera = new PdfPTable(7);
            cabecera.DefaultCell.Border = Rectangle.NO_BORDER;

            //Imagen
            iTextSharp.text.Image imagen = iTextSharp.text.Image.GetInstance(FilePath);
            imagen.ScalePercent(10);

            PdfPCell logocell = new PdfPCell(imagen, true);
            logocell.BackgroundColor = new BaseColor(60, 66, 82);
            logocell.Colspan = 2;
            logocell.Border = Rectangle.NO_BORDER;
            logocell.HorizontalAlignment = Element.ALIGN_LEFT;
            logocell.VerticalAlignment = Element.ALIGN_MIDDLE;

            //Titulo
            PdfPCell TituloCotizador = new PdfPCell(new Phrase("SOLICITUD " + solicitud.Tipo.ToUpper() + " N° " + solicitud.CodigoSolicitud + "      ", new Font(customfontbold, 9, 0, new BaseColor(255, 255, 255))));
            TituloCotizador.Colspan = 5;
            TituloCotizador.Border = Rectangle.NO_BORDER;
            TituloCotizador.HorizontalAlignment = Element.ALIGN_RIGHT;
            TituloCotizador.VerticalAlignment = Element.ALIGN_MIDDLE;
            TituloCotizador.BackgroundColor = new BaseColor(60, 66, 82);
            TituloCotizador.FixedHeight = 50f;

            cabecera.AddCell(logocell);
            cabecera.AddCell(TituloCotizador);

            document.Add(cabecera);

            //Salto Linea      
            PdfPTable saltoLinea1 = new PdfPTable(4);

            PdfPCell EtiquetaSaltoLinea = new PdfPCell(new Phrase(" ", new Font(customfontbold, 9)));
            EtiquetaSaltoLinea.Colspan = 4;
            EtiquetaSaltoLinea.Border = Rectangle.NO_BORDER;
            EtiquetaSaltoLinea.HorizontalAlignment = Element.ALIGN_CENTER;
            EtiquetaSaltoLinea.VerticalAlignment = Element.ALIGN_MIDDLE;
            EtiquetaSaltoLinea.FixedHeight = 8f;

            saltoLinea1.AddCell(EtiquetaSaltoLinea);

            document.Add(saltoLinea1);

            //Informacion de la solicitud         
            PdfPTable informacion = new PdfPTable(6);

            PdfPCell EtiquetaFechaCotizacion = new PdfPCell(new Phrase("Fecha y Hora: ", new Font(customfontbold, 7)));
            EtiquetaFechaCotizacion.Colspan = 1;
            EtiquetaFechaCotizacion.HorizontalAlignment = Element.ALIGN_RIGHT;
            EtiquetaFechaCotizacion.VerticalAlignment = Element.ALIGN_MIDDLE;
            EtiquetaFechaCotizacion.FixedHeight = 16f;
            EtiquetaFechaCotizacion.BackgroundColor = new BaseColor(225, 225, 225);
            EtiquetaFechaCotizacion.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell ValorFechaCotizacion = new PdfPCell(new Phrase("   " + solicitud.FechaHora, new Font(customfont, 7)));
            ValorFechaCotizacion.Colspan = 2;
            ValorFechaCotizacion.HorizontalAlignment = Element.ALIGN_LEFT;
            ValorFechaCotizacion.VerticalAlignment = Element.ALIGN_MIDDLE;
            ValorFechaCotizacion.FixedHeight = 16f;
            ValorFechaCotizacion.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell EtiquetaSolicitante = new PdfPCell(new Phrase("Solicitante:  ", new Font(customfontbold, 7)));
            EtiquetaSolicitante.Colspan = 1;
            EtiquetaSolicitante.HorizontalAlignment = Element.ALIGN_RIGHT;
            EtiquetaSolicitante.VerticalAlignment = Element.ALIGN_MIDDLE;
            EtiquetaSolicitante.FixedHeight = 16f;
            EtiquetaSolicitante.BackgroundColor = new BaseColor(225, 225, 225);
            EtiquetaSolicitante.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell ValorSolicitante = new PdfPCell(new Phrase("   " + solicitud.Solicitante, new Font(customfont, 7)));
            ValorSolicitante.Colspan = 2;
            ValorSolicitante.HorizontalAlignment = Element.ALIGN_LEFT;
            ValorSolicitante.VerticalAlignment = Element.ALIGN_MIDDLE;
            ValorSolicitante.FixedHeight = 16f;
            ValorSolicitante.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell EtiquetaArea1 = new PdfPCell(new Phrase("Solicitante DCE", new Font(customfontbold, 7)));
            EtiquetaArea1.Colspan = 1;
            EtiquetaArea1.HorizontalAlignment = Element.ALIGN_RIGHT;
            EtiquetaArea1.VerticalAlignment = Element.ALIGN_MIDDLE;
            EtiquetaArea1.FixedHeight = 16f;
            EtiquetaArea1.BackgroundColor = new BaseColor(225, 225, 225);
            EtiquetaArea1.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell ValorArea1 = new PdfPCell(new Phrase("   " + solicitud.Area1, new Font(customfont, 7)));
            ValorArea1.Colspan = 2;
            ValorArea1.HorizontalAlignment = Element.ALIGN_LEFT;
            ValorArea1.VerticalAlignment = Element.ALIGN_MIDDLE;
            ValorArea1.FixedHeight = 16f;
            ValorArea1.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell EtiquetaArea2 = new PdfPCell(new Phrase("Área Solicitante DCE:  ", new Font(customfontbold, 7)));
            EtiquetaArea2.Colspan = 1;
            EtiquetaArea2.HorizontalAlignment = Element.ALIGN_RIGHT;
            EtiquetaArea2.VerticalAlignment = Element.ALIGN_MIDDLE;
            EtiquetaArea2.FixedHeight = 16f;
            EtiquetaArea2.BackgroundColor = new BaseColor(225, 225, 225);
            EtiquetaArea2.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell ValorArea2 = new PdfPCell(new Phrase("   " + solicitud.Area2, new Font(customfont, 7)));
            ValorArea2.Colspan = 2;
            ValorArea2.HorizontalAlignment = Element.ALIGN_LEFT;
            ValorArea2.VerticalAlignment = Element.ALIGN_MIDDLE;
            ValorArea2.FixedHeight = 16f;
            ValorArea2.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell EtiquetaImplementacion = new PdfPCell(new Phrase("Preheader:  ", new Font(customfontbold, 7)));
            EtiquetaImplementacion.Colspan = 1;
            EtiquetaImplementacion.HorizontalAlignment = Element.ALIGN_RIGHT;
            EtiquetaImplementacion.VerticalAlignment = Element.ALIGN_MIDDLE;
            EtiquetaImplementacion.FixedHeight = 16f;
            EtiquetaImplementacion.BackgroundColor = new BaseColor(225, 225, 225);
            EtiquetaImplementacion.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell ValorImplementacion = new PdfPCell(new Phrase("   " + solicitud.Implementacion, new Font(customfont, 7)));
            ValorImplementacion.Colspan = 5;
            ValorImplementacion.HorizontalAlignment = Element.ALIGN_LEFT;
            ValorImplementacion.VerticalAlignment = Element.ALIGN_MIDDLE;
            ValorImplementacion.FixedHeight = 16f;
            ValorImplementacion.BorderColor = new BaseColor(60, 66, 82);

            informacion.AddCell(EtiquetaFechaCotizacion);
            informacion.AddCell(ValorFechaCotizacion);
            informacion.AddCell(EtiquetaSolicitante);
            informacion.AddCell(ValorSolicitante);

            informacion.AddCell(EtiquetaArea1);
            informacion.AddCell(ValorArea1);
            informacion.AddCell(EtiquetaArea2);
            informacion.AddCell(ValorArea2);

            informacion.AddCell(EtiquetaImplementacion);
            informacion.AddCell(ValorImplementacion);

            document.Add(informacion);

            //Salto Linea       
            PdfPTable saltoLinea2 = new PdfPTable(4);

            PdfPCell EtiquetaSaltoLinea2 = new PdfPCell(new Phrase(" ", new Font(customfontbold, 9)));
            EtiquetaSaltoLinea2.Colspan = 4;
            EtiquetaSaltoLinea2.Border = Rectangle.NO_BORDER;
            EtiquetaSaltoLinea2.HorizontalAlignment = Element.ALIGN_CENTER;
            EtiquetaSaltoLinea2.VerticalAlignment = Element.ALIGN_MIDDLE;
            EtiquetaSaltoLinea2.FixedHeight = 8f;

            saltoLinea2.AddCell(EtiquetaSaltoLinea2);

            document.Add(saltoLinea2);

            //Informacion de la Solicitud       
            PdfPTable datosGenerales = new PdfPTable(6);

            PdfPCell EtiquetaTipo = new PdfPCell(new Phrase("Tipo: ", new Font(customfontbold, 7)));
            EtiquetaTipo.Colspan = 1;
            EtiquetaTipo.HorizontalAlignment = Element.ALIGN_RIGHT;
            EtiquetaTipo.VerticalAlignment = Element.ALIGN_MIDDLE;
            EtiquetaTipo.FixedHeight = 16f;
            EtiquetaTipo.BackgroundColor = new BaseColor(225, 225, 225);
            EtiquetaTipo.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell ValorTipo = new PdfPCell(new Phrase("   " + solicitud.Tipo, new Font(customfont, 7)));
            ValorTipo.Colspan = 5;
            ValorTipo.HorizontalAlignment = Element.ALIGN_LEFT;
            ValorTipo.VerticalAlignment = Element.ALIGN_MIDDLE;
            ValorTipo.FixedHeight = 16f;
            ValorTipo.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell EtiquetaSubtipo = new PdfPCell(new Phrase("Subtipo: ", new Font(customfontbold, 7)));
            EtiquetaSubtipo.Colspan = 1;
            EtiquetaSubtipo.HorizontalAlignment = Element.ALIGN_RIGHT;
            EtiquetaSubtipo.VerticalAlignment = Element.ALIGN_MIDDLE;
            EtiquetaSubtipo.FixedHeight = 16f;
            EtiquetaSubtipo.BackgroundColor = new BaseColor(225, 225, 225);
            EtiquetaSubtipo.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell ValorSubtipo = new PdfPCell(new Phrase("   " + solicitud.Subtipo, new Font(customfont, 7)));
            ValorSubtipo.Colspan = 5;
            ValorSubtipo.HorizontalAlignment = Element.ALIGN_LEFT;
            ValorSubtipo.VerticalAlignment = Element.ALIGN_MIDDLE;
            ValorSubtipo.FixedHeight = 16f;
            ValorSubtipo.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell EtiquetaMarca = new PdfPCell(new Phrase("Marca: ", new Font(customfontbold, 7)));
            EtiquetaMarca.Colspan = 1;
            EtiquetaMarca.HorizontalAlignment = Element.ALIGN_RIGHT;
            EtiquetaMarca.VerticalAlignment = Element.ALIGN_MIDDLE;
            EtiquetaMarca.FixedHeight = 16f;
            EtiquetaMarca.BackgroundColor = new BaseColor(225, 225, 225);
            EtiquetaMarca.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell ValorMarca = new PdfPCell(new Phrase("   " + solicitud.Marca, new Font(customfont, 7)));
            ValorMarca.Colspan = 2;
            ValorMarca.HorizontalAlignment = Element.ALIGN_LEFT;
            ValorMarca.VerticalAlignment = Element.ALIGN_MIDDLE;
            ValorMarca.FixedHeight = 16f;
            ValorMarca.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell EtiquetaCantidad = new PdfPCell(new Phrase("Cantidad: ", new Font(customfontbold, 7)));
            EtiquetaCantidad.Colspan = 1;
            EtiquetaCantidad.HorizontalAlignment = Element.ALIGN_RIGHT;
            EtiquetaCantidad.VerticalAlignment = Element.ALIGN_MIDDLE;
            EtiquetaCantidad.FixedHeight = 16f;
            EtiquetaCantidad.BackgroundColor = new BaseColor(225, 225, 225);
            EtiquetaCantidad.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell ValorCantidad = new PdfPCell(new Phrase("   " + solicitud.Cantidad, new Font(customfont, 7)));
            ValorCantidad.Colspan = 2;
            ValorCantidad.HorizontalAlignment = Element.ALIGN_LEFT;
            ValorCantidad.VerticalAlignment = Element.ALIGN_MIDDLE;
            ValorCantidad.FixedHeight = 16f;
            ValorCantidad.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell EtiquetaOP = new PdfPCell(new Phrase("OP: ", new Font(customfontbold, 7)));
            EtiquetaOP.Colspan = 1;
            EtiquetaOP.HorizontalAlignment = Element.ALIGN_RIGHT;
            EtiquetaOP.VerticalAlignment = Element.ALIGN_MIDDLE;
            EtiquetaOP.FixedHeight = 16f;
            EtiquetaOP.BackgroundColor = new BaseColor(225, 225, 225);
            EtiquetaOP.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell ValorOP = new PdfPCell(new Phrase("   " + solicitud.OP, new Font(customfont, 7)));
            ValorOP.Colspan = 2;
            ValorOP.HorizontalAlignment = Element.ALIGN_LEFT;
            ValorOP.VerticalAlignment = Element.ALIGN_MIDDLE;
            ValorOP.FixedHeight = 16f;
            ValorOP.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell EtiquetaMKT = new PdfPCell(new Phrase("MKT: ", new Font(customfontbold, 7)));
            EtiquetaMKT.Colspan = 1;
            EtiquetaMKT.HorizontalAlignment = Element.ALIGN_RIGHT;
            EtiquetaMKT.VerticalAlignment = Element.ALIGN_MIDDLE;
            EtiquetaMKT.FixedHeight = 16f;
            EtiquetaMKT.BackgroundColor = new BaseColor(225, 225, 225);
            EtiquetaMKT.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell ValorMKT = new PdfPCell(new Phrase("   " + solicitud.MKT, new Font(customfont, 7)));
            ValorMKT.Colspan = 2;
            ValorMKT.HorizontalAlignment = Element.ALIGN_LEFT;
            ValorMKT.VerticalAlignment = Element.ALIGN_MIDDLE;
            ValorMKT.FixedHeight = 16f;
            ValorMKT.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell EtiquetaNombre = new PdfPCell(new Phrase("Nombre Campaña: ", new Font(customfontbold, 7)));
            EtiquetaNombre.Colspan = 1;
            EtiquetaNombre.HorizontalAlignment = Element.ALIGN_RIGHT;
            EtiquetaNombre.VerticalAlignment = Element.ALIGN_MIDDLE;
            EtiquetaNombre.FixedHeight = 16f;
            EtiquetaNombre.BackgroundColor = new BaseColor(225, 225, 225);
            EtiquetaNombre.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell ValorNombre = new PdfPCell(new Phrase("   " + solicitud.Nombre, new Font(customfont, 7)));
            ValorNombre.Colspan = 5;
            ValorNombre.HorizontalAlignment = Element.ALIGN_LEFT;
            ValorNombre.VerticalAlignment = Element.ALIGN_MIDDLE;
            ValorNombre.FixedHeight = 16f;
            ValorNombre.BorderColor = new BaseColor(60, 66, 82);

            datosGenerales.AddCell(EtiquetaTipo);
            datosGenerales.AddCell(ValorTipo);

            datosGenerales.AddCell(EtiquetaSubtipo);
            datosGenerales.AddCell(ValorSubtipo);

            datosGenerales.AddCell(EtiquetaMarca);
            datosGenerales.AddCell(ValorMarca);

            datosGenerales.AddCell(EtiquetaCantidad);
            datosGenerales.AddCell(ValorCantidad);

            datosGenerales.AddCell(EtiquetaOP);
            datosGenerales.AddCell(ValorOP);
            datosGenerales.AddCell(EtiquetaMKT);
            datosGenerales.AddCell(ValorMKT);

            datosGenerales.AddCell(EtiquetaNombre);
            datosGenerales.AddCell(ValorNombre);

            document.Add(datosGenerales);

            //Salto Linea       
            PdfPTable saltoLinea3 = new PdfPTable(4);

            PdfPCell EtiquetaSaltoLinea3 = new PdfPCell(new Phrase(" ", new Font(customfontbold, 9)));
            EtiquetaSaltoLinea3.Colspan = 4;
            EtiquetaSaltoLinea3.Border = Rectangle.NO_BORDER;
            EtiquetaSaltoLinea3.HorizontalAlignment = Element.ALIGN_CENTER;
            EtiquetaSaltoLinea3.VerticalAlignment = Element.ALIGN_MIDDLE;
            EtiquetaSaltoLinea3.FixedHeight = 8f;

            saltoLinea3.AddCell(EtiquetaSaltoLinea3);

            document.Add(saltoLinea3);

            //Campos Personalizados
            PdfPTable camposDetallePersonalizados = new PdfPTable(6);

            PdfPCell EtiquetaCamposPersonalizados = new PdfPCell(new Phrase("CAMPOS PERSONALIZABLES", new Font(customfontbold, 9, 0, new BaseColor(255, 255, 255))));
            EtiquetaCamposPersonalizados.Colspan = 6;
            EtiquetaCamposPersonalizados.HorizontalAlignment = Element.ALIGN_CENTER;
            EtiquetaCamposPersonalizados.VerticalAlignment = Element.ALIGN_MIDDLE;
            EtiquetaCamposPersonalizados.FixedHeight = 16f;
            EtiquetaCamposPersonalizados.BackgroundColor = new BaseColor(60, 66, 82);
            EtiquetaCamposPersonalizados.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell EtiquetaCampo = new PdfPCell(new Phrase("Campo", new Font(customfontbold, 7)));
            EtiquetaCampo.Colspan = 1;
            EtiquetaCampo.HorizontalAlignment = Element.ALIGN_CENTER;
            EtiquetaCampo.VerticalAlignment = Element.ALIGN_MIDDLE;
            EtiquetaCampo.FixedHeight = 16f;
            EtiquetaCampo.BackgroundColor = new BaseColor(225, 225, 225);
            EtiquetaCampo.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell EtiquetaNombreCampo = new PdfPCell(new Phrase("Nombre", new Font(customfontbold, 7)));
            EtiquetaNombreCampo.Colspan = 5;
            EtiquetaNombreCampo.HorizontalAlignment = Element.ALIGN_LEFT;
            EtiquetaNombreCampo.VerticalAlignment = Element.ALIGN_MIDDLE;
            EtiquetaNombreCampo.FixedHeight = 16f;
            EtiquetaNombreCampo.BackgroundColor = new BaseColor(225, 225, 225);
            EtiquetaNombreCampo.BorderColor = new BaseColor(60, 66, 82);

            camposDetallePersonalizados.AddCell(EtiquetaCamposPersonalizados);

            camposDetallePersonalizados.AddCell(EtiquetaCampo);
            camposDetallePersonalizados.AddCell(EtiquetaNombreCampo);

            foreach (CamposPersonalizadosSolicitud campo in ListadoCamposPersonalizados)
            {
                PdfPCell EtiquetaCampoDetalle = new PdfPCell(new Phrase(camposPersonalizados.ToString(), new Font(customfont, 7)));
                EtiquetaCampoDetalle.Colspan = 1;
                EtiquetaCampoDetalle.HorizontalAlignment = Element.ALIGN_CENTER;
                EtiquetaCampoDetalle.VerticalAlignment = Element.ALIGN_MIDDLE;
                EtiquetaCampoDetalle.FixedHeight = 23f;
                EtiquetaCampoDetalle.BorderColor = new BaseColor(60, 66, 82);

                PdfPCell ValorDetalle = new PdfPCell(new Phrase(campo.nombre_campo, new Font(customfont, 7)));
                ValorDetalle.Colspan = 5;
                ValorDetalle.HorizontalAlignment = Element.ALIGN_LEFT;
                ValorDetalle.VerticalAlignment = Element.ALIGN_MIDDLE;
                ValorDetalle.FixedHeight = 23f;
                ValorDetalle.BorderColor = new BaseColor(60, 66, 82);

                camposDetallePersonalizados.AddCell(EtiquetaCampoDetalle);
                camposDetallePersonalizados.AddCell(ValorDetalle);

                camposPersonalizados = camposPersonalizados + 1;
            }

            if (ListadoCamposPersonalizados.Count > 0)
            {
                document.Add(camposDetallePersonalizados);

                //Salto Linea       
                PdfPTable saltoLinea4 = new PdfPTable(4);

                PdfPCell EtiquetaSaltoLinea4 = new PdfPCell(new Phrase(" ", new Font(customfontbold, 9)));
                EtiquetaSaltoLinea4.Colspan = 4;
                EtiquetaSaltoLinea4.Border = Rectangle.NO_BORDER;
                EtiquetaSaltoLinea4.HorizontalAlignment = Element.ALIGN_CENTER;
                EtiquetaSaltoLinea4.VerticalAlignment = Element.ALIGN_MIDDLE;
                EtiquetaSaltoLinea4.FixedHeight = 8f;

                saltoLinea4.AddCell(EtiquetaSaltoLinea4);

                document.Add(saltoLinea4);
            }

            //Url Externo  
            PdfPTable urlsExterno = new PdfPTable(7);

            PdfPCell EtiquetaCampoUrls = new PdfPCell(new Phrase("URL's Externos", new Font(customfontbold, 9, 0, new BaseColor(255, 255, 255))));
            EtiquetaCampoUrls.Colspan = 7;
            EtiquetaCampoUrls.HorizontalAlignment = Element.ALIGN_CENTER;
            EtiquetaCampoUrls.VerticalAlignment = Element.ALIGN_MIDDLE;
            EtiquetaCampoUrls.FixedHeight = 16f;
            EtiquetaCampoUrls.BackgroundColor = new BaseColor(60, 66, 82);
            EtiquetaCampoUrls.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell EtiquetaTipoUrl = new PdfPCell(new Phrase("Tipo", new Font(customfontbold, 7)));
            EtiquetaTipoUrl.Colspan = 2;
            EtiquetaTipoUrl.HorizontalAlignment = Element.ALIGN_CENTER;
            EtiquetaTipoUrl.VerticalAlignment = Element.ALIGN_MIDDLE;
            EtiquetaTipoUrl.FixedHeight = 16f;
            EtiquetaTipoUrl.BackgroundColor = new BaseColor(225, 225, 225);
            EtiquetaTipoUrl.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell EtiquetaDetalleUrl = new PdfPCell(new Phrase("Detalle", new Font(customfontbold, 7)));
            EtiquetaDetalleUrl.Colspan = 2;
            EtiquetaDetalleUrl.HorizontalAlignment = Element.ALIGN_CENTER;
            EtiquetaDetalleUrl.VerticalAlignment = Element.ALIGN_MIDDLE;
            EtiquetaDetalleUrl.FixedHeight = 16f;
            EtiquetaDetalleUrl.BackgroundColor = new BaseColor(225, 225, 225);
            EtiquetaDetalleUrl.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell EtiquetaNumeroTelefono = new PdfPCell(new Phrase("N. Teléfono", new Font(customfontbold, 7)));
            EtiquetaNumeroTelefono.Colspan = 1;
            EtiquetaNumeroTelefono.HorizontalAlignment = Element.ALIGN_CENTER;
            EtiquetaNumeroTelefono.VerticalAlignment = Element.ALIGN_MIDDLE;
            EtiquetaNumeroTelefono.FixedHeight = 16f;
            EtiquetaNumeroTelefono.BackgroundColor = new BaseColor(225, 225, 225);
            EtiquetaNumeroTelefono.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell EtiquetaUrlLink = new PdfPCell(new Phrase("Url", new Font(customfontbold, 7)));
            EtiquetaUrlLink.Colspan = 3;
            EtiquetaUrlLink.HorizontalAlignment = Element.ALIGN_CENTER;
            EtiquetaUrlLink.VerticalAlignment = Element.ALIGN_MIDDLE;
            EtiquetaUrlLink.FixedHeight = 16f;
            EtiquetaUrlLink.BackgroundColor = new BaseColor(225, 225, 225);
            EtiquetaUrlLink.BorderColor = new BaseColor(60, 66, 82);

            urlsExterno.AddCell(EtiquetaCampoUrls);

            urlsExterno.AddCell(EtiquetaTipoUrl);
            urlsExterno.AddCell(EtiquetaDetalleUrl);
            urlsExterno.AddCell(EtiquetaNumeroTelefono);
            urlsExterno.AddCell(EtiquetaUrlLink);

            PdfPTable urlsExternoLinkCompleto = new PdfPTable(9);

            PdfPCell EtiquetaCampoUrlsLinkCompleto = new PdfPCell(new Phrase("URL's Externos", new Font(customfontbold, 9, 0, new BaseColor(255, 255, 255))));
            EtiquetaCampoUrlsLinkCompleto.Colspan = 9;
            EtiquetaCampoUrlsLinkCompleto.HorizontalAlignment = Element.ALIGN_CENTER;
            EtiquetaCampoUrlsLinkCompleto.VerticalAlignment = Element.ALIGN_MIDDLE;
            EtiquetaCampoUrlsLinkCompleto.FixedHeight = 16f;
            EtiquetaCampoUrlsLinkCompleto.BackgroundColor = new BaseColor(60, 66, 82);
            EtiquetaCampoUrlsLinkCompleto.BorderColor = new BaseColor(60, 66, 82);

            urlsExternoLinkCompleto.AddCell(EtiquetaCampoUrlsLinkCompleto);

            foreach (UrlExternosSolicitud urls in ListadoUrlExterno)
            {
                PdfPCell ValorTipoUrlDetalle = new PdfPCell(new Phrase(urls.tipo.ToString(), new Font(customfont, 7)));
                ValorTipoUrlDetalle.Colspan = 2;
                ValorTipoUrlDetalle.HorizontalAlignment = Element.ALIGN_CENTER;
                ValorTipoUrlDetalle.VerticalAlignment = Element.ALIGN_MIDDLE;
                //ValorTipoUrlDetalle.FixedHeight = 23f;
                ValorTipoUrlDetalle.BorderColor = new BaseColor(60, 66, 82);

                PdfPCell ValorDetalleUrlDetalle = new PdfPCell(new Phrase(urls.detalle.ToString(), new Font(customfont, 7)));
                ValorDetalleUrlDetalle.Colspan = 2;
                ValorDetalleUrlDetalle.HorizontalAlignment = Element.ALIGN_CENTER;
                ValorDetalleUrlDetalle.VerticalAlignment = Element.ALIGN_MIDDLE;
                //ValorDetalleUrlDetalle.FixedHeight = 23f;
                ValorDetalleUrlDetalle.BorderColor = new BaseColor(60, 66, 82);
                 
                //NUEVO CAMPO NÚMERO DE TELÉFONO
                PdfPCell ValorNumeroTelefono = new PdfPCell(new Phrase((urls.NumeroTelefono ?? string.Empty).ToString(), new Font(customfont, 7)));
                ValorNumeroTelefono.Colspan = 1;
                ValorNumeroTelefono.HorizontalAlignment = Element.ALIGN_CENTER;
                ValorNumeroTelefono.VerticalAlignment = Element.ALIGN_MIDDLE;
                ValorNumeroTelefono.BorderColor = new BaseColor(60, 66, 82);
                 
                PdfPCell ValorUrlLinkDetalle = new PdfPCell(new Phrase(urls.url.ToString(), new Font(customfont, 7)));
                ValorUrlLinkDetalle.Colspan = 3;
                ValorUrlLinkDetalle.HorizontalAlignment = Element.ALIGN_CENTER;
                ValorUrlLinkDetalle.VerticalAlignment = Element.ALIGN_MIDDLE;
                ValorUrlLinkDetalle.HasFixedHeight();
                ValorUrlLinkDetalle.BorderColor = new BaseColor(60, 66, 82);

                urlsExterno.AddCell(ValorTipoUrlDetalle);
                urlsExterno.AddCell(ValorDetalleUrlDetalle);
                urlsExterno.AddCell(ValorNumeroTelefono);
                urlsExterno.AddCell(ValorUrlLinkDetalle);

                PdfPCell ValorUrlLinkDetalleCompleto = new PdfPCell(new Phrase(urls.url.ToString(), new Font(customfont, 6)));
                ValorUrlLinkDetalleCompleto.Colspan = 9; 
                ValorUrlLinkDetalleCompleto.VerticalAlignment = Element.ALIGN_MIDDLE; 
                ValorUrlLinkDetalleCompleto.BorderColor = new BaseColor(60, 66, 82);
                ValorUrlLinkDetalleCompleto.FixedHeight = 16f;

                urlsExternoLinkCompleto.AddCell(ValorUrlLinkDetalleCompleto);
            }

            if (ListadoUrlExterno.Count > 0)
            {
                document.Add(urlsExterno);

                //Salto Linea       
                PdfPTable saltoLinea5 = new PdfPTable(4);

                PdfPCell EtiquetaSaltoLinea5 = new PdfPCell(new Phrase(" ", new Font(customfontbold, 9)));
                EtiquetaSaltoLinea5.Colspan = 4;
                EtiquetaSaltoLinea5.Border = Rectangle.NO_BORDER;
                EtiquetaSaltoLinea5.HorizontalAlignment = Element.ALIGN_CENTER;
                EtiquetaSaltoLinea5.VerticalAlignment = Element.ALIGN_MIDDLE;
                EtiquetaSaltoLinea5.FixedHeight = 8f;

                saltoLinea5.AddCell(EtiquetaSaltoLinea5);

                document.Add(saltoLinea5);

                document.Add(urlsExternoLinkCompleto);

                document.Add(saltoLinea5); 

            }

            //Archivos Adjuntos
            PdfPTable camposAdjuntos = new PdfPTable(6);

            PdfPCell EtiquetaPiezaReferencia = new PdfPCell(new Phrase("PIEZA DE REFERENCIA", new Font(customfontbold, 9, 0, new BaseColor(255, 255, 255))));
            EtiquetaPiezaReferencia.Colspan = 6;
            EtiquetaPiezaReferencia.HorizontalAlignment = Element.ALIGN_CENTER;
            EtiquetaPiezaReferencia.VerticalAlignment = Element.ALIGN_MIDDLE;
            EtiquetaPiezaReferencia.FixedHeight = 16f;
            EtiquetaPiezaReferencia.BackgroundColor = new BaseColor(60, 66, 82);
            EtiquetaPiezaReferencia.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell EtiquetaNumero = new PdfPCell(new Phrase("# Adjunto", new Font(customfontbold, 7)));
            EtiquetaNumero.Colspan = 1;
            EtiquetaNumero.HorizontalAlignment = Element.ALIGN_CENTER;
            EtiquetaNumero.VerticalAlignment = Element.ALIGN_MIDDLE;
            EtiquetaNumero.FixedHeight = 16f;
            EtiquetaNumero.BackgroundColor = new BaseColor(225, 225, 225);
            EtiquetaNumero.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell EtiquetaNombreArchivo = new PdfPCell(new Phrase("Nombre Archivo", new Font(customfontbold, 7)));
            EtiquetaNombreArchivo.Colspan = 5;
            EtiquetaNombreArchivo.HorizontalAlignment = Element.ALIGN_LEFT;
            EtiquetaNombreArchivo.VerticalAlignment = Element.ALIGN_MIDDLE;
            EtiquetaNombreArchivo.FixedHeight = 16f;
            EtiquetaNombreArchivo.BackgroundColor = new BaseColor(225, 225, 225);
            EtiquetaNombreArchivo.BorderColor = new BaseColor(60, 66, 82);

            camposAdjuntos.AddCell(EtiquetaPiezaReferencia);

            camposAdjuntos.AddCell(EtiquetaNumero);
            camposAdjuntos.AddCell(EtiquetaNombreArchivo);

            foreach (AdjuntosSolicitud adjunto in ListadoAdjuntos)
            {
                PdfPCell EtiquetaCampoNumero = new PdfPCell(new Phrase(contadorAdjuntos.ToString(), new Font(customfont, 7)));
                EtiquetaCampoNumero.Colspan = 1;
                EtiquetaCampoNumero.HorizontalAlignment = Element.ALIGN_CENTER;
                EtiquetaCampoNumero.VerticalAlignment = Element.ALIGN_MIDDLE;
                EtiquetaCampoNumero.FixedHeight = 23f;
                EtiquetaCampoNumero.BorderColor = new BaseColor(60, 66, 82);

                PdfPCell ValorNombreAdjunto = new PdfPCell(new Phrase(" " + adjunto.nombre_adjunto, new Font(customfont, 7)));
                ValorNombreAdjunto.Colspan = 5;
                ValorNombreAdjunto.HorizontalAlignment = Element.ALIGN_LEFT;
                ValorNombreAdjunto.VerticalAlignment = Element.ALIGN_MIDDLE;
                ValorNombreAdjunto.FixedHeight = 23f;
                ValorNombreAdjunto.BorderColor = new BaseColor(60, 66, 82);

                camposAdjuntos.AddCell(EtiquetaCampoNumero);
                camposAdjuntos.AddCell(ValorNombreAdjunto);

                contadorAdjuntos = contadorAdjuntos + 1;
            }

            if (ListadoAdjuntos.Count > 0)
            {
                document.Add(camposAdjuntos);

                //Salto Linea       
                PdfPTable saltoLinea6 = new PdfPTable(4);

                PdfPCell EtiquetaSaltoLinea6 = new PdfPCell(new Phrase(" ", new Font(customfontbold, 9)));
                EtiquetaSaltoLinea6.Colspan = 4;
                EtiquetaSaltoLinea6.Border = Rectangle.NO_BORDER;
                EtiquetaSaltoLinea6.HorizontalAlignment = Element.ALIGN_CENTER;
                EtiquetaSaltoLinea6.VerticalAlignment = Element.ALIGN_MIDDLE;
                EtiquetaSaltoLinea6.FixedHeight = 8f;

                saltoLinea6.AddCell(EtiquetaSaltoLinea6);

                document.Add(saltoLinea6);
            }

            //Cerrar Documento
            document.Close();

        }

        //Reporte Solicitud Mantenimiento
        public void GenerarPDFMantenimiento(int? id)
        {
            int contadorUrls = 1;
            int contadorAdjuntos = 1;

            //Obtener los datos de la Solicitud
            DatosSolicitudCliente solicitud = db.DatosSolicitudCliente(id).First();
            MarcaPortales marcaPortales = db.ConsultarPortales(id).First();

            //Url Externo 
            List<UrlExternosSolicitud> ListadoUrlExterno = db.UrlExternosSolicitud(id).ToList();

            //Adjuntos
            List<AdjuntosSolicitud> ListadoAdjuntos = db.AdjuntosSolicitud(id).ToList();

            //Fonts para el PDF
            //Regular
            string fontpath = Server.MapPath("~/Content/fonts/ubuntu/Raleway-Regular.ttf");
            BaseFont customfont = BaseFont.CreateFont(fontpath, BaseFont.CP1252, BaseFont.EMBEDDED);

            //Bold
            string fontpathbold = Server.MapPath("~/Content/fonts/ubuntu/Raleway-Bold.ttf");
            BaseFont customfontbold = BaseFont.CreateFont(fontpathbold, BaseFont.CP1252, BaseFont.EMBEDDED);

            //ruta imagen del formulario
            String FilePath = Server.MapPath("~/Content/img/LogoPPMPDF.png");
            var physicalPath = Server.MapPath("~/Content/img/LogoPPMPDF.png");

            //anio Actual
            var anioActual = System.DateTime.Now.Year.ToString();

            string basePath = ConfigurationManager.AppSettings["RepositorioDocumentos"];
            string rutaArchivos = basePath + "\\GESTION_PPM\\ADJUNTOS_SOLICITUDES\\" + anioActual + '\\' + solicitud.Id.ToString();

            // En caso de que no exista el directorio, crearlo.
            bool directorio = Directory.Exists(rutaArchivos);
            if (!directorio)
                Directory.CreateDirectory(rutaArchivos);

            //ruta del archivo PDF que se va a crear
            string rutaDocumentos = rutaArchivos + "\\Solicitud-" + solicitud.CodigoSolicitud.ToString() + ".pdf";

            //Crear el formulario PDF
            FileStream fs = new FileStream(rutaDocumentos, FileMode.Create);
            Document document = new Document(iTextSharp.text.PageSize.A4, -20, -20, 0, 0);

            //Document document = new Document(iTextSharp.text.PageSize.A4);
            PdfWriter pw = PdfWriter.GetInstance(document, fs);

            //Abrir archivo para su edicion
            document.Open();

            //Salto de linea 
            document.Add(new Paragraph(" "));

            //Salto de linea 
            document.Add(new Paragraph(" "));

            //Cabecera del Cotizador         
            PdfPTable cabecera = new PdfPTable(7);
            cabecera.DefaultCell.Border = Rectangle.NO_BORDER;

            //Imagen
            iTextSharp.text.Image imagen = iTextSharp.text.Image.GetInstance(FilePath);
            imagen.ScalePercent(10);

            PdfPCell logocell = new PdfPCell(imagen, true);
            logocell.BackgroundColor = new BaseColor(60, 66, 82);
            logocell.Colspan = 2;
            logocell.Border = Rectangle.NO_BORDER;
            logocell.HorizontalAlignment = Element.ALIGN_LEFT;
            logocell.VerticalAlignment = Element.ALIGN_MIDDLE;

            //Titulo
            PdfPCell TituloCotizador = new PdfPCell(new Phrase("SOLICITUD " + solicitud.Tipo.ToUpper() + " N° " + solicitud.CodigoSolicitud + "      ", new Font(customfontbold, 9, 0, new BaseColor(255, 255, 255))));
            TituloCotizador.Colspan = 5;
            TituloCotizador.Border = Rectangle.NO_BORDER;
            TituloCotizador.HorizontalAlignment = Element.ALIGN_RIGHT;
            TituloCotizador.VerticalAlignment = Element.ALIGN_MIDDLE;
            TituloCotizador.BackgroundColor = new BaseColor(60, 66, 82);
            TituloCotizador.FixedHeight = 50f;

            cabecera.AddCell(logocell);
            cabecera.AddCell(TituloCotizador);

            document.Add(cabecera);

            //Salto Linea      
            PdfPTable saltoLinea1 = new PdfPTable(4);

            PdfPCell EtiquetaSaltoLinea = new PdfPCell(new Phrase(" ", new Font(customfontbold, 9)));
            EtiquetaSaltoLinea.Colspan = 4;
            EtiquetaSaltoLinea.Border = Rectangle.NO_BORDER;
            EtiquetaSaltoLinea.HorizontalAlignment = Element.ALIGN_CENTER;
            EtiquetaSaltoLinea.VerticalAlignment = Element.ALIGN_MIDDLE;
            EtiquetaSaltoLinea.FixedHeight = 8f;

            saltoLinea1.AddCell(EtiquetaSaltoLinea);

            document.Add(saltoLinea1);

            //Informacion de la solicitud         
            PdfPTable informacion = new PdfPTable(6);

            PdfPCell EtiquetaFechaCotizacion = new PdfPCell(new Phrase("Fecha y Hora: ", new Font(customfontbold, 7)));
            EtiquetaFechaCotizacion.Colspan = 1;
            EtiquetaFechaCotizacion.HorizontalAlignment = Element.ALIGN_RIGHT;
            EtiquetaFechaCotizacion.VerticalAlignment = Element.ALIGN_MIDDLE;
            EtiquetaFechaCotizacion.FixedHeight = 16f;
            EtiquetaFechaCotizacion.BackgroundColor = new BaseColor(225, 225, 225);
            EtiquetaFechaCotizacion.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell ValorFechaCotizacion = new PdfPCell(new Phrase("   " + solicitud.FechaHora, new Font(customfont, 7)));
            ValorFechaCotizacion.Colspan = 2;
            ValorFechaCotizacion.HorizontalAlignment = Element.ALIGN_LEFT;
            ValorFechaCotizacion.VerticalAlignment = Element.ALIGN_MIDDLE;
            ValorFechaCotizacion.FixedHeight = 16f;
            ValorFechaCotizacion.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell EtiquetaSolicitante = new PdfPCell(new Phrase("Solicitante:  ", new Font(customfontbold, 7)));
            EtiquetaSolicitante.Colspan = 1;
            EtiquetaSolicitante.HorizontalAlignment = Element.ALIGN_RIGHT;
            EtiquetaSolicitante.VerticalAlignment = Element.ALIGN_MIDDLE;
            EtiquetaSolicitante.FixedHeight = 16f;
            EtiquetaSolicitante.BackgroundColor = new BaseColor(225, 225, 225);
            EtiquetaSolicitante.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell ValorSolicitante = new PdfPCell(new Phrase("   " + solicitud.Solicitante, new Font(customfont, 7)));
            ValorSolicitante.Colspan = 2;
            ValorSolicitante.HorizontalAlignment = Element.ALIGN_LEFT;
            ValorSolicitante.VerticalAlignment = Element.ALIGN_MIDDLE;
            ValorSolicitante.FixedHeight = 16f;
            ValorSolicitante.BorderColor = new BaseColor(60, 66, 82);

            informacion.AddCell(EtiquetaFechaCotizacion);
            informacion.AddCell(ValorFechaCotizacion);
            informacion.AddCell(EtiquetaSolicitante);
            informacion.AddCell(ValorSolicitante);

            document.Add(informacion);

            //Salto Linea       
            PdfPTable saltoLinea2 = new PdfPTable(4);

            PdfPCell EtiquetaSaltoLinea2 = new PdfPCell(new Phrase(" ", new Font(customfontbold, 9)));
            EtiquetaSaltoLinea2.Colspan = 4;
            EtiquetaSaltoLinea2.Border = Rectangle.NO_BORDER;
            EtiquetaSaltoLinea2.HorizontalAlignment = Element.ALIGN_CENTER;
            EtiquetaSaltoLinea2.VerticalAlignment = Element.ALIGN_MIDDLE;
            EtiquetaSaltoLinea2.FixedHeight = 8f;

            saltoLinea2.AddCell(EtiquetaSaltoLinea2);

            document.Add(saltoLinea2);

            //Informacion de la Solicitud       
            PdfPTable datosGenerales = new PdfPTable(6);

            PdfPCell EtiquetaTipo = new PdfPCell(new Phrase("Tipo: ", new Font(customfontbold, 7)));
            EtiquetaTipo.Colspan = 1;
            EtiquetaTipo.HorizontalAlignment = Element.ALIGN_RIGHT;
            EtiquetaTipo.VerticalAlignment = Element.ALIGN_MIDDLE;
            EtiquetaTipo.FixedHeight = 16f;
            EtiquetaTipo.BackgroundColor = new BaseColor(225, 225, 225);
            EtiquetaTipo.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell ValorTipo = new PdfPCell(new Phrase("   " + solicitud.Tipo, new Font(customfont, 7)));
            ValorTipo.Colspan = 2;
            ValorTipo.HorizontalAlignment = Element.ALIGN_LEFT;
            ValorTipo.VerticalAlignment = Element.ALIGN_MIDDLE;
            ValorTipo.FixedHeight = 16f;
            ValorTipo.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell EtiquetaSubtipo = new PdfPCell(new Phrase("Subtipo: ", new Font(customfontbold, 7)));
            EtiquetaSubtipo.Colspan = 1;
            EtiquetaSubtipo.HorizontalAlignment = Element.ALIGN_RIGHT;
            EtiquetaSubtipo.VerticalAlignment = Element.ALIGN_MIDDLE;
            EtiquetaSubtipo.FixedHeight = 16f;
            EtiquetaSubtipo.BackgroundColor = new BaseColor(225, 225, 225);
            EtiquetaSubtipo.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell ValorSubtipo = new PdfPCell(new Phrase("   " + solicitud.Subtipo, new Font(customfont, 7)));
            ValorSubtipo.Colspan = 2;
            ValorSubtipo.HorizontalAlignment = Element.ALIGN_LEFT;
            ValorSubtipo.VerticalAlignment = Element.ALIGN_MIDDLE;
            ValorSubtipo.FixedHeight = 16f;
            ValorSubtipo.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell EtiquetaMarca = new PdfPCell(new Phrase("Marca(s): ", new Font(customfontbold, 7)));
            EtiquetaMarca.Colspan = 1;
            EtiquetaMarca.HorizontalAlignment = Element.ALIGN_RIGHT;
            EtiquetaMarca.VerticalAlignment = Element.ALIGN_MIDDLE;
            EtiquetaMarca.FixedHeight = 16f;
            EtiquetaMarca.BackgroundColor = new BaseColor(225, 225, 225);
            EtiquetaMarca.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell ValorMarca = new PdfPCell(new Phrase("   " + marcaPortales.Portales, new Font(customfont, 7)));
            ValorMarca.Colspan = 5;
            ValorMarca.HorizontalAlignment = Element.ALIGN_LEFT;
            ValorMarca.VerticalAlignment = Element.ALIGN_MIDDLE;
            ValorMarca.FixedHeight = 16f;
            ValorMarca.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell EtiquetaAreaSolicitante = new PdfPCell(new Phrase("Aréa Solicitante: ", new Font(customfontbold, 7)));
            EtiquetaAreaSolicitante.Colspan = 1;
            EtiquetaAreaSolicitante.HorizontalAlignment = Element.ALIGN_RIGHT;
            EtiquetaAreaSolicitante.VerticalAlignment = Element.ALIGN_MIDDLE;
            EtiquetaAreaSolicitante.FixedHeight = 16f;
            EtiquetaAreaSolicitante.BackgroundColor = new BaseColor(225, 225, 225);
            EtiquetaAreaSolicitante.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell AreaSolicitante = new PdfPCell(new Phrase("   " + solicitud.AreaEncargada, new Font(customfont, 7)));
            AreaSolicitante.Colspan = 2;
            AreaSolicitante.HorizontalAlignment = Element.ALIGN_LEFT;
            AreaSolicitante.VerticalAlignment = Element.ALIGN_MIDDLE;
            AreaSolicitante.FixedHeight = 16f;
            AreaSolicitante.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell EtiquetaUrgente = new PdfPCell(new Phrase("Urgente:  ", new Font(customfontbold, 7)));
            EtiquetaUrgente.Colspan = 1;
            EtiquetaUrgente.HorizontalAlignment = Element.ALIGN_RIGHT;
            EtiquetaUrgente.VerticalAlignment = Element.ALIGN_MIDDLE;
            EtiquetaUrgente.FixedHeight = 16f;
            EtiquetaUrgente.BackgroundColor = new BaseColor(225, 225, 225);
            EtiquetaUrgente.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell Urgente = new PdfPCell(new Phrase("   " + solicitud.Urgente, new Font(customfont, 7)));
            Urgente.Colspan = 2;
            Urgente.HorizontalAlignment = Element.ALIGN_LEFT;
            Urgente.VerticalAlignment = Element.ALIGN_MIDDLE;
            Urgente.FixedHeight = 16f;
            Urgente.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell EtiquetaImplementacion = new PdfPCell(new Phrase("Asunto:  ", new Font(customfontbold, 7)));
            EtiquetaImplementacion.Colspan = 1;
            EtiquetaImplementacion.HorizontalAlignment = Element.ALIGN_RIGHT;
            EtiquetaImplementacion.VerticalAlignment = Element.ALIGN_MIDDLE;
            EtiquetaImplementacion.FixedHeight = 16f;
            EtiquetaImplementacion.BackgroundColor = new BaseColor(225, 225, 225);
            EtiquetaImplementacion.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell ValorImplementacion = new PdfPCell(new Phrase("   " + solicitud.Asunto, new Font(customfont, 7)));
            ValorImplementacion.Colspan = 5;
            ValorImplementacion.HorizontalAlignment = Element.ALIGN_LEFT;
            ValorImplementacion.VerticalAlignment = Element.ALIGN_MIDDLE;
            ValorImplementacion.FixedHeight = 16f;
            ValorImplementacion.BorderColor = new BaseColor(60, 66, 82);


            datosGenerales.AddCell(EtiquetaTipo);
            datosGenerales.AddCell(ValorTipo);

            datosGenerales.AddCell(EtiquetaSubtipo);
            datosGenerales.AddCell(ValorSubtipo);

            datosGenerales.AddCell(EtiquetaMarca);
            datosGenerales.AddCell(ValorMarca);

            datosGenerales.AddCell(EtiquetaAreaSolicitante);
            datosGenerales.AddCell(AreaSolicitante);

            datosGenerales.AddCell(EtiquetaUrgente);
            datosGenerales.AddCell(Urgente);

            datosGenerales.AddCell(EtiquetaImplementacion);
            datosGenerales.AddCell(ValorImplementacion);


            document.Add(datosGenerales);


            //Salto Linea       
            PdfPTable saltoLinea3 = new PdfPTable(4);

            PdfPCell EtiquetaSaltoLinea3 = new PdfPCell(new Phrase(" ", new Font(customfontbold, 9)));
            EtiquetaSaltoLinea3.Colspan = 4;
            EtiquetaSaltoLinea3.Border = Rectangle.NO_BORDER;
            EtiquetaSaltoLinea3.HorizontalAlignment = Element.ALIGN_CENTER;
            EtiquetaSaltoLinea3.VerticalAlignment = Element.ALIGN_MIDDLE;
            EtiquetaSaltoLinea3.FixedHeight = 8f;

            saltoLinea3.AddCell(EtiquetaSaltoLinea3);

            document.Add(saltoLinea3);

            //Descripcion del requermiento    
            PdfPTable entregables = new PdfPTable(6);

            PdfPCell EtiquetaDescripcion = new PdfPCell(new Phrase("DESCRIPCIÓN DEL REQUERIMIENTO", new Font(customfontbold, 9, 0, new BaseColor(255, 255, 255))));
            EtiquetaDescripcion.Colspan = 6;
            EtiquetaDescripcion.HorizontalAlignment = Element.ALIGN_CENTER;
            EtiquetaDescripcion.VerticalAlignment = Element.ALIGN_MIDDLE;
            EtiquetaDescripcion.FixedHeight = 16f;
            EtiquetaDescripcion.BackgroundColor = new BaseColor(60, 66, 82);
            EtiquetaDescripcion.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell ValorDescripcion = new PdfPCell(new Phrase("   " + solicitud.Descripcion, new Font(customfont, 7)));
            ValorDescripcion.Colspan = 6;
            ValorDescripcion.Rowspan = 6;
            ValorDescripcion.VerticalAlignment = Element.ALIGN_MIDDLE;
            ValorDescripcion.BorderColor = new BaseColor(60, 66, 82);

            //validar el tamaño de las celdas
            var caracteres = solicitud.Descripcion.Count();
            if (caracteres <= 70)
            {
                ValorDescripcion.FixedHeight = 16f;
            }
            else if (caracteres <= 140)
            {
                ValorDescripcion.FixedHeight = 32f;
            }
            else if (caracteres <= 210)
            {
                ValorDescripcion.FixedHeight = 46f;
            }
            else if (caracteres <= 280)
            {
                ValorDescripcion.FixedHeight = 58f;
            }
            else if (caracteres <= 350)
            {
                ValorDescripcion.FixedHeight = 70f;
            }
            else if (caracteres <= 420)
            {
                ValorDescripcion.FixedHeight = 82f;
            }
            else if (caracteres <= 550)
            {
                ValorDescripcion.FixedHeight = 94f;
            }
            else if (caracteres <= 680)
            {
                ValorDescripcion.FixedHeight = 106f;
            }
            else if (caracteres <= 810)
            {
                ValorDescripcion.FixedHeight = 128f;
            }
            else
            {
                ValorDescripcion.FixedHeight = 140f;
            }

            entregables.AddCell(EtiquetaDescripcion);
            entregables.AddCell(ValorDescripcion);

            document.Add(entregables);

            //Salto Linea       
            PdfPTable saltoLinea4 = new PdfPTable(4);

            PdfPCell EtiquetaSaltoLinea4 = new PdfPCell(new Phrase(" ", new Font(customfontbold, 9)));
            EtiquetaSaltoLinea4.Colspan = 4;
            EtiquetaSaltoLinea4.Border = Rectangle.NO_BORDER;
            EtiquetaSaltoLinea4.HorizontalAlignment = Element.ALIGN_CENTER;
            EtiquetaSaltoLinea4.VerticalAlignment = Element.ALIGN_MIDDLE;
            EtiquetaSaltoLinea4.FixedHeight = 8f;

            saltoLinea4.AddCell(EtiquetaSaltoLinea4);

            document.Add(saltoLinea4);

            //Url Externo  
            PdfPTable urlsExterno = new PdfPTable(6);

            PdfPCell EtiquetaCampoUrls = new PdfPCell(new Phrase("URL's de Soporte", new Font(customfontbold, 9, 0, new BaseColor(255, 255, 255))));
            EtiquetaCampoUrls.Colspan = 6;
            EtiquetaCampoUrls.HorizontalAlignment = Element.ALIGN_CENTER;
            EtiquetaCampoUrls.VerticalAlignment = Element.ALIGN_MIDDLE;
            EtiquetaCampoUrls.FixedHeight = 16f;
            EtiquetaCampoUrls.BackgroundColor = new BaseColor(60, 66, 82);
            EtiquetaCampoUrls.BorderColor = new BaseColor(60, 66, 82);

            urlsExterno.AddCell(EtiquetaCampoUrls);

            foreach (UrlExternosSolicitud urls in ListadoUrlExterno)
            {
                PdfPCell ValorTipoUrlDetalle = new PdfPCell(new Phrase(contadorUrls.ToString(), new Font(customfont, 7)));
                ValorTipoUrlDetalle.Colspan = 1;
                ValorTipoUrlDetalle.HorizontalAlignment = Element.ALIGN_CENTER;
                ValorTipoUrlDetalle.VerticalAlignment = Element.ALIGN_MIDDLE;
                ValorTipoUrlDetalle.FixedHeight = 23f;
                ValorTipoUrlDetalle.BorderColor = new BaseColor(60, 66, 82);

                PdfPCell ValorDetalleUrlDetalle = new PdfPCell(new Phrase(urls.url.ToString(), new Font(customfont, 7)));
                ValorDetalleUrlDetalle.Colspan = 5;
                ValorDetalleUrlDetalle.HorizontalAlignment = Element.ALIGN_LEFT;
                ValorDetalleUrlDetalle.VerticalAlignment = Element.ALIGN_MIDDLE;
                ValorDetalleUrlDetalle.FixedHeight = 23f;
                ValorDetalleUrlDetalle.BorderColor = new BaseColor(60, 66, 82);

                urlsExterno.AddCell(ValorTipoUrlDetalle);
                urlsExterno.AddCell(ValorDetalleUrlDetalle);
            }

            if (ListadoUrlExterno.Count > 0)
            {
                document.Add(urlsExterno);

                //Salto Linea       
                PdfPTable saltoLinea5 = new PdfPTable(4);

                PdfPCell EtiquetaSaltoLinea5 = new PdfPCell(new Phrase(" ", new Font(customfontbold, 9)));
                EtiquetaSaltoLinea5.Colspan = 4;
                EtiquetaSaltoLinea5.Border = Rectangle.NO_BORDER;
                EtiquetaSaltoLinea5.HorizontalAlignment = Element.ALIGN_CENTER;
                EtiquetaSaltoLinea5.VerticalAlignment = Element.ALIGN_MIDDLE;
                EtiquetaSaltoLinea5.FixedHeight = 8f;

                saltoLinea5.AddCell(EtiquetaSaltoLinea5);

                document.Add(saltoLinea5);
            }
            //Archivos Adjuntos
            PdfPTable camposAdjuntos = new PdfPTable(6);

            PdfPCell EtiquetaPiezaReferencia = new PdfPCell(new Phrase("Archivos Adjunto Soporte", new Font(customfontbold, 9, 0, new BaseColor(255, 255, 255))));
            EtiquetaPiezaReferencia.Colspan = 6;
            EtiquetaPiezaReferencia.HorizontalAlignment = Element.ALIGN_CENTER;
            EtiquetaPiezaReferencia.VerticalAlignment = Element.ALIGN_MIDDLE;
            EtiquetaPiezaReferencia.FixedHeight = 16f;
            EtiquetaPiezaReferencia.BackgroundColor = new BaseColor(60, 66, 82);
            EtiquetaPiezaReferencia.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell EtiquetaNumero = new PdfPCell(new Phrase("# Adjunto", new Font(customfontbold, 7)));
            EtiquetaNumero.Colspan = 1;
            EtiquetaNumero.HorizontalAlignment = Element.ALIGN_CENTER;
            EtiquetaNumero.VerticalAlignment = Element.ALIGN_MIDDLE;
            EtiquetaNumero.FixedHeight = 16f;
            EtiquetaNumero.BackgroundColor = new BaseColor(225, 225, 225);
            EtiquetaNumero.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell EtiquetaNombreArchivo = new PdfPCell(new Phrase("Nombre Archivo", new Font(customfontbold, 7)));
            EtiquetaNombreArchivo.Colspan = 5;
            EtiquetaNombreArchivo.HorizontalAlignment = Element.ALIGN_LEFT;
            EtiquetaNombreArchivo.VerticalAlignment = Element.ALIGN_MIDDLE;
            EtiquetaNombreArchivo.FixedHeight = 16f;
            EtiquetaNombreArchivo.BackgroundColor = new BaseColor(225, 225, 225);
            EtiquetaNombreArchivo.BorderColor = new BaseColor(60, 66, 82);

            camposAdjuntos.AddCell(EtiquetaPiezaReferencia);

            camposAdjuntos.AddCell(EtiquetaNumero);
            camposAdjuntos.AddCell(EtiquetaNombreArchivo);

            foreach (AdjuntosSolicitud adjunto in ListadoAdjuntos)
            {
                PdfPCell EtiquetaCampoNumero = new PdfPCell(new Phrase(contadorAdjuntos.ToString(), new Font(customfont, 7)));
                EtiquetaCampoNumero.Colspan = 1;
                EtiquetaCampoNumero.HorizontalAlignment = Element.ALIGN_CENTER;
                EtiquetaCampoNumero.VerticalAlignment = Element.ALIGN_MIDDLE;
                EtiquetaCampoNumero.FixedHeight = 23f;
                EtiquetaCampoNumero.BorderColor = new BaseColor(60, 66, 82);

                PdfPCell ValorNombreAdjunto = new PdfPCell(new Phrase(adjunto.nombre_adjunto, new Font(customfont, 7)));
                ValorNombreAdjunto.Colspan = 5;
                ValorNombreAdjunto.HorizontalAlignment = Element.ALIGN_LEFT;
                ValorNombreAdjunto.VerticalAlignment = Element.ALIGN_MIDDLE;
                ValorNombreAdjunto.FixedHeight = 23f;
                ValorNombreAdjunto.BorderColor = new BaseColor(60, 66, 82);

                camposAdjuntos.AddCell(EtiquetaCampoNumero);
                camposAdjuntos.AddCell(ValorNombreAdjunto);

                contadorAdjuntos = contadorAdjuntos + 1;
            }

            if (ListadoAdjuntos.Count > 0)
            {
                document.Add(camposAdjuntos);

                //Salto Linea       
                PdfPTable saltoLinea6 = new PdfPTable(4);

                PdfPCell EtiquetaSaltoLinea6 = new PdfPCell(new Phrase(" ", new Font(customfontbold, 9)));
                EtiquetaSaltoLinea6.Colspan = 4;
                EtiquetaSaltoLinea6.Border = Rectangle.NO_BORDER;
                EtiquetaSaltoLinea6.HorizontalAlignment = Element.ALIGN_CENTER;
                EtiquetaSaltoLinea6.VerticalAlignment = Element.ALIGN_MIDDLE;
                EtiquetaSaltoLinea6.FixedHeight = 8f;

                saltoLinea6.AddCell(EtiquetaSaltoLinea6);
            }

            document.Close();
        }

        //Reporte Solicitud Envio Icare
        public void GenerarPDFEnvioIcare(int? id)
        {
            //Obtener los datos de la Solicitud
            DatosSolicitudCliente solicitud = db.DatosSolicitudCliente(id).First();

            //Adjuntos
            List<AdjuntosSolicitud> ListadoAdjuntos = db.AdjuntosSolicitud(id).ToList();

            //Fonts para el PDF
            //Regular
            string fontpath = Server.MapPath("~/Content/fonts/ubuntu/Raleway-Regular.ttf");
            BaseFont customfont = BaseFont.CreateFont(fontpath, BaseFont.CP1252, BaseFont.EMBEDDED);

            //Bold
            string fontpathbold = Server.MapPath("~/Content/fonts/ubuntu/Raleway-Bold.ttf");
            BaseFont customfontbold = BaseFont.CreateFont(fontpathbold, BaseFont.CP1252, BaseFont.EMBEDDED);

            //ruta imagen del formulario
            String FilePath = Server.MapPath("~/Content/img/LogoPPMPDF.png");
            var physicalPath = Server.MapPath("~/Content/img/LogoPPMPDF.png");

            //anio Actual
            var anioActual = System.DateTime.Now.Year.ToString();

            string basePath = ConfigurationManager.AppSettings["RepositorioDocumentos"];
            string rutaArchivos = basePath + "\\GESTION_PPM\\ADJUNTOS_SOLICITUDES\\" + anioActual + '\\' + solicitud.Id.ToString();

            // En caso de que no exista el directorio, crearlo.
            bool directorio = Directory.Exists(rutaArchivos);
            if (!directorio)
                Directory.CreateDirectory(rutaArchivos);

            //ruta del archivo PDF que se va a crear
            string rutaDocumentos = rutaArchivos + "\\Solicitud-" + solicitud.CodigoSolicitud.ToString() + ".pdf";

            //Crear el formulario PDF
            FileStream fs = new FileStream(rutaDocumentos, FileMode.Create);
            Document document = new Document(iTextSharp.text.PageSize.A4, -20, -20, 0, 0);

            //Document document = new Document(iTextSharp.text.PageSize.A4);
            PdfWriter pw = PdfWriter.GetInstance(document, fs);

            //Abrir archivo para su edicion
            document.Open();

            //Salto de linea 
            document.Add(new Paragraph(" "));

            //Salto de linea 
            document.Add(new Paragraph(" "));

            //Cabecera del Cotizador         
            PdfPTable cabecera = new PdfPTable(7);
            cabecera.DefaultCell.Border = Rectangle.NO_BORDER;

            //Imagen
            iTextSharp.text.Image imagen = iTextSharp.text.Image.GetInstance(FilePath);
            imagen.ScalePercent(10);

            PdfPCell logocell = new PdfPCell(imagen, true);
            logocell.BackgroundColor = new BaseColor(60, 66, 82);
            logocell.Colspan = 2;
            logocell.Border = Rectangle.NO_BORDER;
            logocell.HorizontalAlignment = Element.ALIGN_LEFT;
            logocell.VerticalAlignment = Element.ALIGN_MIDDLE;

            //Titulo
            PdfPCell TituloCotizador = new PdfPCell(new Phrase("SOLICITUD " + solicitud.Tipo.ToUpper() + " N° " + solicitud.CodigoSolicitud + "      ", new Font(customfontbold, 9, 0, new BaseColor(255, 255, 255))));
            TituloCotizador.Colspan = 5;
            TituloCotizador.Border = Rectangle.NO_BORDER;
            TituloCotizador.HorizontalAlignment = Element.ALIGN_RIGHT;
            TituloCotizador.VerticalAlignment = Element.ALIGN_MIDDLE;
            TituloCotizador.BackgroundColor = new BaseColor(60, 66, 82);
            TituloCotizador.FixedHeight = 50f;

            cabecera.AddCell(logocell);
            cabecera.AddCell(TituloCotizador);

            document.Add(cabecera);

            //Salto Linea      
            PdfPTable saltoLinea1 = new PdfPTable(4);

            PdfPCell EtiquetaSaltoLinea = new PdfPCell(new Phrase(" ", new Font(customfontbold, 9)));
            EtiquetaSaltoLinea.Colspan = 4;
            EtiquetaSaltoLinea.Border = Rectangle.NO_BORDER;
            EtiquetaSaltoLinea.HorizontalAlignment = Element.ALIGN_CENTER;
            EtiquetaSaltoLinea.VerticalAlignment = Element.ALIGN_MIDDLE;
            EtiquetaSaltoLinea.FixedHeight = 8f;

            saltoLinea1.AddCell(EtiquetaSaltoLinea);

            document.Add(saltoLinea1);

            //Informacion de la solicitud         
            PdfPTable informacion = new PdfPTable(6);

            PdfPCell EtiquetaFechaCotizacion = new PdfPCell(new Phrase("Fecha y Hora: ", new Font(customfontbold, 7)));
            EtiquetaFechaCotizacion.Colspan = 1;
            EtiquetaFechaCotizacion.HorizontalAlignment = Element.ALIGN_RIGHT;
            EtiquetaFechaCotizacion.VerticalAlignment = Element.ALIGN_MIDDLE;
            EtiquetaFechaCotizacion.FixedHeight = 16f;
            EtiquetaFechaCotizacion.BackgroundColor = new BaseColor(225, 225, 225);
            EtiquetaFechaCotizacion.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell ValorFechaCotizacion = new PdfPCell(new Phrase("   " + solicitud.FechaHora, new Font(customfont, 7)));
            ValorFechaCotizacion.Colspan = 2;
            ValorFechaCotizacion.HorizontalAlignment = Element.ALIGN_LEFT;
            ValorFechaCotizacion.VerticalAlignment = Element.ALIGN_MIDDLE;
            ValorFechaCotizacion.FixedHeight = 16f;
            ValorFechaCotizacion.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell EtiquetaSolicitante = new PdfPCell(new Phrase("Solicitante:  ", new Font(customfontbold, 7)));
            EtiquetaSolicitante.Colspan = 1;
            EtiquetaSolicitante.HorizontalAlignment = Element.ALIGN_RIGHT;
            EtiquetaSolicitante.VerticalAlignment = Element.ALIGN_MIDDLE;
            EtiquetaSolicitante.FixedHeight = 16f;
            EtiquetaSolicitante.BackgroundColor = new BaseColor(225, 225, 225);
            EtiquetaSolicitante.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell ValorSolicitante = new PdfPCell(new Phrase("   " + solicitud.Solicitante, new Font(customfont, 7)));
            ValorSolicitante.Colspan = 2;
            ValorSolicitante.HorizontalAlignment = Element.ALIGN_LEFT;
            ValorSolicitante.VerticalAlignment = Element.ALIGN_MIDDLE;
            ValorSolicitante.FixedHeight = 16f;
            ValorSolicitante.BorderColor = new BaseColor(60, 66, 82);

            informacion.AddCell(EtiquetaFechaCotizacion);
            informacion.AddCell(ValorFechaCotizacion);
            informacion.AddCell(EtiquetaSolicitante);
            informacion.AddCell(ValorSolicitante);

            document.Add(informacion);

            //Salto Linea       
            PdfPTable saltoLinea2 = new PdfPTable(4);

            PdfPCell EtiquetaSaltoLinea2 = new PdfPCell(new Phrase(" ", new Font(customfontbold, 9)));
            EtiquetaSaltoLinea2.Colspan = 4;
            EtiquetaSaltoLinea2.Border = Rectangle.NO_BORDER;
            EtiquetaSaltoLinea2.HorizontalAlignment = Element.ALIGN_CENTER;
            EtiquetaSaltoLinea2.VerticalAlignment = Element.ALIGN_MIDDLE;
            EtiquetaSaltoLinea2.FixedHeight = 8f;

            saltoLinea2.AddCell(EtiquetaSaltoLinea2);

            document.Add(saltoLinea2);

            //Informacion de la Solicitud       
            PdfPTable datosGenerales = new PdfPTable(6);

            PdfPCell EtiquetaTipo = new PdfPCell(new Phrase("Tipo: ", new Font(customfontbold, 7)));
            EtiquetaTipo.Colspan = 1;
            EtiquetaTipo.HorizontalAlignment = Element.ALIGN_RIGHT;
            EtiquetaTipo.VerticalAlignment = Element.ALIGN_MIDDLE;
            EtiquetaTipo.FixedHeight = 16f;
            EtiquetaTipo.BackgroundColor = new BaseColor(225, 225, 225);
            EtiquetaTipo.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell ValorTipo = new PdfPCell(new Phrase("   " + solicitud.Tipo, new Font(customfont, 7)));
            ValorTipo.Colspan = 5;
            ValorTipo.HorizontalAlignment = Element.ALIGN_LEFT;
            ValorTipo.VerticalAlignment = Element.ALIGN_MIDDLE;
            ValorTipo.FixedHeight = 16f;
            ValorTipo.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell EtiquetaSubtipo = new PdfPCell(new Phrase("Subtipo: ", new Font(customfontbold, 7)));
            EtiquetaSubtipo.Colspan = 1;
            EtiquetaSubtipo.HorizontalAlignment = Element.ALIGN_RIGHT;
            EtiquetaSubtipo.VerticalAlignment = Element.ALIGN_MIDDLE;
            EtiquetaSubtipo.FixedHeight = 16f;
            EtiquetaSubtipo.BackgroundColor = new BaseColor(225, 225, 225);
            EtiquetaSubtipo.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell ValorSubtipo = new PdfPCell(new Phrase("   " + solicitud.Subtipo, new Font(customfont, 7)));
            ValorSubtipo.Colspan = 5;
            ValorSubtipo.HorizontalAlignment = Element.ALIGN_LEFT;
            ValorSubtipo.VerticalAlignment = Element.ALIGN_MIDDLE;
            ValorSubtipo.FixedHeight = 16f;
            ValorSubtipo.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell EtiquetaOP = new PdfPCell(new Phrase("OP: ", new Font(customfontbold, 7)));
            EtiquetaOP.Colspan = 1;
            EtiquetaOP.HorizontalAlignment = Element.ALIGN_RIGHT;
            EtiquetaOP.VerticalAlignment = Element.ALIGN_MIDDLE;
            EtiquetaOP.FixedHeight = 16f;
            EtiquetaOP.BackgroundColor = new BaseColor(225, 225, 225);
            EtiquetaOP.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell ValorOP = new PdfPCell(new Phrase("   " + solicitud.OP, new Font(customfont, 7)));
            ValorOP.Colspan = 2;
            ValorOP.HorizontalAlignment = Element.ALIGN_LEFT;
            ValorOP.VerticalAlignment = Element.ALIGN_MIDDLE;
            ValorOP.FixedHeight = 16f;
            ValorOP.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell EtiquetaMKT = new PdfPCell(new Phrase("MKT: ", new Font(customfontbold, 7)));
            EtiquetaMKT.Colspan = 1;
            EtiquetaMKT.HorizontalAlignment = Element.ALIGN_RIGHT;
            EtiquetaMKT.VerticalAlignment = Element.ALIGN_MIDDLE;
            EtiquetaMKT.FixedHeight = 16f;
            EtiquetaMKT.BackgroundColor = new BaseColor(225, 225, 225);
            EtiquetaMKT.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell ValorMKT = new PdfPCell(new Phrase("   " + solicitud.MKT, new Font(customfont, 7)));
            ValorMKT.Colspan = 2;
            ValorMKT.HorizontalAlignment = Element.ALIGN_LEFT;
            ValorMKT.VerticalAlignment = Element.ALIGN_MIDDLE;
            ValorMKT.FixedHeight = 16f;
            ValorMKT.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell EtiquetaNombre = new PdfPCell(new Phrase("Nombre Campaña: ", new Font(customfontbold, 7)));
            EtiquetaNombre.Colspan = 1;
            EtiquetaNombre.HorizontalAlignment = Element.ALIGN_RIGHT;
            EtiquetaNombre.VerticalAlignment = Element.ALIGN_MIDDLE;
            EtiquetaNombre.FixedHeight = 16f;
            EtiquetaNombre.BackgroundColor = new BaseColor(225, 225, 225);
            EtiquetaNombre.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell ValorNombre = new PdfPCell(new Phrase("   " + solicitud.Nombre, new Font(customfont, 7)));
            ValorNombre.Colspan = 5;
            ValorNombre.HorizontalAlignment = Element.ALIGN_LEFT;
            ValorNombre.VerticalAlignment = Element.ALIGN_MIDDLE;
            ValorNombre.FixedHeight = 16f;
            ValorNombre.BorderColor = new BaseColor(60, 66, 82);

            datosGenerales.AddCell(EtiquetaTipo);
            datosGenerales.AddCell(ValorTipo);

            datosGenerales.AddCell(EtiquetaSubtipo);
            datosGenerales.AddCell(ValorSubtipo);

            datosGenerales.AddCell(EtiquetaOP);
            datosGenerales.AddCell(ValorOP);
            datosGenerales.AddCell(EtiquetaMKT);
            datosGenerales.AddCell(ValorMKT);

            datosGenerales.AddCell(EtiquetaNombre);
            datosGenerales.AddCell(ValorNombre);

            document.Add(datosGenerales);

            //Salto Linea       
            PdfPTable saltoLinea3 = new PdfPTable(4);

            PdfPCell EtiquetaSaltoLinea3 = new PdfPCell(new Phrase(" ", new Font(customfontbold, 9)));
            EtiquetaSaltoLinea3.Colspan = 4;
            EtiquetaSaltoLinea3.Border = Rectangle.NO_BORDER;
            EtiquetaSaltoLinea3.HorizontalAlignment = Element.ALIGN_CENTER;
            EtiquetaSaltoLinea3.VerticalAlignment = Element.ALIGN_MIDDLE;
            EtiquetaSaltoLinea3.FixedHeight = 8f;

            saltoLinea3.AddCell(EtiquetaSaltoLinea3);

            document.Add(saltoLinea3);

            //Archivos Adjuntos
            PdfPTable camposAdjuntos = new PdfPTable(6);

            PdfPCell EtiquetaPiezaReferencia = new PdfPCell(new Phrase("Archivos Arjuntos", new Font(customfontbold, 9, 0, new BaseColor(255, 255, 255))));
            EtiquetaPiezaReferencia.Colspan = 6;
            EtiquetaPiezaReferencia.HorizontalAlignment = Element.ALIGN_CENTER;
            EtiquetaPiezaReferencia.VerticalAlignment = Element.ALIGN_MIDDLE;
            EtiquetaPiezaReferencia.FixedHeight = 16f;
            EtiquetaPiezaReferencia.BackgroundColor = new BaseColor(60, 66, 82);
            EtiquetaPiezaReferencia.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell EtiquetaNumero = new PdfPCell(new Phrase("Tipo", new Font(customfontbold, 7)));
            EtiquetaNumero.Colspan = 2;
            EtiquetaNumero.HorizontalAlignment = Element.ALIGN_CENTER;
            EtiquetaNumero.VerticalAlignment = Element.ALIGN_MIDDLE;
            EtiquetaNumero.FixedHeight = 16f;
            EtiquetaNumero.BackgroundColor = new BaseColor(225, 225, 225);
            EtiquetaNumero.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell EtiquetaNombreArchivo = new PdfPCell(new Phrase("Nombre Archivo", new Font(customfontbold, 7)));
            EtiquetaNombreArchivo.Colspan = 4;
            EtiquetaNombreArchivo.HorizontalAlignment = Element.ALIGN_LEFT;
            EtiquetaNombreArchivo.VerticalAlignment = Element.ALIGN_MIDDLE;
            EtiquetaNombreArchivo.FixedHeight = 16f;
            EtiquetaNombreArchivo.BackgroundColor = new BaseColor(225, 225, 225);
            EtiquetaNombreArchivo.BorderColor = new BaseColor(60, 66, 82);

            camposAdjuntos.AddCell(EtiquetaPiezaReferencia);

            camposAdjuntos.AddCell(EtiquetaNumero);
            camposAdjuntos.AddCell(EtiquetaNombreArchivo);

            foreach (AdjuntosSolicitud adjunto in ListadoAdjuntos)
            {
                PdfPCell EtiquetaCampoNumero = new PdfPCell(new Phrase(adjunto.tipo.ToString(), new Font(customfont, 7)));
                EtiquetaCampoNumero.Colspan = 2;
                EtiquetaCampoNumero.HorizontalAlignment = Element.ALIGN_CENTER;
                EtiquetaCampoNumero.VerticalAlignment = Element.ALIGN_MIDDLE;
                EtiquetaCampoNumero.FixedHeight = 23f;
                EtiquetaCampoNumero.BorderColor = new BaseColor(60, 66, 82);

                PdfPCell ValorNombreAdjunto = new PdfPCell(new Phrase(adjunto.nombre_adjunto, new Font(customfont, 7)));
                ValorNombreAdjunto.Colspan = 4;
                ValorNombreAdjunto.HorizontalAlignment = Element.ALIGN_LEFT;
                ValorNombreAdjunto.VerticalAlignment = Element.ALIGN_MIDDLE;
                ValorNombreAdjunto.FixedHeight = 23f;
                ValorNombreAdjunto.BorderColor = new BaseColor(60, 66, 82);

                camposAdjuntos.AddCell(EtiquetaCampoNumero);
                camposAdjuntos.AddCell(ValorNombreAdjunto);

            }

            document.Add(camposAdjuntos);

            //Salto Linea       
            PdfPTable saltoLinea6 = new PdfPTable(4);

            PdfPCell EtiquetaSaltoLinea6 = new PdfPCell(new Phrase(" ", new Font(customfontbold, 9)));
            EtiquetaSaltoLinea6.Colspan = 4;
            EtiquetaSaltoLinea6.Border = Rectangle.NO_BORDER;
            EtiquetaSaltoLinea6.HorizontalAlignment = Element.ALIGN_CENTER;
            EtiquetaSaltoLinea6.VerticalAlignment = Element.ALIGN_MIDDLE;
            EtiquetaSaltoLinea6.FixedHeight = 8f;

            saltoLinea6.AddCell(EtiquetaSaltoLinea6);

            document.Add(saltoLinea6);

            //Cerrar Documento
            document.Close();

        }

        //Reporte Solicitud Html5
        public void GenerarPDFHtml5(int? id)
        {
            //Obtener los datos de la Solicitud
            DatosSolicitudCliente solicitud = db.DatosSolicitudCliente(id).First();

            //Adjuntos
            List<AdjuntosSolicitud> ListadoAdjuntos = db.AdjuntosSolicitud(id).ToList();

            //Fonts para el PDF
            //Regular
            string fontpath = Server.MapPath("~/Content/fonts/ubuntu/Raleway-Regular.ttf");
            BaseFont customfont = BaseFont.CreateFont(fontpath, BaseFont.CP1252, BaseFont.EMBEDDED);

            //Bold
            string fontpathbold = Server.MapPath("~/Content/fonts/ubuntu/Raleway-Bold.ttf");
            BaseFont customfontbold = BaseFont.CreateFont(fontpathbold, BaseFont.CP1252, BaseFont.EMBEDDED);

            //ruta imagen del formulario
            String FilePath = Server.MapPath("~/Content/img/LogoPPMPDF.png");
            var physicalPath = Server.MapPath("~/Content/img/LogoPPMPDF.png");

            //anio Actual
            var anioActual = System.DateTime.Now.Year.ToString();

            string basePath = ConfigurationManager.AppSettings["RepositorioDocumentos"];
            string rutaArchivos = basePath + "\\GESTION_PPM\\ADJUNTOS_SOLICITUDES\\" + anioActual + '\\' + solicitud.Id.ToString();

            // En caso de que no exista el directorio, crearlo.
            bool directorio = Directory.Exists(rutaArchivos);
            if (!directorio)
                Directory.CreateDirectory(rutaArchivos);

            //ruta del archivo PDF que se va a crear
            string rutaDocumentos = rutaArchivos + "\\Solicitud-" + solicitud.CodigoSolicitud.ToString() + ".pdf";

            //Crear el formulario PDF
            FileStream fs = new FileStream(rutaDocumentos, FileMode.Create);
            Document document = new Document(iTextSharp.text.PageSize.A4, -20, -20, 0, 0);

            //Document document = new Document(iTextSharp.text.PageSize.A4);
            PdfWriter pw = PdfWriter.GetInstance(document, fs);

            //Abrir archivo para su edicion
            document.Open();

            //Salto de linea 
            document.Add(new Paragraph(" "));

            //Salto de linea 
            document.Add(new Paragraph(" "));

            //Cabecera del Cotizador         
            PdfPTable cabecera = new PdfPTable(7);
            cabecera.DefaultCell.Border = Rectangle.NO_BORDER;

            //Imagen
            iTextSharp.text.Image imagen = iTextSharp.text.Image.GetInstance(FilePath);
            imagen.ScalePercent(10);

            PdfPCell logocell = new PdfPCell(imagen, true);
            logocell.BackgroundColor = new BaseColor(60, 66, 82);
            logocell.Colspan = 2;
            logocell.Border = Rectangle.NO_BORDER;
            logocell.HorizontalAlignment = Element.ALIGN_LEFT;
            logocell.VerticalAlignment = Element.ALIGN_MIDDLE;

            //Titulo
            PdfPCell TituloCotizador = new PdfPCell(new Phrase("SOLICITUD " + solicitud.Tipo.ToUpper() + " N° " + solicitud.CodigoSolicitud + "      ", new Font(customfontbold, 9, 0, new BaseColor(255, 255, 255))));
            TituloCotizador.Colspan = 5;
            TituloCotizador.Border = Rectangle.NO_BORDER;
            TituloCotizador.HorizontalAlignment = Element.ALIGN_RIGHT;
            TituloCotizador.VerticalAlignment = Element.ALIGN_MIDDLE;
            TituloCotizador.BackgroundColor = new BaseColor(60, 66, 82);
            TituloCotizador.FixedHeight = 50f;

            cabecera.AddCell(logocell);
            cabecera.AddCell(TituloCotizador);

            document.Add(cabecera);

            //Salto Linea      
            PdfPTable saltoLinea1 = new PdfPTable(4);

            PdfPCell EtiquetaSaltoLinea = new PdfPCell(new Phrase(" ", new Font(customfontbold, 9)));
            EtiquetaSaltoLinea.Colspan = 4;
            EtiquetaSaltoLinea.Border = Rectangle.NO_BORDER;
            EtiquetaSaltoLinea.HorizontalAlignment = Element.ALIGN_CENTER;
            EtiquetaSaltoLinea.VerticalAlignment = Element.ALIGN_MIDDLE;
            EtiquetaSaltoLinea.FixedHeight = 8f;

            saltoLinea1.AddCell(EtiquetaSaltoLinea);

            document.Add(saltoLinea1);

            //Informacion de la solicitud         
            PdfPTable informacion = new PdfPTable(6);

            PdfPCell EtiquetaFechaCotizacion = new PdfPCell(new Phrase("Fecha y Hora: ", new Font(customfontbold, 7)));
            EtiquetaFechaCotizacion.Colspan = 1;
            EtiquetaFechaCotizacion.HorizontalAlignment = Element.ALIGN_RIGHT;
            EtiquetaFechaCotizacion.VerticalAlignment = Element.ALIGN_MIDDLE;
            EtiquetaFechaCotizacion.FixedHeight = 16f;
            EtiquetaFechaCotizacion.BackgroundColor = new BaseColor(225, 225, 225);
            EtiquetaFechaCotizacion.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell ValorFechaCotizacion = new PdfPCell(new Phrase("   " + solicitud.FechaHora, new Font(customfont, 7)));
            ValorFechaCotizacion.Colspan = 2;
            ValorFechaCotizacion.HorizontalAlignment = Element.ALIGN_LEFT;
            ValorFechaCotizacion.VerticalAlignment = Element.ALIGN_MIDDLE;
            ValorFechaCotizacion.FixedHeight = 16f;
            ValorFechaCotizacion.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell EtiquetaSolicitante = new PdfPCell(new Phrase("Solicitante:  ", new Font(customfontbold, 7)));
            EtiquetaSolicitante.Colspan = 1;
            EtiquetaSolicitante.HorizontalAlignment = Element.ALIGN_RIGHT;
            EtiquetaSolicitante.VerticalAlignment = Element.ALIGN_MIDDLE;
            EtiquetaSolicitante.FixedHeight = 16f;
            EtiquetaSolicitante.BackgroundColor = new BaseColor(225, 225, 225);
            EtiquetaSolicitante.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell ValorSolicitante = new PdfPCell(new Phrase("   " + solicitud.Solicitante, new Font(customfont, 7)));
            ValorSolicitante.Colspan = 2;
            ValorSolicitante.HorizontalAlignment = Element.ALIGN_LEFT;
            ValorSolicitante.VerticalAlignment = Element.ALIGN_MIDDLE;
            ValorSolicitante.FixedHeight = 16f;
            ValorSolicitante.BorderColor = new BaseColor(60, 66, 82);

            informacion.AddCell(EtiquetaFechaCotizacion);
            informacion.AddCell(ValorFechaCotizacion);
            informacion.AddCell(EtiquetaSolicitante);
            informacion.AddCell(ValorSolicitante);

            document.Add(informacion);

            //Salto Linea       
            PdfPTable saltoLinea2 = new PdfPTable(4);

            PdfPCell EtiquetaSaltoLinea2 = new PdfPCell(new Phrase(" ", new Font(customfontbold, 9)));
            EtiquetaSaltoLinea2.Colspan = 4;
            EtiquetaSaltoLinea2.Border = Rectangle.NO_BORDER;
            EtiquetaSaltoLinea2.HorizontalAlignment = Element.ALIGN_CENTER;
            EtiquetaSaltoLinea2.VerticalAlignment = Element.ALIGN_MIDDLE;
            EtiquetaSaltoLinea2.FixedHeight = 8f;

            saltoLinea2.AddCell(EtiquetaSaltoLinea2);

            document.Add(saltoLinea2);

            //Informacion de la Solicitud       
            PdfPTable datosGenerales = new PdfPTable(6);

            PdfPCell EtiquetaTipo = new PdfPCell(new Phrase("Tipo: ", new Font(customfontbold, 7)));
            EtiquetaTipo.Colspan = 1;
            EtiquetaTipo.HorizontalAlignment = Element.ALIGN_RIGHT;
            EtiquetaTipo.VerticalAlignment = Element.ALIGN_MIDDLE;
            EtiquetaTipo.FixedHeight = 16f;
            EtiquetaTipo.BackgroundColor = new BaseColor(225, 225, 225);
            EtiquetaTipo.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell ValorTipo = new PdfPCell(new Phrase("   " + solicitud.Tipo, new Font(customfont, 7)));
            ValorTipo.Colspan = 5;
            ValorTipo.HorizontalAlignment = Element.ALIGN_LEFT;
            ValorTipo.VerticalAlignment = Element.ALIGN_MIDDLE;
            ValorTipo.FixedHeight = 16f;
            ValorTipo.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell EtiquetaSubtipo = new PdfPCell(new Phrase("Subtipo: ", new Font(customfontbold, 7)));
            EtiquetaSubtipo.Colspan = 1;
            EtiquetaSubtipo.HorizontalAlignment = Element.ALIGN_RIGHT;
            EtiquetaSubtipo.VerticalAlignment = Element.ALIGN_MIDDLE;
            EtiquetaSubtipo.FixedHeight = 16f;
            EtiquetaSubtipo.BackgroundColor = new BaseColor(225, 225, 225);
            EtiquetaSubtipo.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell ValorSubtipo = new PdfPCell(new Phrase("   " + solicitud.Subtipo, new Font(customfont, 7)));
            ValorSubtipo.Colspan = 5;
            ValorSubtipo.HorizontalAlignment = Element.ALIGN_LEFT;
            ValorSubtipo.VerticalAlignment = Element.ALIGN_MIDDLE;
            ValorSubtipo.FixedHeight = 16f;
            ValorSubtipo.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell EtiquetaOP = new PdfPCell(new Phrase("OP: ", new Font(customfontbold, 7)));
            EtiquetaOP.Colspan = 1;
            EtiquetaOP.HorizontalAlignment = Element.ALIGN_RIGHT;
            EtiquetaOP.VerticalAlignment = Element.ALIGN_MIDDLE;
            EtiquetaOP.FixedHeight = 16f;
            EtiquetaOP.BackgroundColor = new BaseColor(225, 225, 225);
            EtiquetaOP.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell ValorOP = new PdfPCell(new Phrase("   " + solicitud.OP, new Font(customfont, 7)));
            ValorOP.Colspan = 2;
            ValorOP.HorizontalAlignment = Element.ALIGN_LEFT;
            ValorOP.VerticalAlignment = Element.ALIGN_MIDDLE;
            ValorOP.FixedHeight = 16f;
            ValorOP.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell EtiquetaMKT = new PdfPCell(new Phrase("MKT: ", new Font(customfontbold, 7)));
            EtiquetaMKT.Colspan = 1;
            EtiquetaMKT.HorizontalAlignment = Element.ALIGN_RIGHT;
            EtiquetaMKT.VerticalAlignment = Element.ALIGN_MIDDLE;
            EtiquetaMKT.FixedHeight = 16f;
            EtiquetaMKT.BackgroundColor = new BaseColor(225, 225, 225);
            EtiquetaMKT.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell ValorMKT = new PdfPCell(new Phrase("   " + solicitud.MKT, new Font(customfont, 7)));
            ValorMKT.Colspan = 2;
            ValorMKT.HorizontalAlignment = Element.ALIGN_LEFT;
            ValorMKT.VerticalAlignment = Element.ALIGN_MIDDLE;
            ValorMKT.FixedHeight = 16f;
            ValorMKT.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell EtiquetaNombre = new PdfPCell(new Phrase("Nombre Campaña: ", new Font(customfontbold, 7)));
            EtiquetaNombre.Colspan = 1;
            EtiquetaNombre.HorizontalAlignment = Element.ALIGN_RIGHT;
            EtiquetaNombre.VerticalAlignment = Element.ALIGN_MIDDLE;
            EtiquetaNombre.FixedHeight = 16f;
            EtiquetaNombre.BackgroundColor = new BaseColor(225, 225, 225);
            EtiquetaNombre.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell ValorNombre = new PdfPCell(new Phrase("   " + solicitud.Nombre, new Font(customfont, 7)));
            ValorNombre.Colspan = 5;
            ValorNombre.HorizontalAlignment = Element.ALIGN_LEFT;
            ValorNombre.VerticalAlignment = Element.ALIGN_MIDDLE;
            ValorNombre.FixedHeight = 16f;
            ValorNombre.BorderColor = new BaseColor(60, 66, 82);

            datosGenerales.AddCell(EtiquetaTipo);
            datosGenerales.AddCell(ValorTipo);

            datosGenerales.AddCell(EtiquetaSubtipo);
            datosGenerales.AddCell(ValorSubtipo);

            datosGenerales.AddCell(EtiquetaOP);
            datosGenerales.AddCell(ValorOP);
            datosGenerales.AddCell(EtiquetaMKT);
            datosGenerales.AddCell(ValorMKT);

            datosGenerales.AddCell(EtiquetaNombre);
            datosGenerales.AddCell(ValorNombre);

            document.Add(datosGenerales);

            //Salto Linea       
            PdfPTable saltoLinea3 = new PdfPTable(4);

            PdfPCell EtiquetaSaltoLinea3 = new PdfPCell(new Phrase(" ", new Font(customfontbold, 9)));
            EtiquetaSaltoLinea3.Colspan = 4;
            EtiquetaSaltoLinea3.Border = Rectangle.NO_BORDER;
            EtiquetaSaltoLinea3.HorizontalAlignment = Element.ALIGN_CENTER;
            EtiquetaSaltoLinea3.VerticalAlignment = Element.ALIGN_MIDDLE;
            EtiquetaSaltoLinea3.FixedHeight = 8f;

            saltoLinea3.AddCell(EtiquetaSaltoLinea3);

            document.Add(saltoLinea3);

            //Archivos Adjuntos
            PdfPTable camposAdjuntos = new PdfPTable(6);

            PdfPCell EtiquetaPiezaReferencia = new PdfPCell(new Phrase("Archivos Arjuntos", new Font(customfontbold, 9, 0, new BaseColor(255, 255, 255))));
            EtiquetaPiezaReferencia.Colspan = 6;
            EtiquetaPiezaReferencia.HorizontalAlignment = Element.ALIGN_CENTER;
            EtiquetaPiezaReferencia.VerticalAlignment = Element.ALIGN_MIDDLE;
            EtiquetaPiezaReferencia.FixedHeight = 16f;
            EtiquetaPiezaReferencia.BackgroundColor = new BaseColor(60, 66, 82);
            EtiquetaPiezaReferencia.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell EtiquetaNumero = new PdfPCell(new Phrase("Tipo", new Font(customfontbold, 7)));
            EtiquetaNumero.Colspan = 2;
            EtiquetaNumero.HorizontalAlignment = Element.ALIGN_CENTER;
            EtiquetaNumero.VerticalAlignment = Element.ALIGN_MIDDLE;
            EtiquetaNumero.FixedHeight = 16f;
            EtiquetaNumero.BackgroundColor = new BaseColor(225, 225, 225);
            EtiquetaNumero.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell EtiquetaNombreArchivo = new PdfPCell(new Phrase("Nombre Archivo", new Font(customfontbold, 7)));
            EtiquetaNombreArchivo.Colspan = 4;
            EtiquetaNombreArchivo.HorizontalAlignment = Element.ALIGN_LEFT;
            EtiquetaNombreArchivo.VerticalAlignment = Element.ALIGN_MIDDLE;
            EtiquetaNombreArchivo.FixedHeight = 16f;
            EtiquetaNombreArchivo.BackgroundColor = new BaseColor(225, 225, 225);
            EtiquetaNombreArchivo.BorderColor = new BaseColor(60, 66, 82);

            camposAdjuntos.AddCell(EtiquetaPiezaReferencia);

            camposAdjuntos.AddCell(EtiquetaNumero);
            camposAdjuntos.AddCell(EtiquetaNombreArchivo);

            foreach (AdjuntosSolicitud adjunto in ListadoAdjuntos)
            {
                PdfPCell EtiquetaCampoNumero = new PdfPCell(new Phrase(adjunto.tipo.ToString(), new Font(customfont, 7)));
                EtiquetaCampoNumero.Colspan = 2;
                EtiquetaCampoNumero.HorizontalAlignment = Element.ALIGN_CENTER;
                EtiquetaCampoNumero.VerticalAlignment = Element.ALIGN_MIDDLE;
                EtiquetaCampoNumero.FixedHeight = 23f;
                EtiquetaCampoNumero.BorderColor = new BaseColor(60, 66, 82);

                PdfPCell ValorNombreAdjunto = new PdfPCell(new Phrase(adjunto.nombre_adjunto, new Font(customfont, 7)));
                ValorNombreAdjunto.Colspan = 4;
                ValorNombreAdjunto.HorizontalAlignment = Element.ALIGN_LEFT;
                ValorNombreAdjunto.VerticalAlignment = Element.ALIGN_MIDDLE;
                ValorNombreAdjunto.FixedHeight = 23f;
                ValorNombreAdjunto.BorderColor = new BaseColor(60, 66, 82);

                camposAdjuntos.AddCell(EtiquetaCampoNumero);
                camposAdjuntos.AddCell(ValorNombreAdjunto);

            }

            document.Add(camposAdjuntos);

            //Salto Linea       
            PdfPTable saltoLinea6 = new PdfPTable(4);

            PdfPCell EtiquetaSaltoLinea6 = new PdfPCell(new Phrase(" ", new Font(customfontbold, 9)));
            EtiquetaSaltoLinea6.Colspan = 4;
            EtiquetaSaltoLinea6.Border = Rectangle.NO_BORDER;
            EtiquetaSaltoLinea6.HorizontalAlignment = Element.ALIGN_CENTER;
            EtiquetaSaltoLinea6.VerticalAlignment = Element.ALIGN_MIDDLE;
            EtiquetaSaltoLinea6.FixedHeight = 8f;

            saltoLinea6.AddCell(EtiquetaSaltoLinea6);

            document.Add(saltoLinea6);

            //Cerrar Documento
            document.Close();

        }

        //Reporte Solicitud Tags
        public void GenerarPDFTags(int? id)
        {
            int contadorUrls = 1;

            //Obtener los datos de la Solicitud
            DatosSolicitudCliente solicitud = db.DatosSolicitudCliente(id).First();

            //Url Externo 
            List<UrlExternosSolicitud> ListadoUrlExterno = db.UrlExternosSolicitud(id).ToList();

            //Fonts para el PDF
            //Regular
            string fontpath = Server.MapPath("~/Content/fonts/ubuntu/Raleway-Regular.ttf");
            BaseFont customfont = BaseFont.CreateFont(fontpath, BaseFont.CP1252, BaseFont.EMBEDDED);

            //Bold
            string fontpathbold = Server.MapPath("~/Content/fonts/ubuntu/Raleway-Bold.ttf");
            BaseFont customfontbold = BaseFont.CreateFont(fontpathbold, BaseFont.CP1252, BaseFont.EMBEDDED);

            //ruta imagen del formulario
            String FilePath = Server.MapPath("~/Content/img/LogoPPMPDF.png");
            var physicalPath = Server.MapPath("~/Content/img/LogoPPMPDF.png");

            //anio Actual
            var anioActual = System.DateTime.Now.Year.ToString();

            string basePath = ConfigurationManager.AppSettings["RepositorioDocumentos"];
            string rutaArchivos = basePath + "\\GESTION_PPM\\ADJUNTOS_SOLICITUDES\\" + anioActual + '\\' + solicitud.Id.ToString();

            // En caso de que no exista el directorio, crearlo.
            bool directorio = Directory.Exists(rutaArchivos);
            if (!directorio)
                Directory.CreateDirectory(rutaArchivos);

            //ruta del archivo PDF que se va a crear
            string rutaDocumentos = rutaArchivos + "\\Solicitud-" + solicitud.CodigoSolicitud.ToString() + ".pdf";

            //Crear el formulario PDF
            FileStream fs = new FileStream(rutaDocumentos, FileMode.Create);
            Document document = new Document(iTextSharp.text.PageSize.A4, -20, -20, 0, 0);

            //Document document = new Document(iTextSharp.text.PageSize.A4);
            PdfWriter pw = PdfWriter.GetInstance(document, fs);

            //Abrir archivo para su edicion
            document.Open();

            //Salto de linea 
            document.Add(new Paragraph(" "));

            //Salto de linea 
            document.Add(new Paragraph(" "));

            //Cabecera del Cotizador         
            PdfPTable cabecera = new PdfPTable(7);
            cabecera.DefaultCell.Border = Rectangle.NO_BORDER;

            //Imagen
            iTextSharp.text.Image imagen = iTextSharp.text.Image.GetInstance(FilePath);
            imagen.ScalePercent(10);

            PdfPCell logocell = new PdfPCell(imagen, true);
            logocell.BackgroundColor = new BaseColor(60, 66, 82);
            logocell.Colspan = 2;
            logocell.Border = Rectangle.NO_BORDER;
            logocell.HorizontalAlignment = Element.ALIGN_LEFT;
            logocell.VerticalAlignment = Element.ALIGN_MIDDLE;

            //Titulo
            PdfPCell TituloCotizador = new PdfPCell(new Phrase("SOLICITUD " + solicitud.Tipo.ToUpper() + " N° " + solicitud.CodigoSolicitud + "      ", new Font(customfontbold, 9, 0, new BaseColor(255, 255, 255))));
            TituloCotizador.Colspan = 5;
            TituloCotizador.Border = Rectangle.NO_BORDER;
            TituloCotizador.HorizontalAlignment = Element.ALIGN_RIGHT;
            TituloCotizador.VerticalAlignment = Element.ALIGN_MIDDLE;
            TituloCotizador.BackgroundColor = new BaseColor(60, 66, 82);
            TituloCotizador.FixedHeight = 50f;

            cabecera.AddCell(logocell);
            cabecera.AddCell(TituloCotizador);

            document.Add(cabecera);

            //Salto Linea      
            PdfPTable saltoLinea1 = new PdfPTable(4);

            PdfPCell EtiquetaSaltoLinea = new PdfPCell(new Phrase(" ", new Font(customfontbold, 9)));
            EtiquetaSaltoLinea.Colspan = 4;
            EtiquetaSaltoLinea.Border = Rectangle.NO_BORDER;
            EtiquetaSaltoLinea.HorizontalAlignment = Element.ALIGN_CENTER;
            EtiquetaSaltoLinea.VerticalAlignment = Element.ALIGN_MIDDLE;
            EtiquetaSaltoLinea.FixedHeight = 8f;

            saltoLinea1.AddCell(EtiquetaSaltoLinea);

            document.Add(saltoLinea1);

            //Informacion de la solicitud         
            PdfPTable informacion = new PdfPTable(6);

            PdfPCell EtiquetaFechaCotizacion = new PdfPCell(new Phrase("Fecha y Hora: ", new Font(customfontbold, 7)));
            EtiquetaFechaCotizacion.Colspan = 1;
            EtiquetaFechaCotizacion.HorizontalAlignment = Element.ALIGN_RIGHT;
            EtiquetaFechaCotizacion.VerticalAlignment = Element.ALIGN_MIDDLE;
            EtiquetaFechaCotizacion.FixedHeight = 16f;
            EtiquetaFechaCotizacion.BackgroundColor = new BaseColor(225, 225, 225);
            EtiquetaFechaCotizacion.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell ValorFechaCotizacion = new PdfPCell(new Phrase("   " + solicitud.FechaHora, new Font(customfont, 7)));
            ValorFechaCotizacion.Colspan = 2;
            ValorFechaCotizacion.HorizontalAlignment = Element.ALIGN_LEFT;
            ValorFechaCotizacion.VerticalAlignment = Element.ALIGN_MIDDLE;
            ValorFechaCotizacion.FixedHeight = 16f;
            ValorFechaCotizacion.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell EtiquetaSolicitante = new PdfPCell(new Phrase("Solicitante:  ", new Font(customfontbold, 7)));
            EtiquetaSolicitante.Colspan = 1;
            EtiquetaSolicitante.HorizontalAlignment = Element.ALIGN_RIGHT;
            EtiquetaSolicitante.VerticalAlignment = Element.ALIGN_MIDDLE;
            EtiquetaSolicitante.FixedHeight = 16f;
            EtiquetaSolicitante.BackgroundColor = new BaseColor(225, 225, 225);
            EtiquetaSolicitante.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell ValorSolicitante = new PdfPCell(new Phrase("   " + solicitud.Solicitante, new Font(customfont, 7)));
            ValorSolicitante.Colspan = 2;
            ValorSolicitante.HorizontalAlignment = Element.ALIGN_LEFT;
            ValorSolicitante.VerticalAlignment = Element.ALIGN_MIDDLE;
            ValorSolicitante.FixedHeight = 16f;
            ValorSolicitante.BorderColor = new BaseColor(60, 66, 82);

            informacion.AddCell(EtiquetaFechaCotizacion);
            informacion.AddCell(ValorFechaCotizacion);
            informacion.AddCell(EtiquetaSolicitante);
            informacion.AddCell(ValorSolicitante);

            document.Add(informacion);

            //Salto Linea       
            PdfPTable saltoLinea2 = new PdfPTable(4);

            PdfPCell EtiquetaSaltoLinea2 = new PdfPCell(new Phrase(" ", new Font(customfontbold, 9)));
            EtiquetaSaltoLinea2.Colspan = 4;
            EtiquetaSaltoLinea2.Border = Rectangle.NO_BORDER;
            EtiquetaSaltoLinea2.HorizontalAlignment = Element.ALIGN_CENTER;
            EtiquetaSaltoLinea2.VerticalAlignment = Element.ALIGN_MIDDLE;
            EtiquetaSaltoLinea2.FixedHeight = 8f;

            saltoLinea2.AddCell(EtiquetaSaltoLinea2);

            document.Add(saltoLinea2);

            //Informacion de la Solicitud       
            PdfPTable datosGenerales = new PdfPTable(6);

            PdfPCell EtiquetaTipo = new PdfPCell(new Phrase("Tipo: ", new Font(customfontbold, 7)));
            EtiquetaTipo.Colspan = 1;
            EtiquetaTipo.HorizontalAlignment = Element.ALIGN_RIGHT;
            EtiquetaTipo.VerticalAlignment = Element.ALIGN_MIDDLE;
            EtiquetaTipo.FixedHeight = 16f;
            EtiquetaTipo.BackgroundColor = new BaseColor(225, 225, 225);
            EtiquetaTipo.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell ValorTipo = new PdfPCell(new Phrase("   " + solicitud.Tipo, new Font(customfont, 7)));
            ValorTipo.Colspan = 5;
            ValorTipo.HorizontalAlignment = Element.ALIGN_LEFT;
            ValorTipo.VerticalAlignment = Element.ALIGN_MIDDLE;
            ValorTipo.FixedHeight = 16f;
            ValorTipo.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell EtiquetaSubtipo = new PdfPCell(new Phrase("Subtipo: ", new Font(customfontbold, 7)));
            EtiquetaSubtipo.Colspan = 1;
            EtiquetaSubtipo.HorizontalAlignment = Element.ALIGN_RIGHT;
            EtiquetaSubtipo.VerticalAlignment = Element.ALIGN_MIDDLE;
            EtiquetaSubtipo.FixedHeight = 16f;
            EtiquetaSubtipo.BackgroundColor = new BaseColor(225, 225, 225);
            EtiquetaSubtipo.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell ValorSubtipo = new PdfPCell(new Phrase("   " + solicitud.Subtipo, new Font(customfont, 7)));
            ValorSubtipo.Colspan = 5;
            ValorSubtipo.HorizontalAlignment = Element.ALIGN_LEFT;
            ValorSubtipo.VerticalAlignment = Element.ALIGN_MIDDLE;
            ValorSubtipo.FixedHeight = 16f;
            ValorSubtipo.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell EtiquetaOP = new PdfPCell(new Phrase("OP: ", new Font(customfontbold, 7)));
            EtiquetaOP.Colspan = 1;
            EtiquetaOP.HorizontalAlignment = Element.ALIGN_RIGHT;
            EtiquetaOP.VerticalAlignment = Element.ALIGN_MIDDLE;
            EtiquetaOP.FixedHeight = 16f;
            EtiquetaOP.BackgroundColor = new BaseColor(225, 225, 225);
            EtiquetaOP.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell ValorOP = new PdfPCell(new Phrase("   " + solicitud.OP, new Font(customfont, 7)));
            ValorOP.Colspan = 2;
            ValorOP.HorizontalAlignment = Element.ALIGN_LEFT;
            ValorOP.VerticalAlignment = Element.ALIGN_MIDDLE;
            ValorOP.FixedHeight = 16f;
            ValorOP.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell EtiquetaMKT = new PdfPCell(new Phrase("MKT: ", new Font(customfontbold, 7)));
            EtiquetaMKT.Colspan = 1;
            EtiquetaMKT.HorizontalAlignment = Element.ALIGN_RIGHT;
            EtiquetaMKT.VerticalAlignment = Element.ALIGN_MIDDLE;
            EtiquetaMKT.FixedHeight = 16f;
            EtiquetaMKT.BackgroundColor = new BaseColor(225, 225, 225);
            EtiquetaMKT.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell ValorMKT = new PdfPCell(new Phrase("   " + solicitud.MKT, new Font(customfont, 7)));
            ValorMKT.Colspan = 2;
            ValorMKT.HorizontalAlignment = Element.ALIGN_LEFT;
            ValorMKT.VerticalAlignment = Element.ALIGN_MIDDLE;
            ValorMKT.FixedHeight = 16f;
            ValorMKT.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell EtiquetaNombre = new PdfPCell(new Phrase("Nombre Campaña: ", new Font(customfontbold, 7)));
            EtiquetaNombre.Colspan = 1;
            EtiquetaNombre.HorizontalAlignment = Element.ALIGN_RIGHT;
            EtiquetaNombre.VerticalAlignment = Element.ALIGN_MIDDLE;
            EtiquetaNombre.FixedHeight = 16f;
            EtiquetaNombre.BackgroundColor = new BaseColor(225, 225, 225);
            EtiquetaNombre.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell ValorNombre = new PdfPCell(new Phrase("   " + solicitud.Nombre, new Font(customfont, 7)));
            ValorNombre.Colspan = 5;
            ValorNombre.HorizontalAlignment = Element.ALIGN_LEFT;
            ValorNombre.VerticalAlignment = Element.ALIGN_MIDDLE;
            ValorNombre.FixedHeight = 16f;
            ValorNombre.BorderColor = new BaseColor(60, 66, 82);

            datosGenerales.AddCell(EtiquetaTipo);
            datosGenerales.AddCell(ValorTipo);

            datosGenerales.AddCell(EtiquetaSubtipo);
            datosGenerales.AddCell(ValorSubtipo);

            datosGenerales.AddCell(EtiquetaOP);
            datosGenerales.AddCell(ValorOP);
            datosGenerales.AddCell(EtiquetaMKT);
            datosGenerales.AddCell(ValorMKT);

            datosGenerales.AddCell(EtiquetaNombre);
            datosGenerales.AddCell(ValorNombre);

            document.Add(datosGenerales);

            //Salto Linea       
            PdfPTable saltoLinea3 = new PdfPTable(4);

            PdfPCell EtiquetaSaltoLinea3 = new PdfPCell(new Phrase(" ", new Font(customfontbold, 9)));
            EtiquetaSaltoLinea3.Colspan = 4;
            EtiquetaSaltoLinea3.Border = Rectangle.NO_BORDER;
            EtiquetaSaltoLinea3.HorizontalAlignment = Element.ALIGN_CENTER;
            EtiquetaSaltoLinea3.VerticalAlignment = Element.ALIGN_MIDDLE;
            EtiquetaSaltoLinea3.FixedHeight = 8f;

            saltoLinea3.AddCell(EtiquetaSaltoLinea3);

            document.Add(saltoLinea3);

            //Url Externo  
            PdfPTable urlsExterno = new PdfPTable(6);

            PdfPCell EtiquetaCampoUrls = new PdfPCell(new Phrase("URL's de Tipificación", new Font(customfontbold, 9, 0, new BaseColor(255, 255, 255))));
            EtiquetaCampoUrls.Colspan = 6;
            EtiquetaCampoUrls.HorizontalAlignment = Element.ALIGN_CENTER;
            EtiquetaCampoUrls.VerticalAlignment = Element.ALIGN_MIDDLE;
            EtiquetaCampoUrls.FixedHeight = 16f;
            EtiquetaCampoUrls.BackgroundColor = new BaseColor(60, 66, 82);
            EtiquetaCampoUrls.BorderColor = new BaseColor(60, 66, 82);

            urlsExterno.AddCell(EtiquetaCampoUrls);

            foreach (UrlExternosSolicitud urls in ListadoUrlExterno)
            {
                PdfPCell ValorTipoUrlDetalle = new PdfPCell(new Phrase(contadorUrls.ToString(), new Font(customfont, 7)));
                ValorTipoUrlDetalle.Colspan = 1;
                ValorTipoUrlDetalle.HorizontalAlignment = Element.ALIGN_CENTER;
                ValorTipoUrlDetalle.VerticalAlignment = Element.ALIGN_MIDDLE;
                ValorTipoUrlDetalle.FixedHeight = 23f;
                ValorTipoUrlDetalle.BorderColor = new BaseColor(60, 66, 82);

                PdfPCell ValorDetalleUrlDetalle = new PdfPCell(new Phrase(urls.url.ToString(), new Font(customfont, 7)));
                ValorDetalleUrlDetalle.Colspan = 5;
                ValorDetalleUrlDetalle.HorizontalAlignment = Element.ALIGN_LEFT;
                ValorDetalleUrlDetalle.VerticalAlignment = Element.ALIGN_MIDDLE;
                ValorDetalleUrlDetalle.FixedHeight = 23f;
                ValorDetalleUrlDetalle.BorderColor = new BaseColor(60, 66, 82);

                urlsExterno.AddCell(ValorTipoUrlDetalle);
                urlsExterno.AddCell(ValorDetalleUrlDetalle);
            }

            if (ListadoUrlExterno.Count > 0)
            {
                document.Add(urlsExterno);

                //Salto Linea       
                PdfPTable saltoLinea5 = new PdfPTable(4);

                PdfPCell EtiquetaSaltoLinea5 = new PdfPCell(new Phrase(" ", new Font(customfontbold, 9)));
                EtiquetaSaltoLinea5.Colspan = 4;
                EtiquetaSaltoLinea5.Border = Rectangle.NO_BORDER;
                EtiquetaSaltoLinea5.HorizontalAlignment = Element.ALIGN_CENTER;
                EtiquetaSaltoLinea5.VerticalAlignment = Element.ALIGN_MIDDLE;
                EtiquetaSaltoLinea5.FixedHeight = 8f;

                saltoLinea5.AddCell(EtiquetaSaltoLinea5);

                document.Add(saltoLinea5);
            }

            document.Close();
        }

        //Reporte Solicitud Html y Tags
        public void GenerarPDFHtmlTags(int? id)
        {
            int contadorUrls = 1;

            //Obtener los datos de la Solicitud
            DatosSolicitudCliente solicitud = db.DatosSolicitudCliente(id).First();

            //Adjuntos
            List<AdjuntosSolicitud> ListadoAdjuntos = db.AdjuntosSolicitud(id).ToList();

            //Url Externo 
            List<UrlExternosSolicitud> ListadoUrlExterno = db.UrlExternosSolicitud(id).ToList();

            //Fonts para el PDF
            //Regular
            string fontpath = Server.MapPath("~/Content/fonts/ubuntu/Raleway-Regular.ttf");
            BaseFont customfont = BaseFont.CreateFont(fontpath, BaseFont.CP1252, BaseFont.EMBEDDED);

            //Bold
            string fontpathbold = Server.MapPath("~/Content/fonts/ubuntu/Raleway-Bold.ttf");
            BaseFont customfontbold = BaseFont.CreateFont(fontpathbold, BaseFont.CP1252, BaseFont.EMBEDDED);

            //ruta imagen del formulario
            String FilePath = Server.MapPath("~/Content/img/LogoPPMPDF.png");
            var physicalPath = Server.MapPath("~/Content/img/LogoPPMPDF.png");

            //anio Actual
            var anioActual = System.DateTime.Now.Year.ToString();

            string basePath = ConfigurationManager.AppSettings["RepositorioDocumentos"];
            string rutaArchivos = basePath + "\\GESTION_PPM\\ADJUNTOS_SOLICITUDES\\" + anioActual + '\\' + solicitud.Id.ToString();

            // En caso de que no exista el directorio, crearlo.
            bool directorio = Directory.Exists(rutaArchivos);
            if (!directorio)
                Directory.CreateDirectory(rutaArchivos);

            //ruta del archivo PDF que se va a crear
            string rutaDocumentos = rutaArchivos + "\\Solicitud-" + solicitud.CodigoSolicitud.ToString() + ".pdf";

            //Crear el formulario PDF
            FileStream fs = new FileStream(rutaDocumentos, FileMode.Create);
            Document document = new Document(iTextSharp.text.PageSize.A4, -20, -20, 0, 0);

            //Document document = new Document(iTextSharp.text.PageSize.A4);
            PdfWriter pw = PdfWriter.GetInstance(document, fs);

            //Abrir archivo para su edicion
            document.Open();

            //Salto de linea 
            document.Add(new Paragraph(" "));

            //Salto de linea 
            document.Add(new Paragraph(" "));

            //Cabecera del Cotizador         
            PdfPTable cabecera = new PdfPTable(7);
            cabecera.DefaultCell.Border = Rectangle.NO_BORDER;

            //Imagen
            iTextSharp.text.Image imagen = iTextSharp.text.Image.GetInstance(FilePath);
            imagen.ScalePercent(10);

            PdfPCell logocell = new PdfPCell(imagen, true);
            logocell.BackgroundColor = new BaseColor(60, 66, 82);
            logocell.Colspan = 2;
            logocell.Border = Rectangle.NO_BORDER;
            logocell.HorizontalAlignment = Element.ALIGN_LEFT;
            logocell.VerticalAlignment = Element.ALIGN_MIDDLE;

            //Titulo
            PdfPCell TituloCotizador = new PdfPCell(new Phrase("SOLICITUD " + solicitud.Tipo.ToUpper() + " N° " + solicitud.CodigoSolicitud + "      ", new Font(customfontbold, 9, 0, new BaseColor(255, 255, 255))));
            TituloCotizador.Colspan = 5;
            TituloCotizador.Border = Rectangle.NO_BORDER;
            TituloCotizador.HorizontalAlignment = Element.ALIGN_RIGHT;
            TituloCotizador.VerticalAlignment = Element.ALIGN_MIDDLE;
            TituloCotizador.BackgroundColor = new BaseColor(60, 66, 82);
            TituloCotizador.FixedHeight = 50f;

            cabecera.AddCell(logocell);
            cabecera.AddCell(TituloCotizador);

            document.Add(cabecera);

            //Salto Linea      
            PdfPTable saltoLinea1 = new PdfPTable(4);

            PdfPCell EtiquetaSaltoLinea = new PdfPCell(new Phrase(" ", new Font(customfontbold, 9)));
            EtiquetaSaltoLinea.Colspan = 4;
            EtiquetaSaltoLinea.Border = Rectangle.NO_BORDER;
            EtiquetaSaltoLinea.HorizontalAlignment = Element.ALIGN_CENTER;
            EtiquetaSaltoLinea.VerticalAlignment = Element.ALIGN_MIDDLE;
            EtiquetaSaltoLinea.FixedHeight = 8f;

            saltoLinea1.AddCell(EtiquetaSaltoLinea);

            document.Add(saltoLinea1);

            //Informacion de la solicitud         
            PdfPTable informacion = new PdfPTable(6);

            PdfPCell EtiquetaFechaCotizacion = new PdfPCell(new Phrase("Fecha y Hora: ", new Font(customfontbold, 7)));
            EtiquetaFechaCotizacion.Colspan = 1;
            EtiquetaFechaCotizacion.HorizontalAlignment = Element.ALIGN_RIGHT;
            EtiquetaFechaCotizacion.VerticalAlignment = Element.ALIGN_MIDDLE;
            EtiquetaFechaCotizacion.FixedHeight = 16f;
            EtiquetaFechaCotizacion.BackgroundColor = new BaseColor(225, 225, 225);
            EtiquetaFechaCotizacion.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell ValorFechaCotizacion = new PdfPCell(new Phrase("   " + solicitud.FechaHora, new Font(customfont, 7)));
            ValorFechaCotizacion.Colspan = 2;
            ValorFechaCotizacion.HorizontalAlignment = Element.ALIGN_LEFT;
            ValorFechaCotizacion.VerticalAlignment = Element.ALIGN_MIDDLE;
            ValorFechaCotizacion.FixedHeight = 16f;
            ValorFechaCotizacion.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell EtiquetaSolicitante = new PdfPCell(new Phrase("Solicitante:  ", new Font(customfontbold, 7)));
            EtiquetaSolicitante.Colspan = 1;
            EtiquetaSolicitante.HorizontalAlignment = Element.ALIGN_RIGHT;
            EtiquetaSolicitante.VerticalAlignment = Element.ALIGN_MIDDLE;
            EtiquetaSolicitante.FixedHeight = 16f;
            EtiquetaSolicitante.BackgroundColor = new BaseColor(225, 225, 225);
            EtiquetaSolicitante.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell ValorSolicitante = new PdfPCell(new Phrase("   " + solicitud.Solicitante, new Font(customfont, 7)));
            ValorSolicitante.Colspan = 2;
            ValorSolicitante.HorizontalAlignment = Element.ALIGN_LEFT;
            ValorSolicitante.VerticalAlignment = Element.ALIGN_MIDDLE;
            ValorSolicitante.FixedHeight = 16f;
            ValorSolicitante.BorderColor = new BaseColor(60, 66, 82);

            informacion.AddCell(EtiquetaFechaCotizacion);
            informacion.AddCell(ValorFechaCotizacion);
            informacion.AddCell(EtiquetaSolicitante);
            informacion.AddCell(ValorSolicitante);

            document.Add(informacion);

            //Salto Linea       
            PdfPTable saltoLinea2 = new PdfPTable(4);

            PdfPCell EtiquetaSaltoLinea2 = new PdfPCell(new Phrase(" ", new Font(customfontbold, 9)));
            EtiquetaSaltoLinea2.Colspan = 4;
            EtiquetaSaltoLinea2.Border = Rectangle.NO_BORDER;
            EtiquetaSaltoLinea2.HorizontalAlignment = Element.ALIGN_CENTER;
            EtiquetaSaltoLinea2.VerticalAlignment = Element.ALIGN_MIDDLE;
            EtiquetaSaltoLinea2.FixedHeight = 8f;

            saltoLinea2.AddCell(EtiquetaSaltoLinea2);

            document.Add(saltoLinea2);

            //Informacion de la Solicitud       
            PdfPTable datosGenerales = new PdfPTable(6);

            PdfPCell EtiquetaTipo = new PdfPCell(new Phrase("Tipo: ", new Font(customfontbold, 7)));
            EtiquetaTipo.Colspan = 1;
            EtiquetaTipo.HorizontalAlignment = Element.ALIGN_RIGHT;
            EtiquetaTipo.VerticalAlignment = Element.ALIGN_MIDDLE;
            EtiquetaTipo.FixedHeight = 16f;
            EtiquetaTipo.BackgroundColor = new BaseColor(225, 225, 225);
            EtiquetaTipo.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell ValorTipo = new PdfPCell(new Phrase("   " + solicitud.Tipo, new Font(customfont, 7)));
            ValorTipo.Colspan = 5;
            ValorTipo.HorizontalAlignment = Element.ALIGN_LEFT;
            ValorTipo.VerticalAlignment = Element.ALIGN_MIDDLE;
            ValorTipo.FixedHeight = 16f;
            ValorTipo.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell EtiquetaSubtipo = new PdfPCell(new Phrase("Subtipo: ", new Font(customfontbold, 7)));
            EtiquetaSubtipo.Colspan = 1;
            EtiquetaSubtipo.HorizontalAlignment = Element.ALIGN_RIGHT;
            EtiquetaSubtipo.VerticalAlignment = Element.ALIGN_MIDDLE;
            EtiquetaSubtipo.FixedHeight = 16f;
            EtiquetaSubtipo.BackgroundColor = new BaseColor(225, 225, 225);
            EtiquetaSubtipo.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell ValorSubtipo = new PdfPCell(new Phrase("   " + solicitud.Subtipo, new Font(customfont, 7)));
            ValorSubtipo.Colspan = 5;
            ValorSubtipo.HorizontalAlignment = Element.ALIGN_LEFT;
            ValorSubtipo.VerticalAlignment = Element.ALIGN_MIDDLE;
            ValorSubtipo.FixedHeight = 16f;
            ValorSubtipo.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell EtiquetaOP = new PdfPCell(new Phrase("OP: ", new Font(customfontbold, 7)));
            EtiquetaOP.Colspan = 1;
            EtiquetaOP.HorizontalAlignment = Element.ALIGN_RIGHT;
            EtiquetaOP.VerticalAlignment = Element.ALIGN_MIDDLE;
            EtiquetaOP.FixedHeight = 16f;
            EtiquetaOP.BackgroundColor = new BaseColor(225, 225, 225);
            EtiquetaOP.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell ValorOP = new PdfPCell(new Phrase("   " + solicitud.OP, new Font(customfont, 7)));
            ValorOP.Colspan = 2;
            ValorOP.HorizontalAlignment = Element.ALIGN_LEFT;
            ValorOP.VerticalAlignment = Element.ALIGN_MIDDLE;
            ValorOP.FixedHeight = 16f;
            ValorOP.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell EtiquetaMKT = new PdfPCell(new Phrase("MKT: ", new Font(customfontbold, 7)));
            EtiquetaMKT.Colspan = 1;
            EtiquetaMKT.HorizontalAlignment = Element.ALIGN_RIGHT;
            EtiquetaMKT.VerticalAlignment = Element.ALIGN_MIDDLE;
            EtiquetaMKT.FixedHeight = 16f;
            EtiquetaMKT.BackgroundColor = new BaseColor(225, 225, 225);
            EtiquetaMKT.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell ValorMKT = new PdfPCell(new Phrase("   " + solicitud.MKT, new Font(customfont, 7)));
            ValorMKT.Colspan = 2;
            ValorMKT.HorizontalAlignment = Element.ALIGN_LEFT;
            ValorMKT.VerticalAlignment = Element.ALIGN_MIDDLE;
            ValorMKT.FixedHeight = 16f;
            ValorMKT.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell EtiquetaNombre = new PdfPCell(new Phrase("Nombre Campaña: ", new Font(customfontbold, 7)));
            EtiquetaNombre.Colspan = 1;
            EtiquetaNombre.HorizontalAlignment = Element.ALIGN_RIGHT;
            EtiquetaNombre.VerticalAlignment = Element.ALIGN_MIDDLE;
            EtiquetaNombre.FixedHeight = 16f;
            EtiquetaNombre.BackgroundColor = new BaseColor(225, 225, 225);
            EtiquetaNombre.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell ValorNombre = new PdfPCell(new Phrase("   " + solicitud.Nombre, new Font(customfont, 7)));
            ValorNombre.Colspan = 5;
            ValorNombre.HorizontalAlignment = Element.ALIGN_LEFT;
            ValorNombre.VerticalAlignment = Element.ALIGN_MIDDLE;
            ValorNombre.FixedHeight = 16f;
            ValorNombre.BorderColor = new BaseColor(60, 66, 82);

            datosGenerales.AddCell(EtiquetaTipo);
            datosGenerales.AddCell(ValorTipo);

            datosGenerales.AddCell(EtiquetaSubtipo);
            datosGenerales.AddCell(ValorSubtipo);

            datosGenerales.AddCell(EtiquetaOP);
            datosGenerales.AddCell(ValorOP);
            datosGenerales.AddCell(EtiquetaMKT);
            datosGenerales.AddCell(ValorMKT);

            datosGenerales.AddCell(EtiquetaNombre);
            datosGenerales.AddCell(ValorNombre);

            document.Add(datosGenerales);

            //Salto Linea       
            PdfPTable saltoLinea3 = new PdfPTable(4);

            PdfPCell EtiquetaSaltoLinea3 = new PdfPCell(new Phrase(" ", new Font(customfontbold, 9)));
            EtiquetaSaltoLinea3.Colspan = 4;
            EtiquetaSaltoLinea3.Border = Rectangle.NO_BORDER;
            EtiquetaSaltoLinea3.HorizontalAlignment = Element.ALIGN_CENTER;
            EtiquetaSaltoLinea3.VerticalAlignment = Element.ALIGN_MIDDLE;
            EtiquetaSaltoLinea3.FixedHeight = 8f;

            saltoLinea3.AddCell(EtiquetaSaltoLinea3);

            document.Add(saltoLinea3);

            //Url Externo  
            PdfPTable urlsExterno = new PdfPTable(6);

            PdfPCell EtiquetaCampoUrls = new PdfPCell(new Phrase("URL's de Tipificación", new Font(customfontbold, 9, 0, new BaseColor(255, 255, 255))));
            EtiquetaCampoUrls.Colspan = 6;
            EtiquetaCampoUrls.HorizontalAlignment = Element.ALIGN_CENTER;
            EtiquetaCampoUrls.VerticalAlignment = Element.ALIGN_MIDDLE;
            EtiquetaCampoUrls.FixedHeight = 16f;
            EtiquetaCampoUrls.BackgroundColor = new BaseColor(60, 66, 82);
            EtiquetaCampoUrls.BorderColor = new BaseColor(60, 66, 82);

            urlsExterno.AddCell(EtiquetaCampoUrls);

            foreach (UrlExternosSolicitud urls in ListadoUrlExterno)
            {
                PdfPCell ValorTipoUrlDetalle = new PdfPCell(new Phrase(contadorUrls.ToString(), new Font(customfont, 7)));
                ValorTipoUrlDetalle.Colspan = 1;
                ValorTipoUrlDetalle.HorizontalAlignment = Element.ALIGN_CENTER;
                ValorTipoUrlDetalle.VerticalAlignment = Element.ALIGN_MIDDLE;
                ValorTipoUrlDetalle.FixedHeight = 23f;
                ValorTipoUrlDetalle.BorderColor = new BaseColor(60, 66, 82);

                PdfPCell ValorDetalleUrlDetalle = new PdfPCell(new Phrase(urls.url.ToString(), new Font(customfont, 7)));
                ValorDetalleUrlDetalle.Colspan = 5;
                ValorDetalleUrlDetalle.HorizontalAlignment = Element.ALIGN_LEFT;
                ValorDetalleUrlDetalle.VerticalAlignment = Element.ALIGN_MIDDLE;
                ValorDetalleUrlDetalle.FixedHeight = 23f;
                ValorDetalleUrlDetalle.BorderColor = new BaseColor(60, 66, 82);

                urlsExterno.AddCell(ValorTipoUrlDetalle);
                urlsExterno.AddCell(ValorDetalleUrlDetalle);
            }

            if (ListadoUrlExterno.Count > 0)
            {
                document.Add(urlsExterno);

                //Salto Linea       
                PdfPTable saltoLinea5 = new PdfPTable(4);

                PdfPCell EtiquetaSaltoLinea5 = new PdfPCell(new Phrase(" ", new Font(customfontbold, 9)));
                EtiquetaSaltoLinea5.Colspan = 4;
                EtiquetaSaltoLinea5.Border = Rectangle.NO_BORDER;
                EtiquetaSaltoLinea5.HorizontalAlignment = Element.ALIGN_CENTER;
                EtiquetaSaltoLinea5.VerticalAlignment = Element.ALIGN_MIDDLE;
                EtiquetaSaltoLinea5.FixedHeight = 8f;

                saltoLinea5.AddCell(EtiquetaSaltoLinea5);

                document.Add(saltoLinea5);
            }

            //Archivos Adjuntos
            PdfPTable camposAdjuntos = new PdfPTable(6);

            PdfPCell EtiquetaPiezaReferencia = new PdfPCell(new Phrase("Archivos Arjuntos", new Font(customfontbold, 9, 0, new BaseColor(255, 255, 255))));
            EtiquetaPiezaReferencia.Colspan = 6;
            EtiquetaPiezaReferencia.HorizontalAlignment = Element.ALIGN_CENTER;
            EtiquetaPiezaReferencia.VerticalAlignment = Element.ALIGN_MIDDLE;
            EtiquetaPiezaReferencia.FixedHeight = 16f;
            EtiquetaPiezaReferencia.BackgroundColor = new BaseColor(60, 66, 82);
            EtiquetaPiezaReferencia.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell EtiquetaNumero = new PdfPCell(new Phrase("Tipo", new Font(customfontbold, 7)));
            EtiquetaNumero.Colspan = 2;
            EtiquetaNumero.HorizontalAlignment = Element.ALIGN_CENTER;
            EtiquetaNumero.VerticalAlignment = Element.ALIGN_MIDDLE;
            EtiquetaNumero.FixedHeight = 16f;
            EtiquetaNumero.BackgroundColor = new BaseColor(225, 225, 225);
            EtiquetaNumero.BorderColor = new BaseColor(60, 66, 82);

            PdfPCell EtiquetaNombreArchivo = new PdfPCell(new Phrase("Nombre Archivo", new Font(customfontbold, 7)));
            EtiquetaNombreArchivo.Colspan = 4;
            EtiquetaNombreArchivo.HorizontalAlignment = Element.ALIGN_LEFT;
            EtiquetaNombreArchivo.VerticalAlignment = Element.ALIGN_MIDDLE;
            EtiquetaNombreArchivo.FixedHeight = 16f;
            EtiquetaNombreArchivo.BackgroundColor = new BaseColor(225, 225, 225);
            EtiquetaNombreArchivo.BorderColor = new BaseColor(60, 66, 82);

            camposAdjuntos.AddCell(EtiquetaPiezaReferencia);

            camposAdjuntos.AddCell(EtiquetaNumero);
            camposAdjuntos.AddCell(EtiquetaNombreArchivo);

            foreach (AdjuntosSolicitud adjunto in ListadoAdjuntos)
            {
                PdfPCell EtiquetaCampoNumero = new PdfPCell(new Phrase(adjunto.tipo.ToString(), new Font(customfont, 7)));
                EtiquetaCampoNumero.Colspan = 2;
                EtiquetaCampoNumero.HorizontalAlignment = Element.ALIGN_CENTER;
                EtiquetaCampoNumero.VerticalAlignment = Element.ALIGN_MIDDLE;
                EtiquetaCampoNumero.FixedHeight = 23f;
                EtiquetaCampoNumero.BorderColor = new BaseColor(60, 66, 82);

                PdfPCell ValorNombreAdjunto = new PdfPCell(new Phrase(adjunto.nombre_adjunto, new Font(customfont, 7)));
                ValorNombreAdjunto.Colspan = 4;
                ValorNombreAdjunto.HorizontalAlignment = Element.ALIGN_LEFT;
                ValorNombreAdjunto.VerticalAlignment = Element.ALIGN_MIDDLE;
                ValorNombreAdjunto.FixedHeight = 23f;
                ValorNombreAdjunto.BorderColor = new BaseColor(60, 66, 82);

                camposAdjuntos.AddCell(EtiquetaCampoNumero);
                camposAdjuntos.AddCell(ValorNombreAdjunto);

            }

            if (ListadoAdjuntos.Count > 0)
            {
                document.Add(camposAdjuntos);

                //Salto Linea       
                PdfPTable saltoLinea6 = new PdfPTable(4);

                PdfPCell EtiquetaSaltoLinea6 = new PdfPCell(new Phrase(" ", new Font(customfontbold, 9)));
                EtiquetaSaltoLinea6.Colspan = 4;
                EtiquetaSaltoLinea6.Border = Rectangle.NO_BORDER;
                EtiquetaSaltoLinea6.HorizontalAlignment = Element.ALIGN_CENTER;
                EtiquetaSaltoLinea6.VerticalAlignment = Element.ALIGN_MIDDLE;
                EtiquetaSaltoLinea6.FixedHeight = 8f;

                saltoLinea6.AddCell(EtiquetaSaltoLinea6);

                document.Add(saltoLinea6);
            }


            document.Close();
        }

        #endregion

        //Obtener Portales para mantenimiento
        public JsonResult _GetOpcionesPortales(string searchTerm)
        {
            List<MultiSelectJQueryUi> items = new List<MultiSelectJQueryUi>();

            items = CatalogoEntity.ObtenerListadoCatalogosByCodigo("MRC-01").Where(o => o.Text != "MULTIMARCA")
            .Select(o => new MultiSelectJQueryUi(Convert.ToUInt32(o.Value), o.Text, o.Text)).ToList();
            return Json(items, JsonRequestBehavior.AllowGet);

        }
    }
}

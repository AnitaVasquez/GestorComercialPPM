using GestionPPM.Entidades.Metodos;
using GestionPPM.Entidades.Modelo;
using GestionPPM.Repositorios;
using Newtonsoft.Json;
using Seguridad.Helper;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Web.Mvc;

namespace TemplateInicial.Controllers
{
    [Autenticado]
    public class OrganigramaController : BaseAppController
    {
        // GET: Organigrama
        public ActionResult Index()
        {
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

        // GET: Organigrama/Details/5
        public ActionResult Details(int id)
        {
            return View();
        }

        public ActionResult _SeleccionTipoOrganigrama()
        {
            ViewBag.TituloModal = "Seleccionar el tipo de Organigrama";
            return PartialView();
        }

        public ActionResult Organigrama(int? id)
        {
            //Controlar permisos
            var user = ViewData["usuario"] = System.Web.HttpContext.Current.Session["usuario"];
            var usuario = int.Parse(user.ToString());
            string nombreControlador = ControllerContext.RouteData.Values["controller"].ToString();
            ViewBag.NombreControlador = nombreControlador;

            ViewBag.AccionesUsuario = ManejoPermisosEntity.ListadoAccionesCatalogoUsuario(usuario, nombreControlador);

            //Obtener Acciones del controlador
            ViewBag.AccionesControlador = GetMetodosControlador(nombreControlador);

           
            int usuarioPrincipal = int.Parse(ParametrosSistemaEntity.ConsultarParametros(2).valor);

            var listadoUsuarios = new List<ListadoUsuarios>();

            var tipoOrganigrama = OrganigramaEntity.ConsultarTipoOrganigrama(id.Value);
            string nombreOrganigrama = string.Empty;

            bool esInicial = true;

            switch (tipoOrganigrama.Codigo)
            {
                case "TIPO-ORG-INTERNOS":
                    listadoUsuarios = UsuarioEntity.ListarUsuariosOrganigrama(new List<int> { usuarioPrincipal }, "INTERNO");
                    nombreOrganigrama = "Organigrama Usuarios Internos";
                    esInicial = OrganigramaEntity.EsOrganigramaNuevo(id);
                    break;
                case "TIPO-ORG-EXTERNOS":
                    listadoUsuarios = UsuarioEntity.ListarUsuariosOrganigrama(new List<int> { usuarioPrincipal }, "EXTERNO");
                    nombreOrganigrama = "Organigrama Usuarios Externos";
                    esInicial = OrganigramaEntity.EsOrganigramaNuevo(id);
                    break;
                case "TIPO-ORG-GENERICO": // Permite crear múltiples organigramas
                    listadoUsuarios = UsuarioEntity.ListarUsuariosOrganigrama(new List<int> { usuarioPrincipal });
                    nombreOrganigrama = "Organigrama genérico";
                    esInicial = OrganigramaEntity.EsOrganigramaNuevo(id);//OrganigramaEntity.EsOrganigramaNuevo(id, true);
                    break;
                default:
                    listadoUsuarios = UsuarioEntity.ListarUsuariosOrganigrama(new List<int> { usuarioPrincipal });
                    break;
            }

            var model = new Organigrama { Estado = true, EmpresaID = 1, Codigo = "ORG-" + Guid.NewGuid(), TipoOrganigramaID = id.Value, Nombre = nombreOrganigrama };

            ViewBag.esInicial = esInicial;

            if (esInicial)
            {
                ViewBag.Organigrama = JsonConvert.SerializeObject(UsuarioEntity.GetUsuarioAdministradorPrincipalOrganigrama(usuarioPrincipal));
            }
            else
            {
                ViewBag.Organigrama = OrganigramaEntity.GetEstructuraOrganigrama(id.Value);
                model = OrganigramaEntity.ConsultarOrganigramaPrincipal(id);

                var organigramaParcial = JsonConvert.DeserializeObject<List<OrganigramaParcial>>(OrganigramaEntity.GetEstructuraOrganigrama(id.Value));

                var ids = organigramaParcial != null ? organigramaParcial.Select(s => s.id).ToList() : new List<int>();

                if (organigramaParcial == null) {
                    ViewBag.Organigrama = JsonConvert.SerializeObject(UsuarioEntity.GetUsuarioAdministradorPrincipalOrganigrama(usuarioPrincipal));
                }
                
                listadoUsuarios = listadoUsuarios.Where(s => !ids.Contains(s.Id)).ToList();
            }

            ViewBag.ListadoUsuarios = listadoUsuarios;

            return View(model);
        }

        // GET: Organigrama/Create
        public ActionResult Create()
        {
            return View();
        }

        // POST: Organigrama/Create
        [HttpPost]
        public ActionResult Create(Organigrama organigrama)
        {
            RespuestaTransaccion resultado = new RespuestaTransaccion();
            try
            {
                if (organigrama.IDOrganigrama == 0)
                    resultado = OrganigramaEntity.CrearOrganigrama(organigrama);
                else
                    resultado = OrganigramaEntity.EditarOrganigrama(organigrama);

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
            RespuestaTransaccion resultado = OrganigramaEntity.EliminarOrganigrama(id);

            //return RedirectToAction("Index");
            return Json(new { Resultado = resultado }, JsonRequestBehavior.AllowGet);
        }

        // GET: Organigrama/Edit/5
        //public ActionResult Edit(int id)
        //{
        //    return View();
        //}

        // POST: Organigrama/Edit/5
       // [HttpPost]
        //public ActionResult Edit(int id, FormCollection collection)
        //{
        //    try
        //    {
        //        // TODO: Add update logic here

        //        return RedirectToAction("Index");
        //    }
        //    catch
        //    {
        //        return View();
        //    }
        //}

        // GET: Organigrama/Delete/5
        public ActionResult Delete(int id)
        {
            return View();
        }

        // POST: Organigrama/Delete/5
        [HttpPost]
        public ActionResult Delete(int id, FormCollection collection)
        {
            try
            {
                // TODO: Add delete logic here

                return RedirectToAction("Index");
            }
            catch
            {
                return View();
            }
        }

        public ActionResult DescargaOrganigrama()
        {

            try
            {
                return Json(new { Resultado = new RespuestaTransaccion { Estado = true, Respuesta = Mensajes.MensajeTransaccionExitosa } }, JsonRequestBehavior.AllowGet);

            }
            catch (Exception ex)
            {
                return Json(new { Resultado = new RespuestaTransaccion { Estado = false, Respuesta = ex.Message } }, JsonRequestBehavior.AllowGet);
            }


        }
    }
}

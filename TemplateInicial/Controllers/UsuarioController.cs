using GestionPPM.Entidades.Metodos;
using GestionPPM.Entidades.Modelo;
using GestionPPM.Repositorios;
using NLog;
using OfficeOpenXml;
using Seguridad.Helper;
using System; 
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks; 
using System.Web.Mvc;

namespace TemplateInicial.Controllers
{
    [Autenticado]
    public class UsuarioController : BaseAppController
    {
        private const string XlsxContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
        private static readonly Logger Log = LogManager.GetCurrentClassLogger();

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
            ViewBag.NombreListado = Etiquetas.TituloGridUsuario;

            //Controlar permisos
            var user = ViewData["usuario"] = System.Web.HttpContext.Current.Session["usuario"];
            var usuario = int.Parse(user.ToString());
            string nombreControlador = ControllerContext.RouteData.Values["controller"].ToString();
            ViewBag.NombreControlador = nombreControlador;

            ViewBag.AccionesUsuario = ManejoPermisosEntity.ListadoAccionesCatalogoUsuario(usuario, nombreControlador);

            //Obtener Acciones del controlador
            ViewBag.AccionesControlador = GetMetodosControlador(nombreControlador);

            //Búsqueda
            var listado = UsuarioEntity.ListarUsuarios();

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

        public ActionResult CambiarClave()
        {
            CambiarClave cambiaClave = new CambiarClave();
            return View();
        }

        [HttpPost]
        public ActionResult CambiarClave(CambiarClave usuario)
        {
            RespuestaTransaccion resultado = UsuarioEntity.CambiarClave(usuario);

            //Almacenar en una variable de sesion
            Session["Resultado"] = resultado.Respuesta;
            Session["Estado"] = resultado.Estado.ToString();

            if (resultado.Estado.ToString() == "True")
            {
                return RedirectToAction("Login", "Ingreso");
            }
            else
            {
                ViewBag.Resultado = resultado.Respuesta;
                ViewBag.Estado = resultado.Estado.ToString();
                Session["Resultado"] = "";
                Session["Estado"] = "";
                return View(usuario);
            }
        }

        public ActionResult Create()
        {
            try
            {
                //Tipo de Usuario
                ViewBag.tipoUsuarioId = 110;                  

                //Listado Tipo Usuario  
                var tipoUsuario = CatalogoEntity.ListadoCatalogosPorCodigo("TUS-01");
                ViewBag.listadoTipoUsuario = tipoUsuario; 

                //Listado Clientes  
                var clientes = ClienteEntity.ObtenerListadoClientes();
                ViewBag.listadoClientes = clientes;  

                //Listado Rol  
                var roles = RolEntity.ObtenerListadoRoles();
                ViewBag.listadoRoles = roles; 

                //Listado Area Departamento  
                var areasDepartamentos = CatalogoEntity.ListadoCatalogosPorCodigo("DTO-01");
                ViewBag.ListadoAreasDepartamentos = areasDepartamentos; 

                //Listado Rol  
                var cargos = CatalogoEntity.ListadoCatalogosPorCodigo("CRG-01");
                ViewBag.ListadoCargos = cargos; 

                //Listado Pais  
                var paises = CatalogoEntity.ListadoCatalogosPorCodigo("PAI-01");
                ViewBag.listadoPaises = paises; 

                //Listado Ciudades   
                var ciudades = CatalogoEntity.ListadoCatalogosPorCodigoId("CIUDAD", 0);
                ViewBag.listadoCiudades = ciudades; 

                return View();
            }
            catch (Exception ex)
            {
                Log.Error(ex.Message.ToString(),"Started");
                throw;
            }
            
        }

        [HttpPost] 
        public ActionResult Create(UsuarioCE usuario)
        {
            //Tipo de Usuario
            ViewBag.tipoUsuarioId = usuario.tipo_usuario;

            //Listado Tipo Usuario  
            var tipoUsuario = CatalogoEntity.ListadoCatalogosPorCodigo("TUS-01");
            ViewBag.listadoTipoUsuario = tipoUsuario;

            //Listado Clientes  
            var clientes = ClienteEntity.ObtenerListadoClientes();
            ViewBag.listadoClientes = clientes;

            //Listado Rol  
            var roles = RolEntity.ObtenerListadoRoles();
            ViewBag.listadoRoles = roles;

            //Listado Area Departamento  
            var areasDepartamentos = CatalogoEntity.ListadoCatalogosPorCodigo("DTO-01");
            ViewBag.ListadoAreasDepartamentos = areasDepartamentos;

            //Listado Rol  
            var cargos = CatalogoEntity.ListadoCatalogosPorCodigo("CRG-01");
            ViewBag.ListadoCargos = cargos;

            //Listado Pais  
            var paises = CatalogoEntity.ListadoCatalogosPorCodigo("PAI-01");
            ViewBag.listadoPaises = paises;

            //Listado Ciudades  
            var ciudades = CatalogoEntity.ListadoCatalogosPorCodigoId("CIUDAD", Convert.ToInt32(usuario.pais));
            ViewBag.listadoCiudades = ciudades;
             
            RespuestaTransaccion resultado = UsuarioEntity.CrearUsuario(usuario);

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
                return View(usuario);                    
            }   
        }

        public ActionResult Edit(int? id)
        {
            //Listado Tipo Usuario  
            var tipoUsuario = CatalogoEntity.ListadoCatalogosPorCodigo("TUS-01");
            ViewBag.listadoTipoUsuario = tipoUsuario;

            //Listado Clientes  
            var clientes = ClienteEntity.ObtenerListadoClientes();
            ViewBag.listadoClientes = clientes;

            //Listado Rol  
            var roles = RolEntity.ObtenerListadoRoles();
            ViewBag.listadoRoles = roles;

            //Listado Area Departamento  
            var areasDepartamentos = CatalogoEntity.ListadoCatalogosPorCodigo("DTO-01");
            ViewBag.ListadoAreasDepartamentos = areasDepartamentos;

            //Listado Rol  
            var cargos = CatalogoEntity.ListadoCatalogosPorCodigo("CRG-01");
            ViewBag.ListadoCargos = cargos;

            //Listado Pais  
            var paises = CatalogoEntity.ListadoCatalogosPorCodigo("PAI-01");
            ViewBag.listadoPaises = paises;

            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            else
            {
                var usuario = UsuarioEntity.ConsultarUsuario(id.Value);

                if (usuario == null)
                {
                    return HttpNotFound();
                }
                else
                {
                    //Tipo de Usuario
                    ViewBag.tipoUsuarioId = usuario.tipo_usuario;

                    //Listado Ciudades  
                    var ciudades = CatalogoEntity.ListadoCatalogosPorCodigoId("CIUDAD", Convert.ToInt32(usuario.pais));
                    ViewBag.listadoCiudades = ciudades;

                    return View(usuario);
                }
            }
        }

        [HttpPost] 
        public ActionResult Edit(UsuarioCE usuario)
        {
            //Listado Tipo Usuario  
            var tipoUsuario = CatalogoEntity.ListadoCatalogosPorCodigo("TUS-01");
            ViewBag.listadoTipoUsuario = tipoUsuario;

            //Listado Clientes  
            var clientes = ClienteEntity.ObtenerListadoClientes();
            ViewBag.listadoClientes = clientes;

            //Listado Rol  
            var roles = RolEntity.ObtenerListadoRoles();
            ViewBag.listadoRoles = roles;

            //Listado Area Departamento  
            var areasDepartamentos = CatalogoEntity.ListadoCatalogosPorCodigo("DTO-01");
            ViewBag.ListadoAreasDepartamentos = areasDepartamentos;

            //Listado Rol  
            var cargos = CatalogoEntity.ListadoCatalogosPorCodigo("CRG-01");
            ViewBag.ListadoCargos = cargos;

            //Listado Pais  
            var paises = CatalogoEntity.ListadoCatalogosPorCodigo("PAI-01");
            ViewBag.listadoPaises = paises;

            //Listado Ciudades  
            var ciudades = CatalogoEntity.ListadoCatalogosPorCodigoId("CIUDAD", Convert.ToInt32(usuario.pais));
            ViewBag.listadoCiudades = ciudades;

            if (ModelState.IsValid)
            {
                RespuestaTransaccion resultado = UsuarioEntity.ActualizarUsuario(usuario);

                //Almacenar en una variable de sesion
                Session["Resultado"] = resultado.Respuesta;
                Session["Estado"] = resultado.Estado.ToString();

                if (resultado.Estado.ToString() == "True")
                {
                    return RedirectToAction("Index");
                }
                else
                {
                    //Tipo de Usuario
                    ViewBag.tipoUsuarioId = usuario.tipo_usuario;

                    ViewBag.Resultado = resultado.Respuesta;
                    ViewBag.Estado = resultado.Estado.ToString();
                    Session["Resultado"] = "";
                    Session["Estado"] = "";
                    return View(usuario);
                }
            }
            else
            {
                return View(usuario);
            }
        }

        [HttpPost]
        public ActionResult Eliminar(int id)
        {
            if (id == 0)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }

            RespuestaTransaccion resultado = UsuarioEntity.EliminarUsuario(id); 

            //Almacenar en una variable de sesion
            Session["Resultado"] = resultado.Respuesta;
            Session["Estado"] = resultado.Estado.ToString();

            //return RedirectToAction("Index");
            return Json(new { Resultado = resultado }, JsonRequestBehavior.AllowGet);

        }

        //JSON para el combo provincia
        public JsonResult GetCiudad(int id)
        {
            var datos = CatalogoEntity.ListadoCatalogosPorCodigoId("CIUDAD",id);
            return Json(datos, JsonRequestBehavior.AllowGet);
        }

        public ActionResult DescargarReporteFormatoExcel()
        {
            //Seleccionar las columnas a exportar
            var collection = UsuarioEntity.ListarUsuarios();
            var package = new ExcelPackage();
            var nombreHoja = "Usuarios";

            package = Reportes.ExportarExcel(collection.Cast<object>().ToList(), nombreHoja);
            return File(package.GetAsByteArray(), XlsxContentType, "ListadoUsuarios.xlsx");
        }

        public ActionResult DescargarReporteFormatoPDF()
        {
            // Seleccionar las columnas a exportar
            var results = UsuarioEntity.ListarUsuarios();

            var list = Reportes.SerializeToJSON(results);
            return Content(list, "application/json");
        }

        public ActionResult DescargarReporteFormatoCSV()
        {
            var comlumHeadrs = new string[]
            {
                "ID",
                "CODIGO",
                "TIPO_USUARIO",
                "NOMBRES_COMPLETOS", 
                "CLIENTE_ASOCIADO",
                "AREA_O_DEPARTAMENTO",
                "CARGO",
                "PAIS",
                "CIUDAD",
                "DIRECCION",
                "MAIL",
                "TELEFONO",
                "CELULAR",
                "ROL",
                "CODIGO_VENDEDOR",
                "ESTADO"
            };

            var listado = (from item in UsuarioEntity.ListarUsuarios()
                           select new object[]
                           {
                               item.Id,
                               item.Codigo,
                               item.Tipo_Usuario,
                               item.Nombres_Completos,  
                               item.Cliente_Asociado,
                               item.Area_o_Departamento,
                               item.Cargo,
                               item.País,
                               item.Ciudad,
                               item.Direccion,
                               item.Mail,
                               item.Telefono,
                               item.Celular,
                               item.Rol,
                               item.Codigo_Vendedor,
                               item.Estado
                           }).ToList();

            // Build the file content
            var employeecsv = new StringBuilder();
            listado.ForEach(line =>
            {
                employeecsv.AppendLine(string.Join(",", line));
            });

            byte[] buffer = Encoding.Default.GetBytes($"{string.Join(",", comlumHeadrs)}\r\n{employeecsv.ToString()}");
            return File(buffer, "text/csv", $"Listado_Usuarios.csv");
        }

    }


}

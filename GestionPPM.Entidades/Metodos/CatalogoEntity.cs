using GestionPPM.Entidades.Modelo;
using GestionPPM.Repositorios;
using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.Linq;
using System.Threading.Tasks;
using System.Web;
using System.Web.Mvc;

namespace GestionPPM.Entidades.Metodos
{
    public static class CatalogoEntity
    {
        private static readonly GestionPPMEntities db = new GestionPPMEntities();

        public static RespuestaTransaccion CrearCatalogo(Catalogo catalogo)
        {
            try
            {
                //Validar que no exista otro catalogo con ese nombre
                var ValidarCatalogosDuplicados = db.Catalogo.Where(c => c.nombre_catalgo == catalogo.nombre_catalgo).ToList();
                if (ValidarCatalogosDuplicados.Count() > 0)
                {
                    return new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeResgistroExistente };
                }
                else
                {
                    //validar datos obligatorios
                    if (catalogo.nombre_catalgo is null || catalogo.descripcion_catalogo is null || catalogo.codigo_catalogo is null)
                    {
                        return new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeDatosObligatorios };
                    }
                    else
                    {
                        catalogo.nombre_catalgo = catalogo.nombre_catalgo.ToUpper();
                        catalogo.descripcion_catalogo = catalogo.descripcion_catalogo.ToUpper();
                        catalogo.eliminado = false;
                        catalogo.estado_catalogo = true;
                        catalogo.id_empresa = 1;
                        db.Catalogo.Add(catalogo);
                        db.SaveChanges();

                        return new RespuestaTransaccion { Estado = true, Respuesta = Mensajes.MensajeTransaccionExitosa };
                    }
                }
            }
            catch (Exception ex)
            {
                return new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeTransaccionFallida + " ;" + ex.Message.ToString() };
            }
        }

        public static RespuestaTransaccion ActualizarCatalogo(Catalogo catalogo)
        {
            try
            {
                // Por si queda el Attach de la entidad y no deja actualizar
                var local = db.Catalogo.FirstOrDefault(c => c.id_catalogo == catalogo.id_catalogo);
                if (local != null)
                {
                    db.Entry(local).State = EntityState.Detached;
                }

                //validar datos obligatorios
                if (catalogo.nombre_catalgo is null || catalogo.descripcion_catalogo is null || catalogo.codigo_catalogo is null)
                {
                    return new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeDatosObligatorios };
                }
                else
                {
                    //valiar que no se repita el nombre
                    var ValidarCatalogosDuplicados = db.Catalogo.Where(c => c.nombre_catalgo == catalogo.nombre_catalgo && c.id_catalogo != catalogo.id_catalogo).ToList();
                    if (ValidarCatalogosDuplicados.Count() > 0)
                    {
                        return new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeDatosObligatorios };
                    }
                    else
                    {
                        catalogo.nombre_catalgo = catalogo.nombre_catalgo.ToUpper();
                        catalogo.descripcion_catalogo = catalogo.descripcion_catalogo.ToUpper();
                        db.Entry(catalogo).State = EntityState.Modified;
                        db.SaveChanges();
                        return new RespuestaTransaccion { Estado = true, Respuesta = Mensajes.MensajeTransaccionExitosa };
                    }
                }
            }
            catch (Exception ex)
            {
                return new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeTransaccionFallida + " ;" + ex.Message.ToString() };
            }
        }

        public static RespuestaTransaccion EliminarCatalogo(int id)
        {
            try
            {
                var catalogo = db.Catalogo.Find(id);

                if (catalogo.estado_catalogo == true)
                {
                    catalogo.estado_catalogo = false;
                }
                else
                {
                    catalogo.estado_catalogo = true;
                }

                db.Entry(catalogo).State = EntityState.Modified;
                db.SaveChanges();

                return new RespuestaTransaccion { Estado = true, Respuesta = Mensajes.MensajeTransaccionExitosa };
            }
            catch (Exception ex)
            {
                return new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeTransaccionFallida + " ;" + ex.Message.ToString() };
            }
        }

        public static RespuestaTransaccion CrearSubcatalogo(Catalogo catalogo)
        {
            try
            {
                //validar datos obligatorios
                if (catalogo.nombre_catalgo is null)
                {
                    return new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeDatosObligatorios };
                }
                else
                {
                    //Nuevo objeto de tipo catalogo
                    Catalogo subcatalogo = new Catalogo();

                    //Es un subcatalogo del subcatalogo
                    if (catalogo.descripcion_catalogo != null && catalogo.id_catalogo == 0)
                    {
                        //valiar que no se repita el nombre
                        var ValidarCatalogosDuplicados = db.Catalogo.Where(c => c.nombre_catalgo == catalogo.nombre_catalgo && c.codigo_catalogo == catalogo.descripcion_catalogo && c.id_catalogo_padre == catalogo.id_catalogo_padre).ToList();
                        if (ValidarCatalogosDuplicados.Count() > 0)
                        {
                            return new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeResgistroExistente };
                        }
                        else
                        {

                            subcatalogo.nombre_catalgo = catalogo.nombre_catalgo.ToUpper();
                            subcatalogo.descripcion_catalogo = catalogo.nombre_catalgo.ToUpper();
                            subcatalogo.codigo_catalogo = catalogo.descripcion_catalogo;
                            subcatalogo.id_catalogo_padre = catalogo.id_catalogo_padre;
                            subcatalogo.eliminado = false;
                            subcatalogo.id_empresa = 1;
                            subcatalogo.estado_catalogo = true;

                            db.Catalogo.Add(subcatalogo);
                            db.SaveChanges();

                            return new RespuestaTransaccion { Estado = true, Respuesta = Mensajes.MensajeTransaccionExitosa };
                        }
                    }
                    else
                    {
                        //valiar que no se repita el nombre
                        var ValidarCatalogosDuplicados = db.Catalogo.Where(c => c.nombre_catalgo == catalogo.nombre_catalgo && c.id_catalogo_padre == catalogo.id_catalogo).ToList();
                        if (ValidarCatalogosDuplicados.Count() > 0)
                        {
                            return new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeResgistroExistente };
                        }
                        else
                        {
                            if (catalogo.id_catalogo == 774)
                            {
                                subcatalogo.nombre_catalgo = catalogo.nombre_catalgo.ToUpper();
                                subcatalogo.descripcion_catalogo = catalogo.descripcion_catalogo;
                                subcatalogo.codigo_catalogo = catalogo.codigo_catalogo.ToUpper();
                                subcatalogo.id_catalogo_padre = catalogo.id_catalogo;
                                subcatalogo.eliminado = false;
                                subcatalogo.id_empresa = 1;
                                subcatalogo.estado_catalogo = true;

                                db.Catalogo.Add(subcatalogo);
                                db.SaveChanges();
                            }
                            else
                            {

                                subcatalogo.nombre_catalgo = catalogo.nombre_catalgo.ToUpper();
                                subcatalogo.descripcion_catalogo = catalogo.nombre_catalgo.ToUpper();
                                subcatalogo.id_catalogo_padre = catalogo.id_catalogo;
                                subcatalogo.eliminado = false;
                                subcatalogo.id_empresa = 1;
                                subcatalogo.estado_catalogo = true;

                                db.Catalogo.Add(subcatalogo);
                                db.SaveChanges();
                            }

                            return new RespuestaTransaccion { Estado = true, Respuesta = Mensajes.MensajeTransaccionExitosa };
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                return new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeTransaccionFallida + " ;" + ex.Message.ToString() };
            }
        }

        public static RespuestaTransaccion CrearSubcatalogoEtapaCliente(Catalogo catalogo, string general, string detallado, string statusGeneral)
        {
            try
            {
                //Nuevo objeto de tipo catalogo
                Catalogo subcatalogoEG = new Catalogo();
                Catalogo subcatalogoED = new Catalogo();
                Catalogo subcatalogoEGR = new Catalogo();

                //valiar que no se repita el nombre
                var ValidarCatalogosDuplicados = db.Catalogo.Where(c => c.nombre_catalgo == catalogo.nombre_catalgo && c.codigo_catalogo == catalogo.descripcion_catalogo && c.id_catalogo_padre == catalogo.id_catalogo_padre).ToList();
                if (ValidarCatalogosDuplicados.Count() > 0)
                {
                    return new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeResgistroExistente };
                }
                else
                {
                    //Se realiza el resgistro de etapa General
                    subcatalogoEG.codigo_catalogo = "ETAPA-GENERAL";
                    subcatalogoEG.nombre_catalgo = general;
                    subcatalogoEG.descripcion_catalogo = general;
                    subcatalogoEG.id_catalogo_padre = catalogo.id_catalogo_padre;
                    subcatalogoEG.eliminado = false;
                    subcatalogoEG.id_empresa = 1;
                    subcatalogoEG.estado_catalogo = true;

                    db.Catalogo.Add(subcatalogoEG);
                    db.SaveChanges();

                    //Se realiza el resgistro de estatus detallado
                    subcatalogoED.codigo_catalogo = "ESTATUS-DETALLADO";
                    subcatalogoED.nombre_catalgo = detallado;
                    subcatalogoED.descripcion_catalogo = detallado;
                    subcatalogoED.id_catalogo_padre = subcatalogoEG.id_catalogo;
                    subcatalogoED.eliminado = false;
                    subcatalogoED.id_empresa = 1;
                    subcatalogoED.estado_catalogo = true;

                    db.Catalogo.Add(subcatalogoED);
                    db.SaveChanges();

                    //Se realiza el resgistro de estatus general
                    subcatalogoEGR.codigo_catalogo = "ESTATUS-GENERAL";
                    subcatalogoEGR.nombre_catalgo = statusGeneral;
                    subcatalogoEGR.descripcion_catalogo = statusGeneral;
                    subcatalogoEGR.id_catalogo_padre = subcatalogoED.id_catalogo;
                    subcatalogoEGR.eliminado = false;
                    subcatalogoEGR.id_empresa = 1;
                    subcatalogoEGR.estado_catalogo = true;

                    db.Catalogo.Add(subcatalogoEGR);
                    db.SaveChanges();

                    return new RespuestaTransaccion { Estado = true, Respuesta = Mensajes.MensajeTransaccionExitosa };
                }
            }
            catch (Exception ex)
            {
                return new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeTransaccionFallida + " ;" + ex.Message.ToString() };
            }
        }

        public static RespuestaTransaccion ActualizarSubCatalogo(Catalogo catalogo)
        {
            try
            {
                // Por si queda el Attach de la entidad y no deja actualizar
                var local = db.Catalogo.FirstOrDefault(c => c.id_catalogo == catalogo.id_catalogo);
                if (local != null)
                {
                    db.Entry(local).State = EntityState.Detached;
                }

                //validar que no este vacio
                if (catalogo.nombre_catalgo == null)
                {
                    return new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeDatosObligatorios };
                }
                else
                {
                    //valiar que no se repita el nombre
                    var ValidarCatalogosDuplicados = db.Catalogo.Where(c => c.nombre_catalgo == catalogo.nombre_catalgo && c.id_catalogo != catalogo.id_catalogo && c.id_catalogo_padre == catalogo.id_catalogo_padre).ToList();
                    if (ValidarCatalogosDuplicados.Count() > 0)
                    {
                        return new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeDatosObligatorios };
                    }
                    else
                    {
                        catalogo.nombre_catalgo = catalogo.nombre_catalgo.ToUpper();
                        catalogo.descripcion_catalogo = catalogo.nombre_catalgo.ToUpper();
                        db.Entry(catalogo).State = EntityState.Modified;
                        db.SaveChanges();
                        return new RespuestaTransaccion { Estado = true, Respuesta = Mensajes.MensajeTransaccionExitosa };
                    }
                }

            }
            catch (Exception ex)
            {
                return new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeTransaccionFallida + " ;" + ex.Message.ToString() };
            }
        }

        public static RespuestaTransaccion EliminarSubcatalogo(int id)
        {
            try
            {
                var catalogo = db.Catalogo.First(c => c.id_catalogo == id);

                db.Catalogo.Remove(catalogo);
                db.SaveChanges();

                return new RespuestaTransaccion { Estado = true, Respuesta = Mensajes.MensajeTransaccionExitosa };
            }
            catch (Exception ex)
            {
                return new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeTransaccionFallida + " ;" + ex.Message.ToString() };
            }
        }


        public static List<ListadoCatalogosPadres> ListarCatalogos()
        {
            try
            {
                //listado de catalogs padres
                return db.ListadoCatalogosPadres().ToList();
            }
            catch (Exception)
            {
                throw;
            }
        }
        public static List<ListadoSubcatalogos> ListarSubcatalogos()
        {
            try
            {
                //listado de catalogs padres
                return db.ListadoSubcatalogos().ToList();
            }
            catch (Exception)
            {
                throw;
            }
        }

        public static List<Catalogo> ListarCatalogo()
        {
            try
            {
                using (var context = new GestionPPMEntities())
                {
                    return context.Catalogo.ToList();
                }
            }
            catch (Exception)
            {
                throw;
            }
        }

        public static Catalogo ConsultarCatalogo(int? id)
        {
            try
            {
                Catalogo catalogo = db.Catalogo.Find(id);
                return catalogo;
            }
            catch (Exception)
            {
                throw;
            }
        }

        public static string ConsultarNombreCatalogo(int? id)
        {
            try
            {
                if (id.HasValue)
                {
                    var catalogo = db.Catalogo.Find(id.Value);
                    return catalogo.nombre_catalgo;
                }
                else
                {
                    return Mensajes.MensajeCatalogoNoDisponible;
                }
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
        }

        public static int BuscarCatalogo(int id)
        {
            try
            {
                var catalogo = db.Catalogo.Find(id);
                if (catalogo == null)
                {
                    return catalogo.id_catalogo;
                }
                else
                {
                    return catalogo.id_catalogo;
                }
            }
            catch (Exception ex)
            {
                return 0;
            }
        }

        public static bool VerificarExistenciaCatalogo(int id)
        {
            try
            {
                using (var context = new GestionPPMEntities())
                {
                    var catalogo = context.Catalogo.Find(id);
                    if (catalogo == null)
                        return false;
                    else
                        return true;
                }
            }
            catch (Exception ex)
            {
                return false;
            }
        }

        public static IEnumerable<SelectListItem> ListadoHijosoCatalogoPorIdPadre(int id_catalogo)
        {
            var ListadoCatalogo = db.usp_b_catalogo_hijos(id_catalogo).Select(c => new SelectListItem
            {
                Text = c.nombre_catalgo,
                Value = c.id_catalogo.ToString()
            }).ToList();

            return ListadoCatalogo;
        }

        public static IEnumerable<SelectListItem> ListadoCatalogosPorCodigo(string codigo)
        {
            var ListadoCatalogo = db.usp_b_catalogo_codigo(codigo).OrderBy(c => c.nombre_catalgo).Select(c => new SelectListItem
            {
                Text = c.nombre_catalgo.ToUpper(),
                Value = c.id_catalogo.ToString()
            }).ToList();

            return ListadoCatalogo;
        }

        public static bool VerificarExistenciaCatalogoLineaNegocio(int idCatalogo)
        {
            var catalogo = db.usp_b_catalogo_codigo("LNG-CLI-01").Where(s => s.id_catalogo == idCatalogo).FirstOrDefault();
            if (catalogo != null)
                return true;
            else
                return false;
        }

        public static IEnumerable<SelectListItem> ListadoCatalogosPorId(int id)
        {
            var ListadoCatalogo = db.usp_b_catalogo_id(id, "").OrderBy(c => c.Nombre).Select(c => new SelectListItem
            {
                Text = c.Nombre,
                Value = c.Id.ToString()
            }).ToList();

            return ListadoCatalogo;
        }

        public static IEnumerable<SelectListItem> ListadoCatalogosPorIdSinOrdenar(int id)
        {
            var ListadoCatalogo = db.usp_b_catalogo_id(id, "").Select(c => new SelectListItem
            {
                Text = c.Nombre,
                Value = c.Id.ToString()
            }).ToList();

            return ListadoCatalogo;
        }

        public static IEnumerable<SelectListItem> ListadoCatalogosPorCodigoId(string codigo, int id)
        {
            var ListadoCatalogo = db.usp_b_catalogo_codigo_id(codigo, id).OrderBy(c => c.nombre_catalgo).Select(c => new SelectListItem
            {
                Text = c.nombre_catalgo,
                Value = c.id_catalogo.ToString()
            }).ToList();

            return ListadoCatalogo;
        }

        public static int ObtenerNumeroHijosCatalgo(int id)
        {
            try
            {
                var numero_hijos = db.usp_b_nume_hijo_cata(id).First();
                int numero = Convert.ToInt32(numero_hijos.numero_hijos);

                return numero;
            }
            catch (Exception ex)
            {
                return 0;
            }
        }

        public static int ObtenerIdPadre(int id)
        {
            try
            {
                var numero_hijos = db.Catalogo.Where(c => c.id_catalogo == id).First().id_catalogo_padre;
                int idPadre = Convert.ToInt32(numero_hijos);

                return idPadre;
            }
            catch (Exception ex)
            {
                return 0;
            }
        }

        public static List<usp_b_catalogo_id> ListarCatalogosPorId(int id, string filtro)
        {
            try
            {
                var listado = db.usp_b_catalogo_id(id, filtro).ToList();
                return listado;
            }
            catch (Exception)
            {
                throw;
            }
        }

        public static IEnumerable<SelectListItem> ObtenerListadoCatalogos(int id_padre)
        {
            List<SelectListItem> listado = new List<SelectListItem>();
            try
            {
                listado = db.Catalogo.Where(c => c.id_catalogo_padre == id_padre).OrderBy(c => c.nombre_catalgo).Select(c => new SelectListItem
                {
                    Text = c.nombre_catalgo,
                    Value = c.id_catalogo.ToString()
                }).ToList();

                return listado;
            }
            catch (Exception)
            {
                return listado;
                throw;
            }

        }

        public static IEnumerable<SelectListItem> ObtenerListadoCatalogosSublineaNegocio(int value, string seleccionado = null)
        {
            List<SelectListItem> ListadoCatalogo = new List<SelectListItem>();
            try
            {
                using (var context = new GestionPPMEntities())
                {
                    ListadoCatalogo = context.Catalogo.Where(c => c.codigo_catalogo == "SLN-01" && c.id_catalogo_padre == value && c.estado_catalogo.Value).OrderBy(c => c.nombre_catalgo).Select(c => new SelectListItem
                    {
                        Text = c.nombre_catalgo,
                        Value = c.id_catalogo.ToString()
                    }).ToList();

                    if (!string.IsNullOrEmpty(seleccionado))
                    {
                        if (ListadoCatalogo.FirstOrDefault(s => s.Value == seleccionado.ToString()) != null)
                            ListadoCatalogo.FirstOrDefault(s => s.Value == seleccionado.ToString()).Selected = true;
                    }

                    return ListadoCatalogo;
                }
            }
            catch (Exception ex)
            {
                return ListadoCatalogo;
            }
        }

        public static bool VerificarSublineaNegocio(int idSublineaNegocio)
        {
            try
            {
                using (var context = new GestionPPMEntities())
                {
                    var sublinea = context.Catalogo.Where(c => c.codigo_catalogo == "SLN-01" && c.id_catalogo_padre != null && c.estado_catalogo.Value && c.id_catalogo == idSublineaNegocio).FirstOrDefault();

                    if (sublinea == null)
                        return false;
                    else
                        return true;
                }
            }
            catch (Exception ex)
            {
                return false;
            }
        }

        public static IEnumerable<SelectListItem> ObtenerListadoCatalogosByCodigoSeleccion(string codigo, string seleccionado = null)
        {
            List<SelectListItem> ListadoCatalogo = new List<SelectListItem>();
            try
            {
                var listado = db.Catalogo.Where(s => s.codigo_catalogo == codigo).ToList();

                ListadoCatalogo = listado.OrderBy(c => c.nombre_catalgo).Select(c => new SelectListItem
                {
                    Text = c.nombre_catalgo,
                    Value = c.id_catalogo.ToString()
                }).ToList();

                if (!string.IsNullOrEmpty(seleccionado))
                {
                    if (ListadoCatalogo.FirstOrDefault(s => s.Value == seleccionado.ToString()) != null)
                        ListadoCatalogo.FirstOrDefault(s => s.Value == seleccionado.ToString()).Selected = true;
                }

                return ListadoCatalogo;
            }
            catch (Exception ex)
            {
                return ListadoCatalogo;
            }

        }

        public static List<Catalogo> ListadoCatalogosEtapaCliente(int idCatalogoPadre, string codigo = "")
        {
            List<Catalogo> listado = new List<Catalogo>();
            try
            {
                var etapaGeneral = db.Catalogo.FirstOrDefault(s => s.id_catalogo_padre == idCatalogoPadre);

                if (etapaGeneral != null)
                {
                    listado.Add(new Catalogo { nombre_catalgo = etapaGeneral.nombre_catalgo, id_catalogo = etapaGeneral.id_catalogo, codigo_catalogo = etapaGeneral.codigo_catalogo });


                    var estatusDetallado = db.Catalogo.FirstOrDefault(s => s.id_catalogo_padre == etapaGeneral.id_catalogo);

                    if (estatusDetallado != null)
                        listado.Add(new Catalogo { nombre_catalgo = estatusDetallado.nombre_catalgo, id_catalogo = estatusDetallado.id_catalogo, codigo_catalogo = estatusDetallado.codigo_catalogo });


                    var estatusGeneral = db.Catalogo.FirstOrDefault(s => s.id_catalogo_padre == estatusDetallado.id_catalogo);

                    if (estatusGeneral != null)
                        listado.Add(new Catalogo { nombre_catalgo = estatusGeneral.nombre_catalgo, id_catalogo = estatusGeneral.id_catalogo, codigo_catalogo = estatusGeneral.codigo_catalogo });
                }
                return listado;
            }
            catch (Exception)
            {
                return listado;
            }
        }

        public static IEnumerable<SelectListItem> ConsultarCatalogoPorPadre(int id, string codigo)
        {
            try
            {
                using (var context = new GestionPPMEntities())
                {
                    var catalogo = context.Catalogo.Where(s => s.id_catalogo_padre == id && s.codigo_catalogo == codigo).OrderBy(c => c.nombre_catalgo).Select(c => new SelectListItem
                    {
                        Text = c.nombre_catalgo,
                        Value = c.id_catalogo.ToString()
                    }).ToList();

                    return catalogo;
                }

            }
            catch (Exception)
            {
                throw;
            }
        }

        public static IEnumerable<SelectListItem> ObtenerListadoCatalogosByCodigo(string codigo, string seleccionado = null)
        {
            List<SelectListItem> ListadoCatalogo = new List<SelectListItem>();
            try
            {
                var padreCatalogo = db.Catalogo.FirstOrDefault(s => s.codigo_catalogo == codigo);

                if (padreCatalogo == null)
                {
                    return new List<SelectListItem>();
                }

                ListadoCatalogo = db.Catalogo.Where(c => c.id_catalogo_padre == padreCatalogo.id_catalogo && c.estado_catalogo.Value).OrderBy(c => c.nombre_catalgo).Select(c => new SelectListItem
                {
                    Text = c.nombre_catalgo,
                    Value = c.id_catalogo.ToString(),
                }).ToList();

                if (!string.IsNullOrEmpty(seleccionado))
                {
                    if (ListadoCatalogo.FirstOrDefault(s => s.Value == seleccionado.ToString()) != null)
                        ListadoCatalogo.FirstOrDefault(s => s.Value == seleccionado.ToString()).Selected = true;
                }

                return ListadoCatalogo;
            }
            catch (Exception ex)
            {
                return ListadoCatalogo;
            }

        }

        //Para dependientes
        public static IEnumerable<SelectListItem> ConsultarCatalogoPorPadreByCodigo(string codigo, int id, string seleccionado = null)
        {
            SelectListItem vacio = new SelectListItem { Disabled = true, Selected = true, Text = Etiquetas.TituloComboVacio, Value = "" };
            List<SelectListItem> ListadoCatalogo = new List<SelectListItem> { vacio };
            try
            {
                using (var context = new GestionPPMEntities())
                {
                    var padreCatalogo = db.Catalogo.FirstOrDefault(s => s.codigo_catalogo == codigo);

                    if (padreCatalogo == null)
                    {
                        return ListadoCatalogo;
                    }

                    if (id != 0)
                        ListadoCatalogo = context.Catalogo.Where(s => s.id_catalogo_padre == id && s.codigo_catalogo == codigo).OrderBy(c => c.nombre_catalgo).Select(c => new SelectListItem
                        {
                            Text = c.nombre_catalgo,
                            Value = c.id_catalogo.ToString()
                        }).ToList();
                    else
                        return ListadoCatalogo;

                    if (!string.IsNullOrEmpty(seleccionado))
                    {
                        if (ListadoCatalogo.FirstOrDefault(s => s.Value == seleccionado.ToString()) != null)
                            ListadoCatalogo.FirstOrDefault(s => s.Value == seleccionado.ToString()).Selected = true;
                    }

                    return ListadoCatalogo;

                }
            }
            catch (Exception)
            {
                return ListadoCatalogo;
            }
        }

        public static IEnumerable<SelectListItem> ObtenerListadoCatalogosStatusByCodigo(string codigo, string seleccionado = null)
        {
            List<SelectListItem> ListadoCatalogo = new List<SelectListItem>();
            try
            {
                var padreCatalogo = db.Catalogo.FirstOrDefault(s => s.codigo_catalogo == codigo);

                if (padreCatalogo == null)
                {
                    return new List<SelectListItem>();
                }
                List<int> filtro = new List<int> { 68, 69 };
                ListadoCatalogo = db.Catalogo.Where(c => c.id_catalogo_padre == padreCatalogo.id_catalogo && filtro.Contains(c.id_catalogo)).OrderBy(c => c.nombre_catalgo).Select(c => new SelectListItem
                {
                    Text = c.nombre_catalgo,
                    Value = c.id_catalogo.ToString(),
                }).ToList();

                if (!string.IsNullOrEmpty(seleccionado))
                {
                    if (ListadoCatalogo.FirstOrDefault(s => s.Value == seleccionado.ToString()) != null)
                        ListadoCatalogo.FirstOrDefault(s => s.Value == seleccionado.ToString()).Selected = true;
                }

                return ListadoCatalogo;
            }
            catch (Exception ex)
            {
                return ListadoCatalogo;
            }

        }

        public static IEnumerable<SelectListItem> ObtenerListadoCatalogosByCodigoSinOrdenar(string codigo, string seleccionado = null)
        {
            List<SelectListItem> ListadoCatalogo = new List<SelectListItem>();
            try
            {
                var padreCatalogo = db.Catalogo.FirstOrDefault(s => s.codigo_catalogo == codigo);

                if (padreCatalogo == null)
                {
                    return new List<SelectListItem>();
                }


                ListadoCatalogo = db.Catalogo.Where(c => c.id_catalogo_padre == padreCatalogo.id_catalogo).Select(c => new SelectListItem
                {
                    Text = c.nombre_catalgo,
                    Value = c.id_catalogo.ToString(),
                }).ToList();

                if (!string.IsNullOrEmpty(seleccionado))
                {
                    if (ListadoCatalogo.FirstOrDefault(s => s.Value == seleccionado.ToString()) != null)
                        ListadoCatalogo.FirstOrDefault(s => s.Value == seleccionado.ToString()).Selected = true;
                }
               

                return ListadoCatalogo;
            }
            catch (Exception ex)
            {
                return ListadoCatalogo;
            }

        }

        public static bool ValidarExistenciaPrefijo(string numeroPrefijo)
        {
            bool flag = true;
            try
            {
                var padreCatalogo = db.Catalogo.FirstOrDefault(s => s.codigo_catalogo == "PREFIJO" && s.nombre_catalgo == numeroPrefijo);
                if (padreCatalogo == null)
                    flag = false;

                return flag;
            }
            catch (Exception ex)
            {
                return false;
            }
        }


        public static List<Catalogo> ListadoCatalogosEtapaClienteTipoZoho(int idCatalogoPadre)
        {
            //Tipo Zoho Cliente
            if (idCatalogoPadre == 30)
            {

                var EstapaCliente = db.Catalogo.Where(s => s.id_catalogo_padre == 357 && s.nombre_catalgo != "GESTIÓN").ToList();
                return EstapaCliente;
            }
            else
            {
                var EstapaCliente = db.Catalogo.Where(s => s.id_catalogo_padre == 357).ToList();
                return EstapaCliente;
            }
        }

        public static IEnumerable<SelectListItem> ObtenerListadoCatalogos2(int id_padre)
        {
            List<SelectListItem> listado = new List<SelectListItem>();
            try
            {
                listado = db.Catalogo.Where(c => c.id_catalogo_padre == id_padre).OrderBy(c => c.nombre_catalgo).Select(c => new SelectListItem
                {
                    Text = c.nombre_catalgo + "|" + c.descripcion_catalogo,
                    Value = c.id_catalogo.ToString()
                }).ToList();

                return listado;
            }
            catch (Exception)
            {
                return listado;
                throw;
            }

        }
        public static string ConsultarCodigocatalogo(int id)
        {
            try
            {

                Catalogo catalogo = db.Catalogo.Find(id);
                if (catalogo == null)
                {
                    return string.Empty;
                }
                return catalogo.codigo_catalogo;
            }
            catch (Exception)
            {
                return string.Empty;
            }
        }

        public static IEnumerable<SelectListItem> ObtenerListadoCatalogosSublineaNegocioPadre(string seleccionado = null)
        {
            List<SelectListItem> ListadoCatalogo = new List<SelectListItem>();
            try
            {
                using (var context = new GestionPPMEntities())
                {
                    var padresActivos = context.Catalogo.Where(c => c.id_catalogo_padre == 111 && c.estado_catalogo.Value).Select(c => c.id_catalogo).ToList();
                    ListadoCatalogo = context.Catalogo.Where(c => c.codigo_catalogo == "SLN-01" && c.id_catalogo_padre != null && c.estado_catalogo.Value && padresActivos.Contains(c.id_catalogo_padre.Value)).OrderBy(c => c.nombre_catalgo).Select(c => new SelectListItem
                    {
                        Text = c.nombre_catalgo,
                        Value = c.id_catalogo.ToString()
                    }).ToList();

                    if (!string.IsNullOrEmpty(seleccionado))
                    {
                        if (ListadoCatalogo.FirstOrDefault(s => s.Value == seleccionado.ToString()) != null)
                            ListadoCatalogo.FirstOrDefault(s => s.Value == seleccionado.ToString()).Selected = true;
                    }

                    return ListadoCatalogo;
                }
            }
            catch (Exception ex)
            {
                return ListadoCatalogo;
            }
        }

    }
}
using GestionPPM.Entidades.Modelo;
using GestionPPM.Repositorios;
using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.Data.Entity.Validation;
using System.Linq;
using System.Threading.Tasks;
using System.Web;
using System.Web.Mvc;

namespace GestionPPM.Entidades.Metodos
{
    public class ActaEntity
    {
        private static readonly GestionPPMEntities db = new GestionPPMEntities();

        public static RespuestaTransaccion CrearActa(Acta cabecera, DetallesActaParcial cuerpo, ActaInformacionAdicional piePagina, int tipoActa)
        {
            using (var transaction = db.Database.BeginTransaction())
            {

                try
                {
                    string validacion = string.Empty;
                    if (cuerpo.Acuerdos.Any())
                        validacion = ValidarFechaAcuerdos(cuerpo.Acuerdos, cabecera.FechaInicio);

                    if (!string.IsNullOrEmpty(validacion))
                        return new RespuestaTransaccion { Estado = false, Respuesta = validacion };

                    #region Guardar Cabecera Acta
                    var generacionCodigo = db.GenerarCodigoActa(tipoActa).FirstOrDefault();

                    if (generacionCodigo != null)
                    {
                        var existenciaCodigo = db.Acta.Where(s => s.CodigoActa == generacionCodigo.CodigoActa && s.TipoActaID == tipoActa).Any();

                        if (!existenciaCodigo)
                            cabecera.CodigoActa = generacionCodigo.CodigoActa;
                        else
                            return new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeCodigoActaError };
                    }
                    else
                        return new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeCodigoActaError };

                    cabecera.Estado = true;
                    db.Acta.Add(cabecera);
                    db.SaveChanges();
                    #endregion

                    #region Guardar Cuerpo Acta

                    foreach (var detalleEntregables in cuerpo.Entregables)
                    {
                        DetalleActaEntregables entregable = new DetalleActaEntregables
                        {
                            Entregable = detalleEntregables.Entregable,
                            Tipo = detalleEntregables.Tipo
                        };
                        db.DetalleActaEntregables.Add(entregable);
                        db.SaveChanges();

                        DetallesActa detalle = new DetallesActa
                        {
                            DetalleActaEntregablesID = entregable.IDDetalleActaEntregables,
                            ActaID = cabecera.IDActa,
                        };

                        db.DetallesActa.Add(detalle);
                        db.SaveChanges();

                    }

                    foreach (var detalleAcuerdos in cuerpo.Acuerdos)
                    {
                        DetalleActaAcuerdos acuerdos = new DetalleActaAcuerdos
                        {
                            Acuerdo = detalleAcuerdos.Acuerdo,
                            Responsable = detalleAcuerdos.Responsable,
                            Fecha = detalleAcuerdos.Fecha,
                        };
                        db.DetalleActaAcuerdos.Add(acuerdos);
                        db.SaveChanges();

                        DetallesActa detalle = new DetallesActa
                        {
                            DetalleActaAcuerdosID = acuerdos.IDDetalleActaAcuerdos,
                            ActaID = cabecera.IDActa,
                        };

                        db.DetallesActa.Add(detalle);
                        db.SaveChanges();
                    }

                    foreach (var detalleParticipantes in cuerpo.Participantes)
                    {
                        DetalleActaParticipantes participantes = new DetalleActaParticipantes
                        {
                            Nombres = detalleParticipantes.Nombres,
                            Presente = detalleParticipantes.Presente
                        };
                        db.DetalleActaParticipantes.Add(participantes);
                        db.SaveChanges();

                        DetallesActa detalle = new DetallesActa
                        {
                            DetalleActaParticipantesID = participantes.IDDetalleActaParticipantes,
                            ActaID = cabecera.IDActa,
                        };

                        db.DetallesActa.Add(detalle);
                        db.SaveChanges();
                    }

                    foreach (var detalleResponsables in cuerpo.Responsables)
                    {
                        DetalleActaResponsables responsables = new DetalleActaResponsables
                        {
                            Nombres = detalleResponsables.Nombres,
                            Rol = detalleResponsables.Rol,
                            Empresa = detalleResponsables.Empresa
                        };
                        db.DetalleActaResponsables.Add(responsables);
                        db.SaveChanges();

                        DetallesActa detalle = new DetallesActa
                        {
                            DetalleActaResponsablesClienteID = responsables.IDDetalleActaResponsablesCliente,
                            ActaID = cabecera.IDActa,
                        };

                        db.DetallesActa.Add(detalle);
                        db.SaveChanges();
                    }

                    foreach (var detalleTemas in cuerpo.Temas)
                    {
                        DetalleActaTemasTratar temas = new DetalleActaTemasTratar
                        {
                            Responsable = detalleTemas.Responsable,
                            Tema = detalleTemas.Tema
                        };
                        db.DetalleActaTemasTratar.Add(temas);
                        db.SaveChanges();

                        DetallesActa detalle = new DetallesActa
                        {
                            DetalleActaTemasTratarID = temas.IDDetalleActaTemasTratar,
                            ActaID = cabecera.IDActa,
                        };

                        db.DetallesActa.Add(detalle);
                        db.SaveChanges();
                    }

                    foreach (var detalleCondiciones in cuerpo.CondicionesGenerales)
                    {
                        DetalleActaCondicionesGenerales condiciones = new DetalleActaCondicionesGenerales
                        {
                            Condicion = detalleCondiciones.Condicion,
                        };
                        db.DetalleActaCondicionesGenerales.Add(condiciones);
                        db.SaveChanges();

                        DetallesActa detalle = new DetallesActa
                        {
                            DetalleActaCondicionesGeneralesID = Convert.ToInt32(condiciones.IDActaCondicionesGenerales),
                            ActaID = cabecera.IDActa,
                        };

                        db.DetallesActa.Add(detalle);
                        db.SaveChanges();
                    }

                    //Nuevos Detalles
                    foreach (var detalleCliente in cuerpo.DetalleCliente)
                    {
                        DetalleActaCliente cliente = new DetalleActaCliente
                        {
                            id_facturacion_safi = detalleCliente.id_facturacion_safi,
                        };
                        db.DetalleActaCliente.Add(cliente);
                        db.SaveChanges();

                        DetallesActa detalle = new DetallesActa
                        {
                            DetalleActaClienteID = Convert.ToInt32(cliente.IDDetalleActaCliente), // Revisar que tipo de dato
                            ActaID = cabecera.IDActa,
                        };

                        db.DetallesActa.Add(detalle);
                        db.SaveChanges();
                    }

                    //Nuevos Detalles
                    foreach (var detalleCliente in cuerpo.DetalleContabilidad)
                    {
                        DetalleActaContabilidad contabilidad = new DetalleActaContabilidad
                        {
                            id_facturacion_safi = detalleCliente.id_facturacion_safi,
                        };
                        db.DetalleActaContabilidad.Add(contabilidad);
                        db.SaveChanges();

                        DetallesActa detalle = new DetallesActa
                        {
                            DetalleActaContabilidadID = Convert.ToInt32(contabilidad.IDDetalleActaContabilidad), // Revisar que tipo de dato
                            ActaID = cabecera.IDActa,
                        };

                        db.DetallesActa.Add(detalle);
                        db.SaveChanges();
                    }
                    #endregion

                    #region Guardar Pie Pagina Acta
                    piePagina.ActaID = cabecera.IDActa;
                    db.ActaInformacionAdicional.Add(piePagina);
                    db.SaveChanges();
                    #endregion

                    #region Actualizacion Secuencial Acta
                    var SecuencialActa = db.SecuencialTipoActa.Find(tipoActa);
                    SecuencialActa.Secuencial += 1;
                    SecuencialActa.Anio = generacionCodigo.Anio;
                    db.Entry(SecuencialActa).State = EntityState.Modified;
                    db.SaveChanges();
                    #endregion

                    transaction.Commit();

                    return new RespuestaTransaccion { Estado = true, Respuesta = Mensajes.MensajeTransaccionExitosa, EntidadID = cabecera.IDActa };
                }
                catch (Exception ex)
                {
                    transaction.Rollback();
                    return new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeTransaccionFallida + " ;" + ex.Message.ToString() };
                }

            }
        }

        public static bool ActualizarSecuencialActa(int tipoActa)
        {
            try
            {
                var SecuencialActa = db.SecuencialTipoActa.Find(tipoActa);
                SecuencialActa.Secuencial += 1;
                db.Entry(SecuencialActa).State = EntityState.Modified;
                db.SaveChanges();
                return true;
            }
            catch (Exception ex)
            {
                return false;
            }
        }

        public static RespuestaTransaccion ActualizarActa(Acta cabecera, DetallesActaParcial cuerpo, ActaInformacionAdicional piePagina)
        {
            using (var transaction = db.Database.BeginTransaction())
            {

                try
                {
                    string validacion = string.Empty;
                    if (cuerpo.Acuerdos.Any())
                        validacion = ValidarFechaAcuerdos(cuerpo.Acuerdos, cabecera.FechaInicio);

                    if (!string.IsNullOrEmpty(validacion))
                        return new RespuestaTransaccion { Estado = false, Respuesta = validacion };

                    #region Actualizar Cabecera Acta
                    // assume Entity base class have an Id property for all items
                    var entity = db.Acta.Find(cabecera.IDActa);
                    if (entity == null)
                    {
                        return new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeTransaccionFallida };
                    }

                    db.Entry(entity).CurrentValues.SetValues(cabecera);

                    #endregion

                    #region Limpiar detalles anteriores
                    // Limpiar primero los detalles anteriores del acta.
                    var detallesAnteriores = db.DetallesActa.Where(s => s.ActaID == cabecera.IDActa).ToList();
                    foreach (var item in detallesAnteriores)
                    {
                        db.DetallesActa.Remove(item);
                        db.SaveChanges();
                    }

                    foreach (var detalleEntregables in cuerpo.Entregables.Where(s => s.IDDetalleActaEntregables != 0))
                    {
                        var elemento = db.DetalleActaEntregables.Find(detalleEntregables.IDDetalleActaEntregables);
                        db.DetalleActaEntregables.Remove(elemento);
                        db.SaveChanges();
                    }

                    foreach (var detalleAcuerdos in cuerpo.Acuerdos.Where(s => s.IDDetalleActaAcuerdos != 0))
                    {
                        var elemento = db.DetalleActaAcuerdos.Find(detalleAcuerdos.IDDetalleActaAcuerdos);
                        db.DetalleActaAcuerdos.Remove(elemento);
                        db.SaveChanges();
                    }

                    foreach (var detalleParticipantes in cuerpo.Participantes.Where(s => s.IDDetalleActaParticipantes != 0))
                    {
                        var elemento = db.DetalleActaParticipantes.Find(detalleParticipantes.IDDetalleActaParticipantes);
                        db.DetalleActaParticipantes.Remove(elemento);
                        db.SaveChanges();
                    }

                    foreach (var detalleResponsables in cuerpo.Responsables.Where(s => s.IDDetalleActaResponsablesCliente != 0))
                    {
                        var elemento = db.DetalleActaResponsables.Find(detalleResponsables.IDDetalleActaResponsablesCliente);
                        db.DetalleActaResponsables.Remove(elemento);
                        db.SaveChanges();
                    }

                    foreach (var detalleTemas in cuerpo.Temas.Where(s => s.IDDetalleActaTemasTratar != 0))
                    {
                        var elemento = db.DetalleActaTemasTratar.Find(detalleTemas.IDDetalleActaTemasTratar);
                        db.DetalleActaTemasTratar.Remove(elemento);
                        db.SaveChanges();
                    }

                    foreach (var detalleCondiciones in cuerpo.CondicionesGenerales.Where(s => s.IDActaCondicionesGenerales != 0))
                    {
                        var elemento = db.DetalleActaCondicionesGenerales.Find(detalleCondiciones.IDActaCondicionesGenerales);
                        db.DetalleActaCondicionesGenerales.Remove(elemento);
                        db.SaveChanges();
                    }

                    //Nuevos detalles
                    foreach (var detalleCliente in cuerpo.DetalleCliente.Where(s => s.IDDetalleActaCliente != 0))
                    {
                        var elemento = db.DetalleActaCliente.Find(detalleCliente.IDDetalleActaCliente);
                        db.DetalleActaCliente.Remove(elemento);
                        db.SaveChanges();
                    }

                    foreach (var detalleCliente in cuerpo.DetalleContabilidad.Where(s => s.IDDetalleActaContabilidad != 0))
                    {
                        var elemento = db.DetalleActaContabilidad.Find(detalleCliente.IDDetalleActaContabilidad);
                        db.DetalleActaContabilidad.Remove(elemento);
                        db.SaveChanges();
                    }

                    #endregion

                    #region Actualizar Cuerpo Acta

                    foreach (var detalleEntregables in cuerpo.Entregables)
                    {
                        DetalleActaEntregables entregable = new DetalleActaEntregables
                        {
                            Entregable = detalleEntregables.Entregable,
                            Tipo = detalleEntregables.Tipo
                        };
                        db.DetalleActaEntregables.Add(entregable);
                        db.SaveChanges();

                        DetallesActa detalle = new DetallesActa
                        {
                            DetalleActaEntregablesID = entregable.IDDetalleActaEntregables,
                            ActaID = cabecera.IDActa,
                        };

                        db.DetallesActa.Add(detalle);
                        db.SaveChanges();

                    }

                    foreach (var detalleAcuerdos in cuerpo.Acuerdos)
                    {
                        DetalleActaAcuerdos acuerdos = new DetalleActaAcuerdos
                        {
                            Acuerdo = detalleAcuerdos.Acuerdo,
                            Responsable = detalleAcuerdos.Responsable,
                            Fecha = detalleAcuerdos.Fecha,
                        };
                        db.DetalleActaAcuerdos.Add(acuerdos);
                        db.SaveChanges();

                        DetallesActa detalle = new DetallesActa
                        {
                            DetalleActaAcuerdosID = acuerdos.IDDetalleActaAcuerdos,
                            ActaID = cabecera.IDActa,
                        };

                        db.DetallesActa.Add(detalle);
                        db.SaveChanges();
                    }

                    foreach (var detalleParticipantes in cuerpo.Participantes)
                    {
                        DetalleActaParticipantes participantes = new DetalleActaParticipantes
                        {
                            Nombres = detalleParticipantes.Nombres,
                            Presente = detalleParticipantes.Presente
                        };
                        db.DetalleActaParticipantes.Add(participantes);
                        db.SaveChanges();

                        DetallesActa detalle = new DetallesActa
                        {
                            DetalleActaParticipantesID = participantes.IDDetalleActaParticipantes,
                            ActaID = cabecera.IDActa,
                        };

                        db.DetallesActa.Add(detalle);
                        db.SaveChanges();
                    }

                    foreach (var detalleResponsables in cuerpo.Responsables)
                    {
                        DetalleActaResponsables responsables = new DetalleActaResponsables
                        {
                            Nombres = detalleResponsables.Nombres,
                            Rol = detalleResponsables.Rol,
                            Empresa = detalleResponsables.Empresa
                        };
                        db.DetalleActaResponsables.Add(responsables);
                        db.SaveChanges();

                        DetallesActa detalle = new DetallesActa
                        {
                            DetalleActaResponsablesClienteID = responsables.IDDetalleActaResponsablesCliente,
                            ActaID = cabecera.IDActa,
                        };

                        db.DetallesActa.Add(detalle);
                        db.SaveChanges();
                    }

                    foreach (var detalleTemas in cuerpo.Temas)
                    {
                        DetalleActaTemasTratar temas = new DetalleActaTemasTratar
                        {
                            Responsable = detalleTemas.Responsable,
                            Tema = detalleTemas.Tema
                        };
                        db.DetalleActaTemasTratar.Add(temas);
                        db.SaveChanges();

                        DetallesActa detalle = new DetallesActa
                        {
                            DetalleActaTemasTratarID = temas.IDDetalleActaTemasTratar,
                            ActaID = cabecera.IDActa,
                        };

                        db.DetallesActa.Add(detalle);
                        db.SaveChanges();
                    }

                    foreach (var detalleCondiciones in cuerpo.CondicionesGenerales)
                    {
                        DetalleActaCondicionesGenerales condiciones = new DetalleActaCondicionesGenerales
                        {
                            Condicion = detalleCondiciones.Condicion,
                        };
                        db.DetalleActaCondicionesGenerales.Add(condiciones);
                        db.SaveChanges();

                        DetallesActa detalle = new DetallesActa
                        {
                            DetalleActaCondicionesGeneralesID = Convert.ToInt32(condiciones.IDActaCondicionesGenerales),
                            ActaID = cabecera.IDActa,
                        };

                        db.DetallesActa.Add(detalle);
                        db.SaveChanges();
                    }

                    //Nuevos Detalles
                    foreach (var detalleCliente in cuerpo.DetalleCliente)
                    {
                        DetalleActaCliente cliente = new DetalleActaCliente
                        {
                            id_facturacion_safi = detalleCliente.id_facturacion_safi,
                        };
                        db.DetalleActaCliente.Add(cliente);
                        db.SaveChanges();

                        DetallesActa detalle = new DetallesActa
                        {
                            DetalleActaClienteID = Convert.ToInt32(cliente.IDDetalleActaCliente), // Revisar que tipo de dato
                            ActaID = cabecera.IDActa,
                        };

                        db.DetallesActa.Add(detalle);
                        db.SaveChanges();
                    }

                    //Nuevos Detalles
                    foreach (var detalleCliente in cuerpo.DetalleContabilidad)
                    {
                        DetalleActaContabilidad contabilidad = new DetalleActaContabilidad
                        {
                            id_facturacion_safi = detalleCliente.id_facturacion_safi,
                        };
                        db.DetalleActaContabilidad.Add(contabilidad);
                        db.SaveChanges();

                        DetallesActa detalle = new DetallesActa
                        {
                            DetalleActaContabilidadID = Convert.ToInt32(contabilidad.IDDetalleActaContabilidad), // Revisar que tipo de dato
                            ActaID = cabecera.IDActa,
                        };

                        db.DetallesActa.Add(detalle);
                        db.SaveChanges();
                    }

                    #endregion

                    #region Actualizar Pie Pagina Acta
                    //var piePaginaActa = db.ActaInformacionAdicional.Find(piePagina.IDActaInformacionAdicional);
                    //db.Entry(piePaginaActa).State = EntityState.Modified;
                    //db.SaveChanges();

                    var piePaginaActa = db.ActaInformacionAdicional.Find(piePagina.IDActaInformacionAdicional);
                    if (piePaginaActa == null)
                    {
                        return new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeTransaccionFallida };
                    }

                    db.Entry(piePaginaActa).CurrentValues.SetValues(piePagina);
                    db.SaveChanges();

                    #endregion

                    transaction.Commit();
                    return new RespuestaTransaccion { Estado = true, Respuesta = Mensajes.MensajeTransaccionExitosa, EntidadID = cabecera.IDActa };
                }
                catch (Exception ex)
                {
                    transaction.Rollback();
                    return new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeTransaccionFallida + " ;" + ex.Message.ToString() };
                }
            }
        }

        public static RespuestaTransaccion EliminarActa(long IDActa)
        {
            using (var transaction = db.Database.BeginTransaction())
            {
                try
                {
                    var cabeceraActa = db.Acta.Find(IDActa);
                    cabeceraActa.Estado = false;
                    db.Entry(cabeceraActa).State = EntityState.Modified;
                    db.SaveChanges();

                    transaction.Commit();
                    return new RespuestaTransaccion { Estado = true, Respuesta = Mensajes.MensajeTransaccionExitosa };
                }
                catch (Exception ex)
                {
                    transaction.Rollback();
                    return new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeTransaccionFallida + " ;" + ex.Message.ToString() };
                }
            }
        }

        public static ActaInfo ConsultarActa(long IDActa, string codigoActa = null)
        {
            ActaInfo acta = new ActaInfo();
            try
            {
                acta = db.ConsultarActa(IDActa, codigoActa).FirstOrDefault();
                return acta;
            }
            catch (Exception ex)
            {
                return acta;
            }
        }

        public static List<ActaInfo> ConsultarActaInformacionCompleta(long IDActa)
        {
            List<ActaInfo> acta = new List<ActaInfo>();
            try
            {
                acta = db.ConsultarActa(IDActa, null).ToList();
                return acta;
            }
            catch (Exception ex)
            {
                return acta;
            }
        }

        public static List<ActasInformacionGeneralInfo> ListadoActa()
        {
            List<ActasInformacionGeneralInfo> listado = new List<ActasInformacionGeneralInfo>();
            try
            {
                listado = db.ListadoActas().ToList();
                return listado;
            }
            catch (Exception ex)
            {
                return listado;
            }
        }

        public static List<ActaInfo> ListadoActasDetalles()
        {
            List<ActaInfo> listado = new List<ActaInfo>();
            try
            {
                listado = db.ListadoActasDetalles().ToList();
                return listado;
            }
            catch (Exception ex)
            {
                return listado;
            }
        }

        public static List<SecuencialTipoActa> GetTiposActas()
        {
            List<SecuencialTipoActa> listado = new List<SecuencialTipoActa>();
            try
            {
                listado = db.SecuencialTipoActa.ToList();
                return listado;
            }
            catch (Exception ex)
            {
                return listado;
            }
        }

        public static string GetNombreTipo(int tipoActa)
        {
            string nombre = string.Empty;
            try
            {
                var acta = db.SecuencialTipoActa.Where(s => s.IDTipoActa == tipoActa).FirstOrDefault();

                if (acta != null)
                    nombre = acta.NombreTipoActa;

                return nombre;
            }
            catch (Exception ex)
            {
                return nombre;
            }
        }

        public static string GetCodigoTipo(int tipoActa)
        {
            string codigo = string.Empty;
            try
            {
                var acta = db.SecuencialTipoActa.Where(s => s.IDTipoActa == tipoActa).FirstOrDefault();

                if (acta != null)
                    codigo = acta.Codigo;

                return codigo;
            }
            catch (Exception ex)
            {
                return codigo;
            }
        }

        public static IEnumerable<SelectListItem> ObtenerListadoTiposActas(string seleccionado = null)
        {
            List<SelectListItem> listado = new List<SelectListItem>();
            try
            {
                listado = db.SecuencialTipoActa.OrderBy(s => s.NombreTipoActa).Select(c => new SelectListItem
                {
                    Text = c.NombreTipoActa,
                    Value = c.IDTipoActa.ToString()
                }).ToList();


                if (!string.IsNullOrEmpty(seleccionado))
                {
                    if (listado.FirstOrDefault(s => s.Value == seleccionado.ToString()) != null)
                        listado.FirstOrDefault(s => s.Value == seleccionado.ToString()).Selected = true;
                }

                return listado;
            }
            catch (Exception ex)
            {
                return listado;
            }

        }

        public static bool ValidacionRangosFechaInicioFin(DateTime fechaInicio, DateTime fechaFin)
        {
            bool flag = true;
            if (fechaInicio > fechaFin)
                flag = false;
            return flag;
        }

        public static bool ValidacionRangosFechaFinEntrega(DateTime fechaFin, DateTime fechaEntrega)
        {
            bool flag = true;
            if (fechaFin > fechaEntrega)
                flag = false;
            return flag;
        }
        public static bool ValidacionFechaEntrega(DateTime fechaInicio, DateTime fechaEntrega)
        {
            bool flag = true;
            if (fechaInicio > fechaEntrega)
                flag = false;
            return flag;
        }

        public static bool ValidacionHoraInicioFin(string horaInicio, string horaFin)
        {
            TimeSpan hora1 = TimeSpan.Parse(horaInicio);
            TimeSpan hora2 = TimeSpan.Parse(horaFin);


            bool flag = true;
            if (hora1 > hora2)
                flag = false;
            return flag;
        }

        public static string CalcularDuracion(string horaInicio, string horaFin, string dato)
        {
            DateTime hora1 = DateTime.Parse(horaInicio);
            DateTime hora2 = DateTime.Parse(horaFin);

            TimeSpan span = hora2.Subtract(hora1);

            string resultado = span.ToString(@"hh\:mm");

            return resultado + dato;

        }
        public static bool ValidacionHoraInicioFinDiferente(string horaInicio, string horaFin)
        {
            TimeSpan hora1 = TimeSpan.Parse(horaInicio);
            TimeSpan hora2 = TimeSpan.Parse(horaFin);


            bool flag = true;
            if (hora1 == hora2)
                flag = false;
            return flag;
        }

        public static string CalcularHora(string horaInicio)
        {
            DateTime d = DateTime.Parse(horaInicio);
            var resultado = d.ToString("HH:mm");
            return resultado;
        }

        public static string ValidarFechaAcuerdos(List<DetalleActaAcuerdos> acuerdos, DateTime fechaInicio)
        {
            string mensaje = "Revisar los detalles de acuerdos. La fecha {0} no puede ser menor a la Fecha de inicio {1} del acta.";

            var consulta = acuerdos.Where(s => s.Fecha < fechaInicio).FirstOrDefault();

            if (consulta != null)
                mensaje = string.Format(mensaje, consulta.Fecha.ToString("yyyy/MM/dd"), fechaInicio.ToString("yyyy/MM/dd"));
            else
                mensaje = string.Empty;

            return mensaje;
        }

    }
}
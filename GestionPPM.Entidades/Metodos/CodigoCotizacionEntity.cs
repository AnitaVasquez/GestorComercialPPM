using GestionPPM.Entidades.Modelo;
using GestionPPM.Repositorios;
using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.Linq;
using System.Web;

namespace GestionPPM.Entidades.Metodos
{
    public static class CodigoCotizacionEntity
    {
        private static readonly GestionPPMEntities db = new GestionPPMEntities();

        public static RespuestaTransaccion CrearCodigoCotizacion(CodigoCotizacion CodigoCotizacion, int codigoUsuario, List<SublineasNegocioCodigoCotizacionParcial> sublineasNegocio, List<int> idsContactosCotizacion)
        {
            using (var transaction = db.Database.BeginTransaction())
            {

                try
                {
                    CodigoCotizacion.Estado = true;
                    
                    //validar si eta aprobado para coloar creacion ERP si
                    if(CodigoCotizacion.estatus_codigo==68 && CodigoCotizacion.tipo_requerido == 110)
                    {
                        CodigoCotizacion.creacion_safi = 293;
                    }

                    var generacionCodigo = db.GenerarCodigoCotizacion(codigoUsuario).FirstOrDefault();

                    if (generacionCodigo != null)
                    {
                        CodigoCotizacion.codigo_cotizacion = generacionCodigo.CodigoCotizacionGenerado;
                        CodigoCotizacion.id_empresa = 1;
                    }
                    else
                        return new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeCodigoCotizacionError };

                    db.CodigoCotizacion.Add(CodigoCotizacion);
                    db.SaveChanges();

                    foreach (var item in sublineasNegocio)
                    {
                        SublineaNegocioCodigoCotizacion sublinea = new SublineaNegocioCodigoCotizacion
                        {
                            CodigoCatalogoSublineaNegocio = Convert.ToInt32(item.CodigoCatalogoSublineaNegocio),
                            IdCodigoCotizacion = CodigoCotizacion.id_codigo_cotizacion,
                            Valor = Convert.ToDecimal((item.Valor.Replace(".", "")).Replace(",", ".")),
                            Estado = true,
                        };

                        db.SublineaNegocioCodigoCotizacion.Add(sublinea);
                        db.SaveChanges();
                    }

                    foreach (var item in idsContactosCotizacion)
                    {
                        ContactosCodigoCotizacion contacto = new ContactosCodigoCotizacion
                        {
                            idCodigoCotizacion = CodigoCotizacion.id_codigo_cotizacion,
                            idContacto = item
                        };

                        db.ContactosCodigoCotizacion.Add(contacto);
                        db.SaveChanges();
                    }

                    transaction.Commit();
                    return new RespuestaTransaccion { Estado = true, Respuesta = Mensajes.MensajeTransaccionExitosa, CotizacionID = CodigoCotizacion.id_codigo_cotizacion };
                }
                catch (Exception ex)
                {
                    transaction.Rollback();
                    return new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeTransaccionFallida + " ;" + ex.Message.ToString() };
                }
            }
        }

        public static RespuestaTransaccion ActualizarCodigoCotizacion(CodigoCotizacion CodigoCotizacion, List<SublineaNegocioCodigoCotizacionInfo> sublineasNegocio, List<int> idsContactosCotizacion)
        {
            using (var transaction = db.Database.BeginTransaction())
            {
                try
                {
                    // Por si queda el Attach de la entidad y no deja actualizar
                    var local = db.CodigoCotizacion.FirstOrDefault(f => f.id_codigo_cotizacion == CodigoCotizacion.id_codigo_cotizacion);
                    if (local != null)
                    {
                        db.Entry(local).State = EntityState.Detached;
                    }
                    CodigoCotizacion.Estado = true;

                    //validar si eta aprobado para coloar creacion ERP si
                    if (CodigoCotizacion.estatus_codigo == 68 && CodigoCotizacion.tipo_requerido == 110)
                    {
                        CodigoCotizacion.creacion_safi = 293;
                    }
                    db.Entry(CodigoCotizacion).State = EntityState.Modified;
                    db.SaveChanges();

                    // ACTUALIZAR SUBLINEAS DE NEGOCIO CODIGO DE COTIZACION
                    var codigoCotizacionAnterior = db.SublineaNegocioCodigoCotizacion.Where(s => s.IdCodigoCotizacion == CodigoCotizacion.id_codigo_cotizacion).ToList();
                    foreach (var item in codigoCotizacionAnterior)
                    {
                        db.SublineaNegocioCodigoCotizacion.Remove(item);
                        db.SaveChanges();
                    }

                    foreach (var item in sublineasNegocio)
                    {
                        db.SublineaNegocioCodigoCotizacion.Add(new SublineaNegocioCodigoCotizacion
                        {
                            CodigoCatalogoSublineaNegocio = Convert.ToInt32(item.CodigoCatalogoSublineaNegocio),
                            IdCodigoCotizacion = CodigoCotizacion.id_codigo_cotizacion,
                            Valor = Convert.ToDecimal(item.Valor),
                            Estado = true,
                        });
                        db.SaveChanges();
                    }

                    // ACTUALIZAR CONTACTOS CODIGO DE COTIZACION
                    var contactosCodigoCotizacionAnterior = db.ContactosCodigoCotizacion.Where(s => s.idCodigoCotizacion == CodigoCotizacion.id_codigo_cotizacion).ToList();
                    foreach (var item in contactosCodigoCotizacionAnterior)
                    {
                        db.ContactosCodigoCotizacion.Remove(item);
                        db.SaveChanges();
                    }

                    foreach (var item in idsContactosCotizacion)
                    {
                        ContactosCodigoCotizacion contacto = new ContactosCodigoCotizacion
                        {
                            idCodigoCotizacion = CodigoCotizacion.id_codigo_cotizacion,
                            idContacto = item
                        };

                        db.ContactosCodigoCotizacion.Add(contacto);
                        db.SaveChanges();
                    }

                    transaction.Commit();
                    return new RespuestaTransaccion { Estado = true, Respuesta = Mensajes.MensajeTransaccionExitosa, CotizacionID = CodigoCotizacion.id_codigo_cotizacion };
                }
                catch (Exception ex)
                {
                    transaction.Rollback();
                    return new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeTransaccionFallida + " ;" + ex.Message.ToString() };
                }
            }
        }


        public static RespuestaTransaccion CrearCodigoCotizacionCargaMasiva(List<CodigoCotizacionExcel> codigos)
        {
            using (var transaction = db.Database.BeginTransaction())
            {

                try
                {
                    var user = HttpContext.Current.Session["usuario"];
                    var usuarioID = Convert.ToInt16(user);
                    var datosUsuario = UsuarioEntity.ConsultarUsuario(usuarioID);
                    var codigoUsuario = datosUsuario.secu_usua;

                    foreach (var item in codigos)
                    {

                        //Validacion de Mínimo un contacto y una Sublínea de Negocio
                        var sublineas = item.SublineasNegocio;
                        var contactos = item.Contactos;

                        if (!sublineas.Any() || !contactos.Any())
                            return new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeValidacionCodigoCotizacionCargaMasiva };

                        var generacionCodigo = db.GenerarCodigoCotizacion(codigoUsuario).FirstOrDefault();

                        if (generacionCodigo != null)
                        {
                            item.Estado = true;
                            item.codigo_cotizacion = generacionCodigo.CodigoCotizacionGenerado;
                            item.id_empresa = 1;
                        }
                        else
                            return new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeCodigoCotizacionError };

                        var registro = new CodigoCotizacion
                        {
                            //id_codigo_cotizacion = item,
                            codigo_cotizacion = item.codigo_cotizacion,
                            fecha_cotizacion = item.fecha_cotizacion,
                            responsable = item.responsable,
                            estatus_codigo = item.estatus_codigo,
                            id_cliente = item.id_cliente,
                            ejecutivo = item.ejecutivo,
                            tipo_requerido = item.tipo_requerido,
                            tipo_intermediario = item.tipo_intermediario,
                            tipo_proyecto = item.tipo_proyecto,
                            dimension_proyecto = item.dimension_proyecto,
                            aplica_contrato = item.aplica_contrato,
                            forma_pago = item.forma_pago,
                            forma_pago_1 = item.forma_pago_1,
                            forma_pago_2 = item.forma_pago_2,
                            forma_pago_3 = item.forma_pago_3,
                            forma_pago_4 = item.forma_pago_4,
                            etapa_cliente = item.etapa_cliente,
                            etapa_general = item.etapa_general,
                            estatus_detallado = item.estatus_detallado,
                            estatus_general = item.estatus_general,
                            tipo_producto_PtoP = item.tipo_producto_PtoP,
                            tipo_plan = item.tipo_plan,
                            tipo_tarifa = item.tipo_tarifa,
                            tipo_migracion = item.tipo_migracion,
                            tipo_etapa_PtoP = item.tipo_etapa_PtoP,
                            tipo_subsidio = item.tipo_subsidio,
                            nombre_proyecto = item.nombre_proyecto,
                            descripcion_proyecto = item.descripcion_proyecto,
                            tipo_fee = item.tipo_fee,
                            creacion_safi = item.creacion_safi,
                            facturable = item.facturable,
                            area_departamento_usuario = item.area_departamento_usuario,
                            pais = item.pais,
                            ciudad = item.ciudad,
                            direccion = item.direccion,
                            id_empresa = item.id_empresa,
                            estado_cliente = item.estado_cliente,
                            tipo_cliente = item.tipo_cliente,
                            tipo_zoho = item.tipo_zoho,
                            referido = item.referido,
                            Estado = item.Estado,
                        };

                        db.CodigoCotizacion.Add(registro);
                        db.SaveChanges();

                        foreach (var obj in sublineas)
                        {
                            //Verificar si el contacto es de facturación
                            if (!CatalogoEntity.VerificarSublineaNegocio(obj.CodigoCatalogoSublineaNegocio))
                                return new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeValidacionCodigoCotizacionSublineaNegocioCargaMasiva + " Sublínea de Negocio: " + obj.CodigoCatalogoSublineaNegocio + " ." };

                            SublineaNegocioCodigoCotizacion sublinea = new SublineaNegocioCodigoCotizacion
                            {
                                CodigoCatalogoSublineaNegocio = Convert.ToInt32(obj.CodigoCatalogoSublineaNegocio),
                                IdCodigoCotizacion = registro.id_codigo_cotizacion,
                                Valor = obj.Valor,
                                Estado = true,
                            };

                            db.SublineaNegocioCodigoCotizacion.Add(sublinea);
                            db.SaveChanges();
                        }

                        foreach (var obj in contactos)
                        {
                            //Verificar si el contacto es de facturación
                            if (!ContactoClienteEntity.VerificarContactoFacturacion(obj.idContacto.Value))
                                return new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeValidacionCodigoCotizacionContactoFacturacionCargaMasiva + ". Contacto: " + obj.idContacto };

                            ContactosCodigoCotizacion contacto = new ContactosCodigoCotizacion
                            {
                                idCodigoCotizacion = registro.id_codigo_cotizacion,
                                idContacto = obj.idContacto
                            };

                            db.ContactosCodigoCotizacion.Add(contacto);
                            db.SaveChanges();
                        }

                    }



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


        public static RespuestaTransaccion ActualizarStatusCodigoCotizacion(CodigoCotizacion codigo)
        {
            try
            {
                var CodigoCotizacion = db.CodigoCotizacion.Find(codigo.id_codigo_cotizacion);

                CodigoCotizacion.estatus_codigo = codigo.estatus_codigo;

                db.Entry(CodigoCotizacion).State = EntityState.Modified;
                db.SaveChanges();

                return new RespuestaTransaccion { Estado = true, Respuesta = Mensajes.MensajeTransaccionExitosa };
            }
            catch (Exception ex)
            {
                return new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeTransaccionFallida + " ;" + ex.Message.ToString() };
            }
        }

        public static RespuestaTransaccion ActualizarStatusCodigoCotizacionClienteExterno(int codigo, int estatus_codigo)
        {
            try
            {
                var CodigoCotizacion = db.CodigoCotizacion.Find(codigo);

                CodigoCotizacion.estatus_codigo = estatus_codigo;

                db.Entry(CodigoCotizacion).State = EntityState.Modified;
                db.SaveChanges();

                return new RespuestaTransaccion { Estado = true, Respuesta = Mensajes.MensajeTransaccionExitosa };
            }
            catch (Exception ex)
            {
                return new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeTransaccionFallida + " ;" + ex.Message.ToString() };
            }
        }

        //Eliminación Lógica
        public static RespuestaTransaccion EliminarCodigoCotizacion(int id)
        {
            try
            {
                var CodigoCotizacion = db.CodigoCotizacion.Find(id);

                CodigoCotizacion.Estado = false;

                db.Entry(CodigoCotizacion).State = EntityState.Modified;
                db.SaveChanges();

                return new RespuestaTransaccion { Estado = true, Respuesta = Mensajes.MensajeTransaccionExitosa };
            }
            catch (Exception ex)
            {
                return new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeTransaccionFallida + " ;" + ex.Message.ToString() };
            }
        }

        public static List<CodigoCotizacionInfo> ListarCodigoCotizacion()
        {
            try
            {
                return db.ListadoCodigoCotizacion().ToList();
            }
            catch (Exception)
            {
                throw;
            }
        }

        public static List<SublineaNegocioCodigoCotizacionInfo> ListarSublineaCodigoCotizacion()
        {
            try
            {
                return db.ListadoSublineaNegocioCodigoCotizacion().ToList();
            }
            catch (Exception)
            {
                throw;
            }
        }

        public static CodigoCotizacion ConsultarCodigoCotizacion(int id)
        {
            try
            {
                using (var context = new GestionPPMEntities())
                {
                    CodigoCotizacion codigoCotizacion = context.CodigoCotizacion.Find(id);
                    return codigoCotizacion;
                }
            }
            catch (Exception ex)
            {
                throw;
            }
        }

        public static CodigoCotizacion ConsultarCodigoCotizacionValoresDefault(string nombreProyecto, string descripcionProyecto)
        {
            try
            {
                CodigoCotizacion codigoCotizacion = new CodigoCotizacion
                {
                    //id_codigo_cotizacion = "",
                    //codigo_cotizacion = "",
                    fecha_cotizacion = DateTime.Now,
                    responsable = null,
                    //estatus_codigo = 68,
                    id_cliente = null,
                    ejecutivo = null,
                    tipo_requerido = 110,
                    tipo_intermediario = null,
                    tipo_proyecto = 60,
                    dimension_proyecto = 63,
                    aplica_contrato = false,
                    forma_pago = false,
                    forma_pago_1 = 100,
                    forma_pago_2 = 0,
                    forma_pago_3 = 0,
                    forma_pago_4 = 0,
                    //etapa_cliente = 183,
                    etapa_general = 215,
                    estatus_detallado = 215,
                    estatus_general = 231,
                    tipo_producto_PtoP = 77,
                    tipo_plan = 82,
                    tipo_tarifa = 91,
                    tipo_migracion = 95,
                    tipo_etapa_PtoP = 107,
                    tipo_subsidio = 292,
                    nombre_proyecto = nombreProyecto,
                    descripcion_proyecto = descripcionProyecto,
                    tipo_fee = 151,
                    //creacion_safi = 293,
                    facturable = true,
                    area_departamento_usuario = string.Empty,
                    pais = null,
                    ciudad = null,
                    direccion = null,
                    id_empresa = null,
                    estado_cliente = false,
                    tipo_cliente = null,
                    tipo_zoho = 30,
                    referido = 20,
                    Estado = true,
                };
                return codigoCotizacion;
            }
            catch (Exception)
            {
                throw;
            }
        }

        public static bool VerificarExistenciaCodigoCotizacion(string codigoCotizacion)
        {
            try
            {
                CodigoCotizacion codigo = db.CodigoCotizacion.FirstOrDefault(s => s.codigo_cotizacion == codigoCotizacion);
                if (codigo == null)
                    return true;
                else
                    return false;
            }
            catch (Exception)
            {
                return false;
            }
        }

        public static CodigoCotizacionInfo ConsultarCodigoCotizacionCompleto(int id)
        {
            try
            {
                CodigoCotizacionInfo codigoCotizacion = db.ListadoCodigoCotizacion().FirstOrDefault(s => s.id_codigo_cotizacion == id);
                return codigoCotizacion;
            }
            catch (Exception)
            {
                throw;
            }
        }

        public static bool ValidarPagosCodigoCotizacion(CodigoCotizacion codigoCotizacion)
        {
            bool flag = true;

            if (!codigoCotizacion.forma_pago.Value)
            {
                codigoCotizacion.forma_pago_1 = 100;
                codigoCotizacion.forma_pago_2 = 0;
                codigoCotizacion.forma_pago_3 = 0;
                codigoCotizacion.forma_pago_4 = 0;
            }

            var totalPagos = codigoCotizacion.forma_pago_1 + codigoCotizacion.forma_pago_2 + codigoCotizacion.forma_pago_3 + codigoCotizacion.forma_pago_4;

            if (totalPagos != 100)
                flag = false;

            return flag;
        }

        public static bool ValidarPagosCodigoCotizacion2(CodigoCotizacion codigoCotizacion)
        {
            bool flag = true;

            var totalPagos = codigoCotizacion.forma_pago_1 + codigoCotizacion.forma_pago_2 + codigoCotizacion.forma_pago_3 + codigoCotizacion.forma_pago_4;

            if (totalPagos != 100)
                flag = false;

            return flag;
        }

        public static bool ValidarPagosCodigoCotizacion3(CodigoCotizacion codigoCotizacion)
        {
            bool flag = true;

            if (!codigoCotizacion.forma_pago.Value)
            {
                if (codigoCotizacion.forma_pago_1 == 100 && codigoCotizacion.forma_pago_2 == 0 && codigoCotizacion.forma_pago_3 == 0 && codigoCotizacion.forma_pago_4 == 0)
                    return true;
                else
                    return false;
            }
            else
            {
                var totalPagos = codigoCotizacion.forma_pago_1 + codigoCotizacion.forma_pago_2 + codigoCotizacion.forma_pago_3 + codigoCotizacion.forma_pago_4;

                if (totalPagos != 100)
                    flag = false;

            }

            return flag;
        }

        public static bool ValidarPagos1(CodigoCotizacion codigoCotizacion)
        {
            bool flag = true;

            if (codigoCotizacion.forma_pago_1 == 0)
            {
                flag = false;
            }

            return flag;
        }

        public static List<SublineaNegocioCodigoCotizacionInfo> ListarSublineaNegocioCodigoCotizacion(int idCodigoCotizacion)
        {
            try
            {
                var listado = db.ListadoSublineaNegocioCodigoCotizacion().Where(s => s.IdCodigoCotizacion == idCodigoCotizacion).ToList();
                return listado;
            }
            catch (Exception)
            {

                throw;
            }

        }

        public static CodigoCotizacion ConsultarCodigoCotizacionValoresDefault(string nombreProyecto, string descripcionProyecto, int? idCliente = null)
        {
            try
            {
                CodigoCotizacion codigoCotizacion = new CodigoCotizacion
                {
                    //id_codigo_cotizacion = "",
                    //codigo_cotizacion = "",
                    fecha_cotizacion = DateTime.Now,
                    responsable = null,
                    //estatus_codigo = 68,
                    id_cliente = idCliente,// null,
                    ejecutivo = null,
                    tipo_requerido = 110,
                    tipo_intermediario = null,
                    tipo_proyecto = 60,
                    dimension_proyecto = 63,
                    aplica_contrato = false,
                    forma_pago = false,
                    forma_pago_1 = 100,
                    forma_pago_2 = 0,
                    forma_pago_3 = 0,
                    forma_pago_4 = 0,
                    //etapa_cliente = 183,
                    etapa_general = 215,
                    estatus_detallado = 215,
                    estatus_general = 231,
                    tipo_producto_PtoP = 77,
                    tipo_plan = 82,
                    tipo_tarifa = 91,
                    tipo_migracion = 95,
                    tipo_etapa_PtoP = 107,
                    tipo_subsidio = 292,
                    nombre_proyecto = nombreProyecto,
                    descripcion_proyecto = descripcionProyecto,
                    tipo_fee = 151,
                    //creacion_safi = 293,
                    facturable = true,
                    area_departamento_usuario = string.Empty,
                    pais = null,
                    ciudad = null,
                    direccion = null,
                    id_empresa = null,
                    estado_cliente = false,
                    tipo_cliente = null,
                    tipo_zoho = 30,
                    referido = 20,
                    Estado = true,
                };
                return codigoCotizacion;
            }
            catch (Exception)
            {
                throw;
            }
        }
    }
}
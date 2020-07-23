using GestionPPM.Entidades.Modelo;
using GestionPPM.Repositorios;
using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.Linq;
using System.Web.Mvc;

namespace GestionPPM.Entidades.Metodos
{
    public static class ComercioEntity
    {
        private static readonly GestionPPMEntities db = new GestionPPMEntities();
         
        public static RespuestaTransaccion CrearComercio(Comercio comercio)
        {
            using (var transaction = db.Database.BeginTransaction())
            {
                try
                {
                    //Pasar a mayusculas el nombre del comercio
                    if (comercio.nombre_comercio != null)
                    {
                        comercio.nombre_comercio = comercio.nombre_comercio.ToUpper();
                    } 
                    
                    //Pasar a mayusculas el nombre del comercio
                    if (comercio.id_comercio_predictive != null)
                    {
                        comercio.id_comercio_predictive = comercio.id_comercio_predictive.ToUpper();
                    }
                    //Validar si el estado del comercio en base a la fecha de salida
                    if (comercio.fecha_inactivo == null)
                    { 
                        comercio.estado_comercio = true;
                    }
                    else
                    {
                        comercio.estado_comercio = false;
                    }

                    db.Comercio.Add(comercio);
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

        public static RespuestaTransaccion ActualizarComercio(Comercio comercio)
        {
            using (var transaction = db.Database.BeginTransaction())
            {
                try
                {
                    // Por si queda el Attach de la entidad y no deja actualizar
                    var local = db.Comercio.FirstOrDefault(c => c.id_comercio == comercio.id_comercio);
                    if (local != null)
                    {
                        db.Entry(local).State = EntityState.Detached;
                    }

                    //Pasar a mayusculas el nombre del comercio
                    if (comercio.nombre_comercio != null)
                    {
                        comercio.nombre_comercio = comercio.nombre_comercio.ToUpper();
                    }

                    //Pasar a mayusculas el nombre del comercio
                    if (comercio.id_comercio_predictive != null)
                    {
                        comercio.id_comercio_predictive = comercio.id_comercio_predictive.ToUpper();
                    }
                    //Validar si el estado del comercio en base a la fecha de salida
                    if (comercio.fecha_inactivo == null)
                    {
                        comercio.estado_comercio = true;
                    }
                    else
                    {
                        comercio.estado_comercio = false;
                    }

                    db.Entry(comercio).State = EntityState.Modified;
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
         
        //Eliminación Lógica
        public static RespuestaTransaccion EliminarComercio(int id)
        {
            try
            {
                var Cliente = db.Cliente.Find(id);

                if (Cliente.estado_cliente == true)
                {
                    Cliente.estado_cliente = false;
                }
                else
                {
                    Cliente.estado_cliente = true;
                }

                db.Entry(Cliente).State = EntityState.Modified;
                db.SaveChanges();

                return new RespuestaTransaccion { Estado = true, Respuesta = Mensajes.MensajeTransaccionExitosa };
            }
            catch (Exception ex)
            {
                return new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeTransaccionFallida + " ;" + ex.Message.ToString() };
            }
        }

        public static List<ListadoComercios> ListarComercios()
        {
            try
            {
                return db.ListadoComercios().ToList();
            }
            catch (Exception)
            {
                throw;
            }
        }

        public static IEnumerable<SelectListItem> ListarComerciosPTOP()
        {
            List<SelectListItem> ListadoComercios = new List<SelectListItem>();
            try
            {
                ListadoComercios = db.ListadoComercios().Select(c => new SelectListItem
                {
                    Text = c.ID_Comercio,
                    Value = c.Codigo.ToString(),
                }).ToList();
                 
                return ListadoComercios;
            }
            catch (Exception ex)
            {
                return ListadoComercios;
            }

        }

        public static Comercio ConsultarComercio(int id)
        {
            try
            {
                Comercio comercio = db.Comercio.Find(id);
                return comercio;
            }
            catch (Exception)
            {
                throw;
            }
        }

        public static RespuestaTransaccion CrearActualizarComerciosCargaMasiva(List<ComercioExcel> Comercios)
        {
            using (var transaction = db.Database.BeginTransaction())
            {
                try
                {
                    foreach (var item in Comercios)
                    {
                        if (item.id_comercio != null)
                        {
                            item.id_comercio = item.id_comercio.ToUpper();
                        }

                        if (item.comercio != null)
                        {
                            item.comercio = item.comercio.ToUpper();
                        }

                        //verificar si existe para crear o actualizar
                        var existeComercio = db.Comercio.Where(s => s.id_comercio_predictive == item.id_comercio).FirstOrDefault();

                        if (existeComercio != null) // Actualiza el cliente
                        {
                            //obtener los datos del Comercio a Actualiza
                            var comercioActualizar = db.Comercio.Find(existeComercio.id_comercio);

                            //setear datos nuevos
                            comercioActualizar.nombre_comercio = item.comercio.ToUpper();

                            //Obtener el codigo del cliente 
                            var cliente = ClienteEntity.ConsultarClienteLineaNegocio("Place to Pay", item.ruc_cliente);
                            if (cliente != null)
                            {
                                comercioActualizar.id_cliente = cliente.id_cliente;
                            }
                            comercioActualizar.fecha_salida_produccion = Convert.ToDateTime(item.fecha_salida_produccion);

                            //Validar si el estado del comercio en base a la fecha de salida
                            if (item.fecha_inactivo == null || item.fecha_inactivo == "")
                            {
                                comercioActualizar.fecha_inactivo = null;
                                comercioActualizar.estado_comercio = true;
                            }
                            else
                            {
                                comercioActualizar.fecha_inactivo = Convert.ToDateTime(item.fecha_inactivo);
                                comercioActualizar.estado_comercio = false;
                            }

                            //Estatus contrato                       
                            var estatus = CatalogoEntity.ListadoCatalogosPorCodigo("ECP-01").FirstOrDefault(c => c.Text == item.estatus_contrato.ToUpper());
                            comercioActualizar.id_estatus_contrato = Convert.ToInt32(estatus.Value.ToString());

                            //Tipo Subsidio                   
                            var tipo = CatalogoEntity.ListadoCatalogosPorCodigo("TSP-01").FirstOrDefault(c => c.Text == item.tipo_subsidio.ToUpper());
                            comercioActualizar.id_tipo_subsidio = Convert.ToInt32(tipo.Value.ToString()); 

                            //Agrupacion                       
                            var agrupacion = CatalogoEntity.ListadoCatalogosPorCodigo("AGP-01").FirstOrDefault(c => c.Text == item.agrupacion.ToUpper());
                            comercioActualizar.id_agrupacion = Convert.ToInt32(agrupacion.Value.ToString());

                            //Compartido PTOP                      
                            if(item.compartido_ptop.ToUpper() == "SI")
                            {
                                comercioActualizar.compartido_ptop = true;
                            }
                            else
                            {
                                comercioActualizar.compartido_ptop = false;
                            }
                            //Porcentaje Decimal                        
                            comercioActualizar.descuento = Convert.ToDecimal(item.descuento);

                            //Porcentaje Iva                        
                            comercioActualizar.porcentaje_iva = Convert.ToDecimal(item.porcentaje_iva);

                            //Desceunto en porcentaje 
                            if (item.cobro_porcentaje.ToUpper() == "SI")
                            {
                                comercioActualizar.cobro_porcentaje = true;
                            }
                            else
                            {
                                comercioActualizar.cobro_porcentaje = false;
                            }

                            db.Entry(comercioActualizar).State = EntityState.Modified;
                            db.SaveChanges(); 
                        }
                        else
                        { 
                            // Guarda el comercio                              
                            var comercioNuevo = new Comercio
                            {
                                id_comercio_predictive = item.id_comercio.ToUpper(),
                                nombre_comercio = item.comercio.ToUpper(), 
                            };

                            //Obtener el codigo del cliente 
                            var cliente = ClienteEntity.ConsultarClienteLineaNegocio("Place to Pay", item.ruc_cliente);
                            if (cliente != null)
                            {
                                comercioNuevo.id_cliente = cliente.id_cliente;
                            }
                            comercioNuevo.fecha_salida_produccion = Convert.ToDateTime(item.fecha_salida_produccion);

                            //Validar si el estado del comercio en base a la fecha de salida
                            if (item.fecha_inactivo == null || item.fecha_inactivo == "")
                            {
                                comercioNuevo.fecha_inactivo = null;
                                comercioNuevo.estado_comercio = true;
                            }
                            else
                            {
                                comercioNuevo.fecha_inactivo = Convert.ToDateTime(item.fecha_inactivo);
                                comercioNuevo.estado_comercio = false;
                            }

                            //Estatus contrato                       
                            var estatus = CatalogoEntity.ListadoCatalogosPorCodigo("ECP-01").FirstOrDefault(c => c.Text == item.estatus_contrato.ToUpper());
                            comercioNuevo.id_estatus_contrato = Convert.ToInt32(estatus.Value.ToString());

                            //Tipo Subsidio                   
                            var tipo = CatalogoEntity.ListadoCatalogosPorCodigo("TSP-01").FirstOrDefault(c => c.Text == item.tipo_subsidio.ToUpper());
                            comercioNuevo.id_tipo_subsidio = Convert.ToInt32(tipo.Value.ToString());

                            //Agrupacion                       
                            var agrupacion = CatalogoEntity.ListadoCatalogosPorCodigo("AGP-01").FirstOrDefault(c => c.Text == item.agrupacion.ToUpper());
                            comercioNuevo.id_agrupacion = Convert.ToInt32(agrupacion.Value.ToString());

                            //Compartido PTOP                      
                            if (item.compartido_ptop.ToUpper() == "SI")
                            {
                                comercioNuevo.compartido_ptop = true;
                            }
                            else
                            {
                                comercioNuevo.compartido_ptop = false;
                            }

                            //Descuento en porcentaje 
                            if (item.cobro_porcentaje.ToUpper() == "SI")
                            {
                                comercioNuevo.cobro_porcentaje = true;
                            }
                            else
                            {
                                comercioNuevo.cobro_porcentaje = false;
                            }

                            //Porcentaje Decimal                        
                            comercioNuevo.descuento = Convert.ToDecimal(item.descuento);

                            //Porcentaje Iva                        
                            comercioNuevo.porcentaje_iva = Convert.ToDecimal(item.porcentaje_iva);

                            db.Comercio.Add(comercioNuevo);
                            db.SaveChanges();                          
                        }
                    }

                    transaction.Commit();
                    return new RespuestaTransaccion { Estado = true, Respuesta = Mensajes.MensajeTransaccionExitosa };
                }
                catch (Exception ex)
                {
                    //transaction.Rollback();
                    return new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeTransaccionFallida + " ;" + ex.Message.ToString() };
                }
            }
        }

    }
}
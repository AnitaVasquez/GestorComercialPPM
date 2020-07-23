using GestionPPM.Entidades.Modelo;
using GestionPPM.Repositorios;
using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace GestionPPM.Entidades.Metodos
{
    public static class ProyectosAsigandosEntity
    {
        private static readonly GestionPPMEntities db = new GestionPPMEntities();

        public static RespuestaTransaccion CrearAvanceProyecto(DetalleProyectosAsignados detalleAvance, int id_codigo_cotizacion, int proyecto, int etapa_cliente, int etapa_general, int estatus_detallado, int estatus_general, DateTime fecha_inicio_programado, DateTime fecha_fin_programado, DateTime fecha_inicio_real, DateTime fecha_fin_real, int horas_programadas, int horas_reales)
        {
            using (var transaction = db.Database.BeginTransaction())
            { 
                try
                {
                    // Guardando etapas cliente
                    var codigoCotizacion = db.CodigoCotizacion.Find(id_codigo_cotizacion);
                    codigoCotizacion.etapa_cliente = etapa_cliente;
                    codigoCotizacion.etapa_general = etapa_general;
                    codigoCotizacion.estatus_detallado = estatus_detallado;
                    codigoCotizacion.estatus_general = estatus_general;
                    db.Entry(codigoCotizacion).State = EntityState.Modified;
                    db.SaveChanges();

                    //validar si tiene el registro del proyecto
                    if(proyecto == 0)
                    {
                        //crear la cabecera
                        ProyectosAsignados cabeceraProyecto = new ProyectosAsignados();
                        cabeceraProyecto.id_codigo_cotizacion = id_codigo_cotizacion;
                        cabeceraProyecto.numero_horas_programado = horas_programadas;
                        cabeceraProyecto.fecha_inicio_programado = fecha_inicio_programado;
                        cabeceraProyecto.fecha_fin_programado = fecha_fin_programado;
                        if(fecha_inicio_real.ToString() != "01/01/1990 00:00:00")
                        { 
                            cabeceraProyecto.fecha_inicio_real = fecha_inicio_real;
                        }
                        if (fecha_fin_real.ToString() != "01/01/1990 00:00:00")
                        {
                            cabeceraProyecto.fecha_fin_real = fecha_fin_real;
                        }
                        cabeceraProyecto.estado_proyecto = true;

                        db.ProyectosAsignados.Add(cabeceraProyecto);
                        db.SaveChanges();

                        proyecto = cabeceraProyecto.id_proyecto;
                    }
                     
                    //Realizar el registro del avance
                    detalleAvance.id_proyecto = proyecto;
                    db.DetalleProyectosAsignados.Add(detalleAvance);
                    db.SaveChanges();

                    //Actualizar el porcentaje de cumpliento en el proyecto general
                    var proyectoAsignado = db.ProyectosAsignados.Find(proyecto);
                    proyectoAsignado.porcentaje_cumplimiento = detalleAvance.porcentaje_avance;
                    proyectoAsignado.fecha_ultimo_avance = detalleAvance.fecha_avance;
                    if (fecha_inicio_real.ToString() != "01/01/1990 00:00:00")
                    {
                        proyectoAsignado.fecha_inicio_real = fecha_inicio_real;
                    }
                    if (fecha_fin_real.ToString() != "01/01/1990 00:00:00")
                    {
                        proyectoAsignado.fecha_fin_real = fecha_fin_real;
                    }
                    proyectoAsignado.numero_horas_real = horas_reales;

                    db.Entry(proyectoAsignado).State = EntityState.Modified;
                    db.SaveChanges();

                    //enviar notificaciones
                    db.usp_guarda_envio_correo_notificaciones(2, detalleAvance.id_detalle_proyecto, "", 1, "");
                    db.SaveChanges();

                    transaction.Commit();
                    return new RespuestaTransaccion { Estado = true, Respuesta = Mensajes.MensajeTransaccionExitosa};
                }
                catch (Exception ex)
                {
                    transaction.Rollback();
                    return new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeTransaccionFallida + " ;" + ex.Message.ToString() };
                }
            }
        }
               
        //Metodo para listar proyectos asigandos ativos
        public static IEnumerable<SelectListItem> ObtenerListadoProyectos(int? id = null, string seleccionado = null)
        {
            var listadoProyectos = db.ListadoProyectos(id).Select(x => new SelectListItem
            {
                Text = x.codigo_cotizacion + " - " + x.nombre_proyecto,
                Value = x.id_codigo_cotizacion.ToString()
            }).ToList();

            if (!string.IsNullOrEmpty(seleccionado))
            {
                if (listadoProyectos.FirstOrDefault(s => s.Value == seleccionado.ToString()) != null)
                    listadoProyectos.FirstOrDefault(s => s.Value == seleccionado.ToString()).Selected = true;
            }

            return listadoProyectos;
        }

        //Metodo para devolver los datos del codigo cotizacion
        public static DatosProyecto ConsultarDatosCodigoCotizacion(int? id)
        {
            try
            {
                DatosProyecto datosProyecto = db.DatosProyecto(id).First();
                return datosProyecto;
            }
            catch (Exception ex)
            {
                DatosProyecto datosProyecto = new DatosProyecto();
                return datosProyecto; 
            }
        }

        //Metodo para devolver los datos del codigo cotizacion
        public static DetalleUltimoAvance ConsultarUltimoAvance(int? id)
        {
            try
            {
                DetalleUltimoAvance ultimoAvance = db.DetalleUltimoAvance(id).First();
                return ultimoAvance;
            }
            catch (Exception ex)
            {
                throw;
            }
        }

        //Metodo para devolver los datos del codigo cotizacion
        public static DatosProyectoCotizacion ConsultarProyectoPorCodigoCotizacion(int? id)
        {
            try
            {
                DatosProyectoCotizacion datosProyecto = db.DatosProyectoCotizacion(id).First();
                return datosProyecto;
            }
            catch (Exception ex)
            {
                DatosProyectoCotizacion datosProyecto = new DatosProyectoCotizacion();
                ex.ToString();
                return datosProyecto;
            }
        }

        //Listar Avances
        public static List<ListadoAvanceProyectoUsuario> ListarProyectos(int usuario)
        {
            try
            {
                return db.ListadoAvanceProyectoUsuario(usuario).ToList();
            }
            catch (Exception)
            {
                throw;
            }
        }

        //Listar Avances Totales
        public static List<ListadoAvanceProyectosTotal> ListarProyectosTotales()
        {
            try
            {
                return db.ListadoAvanceProyectosTotal().ToList();
            }
            catch (Exception)
            {
                throw;
            }
        }

        public static List<ListadoProyectos> ListadoProyectos(int id)
        {
            try
            {
                return db.ListadoProyectos(id).ToList();
            }
            catch (Exception)
            {
                throw;
            }
        }

    }
}
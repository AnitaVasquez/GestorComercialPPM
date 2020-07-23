using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace GestionPPM.Entidades.Modelo
{
    public class ActaCompleta
    {
        public ActaCompleta(List<ActaInfo> InformacionActaCompleta)
        {
            var acta = InformacionActaCompleta.FirstOrDefault();
            Cabecera = new ActaInfo
            {
                IDActa = acta.IDActa,
                CodigoActa = acta.CodigoActa,
                FechaCreacion = acta.FechaCreacion,
                ElaboradoPor = acta.ElaboradoPor,
                NombresElaboradoPor = acta.NombresElaboradoPor,
                FechaInicio = acta.FechaInicio,
                FechaFin = acta.FechaFin,
                FechaEntrega = acta.FechaEntrega,
                Lugar = acta.Lugar,
                NumeroReunion = acta.NumeroReunion,
                HoraInicio = acta.HoraInicio,
                Cargo = acta.Cargo,
                HoraFin = acta.HoraFin,
                ReferenciaCliente = acta.ReferenciaCliente,
                AlcanceObjetivo = acta.AlcanceObjetivo,
                FacilitadorModerador = acta.FacilitadorModerador,
                CodigoCotizacionID = acta.CodigoCotizacionID,
                CodigoCotizacion = acta.CodigoCotizacion,
                Cliente = acta.Cliente,
                NombreProyecto = acta.NombreProyecto,
                DescripcionProyecto = acta.DescripcionProyecto,
                Observaciones = acta.Observaciones,
                Suspendida = acta.Suspendida,
                TipoActaID = acta.TipoActaID,
                Estado = true,//acta.Estado,
                CodigoTipoActa = acta.CodigoTipoActa
            };

            Cuerpo = GetDetallesActaAcuerdos(InformacionActaCompleta);

            PiePagina = new ActaInformacionAdicional
            {
                IDActaInformacionAdicional = acta.IDActaInformacionAdicional.Value,
                ActaID = acta.IDActa,
                AcuerdoConformidad = acta.AcuerdoConformidad,
                Firmas = acta.Firmas,
            };
        }

        public ActaInfo Cabecera { get; set; }
        public DetallesActaParcial Cuerpo { get; set; }
        public ActaInformacionAdicional PiePagina { get; set; }

        public DetallesActaParcial GetDetallesActaAcuerdos(List<ActaInfo> ActaDetalles)
        {
            DetallesActaParcial listadoCuerpo = new DetallesActaParcial();
            try
            {
                foreach (var item in ActaDetalles)
                {
                    if (item.DetalleActaAcuerdosID.HasValue)
                    {
                        listadoCuerpo.Acuerdos.Add(new DetalleActaAcuerdos
                        {
                            IDDetalleActaAcuerdos = item.DetalleActaAcuerdosID.Value,
                            Acuerdo = item.Acuerdo,
                            Fecha = item.Fecha.Value,
                            Responsable = item.ResponsableAcuerdo,
                        });

                    }
                    if (item.DetalleActaCondicionesGeneralesID.HasValue)
                    {
                        listadoCuerpo.CondicionesGenerales.Add(new DetalleActaCondicionesGenerales
                        {
                            IDActaCondicionesGenerales = item.DetalleActaCondicionesGeneralesID.Value,
                            Condicion = item.Condicion,
                        });
                    }
                    if (item.DetalleActaEntregablesID.HasValue)
                    {
                        listadoCuerpo.Entregables.Add(new DetalleActaEntregables
                        {
                            IDDetalleActaEntregables = item.DetalleActaEntregablesID.Value,
                            Entregable = item.Entregable,
                            Tipo = item.Tipo,
                        });
                    }
                    if (item.DetalleActaParticipantesID.HasValue)
                    {
                        listadoCuerpo.Participantes.Add(new DetalleActaParticipantes
                        {
                            IDDetalleActaParticipantes = item.DetalleActaParticipantesID.Value,
                            Nombres = item.NombresParticipante,
                            Presente = item.Presente.Value,
                        });
                    }
                    if (item.DetalleActaResponsablesClienteID.HasValue)
                    {
                        listadoCuerpo.Responsables.Add(new DetalleActaResponsables
                        {
                            IDDetalleActaResponsablesCliente = item.DetalleActaResponsablesClienteID.Value,
                            Nombres = item.NombresResponsable,
                            Rol = item.Rol,
                            Empresa = item.Empresa,
                        });
                    }
                    if (item.DetalleActaTemasTratarID.HasValue)
                    {
                        listadoCuerpo.Temas.Add(new DetalleActaTemasTratar
                        {
                            IDDetalleActaTemasTratar = item.DetalleActaTemasTratarID.Value,
                            Tema = item.Tema,
                            Responsable = item.ResponsableTema
                        });
                    }

                    //Nuevos Detalles
                    if (item.DetalleActaClienteID.HasValue)
                    {
                        listadoCuerpo.DetalleCliente.Add(new DetalleActaCliente
                        {
                            IDDetalleActaCliente = item.DetalleActaClienteID.Value,
                            id_facturacion_safi = item.id_facturacion_safi_ActaCliente.Value,
                        });
                    }
                    if (item.DetalleActaContabilidadID.HasValue)
                    {
                        listadoCuerpo.DetalleContabilidad.Add(new DetalleActaContabilidad
                        {
                            IDDetalleActaContabilidad = item.DetalleActaContabilidadID.Value,
                            id_facturacion_safi = item.id_facturacion_safi_ActaContabilidad.Value,
                        });
                    }

                }
                return listadoCuerpo;
            }
            catch (Exception ex)
            {
                return listadoCuerpo;
            }

        }
    }
}
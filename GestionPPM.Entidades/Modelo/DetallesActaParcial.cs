using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace GestionPPM.Entidades.Modelo
{
    public class DetallesActaParcial
    {
        public DetallesActaParcial()
        {
            Acuerdos = new List<DetalleActaAcuerdos>();
            Entregables = new List<DetalleActaEntregables>();
            Participantes = new List<DetalleActaParticipantes>();
            Responsables = new List<DetalleActaResponsables>();
            Temas = new List<DetalleActaTemasTratar>();
            CondicionesGenerales = new List<DetalleActaCondicionesGenerales>();

            DetalleCliente = new List<DetalleActaCliente>();
            DetalleContabilidad = new List<DetalleActaContabilidad>();
        }

        public DetallesActaParcial(List<DetalleActaAcuerdos> _acuerdos, List<DetalleActaEntregables> _entregables, List<DetalleActaParticipantes> _participantes,
            List<DetalleActaResponsables> _responsables, List<DetalleActaTemasTratar> _temas, List<DetalleActaCondicionesGenerales> _condiciones)
        {
            Acuerdos = _acuerdos;
            Entregables = _entregables;
            Participantes = _participantes;
            Responsables = _responsables;
            Temas = _temas;
            CondicionesGenerales = _condiciones;
        }

        public long ActaID { get; set; }
        public List<DetalleActaAcuerdos> Acuerdos { get; set; }
        public List<DetalleActaEntregables> Entregables { get; set; }
        public List<DetalleActaParticipantes> Participantes { get; set; }
        public List<DetalleActaResponsables> Responsables { get; set; }
        public List<DetalleActaTemasTratar> Temas { get; set; }
        public List<DetalleActaCondicionesGenerales> CondicionesGenerales { get; set; }
        public List<DetalleActaCliente> DetalleCliente { get; set; }
        public List<DetalleActaContabilidad> DetalleContabilidad { get; set; }
    }
}
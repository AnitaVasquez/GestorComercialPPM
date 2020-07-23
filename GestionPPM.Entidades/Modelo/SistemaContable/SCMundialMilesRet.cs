using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace GestionPPM.Entidades.Modelo.SistemaContable
{
    public class SCMundialMilesRet
    {
        public List<SCCabeceraRet> Cabeceras { get; set; }
        public List<SCDetalleRet> Detalles { get; set; }
        public SCMundialMilesRet()
        {
            Cabeceras = new List<SCCabeceraRet>();
            Detalles = new List<SCDetalleRet>();
        }
    }
}
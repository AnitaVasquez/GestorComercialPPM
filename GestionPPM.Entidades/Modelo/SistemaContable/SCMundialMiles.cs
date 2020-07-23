using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace GestionPPM.Entidades.Modelo.SistemaContable
{
    public class SCMundialMiles
    {
        public List<SCCabecera> Cabeceras { get; set; }
        public List<SCDetalle> Detalles { get; set; }
        public SCMundialMiles()
        {
            Cabeceras = new List<SCCabecera>();
            Detalles = new List<SCDetalle>();
        }

    }
}
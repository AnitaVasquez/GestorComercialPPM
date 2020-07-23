using GestionPPM.Entidades.Modelo.SistemaContable;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace GestionPPM.Entidades.Modelo
{
    public class Documento
    {
        public SCCabecera Cabeceras { get; set; }
        public List<SCDetalle> Detalles { get; set; }
        public Documento()
        {
            Detalles = new List<SCDetalle>();
        }
    }
}
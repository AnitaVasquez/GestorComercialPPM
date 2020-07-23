using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace GestionPPM.Entidades.Modelo.SistemaContable
{
    public class SCCabeceraRet
    {
        public string SecuencialFactura { get; set; }
        public string Factura { get; set; }
        public string Motivo { get; set; }
        public string TotalFactura { get; set; }
        public string Segmento { get; set; }
        public string CantDetalle { get; set; }
    }
}
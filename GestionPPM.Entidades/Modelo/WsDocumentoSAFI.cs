using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace GestionPPM.Entidades.Modelo
{
    public class WsDocumentoSAFI
    {
        public long id { get; set; }
        public string tipo { get; set; }
        public string factura { get; set; }
        public int? Estado { get; set; }
        public string AutorizacionSRI { get; set; }
        public string ClaveAcceso { get; set; }
        public DateTime? FechaAutorizacion { get; set; }
        public int? idFactura { get; set; }
        public string DescripcionError { get; set; }
        public string CodigoError { get; set; }
        public string CodigoFormaPago { get; set; }
        public string Plazo { get; set; }
        public string secuencialerp { get; set; }
        public string numdocumento { get; set; }
        public string email_cliente { get; set; }
    }
}
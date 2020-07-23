using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace GestionPPM.Repositorios
{
    public class RespuestaTransaccion
    {
        public bool Estado { get; set; }
        public string Respuesta { get; set; }

        // Campo que solo se utiliza para el código de cotización
        public int? CotizacionID { get; set; }

        // Parametro para cotizador
        public int? CotizadorID { get; set; }

        // Parametro para idSolicitud
        public int? SolicitudID { get; set; }

        // Parametro generico EntidadID
        public long? EntidadID { get; set; }

        // Parámetro para número de Prefactura
        public int? DocumentoSAFIID { get; set; }
    }
}
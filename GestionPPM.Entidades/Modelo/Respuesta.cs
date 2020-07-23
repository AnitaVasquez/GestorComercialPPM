using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace GestionPPM.Entidades.Modelo
{
    public class Respuesta
    {
        public string mensaje { get; set; }
        public string codigoRetorno { get; set; }
        public string estado { get; set; }
        public string numeroDocumento { get; set; }
    }

    public class Respuesta2
    {
        public string mensaje { get; set; }
        public string codigoRetorno { get; set; }
        public string estado { get; set; }
        public string numeroDocumento { get; set; }
        public string secuencial { get; set; }
    } 
}
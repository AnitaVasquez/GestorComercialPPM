using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace GestionPPM.Entidades.Modelo
{
    public class CargaMasiva
    {
        public Int64 Fila { get; set; }
        public Int64 Columna { get; set; }
        public string Valor { get; set; }
        public string Error { get; set; }
    }
}
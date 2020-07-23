using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace GestionPPM.Entidades.Modelo
{
    public class OrganigramaParcial
    {
        public OrganigramaParcial()
        {
            parent = string.Empty;
        }

        public int id { get; set; }
        public string usuario { get; set; }
        public string cargo { get; set; }
        public string mail { get; set; }
        public string departamento { get; set; }
        public string codigo { get; set; }
        public string parent { get; set; }
    }
}
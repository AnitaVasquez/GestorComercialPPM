using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace GestionPPM.Entidades.Modelo
{
    public class MultiSelectJQueryUi
    {
        public MultiSelectJQueryUi(long ID, string texto, string descripcion) {
            id = ID;
            text = texto;
            desc = descripcion;
        }
        public long id { get; set; }
        public string text { get; set; }
        public string desc { get; set; }
    }
}
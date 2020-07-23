using GestionPPM.Entidades.Modelo.PlaceToPay;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace GestionPPM.Entidades.Modelo
{
    public class AutoCompleteUI
    {
        public AutoCompleteUI(long ID, string texto, string descripcion)
        {
            id = ID;
            text = texto;
            desc = descripcion;
        }
        
        // Constructor especial para enviar parámetros adicionales de código de cotización
        public AutoCompleteUI(long ID, string texto, string descripcion, Dictionary<string, CodigoCotizacionInfo> listadoAuxiliar)
        {
            id = ID;
            text = texto;
            desc = descripcion;
            auxiliares = listadoAuxiliar; 
        }

        // Constructor especial para enviar parámetros adicionales de código de cotización
        public AutoCompleteUI(long ID, string texto, string descripcion, Dictionary<string, ComercioPlaceToPayInfo> listadoAuxiliar)
        {
            id = ID;
            text = texto;
            desc = descripcion;
            auxiliares2 = listadoAuxiliar;
        }

        public Dictionary<string, CodigoCotizacionInfo> auxiliares { get; set; }

        public Dictionary<string, ComercioPlaceToPayInfo> auxiliares2 { get; set; }

        public long id { get; set; }
        public string text { get; set; }
        public string desc { get; set; }
        public string icon { get; set; }
    }
}
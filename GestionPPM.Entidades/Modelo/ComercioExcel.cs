using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace GestionPPM.Entidades.Modelo
{
    public class ComercioExcel
    {
        public int id { get; set; }
        public string id_comercio  { get; set; }
        public string comercio { get; set; }
        public string ruc_cliente { get; set; }
        public string fecha_salida_produccion { get; set; }
        public string fecha_inactivo { get; set; }
        public string estatus_contrato { get; set; }
        public string tipo_subsidio { get; set; }
        public string agrupacion { get; set; }
        public string compartido_ptop { get; set; }
        public decimal descuento { get; set; }
        public decimal porcentaje_iva { get; set; }
        public string cobro_porcentaje { get; set; }
    }
}
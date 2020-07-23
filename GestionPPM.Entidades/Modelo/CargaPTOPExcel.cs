using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace GestionPPM.Entidades.Modelo
{
    public class CargaPTOPExcel
    {
        public int id { get; set; }
        public string id_comercio  { get; set; }
        public string plan { get; set; }
        public string facturable_certificacion { get; set; }
        public string facturable_mensual { get; set; }
        public int mes { get; set; }
        public int anio { get; set; }
        public string detalle { get; set; }
        public decimal valor_certificacion { get; set; }
        public int transacciones_aprobadas { get; set; }
        public int transacciones_rechazadas { get; set; }
        public decimal monto_aprobado { get; set; }
        public decimal monto_rechazado { get; set; }
        public string observaciones { get; set; }
        public string email { get; set; }
    }
}
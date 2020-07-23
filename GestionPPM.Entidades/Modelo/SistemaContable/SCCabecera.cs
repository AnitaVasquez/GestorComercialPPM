using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace GestionPPM.Entidades.Modelo.SistemaContable
{
    public class SCCabecera
    {
        public int ID { get; set; } // Para controlar facturas con múltiples detalles
        public string NombreCliente { get; set; }
        public string RucCliente { get; set; }
        public string DirCliente { get; set; }
        public string MailCliente { get; set; }
        public string TelefonoCliente { get; set; }
        public string SecuencialFactura { get; set; }
        public string ObservFactura { get; set; }
        public string SubTotalFactura { get; set; }
        public string TotalFactura { get; set; }
        public string Segmento { get; set; }
        public string CantDetalle { get; set; }
        public int CantidadDetalle { get; set; } // Factura con múltiples detalles
        public string Descuento { get; set; }
        public string FechaVenta { get; set; }

        // Campos para cotizaciones
        public string Vencimiento { get; set; }
        public string Comentario1 { get; set; }
        public string PlazoDias { get; set; }
        public string FhComent { get; set; }
        public string FhComent1 { get; set; }
        public string FhComent2 { get; set; }
        public string Bodega { get; set; }
    }
}
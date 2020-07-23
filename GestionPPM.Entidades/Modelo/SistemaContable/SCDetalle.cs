using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace GestionPPM.Entidades.Modelo.SistemaContable
{
    public class SCDetalle
    {
        public int ID { get; set; } // Para factura con multiples detalles
        public string Cantidad { get; set; }
        public string Detalle { get; set; }
        public string Valor { get; set; }
        public string Descuento { get; set; }
        public string SubTotal { get; set; }
        public string Total { get; set; }
        public string CodigoCategoria { get; set; }
        public string CodigoProducto { get; set; }//validar
        public string RucProveedor { get; set; }
        public string Proveedor { get; set; }
        public string DetalleFactura { get; set; }
        public string CostoUnitario { get; set; }
        public string FechaVenta { get; set; }
        public string SecuencialFactura { get; set; }
    }
}
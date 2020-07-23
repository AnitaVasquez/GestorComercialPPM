using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace GestionPPM.Entidades.Modelo.SistemaContable
{
    public partial class WraperNotaCreditoWs
    {
        public CabeceraWs NotaCredito { get; set; }
    }
    public partial class CabeceraWs
    {
        public ClienteSAFI Cliente { get; set; }
        public DetallesCotizacion CotizacionDetalle { get; set; }
        public NotaCreditoWs Detalle { get; set; }
    } 
    public partial class DetalleNotaCreditoWs
    { 
        public int Cantidad { get; set; }         
        public string Detalle { get; set; }         
        public decimal Valor { get; set; }     
        public decimal SubTotal { get; set; }
        public string CodigoCategoria { get; set; }
        public string CodigoProducto { get; set; }
        public string RUCProveedor { get; set; }
        public string Proveedor { get; set; }
        public decimal CostoUnitario { get; set; }
        public DateTime FechaVenta { get; set; }
    }
    public partial class NotaCreditoWs
    {
        public string Idfactura { get; set; }
        public string Motivo { get; set; }
        public decimal Valor { get; set; }
        public string Secuencial { get; set; }
        public int Estado { get; set; }
        public List<DetalleNotaCreditoWs> DetalleNotaCredito { get; set; } 
        public NotaCreditoWs()
        {
            DetalleNotaCredito = new List<DetalleNotaCreditoWs>();
        }
    }

}
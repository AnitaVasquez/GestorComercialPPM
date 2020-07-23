using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace GestionPPM.Entidades.Modelo.SistemaContable
{
    public partial class Wrapper
    {
        public Cabecera Cabecera { get; set; }
        public Detalle Detalle { get; set; }
    }
    public partial class Cabecera
    {
        public ClienteSAFI Cliente { get; set; }
        public DetallesCotizacion CotizacionDetalle { get; set; }
    }
    public partial class Detalle
    {
        public Factura Factura { get; set; }
    }
    public partial class ClienteSAFI
    {
        public string CodigoCliente { get; set; }
        public string NombreCliente { get; set; }
        public string Identificacion { get; set; }
        public string Direccion { get; set; }
        public string Mail { get; set; }
        public string Telefono { get; set; }
        public string Segmento { get; set; }
    }

    public partial class DetallesCotizacion
    {
        // Campos para cotizaciones
        public string Vencimiento { get; set; }
        public string Comentario1 { get; set; }
        public string PlazoDias { get; set; }
        public string FhComent { get; set; }
        public string FhComent1 { get; set; }
        public string FhComent2 { get; set; }
        public string Bodega { get; set; }
        public string UGE { get; set; }
        public string FormaPago { get; set; }
    }

    public partial class Factura
    {
        public string Secuencial { get; set; }
        public string Observacion { get; set; }
        public decimal PorcentajeIva { get; set; }
        public decimal SubTotal { get; set; }
        public decimal Total { get; set; }
        public decimal Descuento { get; set; }
        public List<DetalleFactura> DetalleFactura { get; set; }
        public int Estado { get; set; }
        public Factura()
        {
            DetalleFactura = new List<DetalleFactura>();
        }
    }
    public partial class DetalleFactura
    {
        public int Cantidad { get; set; }
        public string Detalle { get; set; }
        public decimal PorcentajeIva { get; set; }
        public decimal Valor { get; set; }
        public decimal SubTotal { get; set; }
        public decimal Descuento { get; set; }
        public decimal Total { get; set; }
        public string CodigoCategoria { get; set; }
        public string CodigoProducto { get; set; }
        public string NombreProducto { get; set; }
        public string RUCProveedor { get; set; }
        public string Proveedor { get; set; }
        public decimal CostoUnitario { get; set; }
        public DateTime FechaVenta { get; set; }
    }
    public partial class FacturaCompleta
    {
        public int nfactura { get; set; }
        public int ncod_cli { get; set; }
        public int dfecha { get; set; }
        public int nporc_iva { get; set; }
        public int ncod_ciu { get; set; }
        public int mobserva { get; set; }
        public int ytotal { get; set; }
        public int ysaldo { get; set; }
        public int dfec_ven { get; set; }
        public int dfec_rec { get; set; }
        public int nfacturacion { get; set; }
        public int ntip_ref { get; set; }
        public int ncod_rub { get; set; }
        public int ncod_pro { get; set; }
        public int nporc_fee { get; set; }
        public int nnum_cotiza { get; set; }
        public int lprincipal { get; set; }
        public int ccon_mar { get; set; }
        public int cnom_ent { get; set; }
        public int ccon_fac { get; set; }
        public int val_canje { get; set; }
        public int val_efectivo { get; set; }
        public int nsubpro { get; set; }
        public int notrospos { get; set; }
        public int cotrospos { get; set; }
        public int notrosneg { get; set; }
        public int cotrosneg { get; set; }
        public int ncod_ven { get; set; }
        public int nporc_comp { get; set; }
        public int ntipo_neg { get; set; }
        public int dfec_emi { get; set; }
        public int dfec_ent { get; set; }
        public int dfec_comi { get; set; }
        public int ntipo_doc { get; set; }
        public int id_catalogo { get; set; }
        public int estado { get; set; }
        public int autorizacionSRI { get; set; }
        public int claveAcceso { get; set; }
        public int fechaAutorizacion { get; set; }
        public int idFactura { get; set; }
        public int descripcionError { get; set; }
        public int codigoError { get; set; }
        public int tcliente { get; set; }
        public int tfactura_ref { get; set; }
        public int tproducto { get; set; }
    }
}
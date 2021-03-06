//------------------------------------------------------------------------------
// <auto-generated>
//     Este código se generó a partir de una plantilla.
//
//     Los cambios manuales en este archivo pueden causar un comportamiento inesperado de la aplicación.
//     Los cambios manuales en este archivo se sobrescribirán si se regenera el código.
// </auto-generated>
//------------------------------------------------------------------------------

namespace GestionPPM.Entidades.Modelo
{
    using System;
    
    public partial class ActaInfo
    {
        public long IDActa { get; set; }
        public string CodigoActa { get; set; }
        public System.DateTime FechaCreacion { get; set; }
        public int ElaboradoPor { get; set; }
        public string NombresElaboradoPor { get; set; }
        public string Cargo { get; set; }
        public System.DateTime FechaInicio { get; set; }
        public System.DateTime FechaFin { get; set; }
        public Nullable<System.DateTime> FechaEntrega { get; set; }
        public string Lugar { get; set; }
        public Nullable<int> NumeroReunion { get; set; }
        public string HoraInicio { get; set; }
        public string HoraFin { get; set; }
        public string ReferenciaCliente { get; set; }
        public string AlcanceObjetivo { get; set; }
        public string FacilitadorModerador { get; set; }
        public Nullable<int> CodigoCotizacionID { get; set; }
        public string CodigoCotizacion { get; set; }
        public string Cliente { get; set; }
        public string NombreProyecto { get; set; }
        public string DescripcionProyecto { get; set; }
        public string Observaciones { get; set; }
        public bool Suspendida { get; set; }
        public int TipoActaID { get; set; }
        public string NombreTipoActa { get; set; }
        public string CodigoTipoActa { get; set; }
        public Nullable<long> SecuencialTipoActa { get; set; }
        public Nullable<int> AnioTipoActa { get; set; }
        public bool Estado { get; set; }
        public Nullable<int> DetalleActaParticipantesID { get; set; }
        public string NombresParticipante { get; set; }
        public Nullable<bool> Presente { get; set; }
        public Nullable<int> DetalleActaResponsablesClienteID { get; set; }
        public string NombresResponsable { get; set; }
        public string Empresa { get; set; }
        public string Rol { get; set; }
        public Nullable<int> DetalleActaEntregablesID { get; set; }
        public string Entregable { get; set; }
        public string Tipo { get; set; }
        public Nullable<int> DetalleActaTemasTratarID { get; set; }
        public string Tema { get; set; }
        public string ResponsableTema { get; set; }
        public Nullable<int> DetalleActaAcuerdosID { get; set; }
        public string Acuerdo { get; set; }
        public string ResponsableAcuerdo { get; set; }
        public Nullable<System.DateTime> Fecha { get; set; }
        public Nullable<int> DetalleActaCondicionesGeneralesID { get; set; }
        public string Condicion { get; set; }
        public Nullable<long> IDActaInformacionAdicional { get; set; }
        public string AcuerdoConformidad { get; set; }
        public string Firmas { get; set; }
        public string TipoProyecto { get; set; }
        public Nullable<int> DetalleActaClienteID { get; set; }
        public Nullable<int> IDDetalleActaCliente { get; set; }
        public Nullable<int> id_facturacion_safi_ActaCliente { get; set; }
        public Nullable<int> id_codigo_cotizacion_ActaCliente { get; set; }
        public string codigo_cotizacion_ActaCliente { get; set; }
        public string TipoIntermediario_ActaCliente { get; set; }
        public Nullable<int> ClienteID_ActaCliente { get; set; }
        public string nombre_comercial_cliente_ActaCliente { get; set; }
        public string direccion_ActaCliente { get; set; }
        public string Ciudad_ActaCliente { get; set; }
        public string Ejecutivo_ActaCliente { get; set; }
        public string TelefonoEjecutivo_ActaCliente { get; set; }
        public string nombre_proyecto_ActaCliente { get; set; }
        public Nullable<int> UltimaIDCotizacion_ActaCliente { get; set; }
        public Nullable<decimal> UltimaVersionCotizacion_ActaCliente { get; set; }
        public Nullable<decimal> TotalCotizacion_ActaCliente { get; set; }
        public string ObservacionCotizacion_ActaCliente { get; set; }
        public string detalle_cotizacion_ActaCliente { get; set; }
        public Nullable<int> numero_pago_ActaCliente { get; set; }
        public Nullable<int> id_codigo_producto_ActaCliente { get; set; }
        public string nombre_producto_ActaCliente { get; set; }
        public string codigo_producto_ActaCliente { get; set; }
        public Nullable<int> cantidad_ActaCliente { get; set; }
        public Nullable<decimal> precio_unitario_ActaCliente { get; set; }
        public Nullable<decimal> subtotal_pago_ActaCliente { get; set; }
        public Nullable<decimal> iva_pago_ActaCliente { get; set; }
        public Nullable<decimal> descuento_pago_ActaCliente { get; set; }
        public Nullable<decimal> total_pago_ActaCliente { get; set; }
        public Nullable<System.DateTime> fecha_prefactura_ActaCliente { get; set; }
        public string numero_prefactura_ActaCliente { get; set; }
        public Nullable<System.DateTime> fecha_factura_ActaCliente { get; set; }
        public string numero_factura_ActaCliente { get; set; }
        public Nullable<System.DateTime> fecha_nota_credito_ActaCliente { get; set; }
        public string numero_nota_credito_ActaCliente { get; set; }
        public string numero_retencion_ActaCliente { get; set; }
        public Nullable<bool> estado_ActaCliente { get; set; }
        public Nullable<int> DetalleActaContabilidadID { get; set; }
        public Nullable<int> IDDetalleActaContabilidad { get; set; }
        public Nullable<int> id_facturacion_safi_ActaContabilidad { get; set; }
        public Nullable<int> id_codigo_cotizacion_ActaContabilidad { get; set; }
        public string codigo_cotizacion_ActaContabilidad { get; set; }
        public string TipoIntermediario_ActaContabilidad { get; set; }
        public string nombre_comercial_cliente_ActaContabilidad { get; set; }
        public string direccion_ActaContabilidad { get; set; }
        public string Ciudad_ActaContabilidad { get; set; }
        public string Ejecutivo_ActaContabilidad { get; set; }
        public string TelefonoEjecutivo_ActaContabilidad { get; set; }
        public string nombre_proyecto_ActaContabilidad { get; set; }
        public Nullable<int> UltimaIDCotizacion_ActaContabilidad { get; set; }
        public Nullable<decimal> UltimaVersionCotizacion_ActaContabilidad { get; set; }
        public Nullable<decimal> TotalCotizacion_ActaContabilidad { get; set; }
        public string ObservacionCotizacion_ActaContabilidad { get; set; }
        public string detalle_cotizacion_ActaContabilidad { get; set; }
        public Nullable<int> numero_pago_ActaContabilidad { get; set; }
        public Nullable<int> id_codigo_producto_ActaContabilidad { get; set; }
        public string nombre_producto_ActaContabilidad { get; set; }
        public string codigo_producto_ActaContabilidad { get; set; }
        public Nullable<int> cantidad_ActaContabilidad { get; set; }
        public Nullable<decimal> precio_unitario_ActaContabilidad { get; set; }
        public Nullable<decimal> subtotal_pago_ActaContabilidad { get; set; }
        public Nullable<decimal> iva_pago_ActaContabilidad { get; set; }
        public Nullable<decimal> descuento_pago_ActaContabilidad { get; set; }
        public Nullable<decimal> total_pago_ActaContabilidad { get; set; }
        public Nullable<System.DateTime> fecha_prefactura_ActaContabilidad { get; set; }
        public string numero_prefactura_ActaContabilidad { get; set; }
        public Nullable<System.DateTime> fecha_factura_ActaContabilidad { get; set; }
        public string numero_factura_ActaContabilidad { get; set; }
        public Nullable<System.DateTime> fecha_nota_credito_ActaContabilidad { get; set; }
        public string numero_nota_credito_ActaContabilidad { get; set; }
        public string numero_retencion_ActaContabilidad { get; set; }
        public Nullable<bool> estado_ActaContabilidad { get; set; }
    }
}

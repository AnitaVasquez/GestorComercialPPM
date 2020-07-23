using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace GestionPPM.Entidades.Modelo
{
    public partial class CodigoCotizacionExcel
    {
        public int N { get; set; } // N.-   ---> Referencia a columna en Excel para asociar contactos y sublineas de negocio
        public int id_codigo_cotizacion { get; set; }
        public string codigo_cotizacion { get; set; }
        public Nullable<System.DateTime> fecha_cotizacion { get; set; }
        public Nullable<int> responsable { get; set; }
        public Nullable<int> estatus_codigo { get; set; }
        public Nullable<int> id_cliente { get; set; }
        public Nullable<int> ejecutivo { get; set; }
        public Nullable<int> tipo_requerido { get; set; }
        public Nullable<int> tipo_intermediario { get; set; }
        public Nullable<int> tipo_proyecto { get; set; }
        public Nullable<int> dimension_proyecto { get; set; }
        public Nullable<bool> aplica_contrato { get; set; }
        public Nullable<bool> forma_pago { get; set; }
        public Nullable<decimal> forma_pago_1 { get; set; }
        public Nullable<decimal> forma_pago_2 { get; set; }
        public Nullable<decimal> forma_pago_3 { get; set; }
        public Nullable<decimal> forma_pago_4 { get; set; }
        public Nullable<int> etapa_cliente { get; set; }
        public Nullable<int> etapa_general { get; set; }
        public Nullable<int> estatus_detallado { get; set; }
        public Nullable<int> estatus_general { get; set; }
        public Nullable<int> tipo_producto_PtoP { get; set; }
        public Nullable<int> tipo_plan { get; set; }
        public Nullable<int> tipo_tarifa { get; set; }
        public Nullable<int> tipo_migracion { get; set; }
        public Nullable<int> tipo_etapa_PtoP { get; set; }
        public Nullable<int> tipo_subsidio { get; set; }
        public string nombre_proyecto { get; set; }
        public string descripcion_proyecto { get; set; }
        public Nullable<int> tipo_fee { get; set; }
        public Nullable<int> creacion_safi { get; set; }
        public Nullable<bool> facturable { get; set; }
        public string area_departamento_usuario { get; set; }
        public Nullable<int> pais { get; set; }
        public Nullable<int> ciudad { get; set; }
        public string direccion { get; set; }
        public Nullable<int> id_empresa { get; set; }
        public bool estado_cliente { get; set; }
        public Nullable<int> tipo_cliente { get; set; }
        public Nullable<int> tipo_zoho { get; set; }
        public Nullable<int> referido { get; set; }
        public Nullable<bool> Estado { get; set; }
        public List<ContactosCodigoCotizacion> Contactos { get; set; }
        public List<SublineaNegocioCodigoCotizacion> SublineasNegocio { get; set; }
    }
}
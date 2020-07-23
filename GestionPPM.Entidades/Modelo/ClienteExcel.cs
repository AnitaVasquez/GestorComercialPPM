using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace GestionPPM.Entidades.Modelo
{
    public class ClienteExcel
    {
        public int N { get; set; }
        public int id_cliente { get; set; }
        public int? usuario_id { get; set; }
        public int? referido { get; set; }
        public int? intermediario { get; set; }
        public int? tipo_zoho { get; set; }
        public int? tipo_cliente { get; set; }
        public int? tamanio_empresa { get; set; }
        public int? etapa_cliente { get; set; }
        public int? potencial_crecimiento { get; set; }
        public int? categorizacion_cliente { get; set; }
        public string ruc_ci_cliente { get; set; }
        public string razon_social_cliente { get; set; }
        public string nombre_comercial_cliente { get; set; }
        public decimal? ingresos_anuales_cliente { get; set; }
        public int? sector { get; set; }
        public int? pais { get; set; }
        public int? ciudad { get; set; }
        public string direccion_cliente { get; set; }
        public bool? estado_cliente { get; set; }
        public List<ContactosClientes> contactos { get; set; }
        public List<ClienteLineaNegocio> lineasNegocio { get; set; }
    }
}
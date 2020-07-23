using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace GestionPPM.Entidades.Modelo
{
    public partial class UsuarioCE
    {
        public int id_usuario { get; set; }
        public string nombre_usuario { get; set; }
        public string apellido_usuario { get; set; }
        public Nullable<int> cliente_asociado { get; set; }
        public Nullable<int> tipo_usuario { get; set; }
        public string area_departamento { get; set; }
        public string area_departamento_texto { get; set; }
        public Nullable<int> pais { get; set; }
        public Nullable<int> ciudad { get; set; }
        public string direccion_usuario { get; set; }
        public string mail_usuario { get; set; }
        public string telefono_usuario { get; set; }
        public string celular_usuario { get; set; }
        public string codigo_usuario { get; set; }
        public string clave_usuario { get; set; }
        public Nullable<bool> estado_usuario { get; set; }
        public Nullable<bool> activo_usuario { get; set; }
        public string cargo_usuario { get; set; }
        public string cargo_usuario_texto { get; set; }
        public int id_rol { get; set; }
        public int id_empresa { get; set; }
        public int secu_usua { get; set; }
        public bool reset_clave { get; set; }
        public string link_firma { get; set; }
        public bool validacion_correo { get; set; }
        public string NombreEmpresa { get; set; }
        public bool usar_organigrama { get; set; }
    }
}
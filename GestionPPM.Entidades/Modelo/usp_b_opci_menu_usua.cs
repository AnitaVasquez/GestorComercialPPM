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
    
    public partial class usp_b_opci_menu_usua
    {
        public string nombre_menu { get; set; }
        public Nullable<int> id_menu_padre { get; set; }
        public string carpeta { get; set; }
        public string opcion { get; set; }
        public Nullable<int> hijos { get; set; }
        public string permisos { get; set; }
        public int id_menu { get; set; }
        public Nullable<int> orden_menu { get; set; }
    }
}

using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace GestionPPM.Entidades.Modelo
{
    public partial class ProductosGestorCE
    {
        public int id_tipo_producto { get; set; }

        public int id_linea_negocio { get; set; }

        public int id_sublinea_negocio { get; set; }

        public int id_producto_general { get; set; }

        public int id_producto_gestor { get; set; } 

        public string nombre { get; set; }

        public string descripcon { get; set; }

        public bool estado { get; set; }

    }
}
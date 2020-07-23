using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace GestionPPM.Entidades.Modelo
{
    public class FirmasActaParcial
    {
        //[{"nombre":"Nombrejkasjkajksas","cargo":"Cargo","empresa":"Empresa"},{"usuarioNombre":"Nombre","usuarioCargo":"Cargo","usuarioEmpresa":"Empresa"}]
        public string nombre { get; set; }
        public string cargo { get; set; }
        public string empresa { get; set; }
        public string fecha { get; set; }
        public string usuarioNombre { get; set; }
        public string usuarioCargo { get; set; }
        public string usuarioEmpresa { get; set; }
        public string Usuariofecha { get; set; }
    }
}
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Web;

namespace GestionPPM.Entidades.Modelo
{
    public class CambiarClave
    {
        public long UsuaCodi { get; set; }


        //[Required(ErrorMessage = "Ingrese la Contraseña Actual")]
        //[DataType(DataType.Password)]
        public string ContraseniaActual { get; set; }


        //[Required(ErrorMessage = "Ingrese la Nueva Contraseña")]
        //[DataType(DataType.Password)]
        public string ContraseniaNueva { get; set; }


        //[Required(ErrorMessage = "Confirme la Nueva Contraseña")]
        //[DataType(DataType.Password)]
        public string ConfirmarContrasenia { get; set; }

    }
}
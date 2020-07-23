using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace GestionPPM.Repositorios
{
    public class SessionHelperR
    {
        //Obtener la respuesta
        public static string GetRespuesta()
        {
            if (!(HttpContext.Current.Session["Resultado"] is string resultado))
                return "";
            else
                return resultado;
        }

        //Obtener el estado
        public static string GetEstado()
        {
            if (!(HttpContext.Current.Session["Estado"] is string estado))
                return "";
            else
                return estado;
        }
    }
}
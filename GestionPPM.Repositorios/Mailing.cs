using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Threading.Tasks;
using System.Web;
using System.Web.Hosting;

namespace GestionPPM.Repositorios
{
    public static class Mailing
    {
        // Contruye el HTML a enviar
        public static string EmailFormatoHTMLFinal(string template, string cuerpoCorreo)
        {
            try
            {
                var message = GetEmailTemplate(template);
                message = message.Replace("@ViewBag.CuerpoCorreo", cuerpoCorreo); //Cuerpo del correo
                return message;
            }
            catch (Exception ex)
            {
                return string.Empty;
            }
        }

        // Obtiene la plantilla del Email
        public static string GetEmailTemplate(string nombreTemplate)
        {
            try
            {
                string body = System.IO.File.ReadAllText(HostingEnvironment.MapPath("~/Content/templates/") + nombreTemplate + ".cshtml");
                return body.ToString();
            }
            catch (Exception ex)
            {
                string plantillaPorDefecto = "Default";
                string body = System.IO.File.ReadAllText(HostingEnvironment.MapPath("~/Content/templates/") + plantillaPorDefecto + ".cshtml");
                return body.ToString();
            }
        }
    }
}
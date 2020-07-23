using System;
using System.Threading.Tasks;
using GestionPPM.Entidades.Metodos; 
using Owin;
 

namespace TemplateInicial
{
    public class Startup
    {
        public void Configuration(IAppBuilder app)
        {
            // Para obtener más información sobre cómo configurar la aplicación, visite https://go.microsoft.com/fwlink/?LinkID=316888

            //GlobalConfiguration.Configuration.UseSqlServerStorage("DefaultConnection");
            //app.UseHangfireDashboard("/visor", new DashboardOptions
            //{
            //    AppPath = null,
            //});

            //Creacion siempre y cuando no existan de perfiles de carga masiva
            PerfilesEntity.CreacionPerfilesCargaMasiva();

        }
    }
}

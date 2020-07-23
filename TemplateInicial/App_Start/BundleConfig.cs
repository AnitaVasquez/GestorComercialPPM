using System.Web.Optimization; 

namespace TemplateInicial
{
    public class BundleConfig
    {
        public static void RegisterBundles(BundleCollection bundles)
        {

            bundles.Add(new ScriptBundle("~/bundles/jquery").Include(
                        "~/Scripts/jquery-3.3.1.min.js"));

            bundles.Add(new ScriptBundle("~/bundles/jqueryval").Include(
                        "~/Scripts/jquery.validate*"));

            // Utilice la versión de desarrollo de Modernizr para desarrollar y obtener información. De este modo, estará
            // para la producción, use la herramienta de compilación disponible en https://modernizr.com para seleccionar solo las pruebas que necesite.
            bundles.Add(new ScriptBundle("~/bundles/modernizr").Include(
                        "~/Scripts/modernizr-*"));


            bundles.Add(new ScriptBundle("~/bundles/bootstrap").Include(
                        "~/Scripts/bootstrap.js", 
                      "~/Scripts/respond.js"));

            bundles.Add(new StyleBundle("~/Content/css").Include(
                      "~/Content/bootstrap.css",
                      "~/Content/site.css" 
                      ));

            bundles.Add(new StyleBundle("~/Bundles/css")
                .Include("~/Content/bootstrap.min.css")
                .Include("~/Content/css/bootstrap-select.css", "~/Content/bootstrap-sortable.css")
                .Include("~/Content/css/bootstrap-datepicker3.min.css")
                .Include("~/Content/css/font-awesome.min.css")
                //.Include("~/Content/css/icheck/blue.min.css")
                .Include("~/Content/css/bootstrap-wysihtml5/bootstrap3-wysihtml5.min.css")
                .Include("~/Content/css/AdminLTE.css")
                .Include("~/Content/css/skins/skin-blue.css"));

            bundles.Add(new ScriptBundle("~/Bundles/js")
                //.Include("~/Content/js/plugins/jquery/jquery-3.3.1.js")
                .Include("~/Content/js/plugins/bootstrap/bootstrap.js")
                .Include("~/Content/js/plugins/fastclick/fastclick.js")
                .Include("~/Content/js/plugins/slimscroll/jquery.slimscroll.js")
                .Include("~/Content/js/plugins/bootstrap-select/bootstrap-select.js")
                .Include("~/Content/js/plugins/moment/moment.js")
                .Include("~/Content/js/plugins/datepicker/bootstrap-datepicker.js")
                //.Include("~/Content/js/plugins/icheck/icheck.js")
                .Include("~/Content/js/plugins/validator/validator.js")
                .Include("~/Content/js/plugins/inputmask/jquery.inputmask.bundle.js")
                .Include("~/Content/js/plugins/bootstrap-wysihtml5/bootstrap3-wysihtml5.all.min.js")
                .Include("~/Content/js/adminlte.js") 

                .Include("~/Scripts/AwesomeMvc.js")
                .Include("~/Scripts/awem.js")
                .Include("~/Scripts/utils.js")
                //.Include("~/Scripts/jquery-3.3.1")

                .Include("~/Content/js/init.js"));

//#if DEBUG
//            BundleTable.EnableOptimizations = false;
//#else
//            BundleTable.EnableOptimizations = true;
//#endif
        }
    }
}

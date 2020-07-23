using System.Web.Mvc;

namespace TemplateInicial.Controllers
{
    public class ErrorController : Controller
    {
        [HttpGet]
        public ActionResult InternalServerError()
        {
            return View();
        }

        [HttpGet]
        public ActionResult NotFound()
        {
            return View();
        }

        [HttpGet]
        public ActionResult NotForbbiden()
        {
            return View();
        }
    }
}

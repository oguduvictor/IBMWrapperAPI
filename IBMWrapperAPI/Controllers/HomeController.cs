using Microsoft.AspNetCore.Mvc;

namespace IBMWrapperAPI.Controllers
{
    /// <summary>
    /// Default HomePage Controller
    /// </summary>
    public class HomeController : Controller
    {
        /// <summary>
        /// Action for routing to Swagger UI
        /// </summary>
        /// <returns></returns>
        public IActionResult Index()
        {
            return Redirect("/swagger");
        }
    }
}
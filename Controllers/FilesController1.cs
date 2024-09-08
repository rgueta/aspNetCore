using Microsoft.AspNetCore.Mvc;


namespace aspNetCore.Controllers
{
    public class FilesController1 : Controller
    {
        public IActionResult Index()
        {
            return View();
        }
    }
}

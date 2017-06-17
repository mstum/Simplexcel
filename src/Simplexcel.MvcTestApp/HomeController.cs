using System;
using System.Collections.Generic;
using System.Web.Mvc;

namespace Simplexcel.MvcTestApp.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult Index()
        {
            return View();
        }

        public ActionResult ExcelTest()
        {
            var businessData = new List<string>();
            for (int i = 1; i < 55; i++)
            {
                businessData.Add(Guid.NewGuid().ToString());
            }

            return new ExcelTestActionResult("test.xlsx", businessData);
        }
    }
}
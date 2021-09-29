using hospital.Models;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web.Mvc;

namespace hospital.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult Index()
        {
            return View();
        }

        public ActionResult Files_Download()
        {
            //Fetch all files in the Folder (Directory).
            string[] filePaths = Directory.GetFiles(Server.MapPath("~/Files/"));

            //Copy File names to Model collection.
            List<string> files = new List<string>();
            foreach (string filePath in filePaths)
            {
                files.Add(Path.GetFileName(filePath));
            }
            ViewBag.files = files;
            return View();
        }

        public ActionResult DownloadFile(string fileName)
        {
            //Build the File Path.
            string path = Server.MapPath("~/Files/") + fileName;

            if (System.IO.File.Exists(path))
            {
                //Read the File data into Byte Array.
                byte[] bytes = System.IO.File.ReadAllBytes(path);
                //Send the File to Download.
                return File(bytes, "application/octet-stream", fileName);
            }
            else
            {
                return View("Error");
            }
        }

    }
}
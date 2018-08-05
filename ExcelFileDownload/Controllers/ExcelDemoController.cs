using System.Web;
using System.Web.Mvc;
using ExcelFileDownload.Models;
using System.Data;

namespace ExcelFileDownload.Controllers
{
    public class ExcelDemoController : Controller
    {
        // GET: ExcelDemo
        public ActionResult ExcelUpload()
        {
            return View();
        }

        // GET: ExcelDemo/Details/5
        [HttpPost]
        public ActionResult UploadExcel(HttpPostedFileBase fileUpload)
        {
            if (fileUpload != null)
            {
                bool valid;
                string directorFile = "~/ DetailFormatInExcel / ";
                string[] cols = new string[1];
                cols[0] = "SerialNo";
                DataTable data_table =Extentions.ExcelFileToDataTable(fileUpload, cols, out valid, directorFile);
              
                return new JsonNetResult() { Data = data_table, JsonRequestBehavior = JsonRequestBehavior.AllowGet };
            }
            else
            {
                return RedirectToAction("CreateEdit");

            }
        }
    }
}

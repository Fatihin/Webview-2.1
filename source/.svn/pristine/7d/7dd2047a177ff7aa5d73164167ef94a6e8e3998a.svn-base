using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.Data.EntityClient;
using System.Data.EntityModel;
using Oracle.DataAccess.Client;
using System.Data.Objects;
using System.Data.Common;
using System.Data;
using WebView.Models;
using WebView.Library;
using System.IO;


namespace WebView.Controllers
{
    public class InstallMaintenanceController : Controller
    {
        //
        // GET: /InstallMaintenance/

        public ActionResult Index()
        {
            return View();
        }
        
        public ActionResult List()
        {
            WebView.WebService._base myWebService;
            myWebService = new WebService._base();

            WebService._base.OSPInstMaintenance instM = new WebService._base.OSPInstMaintenance();
            instM = myWebService.GetOSPInstMaintenance(0, 10);

            ViewData["data4"] = instM.InstMaintenanceList;

            string input = "\\\\adsvr";
            //string input = "\\\\server\\d$\\x\\y\\z\\AAA";
            string output = String.Format("http:{0}", input.Replace("\\d$\\x\\y", String.Empty).Replace("\\", "/"));

            ViewBag.output = output;

            return View();
        }

        public ActionResult NewInstallMaintenance()
        {
            List<SelectListItem> list = new List<SelectListItem>();
            list.Add(new SelectListItem() { Text = "EXE", Value = "EXE" });
            list.Add(new SelectListItem() { Text = "NEX", Value = "NEX" });
            ViewBag.Rank = list;

           return View();
        }
        
        [HttpPost]
        public ActionResult NewInstallMaintenance(InstallMaintenanceModel m)
        {
            bool success = true;
            WebView.WebService._base myWebService;
            myWebService = new WebService._base();

            bool selected = false;

            WebService._base.InstallMaintenance newInstM = new WebService._base.InstallMaintenance();

            newInstM.INST_CODE = Convert.ToString(m.InstallCode);
            newInstM.INST_NAME = m.InstallName;
            newInstM.SUP_NORM_RATE = Convert.ToString(m.SUPNormRate) ;
            newInstM.SUP_RWO_RATE = Convert.ToString(m.SUPRWORate) ;
            newInstM.IMPL_NORM_RATE = Convert.ToString(m.IMPLNormalRate) ;
            newInstM.IMPL_RWO_RATE = Convert.ToString(m.IMPLRWORate);
            newInstM.INST_TYPE = m.InstallType;
            newInstM.RANK = m.Rank;
          

            success = myWebService.AddInstMaintenance(newInstM);
            selected = true;

            if (ModelState.IsValid && selected)
            {
                if (success == true)
                    return RedirectToAction("NewSave");
                else
                    return RedirectToAction("NewSaveFail"); // store to db failed.
            }

            return View(m);
        }

        public ActionResult NewSave()
        {
            return View();
        }

        [HttpPost]
        public ActionResult GetDetails(string id)
        {
            string fileListStr = "";
            System.Diagnostics.Debug.WriteLine("############");
            System.Diagnostics.Debug.WriteLine(id);
            using (Entities ctxData = new Entities())
            {
                var query = (from q in ctxData.WV_INSTALL_MAST
                             where q.CODE_ID == id
                             select q).Single();

                return Json(new
                {
                    Success = true,
                    codeid = id,
                    code = query.INST_CODE,                    
                    name = query.INST_NAME,
                    supNorm = query.SUP_NORM_RATE,
                    supRwo = query.SUP_RWO_RATE,
                    implNorm = query.IMPL_NORM_RATE,
                    implRwo = query.IMPL_RWO_RATE,
                    type = query.INST_TYPE,
                    rank = query.RANK,                    
                    files_list = new
                    {
                        files = fileListStr
                    }
                }, JsonRequestBehavior.AllowGet); //
            }
        }

        [HttpPost]
        public ActionResult DeleteData(string targetCode)
        {
            bool success;

            WebView.WebService._base myWebService;
            myWebService = new WebService._base();

            success = myWebService.DeleteInstMaintenance(targetCode);

            return Json(new
            {
                Success = success
            }, JsonRequestBehavior.AllowGet);
        }

        public ActionResult UpdateData(string codeVal, string txtCode, string txtName, string txtsupNorm, string txtsupRwo, string txtimplNorm, string txtimplRwo, string txtType, string txtRank)
        {
            bool success = true;

            WebView.WebService._base myWebService;
            myWebService = new WebService._base();

            WebService._base.InstallMaintenance newCode = new WebService._base.InstallMaintenance();
            newCode.INST_CODE = txtCode;
            newCode.INST_NAME = txtName;
            newCode.SUP_NORM_RATE = txtsupNorm;
            newCode.SUP_RWO_RATE = txtsupRwo;
            newCode.IMPL_NORM_RATE = txtimplNorm;
            newCode.IMPL_RWO_RATE = txtimplRwo;
            newCode.INST_TYPE = txtType;
            newCode.RANK = txtRank;
            success = myWebService.UpdateInstMaintenance(newCode, codeVal);

            return Json(new
            {
                Success = success
            }, JsonRequestBehavior.AllowGet);
        }
    }
}

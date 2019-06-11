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
using PagedList;

namespace WebView.Controllers
{
    public class ANDController : Controller
    {
        //
        // GET: /AND/

        public ActionResult Index()
        {
            return View();
        }
        public ActionResult List(string searchKey, int? page)
        {
            WebView.WebService._base myWebService;
            myWebService = new WebService._base();

            WebService._base.OSPANDMaintenance ANDMain = new WebService._base.OSPANDMaintenance();

            if (searchKey != null)
            {
                if (searchKey.Equals(""))
                    ANDMain = myWebService.GetOSPANDMaintenance(0, 100, null);
                else
                {
                    ANDMain = myWebService.GetOSPANDMaintenance(0, 100, searchKey);
                    ViewBag.searchKey = searchKey;
                }
            }
            else
            {
                ANDMain = myWebService.GetOSPANDMaintenance(0, 100, null);
                ViewBag.searchKey = "";
            }

            ViewData["data7"] = ANDMain.ANDMaintenanceList;

            string input = "\\\\adsvr";
            //string input = "\\\\server\\d$\\x\\y\\z\\AAA";
            string output = String.Format("http:{0}", input.Replace("\\d$\\x\\y", String.Empty).Replace("\\", "/"));

            ViewBag.output = output;

            //return View();
            int pageSize = 10;
            int pageNumber = (page ?? 1);
            return View(ANDMain.ANDMaintenanceList.ToPagedList(pageNumber, pageSize));
        }

        public ActionResult NewAND()
        {
            return View();
        }

        [HttpPost]
        public ActionResult NewAND(ANDModel m)
        {
            bool success = true;
            WebView.WebService._base myWebService;
            myWebService = new WebService._base();

            bool selected = false;

            WebService._base.ANDMaintenance newANDMain = new WebService._base.ANDMaintenance();

            newANDMain.AND_ID = m.CanCode;
            newANDMain.AND_NAME = m.CanName;
            newANDMain.REGION_ID = m.RegionId;
            newANDMain.TARGET_ECP = m.TargetEcp;
            success = myWebService.AddANDMaintenance(newANDMain);
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

            using (Entities ctxData = new Entities())
            {
                var query = (from q in ctxData.WV_AND_MAST
                             where q.AND_ID == id
                             select q).Single();

                return Json(new
                {
                    Success = true,
                    CanCode = id,
                    CanName = query.AND_NAME,
                    RegionId = query.REGION_ID,
                    TargetEcp = query.TARGET_ECP,
                    files_list = new
                    {
                        files = fileListStr
                    }
                }, JsonRequestBehavior.AllowGet); //
            }
        }

        [HttpPost]
        public ActionResult DeleteData(string targetCanCode)
        {
            bool success;

            WebView.WebService._base myWebService;
            myWebService = new WebService._base();

            success = myWebService.DeleteANDMaintenance(targetCanCode);

            return Json(new
            {
                Success = success
            }, JsonRequestBehavior.AllowGet);
        }

        public ActionResult UpdateData(string txtCanCode, string txtCanName, string txtRegionId, string txtTargetEcp)
        {
            bool success = true;

            WebView.WebService._base myWebService;
            myWebService = new WebService._base();

            WebService._base.ANDMaintenance newANDMain = new WebService._base.ANDMaintenance();
            newANDMain.AND_NAME = txtCanName;
            newANDMain.REGION_ID = txtRegionId;
            newANDMain.TARGET_ECP = txtTargetEcp;
            success = myWebService.UpdateANDMaintenance(newANDMain, txtCanCode);

            return Json(new
            {
                Success = success
            }, JsonRequestBehavior.AllowGet);
        }

    }
}

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
    public class RegionMaintenanceController : Controller
    {
        public ActionResult Index()
        {
            return View();
        }

        public ActionResult List(string searchKey, int? page)
        {
            WebView.WebService._base myWebService;
            myWebService = new WebService._base();

            WebService._base.OSPRegionMaintenance RegionMain = new WebService._base.OSPRegionMaintenance();
            if (searchKey != null)
            {
                System.Diagnostics.Debug.WriteLine("!!!!!!!!!!!!!!!!!!!!!!!!!!!");
                System.Diagnostics.Debug.WriteLine(searchKey);
                if (searchKey.Equals(""))
                    RegionMain = myWebService.GetOSPRegionMaintenance(0, 100, null);
                else
                {
                    RegionMain = myWebService.GetOSPRegionMaintenance(0, 100, searchKey);
                    ViewBag.searchKey = searchKey;
                }
            }
            else
            {
                System.Diagnostics.Debug.WriteLine("!!!!!!!!!!!!!!!!!!!!!!!!!!!");
                System.Diagnostics.Debug.WriteLine(searchKey);
                RegionMain = myWebService.GetOSPRegionMaintenance(0, 100, null);
                ViewBag.searchKey = "";
            }        

            ViewData["data8"] = RegionMain.RegionMaintenanceList;

            string input = "\\\\adsvr";
            //string input = "\\\\server\\d$\\x\\y\\z\\AAA";
            string output = String.Format("http:{0}", input.Replace("\\d$\\x\\y", String.Empty).Replace("\\", "/"));

            ViewBag.output = output;

            int pageSize = 10;
            int pageNumber = (page ?? 1);
            return View(RegionMain.RegionMaintenanceList.ToPagedList(pageNumber, pageSize));
        }

        public ActionResult NewRegionMaintenance()
        {
            

            return View();
        }

        [HttpPost]
        public ActionResult NewRegionMaintenance(RegionMaintenanceModel m)
        {
            bool success = true;
            WebView.WebService._base myWebService;
            myWebService = new WebService._base();

            bool selected = false;

            WebService._base.RegionMaintenance newRegionMain = new WebService._base.RegionMaintenance();

            newRegionMain.REGION_ID = m.RegionId;
            newRegionMain.REGION_NAME = m.RegionName;
            newRegionMain.REGION_NO = m.RegionNo.ToString();
            success = myWebService.AddRegionMaintenance(newRegionMain);
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
                var query = (from q in ctxData.WV_REGION_MAST
                             where q.REGION_ID == id
                             select q).Single();

                return Json(new
                {
                    Success = true,
                    RegionId = id,
                    RegionName = query.REGION_NAME.Trim(),
                    RegionNo = query.REGION_NO,                    
                    files_list = new
                    {
                        files = fileListStr
                    }
                }, JsonRequestBehavior.AllowGet); //
            }
        }

        [HttpPost]
        public ActionResult DeleteData(string targetRegionId)
        {
            bool success;

            WebView.WebService._base myWebService;
            myWebService = new WebService._base();

            success = myWebService.DeleteRegionMaintenance(targetRegionId);

            return Json(new
            {
                Success = success
            }, JsonRequestBehavior.AllowGet);
        }

        public ActionResult UpdateData(string RegionIdVal, string txtRegionName, string txtRegionNo)
        {
            bool success = true;

            WebView.WebService._base myWebService;
            myWebService = new WebService._base();

            WebService._base.RegionMaintenance newRegionMain = new WebService._base.RegionMaintenance();
            newRegionMain.REGION_NAME = txtRegionName;
            newRegionMain.REGION_NO = txtRegionNo;

            success = myWebService.UpdateRegionMaintenance(newRegionMain, RegionIdVal);

            return Json(new
            {
                Success = success
            }, JsonRequestBehavior.AllowGet);
        }
    }
}

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
    public class StateMaintenanceController : Controller
    {
        public ActionResult Index()
        {
            return View();
        }

        public ActionResult List(string searchKey, int? page)
        {
            WebView.WebService._base myWebService;
            myWebService = new WebService._base();

            WebService._base.OSPStateMaintenance StateMain = new WebService._base.OSPStateMaintenance();
            if (searchKey != null)
            {
                if (searchKey.Equals(""))
                    StateMain = myWebService.GetOSPStateMaintenance(0, 100, null);
                else
                {
                    StateMain = myWebService.GetOSPStateMaintenance(0, 100, searchKey);
                    ViewBag.searchKey = searchKey;
                }
            }
            else
            {
                StateMain = myWebService.GetOSPStateMaintenance(0, 100, null);
                ViewBag.searchKey = "";
            }         

            ViewData["data9"] = StateMain.StateMaintenanceList;

            string input = "\\\\adsvr";
            //string input = "\\\\server\\d$\\x\\y\\z\\AAA";
            string output = String.Format("http:{0}", input.Replace("\\d$\\x\\y", String.Empty).Replace("\\", "/"));

            ViewBag.output = output;

            int pageSize = 10;
            int pageNumber = (page ?? 1);
            return View(StateMain.StateMaintenanceList.ToPagedList(pageNumber, pageSize));
        }

        public ActionResult NewStateMaintenance()
        {
            List<SelectListItem> list = new List<SelectListItem>();
            list.Add(new SelectListItem() { Text = "SBH", Value = "SBH" });
            list.Add(new SelectListItem() { Text = "SEL", Value = "SEL" });
            list.Add(new SelectListItem() { Text = "SWK", Value = "SWK" });
            list.Add(new SelectListItem() { Text = "TGH", Value = "TGH" });
            list.Add(new SelectListItem() { Text = "TMR", Value = "TMR" });
            list.Add(new SelectListItem() { Text = "UTR", Value = "UTR" });
            ViewBag.RegionId = list;

            return View();
        }

        [HttpPost]
        public ActionResult NewStateMaintenance(StateMaintenanceModel m)
        {
            bool success = true;
            WebView.WebService._base myWebService;
            myWebService = new WebService._base();

            bool selected = false;

            WebService._base.StateMaintenance newStateMain = new WebService._base.StateMaintenance();

            newStateMain.STATE_ID = m.StateId;
            newStateMain.STATE_NAME = m.StateName;
            newStateMain.REGION_ID = m.RegionId;
            success = myWebService.AddStateMaintenance(newStateMain);
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
                var query = (from q in ctxData.WV_STATE_MAST
                             where q.STATE_ID == id
                             select q).Single();

                return Json(new
                {
                    Success = true,
                    StateId = id,
                    StateName = query.STATE_NAME,
                    RegionId = query.REGION_ID,
                    files_list = new
                    {
                        files = fileListStr
                    }
                }, JsonRequestBehavior.AllowGet); //
            }
        }

        [HttpPost]
        public ActionResult DeleteData(string targetStateId)
        {
            bool success;

            WebView.WebService._base myWebService;
            myWebService = new WebService._base();

            success = myWebService.DeleteStateMaintenance(targetStateId);

            return Json(new
            {
                Success = success
            }, JsonRequestBehavior.AllowGet);
        }

        public ActionResult UpdateData(string StateIdVal, string txtStateName, string txtRegionId)
        {
            bool success = true;

            WebView.WebService._base myWebService;
            myWebService = new WebService._base();

            WebService._base.StateMaintenance newStateMain = new WebService._base.StateMaintenance();
            newStateMain.STATE_NAME = txtStateName;
            newStateMain.REGION_ID = txtRegionId;

            success = myWebService.UpdateStateMaintenance(newStateMain, StateIdVal);

            return Json(new
            {
                Success = success
            }, JsonRequestBehavior.AllowGet);
        }

    }
}

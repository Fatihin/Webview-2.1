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
    public class PTTMaintenanceController : Controller
    {
        public ActionResult Index()
        {
            return View();
        }

        public ActionResult List(string searchKey, int? page)
        {
            WebView.WebService._base myWebService;
            myWebService = new WebService._base();

            WebService._base.OSPPTTMaintenance PTTMain = new WebService._base.OSPPTTMaintenance();

            if (searchKey != null)
            {
                if (searchKey.Equals(""))
                    PTTMain = myWebService.GetOSPPTTMaintenance(0, 100, null);
                else
                {
                    PTTMain = myWebService.GetOSPPTTMaintenance(0, 100, searchKey);
                    ViewBag.searchKey = searchKey;
                }
            }
            else
            {
                PTTMain = myWebService.GetOSPPTTMaintenance(0, 100, null);
                ViewBag.searchKey = "";
            }         

            ViewData["data7"] = PTTMain.PTTMaintenanceList;

            List<SelectListItem> list2 = new List<SelectListItem>();
            using (Entities ctxData = new Entities())
            {
                
                var sStateName = (from d in ctxData.WV_STATE_MAST 
                            select d);

                list2.Add(new SelectListItem() { Text = "", Value = "" });

                foreach (var a in sStateName.Distinct().OrderBy(it => it.STATE_ID))
                {
                    list2.Add(new SelectListItem() { Text = a.STATE_NAME, Value = a.STATE_ID.ToString() });
                    System.Diagnostics.Debug.WriteLine("B = State Name :" + a.STATE_NAME  + " C = State ID :" + a.STATE_ID );

                }
                //Console.WriteLine(ListSatateName);
                System.Diagnostics.Debug.WriteLine("B = State Name :" + list2);
            }
            ViewBag.StateName = list2;
            
            string input = "\\\\adsvr";
            //string input = "\\\\server\\d$\\x\\y\\z\\AAA";
            string output = String.Format("http:{0}", input.Replace("\\d$\\x\\y", String.Empty).Replace("\\", "/"));

            ViewBag.output = output;

            int pageSize = 10;
            int pageNumber = (page ?? 1);
            return View(PTTMain.PTTMaintenanceList.ToPagedList(pageNumber, pageSize));
        }

        public ActionResult NewPTTMaintenance()
        {
            List<SelectListItem> list2 = new List<SelectListItem>();
            List<SelectListItem> list3 = new List<SelectListItem>();
            using (Entities ctxData = new Entities())
            {

                var sStateName = (from d in ctxData.WV_STATE_MAST
                                  select d);

                list2.Add(new SelectListItem() { Text = "", Value = "" });

                foreach (var a in sStateName.Distinct().OrderBy(it => it.STATE_ID))
                {
                    list2.Add(new SelectListItem() { Text = a.STATE_NAME, Value = a.STATE_ID.ToString() });
                    //System.Diagnostics.Debug.WriteLine("B = State Name :" + a.STATE_NAME + " C = State ID :" + a.STATE_ID);

                }


                var queryRegion = (from d in ctxData.WV_REGION_MAST 
                                  select d);

                list3.Add(new SelectListItem() { Text = "", Value = "" });

                foreach (var a in queryRegion.Distinct().OrderBy(it => it.REGION_ID ))
                {
                    list3.Add(new SelectListItem() { Text = a.REGION_NAME , Value = a.REGION_ID .ToString() });
                   // System.Diagnostics.Debug.WriteLine("B = State Name :" + a.STATE_NAME + " C = State ID :" + a.STATE_ID);

                }
                //Console.WriteLine(ListSatateName);
                //System.Diagnostics.Debug.WriteLine("B = State Name :" + list2);
            }
            ViewBag.StateName = list2;
            ViewBag.RegionName = list3;

            return View();
        }

        [HttpPost]
        public ActionResult NewPTTMaintenance(PTTMaintenanceModel m)
        {
            bool success = true;
            WebView.WebService._base myWebService;
            myWebService = new WebService._base();

            bool selected = false;

            WebService._base.PTTMaintenance newPTTMain = new WebService._base.PTTMaintenance();
            //System.Diagnostics.Debug.WriteLine("State Name :" + m.RegionId);
            newPTTMain.PTT_ID= m.PTTId;
            newPTTMain.PTT_DESC = m.PTTDesc;
            newPTTMain.REGION_ID = m.RegionId;
            newPTTMain.COST_CENTRE = m.CostCentre;
            newPTTMain.STATE_CODE = m.StateCode;
            newPTTMain.STATE_NAME = m.StateName;

            
            success = myWebService.AddPTTMaintenance(newPTTMain);
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
            string dataStateName = "";
            string regionSelect ="";
           

            using (Entities ctxData = new Entities())
            {
                var query = (from q in ctxData.WV_PTT_MAST
                             where q.PTT_ID == id
                             select q).Single();

                var getStateCode = query.STATE_CODE;
                
                var queryStateCode = (from q in ctxData.WV_STATE_MAST
                                      where q.STATE_ID == getStateCode
                                      select new { q.STATE_NAME });

                foreach (var a in queryStateCode.Distinct())
                {
                    dataStateName = dataStateName + a.STATE_NAME;
                }

                string stateListName = "";
                var queryAllStateCode = (from q in ctxData.WV_STATE_MAST
                                         select new {q.STATE_NAME,q.STATE_ID});

                foreach (var b in queryAllStateCode.Distinct())
                {
                    stateListName = stateListName + b.STATE_ID  + ":  " + b.STATE_NAME + "|";
                }
                //System.Diagnostics.Debug.WriteLine("A :" + stateListName);


                //***********Region List***********
                var getRegion = query.REGION_ID;

                var queryRegionSelect = (from a in ctxData.WV_REGION_MAST
                                         where a.REGION_ID == getRegion
                                         select new { a.REGION_NAME });

                foreach (var b in queryRegionSelect)
                {
                    regionSelect = regionSelect + b.REGION_NAME;
                }

                var queryRegion = (from a in ctxData.WV_REGION_MAST
                                   select new { a.REGION_NAME, a.REGION_ID });
                string regionList ="";

                foreach (var b in queryRegion)
                {
                    regionList = regionList + b.REGION_ID + ": " + b.REGION_NAME + "|";
                }

                return Json(new
                {
                    Success = true,
                    PTTId = id,
                    PTTDesc = query.PTT_DESC,
                    RegionId = query.REGION_ID,
                    CostCentre = query.COST_CENTRE ,
                    StateCode = query.STATE_CODE,
                    StateName = dataStateName,
                    StateList = stateListName,
                    dregionList = regionList,
                    dregionSelect = regionSelect,

                    files_list = new
                    {
                        files = fileListStr
                    }
                }, JsonRequestBehavior.AllowGet); //
            }
        }

        [HttpPost]
        public ActionResult DeleteData(string targetPTTId)
        {
            bool success;

            WebView.WebService._base myWebService;
            myWebService = new WebService._base();

            success = myWebService.DeletePTTMaintenance(targetPTTId);

            return Json(new
            {
                Success = success
            }, JsonRequestBehavior.AllowGet);
        }

        public ActionResult UpdateData(string PTTIdVal, string txtPTTDesc, string txtRegionId, string txtCostCentre, string txtStateName)
        {
            bool success = true;

            WebView.WebService._base myWebService;
            myWebService = new WebService._base();

            WebService._base.PTTMaintenance newPTTMain = new WebService._base.PTTMaintenance();
            newPTTMain.PTT_DESC = txtPTTDesc;
            newPTTMain.REGION_ID = txtRegionId;
            newPTTMain.COST_CENTRE = txtCostCentre;
            newPTTMain.STATE_CODE = txtStateName;
            success = myWebService.UpdatePTTMaintenance(newPTTMain, PTTIdVal);

            return Json(new
            {
                Success = success
            }, JsonRequestBehavior.AllowGet);
        }
    }
}

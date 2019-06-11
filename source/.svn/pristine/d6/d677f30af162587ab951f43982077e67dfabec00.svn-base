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
    public class CityMaintenanceController : Controller
    {
        public ActionResult Index()
        {
            return View();
        }

        public ActionResult List(string searchKey, string searchCity, int? page)
        {
            WebView.WebService._base myWebService;
            myWebService = new WebService._base();
            int i = 0;

       //     System.Diagnostics.Debug.WriteLine("Search Name Controler : " + searchName );

            WebService._base.OSPCityMaintenance CityMain = new WebService._base.OSPCityMaintenance();
            if (searchKey != null || searchCity != null)
            {
                if (searchKey == null && searchCity == null)
                {
                    System.Diagnostics.Debug.WriteLine("B : " + searchKey + " " + searchCity);
                    CityMain = myWebService.GetOSPCityMaintenance(0, 100, null, null);
                    ViewBag.searchKey = "";
                    ViewBag.CitySearch = "";
                }
                else
                {
                    //System.Diagnostics.Debug.WriteLine("Search Name P1 : " + searchName + "  " + CitySearch + "  " + searchKey);
                    CityMain = myWebService.GetOSPCityMaintenance(0, 100, searchKey, searchCity);
                    System.Diagnostics.Debug.WriteLine("C : " + searchKey + " " + searchCity);
                    ViewBag.searchKey = searchKey;
                    ViewBag.CitySearch2 = searchCity;
                }
            }
            else
            {
                CityMain = myWebService.GetOSPCityMaintenance(0, 100, null , null);
                ViewBag.searchKey = "";
                ViewBag.CitySearch = "";
            }

            ViewData["data9"] = CityMain.CityMaintenanceList;

            using (Entities ctxData = new Entities())
            {
                //string StateName = "";

                var query = from d in ctxData.WV_CITY_MAST
                            select new { Text = d.DESCRIPTION_STATE.Trim (), Value = d.STATE_CODE };

                List<SelectListItem> listStateID = new List<SelectListItem>();
                listStateID.Add(new SelectListItem() { Text = "", Value = "" });

                foreach (var a in query.Distinct().Distinct().OrderBy(it => it.Text))
                {
                    if (a.Value != null)
                    {
                        listStateID.Add(new SelectListItem() { Text = a.Text, Value = a.Value });
                    }
                }
                ViewBag.LCityID = listStateID;
            }

            string input = "\\\\adsvr";
            //string input = "\\\\server\\d$\\x\\y\\z\\AAA";
            string output = String.Format("http:{0}", input.Replace("\\d$\\x\\y", String.Empty).Replace("\\", "/"));

            ViewBag.output = output;

            int pageSize = 10;
            int pageNumber = (page ?? 1);
            return View(CityMain.CityMaintenanceList.ToPagedList(pageNumber, pageSize));
        }

        public ActionResult NewCityMaintenance()
        {
            List<SelectListItem> listStateID = new List<SelectListItem>();
            using (Entities ctxData = new Entities())
            {
                listStateID.Add(new SelectListItem() { Text = "", Value = "" });

                var queryState = (from q in ctxData.WV_CITY_MAST
                                  where q.STATE_CODE.Trim () != null
                                  select new { Test = q.DESCRIPTION_STATE.Trim () , Value = q.STATE_CODE.Trim () + ":" + q.DESCRIPTION_STATE });

                foreach (var a in queryState.Distinct().OrderBy(it => it.Value))
                {
                    listStateID.Add(new SelectListItem() { Text = a.Test, Value = a.Value });
                }
                ViewBag.StateID = listStateID;
            }
            return View();
        }

        [HttpPost]
        public ActionResult NewCityMaintenance(CityMaintenanceModel m)
        {
            bool success = true;
            WebView.WebService._base myWebService;
            myWebService = new WebService._base();

            bool selected = false;

            WebService._base.CityMaintenance newCityMain = new WebService._base.CityMaintenance();
            //System.Diagnostics.Debug.WriteLine("City NAME : " + m.City_Name + "   State ID : " + m.State_CODE);
            string SCode = m.State_CODE;
            //System.Diagnostics.Debug.WriteLine("Success1 :: " + SCode);
            char s = (':');
            string[] p = SCode.Split(s);

            //System.Diagnostics.Debug.WriteLine("Success :: " + p[0]);
            //System.Diagnostics.Debug.WriteLine("Success :: " + p[1]);
            newCityMain.STATE_CODE = p[0];
            newCityMain.CITY_NAME  = m.City_Name ;
            newCityMain.DESCRIPTION_STATE = p[1];

            success = myWebService.AddCityMaintenance(newCityMain);
            selected = true;
            //System.Diagnostics.Debug.WriteLine("Success :: " + success);

            //System.Diagnostics.Debug.WriteLine("ModelState :: " + ModelState.IsValid);
            //System.Diagnostics.Debug.WriteLine("selected :: " + selected);
            if (ModelState.IsValid && selected)
            {

                if (success == true)
                {
                    System.Diagnostics.Debug.WriteLine("pizal 1");
                    return RedirectToAction("NewSave");
                }
                else
                {
                    System.Diagnostics.Debug.WriteLine("pizal 1");
                    return RedirectToAction("NewSaveFail"); // store to db failed.

                }
            }
            List<SelectListItem> listStateID = new List<SelectListItem>();
            using (Entities ctxData = new Entities())
            {
                listStateID.Add(new SelectListItem() { Text = "", Value = "" });

                var queryState = (from q in ctxData.WV_STATE_MAST
                                  orderby q.STATE_ID descending
                                  select new { q.STATE_ID, q.STATE_NAME });

                foreach (var a in queryState.Distinct())
                {
                    listStateID.Add(new SelectListItem() { Text = a.STATE_NAME, Value = a.STATE_ID });
                }
                ViewBag.StateID = listStateID;
            }

            return View(m);
        }

        public ActionResult NewSave()
        {
            return View();
        }

        [HttpPost]
        public ActionResult GetDetails(string id, string SCode)
        {
            string fileListStr = "";
            //string StateName = "";
            string StateListCity ="";
            string City_Name = "";
            //string State_Name = "";

            using (Entities ctxData = new Entities())
            {

                System.Diagnostics.Debug.WriteLine("ID :" + id + "  SCODE :" + SCode);

                var query = (from q in ctxData.WV_CITY_MAST
                             where q.CITY_NAME.Trim() == id.Trim() && q.STATE_CODE.Trim() == SCode.Trim()
                             select q).Single (); //new { q.CITY_NAME, q.STATE_CODE });

                System.Diagnostics.Debug.WriteLine("Query : " + query.ToString());
        
                var queryCity = (from q in ctxData.WV_CITY_MAST
                                 select new { Value = q.STATE_CODE.Trim()});

                foreach (var q in queryCity.Distinct().OrderBy(it => it.Value))
                {
                    //System.Diagnostics.Debug.WriteLine("Detail : " + q.Value);
                    StateListCity = StateListCity + q.Value + "|";
                }

                return Json(new
                {
                    Success = true,
                    CityId = id + ":" +  query.STATE_CODE,
                    CityName = query.CITY_NAME,
                    StateID = query.STATE_CODE,
                    StateName = query.DESCRIPTION_STATE,
                    StateListCity = StateListCity,

                    files_list = new
                    {
                        files = fileListStr
                    }
                }, JsonRequestBehavior.AllowGet); //
            }
        }

        [HttpPost]
        public ActionResult DeleteData(string targetCity)
        {
            bool success;

            WebView.WebService._base myWebService;
            myWebService = new WebService._base();

            success = myWebService.DeleteCityMaintenance(targetCity);
            System.Diagnostics.Debug.WriteLine("DELETE !: " + targetCity);
            return Json(new
            {
                Success = success
            }, JsonRequestBehavior.AllowGet);
        }

        [HttpPost]
        public ActionResult DescriptionName(string StateCode)
        {
            string DescripName = "";

            using (Entities ctxData = new Entities())
            {
                System.Diagnostics.Debug.WriteLine("StateCODE : " + StateCode);
                var query = (from q in ctxData.WV_CITY_MAST
                             where q.STATE_CODE.Trim() == StateCode.Trim()
                             select new {Value = q.DESCRIPTION_STATE.Trim()});

                foreach(var i in query.Distinct().OrderBy (it => it.Value))
                {
                    DescripName = DescripName + i.Value + ":";
                }

                System.Diagnostics.Debug.WriteLine("StateCODE : " + DescripName);
                return Json(new
                {
                    Des = DescripName,
                }, JsonRequestBehavior.AllowGet);
            }
           
        }

        public ActionResult UpdateData(string CityIdVal, string txtCity, string txtStateID, string txtSName)
        {
            bool success = true;
            System.Diagnostics.Debug.WriteLine("pizal");

            WebView.WebService._base myWebService;
            myWebService = new WebService._base();

            WebService._base.CityMaintenance newCityMain = new WebService._base.CityMaintenance();
            newCityMain.STATE_CODE = txtStateID;
            newCityMain.CITY_NAME = txtCity;
            newCityMain.DESCRIPTION_STATE = txtSName;

            success = myWebService.UpdateCityMaintenance(newCityMain, CityIdVal);
            System.Diagnostics.Debug.WriteLine("pizal" + success );
            return Json(new
            {
                Success = success
            }, JsonRequestBehavior.AllowGet);
        }
    }
}

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
    public class ExchangeMaintenanceController : Controller
    {
        public ActionResult Index()
        {
            return View();
        }

        public ActionResult List(string searchKey, string pttID, int? page)
        {
            WebView.WebService._base myWebService;
            myWebService = new WebService._base();

            WebService._base.OSPExchangeMaintenance ExcMain = new WebService._base.OSPExchangeMaintenance();
            if (searchKey != null || pttID != null)
            {
                if (searchKey == "" && pttID == "")
                    ExcMain = myWebService.GetOSPExchangeMaintenance(0, 50000, null, null);

                else
                {
                    ExcMain = myWebService.GetOSPExchangeMaintenance(0, 50000, searchKey, pttID);
                    ViewBag.searchKey = searchKey;
                    ViewBag.pttID2 = pttID;
                }
            }
            else
            {
                ExcMain = myWebService.GetOSPExchangeMaintenance(0, 50000, null, null);
                ViewBag.searchKey = "";
                ViewBag.pttID2 = "";
            }        

            ViewData["data10"] = ExcMain.ExchangeMaintenanceList;

            using (Entities ctxData = new Entities())
            {
                List<SelectListItem> list2 = new List<SelectListItem>();
                var query3 = from p in ctxData.WV_EXC_MAST
                             //orderby p.PTT_ID ascending
                             select new { p.PTT_ID };
                list2.Add(new SelectListItem() { Text = "", Value = "" });
                foreach (var a in query3.Distinct().OrderBy(it => it.PTT_ID))
                {
                    list2.Add(new SelectListItem() { Text = a.PTT_ID, Value = a.PTT_ID });
                }
                ViewBag.PttId = list2;
            }
            string input = "\\\\adsvr";
            //string input = "\\\\server\\d$\\x\\y\\z\\AAA";
            string output = String.Format("http:{0}", input.Replace("\\d$\\x\\y", String.Empty).Replace("\\", "/"));

            ViewBag.output = output;

            //return View();
            int pageSize = 10;
            int pageNumber = (page ?? 1);
            return View(ExcMain.ExchangeMaintenanceList.ToPagedList(pageNumber, pageSize));

          
        }

        public ActionResult NewExchangeMaintenance()
        {
            using (Entities ctxData = new Entities())
            {
                List<SelectListItem> list2 = new List<SelectListItem>();
                var query3 = from p in ctxData.WV_EXC_MAST
                             //orderby p.PTT_ID ascending
                             select new { p.PTT_ID  };

                foreach (var a in query3.Distinct().OrderBy(it => it.PTT_ID))
                {
                    list2.Add(new SelectListItem() { Text = a.PTT_ID, Value = a.PTT_ID });
                }
                ViewBag.PttId = list2;
                ViewData["Submarkets"] = new SelectList(list2, "id", "name");
                List<SelectListItem> listEXC = new List<SelectListItem>();
                 //foreach (var a in query3 .Distinct ())
                 //{
                    listEXC.Add(new SelectListItem() { Text = "M", Value = "M" });
                    listEXC.Add(new SelectListItem() { Text = "S", Value = "S" });
                // }

                ViewBag.EXCType = listEXC;

                List<SelectListItem> list3 = new List<SelectListItem>();
                var queryState = (from q in ctxData.WV_STATE_MAST
                                  orderby q.STATE_ID
                             select q);
                foreach (var a in queryState)
                {
                    list3.Add(new SelectListItem() { Text = a.STATE_ID, Value = a.STATE_ID });

                }
                ViewBag.StateId = list3;

                List<SelectListItem> list4 = new List<SelectListItem>();
                list4.Add(new SelectListItem() { Text = "U - URBAN", Value = "U" });
                list4.Add(new SelectListItem() { Text = "R - RURAL", Value = "R" });
                list4.Add(new SelectListItem() { Text = "M - METROPOLITAN", Value = "M" });
                list4.Add(new SelectListItem() { Text = "S - SEMI-URBAN", Value = "S" });
                list4.Add(new SelectListItem() { Text = "L - SPECIAL", Value = "L" });
                ViewBag.AreaClass = list4;

                List<SelectListItem> list5 = new List<SelectListItem>();
                var querySegment = (from q in ctxData.WV_EXC_MAST
                                    orderby q.SEGMENT
                                    select new { q.SEGMENT });
                foreach (var a in querySegment.Distinct ())
                {
                        list5.Add(new SelectListItem() { Text = a.SEGMENT, Value = a.SEGMENT });
                }

                ViewBag.Segment = list5;
            }
            //load Exc_type
           // List<SelectListItem> listEXC = new List<SelectListItem>();
            //using (Entities ctxData = new Entities())
            //{
            //    var query = from p in ctxData.REF_COM_EXCHABB
            //                select new { Text = p.PL_VALUE, Value = p.PL_VALUE};

            //    foreach (var a in )
            //    {
            //        listEXC.Add(new SelectListItem() { Text = a.Text, Value = a.Value.ToString() });
            //    }

            //    ViewBag.EXCType = listEXC;
            //}

            List<SelectListItem> ListCity = new List<SelectListItem>();
            using (Entities ctxData = new Entities())
            {
                string allCity = "";
                var city = (from d in ctxData.WV_CITY_MAST
                            where d.STATE_CODE == "JH"
                            select new { Value = d.CITY_NAME.Trim(), Text = d.CITY_NAME.Trim() });

                foreach (var a in city.Distinct().OrderBy(it => it.Text))
                {
                    ListCity.Add(new SelectListItem() { Text = a.Text , Value = a.Value });
                }
                ViewBag.City = ListCity;
            }

            return View();
        }

        [HttpPost]
        public ActionResult NewExchangeMaintenance(ExchangeMaintenanceModel m)
        {
            bool success = true;
            WebView.WebService._base myWebService;
            myWebService = new WebService._base();

            //string Success = "false";
            bool selected = false;

            WebService._base.ExchangeMaintenance newExcMain = new WebService._base.ExchangeMaintenance();
            System.Diagnostics.Debug.WriteLine("A :" + m.PTT_ID);
            newExcMain.EXC_ABB = m.EXC_ABB;
            newExcMain.EXC_NAME = m.EXC_NAME;
            newExcMain.PTT_ID = m.PTT_ID;
            newExcMain.LOC_NO = m.LOC_NO;
            newExcMain.STATE_ID = m.STATE_ID;
            newExcMain.MAIN_NO = m.MAIN_NO.ToString();
            newExcMain.CAT_CODE = m.CAT_CODE;
            newExcMain.EXC_TYPE = m.EXC_TYPE;
            newExcMain.EXC_ID = m.EXC_ID;
            newExcMain.SEGMENT = m.SEGMENT;
            newExcMain.ADD1 = m.ADDR1;
            newExcMain.ADD2 = m.ADDR2;
            newExcMain.CITY = m.CITY;

            success = myWebService.AddExchangeMaintenance(newExcMain);
            selected = true;
            System.Diagnostics.Debug.WriteLine("C : "+m.SEGMENT);
            //if (ModelState.IsValid && selected)
            //{
            
                if (success == true)
                    return RedirectToAction("NewSave");
                else
                    return RedirectToAction("NewSaveFail"); // store to db failed.
            //}

            //return View(m);
        }

        public ActionResult NewSave()
        {
            return View();
        }

        [HttpPost]
        public ActionResult GetDetails(string id)
        {
            string fileListStr = "";
            string Exctype = "";
            string PTTList = "";
            string StateList = "";
            string City = "";
            string dataSegment="";
            string dataSegment_2 = "";
            string dataPttid = "";
            string CityLists = "";

            using (Entities ctxData = new Entities())
            {
                var query = (from q in ctxData.WV_EXC_MAST
                             where q.EXC_ABB == id
                             select q).Single();

                dataSegment = query.SEGMENT;

                var querySegment = from q in ctxData.WV_EXC_MAST
                                //    where q.SEGMENT != dataSegment
                                    select new {Value = q.SEGMENT};
                
                var queryExcType = from p in ctxData.REF_COM_EXCHABB
                                   where p.SEGMENT == dataSegment
                                   select new { p.PL_VALUE, p.PL_NUM, p.SEGMENT  };


                foreach (var a in queryExcType)
                {
                    Exctype = Exctype + a.PL_VALUE  + "|";
                }

                foreach (var b in querySegment.Distinct().OrderBy (it => it.Value))
                {
                    if (b != null)
                    {
                        dataSegment_2 = dataSegment_2 + b.Value  + ": " + b.Value + "|";
                    }
                }
                //foreach (var a in ctxData.WV_EXC_MAST)
                //{
                //    Exctype = Exctype + a.EXC_TYPE + "|";
                //}

                var queryPTTExc = from p in ctxData.WV_PTT_MAST
                                   select new { p.PTT_ID };

                foreach (var a in queryPTTExc.Distinct().OrderBy(it => it.PTT_ID))
                {
                    PTTList = PTTList + a.PTT_ID + ":  " + a.PTT_ID + "|";
                }

                var queryState = from p in ctxData.WV_STATE_MAST
                                  select new { p.STATE_ID };

                foreach (var a in queryState.Distinct().OrderBy(it => it.STATE_ID))
                {
                    StateList = StateList + a.STATE_ID + ":";
                }
                var queryPTTID = from p in ctxData.WV_EXC_MAST 
                                 select new { p.PTT_ID};

                foreach (var a in queryPTTID.Distinct())
                {
                   dataPttid = dataPttid + a.PTT_ID + ":  " + a.PTT_ID + "|";
                }

                if(query.STATE_ID != "")
                {
                    var S = query.STATE_ID;

                    var city = (from d in ctxData.WV_STATE_MAST
                            where d.STATE_ID.Trim() == S.Trim()
                            select d).Single();

                    var GDS_State = city.GDS_STATE_CODE;

                    var CityName = (from d in ctxData.WV_CITY_MAST
                                where d.STATE_CODE.Trim() == GDS_State.Trim()
                                select new { Value = d.CITY_NAME });

                     foreach (var i in CityName.Distinct().OrderBy(it => it.Value))
                     {
                       CityLists = CityLists + i.Value + ":";
                     }
                }

                return Json(new
                {
                    Success = true,
                    ExcAbb = id,
                    ExcName = query.EXC_NAME,
                    PttId = query.PTT_ID,
                    LocNo = query.LOC_NO,
                    StateId = query.STATE_ID,
                    MainNo = query.MAIN_NO,
                    AreaClass = query.CAT_CODE,
                    ExcType = query.EXC_TYPE,
                    ExcId = query.EXC_ID,
                    Segment = query.SEGMENT,
                    Segment_2 = dataSegment_2,
                    Exctypes = Exctype,
                    datapttid = dataPttid,
                    PTTList = PTTList,
                    StateList = StateList,
                    Add = query.SITE_ADDRESS,
                    Add2 = query .SITE_ADDRESS_2,
                    City = query.CITY,
                    CityLists = CityLists,

                    files_list = new
                    {
                        files = fileListStr
                    }
                }, JsonRequestBehavior.AllowGet); //
            }
        }

        [HttpPost]
        public ActionResult AddCity(string Stateid)
        {
            string CITY = "";
            string CityLists = "";
           // System.Diagnostics.Debug.WriteLine("CITY: " + Stateid);
            using (Entities ctxState = new Entities())
            {
                string allCity = "";
                System.Diagnostics.Debug.WriteLine("StateID : " + Stateid);
                var city = (from d in ctxState.WV_STATE_MAST
                            where d.STATE_ID.Trim() == Stateid.Trim()
                            select d).Single();

                var GDS_State = city.GDS_STATE_CODE;
                System.Diagnostics.Debug.WriteLine("GDS_State : " + GDS_State);

                var CityName = (from d in ctxState.WV_CITY_MAST
                                where d.STATE_CODE.Trim() == GDS_State.Trim()
                                select new { Value = d.CITY_NAME.Trim()});

                foreach (var i in CityName.Distinct().OrderBy(it => it.Value))
                {
                    CityLists = CityLists + i.Value + ":" ;
                }
               // System.Diagnostics.Debug.WriteLine("CITY: " + CityLists );
                return Json( new
                {
                    Success = true,
                    CITY = CityLists
            
                 }, JsonRequestBehavior.AllowGet);
           }
        }

        [HttpPost]
        public ActionResult Pttid(string Stateid)
        {
            string pttid = "";

            using (Entities ctxState = new Entities())
            {
                string allpttid = "";
                var city = (from d in ctxState.WV_EXC_MAST
                            select new { d.PTT_ID  });
                foreach (var a in city)
                {
                    allpttid = allpttid + a.PTT_ID  + "|";
                }

                return Json(new
                {
                    Success = true,
                    pttId = allpttid,

                }, JsonRequestBehavior.AllowGet);
            }
        }

        [HttpPost]
        public ActionResult DeleteData(string targetExcAbb)
        {
            bool success;

            WebView.WebService._base myWebService;
            myWebService = new WebService._base();

            success = myWebService.DeleteExchangeMaintenance(targetExcAbb);

            return Json(new
            {
                Success = success
            }, JsonRequestBehavior.AllowGet);
        }

        public ActionResult UpdateData(string ExcAbbVal, string txtExcName, string txtPttId, string txtLocNo, string txtStateId, string txtMainNo, string txtAreaClass, string txtExcType, string txtExcId, string txtSegment, string txtAdd1, string txtAdd2, string txtCity)
        {
            bool success = true;

            WebView.WebService._base myWebService;
            myWebService = new WebService._base();

            WebService._base.ExchangeMaintenance newExcMain = new WebService._base.ExchangeMaintenance();
            newExcMain.EXC_NAME = txtExcName;
            newExcMain.PTT_ID = txtPttId;
            newExcMain.LOC_NO = txtLocNo;
            newExcMain.STATE_ID = txtStateId;
            newExcMain.MAIN_NO = txtMainNo;
            newExcMain.CAT_CODE = txtAreaClass;
            newExcMain.EXC_TYPE = txtExcType;
            newExcMain.EXC_ID = txtExcId;
            newExcMain.SEGMENT = txtSegment;
           // newExcMain.CITY = txtExcId; // city
            newExcMain.ADD1 = txtAdd1;
            newExcMain.ADD2 = txtAdd2;
            newExcMain.CITY = txtCity;

            System.Diagnostics.Debug.WriteLine("A : "+ newExcMain.ADD1);
            success = myWebService.UpdateExchangeMaintenance(newExcMain, ExcAbbVal);

            return Json(new
            {
                Success = success
            }, JsonRequestBehavior.AllowGet);
        }
    }
}

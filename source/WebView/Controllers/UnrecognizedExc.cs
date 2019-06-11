using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using PagedList;

namespace WebView
{
    public class UnrecognizedExcController : Controller
    {
        //
        // GET: /UnrecognizedExc/

        public ActionResult Index()
        {
            return View();
        }
        public ActionResult List(string searchKey, int? page, string searchExcAbb, string searchSegment, string searchFType, string searchFState)
        {
            WebView.WebService._base myWebService;
            myWebService = new WebService._base();

            WebService._base.OSPUnrecognizedExc ExcMain = new WebService._base.OSPUnrecognizedExc();
            System.Diagnostics.Debug.WriteLine("A");

            if (searchKey == null)
            {
                if (searchKey == null && searchExcAbb == null && searchSegment == null && searchFType == null && searchFState == null)
                {
                    System.Diagnostics.Debug.WriteLine("b");
                    ExcMain = myWebService.GetOSPUnrecognizedExc(0, 100000, null, null, null, null, null);
                    ViewBag.searchKey2 = "";
                    ViewBag.sExcAbb2 = "Select";
                    ViewBag.sSegment2 = "Select";
                    ViewBag.sFTpye2 = "Select";
                    ViewBag.sFState2 = "Select";
                }
                else
                {
                    System.Diagnostics.Debug.WriteLine("c");
                    searchKey = "";
                    ExcMain = myWebService.GetOSPUnrecognizedExc(0, 100000, searchKey, searchExcAbb, searchSegment, searchFType, searchFState);
                    
                    ViewBag.searchKey2 = searchKey;
                    ViewBag.sExcAbb2 = searchExcAbb;
                    ViewBag.sSegment2 = searchSegment;
                    ViewBag.sFTpye2 = searchFType;
                    ViewBag.sFState2 = searchFState;
                }
            }
            else
            {

                if (searchKey == "" && searchExcAbb.Equals("Select") && searchSegment.Equals("Select") && searchFType.Equals("Select") && searchFState.Equals("Select"))
                {
                    System.Diagnostics.Debug.WriteLine("d");
                    ExcMain = myWebService.GetOSPUnrecognizedExc(0, 100000, null, null, null, null, null);
                    ViewBag.searchKey2 = searchKey;
                    ViewBag.sExcAbb2 = searchExcAbb;
                    ViewBag.sSegment2 = searchSegment;
                    ViewBag.sFTpye2 = searchFType;
                    ViewBag.sFState2 = searchFState;
                }
                else
                {
                    System.Diagnostics.Debug.WriteLine("c");
                    ExcMain = myWebService.GetOSPUnrecognizedExc(0, 100000, searchKey, searchExcAbb, searchSegment, searchFType, searchFState);
                 
                    ViewBag.searchKey2 = searchKey;
                    ViewBag.sExcAbb2 = searchExcAbb;
                    ViewBag.sSegment2 = searchSegment;
                    ViewBag.sFTpye2 = searchFType;
                    ViewBag.sFState2 = searchFState;
                }
            }


            using (Entities ctxData = new Entities())
            {
                //LIST EXC MAST
                var qExcMast = (from Mast in ctxData.UNRECOGNIZED_EXC
                                select new { Mast.EXC_ABB });

                List<SelectListItem> listExcMast = new List<SelectListItem>();
                listExcMast.Add(new SelectListItem() { Text = "", Value = "Select" });

                foreach (var excMast in qExcMast.Distinct().OrderBy(it => it.EXC_ABB))
                {
                    listExcMast.Add(new SelectListItem() { Text = excMast.EXC_ABB, Value = excMast.EXC_ABB });
                }
                ViewBag.sExcAbb = listExcMast;

                //LIST SEGMENT
                var qSeg = (from a in ctxData.UNRECOGNIZED_EXC
                            select new { a.SEGMENT });

                List<SelectListItem> listSeg = new List<SelectListItem>();
                listSeg.Add(new SelectListItem() { Text = "", Value = "Select" });

                foreach (var Segm in qSeg.Distinct().OrderBy(it => it.SEGMENT))
                {
                    listSeg.Add(new SelectListItem() { Text = Segm.SEGMENT, Value = Segm.SEGMENT });
                }
                ViewBag.sSegment = listSeg;

                //LIST FEATURE STATE
                var qFeatyure = (from a in ctxData.UNRECOGNIZED_EXC
                                 select new { a.FEATURE_STATE });

                List<SelectListItem> listFeaState = new List<SelectListItem>();
                listFeaState.Add(new SelectListItem() { Text = "", Value = "Select" });

                foreach (var a in qFeatyure.Distinct().OrderBy(it => it.FEATURE_STATE))
                {
                    listFeaState.Add(new SelectListItem() { Text = a.FEATURE_STATE, Value = a.FEATURE_STATE });
                }
                ViewBag.sFState = listFeaState;

                //LIST FEATURE Type
                var qFType = (from a in ctxData.G3E_FEATURE
                              select new { a.G3E_USERNAME });

                List<SelectListItem> ListFType = new List<SelectListItem>();
                ListFType.Add(new SelectListItem() { Text = "", Value = "Select" });

                foreach (var a in qFType.Distinct())
                {
                    ListFType.Add(new SelectListItem() { Text = a.G3E_USERNAME.ToString(), Value = a.G3E_USERNAME.ToString() });
                }
                ViewBag.sFTpye = ListFType;

            }

            int pageSize = 10;
            int pageNumber = (page ?? 1);
            return View(ExcMain.Exc.ToPagedList(pageNumber, pageSize));

        }
        public ActionResult GetDetails(int ID, string SName)
        {

            System.Diagnostics.Debug.WriteLine("getDetails");
            string fileListStr = "";
            int fid;
            int id;
            string segment = "";
            string exc = "";
            string fState = "";
            string sName = "";
            string fno = "";
            string fType = "";
            string ValidExc = "";

            using (Entities ctxData = new Entities())
            {
                var query = (from a in ctxData.UNRECOGNIZED_EXC
                             where a.G3E_ID == ID && a.SCHEME_NAME.Trim().ToUpper() == SName.Trim().ToUpper()
                             select a).Single();

                var queryValidExc = (from a in ctxData.WV_EXC_MAST
                                     select new { a.EXC_ABB });

                foreach (var a in queryValidExc.Distinct().OrderBy(it => it.EXC_ABB))
                {
                    ValidExc = ValidExc + a.EXC_ABB + ";";
                }

                System.Diagnostics.Debug.WriteLine(query.FEATURE_STATE);
                return Json(new
                {
                    Success = true,
                    fid = query.G3E_FID,
                    id = query.G3E_ID,
                    segment = query.SEGMENT,
                    exc = query.EXC_ABB,
                    fState = query.FEATURE_STATE,
                    sName = query.SCHEME_NAME,
                    fno = query.G3E_FNO,
                    fType = query.FEATURE_TYPE,
                    ValidExc = ValidExc,

                    files_list = new
                    {
                        files = fileListStr
                    }
                }, JsonRequestBehavior.AllowGet); //
            }
        }
        public ActionResult UpdateData(int Fid, int id, string Exc, string ValidExc)
        {
            bool success = true;

            WebView.WebService._base myWebService;
            myWebService = new WebService._base();

            System.Diagnostics.Debug.WriteLine("Update 1  = " + Fid + "EXC ABB = " + Exc + "Valid =" + ValidExc);
            WebService._base.UnrecognizedExc ExcMain = new WebService._base.UnrecognizedExc();

            ExcMain.G3E_ID = id;
            ExcMain.G3E_FID = Fid;
            ExcMain.EXC_ABB = Exc;

            success = myWebService.UpdateUrecognizedExc(ExcMain, ValidExc);

            return Json(new
            {
                Success = success
            }, JsonRequestBehavior.AllowGet);
        }
    }
}


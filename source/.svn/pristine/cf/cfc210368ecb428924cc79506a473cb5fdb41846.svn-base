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
            if (searchKey != null)
            {
                if (searchExcAbb.Equals("") && searchSegment.Equals("") && searchFType.Equals("") && searchFState.Equals(""))
                {
                    System.Diagnostics.Debug.WriteLine("b");
                    ExcMain = myWebService.GetOSPUnrecognizedExc(0, 100000, null, null, null, null, null);
                    ViewBag.searchKey = "";
                    ViewBag.searchExcAbb = "";
                    ViewBag.searchSegment = "";
                    ViewBag.searchFType = "";
                    ViewBag.searchFState = "";
                }
                else
                {
                    System.Diagnostics.Debug.WriteLine("c");
                    ExcMain = myWebService.GetOSPUnrecognizedExc(0, 100000, searchKey, searchExcAbb, searchSegment, searchFType, searchFState);
                    ViewBag.searchKey = searchKey;
                    ViewBag.searchExcAbb = searchExcAbb;
                    ViewBag.searchSegment = searchSegment;
                    ViewBag.searchFType = searchFType;
                    ViewBag.searchFState = searchFState;
                }
            }
            else
            {
                System.Diagnostics.Debug.WriteLine("d");
                ExcMain = myWebService.GetOSPUnrecognizedExc(0, 100000, null, null, null, null, null);
                ViewBag.searchKey = "";
                ViewBag.searchExcAbb = "";
                ViewBag.searchSegment = "";
                ViewBag.searchFType = "";
                ViewBag.searchFState = "";
            }

            using (Entities ctxData = new Entities())
            {
                //LIST EXC MAST
                var qExcMast = (from Mast in ctxData.UNRECOGNIZED_EXC
                                select new { Mast.EXC_ABB });

                List<SelectListItem> listExcMast = new List<SelectListItem>();
                foreach (var excMast in qExcMast.Distinct().OrderBy(it => it.EXC_ABB ))
                {
                    listExcMast.Add(new SelectListItem() { Text = excMast.EXC_ABB , Value = excMast.EXC_ABB });
                }
                ViewBag.sExcAbb = listExcMast;

                //LIST SEGMENT
                var qSeg = (from a in ctxData.UNRECOGNIZED_EXC
                                select new { a.SEGMENT });

                List<SelectListItem> listSeg = new List<SelectListItem>();

                foreach (var Segm in qSeg.Distinct().OrderBy(it => it.SEGMENT ))
                {
                    listSeg.Add(new SelectListItem() { Text = Segm.SEGMENT , Value = Segm.SEGMENT  });
                }
                ViewBag.sSegment = listSeg;

                //LIST FEATURE STATE
                var qFeatyure = (from a in ctxData.G3E_FEATURE
                                 select new { a.G3E_USERNAME, a.G3E_FNO });

                List<SelectListItem> listFeaState = new List<SelectListItem>();
                foreach (var a in qFeatyure.Distinct())
                {
                    listFeaState.Add(new SelectListItem() { Text = a.G3E_USERNAME + "        [" + Convert.ToString(a.G3E_FNO) + "]   ", Value = Convert.ToString(a.G3E_FNO) });
                }
                ViewBag.sFType = listFeaState;

            }

            int pageSize = 10;
            int pageNumber = (page ?? 1);
            return View(ExcMain.Exc.ToPagedList(pageNumber, pageSize));
            
        }
        public ActionResult GetDetails(int ID, string SName){

            System.Diagnostics.Debug.WriteLine("getDetails");
            string fileListStr = "";
            string fid = "";
            string id = "";
            string segment = "";
            string exc = "";
            string fState = "";
            string sName = "";
            string fno = "";
            string fType = "";
            string ValidExc = "";

            using (Entities ctxData = new Entities ())
            {
                var query = (from a in ctxData.UNRECOGNIZED_EXC 
                                 where a.G3E_ID == ID && a.SCHEME_NAME.Trim().ToUpper() == SName.Trim().ToUpper()
                                 select a).Single ();

                var queryValidExc = (from a in ctxData.WV_EXC_MAST
                                     select new {a.EXC_ABB});

                foreach (var a in queryValidExc.Distinct().OrderBy(it => it.EXC_ABB)) {
                    ValidExc = ValidExc + a.EXC_ABB + ";";
                }

                System.Diagnostics.Debug.WriteLine(query.FEATURE_STATE);
                return Json(new
                {
                    Success = true,
                    fid =query.G3E_FID,
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
    }
}



using System.Collections.Generic;
using System.Linq;
using System.Web.Mvc;

namespace WebView.Controllers
{

    public class BOQ_MAINController : Controller
    {

        // GET: /BOQ_MAIN/

        public ActionResult Index()
        {
            return View();
        }

        public ActionResult BOQ_MAIN_List(string searchNo, string searchIDNo, string searchType, string searchIDType, string searchExc, string searchIDExc, string searchYear, string searchIDYear)
        {
            WebView.WebService._base myWebService;
            myWebService = new WebService._base();

            WebService._base.OSPBOQ_MAIN BOQ_MAIN = new WebService._base.OSPBOQ_MAIN();

            if (searchNo != null && searchType != null && searchExc != null && searchYear != null)
            {
                if (searchNo.Equals("") && searchType.Equals("") && searchExc.Equals("") && searchYear.Equals(""))
                    BOQ_MAIN = myWebService.GetOSPBOQ_MAIN(0, 100, null, null, null, null, null, null, null, null);
                else
                {
                    BOQ_MAIN = myWebService.GetOSPBOQ_MAIN(0, 100, searchNo, searchIDNo, searchType, searchIDType, searchExc, searchIDExc, searchYear, searchIDYear);
                    ViewBag.searchNo = searchNo;
                    ViewBag.searchType = searchType;
                    ViewBag.searchExc = searchExc;
                    ViewBag.searchYear = searchYear;
                }
            }
            else
            {
                BOQ_MAIN = myWebService.GetOSPBOQ_MAIN(0, 100, null, null, null, null, null, null, null, null);
                ViewBag.searchNo = "";
                ViewBag.searchType = "";
                ViewBag.searchExc = "";
                ViewBag.searchYear = "";
            }

            ViewData["data2"] = BOQ_MAIN.BOQ_MAINList;


            using (Entities ctxData = new Entities())
            {
                var query = from p in ctxData.WV_BOQ_DATA
                            orderby p.SCH_NO ascending
                            select new { Text = p.SCH_NO, Value = p.SCH_NO };

                List<SelectListItem> list = new List<SelectListItem>();
                list.Add(new SelectListItem() { Text = "Select", Value = "Select" });
                foreach (var a in query.Distinct().OrderBy(it => it.Value))
                {  
                    if(a.Value != null ) 
                    list.Add(new SelectListItem() { Text = a.Text.ToString(), Value = a.Value.ToString() });
                }

                ViewBag.x_SCH_NO = list;
            }

            using (Entities ctxData = new Entities())
            {
                var query = from p in ctxData.WV_BOQ_DATA
                            orderby p.SCH_TYPE ascending
                            select new { Text = p.SCH_TYPE, Value = p.SCH_TYPE };

                List<SelectListItem> list = new List<SelectListItem>();
                list.Add(new SelectListItem() { Text = "Select", Value = "Select" });
                foreach (var a in query.Distinct().OrderBy(it => it.Value))
                {
                    if (a.Value != null)
                    list.Add(new SelectListItem() { Text = a.Text.ToString(), Value = a.Value.ToString() });
                }

                ViewBag.x_SCH_TYPE = list;
            }

            using (Entities ctxData = new Entities())
            {
                var query = from p in ctxData.WV_BOQ_DATA
                            orderby p.EXC_ABB ascending
                            select new { Text = p.EXC_ABB, Value = p.EXC_ABB };

                List<SelectListItem> list = new List<SelectListItem>();
                list.Add(new SelectListItem() { Text = "Select", Value = "Select" });
                foreach (var a in query.Distinct().OrderBy(it => it.Value))
                {
                    if (a.Value != null)
                    list.Add(new SelectListItem() { Text = a.Text.ToString(), Value = a.Value.ToString() });
                }

                ViewBag.x_EXC_ABB = list;
            }

            using (Entities ctxData = new Entities())
            {
                var query = from p in ctxData.WV_BOQ_DATA
                            orderby p.YEAR_INSTALL ascending
                            select new { Text = p.YEAR_INSTALL, Value = p.YEAR_INSTALL };

                List<SelectListItem> list = new List<SelectListItem>();
                list.Add(new SelectListItem() { Text = "Select", Value = "Select" });
                foreach (var a in query.Distinct().OrderBy(it => it.Value))
                {
                    if (a.Value != null)
                    list.Add(new SelectListItem() { Text = a.Text.ToString(), Value = a.Value.ToString() });
                }

                ViewBag.x_YEAR_INSTALL = list;
            }

            return View();
        }

        public ActionResult BOQ_MAIN_Excel(string schemename)
        {
            WebView.WebService._base myWebService;
            myWebService = new WebService._base();

            WebService._base.OSPBOQ_MAIN_EXCEL BOQ_MAIN_EXCEL = new WebService._base.OSPBOQ_MAIN_EXCEL();
            
            BOQ_MAIN_EXCEL = myWebService.GetOSPBOQ_MAIN_Excel(0, 100, schemename);            

            ViewData["data2"] = BOQ_MAIN_EXCEL.BOQ_MAIN_Excel;
           
            return View();
        }     

        public ActionResult GetDetails_BOQ_MAIN(string id)
        {          
 

            using (Entities ctxData = new Entities())
            {
                string[] arFea = new string[2];
                char[] splitter = { '-' };
                arFea = id.ToString().Split(splitter);

                string id_scheme = arFea[0].ToString();
                string id_pu = arFea[1].ToString();

                var query = (from p in ctxData.WV_BOQ_DATA
                             where p.SCHEME_NAME == id_scheme && p.PU_ID == id_pu 
                             select p).Single();

                return Json(new
                {    
                   
                    Success = true,
                    x_SCHEME_NAME = query.SCHEME_NAME,                                       
                    x_YEAR_INSTALL = query.YEAR_INSTALL,
                    x_SCH_TYPE = query.SCH_TYPE,
                    x_SCH_NO = query.SCH_NO,
                    x_EXC_ABB = query.EXC_ABB,
                    x_PU_ID = query.PU_ID,
                    x_BQ_MAT_PRICE = query.BQ_MAT_PRICE,
                    x_BQ_INSTALL_PRICE = query.BQ_INSTALL_PRICE,
                    x_PU_QTY = query.PU_QTY,                  
                    x_RATE_INDICATOR = query.RATE_INDICATOR,
                }, JsonRequestBehavior.AllowGet); //
            }
        }

        public ActionResult UpdateData_BOQ_MAIN(string txtEXC_ABB, string txtYEAR_INSTALL, string txtSCH_TYPE, string txtSCH_NO, string txtSCHEME_NAME, string txtPU_ID, string txtBQ_MAT_PRICE, string txtBQ_INSTALL_PRICE, string txtPU_QTY, string txtRATE_INDICATOR)
        {           
            bool success = true;                   

            WebView.WebService._base myWebService;
            myWebService = new WebService._base();

            WebService._base.BOQ_MAIN BOQ_MAIN_newjob = new WebService._base.BOQ_MAIN();
            BOQ_MAIN_newjob.x_EXC_ABB = txtEXC_ABB;          
            BOQ_MAIN_newjob.x_YEAR_INSTALL = txtYEAR_INSTALL;
            BOQ_MAIN_newjob.x_SCH_TYPE = txtSCH_TYPE;
            BOQ_MAIN_newjob.x_SCH_NO = txtSCH_NO;
            BOQ_MAIN_newjob.x_SCHEME_NAME = txtSCHEME_NAME;
            BOQ_MAIN_newjob.x_PU_ID = txtPU_ID;
            BOQ_MAIN_newjob.x_BQ_MAT_PRICE = txtBQ_MAT_PRICE;
            BOQ_MAIN_newjob.x_BQ_INSTALL_PRICE = txtBQ_INSTALL_PRICE;
            BOQ_MAIN_newjob.x_PU_QTY = txtPU_QTY;           
            BOQ_MAIN_newjob.x_RATE_INDICATOR = txtRATE_INDICATOR;
            success = myWebService.UpdateBOQ_MAIN(BOQ_MAIN_newjob, txtSCHEME_NAME);

            return Json(new
            {
               Success = success
            }, JsonRequestBehavior.AllowGet);
        }

        [HttpPost]
        public ActionResult DeleteData_BOQ_MAIN(string targetBOQ_MAIN)
        {
            bool success;

            WebView.WebService._base myWebService;
            myWebService = new WebService._base();

            success = myWebService.Delete_BOQ_MAIN(targetBOQ_MAIN);

            return Json(new
            {
                Success = success
            }, JsonRequestBehavior.AllowGet);
        }
    }
}

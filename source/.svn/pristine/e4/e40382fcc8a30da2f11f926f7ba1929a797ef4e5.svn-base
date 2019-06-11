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

    public class PlanUnitController : Controller
    {
       
        // GET: /PlanUnit/

        public ActionResult Index()
        {
            return View();
        }

        public ActionResult PlanUnit_List(string searchKey, string billrate, string searchUOM, int? page)
        {
            WebView.WebService._base myWebService;
            myWebService = new WebService._base();
            WebService._base.OSPPlanUnit PlanUnit = new WebService._base.OSPPlanUnit();
           
            //if (billrate == null)
            //{
            //    billrate = "Select";
            //}
            //if (UOM == null)
            //{
            //    UOM = "Select";
            //}
            if (searchKey != null || billrate != "Select" || searchUOM != null){
                if (searchKey == "" && billrate == "Select" && searchUOM.Equals("Select"))
                {
                    System.Diagnostics.Debug.WriteLine("Controller 1 " + "  Search key : " + searchKey + " | search UOM : " + searchUOM + " search bill : " + billrate);
                  
                    PlanUnit = myWebService.GetOSPPlanUnit(0, 100000, null, null, null);
                    ViewBag.searchKey = "";
                    ViewBag.billrate2 = "Select";
                    ViewBag.UOM2 = "Select";
                
                }else
                {
                    System.Diagnostics.Debug.WriteLine("Controller 2 " + "  Search key : " + searchKey + " | search UOM : " + searchUOM + " search bill : " + billrate);
   
                    PlanUnit = myWebService.GetOSPPlanUnit(0, 100000, searchKey, billrate, searchUOM);
                    ViewBag.searchKey = searchKey;
                    ViewBag.billrate2 = billrate;
                    ViewBag.UOM2 = searchUOM;
                }
            }
           
            else
            {
                System.Diagnostics.Debug.WriteLine("Controller 3 " + "  Search key : " + searchKey + " | search UOM : " + searchUOM + " search bill : " + billrate );
                PlanUnit = myWebService.GetOSPPlanUnit(0, 100000, null, null, null);
                ViewBag.searchKey = "";
                ViewBag.billrate2 = "Select";
                ViewBag.searchUOM = "Select";
            }

            ViewData["data7"] = PlanUnit.PlanUnitList;

            string input = "\\\\adsvr";
            //string input = "\\\\server\\d$\\x\\y\\z\\AAA";
            string output = String.Format("http:{0}", input.Replace("\\d$\\x\\y", String.Empty).Replace("\\", "/"));

            ViewBag.output = output;

            List<SelectListItem> list = new List<SelectListItem>();
            list.Add(new SelectListItem() { Text = "", Value = "Select" });
            list.Add(new SelectListItem() { Text = "D", Value = "D" });
            list.Add(new SelectListItem() { Text = "N", Value = "N" });
            list.Add(new SelectListItem() { Text = "P", Value = "P" });
            list.Add(new SelectListItem() { Text = "W", Value = "W" });

            ViewBag.billrate = list;

            using (Entities ctxData = new Entities())
            {
                var query = (from q in ctxData.WV_PU_MAST
                             select new { Text = q.PU_UOM, Value = q.PU_UOM });

                List<SelectListItem> list2 = new List<SelectListItem>();
                list2.Add(new SelectListItem() { Text = "", Value = "Select" });
                foreach (var a in query.Distinct().OrderBy(it => it.Value))
                {
                    list2.Add(new SelectListItem() { Text = a.Text , Value = a.Value });
                }
                ViewBag.UOM = list2;
            }
            //return View();
            int pageSize =10;
            int pageNumber = (page ?? 1);
            return View(PlanUnit.PlanUnitList.ToPagedList(pageNumber, pageSize));                 
        }       

        public ActionResult PlanUnit_NewJob()
        {
            using (Entities ctxData = new Entities())
            {
                var query = (from q in ctxData.WV_PU_MAST
                             select new { Text = q.PU_UOM, Value = q.PU_UOM });

                List<SelectListItem> list2 = new List<SelectListItem>();
                list2.Add(new SelectListItem() { Text = "", Value = "Select" });

                foreach (var a in query.Distinct().OrderBy(it => it.Value))
                {
                    list2.Add(new SelectListItem() { Text = a.Text, Value = a.Value });
                }
                ViewBag.UOM = list2;
            }

            return View();
        }

        [HttpPost]
        public ActionResult PlanUnit_NewJob(PlanUnit_NewJobModel m)
        {
            bool success = true;
            WebView.WebService._base myWebService;
            myWebService = new WebService._base();

            bool selected = false;

                WebService._base.PlanUnit PlanUnitNewjob = new WebService._base.PlanUnit();
                PlanUnitNewjob.PU_ID = m.x_PU_ID;
                PlanUnitNewjob.PU_DESC = m.x_PU_DESC;
                PlanUnitNewjob.PU_INST_PR = m.x_PU_INST_PR.ToString();
                PlanUnitNewjob.PU_MAT_PR = m.x_PU_MAT_PR.ToString();
                System.Diagnostics.Debug.WriteLine("Master {} :" + m.x_PU_UOM.ToString()); 
                PlanUnitNewjob.PU_UOM = m.x_PU_UOM;
                PlanUnitNewjob.PU_DURATION = m.x_PU_DURATION.ToString();
                PlanUnitNewjob.PU_NETWORK_FLAG = m.x_PU_NETWORK_FLAG;
                PlanUnitNewjob.PU_ACT_CODE = m.x_PU_ACT_CODE;
                PlanUnitNewjob.PU_SUP_DURATION = m.x_PU_SUP_DURATION.ToString();
                PlanUnitNewjob.PU_BILL_RATE = m.x_PU_BILL_RATE;
                success = myWebService.AddPlanUnit(PlanUnitNewjob);

                selected = true;
                
                if(success == true) 
                    return RedirectToAction("PlanUnit_NewSave");
                else
                    return RedirectToAction("NewSaveFail"); // store to db failed.
                   

            //return View(m);
        }    

        public ActionResult PlanUnit_NewSave()
        {
            return View();
        }

        public ActionResult GetDetails_PlanUnit(string id, string code)
        {
            string fileListStr = "";
            string dataListUOM = "";

            using (Entities ctxData = new Entities())
            {
                var query = (from p in ctxData.WV_PU_MAST
                             orderby p.PU_UOM descending
                             where p.PU_ID == id && p.BILL_RATE == code
                             select p).Single();

                var QueryUOM = (from d in ctxData.WV_PU_MAST
                                select new { Value = d.PU_UOM });

                foreach (var i in QueryUOM.Distinct().OrderBy(it => it.Value))
                {
                    dataListUOM = dataListUOM + i.Value + ":";
                }
                return Json(new
                {
                    Success = true,
                    Puid = id,
                    Pudesc = query.PU_DESC,
                    PUInstPR = query.PU_INST_PR,
                    PUMatPR = query.PU_MAT_PR,
                    PUUom = query.PU_UOM,
                    PUDuration = query.DURATION,
                    PUNetworkFlag = query.NETWORK_FLAG,
                    PUActCode = query.ACT_CODE,
                    PUSupDuration = query.SUP_DURATION,
                    PUBillRate = query.BILL_RATE,
                    ListUOM = dataListUOM,

                    files_list = new
                    {
                        files = fileListStr
                    }
                }, JsonRequestBehavior.AllowGet); 
            }
        }

        public ActionResult UpdateData_PlanUnit(string txtPU_ID1, string txtPU_DESC, string txtPU_INST_PR, string txtPU_MAT_PR, string txtPU_UOM, string txtDURATION, string txtNETWORK_FLAG, string txtACT_CODE, string txtSUP_DURATION, string txtPUBillRate)
        {           
            bool success = true;                   

            WebView.WebService._base myWebService;
            myWebService = new WebService._base();
            System.Diagnostics.Debug.WriteLine("txtPU_ID AKU :" + txtPU_DESC);
            WebService._base.PlanUnit planUnitnewjob = new WebService._base.PlanUnit();
            planUnitnewjob.PU_ID = txtPU_ID1;
            planUnitnewjob.PU_DESC = txtPU_DESC;
            planUnitnewjob.PU_INST_PR = txtPU_INST_PR;
            planUnitnewjob.PU_MAT_PR = txtPU_MAT_PR;
            planUnitnewjob.PU_UOM = txtPU_UOM;
            planUnitnewjob.PU_DURATION = txtDURATION;
            planUnitnewjob.PU_NETWORK_FLAG = txtNETWORK_FLAG;
            planUnitnewjob.PU_ACT_CODE = txtACT_CODE;
            planUnitnewjob.PU_SUP_DURATION = txtSUP_DURATION;
            planUnitnewjob.PU_BILL_RATE = txtPUBillRate;
            success = myWebService.UpdatePlanUnit(planUnitnewjob, txtPU_ID1, txtPUBillRate);

            return Json(new
            {
               Success = success
            }, JsonRequestBehavior.AllowGet);
        }

        [HttpPost]
        public ActionResult DeleteData_PlanUnit(string targetPlanUnit, string targetBillRate)
        {
            bool success;

            WebView.WebService._base myWebService;
            myWebService = new WebService._base();
            System.Diagnostics.Debug.WriteLine("RES :" + targetPlanUnit);

            success = myWebService.Delete_PlanUnit(targetPlanUnit, targetBillRate);

            return Json(new
            {
                Success = success
            }, JsonRequestBehavior.AllowGet);
        }

        public string UOM { get; set; }

        public string searchUOM { get; set; }
    }
}

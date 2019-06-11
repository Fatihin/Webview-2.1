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

    public class CSMController : Controller
    {
       
        // GET: /CSM/

        public ActionResult Index()
        {
            return View();
        }

        public ActionResult CSM_List(string searchKey, int? page, string itemNo)
        {
            WebView.WebService._base myWebService;
            myWebService = new WebService._base();

            WebService._base.OSPCSM CSM = new WebService._base.OSPCSM();

            if (searchKey != null)
            {
                    CSM = myWebService.GetOSPCSM(0, 100000, searchKey, itemNo);
                    ViewBag.searchKey = searchKey;
                    ViewBag.itemNo2 = itemNo;
            }
            else
            {
                CSM = myWebService.GetOSPCSM(0, 100000, null, null);
                ViewBag.searchKey = "";
                ViewBag.itemNo2 = "Select";
            }

            ViewData["data7"] = CSM.CSMList;

            string input = "\\\\adsvr";
            //string input = "\\\\server\\d$\\x\\y\\z\\AAA";
            string output = String.Format("http:{0}", input.Replace("\\d$\\x\\y", String.Empty).Replace("\\", "/"));

            ViewBag.output = output;

            List<SelectListItem> list = new List<SelectListItem>();
            using (Entities ctxData = new Entities())
            {
                var queryItemNo = from p in ctxData.WV_CONTRACT_MAST
                                  select new { Text = p.ITEM_NO, Value = p.ITEM_NO };

                list.Add(new SelectListItem() { Text = "", Value = "Select" });
                foreach (var a in queryItemNo.Distinct().OrderBy(it => it.Value))
                {
                    if (a.Value != null)
                        list.Add(new SelectListItem() { Text = a.Text.ToString(), Value = a.Value.ToString() });
                }
            }
            ViewBag.itemNo = list;

            //return View();
            int pageSize = 10;
            int pageNumber = (page ?? 1);
            return View(CSM.CSMList.ToPagedList(pageNumber, pageSize));
        }
    

        public ActionResult CSM_NewJob()
        {           
            return View();
        }

        [HttpPost]
        public ActionResult CSM_NewJob(CSM_NewJobModel model)
        {
            bool success = true;
            WebView.WebService._base myWebService;
            myWebService = new WebService._base();

            bool selected = false;

            WebService._base.CSM CSM_newjob = new WebService._base.CSM();
            CSM_newjob.CONT_ID = model.x_CONT_ID;
            CSM_newjob.CONT_DESC = model.x_CONT_DESC;
            CSM_newjob.CONTRACT_ID = model.x_CONTRACT_ID;
            CSM_newjob.CONTRACT_NO = model.x_CONTRACT_NO;
            CSM_newjob.ACT_CODE = model.x_ACT_CODE;
            CSM_newjob.PU_UOM = model.x_PU_UOM;
            CSM_newjob.CONTRACT_PR = model.x_CONTRACT_PR;
            success = myWebService.AddCSM(CSM_newjob);

            selected = true;


            if (ModelState.IsValid && selected)
            {
                if (success == true)
                    return RedirectToAction("CSM_NewSave");
                else
                    return RedirectToAction("NewSaveFail"); // store to db failed.
            }

            return View(model);
        }

        public ActionResult CSM_NewSave()
        {
            return View();
        }

        public ActionResult GetDetails_CSM(string id, string code)
        {
            using (Entities ctxData = new Entities())
            {
                int itemNo_con = Convert.ToInt32(code);
                var query = (from p in ctxData.WV_CONTRACT_MAST
                             where p.CONTRACT_NO == id && p.ITEM_NO == itemNo_con
                             select p).Single();

                return Json(new
                {    
                   
                    Success = true,
                    x_CONT_ID = id,
                    x_CONT_DESC = query.ITEM_NO,
                    x_CONTRACT_ID = query.CONTRACT_DESC,
                    x_CONTRACT_NO = query.NET_PRICE,
                    x_PU_UOM = query.ORDER_UNIT,
                    x_CONTRACT_PR = query.ITEM_CATEGORY,
                    x_ACT_CODE = query.NETWORK_FLAG,                    
                }, JsonRequestBehavior.AllowGet); //
            }
        }

        // Todo:Save update data
        public ActionResult UpdateData_CSM(string txtCONT_ID, string txtCONT_DESC, string txtCONTRACT_ID, string txtCONTRACT_NO, string txtACT_CODE, string txtPU_UOM, string txtCONTRACT_PR)
        {
            bool success = true;

            WebView.WebService._base myWebService;
            myWebService = new WebService._base();

            WebService._base.CSM CSM_newjob = new WebService._base.CSM();
            CSM_newjob.CONT_ID = txtCONT_ID;
            CSM_newjob.CONT_DESC = txtCONT_DESC;
            CSM_newjob.CONTRACT_ID = txtCONTRACT_ID;
            CSM_newjob.CONTRACT_NO = txtCONTRACT_NO;
            CSM_newjob.ACT_CODE = txtACT_CODE;
            CSM_newjob.PU_UOM = txtPU_UOM;
            CSM_newjob.CONTRACT_PR = txtCONTRACT_PR;

            success = myWebService.UpdateCSM(CSM_newjob, txtCONT_ID, txtCONT_DESC);
            System.Diagnostics.Debug.WriteLine(success);
            return Json(new
            {
                Success = success
            }, JsonRequestBehavior.AllowGet);
        }


        [HttpPost]
        public ActionResult DeleteData_CSM(string targetID, string targetItem)
        {
            bool success;

            WebView.WebService._base myWebService;
            myWebService = new WebService._base();
            System.Diagnostics.Debug.WriteLine(targetID);
            System.Diagnostics.Debug.WriteLine(targetItem);
            success = myWebService.Delete_CSM(targetID, targetItem);
            System.Diagnostics.Debug.WriteLine(success);
            return Json(new
            {
                Success = success
            }, JsonRequestBehavior.AllowGet);
        }
    }
}
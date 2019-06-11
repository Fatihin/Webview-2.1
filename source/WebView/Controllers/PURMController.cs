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

    public class PURMController : Controller
    {
       
        // GET: /PURM/

        public ActionResult Index()
        {
            return View();
        }

        public ActionResult PURM_List(string searchKey, int? page)
        {
            WebView.WebService._base myWebService;
            myWebService = new WebService._base();

            WebService._base.OSPPURM PURM = new WebService._base.OSPPURM();

            if (searchKey != null)
            {
                if (searchKey.Equals(""))
                    PURM = myWebService.GetOSPPURM(0, 1000, null);
                else
                {
                    PURM = myWebService.GetOSPPURM(0, 1000, searchKey);
                    ViewBag.searchKey = searchKey;
                }
            }
            else
            {
                PURM = myWebService.GetOSPPURM(0, 1000, null);
                ViewBag.searchKey = "";
            }

            ViewData["data7"] = PURM.PURMList;

            string input = "\\\\adsvr";
            //string input = "\\\\server\\d$\\x\\y\\z\\AAA";
            string output = String.Format("http:{0}", input.Replace("\\d$\\x\\y", String.Empty).Replace("\\", "/"));

            ViewBag.output = output;

            //return View();
            int pageSize = 5;
            int pageNumber = (page ?? 1);
            return View(PURM.PURMList.ToPagedList(pageNumber, pageSize));
        }
    

        public ActionResult PURM_NewJob()
        {           
            return View();
        }

        [HttpPost]
        public ActionResult PURM_NewJob(PURM_NewJobModel model)
        {
            bool success = true;
            WebView.WebService._base myWebService;
            myWebService = new WebService._base();

            bool selected = false;

            // create job in OSP (SOAP)
           
                WebService._base.PURM PURM_newjob = new WebService._base.PURM();              
                PURM_newjob.PU_ID = model.x_PU_ID;
                PURM_newjob.PU_DESC = model.x_PU_DESC ;             
                PURM_newjob.INST_CODE = model.x_INST_CODE;
                PURM_newjob.INST_NAME = model.x_INST_NAME;
                PURM_newjob.CTRT_SUP_QTY = model.x_CTRT_SUP_QTY;
                PURM_newjob.SUB_LAB_QTY = model.x_SUB_LAB_QTY;
                PURM_newjob.IMPL_LAB_QTY = model.x_IMPL_LAB_QTY;
                PURM_newjob.DUR_SPVR = model.x_DUR_SPVR;
                PURM_newjob.DUR_IMPL = model.x_DUR_IMPL; 
                success = myWebService.AddPURM (PURM_newjob);

                selected = true;
           

            if (ModelState.IsValid && selected)
            {
                if(success == true) 
                    return RedirectToAction("PURM_NewSave");
                else
                    return RedirectToAction("NewSaveFail"); // store to db failed.
            }           

            return View(model);
        }    

        public ActionResult PURM_NewSave()
        {
            return View();
        }

        public ActionResult GetDetails_PURM(string id)
        {
            using (Entities ctxData = new Entities())
            {
                var query = (from p in ctxData.WV_PURM_MAST
                             where p.PU_ID == id
                             select p).Single();

                return Json(new
                {    
                   
                    Success = true,
                    x_PU_ID = id,                    
                    x_PU_DESC = query.PU_DESC,
                    x_INST_CODE = query.INST_CODE,
                    x_INST_NAME = query.INST_NAME,
                    x_CTRT_SUP_QTY = query.CTRT_SUP_QTY,
                    x_SUB_LAB_QTY = query.SUB_LAB_QTY,
                    x_IMPL_LAB_QTY = query.IMPL_LAB_QTY,
                    x_DUR_SPVR = query.DUR_SPVR,
                    x_DUR_IMPL = query.DUR_IMPL,
                }, JsonRequestBehavior.AllowGet); //
            }
        }

        public ActionResult UpdateData_PURM(string txtPU_ID, string txtPU_DESC, string txtINST_CODE, string txtINST_NAME, string txtCTRT_SUP_QTY, string txtSUB_LAB_QTY, string txtIMPL_LAB_QTY, string txtDUR_SPVR, string txtDUR_IMPL)
        {
            bool success = true;

            WebView.WebService._base myWebService;
            myWebService = new WebService._base();

            WebService._base.PURM PURM_newjob = new WebService._base.PURM();
            PURM_newjob.PU_ID = txtPU_ID;
            PURM_newjob.PU_DESC = txtPU_DESC;
            PURM_newjob.INST_CODE = txtINST_CODE;
            PURM_newjob.INST_NAME = txtINST_NAME;
            PURM_newjob.CTRT_SUP_QTY = txtCTRT_SUP_QTY;
            PURM_newjob.SUB_LAB_QTY = txtSUB_LAB_QTY;
            PURM_newjob.IMPL_LAB_QTY = txtIMPL_LAB_QTY;
            PURM_newjob.DUR_SPVR = txtDUR_SPVR;
            PURM_newjob.DUR_IMPL = txtDUR_IMPL;
            success = myWebService.UpdatePURM(PURM_newjob, txtPU_ID);

            return Json(new
            {
                Success = success
            }, JsonRequestBehavior.AllowGet);
        }

        [HttpPost]
        public ActionResult DeleteData_PURM(string targetPURM)
        {
            bool success;

            WebView.WebService._base myWebService;
            myWebService = new WebService._base();

            success = myWebService.Delete_PURM(targetPURM);

            return Json(new
            {
                Success = success
            }, JsonRequestBehavior.AllowGet);
        }
    }
}

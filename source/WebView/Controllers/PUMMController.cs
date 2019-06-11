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

    public class PUMMController : Controller
    {
       
        // GET: /PUMM/

        public ActionResult Index()
        {
            return View();
        }

        public ActionResult PUMM_List(string searchKey, int? page)
        {
            WebView.WebService._base myWebService;
            myWebService = new WebService._base();

            WebService._base.OSPPUMM PUMM = new WebService._base.OSPPUMM();

            if (searchKey != null)
            {
                if (searchKey.Equals(""))
                    PUMM = myWebService.GetOSPPUMM(0, 1000, null);
                else
                {
                    PUMM = myWebService.GetOSPPUMM(0, 1000, searchKey);
                    ViewBag.searchKey = searchKey;
                }
            }
            else
            {
                PUMM = myWebService.GetOSPPUMM(0, 1000, null);
                ViewBag.searchKey = "";
            }
            
            ViewData["data7"] = PUMM.PUMMList;

            string input = "\\\\adsvr";
            //string input = "\\\\server\\d$\\x\\y\\z\\AAA";
            string output = String.Format("http:{0}", input.Replace("\\d$\\x\\y", String.Empty).Replace("\\", "/"));

            ViewBag.output = output;

            //return View();
            int pageSize = 5;
            int pageNumber = (page ?? 1);
            return View(PUMM.PUMMList.ToPagedList(pageNumber, pageSize));
        }
    
        public ActionResult PUMM_NewJob()
        {           
            return View();
        }

        [HttpPost]
        public ActionResult PUMM_NewJob(PUMM_NewJobModel model)
        {
            bool success = true;
            WebView.WebService._base myWebService;
            myWebService = new WebService._base();

            bool selected = false;          
           
                WebService._base.PUMM PUMM_newjob = new WebService._base.PUMM();              
                PUMM_newjob.PU_ID = model.x_PU_ID;
                PUMM_newjob.PU_DESC = model.x_PU_DESC ;             
                PUMM_newjob.MAT_ID = model.x_MAT_ID;
                PUMM_newjob.MAT_NAME = model.x_MAT_NAME;
                PUMM_newjob.MAT_QTY = model.x_MAT_QTY;             
                success = myWebService.AddPUMM (PUMM_newjob);

                selected = true;
           
            if (ModelState.IsValid && selected)
            {
                if(success == true) 
                    return RedirectToAction("PUMM_NewSave");
                else
                    return RedirectToAction("NewSaveFail"); // store to db failed.
            }           

            return View(model);
        }    

        public ActionResult PUMM_NewSave()
        {
            return View();
        }

        public ActionResult GetDetails_PUMM(string id)
        {
            using (Entities ctxData = new Entities())
            {
                var query = (from p in ctxData.WV_PUMM_MAST
                             where p.PU_ID == id
                             select p).Single();

                return Json(new
                {    
                   
                    Success = true,
                    x_PU_ID = id,                    
                    x_PU_DESC = query.PU_DESC,
                    x_MAT_ID = query.MAT_ID,
                    x_MAT_NAME = query.MAT_NAME,
                    x_MAT_QTY = query.MAT_QTY,
                
                }, JsonRequestBehavior.AllowGet); //
            }
        }

        public ActionResult UpdateData_PUMM(string txtPU_ID, string txtPU_DESC, string txtMAT_ID, string txtMAT_NAME, string txtMAT_QTY)
        {
            bool success = true;

            WebView.WebService._base myWebService;
            myWebService = new WebService._base();

            WebService._base.PUMM PUMM_newjob = new WebService._base.PUMM();
            PUMM_newjob.PU_ID = txtPU_ID;
            PUMM_newjob.PU_DESC = txtPU_DESC;
            PUMM_newjob.MAT_ID = txtMAT_ID;
            PUMM_newjob.MAT_NAME = txtMAT_NAME;
            PUMM_newjob.MAT_QTY = txtMAT_QTY;            
            success = myWebService.UpdatePUMM(PUMM_newjob, txtPU_ID);

            return Json(new
            {
                Success = success
            }, JsonRequestBehavior.AllowGet);
        }

        [HttpPost]
        public ActionResult DeleteData_PUMM(string targetPUMM)
        {
            bool success;

            WebView.WebService._base myWebService;
            myWebService = new WebService._base();

            success = myWebService.Delete_PUMM(targetPUMM);

            return Json(new
            {
                Success = success
            }, JsonRequestBehavior.AllowGet);
        }
    }
}

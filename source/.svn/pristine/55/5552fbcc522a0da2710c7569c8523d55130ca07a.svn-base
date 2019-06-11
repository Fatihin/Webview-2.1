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

    public class CSFRMController : Controller
    {
       
        // GET: /CSFRM/

        public ActionResult Index()
        {
            return View();
        }

        public ActionResult CSFRM_List(string searchKey, int? page)
        {
            WebView.WebService._base myWebService;
            myWebService = new WebService._base();

            WebService._base.OSPCSFRM CSFRM = new WebService._base.OSPCSFRM();

            if (searchKey != null)
            {
                if (searchKey.Equals(""))
                    CSFRM = myWebService.GetOSPCSFRM(0, 1000, null);
                else
                {
                    CSFRM = myWebService.GetOSPCSFRM(0, 1000, searchKey);
                    ViewBag.searchKey = searchKey;
                }
            }
            else
            {
                CSFRM = myWebService.GetOSPCSFRM(0, 1000, null);
                ViewBag.searchKey = "";
            }

            ViewData["data7"] = CSFRM.CSFRMList;

            string input = "\\\\adsvr";
            //string input = "\\\\server\\d$\\x\\y\\z\\AAA";
            string output = String.Format("http:{0}", input.Replace("\\d$\\x\\y", String.Empty).Replace("\\", "/"));

            ViewBag.output = output;

            //return View();
            int pageSize = 5;
            int pageNumber = (page ?? 1);
            return View(CSFRM.CSFRMList.ToPagedList(pageNumber, pageSize));
        }
    

        public ActionResult CSFRM_NewJob()
        {           
            return View();
        }

        [HttpPost]
        public ActionResult CSFRM_NewJob(CSFRM_NewJobModel model)
        {
            bool success = true;
            WebView.WebService._base myWebService;
            myWebService = new WebService._base();

            bool selected = false;

            // create job in OSP (SOAP)
           
                WebService._base.CSFRM CSFRM_newjob = new WebService._base.CSFRM();              
                CSFRM_newjob.CONT_ID = model.x_CONT_ID;
                CSFRM_newjob.CONT_DESC = model.x_CONT_DESC ;             
                CSFRM_newjob.FEAT_TYPE = model.x_FEAT_TYPE;
                CSFRM_newjob.QTY_IND = model.x_QTY_IND;
                CSFRM_newjob.MUL_FAC = model.x_MUL_FAC;
                CSFRM_newjob.ATT1 = model.x_ATT1;
                CSFRM_newjob.ATT2 = model.x_ATT2;
                CSFRM_newjob.ATT3 = model.x_ATT3;
                CSFRM_newjob.ATT4 = model.x_ATT4;
                CSFRM_newjob.ATT5 = model.x_ATT5;
                CSFRM_newjob.ATT6 = model.x_ATT6;
                CSFRM_newjob.ATT7 = model.x_ATT7;
                CSFRM_newjob.ATT8 = model.x_ATT8;
                CSFRM_newjob.ATT9 = model.x_ATT9;
                CSFRM_newjob.ATT10 = model.x_ATT10;
                CSFRM_newjob.ATT11 = model.x_ATT11;
                CSFRM_newjob.ATT12 = model.x_ATT12;
                success = myWebService.AddCSFRM (CSFRM_newjob);

                selected = true;
           

            if (ModelState.IsValid && selected)
            {
                if(success == true) 
                    return RedirectToAction("CSFRM_NewSave");
                else
                    return RedirectToAction("NewSaveFail"); // store to db failed.
            }           

            return View(model);
        }    

        public ActionResult CSFRM_NewSave()
        {
            return View();
        }

        public ActionResult GetDetails_CSFRM(string id)
        {
            using (Entities ctxData = new Entities())
            {
                var query = (from p in ctxData.WV_CSFRM_MAST
                             where p.CONT_ID == id
                             select p).Single();

                return Json(new
                {    
                   
                    Success = true,
                    CONT_ID = id,                    
                    CONT_DESC = query.CONT_DESC,
                    FEAT_TYPE = query.FEAT_TYPE,
                    QTY_IND = query.QTY_IND,
                    MUL_FAC = query.MUL_FAC,
                    ATT1 = query.ATT1,
                    ATT2 = query.ATT2,
                    ATT3 = query.ATT3,
                    ATT4 = query.ATT4,
                    ATT5 = query.ATT5,
                    ATT6 = query.ATT6,
                    ATT7 = query.ATT7,
                    ATT8 = query.ATT8,
                    ATT9 = query.ATT9,
                    ATT10 = query.ATT10,
                    ATT11 = query.ATT11,
                    ATT12 = query.ATT12,    
                }, JsonRequestBehavior.AllowGet); //
            }
        }

        // Todo:Save update data
        public ActionResult UpdateData_CSFRM(string txtCONT_ID, string txtCONT_DESC, string txtFEAT_TYPE, string txtQTY_IND, string txtMUL_FAC, string txtATT1, string txtATT2, string txtATT3, string txtATT4, string txtATT5, string txtATT6, string txtATT7, string txtATT8, string txtATT9, string txtATT10, string txtATT11, string txtATT12)
        {
            bool success = true;

            WebView.WebService._base myWebService;
            myWebService = new WebService._base();

            WebService._base.CSFRM CSFRMnewjob = new WebService._base.CSFRM();
            CSFRMnewjob.CONT_ID = txtCONT_ID;
            CSFRMnewjob.CONT_DESC = txtCONT_DESC;
            CSFRMnewjob.FEAT_TYPE = txtFEAT_TYPE;
            CSFRMnewjob.QTY_IND = txtQTY_IND;
            CSFRMnewjob.MUL_FAC = txtMUL_FAC;
            CSFRMnewjob.ATT1 = txtATT1;
            CSFRMnewjob.ATT2 = txtATT2;
            CSFRMnewjob.ATT3 = txtATT3;
            CSFRMnewjob.ATT4 = txtATT4;
            CSFRMnewjob.ATT5 = txtATT5;
            CSFRMnewjob.ATT6 = txtATT6;
            CSFRMnewjob.ATT7 = txtATT7;
            CSFRMnewjob.ATT8 = txtATT8;
            CSFRMnewjob.ATT9 = txtATT9;
            CSFRMnewjob.ATT10 = txtATT10;
            CSFRMnewjob.ATT11 = txtATT11;
            CSFRMnewjob.ATT12 = txtATT12;
            success = myWebService.UpdateCSFRM(CSFRMnewjob, txtCONT_ID);

            return Json(new
            {
                Success = success
            }, JsonRequestBehavior.AllowGet);
        }

        [HttpPost]
        public ActionResult DeleteData_PUFRM(string targetPUFRM)
        {
            bool success;

            WebView.WebService._base myWebService;
            myWebService = new WebService._base();

            success = myWebService.Delete_CSFRM(targetPUFRM);

            return Json(new
            {
                Success = success
            }, JsonRequestBehavior.AllowGet);
        }
    
    }
}

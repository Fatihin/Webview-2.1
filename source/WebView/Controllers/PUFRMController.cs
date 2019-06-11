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

    public class PUFRMController : Controller
    {
       
        // GET: /PUFRM/

        public ActionResult Index()
        {
            return View();
        }

        public ActionResult PUFRM_List(string searchKey, int? page)
        {
            WebView.WebService._base myWebService;
            myWebService = new WebService._base();

            WebService._base.OSPPUFRM PUFRM = new WebService._base.OSPPUFRM();

            if (searchKey != null)
            {
                if (searchKey.Equals(""))
                    PUFRM = myWebService.GetOSPPUFRM(0, 1000, null);
                else
                {
                    PUFRM = myWebService.GetOSPPUFRM(0, 1000, searchKey);
                    ViewBag.searchKey = searchKey;
                }
            }
            else
            {
                PUFRM = myWebService.GetOSPPUFRM(0, 1000, null);
                ViewBag.searchKey = "";
            }

            ViewData["data7"] = PUFRM.PUFRMList;

            string input = "\\\\adsvr";
            //string input = "\\\\server\\d$\\x\\y\\z\\AAA";
            string output = String.Format("http:{0}", input.Replace("\\d$\\x\\y", String.Empty).Replace("\\", "/"));

            ViewBag.output = output;

            //return View();
            int pageSize = 5;
            int pageNumber = (page ?? 1);
            return View(PUFRM.PUFRMList.ToPagedList(pageNumber, pageSize));
        }
    
        public ActionResult PUFRM_NewJob()
        {
            List<SelectListItem> list = new List<SelectListItem>();
            using (Entities ctxData = new Entities())
            {
                var query = from p in ctxData.WV_PU_MAST
                            select new { Text = p.PU_ID, Value = p.PU_ID};

                foreach (var a in query)
                {
                    list.Add(new SelectListItem() { Text = a.Text, Value = a.Value.ToString() });
                }

                ViewBag.PUId = list;
            }
            return View();
        }

        [HttpPost]
        public ActionResult PUFRM_NewJob(PUFRM_NewJobModel model)
        {
            bool success = true;
            WebView.WebService._base myWebService;
            myWebService = new WebService._base();

            bool selected = false;
            if (model.x_QTY_IND.Trim().Length > 1)
            {
                ModelState.AddModelError("x_QTY_IND", "x_QTY_IND must below 10.");
            }
            else
            {
                // create job in OSP (SOAP)

                WebService._base.PUFRM PUFRM_newjob = new WebService._base.PUFRM();
                PUFRM_newjob.PU_ID = model.x_PU_ID;
                PUFRM_newjob.PU_DESC = model.x_PU_DESC;
                PUFRM_newjob.FEAT_TYPE = model.x_FEAT_TYPE;
                PUFRM_newjob.QTY_IND = model.x_QTY_IND;
                PUFRM_newjob.MUL_FAC = model.x_MUL_FAC;
                PUFRM_newjob.ATT1 = model.x_ATT1;
                PUFRM_newjob.ATT2 = model.x_ATT2;
                PUFRM_newjob.ATT3 = model.x_ATT3;
                PUFRM_newjob.ATT4 = model.x_ATT4;
                PUFRM_newjob.ATT5 = model.x_ATT5;
                PUFRM_newjob.ATT6 = model.x_ATT6;
                PUFRM_newjob.ATT7 = model.x_ATT7;
                PUFRM_newjob.ATT8 = model.x_ATT8;
                PUFRM_newjob.ATT9 = model.x_ATT9;
                PUFRM_newjob.ATT10 = model.x_ATT10;
                PUFRM_newjob.ATT11 = model.x_ATT11;
                PUFRM_newjob.ATT12 = model.x_ATT12;
                success = myWebService.AddPUFRM(PUFRM_newjob);

                selected = true;
            }

            if (ModelState.IsValid && selected)
            {
                if(success == true) 
                    return RedirectToAction("PUFRM_NewSave");
                else
                    return RedirectToAction("NewSaveFail"); // store to db failed.
            }           

            return View(model);
        }    

        public ActionResult PUFRM_NewSave()
        {
            return View();
        }

        public ActionResult GetDetails_PUFRM(string id)
        {
            using (Entities ctxData = new Entities())
            {
                var query = (from p in ctxData.WV_PUFRM_MAST
                             where p.PU_ID == id
                             select p).Single();

                return Json(new
                {    
                   
                    Success = true,
                    x_PU_ID = id,                    
                    x_PU_DESC = query.PU_DESC,
                    x_FEAT_TYPE = query.FEAT_TYPE,
                    x_QTY_IND = query.QTY_IND,
                    x_MUL_FAC = query.MUL_FAC,
                    x_ATT1 = query.ATT1,
                    x_ATT2 = query.ATT2,
                    x_ATT3 = query.ATT3,
                    x_ATT4 = query.ATT4,
                    x_ATT5 = query.ATT5,
                    x_ATT6 = query.ATT6,
                    x_ATT7 = query.ATT7,
                    x_ATT8 = query.ATT8,
                    x_ATT9 = query.ATT9,
                    x_ATT10 = query.ATT10,
                    x_ATT11 = query.ATT11,
                    x_ATT12 = query.ATT12,    
                }, JsonRequestBehavior.AllowGet); //
            }
        }

        public ActionResult UpdateData_PUFRM(string txtPU_ID, string txtPU_DESC, string txtFEAT_TYPE, string txtQTY_IND, string txtMUL_FAC, string txtATT1, string txtATT2, string txtATT3, string txtATT4, string txtATT5, string txtATT6, string txtATT7, string txtATT8, string txtATT9, string txtATT10, string txtATT11, string txtATT12)
        {
            bool success = true;

            WebView.WebService._base myWebService;
            myWebService = new WebService._base();

            WebService._base.PUFRM PUFRMnewjob = new WebService._base.PUFRM();
            PUFRMnewjob.PU_ID = txtPU_ID;
            PUFRMnewjob.PU_DESC = txtPU_DESC;
            PUFRMnewjob.FEAT_TYPE = txtFEAT_TYPE;
            PUFRMnewjob.QTY_IND = txtQTY_IND;
            PUFRMnewjob.MUL_FAC = txtMUL_FAC;
            PUFRMnewjob.ATT1 = txtATT1;
            PUFRMnewjob.ATT2 = txtATT2;
            PUFRMnewjob.ATT3 = txtATT3;
            PUFRMnewjob.ATT4 = txtATT4;
            PUFRMnewjob.ATT5 = txtATT5;
            PUFRMnewjob.ATT6 = txtATT6;
            PUFRMnewjob.ATT7 = txtATT7;
            PUFRMnewjob.ATT8 = txtATT8;
            PUFRMnewjob.ATT9 = txtATT9;
            PUFRMnewjob.ATT10 = txtATT10;
            PUFRMnewjob.ATT11 = txtATT11;
            PUFRMnewjob.ATT12 = txtATT12;
            success = myWebService.UpdatePUFRM(PUFRMnewjob, txtPU_ID);

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

            success = myWebService.Delete_PUFRM(targetPUFRM);

            return Json(new
            {
                Success = success
            }, JsonRequestBehavior.AllowGet);
        }
    }
}

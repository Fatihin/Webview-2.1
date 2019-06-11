
using System;
using System.Linq;
using System.Web.Mvc;
using WebView.Models;

namespace WebView.Controllers
{

    public class BOQ_JKHController : Controller
    {

        // GET: /BOQ_JKH/

        public ActionResult Index()
        {
            return View();
        }

        public ActionResult BOQ_JKH_List()
        {
            WebView.WebService._base myWebService;
            myWebService = new WebService._base();

            WebService._base.OSPBOQ_JKH BOQ_JKH = new WebService._base.OSPBOQ_JKH();
            BOQ_JKH = myWebService.GetOSPBOQ_JKH(0, 10);

            ViewData["data2"] = BOQ_JKH.BOQ_JKHList;
          
            //string input = "\\\\adsvr";           
            //string output = String.Format("http:{0}", input.Replace("\\d$\\x\\y", String.Empty).Replace("\\", "/"));

            //ViewBag.output = output;

            return View();          
           
        }

        public ActionResult BOQ_JKH_NewJob()
        {           
            return View();
        }

        [HttpPost]
        public ActionResult BOQ_JKH_NewJob(BOQ_JKH_NewJobModel model)
        {
            bool success = true;
            WebView.WebService._base myWebService;
            myWebService = new WebService._base();

            bool selected = false;

            // create job in OSP (SOAP)

            WebService._base.BOQ_JKH BOQ_JKH_newjob = new WebService._base.BOQ_JKH();
            BOQ_JKH_newjob.x_EXC_ABB = model.x_EXC_ABB;
            BOQ_JKH_newjob.x_PU_DESC = model.x_PU_DESC;
            BOQ_JKH_newjob.x_YEAR_INSTALL = model.x_YEAR_INSTALL.ToString();
            BOQ_JKH_newjob.x_SCH_TYPE = model.x_SCH_TYPE.ToString();
            BOQ_JKH_newjob.x_SCH_NO = model.x_SCH_NO;
            BOQ_JKH_newjob.x_SCHEME_NAME = model.x_SCHEME_NAME;
            BOQ_JKH_newjob.x_PU_ID = model.x_PU_ID;
            BOQ_JKH_newjob.x_BQ_MAT_PRICE = model.x_BQ_MAT_PRICE.ToString();
            BOQ_JKH_newjob.x_BQ_INSTALL_PRICE = model.x_BQ_INSTALL_PRICE.ToString();
            BOQ_JKH_newjob.x_PU_QTY = model.x_PU_QTY.ToString();
            BOQ_JKH_newjob.x_OLD_MAT_PR = model.x_OLD_MAT_PR.ToString();
            BOQ_JKH_newjob.x_OLD_INSTALL_PR = model.x_OLD_INSTALL_PR.ToString();
            BOQ_JKH_newjob.x_CONSTRUCT_BY = model.x_CONSTRUCT_BY;
            BOQ_JKH_newjob.x_RECVR_QTY = model.x_RECVR_QTY.ToString();
            BOQ_JKH_newjob.x_RATE_INDICATOR = model.x_RATE_INDICATOR;
            success = myWebService.AddBOQ_JKH(BOQ_JKH_newjob);

                selected = true;
           

            if (ModelState.IsValid && selected)
            {
                if(success == true)
                    return RedirectToAction("BOQ_JKH_NewSave");
                else
                    return RedirectToAction("NewSaveFail"); // store to db failed.
            }           

            return View(model);
        }

        public ActionResult BOQ_JKH_NewSave()
        {
            return View();
        }

        public ActionResult GetDetails_BOQ_JKH(string id)
        {           

            using (Entities ctxData = new Entities())
            {
                var query = (from p in ctxData.WV_BOQ_DATA
                             where p.EXC_ABB == id
                             select p).Single();

                return Json(new
                {    
                   
                    Success = true,
                    x_EXC_ABB = id,                    
                    x_PU_DESC = query.PU_DESC,
                    x_YEAR_INSTALL = query.YEAR_INSTALL,
                    x_SCH_TYPE = query.SCH_TYPE,
                    x_SCH_NO = query.SCH_NO,
                    x_SCHEME_NAME = query.SCHEME_NAME,
                    x_PU_ID = query.PU_ID,
                    x_BQ_MAT_PRICE = query.BQ_MAT_PRICE,
                    x_BQ_INSTALL_PRICE = query.BQ_INSTALL_PRICE,
                    x_PU_QTY = query.PU_QTY,
                    x_OLD_MAT_PR = query.OLD_MAT_PR,
                    x_OLD_INSTALL_PR = query.OLD_INSTALL_PR,
                    x_CONSTRUCT_BY = query.CONSTRUCT_BY,
                    x_RECVR_QTY = query.RECVR_QTY,
                    x_RATE_INDICATOR = query.RATE_INDICATOR,
                }, JsonRequestBehavior.AllowGet); //
            }
        }

        public ActionResult UpdateData_BOQ_JKH(string txtEXC_ABB, string txtPU_DESC, string txtYEAR_INSTALL, string txtSCH_TYPE, string txtSCH_NO, string txtSCHEME_NAME, string txtPU_ID, string txtBQ_MAT_PRICE, string txtBQ_INSTALL_PRICE, string txtPU_QTY, string txtOLD_MAT_PR, string txtOLD_INSTALL_PR, string txtCONSTRUCT_BY, string txtRECVR_QTY, string txtRATE_INDICATOR)
        {           
            bool success = true;                   

            WebView.WebService._base myWebService;
            myWebService = new WebService._base();

            WebService._base.BOQ_JKH BOQ_JKH_newjob = new WebService._base.BOQ_JKH();
            BOQ_JKH_newjob.x_EXC_ABB = txtEXC_ABB;
            BOQ_JKH_newjob.x_PU_DESC = txtPU_DESC;
            BOQ_JKH_newjob.x_YEAR_INSTALL = txtYEAR_INSTALL;
            BOQ_JKH_newjob.x_SCH_TYPE = txtSCH_TYPE;
            BOQ_JKH_newjob.x_SCH_NO = txtSCH_NO;
            BOQ_JKH_newjob.x_SCHEME_NAME = txtSCHEME_NAME;
            BOQ_JKH_newjob.x_PU_ID = txtPU_ID;
            BOQ_JKH_newjob.x_BQ_MAT_PRICE = txtBQ_MAT_PRICE;
            BOQ_JKH_newjob.x_BQ_INSTALL_PRICE = txtBQ_INSTALL_PRICE;
            BOQ_JKH_newjob.x_PU_QTY = txtPU_QTY;
            BOQ_JKH_newjob.x_OLD_MAT_PR = txtOLD_MAT_PR;
            BOQ_JKH_newjob.x_OLD_INSTALL_PR = txtOLD_INSTALL_PR;
            BOQ_JKH_newjob.x_CONSTRUCT_BY = txtCONSTRUCT_BY;
            BOQ_JKH_newjob.x_RECVR_QTY = txtRECVR_QTY;
            BOQ_JKH_newjob.x_RATE_INDICATOR = txtRATE_INDICATOR;
            success = myWebService.UpdateBOQ_JKH(BOQ_JKH_newjob, txtEXC_ABB);

            return Json(new
            {
               Success = success
            }, JsonRequestBehavior.AllowGet);
        }

        [HttpPost]
        public ActionResult DeleteData_BOQ_JKH(string targetBOQ_JKH)
        {
            bool success;

            WebView.WebService._base myWebService;
            myWebService = new WebService._base();

            success = myWebService.Delete_BOQ_JKH(targetBOQ_JKH);

            return Json(new
            {
                Success = success
            }, JsonRequestBehavior.AllowGet);
        }
    }
}

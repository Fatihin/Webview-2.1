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


namespace WebView.Controllers
{
    public class MileageRatesController : Controller
    {
        //
        // GET: /MileageRates/

        public ActionResult Index()
        {
            return View();
        }

        public ActionResult List()
        {
            WebView.WebService._base myWebService;
            myWebService = new WebService._base();

            WebService._base.OSPMRates mRates = new WebService._base.OSPMRates();
            mRates = myWebService.GetOSPMileageRates(0, 10);

            ViewData["data3"] = mRates.MileageRatesList;

            string input = "\\\\adsvr";
            //string input = "\\\\server\\d$\\x\\y\\z\\AAA";
            string output = String.Format("http:{0}", input.Replace("\\d$\\x\\y", String.Empty).Replace("\\", "/"));

            ViewBag.output = output;

            return View();
        }

        public ActionResult NewMileageRates()
        {
           return View();
        }
        
        [HttpPost]
        public ActionResult NewMileageRates(MileageRatesModel m)
        {
            bool success = true;
            WebView.WebService._base myWebService;
            myWebService = new WebService._base();

            bool selected = false;

            WebService._base.MileageRates newgrade = new WebService._base.MileageRates();
            newgrade.GRADE = m.Grade;
            newgrade.GRADE_DESC = m.GradeDescription;
            newgrade.KM_LLEV1 = Convert.ToString(m.LowLev1);
            newgrade.KM_LLEV2 = Convert.ToString(m.LowLev2);
            newgrade.KM_LLEV3 = Convert.ToString(m.LowLev3);
            newgrade.KM_HLEV1 = Convert.ToString(m.HighLev1);
            newgrade.KM_HLEV2 = Convert.ToString(m.HighLev2);
            newgrade.KM_HLEV3 = Convert.ToString(m.HighLev3);
            newgrade.KM_RATE1 = Convert.ToString(m.Rate1);
            newgrade.KM_RATE2 = Convert.ToString(m.Rate2);
            newgrade.KM_RATE3 = Convert.ToString(m.Rate3);           

            success = myWebService.AddMRates(newgrade);
            selected = true;

            if (ModelState.IsValid && selected)
            {
                if (success == true)
                    return RedirectToAction("NewSave");
                else
                    return RedirectToAction("NewSaveFail"); // store to db failed.
            }
                     

            return View(m);
        }

        public ActionResult NewSave()
        {
            return View();
        }

        [HttpPost]
        public ActionResult GetDetails(string id)
        {
            
            string fileListStr = "";

            using (Entities ctxData = new Entities())
            {
                var query = (from p in ctxData.WV_MILEAGE_RATES
                             where p.GRADE== id
                             select p).Single();

                return Json(new
                {
                    Success = true,
                    grade = id,                    
                    description = query.GRADE_DESC,
                    kmHLev1 = query.KM_HLEV1,
                    kmHLev2 = query.KM_HLEV2,
                    kmHLev3 = query.KM_HLEV3,
                    kmLLev1 = query.KM_LLEV1,
                    kmLLev2 = query.KM_LLEV2,
                    kmLLev3 = query.KM_LLEV3,
                    kmRate1 = query.KM_RATE1,
                    kmRate2 = query.KM_RATE2,
                    kmRate3 = query.KM_RATE3,
                    files_list = new
                    {
                        files = fileListStr
                    }
                }, JsonRequestBehavior.AllowGet); //
            }
        }

        [HttpPost]
        public ActionResult DeleteData(string targetGrade)
        {
            bool success;

            WebView.WebService._base myWebService;
            myWebService = new WebService._base();

            success = myWebService.DeleteMRates(targetGrade);

            return Json(new
            {
                Success = success
            }, JsonRequestBehavior.AllowGet);
        }

        public ActionResult UpdateData(string GradeVal, string txtDescription, string txtLLev1, string txtLLev2, string txtLLev3, string txtHLev1, string txtHLev2, string txtHLev3, string txtRate1, string txtRate2, string txtRate3)
        {
            bool success = true;           

            WebView.WebService._base myWebService;
            myWebService = new WebService._base();

            WebService._base.MileageRates newRates= new WebService._base.MileageRates();
            newRates.GRADE_DESC = txtDescription;
            newRates.KM_LLEV1 = txtLLev1;
            newRates.KM_LLEV2 = txtLLev2;
            newRates.KM_LLEV3 = txtLLev3;
            newRates.KM_HLEV1 = txtHLev1;
            newRates.KM_HLEV2 = txtHLev2;
            newRates.KM_HLEV3 = txtHLev3;
            newRates.KM_RATE1 = txtRate1;
            newRates.KM_RATE2 = txtRate2;
            newRates.KM_RATE3 = txtRate3;
            success = myWebService.UpdateMRates(newRates, GradeVal);

            return Json(new
            {
                Success = success
            }, JsonRequestBehavior.AllowGet);
        }
    }
}

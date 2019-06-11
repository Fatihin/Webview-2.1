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
    public class MaterialClassMaintenanceController : Controller
    {
        //
        // GET: /MaterialClassMaintenance/

        public ActionResult Index()
        {
            return View();
        }

        public ActionResult List(string searchKey, int? page)
        {
            WebView.WebService._base myWebService;
            myWebService = new WebService._base();

            WebService._base.OSPMaterialClassMaintenance MatClassM = new WebService._base.OSPMaterialClassMaintenance();
            MatClassM = myWebService.GetOSPMaterialClassMaintenance(0, 100000, searchKey);

            string input = "\\\\adsvr";
            //string input = "\\\\server\\d$\\x\\y\\z\\AAA";
            string output = String.Format("http:{0}", input.Replace("\\d$\\x\\y", String.Empty).Replace("\\", "/"));

            ViewBag.output = output;

            int pageSize = 10;
            int pageNumber = (page ?? 1);
            return View(MatClassM.MatClassMaintenanceList.ToPagedList(pageNumber, pageSize));
        }

        public ActionResult NewMaterialClassMaintenance()
        {
            return View();
        }

        [HttpPost]
        public ActionResult NewMaterialClassMaintenance(MaterialClassMaintenanceModel m)
        {
            bool success = true;
            WebView.WebService._base myWebService;
            myWebService = new WebService._base();

            bool selected = false;

            WebService._base.MaterialClassMaintenance newMatClassM = new WebService._base.MaterialClassMaintenance();

            newMatClassM.MAT_CLASS = m.MatClass;
            newMatClassM.CLASS_NAME = m.ClassName;            
            success = myWebService.AddMatClassMaintenance(newMatClassM);
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
                var query = (from q in ctxData.WV_MAT_CLASS
                             where q.MAT_CLASS == id
                             select q).Single();

                return Json(new
                {
                    Success = true,
                    MatClass = id,
                    ClassName = query.CLASS_NAME,                    
                    files_list = new
                    {
                        files = fileListStr
                    }
                }, JsonRequestBehavior.AllowGet); //
            }
        }

        [HttpPost]
        public ActionResult DeleteData(string targetMatClass)
        {
            bool success;

            WebView.WebService._base myWebService;
            myWebService = new WebService._base();

            success = myWebService.DeleteMatClassMaintenance(targetMatClass);

            return Json(new
            {
                Success = success
            }, JsonRequestBehavior.AllowGet);
        }

        public ActionResult UpdateData(string MatClassVal, string txtClassName)
        {
            bool success = true;

            WebView.WebService._base myWebService;
            myWebService = new WebService._base();

            WebService._base.MaterialClassMaintenance newMatClassM = new WebService._base.MaterialClassMaintenance();
            newMatClassM.CLASS_NAME = txtClassName;
            success = myWebService.UpdateMatClassMaintenance(newMatClassM, MatClassVal);

            return Json(new
            {
                Success = success
            }, JsonRequestBehavior.AllowGet);
        }
    }
}

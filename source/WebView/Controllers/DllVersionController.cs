using System;
using System.Collections.Generic;
using System.Linq;
using System.Web.Mvc;
using PagedList;
using WebView.Models;

namespace WebView.Controllers
{
    public class DllVersionController : Controller
    {
        public ActionResult List_Version(string searchKey)
        {
            WebView.WebService._base myWebService;
            myWebService = new WebService._base();
            string res = myWebService.GetDllVersionMaintenance(searchKey);
            
            return Json(new
            {
                record = res
            }, JsonRequestBehavior.AllowGet);
        }

        public ActionResult List(string searchKey, int? page)
        {
            WebView.WebService._base myWebService;
            myWebService = new WebService._base();

            WebService._base.ListDllVersion DllV = new WebService._base.ListDllVersion();

            if (searchKey != null)
            {
                if (searchKey.Equals(""))
                {
                    //DllV = myWebService.GetDllVersionMaintenance(0, 1000, null);
                    ViewBag.searchKey = "";
                }
                else
                {
                    //DllV = myWebService.GetDllVersionMaintenance(0, 1000, searchKey);
                    ViewBag.searchKey = searchKey;
                }
            }
            else
            {
                //DllV = myWebService.GetDllVersionMaintenance(0, 1000, null);
                ViewBag.searchKey = "";
            }

            string input = "\\\\adsvr";
            //string input = "\\\\server\\d$\\x\\y\\z\\AAA";
            string output = String.Format("http:{0}", input.Replace("\\d$\\x\\y", String.Empty).Replace("\\", "/"));

            ViewBag.output = output;

            

            int pageSize = 10;
            int pageNumber = (page ?? 1);
            return View(DllV.DllVersionList.ToPagedList(pageNumber, pageSize));
        }

        public ActionResult NewJob()
        {
            List<SelectListItem> list = new List<SelectListItem>();
            using (Entities ctxData = new Entities())
            {
                var query = from p in ctxData.WV_DLL_MASTER
                            select new { Text = p.DLL_NAME, Value = p.DLL_NAME };

                foreach (var a in query)
                {
                    list.Add(new SelectListItem() { Text = a.Text, Value = a.Value.ToString() });
                }

                ViewBag.DllMaster = list;
            }

            return View(new WebView.Models.DllVersion_NewJobModel());
        }

        [HttpPost]
        public ActionResult NewJob(DllVersion_NewJobModel model)
        {
            bool success = true;
            WebView.WebService._base myWebService;
            myWebService = new WebService._base();

            bool selected = false;

            using (Entities ctxData = new Entities())
            {
                var query = (from p in ctxData.WV_DLL_VERSION
                             where p.DLL_MASTER == model.DLL_MASTER && p.DLL_VERSION == model.DLL_VERSION
                             select p).Count();
                if (query == 0)
                {
                    WebService._base.DLLVersion newjob = new WebService._base.DLLVersion();
                    newjob.DLL_MASTER = model.DLL_MASTER;
                    newjob.DLL_NAME = model.DLL_NAME;
                    newjob.DLL_DESCRIPTION = model.DLL_DESCRIPTION;
                    newjob.DLL_VERSION = model.DLL_VERSION;
                    newjob.CREATE_DATE = model.CREATE_DATE;
                    success = myWebService.AddDllVersion(newjob);

                    selected = true;
                    if (ModelState.IsValid && selected)
                    {
                        if (success == true)
                            return RedirectToAction("NewSave");
                        else
                            return RedirectToAction("NewSaveFail"); // store to db failed.
                    }
                }
                else { return RedirectToAction("NewSaveFail"); // store to db failed.
                
                }
            }


            

            return View(model);
        }

        public ActionResult NewSave()
        {
            return View();
        }

        public ActionResult GetDetails_Version(string id, string id2)
        {
            using (Entities ctxData = new Entities())
            {
                var query = (from p in ctxData.WV_DLL_VERSION
                             where p.DLL_MASTER == id && p.DLL_VERSION == id2
                             select p).Single();

                return Json(new
                {

                    Success = true,
                    master = id,
                    name = query.DLL_NAME,
                    version = query.DLL_VERSION,
                    desc = query.DLL_DESCRIPTION,
                    createdAt = query.CREATE_DATE.ToString()
                }, JsonRequestBehavior.AllowGet); //
            }
        }

        public ActionResult UpdateData(string txtMaster, string txtName, string txtDesc, string txtVersion)
        {
            bool success = true;

            WebView.WebService._base myWebService;
            myWebService = new WebService._base();

            WebService._base.DLLVersion upt_data = new WebService._base.DLLVersion();
            upt_data.DLL_MASTER = txtMaster;
            upt_data.DLL_NAME = txtName;
            upt_data.DLL_VERSION = txtVersion;
            upt_data.DLL_DESCRIPTION = txtDesc;
            success = myWebService.UpdateDllVersion(upt_data);

            return Json(new
            {
                Success = success
            }, JsonRequestBehavior.AllowGet);
        }

        [HttpPost]
        public ActionResult DeleteData(string targetId, string targetId2)
        {
            bool success;

            WebView.WebService._base myWebService;
            myWebService = new WebService._base();

            success = myWebService.Delete_DllVersion(targetId, targetId2);

            return Json(new
            {
                Success = success
            }, JsonRequestBehavior.AllowGet);
        }
    }
}
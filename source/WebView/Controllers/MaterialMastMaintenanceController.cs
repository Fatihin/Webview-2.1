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
    public class MaterialMastMaintenanceController : Controller
    {
        //
        // GET: /MaterialMastMaintenance/

        public ActionResult Index()
        {
            return View();
        }

        public ActionResult List(string searchKey,string searchUOM, int? page)
        {
            WebView.WebService._base myWebService;
            myWebService = new WebService._base();
            //System.Diagnostics.Debug.WriteLine("Controller 1 :" + "  search key : " + searchKey + " search UOM : " + searchUOM);
            WebService._base.OSPMaterialMasterMaintenance MatMastM = new WebService._base.OSPMaterialMasterMaintenance();

            MatMastM = myWebService.GetOSPMaterialMasterMaintenance(0, 100000, searchKey, searchUOM);
                ViewBag.searchKey = searchKey;
                ViewBag.searchUOM2 = searchUOM;

            string input = "\\\\adsvr";
            //string input = "\\\\server\\d$\\x\\y\\z\\AAA";
            string output = String.Format("http:{0}", input.Replace("\\d$\\x\\y", String.Empty).Replace("\\", "/"));

            ViewBag.output = output;


            List<SelectListItem> listMatUOM = new List<SelectListItem>();
            using (Entities ctxData = new Entities())
            {
                var query = (from p in ctxData.WV_MAT_MAST
                             orderby p.MAT_UOM
                             select new { p.MAT_UOM, });

                listMatUOM.Add(new SelectListItem() { Text = "", Value = "Select" });

                foreach (var a in query.Distinct())
                {
                    if (a.MAT_UOM != null)
                    {
                        listMatUOM.Add(new SelectListItem() { Text = a.MAT_UOM, Value = a.MAT_UOM });
                    }
                }

                ViewBag.MatUOM = listMatUOM;
            }

            int pageSize = 10;
            int pageNumber = (page ?? 1);
            return View(MatMastM.MatMastMaintenanceList.ToPagedList(pageNumber, pageSize));
        }

        public ActionResult NewMaterialMastMaintenance()
        {
            List<SelectListItem> listMatUOM = new List<SelectListItem>();
            using (Entities ctxData = new Entities())
            {
                var query = (from p in ctxData.WV_MAT_MAST 
                             orderby p.MAT_UOM descending 
                            select new {p.MAT_UOM  });

                foreach (var a in query.Distinct())
                {
                    if (a.MAT_UOM != null)
                    {
                        listMatUOM.Add(new SelectListItem() { Text = a.MAT_UOM, Value = a.MAT_UOM });
                    }
                }

                ViewBag.MatUOM = listMatUOM;
            }

            List<SelectListItem> listMatMatCat = new List<SelectListItem>();
            using (Entities ctxData = new Entities())
            {
                var query = from p in ctxData.WV_MAT_MAST
                            select new { Text = p.MAT_CAT, Value = p.MAT_CAT };

                foreach (var a in query.Distinct())
                {
                    listMatMatCat.Add(new SelectListItem() { Text = a.Text, Value = a.Value.ToString() });
                }

                ViewBag.MatCat = listMatMatCat;
            }
            return View();
        }

        [HttpPost]
        public ActionResult NewMaterialMastMaintenance(MaterialMastMaintenanceModel m)
        {
            bool success = true;
            WebView.WebService._base myWebService;
            myWebService = new WebService._base();

            bool selected = false;

            WebService._base.MaterialMasterMaintenance newMatMastM = new WebService._base.MaterialMasterMaintenance();

            System.Diagnostics.Debug.WriteLine("base :2  Mat Name [" + newMatMastM.MAT_NAME + "]   Mat UOM  [" + newMatMastM.MAT_UOM + "]   Mat Price [" + newMatMastM.MAT_PRICE + "] Mat Cat [" + newMatMastM.MAT_CAT + "]  Record Type [" + newMatMastM.RECORD_TYPE + "]  Tras Type [" + newMatMastM.TRANS_TYPE + "] ");
           
            newMatMastM.MAT_NAME = m.MatName;
            newMatMastM.MAT_UOM = m.MatUOM;
            newMatMastM.MAT_PRICE = m.MatPrice.ToString();
            newMatMastM.MAT_CAT = m.MatCat;
            newMatMastM.RECORD_TYPE = m.RecordType;
            newMatMastM.TRANS_TYPE = m.TransType;
            success = myWebService.AddMatMastMaintenance(newMatMastM);
            selected = true;

            //if (ModelState.IsValid && selected)
            //{
            if (success == true)
            {
                return RedirectToAction("NewSave");
            }
            else
            {
                return RedirectToAction("NewSaveFail"); // store to db failed.
            }
            //}

            //return View(m);
        }

        public ActionResult NewSave()
        {
            return View();
        }

        [HttpPost]
        public ActionResult GetDetails(string id)
        {
            string fileListStr = "";
            string datauom = "";
            string dataCat = "";

            using (Entities ctxData = new Entities())
            {
                var query = (from q in ctxData.WV_MAT_MAST
                             where q.MAT_ID == id
                             select q).Single();

                var queryUOM = (from q in ctxData.WV_MAT_MAST
                                select new { q.MAT_UOM, q.MAT_CAT });

                foreach (var a in queryUOM.Distinct())
                {
                    datauom = datauom + a.MAT_UOM + ":";
                }
                foreach (var a in queryUOM.Distinct())
                {
                    dataCat = dataCat + a.MAT_CAT + ":";
                }
                
                System.Diagnostics.Debug.WriteLine("OUM :" + datauom + "  Cat :" + dataCat);
                return Json(new
                {
                    Success = true,
                    MatId = id,
                    MatName = query.MAT_NAME,
                    MatUOM = query.MAT_UOM,
                    MatPrice = query.MAT_PRICE,
                    MatCat = query.MAT_CAT,
                    RecType = query.RECORD_TYPE,
                    TransType = query.TRANS_TYPE, 
                    UOM = datauom,
                    Cat = dataCat,

                    files_list = new
                    {
                        files = fileListStr
                    }
                }, JsonRequestBehavior.AllowGet); //
            }
        }

        [HttpPost]
        public ActionResult DeleteData(string targetMatId)
        {
            bool success;

            WebView.WebService._base myWebService;
            myWebService = new WebService._base();

            success = myWebService.DeleteMatMastMaintenance(targetMatId);

            return Json(new
            {
                Success = success
            }, JsonRequestBehavior.AllowGet);
        }

        public ActionResult UpdateData(string MatIdVal, string txtMatName, string txtMatUOM, string txtMatPrice, string txtMatCat, string txtRecType, string txtTransType)
        {
            bool success = true;

            WebView.WebService._base myWebService;
            myWebService = new WebService._base();

            WebService._base.MaterialMasterMaintenance newMatMast = new WebService._base.MaterialMasterMaintenance();
            newMatMast.MAT_NAME = txtMatName;
            newMatMast.MAT_UOM = txtMatUOM;
            newMatMast.MAT_PRICE = txtMatPrice;
            newMatMast.MAT_CAT = txtMatCat;
            newMatMast.RECORD_TYPE = txtRecType;
            newMatMast.TRANS_TYPE = txtTransType;          
            success = myWebService.UpdateMatMastMaintenance(newMatMast, MatIdVal);

            return Json(new
            {
                Success = success
            }, JsonRequestBehavior.AllowGet);
        }
    }
}

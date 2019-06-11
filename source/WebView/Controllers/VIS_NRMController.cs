using System;
using System.Collections.Generic;
using System.Linq;
using WebView.Library;
using System.Net;
using System.Net.Mail;
using System.Web.Mvc;
using System.Configuration;
using PagedList;
using WebView.Models;

namespace WebView.Controllers
{
    public class VIS_NRMController : Controller
    {
        // GET: /PUFRM/
        private string connString = ConfigurationManager.AppSettings.Get("connString");

        public ActionResult Index()
        {
            return View();
        }

        public ActionResult NewPath(string searchKey, int? page)
        {
            using (Entities ctxData = new Entities())
            {
                //filter PTT
                List<SelectListItem> list = new List<SelectListItem>();
                var queryEXC = from p in ctxData.WV_EXC_MAST
                               select new { Text = p.PTT_ID, Value = p.PTT_ID };

                list.Add(new SelectListItem() { Text = "", Value = "Select" });
                foreach (var a in queryEXC.Distinct().OrderBy(it => it.Value))
                {
                    if (a.Value != null)
                        list.Add(new SelectListItem() { Text = a.Text.ToString(), Value = a.Value.ToString() });
                }
                ViewBag.PTT = list;

                List<SelectListItem> list1 = new List<SelectListItem>();
                var querySCHEME = from p in ctxData.WV_ISP_JOB
                                  where (p.G3E_OWNER.ToUpper() == User.Identity.Name.ToUpper() || p.G3E_OWNER.ToLower() == User.Identity.Name.ToLower()) && (p.JOB_TYPE == "METROE" || p.JOB_TYPE == "IPCORE" || p.JOB_TYPE.Contains("CID"))
                                  select new { Text = p.G3E_IDENTIFIER, Value = p.G3E_IDENTIFIER };

                list1.Add(new SelectListItem() { Text = "", Value = "Select" });
                foreach (var a in querySCHEME.Distinct().OrderBy(it => it.Value))
                {
                    if (a.Value != null)
                        list1.Add(new SelectListItem() { Text = a.Text.ToString(), Value = a.Value.ToString() });
                }

                ViewBag.job = list1;

                List<SelectListItem> list2 = new List<SelectListItem>();
                list2.Add(new SelectListItem() { Text = "", Value = "Select" });
                ViewBag.pathType = list2;
                ViewBag.equipA = list2;
                ViewBag.equipZ = list2;
                ViewBag.cardA = list2;
                ViewBag.cardZ = list2;
                ViewBag.exc = list2;
            }
            return View();
        }

        [HttpPost]
        public ActionResult updatePath(string job) // Update equip
        {
            string[] jobSplit = job.Split('-');
            string jobType = jobSplit[1];
            string listData = "";
            string listJob = "";
            using (Entities ctxData = new Entities())
            {
                List<SelectListItem> list2 = new List<SelectListItem>();
                if (jobType == "METROE")
                {
                    var queryPATH = from p in ctxData.WV_PATH_MATRIX
                                    where (p.PATH_TYPE == "INTERCONNECT" || p.PATH_TYPE == "METRO ETHERNET")
                                    select new { p.PATH_TYPE, p.A_TYPE, p.Z_TYPE };

                    foreach (var a in queryPATH.Distinct().OrderBy(it => it.PATH_TYPE))
                    {
                        listData = listData + a.PATH_TYPE + ":" + a.A_TYPE + ":" + a.Z_TYPE + ":" + "|";
                    }
                }
                else if (jobType.Contains("CID"))
                {
                    var queryPATH = from p in ctxData.WV_PATH_MATRIX
                                    where (p.PATH_TYPE == "INTERCONNECT" || p.PATH_TYPE == "IPCORE" || p.PATH_TYPE == "METRO ETHERNET")
                                    select new { p.PATH_TYPE, p.A_TYPE, p.Z_TYPE };

                    foreach (var a in queryPATH.Distinct().OrderBy(it => it.PATH_TYPE))
                    {
                        listData = listData + a.PATH_TYPE + ":" + a.A_TYPE + ":" + a.Z_TYPE + ":" + "|";
                    }
                }
                else if (jobType.Contains("HSBB"))
                {
                    var queryPATH = from p in ctxData.WV_PATH_MATRIX
                                    where p.PATH_TYPE.Trim() == "BEARER FDC - OLT"
                                    select new { p.PATH_TYPE, p.A_TYPE, p.Z_TYPE };

                    foreach (var a in queryPATH.Distinct().OrderBy(it => it.PATH_TYPE))
                    {
                        listData = listData + a.PATH_TYPE + ":" + a.A_TYPE + ":" + a.Z_TYPE + ":" + "|";
                    }
                }

                var queryJOB = from p in ctxData.WV_LOAD_PATH
                               where p.JOBID == job
                               select new { p.ANAME, p.ATYPE, p.ASITE, p.ASLOT, p.ACARD, p.APORT, p.ZNAME, p.ZTYPE, p.ZSITE, p.ZSLOT, p.ZCARD, p.ZPORT, p.PRIMARYSECONDARY, p.PATHBANDWIDTH, p.ID };

                foreach (var a in queryJOB.OrderBy(it => it.ID))
                {

                    listJob = listJob + a.ANAME + ":" + a.ATYPE + ":" + a.ASITE + ":" + a.ASLOT + ":" + a.ACARD + ":" + a.APORT + ":" + a.ZNAME + ":" + a.ZTYPE + ":" + a.ZSITE + ":" + a.ZSLOT + ":" + a.ZCARD + ":" + a.ZPORT + ":" + a.PRIMARYSECONDARY + ":" + a.PATHBANDWIDTH + ":" + a.ID + ":" + "|";
                }
                if (queryJOB.Count() == 0)
                {
					// region Mubin - CR14-20180330
                    using (Entities9 ctxData2 = new Entities9())
                    {
                        var queryJOBFTTH = from p in ctxData2.WV_LOAD_PATHCONSUMER
                                           where p.JOBID == job
                                           select new { p.APORT2, p.ANAME, p.ATYPE, p.ASITE, p.ACARD3, p.APORT3, p.ZNAME, p.ZTYPE, p.ZSITE, p.ZCARD, p.ZPORT, p.DPNAME, p.ID, p.DPPORT ,p.REMARKS};

                        foreach (var a in queryJOBFTTH.OrderBy(it => it.DPNAME))
                        {
                            string St_REMARKS = "";
                            if (a.REMARKS != "" && a.REMARKS != null)
                            {
                                St_REMARKS = " (Remarks -  " + a.REMARKS + " ) ";
                            }

                            listJob = listJob + a.ANAME + ":" + a.ATYPE + ":" + a.ASITE + ":" + a.ACARD3 + ":" + a.APORT2 + ":" + a.DPPORT + ":" + a.ZNAME + ":" + a.ZTYPE + ":" + a.ZSITE + ":" + a.ZCARD + ":" + a.ZPORT + ":" + a.DPNAME + St_REMARKS + ":" + a.ID + ":" + "|";
                        }
                    }
                }
            }
            /*using (Entities7 ctxData = new Entities7())
            {
                int data_count = (from p in ctxData.WV_LOAD_PATH_CONSUMER_ODP
                                  where p.JOBID == job
                                  select p).Count();
                if (data_count > 0)
                {
                    var queryJOBFTTH = from p in ctxData.WV_LOAD_PATH_CONSUMER_ODP
                                       where p.JOBID == job
                                       select new { p.APORT2, p.ANAME, p.ATYPE, p.ASITE, p.ACARD3, p.APORT3, p.ZNAME, p.ZTYPE, p.ZSITE, p.ZCARD, p.ZPORT, p.DPNAME, p.ID, p.DPPORT,p.REMARKS };

                    foreach (var a in queryJOBFTTH.OrderBy(it => it.DPNAME))
                    {
                        string St_REMARKS = "";
                        if (a.REMARKS != "" && a.REMARKS != null)
                        {
                            St_REMARKS = " (Remarks - " + a.REMARKS + " ) ";
                        }
                        listJob = listJob + a.ANAME + ":" + a.ATYPE + ":" + a.ASITE + ":" + a.ACARD3 + ":" + a.APORT2 + ":" + a.DPPORT + ":" + a.ZNAME + ":" + a.ZTYPE + ":" + a.ZSITE + ":" + a.ZCARD + ":" + a.ZPORT + ":" + a.DPNAME + St_REMARKS + ":" + a.ID + ":" + "|";
                    }
                }
            }*/
			// endRegion
          
            return Json(new
            {
                Success = true,
                listData = listData,
                listJob = listJob
            }, JsonRequestBehavior.AllowGet); //
        }

        [HttpPost]
        public ActionResult updateEquipA(string job, string pathType, string ZEXC) // Update card
        {
            string[] jobSplit = job.Split('-');
            string exc = jobSplit[0];
            string[] pathTypeSplit1 = pathType.Split('-');
            string pathType1 = pathTypeSplit1[0];
            string pathType2 = pathTypeSplit1[1];

            string listData1 = "";
            string listData2 = "";
            string condition1 = pathType1;
            string condition2 = pathType1;
            string condition3 = pathType2;
            string condition4 = pathType2;
            if (pathType1 == "BRAS") 
            {
                condition1 = "ERX";
                condition2 = "BRS";
            }
            if (pathType2 == "BRAS")
            {
                condition3 = "ERX";
                condition4 = "BRS";
            }
            System.Diagnostics.Debug.WriteLine("exc Type : " + exc);
            System.Diagnostics.Debug.WriteLine("Path Type : " + condition1);
            System.Diagnostics.Debug.WriteLine("ZEXC Type : " + ZEXC);
            using (Entities_NRM ctxData = new Entities_NRM())
            {
                List<SelectListItem> list1 = new List<SelectListItem>();
                var queryEquipA = from p in ctxData.VFDITEMs
                                  join fx in ctxData.VFDITEMPLACEMENTs on p.ID equals fx.ITEM_ID
                                  join e in ctxData.VFDITEMs on fx.CONTAINER_ID equals e.ID
                                  join fxx in ctxData.VFDITEMPLACEMENTs on e.ID equals fxx.CONTAINER_ID
                                  where e.AUTODISPLAYNAME.Contains(exc) && e.DTYPE == "Rack" &&
                                        (p.AUTODISPLAYNAME.Contains(condition1) || p.AUTODISPLAYNAME.Contains(condition2))
                                  select new { p.AUTODISPLAYNAME, p.ID };

                foreach (var a in queryEquipA.Distinct().OrderBy(it => it.AUTODISPLAYNAME))
                {
                    listData1 = listData1 + a.AUTODISPLAYNAME + ":" + a.ID + "|";
                }

                List<SelectListItem> list2 = new List<SelectListItem>();
                var queryEquipZ = from p in ctxData.VFDITEMs
                                  join fx in ctxData.VFDITEMPLACEMENTs on p.ID equals fx.ITEM_ID
                                  join e in ctxData.VFDITEMs on fx.CONTAINER_ID equals e.ID
                                  join fxx in ctxData.VFDITEMPLACEMENTs on e.ID equals fxx.CONTAINER_ID
                                  where e.AUTODISPLAYNAME.Contains(ZEXC) && e.DTYPE == "Rack" &&
                                        (p.AUTODISPLAYNAME.Contains(condition3) || p.AUTODISPLAYNAME.Contains(condition4))
                                  select new { p.AUTODISPLAYNAME, p.ID, p.USERPARTIALNAME };

                foreach (var a in queryEquipZ.Distinct().OrderBy(it => it.AUTODISPLAYNAME))
                {
                    listData2 = listData2 + a.AUTODISPLAYNAME + ":" + a.ID + ":" + a.USERPARTIALNAME  + "|";
                    System.Diagnostics.Debug.WriteLine("USER PARTIAL Type : " + a.USERPARTIALNAME);
                }
            }

            return Json(new
            {
                Success = true,
                listData = listData1,
                listData2 = listData2
            }, JsonRequestBehavior.AllowGet); //
        }

        [HttpPost]
        public ActionResult updateCardA(int equip) // Update port
        {
            string listData = "";
            using (Entities_NRM ctxData = new Entities_NRM())
            {
                List<SelectListItem> list2 = new List<SelectListItem>();
                var queryParent = from p in ctxData.VFDITEMs
                                  join fx in ctxData.VFDITEMPLACEMENTs on p.ID equals fx.ITEM_ID
                                  //join e in ctxData.VFDITEMs on fx.CONTAINER_ID equals e.ID
                                  //join fxx in ctxData.VFDITEMPLACEMENTs on e.ID equals fxx.ITEM_ID
                                  where fx.CONTAINER_ID == equip
                                  select new { p.AUTODISPLAYNAME, p.ID };

                foreach (var a in queryParent.Distinct().OrderBy(it => it.AUTODISPLAYNAME))
                {
                    listData = listData + a.AUTODISPLAYNAME + ":" + a.ID + "|";
                }
                //sub card
                var queryPATH = from p in ctxData.VFDITEMs
                                join fx in ctxData.VFDITEMPLACEMENTs on p.ID equals fx.ITEM_ID
                                join e in ctxData.VFDITEMs on fx.CONTAINER_ID equals e.ID
                                join fxx in ctxData.VFDITEMPLACEMENTs on e.ID equals fxx.ITEM_ID
                                where fxx.CONTAINER_ID == equip
                                select new { p.AUTODISPLAYNAME, p.ID };

                foreach (var a in queryPATH.Distinct().OrderBy(it => it.AUTODISPLAYNAME))
                {
                    listData = listData + a.AUTODISPLAYNAME + ":" + a.ID + "|";
                }
            }

            return Json(new
            {
                Success = true,
                listData = listData
            }, JsonRequestBehavior.AllowGet); //
        }

        [HttpPost]
        public ActionResult updatePortA(string cardA) // Update list port
        {
            string listData = "";
            string[] card = cardA.Split('-');
            string testc = card[0];
            int cardID = Convert.ToInt32(testc);
            System.Diagnostics.Debug.WriteLine("cardA : " + cardA);

            using (Entities_NRM ctxData = new Entities_NRM())
            {
                List<SelectListItem> list2 = new List<SelectListItem>();
                var queryPATH = from p in ctxData.VFDITEMs
                                join fx in ctxData.VFDMODELs on p.MODEL_ID equals fx.ID
                                join e in ctxData.PORTDESCRIPTIONs on p.MODEL_ID equals e.DEVICEMODEL_ID
                                where p.ID == cardID
                                select new { p.AUTODISPLAYNAME, p.ID, fx.NAME, e.NUMBER_, e.LABEL };

                foreach (var a in queryPATH.Distinct().OrderBy(it => it.NUMBER_))
                {
                    listData = listData + a.AUTODISPLAYNAME + ":" + a.ID + ":" + a.NAME + ":" + a.NUMBER_ + ":" + a.LABEL + "|";
                }
            }

            return Json(new
            {
                Success = true,
                listData = listData
            }, JsonRequestBehavior.AllowGet); //
        }

        [HttpPost]
        public ActionResult updataEXC(string PTT_ID) // NIS Network Element
        {
            string PuList = "";
            //string PuList2 = "";
            using (Entities ctxData = new Entities())
            {
                //filter PTT
                List<SelectListItem> list = new List<SelectListItem>();
                var queryEXC = from p in ctxData.WV_EXC_MAST
                               where p.PTT_ID.Trim() == PTT_ID.Trim()
                               orderby p.EXC_ABB
                               select new { p.EXC_ABB };

                foreach (var a in queryEXC)
                {
                    PuList = PuList + a.EXC_ABB + ":  " + a.EXC_ABB + "|";
                }
            }

            return Json(new
            {
                Success = true,
                PuList = PuList
            }, JsonRequestBehavior.AllowGet); //
        }

        [HttpPost]
        public ActionResult InsertLoadPath(string JOBID, string PATHTYPE, string PATHBANDWIDTH, int ANAME, int ZNAME, string PORT, string ASLOT, string ZSLOT, string PRIMARY, string ZEXC, string STATUS,
            string NISBEARERID, string NISMEDIAPATHID, string IPDSLAMNODENAME, string UPLINKPORT, string IPDSLAMDOWNLINKPORT, string MSENODENAME, string MSEDOWNLINKPORT, string ACARDNAME, string ZCARDNAME) // Update list port
        {
            bool success;
            string ANAMEtext;
            string ZNAMEtext;
            string listJob = "";

            using (Entities_NRM ctxData = new Entities_NRM())
            {
                var queryEquipA = (from p in ctxData.VFDITEMs
                                   where p.ID == ANAME
                                   select new { p.AUTODISPLAYNAME, p.ID }).Single();
                ANAMEtext = queryEquipA.AUTODISPLAYNAME;

                var queryEquipZ = (from p in ctxData.VFDITEMs
                                   where p.ID == ZNAME
                                   select new { p.AUTODISPLAYNAME, p.ID , p.USERPARTIALNAME}).Single();

                //if (queryEquipZ.USERPARTIALNAME == null)
                //{
                //    ZNAMEtext = queryEquipZ.AUTODISPLAYNAME;
                //}
                //else
                //{
                    ZNAMEtext = queryEquipZ.AUTODISPLAYNAME;
                //}
            }

            string[] path = PATHTYPE.Split('-');
            string[] exc = JOBID.Split('-');

            WebView.WebService._base myWebService = new WebService._base();
            WebService._base.LoadPath cs = new WebService._base.LoadPath();
            cs.PATHNAME = ANAMEtext + "-" + ZNAMEtext;
            cs.ANAME = ANAMEtext;
            cs.ATYPE = path[0];
            cs.ASITE = exc[0];
            cs.ZNAME = ZNAMEtext;
            cs.ZTYPE = path[1];
            cs.ZSITE = ZEXC;
            cs.PATHTYPE = path[2];
            cs.PATHBANDWIDTH = PATHBANDWIDTH;
            cs.PATHSTATUS = STATUS;
            cs.ASLOT = ASLOT;
            //string ACARD =
            //string APORT =
            cs.ZSLOT = ZSLOT;
            //string ZCARD =
            cs.ZPORT = PORT;
            cs.JOBID = JOBID;
            cs.PRIMARY = PRIMARY;
            cs.NISBEARERID = NISBEARERID;
            cs.NISMEDIAPATHID = NISMEDIAPATHID;
            cs.IPDSLAMNODENAME = IPDSLAMNODENAME;
            cs.UPLINKPORT = UPLINKPORT;
            cs.IPDSLAMDOWNLINKPORT = IPDSLAMDOWNLINKPORT;
            cs.MSENODENAME = MSENODENAME;
            cs.MSEDOWNLINKPORT = MSEDOWNLINKPORT;
            cs.ACARDNAME = ACARDNAME;
            cs.ZCARDNAME = ZCARDNAME;

            success = myWebService.AddLoadPath(cs);
            using (Entities ctxData = new Entities())
            {
                var queryJOB = from p in ctxData.WV_LOAD_PATH
                               where p.JOBID == JOBID
                               select new { p.ANAME, p.ATYPE, p.ASITE, p.ASLOT, p.ACARD, p.APORT, p.ZNAME, p.ZTYPE, p.ZSITE, p.ZSLOT, p.ZCARD, p.ZPORT, p.PRIMARYSECONDARY, p.PATHBANDWIDTH, p.ID };

                foreach (var a in queryJOB.OrderBy(it => it.APORT))
                {
                    listJob = listJob + a.ANAME + ":" + a.ATYPE + ":" + a.ASITE + ":" + a.ASLOT + ":" + a.ACARD + ":" + a.APORT + ":" + a.ZNAME + ":" + a.ZTYPE + ":" + a.ZSITE + ":" + a.ZSLOT + ":" + a.ZCARD + ":" + a.ZPORT + ":" + a.PRIMARYSECONDARY + ":" + a.PATHBANDWIDTH + ":" + a.ID + ":" + "|";
                }
            }
            return Json(new
            {
                Success = true,
                listJob = listJob
            }, JsonRequestBehavior.AllowGet); //
        }

        [HttpPost]
        public ActionResult UpdateLoadPath(string LoadID, string JOBID, string PATHTYPE, string PATHBANDWIDTH, int ANAME, int ZNAME, string PORT, string ASLOT, string ZSLOT, string PRIMARY, string ZEXC, string STATUS,
            string NISBEARERID, string NISMEDIAPATHID, string IPDSLAMNODENAME, string UPLINKPORT, string IPDSLAMDOWNLINKPORT, string MSENODENAME, string MSEDOWNLINKPORT, string ACARDNAME, string ZCARDNAME) // Update list port
        {
            bool success;
            string ANAMEtext;
            string ZNAMEtext;
            string listJob = "";

            using (Entities_NRM ctxData = new Entities_NRM())
            {
                var queryEquipA = (from p in ctxData.VFDITEMs
                                   where p.ID == ANAME
                                   select new { p.AUTODISPLAYNAME, p.ID }).Single();
                ANAMEtext = queryEquipA.AUTODISPLAYNAME;

                var queryEquipZ = (from p in ctxData.VFDITEMs
                                   where p.ID == ZNAME
                                   select new { p.AUTODISPLAYNAME, p.ID, p.USERPARTIALNAME }).Single();

                if (queryEquipZ.USERPARTIALNAME == null)
                {
                    ZNAMEtext = queryEquipZ.AUTODISPLAYNAME;
                }
                else
                {
                    ZNAMEtext = queryEquipZ.USERPARTIALNAME;
                }
            }

            string[] path = PATHTYPE.Split('-');
            string[] exc = JOBID.Split('-');

            WebView.WebService._base myWebService = new WebService._base();
            WebService._base.LoadPath cs = new WebService._base.LoadPath();
            cs.PATHNAME = ANAMEtext + "-" + ZNAMEtext;
            cs.ANAME = ANAMEtext;
            cs.ATYPE = path[0];
            cs.ASITE = exc[0];
            cs.ZNAME = ZNAMEtext;
            cs.ZTYPE = path[1];
            cs.ZSITE = ZEXC;
            cs.PATHTYPE = path[2];
            cs.PATHBANDWIDTH = PATHBANDWIDTH;
            cs.PATHSTATUS = STATUS;
            cs.ASLOT = ASLOT;
            cs.ZSLOT = ZSLOT;
            cs.ZPORT = PORT;
            cs.JOBID = JOBID;
            cs.PRIMARY = PRIMARY;
            cs.NISBEARERID = NISBEARERID;
            cs.NISMEDIAPATHID = NISMEDIAPATHID;
            cs.IPDSLAMNODENAME = IPDSLAMNODENAME;
            cs.UPLINKPORT = UPLINKPORT;
            cs.IPDSLAMDOWNLINKPORT = IPDSLAMDOWNLINKPORT;
            cs.MSENODENAME = MSENODENAME;
            cs.MSEDOWNLINKPORT = MSEDOWNLINKPORT;
            cs.ACARDNAME = ACARDNAME;
            cs.ZCARDNAME = ZCARDNAME;

            success = myWebService.UpdateLoadPath(cs, LoadID);
            using (Entities ctxData = new Entities())
            {
                var queryJOB = from p in ctxData.WV_LOAD_PATH
                               where p.JOBID == JOBID
                               select new { p.ANAME, p.ATYPE, p.ASITE, p.ASLOT, p.ACARD, p.APORT, p.ZNAME, p.ZTYPE, p.ZSITE, p.ZSLOT, p.ZCARD, p.ZPORT, p.PRIMARYSECONDARY, p.PATHBANDWIDTH, p.ID };

                foreach (var a in queryJOB.OrderBy(it => it.APORT))
                {
                    listJob = listJob + a.ANAME + ":" + a.ATYPE + ":" + a.ASITE + ":" + a.ASLOT + ":" + a.ACARD + ":" + a.APORT + ":" + a.ZNAME + ":" + a.ZTYPE + ":" + a.ZSITE + ":" + a.ZSLOT + ":" + a.ZCARD + ":" + a.ZPORT + ":" + a.PRIMARYSECONDARY + ":" + a.PATHBANDWIDTH + ":" + a.ID + ":" + "|";
                }


            }
            return Json(new
            {
                Success = true,
                listJob = listJob
            }, JsonRequestBehavior.AllowGet); //
        }

        [HttpPost]
        public ActionResult GetDetails(string id)
        {
            string equipA;
            string equipZ;

            using (Entities ctxData = new Entities())
            {
                var query = (from p in ctxData.WV_LOAD_PATH
                             where p.ID == id
                             select p).Single();
                equipA = query.ANAME;
                equipZ = query.ZNAME;
            }

            using (Entities_NRM ctxData = new Entities_NRM())
            {
                var queryEquipA = (from p in ctxData.VFDITEMs
                                   where p.AUTODISPLAYNAME == equipA
                                   select p).Single();
                equipA = queryEquipA.ID.ToString();

                var queryEquipZ = (from p in ctxData.VFDITEMs
                                   where p.AUTODISPLAYNAME == equipZ
                                   select p).Single();
                equipZ = queryEquipZ.ID.ToString();
            }

            using (Entities ctxData = new Entities())
            {
                var query = (from p in ctxData.WV_LOAD_PATH
                             where p.ID == id
                             select p).Single();
                query.PATHTYPE = query.ATYPE + "-" + query.ZTYPE + "-" + query.PATHTYPE;

                System.Diagnostics.Debug.WriteLine("FUCK");
                System.Diagnostics.Debug.WriteLine(equipZ);
                return Json(new
                {
                    Success = true,
                    id = id,
                    job = query.JOBID,
                    zexc = query.ZSITE,
                    pathType = query.PATHTYPE,
                    equipA = equipA,
                    equipZ = equipZ,
                    bandwidth = query.PATHBANDWIDTH,
                    primary = query.PRIMARYSECONDARY,
                    status = query.PATHSTATUS,
                    ipnode = query.IPDSLAMNODENAME,
                    ipuplink = query.UPLINKPORT,
                    ipdownlink = query.IPDSLAMDOWNLINKPORT,
                    bearerid = query.NISBEARERID,
                    bearerpath = query.NISMEDIAPATHID,
                    msenode = query.MSENODENAME,
                    msedownlink = query.MSEDOWNLINKPORT,
                    acardname = query.ACARDNAME,
                    zcardname = query.ZCARDNAME,
                    aport = query.APORT,
                    zport = query.ZPORT
                    //checkfeature = checkfeature,
                }, JsonRequestBehavior.AllowGet); //
            }
        }

        [HttpPost]
        public ActionResult DeleteDetails(string id)
        {
            bool success;

            WebView.WebService._base myWebService = new WebService._base();
            WebService._base.LoadPath cs = new WebService._base.LoadPath();
            success = myWebService.DeleteLoadPath(id);
            return Json(new
            {
                Success = true
            }, JsonRequestBehavior.AllowGet); //
        }

        [HttpPost]
        public ActionResult EmailTest(string JOBID) // NIS Network Element
        {
            string result;
            try 
            {
                MailMessage msg = new MailMessage();
                msg.IsBodyHtml = true;
                msg.From = new MailAddress("neps@tm.com.my", "NEPS");
                msg.To.Add("tay.engsoon@tm.com.my");
                msg.Subject = "EMAIL FROM LOCAL " ;
                msg.Body = "<h1>FILES DETAILS</h1>SCHEME NAME	: <br/><br/>DESCRIPTION : <br/><br/>REDMARK FILE NAME: .xml ";
                msg.Body += "<br/><br/> <h1>RNO DETAILS</h1> <br>";
                msg.Body += "RNO ID : " + User.Identity.Name + "<br/><br/>RNO EMAIL	: <br/><br/>RNO PHONE NUMBER: <br/><br/>Please log in to <a href='http://10.41.101.168/'>NEPS WEBVIEW  </a>to download the file.";
                msg.IsBodyHtml = true;
                SmtpClient emailClient = new SmtpClient("smtp.tm.com.my", 25);
                emailClient.UseDefaultCredentials = false;
                emailClient.Credentials = new NetworkCredential("neps", "nepsadmin", "tmmaster");
                emailClient.EnableSsl = false;
                emailClient.DeliveryMethod = SmtpDeliveryMethod.Network;
                emailClient.Send(msg);
                result = "OK";
            }
            catch (Exception ex)
            {
                result = ex.ToString();
            }

            return Json(new
            {
                Success = true,
                result = result
            }, JsonRequestBehavior.AllowGet); //
        }

        //--------------------------------------------------------------------FTTH PATH

        public ActionResult NewFTTHPath(string searchKey, int? page)
        {
            using (Entities ctxData = new Entities())
            {
                //filter PTT
                List<SelectListItem> list = new List<SelectListItem>();
                var queryEXC = from p in ctxData.WV_EXC_MAST
                               select new { Text = p.PTT_ID, Value = p.PTT_ID };

                list.Add(new SelectListItem() { Text = "", Value = "Select" });
                foreach (var a in queryEXC.Distinct().OrderBy(it => it.Value))
                {
                    if (a.Value != null)
                        list.Add(new SelectListItem() { Text = a.Text.ToString(), Value = a.Value.ToString() });
                }
                ViewBag.PTT = list;

                List<SelectListItem> list1 = new List<SelectListItem>();
                var querySCHEME = from p in ctxData.G3E_JOB
                                  orderby p.G3E_IDENTIFIER
                                  where (p.G3E_OWNER.ToUpper() == User.Identity.Name.ToUpper() || p.G3E_OWNER.ToLower() == User.Identity.Name.ToLower()) && (p.JOB_TYPE.Contains("HSBB"))
                                  select new { Text = p.G3E_IDENTIFIER, Value = p.G3E_IDENTIFIER };

                list1.Add(new SelectListItem() { Text = "", Value = "Select" });
                foreach (var a in querySCHEME.Distinct().OrderBy(it => it.Value))
                {
                    if (a.Value != null)
                        list1.Add(new SelectListItem() { Text = a.Text.ToString(), Value = a.Value.ToString() });
                }

                ViewBag.job = list1;

                List<SelectListItem> list2 = new List<SelectListItem>();
                list2.Add(new SelectListItem() { Text = "", Value = "Select" });
                ViewBag.pathType = list2;
                ViewBag.equipA = list2;
                ViewBag.equipZ = list2;
                ViewBag.cardA = list2;
                ViewBag.cardZ = list2;
                ViewBag.exc = list2;
                ViewBag.dfdp = list2;
            }
            return View();
        }

        
        [HttpPost]
        public ActionResult Updatedfdp(string job)
        {
            string listJob = "";
            string[] jobSplit = job.Split('-');
            string exc = jobSplit[0];

            string listData = "";
            using (Entities6 ctxData = new Entities6())
            {
                List<SelectListItem> list1 = new List<SelectListItem>();
                var queryC888dfdp = from p in ctxData.GC_DFDP
                                    join fx in ctxData.GC_NETELEM_M6 on p.G3E_FID equals fx.G3E_FID
                                    where fx.EXC_ABB == exc
                                    select new { p.FDP_CODE, p.G3E_FID };

                foreach (var a in queryC888dfdp.Distinct().OrderBy(it => it.FDP_CODE))
                {
                    listData = listData + a.FDP_CODE + ":" + a.G3E_FID + "|";
                }
            }

            return Json(new
            {
                Success = true,
                listData = listData,
            }, JsonRequestBehavior.AllowGet);
        }

        
        [HttpPost]
        public ActionResult updateEquipFTTHA(string job, string pathType, string ZEXC) // Update card
        {
            string listJob = "";
            string[] jobSplit = job.Split('-');
            string exc = jobSplit[0];
            //string[] pathTypeSplit1 = pathType.Split('-');
            //string pathType1 = pathTypeSplit1[0];
            //string pathType2 = pathTypeSplit1[1];
            //System.Diagnostics.Debug.WriteLine("[" + pathType1 + "]");
            
            string listData1 = "";
            string listData2 = "";

            string condition3 = ZEXC + "_G";
            
            System.Diagnostics.Debug.WriteLine("[" + ZEXC + "_" + condition3 + "]");

            using (Entities ctxData = new Entities())
            {
                List<SelectListItem> list1 = new List<SelectListItem>();
                var queryEquipA = from p in ctxData.GC_FDC
                                  join fx in ctxData.GC_NETELEM on p.G3E_FID equals fx.G3E_FID
                                  where fx.EXC_ABB == exc
                                  select new { p.FDC_CODE, p.G3E_FID };

                foreach (var a in queryEquipA.Distinct().OrderBy(it => it.FDC_CODE))
                {
                   
                        listData1 = listData1 + a.FDC_CODE + ":" + a.G3E_FID + "|";
                }
            }
            using (Entities_NRM ctxData = new Entities_NRM())
            {
                List<SelectListItem> list2 = new List<SelectListItem>();
                var queryEquipZ = from p in ctxData.VFDITEMs
                                  join fx in ctxData.VFDITEMPLACEMENTs on p.ID equals fx.ITEM_ID
                                  join e in ctxData.VFDITEMs on fx.CONTAINER_ID equals e.ID
                                  join fxx in ctxData.VFDITEMPLACEMENTs on e.ID equals fxx.CONTAINER_ID
                                  where (e.DTYPE == "Rack") &&
                                        (p.USERDISPLAYNAME.Contains(condition3) || p.AUTODISPLAYNAME.Contains(condition3))
                                  select new { p.USERDISPLAYNAME, p.AUTODISPLAYNAME, p.ID };

                foreach (var a in queryEquipZ.Distinct().OrderBy(it => it.USERDISPLAYNAME))
                {
                    string data;
                    if (a.USERDISPLAYNAME == null)
                    {
                        data = a.AUTODISPLAYNAME;
                    }
                    else
                    {
                        data = a.USERDISPLAYNAME;
                    }
                    listData2 = listData2 + data + ":" + a.ID + "|";
                }
            }

            // region Mubin - CR14-20180330
			using (Entities9 ctxData = new Entities9())
            {
                var queryJOB = from p in ctxData.WV_LOAD_PATHCONSUMER
                               where p.JOBID == job
                               select new { p.ANAME, p.ATYPE, p.ASITE, p.ACARD2, p.APORT2, p.ZNAME, p.ZTYPE, p.ZSITE, p.ZCARD, p.ZPORT, p.DPNAME, p.ID,p.REMARKS};

                foreach (var a in queryJOB.OrderBy(it => it.DPNAME))
                {
                    string St_REMARKS = "";
                    if (a.REMARKS != "" && a.REMARKS != null)
                    {
                        St_REMARKS = " (Remarks -  " + a.REMARKS + " ) ";
                    }
                    listJob = listJob + a.ANAME + ":" + a.ATYPE + ":" + a.ASITE + ":" + a.ACARD2 + ":" + a.APORT2 + ":" + a.ZNAME + ":" + a.ZTYPE + ":" + a.ZSITE + ":" + a.ZCARD + ":" + a.ZPORT + ":" + a.DPNAME + St_REMARKS + ":" + a.ID + ":" + "|";
                }

               
            }
            /*using (Entities7 ctxData = new Entities7())
            {
                var queryJOB_ODP = from p in ctxData.WV_LOAD_PATH_CONSUMER_ODP
                                   where p.JOBID == job
                                   select new { p.ANAME, p.ATYPE, p.ASITE, p.ACARD2, p.APORT2, p.ZNAME, p.ZTYPE, p.ZSITE, p.ZCARD, p.ZPORT, p.DPNAME, p.ID,p.REMARKS };

                foreach (var a in queryJOB_ODP.OrderBy(it => it.DPNAME))
                {
                    string St_REMARKS = "";
                    if (a.REMARKS != "" && a.REMARKS != null)
                    {
                        St_REMARKS = " (Remarks -  " + a.REMARKS + " ) ";
                    }

                    listJob = listJob + a.ANAME + ":" + a.ATYPE + ":" + a.ASITE + ":" + a.ACARD2 + ":" + a.APORT2 + ":" + a.ZNAME + ":" + a.ZTYPE + ":" + a.ZSITE + ":" + a.ZCARD + ":" + a.ZPORT + ":" + a.DPNAME +St_REMARKS + ":" + a.ID + ":" + "|";
                }
            }*/
			// endRegion

            return Json(new
            {
                Success = true,
                listData = listData1,
                listData2 = listData2,
                listJob = listJob
            }, JsonRequestBehavior.AllowGet); //
        }

        [HttpPost]
        public ActionResult updateCardFTTHA(int equip, string fdccode) // Update port
        {
            string listData = "";
            bool result;

            WebView.WebService._base myWebService;
            myWebService = new WebService._base();

            //result = myWebService.FindFDC(equip);
            using (Entities ctxData = new Entities())
            {
                List<SelectListItem> list2 = new List<SelectListItem>();

                //var queryParent = (from p in ctxData.GC_FDP
                //                   join q in ctxData.GC_OWNERSHIP on p.G3E_FID equals q.G3E_FID
                //                   join r in ctxData.GC_OWNERSHIP on q.G3E_ID equals r.OWNER1_ID
                //                   join sc in ctxData.GC_SPLICE_CONNECT on p.FDC_FID equals sc.FID1
                //                   join spl in ctxData.GC_FSPLITTER on sc.FID2 equals spl.G3E_FID
                //                   where p.FDC_FID == equip //FDC FID
                //                   //&& (sc.FNO1 == 7200 || sc.FNO1 == 4400 || sc.FNO1 == 4500) || 
                //                   //   (sc.FNO2 == 7200 || sc.FNO1 == 4400 || sc.FNO1 == 4500)
                //                   orderby sc.FID2, sc.LOW2
                //                   select new {  sc.G3E_FID, sc.LOW1, sc.HIGH1, p.FDP_CODE, spl.SPLITTER_CODE, spl.SPLITTER_TYPE}); 

                // D-SIDE (FDP)

                if (fdccode.Contains("C999"))
                {
                    var cap = from p in ctxData.GC_FDP
                              join s in ctxData.GC_SPLICE_CONNECT on p.G3E_FID equals s.G3E_FID
                              join spl in ctxData.GC_FSPLITTER on s.FID2 equals spl.G3E_FID
                              where p.G3E_FID == equip
                              select new { spl.G3E_FID, spl.SPLITTER_CODE, spl.SPLITTER_TYPE };
                    int i = 1;
                    foreach (var a in cap.Distinct().OrderBy(it => it.SPLITTER_CODE))
                    {
                        //listData = listData + a.FDP_CODE + "," + a.LOW1 + "," + a.HIGH1 + "," + a.SPLITTER_CODE + "-" + a.SPLITTER_TYPE + "|";
                        listData = listData + a.SPLITTER_CODE + ",Splitter "+a.G3E_FID+" -" + i + " (" + a.SPLITTER_TYPE + ")|";
                        i++;
                    }
                    //string[] capacity = cap.FDP_CAPACITY.Split(':');
                    //string[] capacity2 = cap.FDP_CAPACITY.Split('=');

                    //int num_cap = Convert.ToInt32(capacity[1]);
                    //for (int i = 0; i < num_cap; i++)
                    //{
                    //    listData = listData + (i + 1) +",(" + capacity2[0] + ")|";
                    //}
                }
                else
                {
                    var queryParent = from p in ctxData.GC_FDC
                                      join s in ctxData.GC_SPLICE_CONNECT on p.G3E_FID equals s.G3E_FID
                                      join spl in ctxData.GC_FSPLITTER on s.FID1 equals spl.G3E_FID
                                      join u in ctxData.GC_NETELEM on s.G3E_FID equals u.G3E_FID
                                      where p.G3E_FID == equip
                                      select new { spl.G3E_FID, spl.SPLITTER_CODE, spl.SPLITTER_TYPE };

                    foreach (var a in queryParent.Distinct().OrderBy(it => it.SPLITTER_CODE))
                    {
                        //listData = listData + a.FDP_CODE + "," + a.LOW1 + "," + a.HIGH1 + "," + a.SPLITTER_CODE + "-" + a.SPLITTER_TYPE + "|";
                        listData = listData + a.G3E_FID + "," + a.SPLITTER_CODE + "(" + a.SPLITTER_TYPE + ")|";
                    }
                }
                System.Diagnostics.Debug.WriteLine(listData);
            }

            return Json(new
            {
                Success = true,
                listData = listData
            }, JsonRequestBehavior.AllowGet); //
        }

        [HttpPost]
        public ActionResult updateCardFTTHZ(int equip, string scode) // Update port
        {
            string listData = "";
            string listData1 = "";
            string[] xcode = scode.Split('-');
            string code = xcode[1];
            string listDataOdf = "";
            using (Entities_NRM ctxData = new Entities_NRM())
            {
                List<SelectListItem> list2 = new List<SelectListItem>();
                var queryParent = from p in ctxData.VFDITEMs
                                  join fx in ctxData.VFDITEMPLACEMENTs on p.ID equals fx.ITEM_ID
                                  //join e in ctxData.VFDITEMs on fx.CONTAINER_ID equals e.ID
                                  //join fxx in ctxData.VFDITEMPLACEMENTs on e.ID equals fxx.ITEM_ID
                                  where fx.CONTAINER_ID == equip
                                  select new { p.AUTODISPLAYNAME, p.ID };

                foreach (var a in queryParent.Distinct().OrderBy(it => it.AUTODISPLAYNAME))
                {
                    listData = listData + a.AUTODISPLAYNAME + ":" + a.ID + "|";
                }
                //sub card
                var queryPATH = from p in ctxData.VFDITEMs
                                join fx in ctxData.VFDITEMPLACEMENTs on p.ID equals fx.ITEM_ID
                                join e in ctxData.VFDITEMs on fx.CONTAINER_ID equals e.ID
                                join fxx in ctxData.VFDITEMPLACEMENTs on e.ID equals fxx.ITEM_ID
                                where fxx.CONTAINER_ID == equip
                                select new { p.AUTODISPLAYNAME, p.ID };

                foreach (var a in queryPATH.Distinct().OrderBy(it => it.AUTODISPLAYNAME))
                {
                    listData = listData + a.AUTODISPLAYNAME + ":" + a.ID + "|";
                }

                //ODF
                string[] dataExc = code.Split('_');
                string Exc = dataExc[0];
                var queryODF = from p in ctxData.VFDITEMs
                               join fx in ctxData.VFDITEMPLACEMENTs on p.ID equals fx.ITEM_ID
                               where (p.USERDISPLAYNAME.Contains(Exc) || p.AUTODISPLAYNAME.Contains(Exc)) && (p.USERDISPLAYNAME.Contains("ODF") || p.AUTODISPLAYNAME.Contains("ODF")) && p.DTYPE == "Rack"
                               select new { p.USERDISPLAYNAME, p.AUTODISPLAYNAME, p.ID };

                foreach (var a in queryODF.Distinct().OrderBy(it => it.USERDISPLAYNAME))
                {
                    var dataDisplay = "";
                    if (a.USERDISPLAYNAME == null)
                    {
                        dataDisplay = a.AUTODISPLAYNAME;
                    }
                    else
                    {
                        dataDisplay = a.USERDISPLAYNAME;
                    }
                    listDataOdf = listDataOdf + dataDisplay + ":" + a.ID + "|";
                }
            }

            //string lastcode = code.Substring(code.Length - 2, 2);
            //string newcode = "G" + lastcode;
            string[] codeSplit = code.Split('_');
            string exc = codeSplit[0];
            string oltcode = codeSplit[1];
            string last_code = oltcode.Substring(oltcode.Length - 2, 2);

            System.Diagnostics.Debug.WriteLine("exc" + exc);
            System.Diagnostics.Debug.WriteLine("olt" + last_code);
            using (Entities ctxData = new Entities())
            {
                List<SelectListItem> list1 = new List<SelectListItem>();
                var queryEquipA = from p in ctxData.GC_FDC
                                  join fx in ctxData.GC_NETELEM on p.G3E_FID equals fx.G3E_FID
                                  where p.OLT_ID.Contains(last_code) && fx.EXC_ABB.Trim() == exc
                                  select new { p.FDC_CODE, p.G3E_FID };

                foreach (var a in queryEquipA.Distinct().OrderBy(it => it.FDC_CODE))
                {
                    if (a.FDC_CODE.Contains("C999"))
                    {
                        var queryFDP = from p in ctxData.GC_FDP
                                       join fx in ctxData.GC_NETELEM on p.G3E_FID equals fx.G3E_FID
                                       where fx.EXC_ABB == exc && p.FDC_CODE == a.FDC_CODE
                                       select new { p.FDP_CODE, p.G3E_FID };

                        foreach (var b in queryFDP.Distinct().OrderBy(it => it.FDP_CODE))
                        {
                            listData1 = listData1 + a.FDC_CODE + "_DP" + b.FDP_CODE + ":" + b.G3E_FID + "|";
                        }
                    }
                    else
                    {
                        listData1 = listData1 + a.FDC_CODE + ":" + a.G3E_FID + "|";
                    }
                }
            }
            // Noraini Ali
            using (Entities6 ctxData = new Entities6())
            {
                List<SelectListItem> list1 = new List<SelectListItem>();
                var queryCheckdfdp = (from p in ctxData.GC_DFDP
                                      join fx in ctxData.GC_NETELEM_M6 on p.G3E_FID equals fx.G3E_FID
                                      where fx.EXC_ABB == exc
                                      select p);
                if (queryCheckdfdp.Count() > 0)
                {
                    listData1 = listData1 + "C888:0|";
                }
            }
            System.Diagnostics.Debug.WriteLine("listData1 = " + listData1);
            return Json(new
            {
                Success = true,
                listData = listData,
                listData1 = listData1,
                last_code = last_code,
                listDataOdf = listDataOdf
            }, JsonRequestBehavior.AllowGet); //
        }

        [HttpPost]
        public ActionResult updatePatchPanelODF(int equip) // Update port
        {
            string listDataPP = "";

            using (Entities_NRM ctxData = new Entities_NRM())
            {
                //PatchPanel
                var queryPP = from p in ctxData.VFDITEMs
                              join fx in ctxData.VFDITEMPLACEMENTs on p.ID equals fx.ITEM_ID
                              join e in ctxData.VFDITEMs on fx.CONTAINER_ID equals e.ID
                              join fxx in ctxData.VFDITEMPLACEMENTs on e.ID equals fxx.ITEM_ID
                              where e.ID == equip
                              select new { p.AUTODISPLAYNAME, p.ID };

                foreach (var a in queryPP.Distinct().OrderBy(it => it.AUTODISPLAYNAME))
                {
                    listDataPP = listDataPP + a.AUTODISPLAYNAME + ":" + a.ID + "|";
                }
            }

            return Json(new
            {
                Success = true,
                listDataPP = listDataPP
            }, JsonRequestBehavior.AllowGet); //
        }

        [HttpPost]
        public ActionResult updatePatchPanelPort(int equip) // Update port
        {
            string listDataPP = "";

            using (Entities_NRM ctxData = new Entities_NRM())
            {
                //Port PatchPanel
                var queryPP = from p in ctxData.JUNCTIONs
                              join fx in ctxData.VFDITEMs on p.JUNCTIONARRAYMODEL_ID equals fx.MODEL_ID
                              where fx.ID == equip
                              select new { p.ROW_, p.COLUMN_, p.ID };

                var queryMaxPP = (from p in ctxData.JUNCTIONs
                                  join fx in ctxData.VFDITEMs on p.JUNCTIONARRAYMODEL_ID equals fx.MODEL_ID
                                  where fx.ID == equip
                                  select new { p.COLUMN_ }).Max(x => x.COLUMN_);

                System.Diagnostics.Debug.WriteLine(queryMaxPP);

                int val = Convert.ToInt32(queryMaxPP) + 1;

                foreach (var a in queryPP.Distinct().OrderBy(it => it.ROW_).ThenBy(it => it.COLUMN_))
                {
                    //int val = runNo + 1;
                    int data = (Convert.ToInt32(a.ROW_) * val) + (Convert.ToInt32(a.COLUMN_) + 1);
                    listDataPP = listDataPP + data + ":" + data + "|";
                    //val++;
                }
            }

            return Json(new
            {
                Success = true,
                listDataPP = listDataPP
            }, JsonRequestBehavior.AllowGet); //
        }

        [HttpPost]
        public ActionResult updatePortFTTHA(int equip, string job) // Update port
        {
            string listData = "";
            bool result;

            WebView.WebService._base myWebService;
            myWebService = new WebService._base();
            //result = myWebService.FindFDC(equip);
            using (Entities ctxData = new Entities())
            {
                var queryDB = (from p in ctxData.GC_DB
                               join sc in ctxData.GC_SPLICE_CONNECT on p.G3E_FID equals sc.G3E_FID
                               join spl in ctxData.GC_FSPLITTER on sc.SRC_FID equals spl.G3E_FID
                               join q in ctxData.GC_OWNERSHIP on sc.SRC_FID equals q.G3E_FID
                               join r in ctxData.GC_OWNERSHIP on q.OWNER1_ID equals r.G3E_ID
                               join s in ctxData.GC_NETELEM on p.G3E_FID equals s.G3E_FID
                               where spl.G3E_FID == equip && s.JOB_ID == job
                               orderby sc.FID2, sc.LOW2
                               select new { sc.LOW1, sc.HIGH1, p.DB_CODE, spl.SPLITTER_CODE, spl.SPLITTER_TYPE, spl.NO_OF_SPLITTER }).Distinct();
                foreach (var a in queryDB.Distinct().OrderBy(it => it.DB_CODE))
                {
                    listData = listData + "DB" + a.DB_CODE + "," + a.LOW1 + "," + a.HIGH1 + ", (" + a.SPLITTER_CODE + "/" + a.SPLITTER_TYPE + ")," + a.NO_OF_SPLITTER + ",|";
                }

                var queryParent = (from p in ctxData.GC_FDP
                                   join sc in ctxData.GC_SPLICE_CONNECT on p.G3E_FID equals sc.G3E_FID
                                   join spl in ctxData.GC_FSPLITTER on sc.SRC_FID equals spl.G3E_FID
                                   join q in ctxData.GC_OWNERSHIP on sc.SRC_FID equals q.G3E_FID
                                   join r in ctxData.GC_OWNERSHIP on q.OWNER1_ID equals r.G3E_ID
                                   join s in ctxData.GC_NETELEM on p.G3E_FID equals s.G3E_FID
                                   where spl.G3E_FID == equip && s.JOB_ID == job //FDC FID
                                   orderby sc.FID2, sc.LOW2
                                   select new { sc.LOW1, sc.HIGH1, p.FDP_CODE, spl.SPLITTER_CODE, spl.SPLITTER_TYPE, spl.NO_OF_SPLITTER }).Distinct();

                foreach (var a in queryParent.Distinct().OrderBy(it => it.FDP_CODE))
                {
                    listData = listData + "DP" + a.FDP_CODE + "," + a.LOW1 + "," + a.HIGH1 + ", (" + a.SPLITTER_CODE + "/" + a.SPLITTER_TYPE + ")," + a.NO_OF_SPLITTER + "|";
                }
                System.Diagnostics.Debug.WriteLine(listData);
            }

            return Json(new
            {
                Success = true,
                listData = listData
            }, JsonRequestBehavior.AllowGet); //
        }

        [HttpPost]
        public ActionResult updatePortFTTHZ(string cardA) // Update list port
        {
            string listData = "";
            string[] card = cardA.Split('-');
            string testc = card[0];
            int cardID = Convert.ToInt32(testc);
            using (Entities_NRM ctxData = new Entities_NRM())
            {
                List<SelectListItem> list2 = new List<SelectListItem>();
                var queryPATH = (from p in ctxData.VFDITEMs
                                 join fx in ctxData.VFDMODELs on p.MODEL_ID equals fx.ID
                                 join e in ctxData.PORTDESCRIPTIONs on p.MODEL_ID equals e.DEVICEMODEL_ID
                                 where p.ID == cardID
                                 select new { p.AUTODISPLAYNAME, p.ID, fx.NAME, e.NUMBER_, e.LABEL }).Single();

                //foreach (var a in queryPATH.Distinct().OrderBy(it => it.NUMBER_))
                //{
                listData = listData + queryPATH.AUTODISPLAYNAME + ":" + queryPATH.ID + ":" + queryPATH.NAME + ":" + queryPATH.NUMBER_ + ":" + queryPATH.LABEL + "|";
                //}
            }

            return Json(new
            {
                Success = true,
                listData = listData
            }, JsonRequestBehavior.AllowGet); //
        }

        // region Mubin - CR14-20180330
		[HttpPost]
        public ActionResult InsertLoadFTTHPath(string JOBID, string ANAME, string ATYPE, int FDC_FID, string ASTER, string ZNAME, string ZSITE, string ZTYPE, string ZSLOT, string ZPORT, string A2Port, string DPPORT, string DP_TAMBAHAN, string InPort, string OutPort, string dfdp, string remarks) // Update list port
        {
            bool success;
            string[] splitS = ASTER.Split('-');
            string splitt = splitS[1];
            string[] splitter = ASTER.Split('(');
            string listJob = "";

            WebView.WebService._base myWebService = new WebService._base();
            WebService._base.LoadPath cs = new WebService._base.LoadPath();
            cs.ANAME = ANAME; // FDC CODE
            cs.ATYPE = ATYPE; // FDC
            cs.ASITE = ZSITE; // exc FDC
            cs.ACARD = splitter[0]; // splitter will be get port and card
            cs.ZNAME = ZNAME; // OLT name
            cs.ZTYPE = ZTYPE; // OLT
            cs.ZSITE = ZSITE; // exc OLT
            cs.ZCARD = ZSLOT; // OLT slot
            cs.ZPORT = ZPORT; // OLT port
            cs.FDC_FID = FDC_FID; // FDC_FID
            cs.JOBID = JOBID;

            success = myWebService.AddLoadFTTHPath(cs, A2Port, DPPORT, DP_TAMBAHAN, InPort, OutPort, dfdp, remarks, ASTER);

            using (Entities9 ctxData = new Entities9())
            {
                var queryJOB = from p in ctxData.WV_LOAD_PATHCONSUMER
                               where p.JOBID == JOBID
                               select new { p.ANAME, p.ATYPE, p.ASITE, p.ACARD3, p.APORT2, p.ZNAME, p.ZTYPE, p.ZSITE, p.ZCARD, p.ZPORT, p.DPNAME, p.ID, p.DPPORT,p.REMARKS };

                foreach (var a in queryJOB.OrderBy(it => it.DPNAME))
                {
                    string St_REMARKS = "";
                    if (a.REMARKS != "" && a.REMARKS != null)
                    {
                        St_REMARKS = " (Remarks -  " + a.REMARKS + " ) ";
                    }
                    listJob = listJob + a.ANAME + ":" + a.ATYPE + ":" + a.ASITE + ":" + a.ACARD3 + ":" + a.APORT2 + ":" + a.DPPORT + ":" + a.ZNAME + ":" + a.ZTYPE + ":" + a.ZSITE + ":" + a.ZCARD + ":" + a.ZPORT + ":" + a.DPNAME + St_REMARKS + ":" + a.ID + ":" + "|";
                }
            }

            /*using (Entities7 ctxData = new Entities7())
            {
                var queryJOB_ODP = from p in ctxData.WV_LOAD_PATH_CONSUMER_ODP
                                   where p.JOBID == JOBID
                                   select new { p.ANAME, p.ATYPE, p.ASITE, p.ACARD3, p.APORT2, p.ZNAME, p.ZTYPE, p.ZSITE, p.ZCARD, p.ZPORT, p.DPNAME, p.ID, p.DPPORT,p.REMARKS };

                foreach (var a in queryJOB_ODP.OrderBy(it => it.DPNAME))
                {
                    string St_REMARKS = "";
                    if (a.REMARKS != "" && a.REMARKS != null)
                    {
                        St_REMARKS = " (Remarks -  " + a.REMARKS + " ) ";
                    }

                    listJob = listJob + a.ANAME + ":" + a.ATYPE + ":" + a.ASITE + ":" + a.ACARD3 + ":" + a.APORT2 + ":" + a.DPPORT + ":" + a.ZNAME + ":" + a.ZTYPE + ":" + a.ZSITE + ":" + a.ZCARD + ":" + a.ZPORT + ":" + a.DPNAME + St_REMARKS + ":" + a.ID + ":" + "|";
                }
            }*/
            return Json(new
            {
                Success = success,
                listJob = listJob
            }, JsonRequestBehavior.AllowGet); //
        }
		// endRegion

        //[HttpPost]
        //public ActionResult UpdateLoadPathFTTH(string JOBID, string ANAME, string ATYPE, int FDC_FID, string ASTER, string ZNAME, string ZSITE, string ZTYPE, string ZSLOT, string ZPORT) // Update list port
        //{
        //    bool success;
        //    string[] splitter = ASTER.Split('(');
        //    string listJob = "";

        //    WebView.WebService._base myWebService = new WebService._base();
        //    WebService._base.LoadPath cs = new WebService._base.LoadPath();
        //    cs.ANAME = ANAME; // FDC CODE
        //    cs.ATYPE = ATYPE; // FDC
        //    cs.ASITE = ZSITE; // exc FDC
        //    cs.ACARD = splitter[0]; // splitter will be get port and card
        //    cs.ZNAME = ZNAME; // OLT name
        //    cs.ZTYPE = ZTYPE; // OLT
        //    cs.ZSITE = ZSITE; // exc OLT
        //    cs.ZCARD = ZSLOT; // OLT slot
        //    cs.ZPORT = ZPORT; // OLT port
        //    cs.FDC_FID = FDC_FID; // FDC_FID
        //    cs.JOBID = JOBID;

        //    success = myWebService.AddLoadFTTHPath(cs);

        //    using (Entities ctxData = new Entities())
        //    {

        //        var queryJOB = from p in ctxData.WV_LOAD_PATH_CONSUMER
        //                       where p.JOBID == JOBID
        //                       select new { p.ANAME, p.ATYPE, p.ASITE, p.ACARD2, p.APORT2, p.ZNAME, p.ZTYPE, p.ZSITE, p.ZCARD, p.ZPORT, p.DPNAME, p.ID };

        //        foreach (var a in queryJOB.OrderBy(it => it.APORT2))
        //        {
        //            listJob = listJob + a.ANAME + ":" + a.ATYPE + ":" + a.ASITE + ":" + a.ACARD2 + ":" + a.APORT2 + ":" + a.ZNAME + ":" + a.ZTYPE + ":" + a.ZSITE + ":" + a.ZCARD + ":" + a.ZPORT + ":" + a.DPNAME + ":" + a.ID + ":" + "|";
        //        }
        //    }
        //    return Json(new
        //    {
        //        Success = true,
        //        listJob = listJob
        //    }, JsonRequestBehavior.AllowGet); //
        //}

        [HttpPost]
        public ActionResult DeleteDetailsFTTH(string id)
        {
            bool success;
            string[] arr = id.Split(':');
            WebView.WebService._base myWebService = new WebService._base();
            WebService._base.LoadPath cs = new WebService._base.LoadPath();
            success = myWebService.DeleteLoadPathFTTH(arr[0],arr[1]);
            return Json(new
            {
                Success = true
            }, JsonRequestBehavior.AllowGet); //
        }

        public ActionResult ManageConnectivity(string searchKey, int? page)
        {
            using (Entities ctxData = new Entities())
            {
                //filter PTT
                List<SelectListItem> list = new List<SelectListItem>();
                var queryEXC = from p in ctxData.WV_EXC_MAST
                               select new { Text = p.PTT_ID, Value = p.PTT_ID };

                list.Add(new SelectListItem() { Text = "", Value = "Select" });
                foreach (var a in queryEXC.Distinct().OrderBy(it => it.Value))
                {
                    if (a.Value != null)
                        list.Add(new SelectListItem() { Text = a.Text.ToString(), Value = a.Value.ToString() });
                }
                ViewBag.PTT = list;

                List<SelectListItem> list2 = new List<SelectListItem>();
                ViewBag.exc = list2;
                ViewBag.elv = list2;
                ViewBag.equipA = list2;
                ViewBag.equipZ = list2;
                ViewBag.cardA = list2;
                ViewBag.cardZ = list2;
            }
            return View();
        }

        [HttpPost]
        public ActionResult findELV(string EXC) // Find Elevation
        {
            string listData = "";
            using (Entities_NRM ctxData = new Entities_NRM())
            {
                List<SelectListItem> list2 = new List<SelectListItem>();
                var queryEquipZ = from fx in ctxData.ELEVATIONs
                                  join p in ctxData.VFDITEMs on fx.PLAN_ID equals p.ID
                                  where (p.DTYPE == "Plan") &&
                                        p.USERDISPLAYNAME.Contains(EXC)
                                  select new { fx.NAME, fx.Z };

                foreach (var a in queryEquipZ.OrderBy(it => it.NAME))
                {
                    listData = listData + a.NAME + ":" + a.Z + "|";
                }
            }

            return Json(new
            {
                Success = true,
                listData = listData
            }, JsonRequestBehavior.AllowGet); //
        }

        [HttpPost]
        public ActionResult findChasiss(string EXC, int lvl) // Find Chasiss
        {
            string listData = "";
            using (Entities_NRM ctxData = new Entities_NRM())
            {
                List<SelectListItem> list2 = new List<SelectListItem>();
                var queryEquipZ = from p in ctxData.VFDITEMs
                                  join fx in ctxData.VFDITEMPLACEMENTs on p.ID equals fx.ITEM_ID
                                  //join e in ctxData.VFDITEMs on fx.CONTAINER_ID equals e.ID
                                  //join fxx in ctxData.VFDITEMPLACEMENTs on e.ID equals fxx.CONTAINER_ID
                                  where (p.DTYPE == "Chassis") &&
                                        p.USERDISPLAYNAME.Contains(EXC) && fx.Z >= lvl
                                  select new { p.USERDISPLAYNAME, p.ID };

                foreach (var a in queryEquipZ.Distinct().OrderBy(it => it.USERDISPLAYNAME))
                {
                    listData = listData + a.USERDISPLAYNAME + ":" + a.ID + "|";
                }
            }

            return Json(new
            {
                Success = true,
                listData = listData
            }, JsonRequestBehavior.AllowGet); //
        }

        [HttpPost]
        public ActionResult findChasissPort(int id) // Find Chasiss
        {
            string listData = "";
            using (Entities_NRM ctxData = new Entities_NRM())
            {
                List<SelectListItem> list2 = new List<SelectListItem>();
                var queryEquipZ = from p in ctxData.VFDITEMs
                                  join fx in ctxData.PORTDESCRIPTIONs on p.MODEL_ID equals fx.DEVICEMODEL_ID
                                  where p.ID == id
                                  select new { p.ID, fx.NUMBER_ };

                foreach (var a in queryEquipZ.Distinct().OrderBy(it => it.NUMBER_))
                {
                    listData = listData + a.NUMBER_ + ":" + a.ID + "|";
                }
            }

            return Json(new
            {
                Success = true,
                listData = listData
            }, JsonRequestBehavior.AllowGet); //
        }

        class MyStub
        {
            public string name { get; set; }
            public int level { get; set; }
            public int categoryId { get; set; }
            public int parentId { get; set; }
        }

        [HttpPost]
        public ActionResult findCard(int id, string A_SUB_TYPE) // Find Chasiss Card
        {
            string listData = "";
            using (Entities_NRM ctxData = new Entities_NRM())
            {
                List<SelectListItem> list2 = new List<SelectListItem>();
                var queryEquipZ = from p in ctxData.VFDITEMs
                                  join fx in ctxData.VFDITEMPLACEMENTs on p.ID equals fx.ITEM_ID
                                  join e in ctxData.VFDITEMs on fx.CONTAINER_ID equals e.ID
                                  join fxx in ctxData.VFDITEMPLACEMENTs on e.ID equals fxx.CONTAINER_ID
                                  where e.ID == id
                                  select new { p.USERDISPLAYNAME, p.ID };

                foreach (var a in queryEquipZ.Distinct().OrderBy(it => it.USERDISPLAYNAME))
                {
                    listData = listData + a.USERDISPLAYNAME + ":" + a.ID + "|";
                }
            }

            string listExist = "";
            //KENA GUNA NAMA EQUIPMENT
            System.Diagnostics.Debug.WriteLine("ID :" + id);
            Tools tool = new Tools();

            string sqlCmdUpdate = "WITH    v (a_id, z_id, a_sub_type) AS (";
            sqlCmdUpdate += " SELECT  a_id, z_id, a_sub_type";
            sqlCmdUpdate += " FROM    CONNECTIVITY";
            sqlCmdUpdate += " WHERE   A_SUB_TYPE = '" + A_SUB_TYPE + "'";
            sqlCmdUpdate += " UNION ALL";
            sqlCmdUpdate += " SELECT  t.a_id, t.z_id, t.a_sub_type";
            sqlCmdUpdate += " FROM    v";
            sqlCmdUpdate += " JOIN    MANAGE_CONNECTIVITY t";
            sqlCmdUpdate += " ON      t.a_id = v.z_id )";
            sqlCmdUpdate += " SELECT  * FROM    v";

            string queryEquip = tool.ExecuteStr(connString, sqlCmdUpdate);
            string[] extData = queryEquip.Split('!');
            //listExist = ;
            //var query = queryEquip.ToList();
            for (int i = 0; i < extData.Count() - 1; i++)
            {
                string firstData = extData[i];
                string[] appearData = firstData.Split(':');
                System.Diagnostics.Debug.WriteLine("appearData :" + appearData[2]);
                listExist = listExist + "(" + appearData[0] + " : " + appearData[2] + " )-->";
            }
                      
            return Json(new
            {
                Success = true,
                listData = listData,
                listExist = listExist
            }, JsonRequestBehavior.AllowGet); //
        }

        [HttpPost]
        public ActionResult findPP(string EXC, int lvl) // Find Patch Panel
        {
            string listData = "";
            using (Entities_NRM ctxData = new Entities_NRM())
            {
                List<SelectListItem> list2 = new List<SelectListItem>();
                var queryEquipZ = from fx in ctxData.JUNCTIONs
                                  join p in ctxData.VFDITEMs on fx.JUNCTIONARRAYMODEL_ID equals p.MODEL_ID 
                                  //join e in ctxData.VFDITEMs on fx.CONTAINER_ID equals e.ID
                                  //join fxx in ctxData.VFDITEMPLACEMENTs on fx.JUNCTIONARRAYMODEL_ID equals fxx.CONTAINER_ID
                                  where (p.DTYPE == "PatchPanel") &&
                                        p.USERDISPLAYNAME.Contains(EXC)
                                  select new { p.USERDISPLAYNAME, p.ID };

                foreach (var a in queryEquipZ.Distinct().OrderBy(it => it.USERDISPLAYNAME))
                {
                    listData = listData + a.USERDISPLAYNAME + ":" + a.ID + "|";
                }
            }

            return Json(new
            {
                Success = true,
                listData = listData
            }, JsonRequestBehavior.AllowGet); //
        }

        [HttpPost]
        public ActionResult findPPort(int id) // Find Chasiss Card
        {
            string listData = "";
            using (Entities_NRM ctxData = new Entities_NRM())
            {
                List<SelectListItem> list2 = new List<SelectListItem>();
                var queryEquipZ = from p in ctxData.JUNCTIONs
                                  join fx in ctxData.VFDITEMs on p.JUNCTIONARRAYMODEL_ID equals fx.MODEL_ID
                                  where fx.ID == id
                                  //p.AUTODISPLAYNAME.Contains(EXC)
                                  select new { p.ROW_, p.COLUMN_, p.ID };

                foreach (var a in queryEquipZ.Distinct().OrderBy(it => it.ROW_).ThenBy(it => it.COLUMN_))
                {
                    listData = listData + a.ROW_ + "-" + a.COLUMN_ + ":" + a.ID + "|";
                }
            }

            return Json(new
            {
                Success = true,
                listData = listData
            }, JsonRequestBehavior.AllowGet); //
        }

        [HttpPost]
        public ActionResult checkACard(string id) // Check A Port
        {
            bool success;
            using (Entities_NEPS ctxData = new Entities_NEPS())
            {
                List<SelectListItem> list2 = new List<SelectListItem>();
                var queryCheckA = (from p in ctxData.CONNECTIVITY
                                   where p.A_ID == id
                                   select p);
                if (queryCheckA.Count() > 0)
                {
                    success = false;
                }
                else
                {
                    success = true;
                }
            }

            return Json(new
            {
                Success = success
            }, JsonRequestBehavior.AllowGet); //
        }

        [HttpPost]
        public ActionResult checkZCard(string id) // Check Z Port
        {
            bool success;
            using (Entities_NEPS ctxData = new Entities_NEPS())
            {
                List<SelectListItem> list2 = new List<SelectListItem>();
                var queryCheckA = (from p in ctxData.CONNECTIVITY
                                   where p.Z_ID == id
                                   select p);
                if (queryCheckA.Count() > 0)
                {
                    success = false;
                }
                else
                {
                    success = true;
                }
            }

            return Json(new
            {
                Success = success
            }, JsonRequestBehavior.AllowGet); //
        }

        [HttpPost]
        public ActionResult InsertManageConn(string EXC_ABB, string ELV_NAME, string A_ID, string A_TYPE, string A_SUB_TYPE, string A_PORT, string Z_ID, string Z_TYPE, string Z_SUB_TYPE, string Z_PORT, string REMARKS, string ATYPE, string ZTYPE) // insert list port
        {
            bool success;

            WebView.WebService._base myWebService = new WebService._base();

            success = myWebService.AddConnectivity(EXC_ABB, ELV_NAME, A_ID, A_TYPE, A_SUB_TYPE, A_PORT, Z_ID, Z_TYPE, Z_SUB_TYPE, Z_PORT, REMARKS, ATYPE, ZTYPE);

            return Json(new
            {
                Success = success
            }, JsonRequestBehavior.AllowGet); //
        }
    
        //------------------------CABINET ADDITIONAL INFO

        public ActionResult ListCabInfo(string searchKey, string pwrCab, string StateID, string zoneID, string pttID, string excID, string rcID, string status, string Model, string Manufacturer, int? page)
        {
            WebView.WebService._base myWebService;
            myWebService = new WebService._base();

            WebService._base.DataCabAddInfo CabInfo1 = new WebService._base.DataCabAddInfo();

            if (searchKey != null || StateID != null || zoneID != null || pttID != null || pttID != null || pwrCab != null || rcID != null || status != null || Model != null || Manufacturer != null)
            {
                if (searchKey == "" && StateID == "" && zoneID == "" && pttID == "" && excID == "" && pwrCab == "" && rcID  == ""&& status == "" && Model == "" && Manufacturer == "")
                {
                    CabInfo1 = myWebService.GetDataCabAddInfo(0, 100, null, null, null, null, null, null, null, null, null, null);
                }
                else if (searchKey != "" || StateID != "" || zoneID != "" || pttID != "" || excID != "" || pwrCab != "" || rcID != "" || status != "" || Model != "" || Manufacturer != "")
                {
                    CabInfo1 = myWebService.GetDataCabAddInfo(0, 100, searchKey, StateID, zoneID, pttID, excID, pwrCab.ToUpper(), rcID, status, Model, Manufacturer);
                    ViewBag.searchKey = searchKey;
                    ViewBag.state = StateID;
                    ViewBag.zone1 = zoneID;
                    ViewBag.ptt = pttID;
                    ViewBag.exc1 = excID;
                    ViewBag.pwrCab = pwrCab;
                    ViewBag.pwrCab1 = pwrCab;
                    ViewBag.status = status;
                    ViewBag.status1 = status;
                    ViewBag.Model = Model;
                    ViewBag.Model1 = Model;
                    ViewBag.Manufacturer = Manufacturer;
                    ViewBag.Manufacturer1 = Manufacturer;
                    ViewBag.rcID = rcID;
                    ViewBag.rcID1 = rcID;
                }
            }
            else
            {
                CabInfo1 = myWebService.GetDataCabAddInfo(0, 100, null, null, null, null, null, null, null, null, null, null);
                ViewBag.searchKey = "";
                ViewBag.ptt = "";
                ViewBag.state = "";
                ViewBag.zone1 = "";
                ViewBag.exc1 = "";
                ViewBag.pwrCab = "";
                ViewBag.pwrCab1 = "";
                ViewBag.status = "";
                ViewBag.status1 = "";
                ViewBag.rcID = "";
                ViewBag.rcID1 = "";
            }

            ViewData["data10"] = CabInfo1.DataCabAddInfoList;

            using (Entities ctxData = new Entities())
            {
                List<SelectListItem> list1 = new List<SelectListItem>();
                var query1 = from p in ctxData.WV_STATE_MAST
                             select new { p.STATE_NAME };
                list1.Add(new SelectListItem() { Text = "", Value = "" });
                foreach (var a in query1.Distinct().OrderBy(it => it.STATE_NAME))
                {
                    list1.Add(new SelectListItem() { Text = a.STATE_NAME, Value = a.STATE_NAME });
                }
                ViewBag.StateId = list1;

                List<SelectListItem> list2 = new List<SelectListItem>();
                var query3 = from p in ctxData.WV_EXC_MAST
                             select new { p.PTT_ID };
                list2.Add(new SelectListItem() { Text = "", Value = "" });
                foreach (var a in query3.Distinct().OrderBy(it => it.PTT_ID))
                {
                    list2.Add(new SelectListItem() { Text = a.PTT_ID, Value = a.PTT_ID });
                }
                ViewBag.PttId = list2;

                List<SelectListItem> list3 = new List<SelectListItem>();
                var query4 = from p in ctxData.WV_EXC_MAST
                             select new { p.EXC_ABB };
                list3.Add(new SelectListItem() { Text = "", Value = "" });
                foreach (var a in query4.Distinct().OrderBy(it => it.EXC_ABB))
                {
                    list3.Add(new SelectListItem() { Text = a.EXC_ABB, Value = a.EXC_ABB });
                }
                ViewBag.Exc = list3;

                List<SelectListItem> list4 = new List<SelectListItem>();
                var query5 = from p in ctxData.WV_CAB_ZONE
                             select new { p.ZONE };
                list4.Add(new SelectListItem() { Text = "", Value = "" });
                foreach (var a in query5.Distinct().OrderBy(it => it.ZONE))
                {
                    list4.Add(new SelectListItem() { Text = a.ZONE, Value = a.ZONE });
                }
                ViewBag.Zone = list4;

                List<SelectListItem> list5 = new List<SelectListItem>();

                var query6 = (from d in ctxData.WV_CAB_RC_BRAND
                              group d by d.BRAND into newGroup
                              orderby newGroup.Key
                              select newGroup);

                list5.Add(new SelectListItem() { Text = "", Value = "" });
                foreach (var a in query6)
                {
                    list5.Add(new SelectListItem() { Text = a.Key, Value = a.Key });
                }
                ViewBag.RC = list5;

                List<SelectListItem> list6 = new List<SelectListItem>();
                var query7 = from p in ctxData.WV_CAB_SA_BRAND //------------BATTERY BRAND
                             select new { p.BRAND };
                list6.Add(new SelectListItem() { Text = "", Value = "" });
                foreach (var a in query7.Distinct().OrderBy(it => it.BRAND))
                {
                    list6.Add(new SelectListItem() { Text = a.BRAND, Value = a.BRAND });
                }
                ViewBag.Batt = list6;

                List<SelectListItem> list8 = new List<SelectListItem>();
                var query8 = from p in ctxData.OSP_FIB_FACILITIES //------------BATTERY BRAND
                             select new { p.MODEL };
                list8.Add(new SelectListItem() { Text = "", Value = "" });
                foreach (var a in query8.Distinct().OrderBy(it => it.MODEL))
                {
                    list8.Add(new SelectListItem() { Text = a.MODEL, Value = a.MODEL });
                }
                ViewBag.Model = list8;

                List<SelectListItem> list9 = new List<SelectListItem>();
                var query9 = from p in ctxData.OSP_FIB_FACILITIES //------------BATTERY BRAND
                             select new { p.MANUFACTURER };
                list9.Add(new SelectListItem() { Text = "", Value = "" });
                foreach (var a in query9.Distinct().OrderBy(it => it.MANUFACTURER))
                {
                    list9.Add(new SelectListItem() { Text = a.MANUFACTURER, Value = a.MANUFACTURER });
                }
                ViewBag.Manufacturer = list9;
            }
            string input = "\\\\adsvr";

            string output = String.Format("http:{0}", input.Replace("\\d$\\x\\y", String.Empty).Replace("\\", "/"));

            ViewBag.output = output;

            //return View();
            int pageSize = 15;
            int pageNumber = (page ?? 1);
            return View(CabInfo1.DataCabAddInfoList.ToPagedList(pageNumber, pageSize));
        }

        public ActionResult EListCabInfo(int g3e_fid, string code, int? page)
        {
            WebView.WebService._base myWebService;
            myWebService = new WebService._base();

            WebService._base.DataCabAddInfo CabInfo1 = new WebService._base.DataCabAddInfo();

            CabInfo1 = myWebService.GetDataCabInfo(0, 50000, g3e_fid);
           
            ViewData["data10"] = CabInfo1.DataCabAddInfoList;

            string input = "\\\\adsvr";

            using (Entities ctxData = new Entities())
            {
                int data = Convert.ToInt32(g3e_fid);
                var query3 = (from p in ctxData.OSP_FIB_FACILITIES
                              where p.G3E_FID == data
                              select new { p.CODE }).Single();
                ViewBag.CODE = query3.CODE;
            }

            string output = String.Format("http:{0}", input.Replace("\\d$\\x\\y", String.Empty).Replace("\\", "/"));

            ViewBag.output = output;

            //return View();
            int pageSize = 15;
            int pageNumber = (page ?? 1);
            return View(CabInfo1.DataCabAddInfoList.ToPagedList(pageNumber, pageSize));
        }

        public ActionResult NewCabInfo()
        {
            using (Entities ctxData = new Entities())
            {
                List<SelectListItem> list2 = new List<SelectListItem>();
                var query3 = from p in ctxData.WV_EXC_MAST
                             select new { p.PTT_ID };
                list2.Add(new SelectListItem() { Text = "", Value = "" });
                foreach (var a in query3.Distinct().OrderBy(it => it.PTT_ID))
                {
                    list2.Add(new SelectListItem() { Text = a.PTT_ID, Value = a.PTT_ID });
                }
                ViewBag.PttId = list2;

                List<SelectListItem> list3 = new List<SelectListItem>();
                ViewBag.EXCType = list3;

                List<SelectListItem> list4 = new List<SelectListItem>();
                list4.Add(new SelectListItem() { Text = "", Value = "" });
                list4.Add(new SelectListItem() { Text = "OK", Value = "OK" });
                list4.Add(new SelectListItem() { Text = "NOT OK", Value = "NOT OK" });
                list4.Add(new SelectListItem() { Text = "N/A", Value = "N/A" });
                ViewBag.Condition = list4;

                List<SelectListItem> list41 = new List<SelectListItem>();
                list41.Add(new SelectListItem() { Text = "", Value = "" });
                list41.Add(new SelectListItem() { Text = "YES", Value = "YES" });
                list41.Add(new SelectListItem() { Text = "NO", Value = "NO" });
                ViewBag.Condition1 = list41;

                List<SelectListItem> list411 = new List<SelectListItem>();
                list411.Add(new SelectListItem() { Text = "", Value = "" });
                list411.Add(new SelectListItem() { Text = "GOOD", Value = "GOOD" });
                list411.Add(new SelectListItem() { Text = "BAD", Value = "BAD" });
                ViewBag.Condition2 = list411;

                List<SelectListItem> list6 = new List<SelectListItem>();
                ViewBag.dnull = list6;

                List<SelectListItem> listELCB = new List<SelectListItem>();
                var elcb = (from d in ctxData.WV_CAB_ELCB_BRAND
                            select new { d.BRAND });

                listELCB.Add(new SelectListItem() { Text = "", Value = "" });
                listELCB.Add(new SelectListItem() { Text = "N/A", Value = "N/A" });

                foreach (var a in elcb.Distinct().OrderBy(it => it.BRAND))
                {
                    listELCB.Add(new SelectListItem() { Text = a.BRAND, Value = a.BRAND });
                }
                listELCB.Add(new SelectListItem() { Text = "OTHERS", Value = "OTHERS" });

                ViewBag.ELCBRAND = listELCB;

                List<SelectListItem> listMC = new List<SelectListItem>();
                var MC = (from d in ctxData.WV_CAB_MCB_BRAND
                          select new { d.BRAND });

                listMC.Add(new SelectListItem() { Text = "", Value = "" });
                listMC.Add(new SelectListItem() { Text = "N/A", Value = "N/A" });
                foreach (var a in MC.Distinct().OrderBy(it => it.BRAND))
                {
                    listMC.Add(new SelectListItem() { Text = a.BRAND, Value = a.BRAND });
                }
                listMC.Add(new SelectListItem() { Text = "OTHERS", Value = "OTHERS" });
                ViewBag.MCBRAND = listMC;

                List<SelectListItem> listSA = new List<SelectListItem>();
                var SA = (from d in ctxData.WV_CAB_SA_BRAND
                          select new { d.BRAND });

                listSA.Add(new SelectListItem() { Text = "", Value = "" });
                listSA.Add(new SelectListItem() { Text = "N/A", Value = "N/A" });
                foreach (var a in SA.Distinct().OrderBy(it => it.BRAND))
                {
                    listSA.Add(new SelectListItem() { Text = a.BRAND, Value = a.BRAND });
                }
                listSA.Add(new SelectListItem() { Text = "OTHERS", Value = "OTHERS" });
                ViewBag.SABRAND = listSA;

                var RC = (from d in ctxData.WV_CAB_RC_BRAND
                          group d by d.BRAND into newGroup
                          orderby newGroup.Key
                          select newGroup );


                List<SelectListItem> listRCB = new List<SelectListItem>();
                listRCB.Add(new SelectListItem() { Text = "", Value = "" });
                listRCB.Add(new SelectListItem() { Text = "N/A", Value = "N/A" });
                foreach (var a in RC)
                {
                    listRCB.Add(new SelectListItem() { Text = a.Key, Value = a.Key });
                }
                listRCB.Add(new SelectListItem() { Text = "OTHERS", Value = "OTHERS" });
                ViewBag.RCBRAND = listRCB;

                ViewBag.RCMODEL = list3;

                List<SelectListItem> listInOut = new List<SelectListItem>();
                listInOut.Add(new SelectListItem() { Text = "", Value = "" });
                listInOut.Add(new SelectListItem() { Text = "INDOOR", Value = "INDOOR" });
                listInOut.Add(new SelectListItem() { Text = "OUTDOOR", Value = "OUTDOOR" });
                ViewBag.IN_OUT = listInOut;

                List<SelectListItem> listAlarmExt = new List<SelectListItem>();
                listAlarmExt.Add(new SelectListItem() { Text = "", Value = "" });
                listAlarmExt.Add(new SelectListItem() { Text = "N/A", Value = "N/A" });
                listAlarmExt.Add(new SelectListItem() { Text = "CPAMS", Value = "CPAMS" });
                listAlarmExt.Add(new SelectListItem() { Text = "FUJITSU", Value = "FUJITSU" });
                listAlarmExt.Add(new SelectListItem() { Text = "EMS", Value = "EMS" });
                ViewBag.ALMEXT = listAlarmExt;

                List<SelectListItem> listAR = new List<SelectListItem>();
                var AR = (from d in ctxData.WV_CAB_AR_BRAND
                          select new { d.BRAND });

                listAR.Add(new SelectListItem() { Text = "", Value = "" });
                listAR.Add(new SelectListItem() { Text = "N/A", Value = "N/A" });
                foreach (var a in AR.Distinct().OrderBy(it => it.BRAND))
                {
                    listAR.Add(new SelectListItem() { Text = a.BRAND, Value = a.BRAND });
                }
                listAR.Add(new SelectListItem() { Text = "OTHERS", Value = "OTHERS" });
                ViewBag.ARBRAND = listAR;
            }
            return View(new WebView.Models.CabInfoModel());
        }

        [HttpPost]
        public ActionResult NewCabInfo(CabInfoModel m)
        {
            bool success = true;
            WebView.WebService._base myWebService;
            myWebService = new WebService._base();
            bool selected = false;
            DateTime? nulldate;
            DateTime sekarang = DateTime.Now;

            WebService._base.CabAddInfo newCabMain = new WebService._base.CabAddInfo();
            newCabMain.G3E_FID = m.G3E_FID2;
            newCabMain.PWR_CAB_ID = "PC_" + m.EXC_ABB + "_" + m.PWR_CAB_ID;
            newCabMain.PWR_CAB_CONDITION = m.PWR_CAB_CONDITION;
            newCabMain.EQUIP_CAB_CONDITION = m.EQUIP_CAB_CONDITION;
            newCabMain.PRO_ELCB_AUTO_BRAND = m.PRO_ELCB_AUTO_BRAND;
            newCabMain.PRO_ELCB_BRAND = m.PRO_ELCB_BRAND;
            newCabMain.PRO_ELCB_VAC = m.PRO_ELCB_VAC;
            newCabMain.PRO_ELCB_LOAC = m.PRO_ELCB_LOAC;
            newCabMain.PRO_ELCB_CAPASITY = m.PRO_ELCB_CAPASITY;
            newCabMain.PRO_ELCB_DATE = Convert.ToDateTime(m.PRO_ELCB_DATE);
            newCabMain.PRO_MCB_BRAND = m.PRO_MCB_BRAND;
            newCabMain.PRO_MCB_CAPASITY = m.PRO_MCB_CAPASITY;
            newCabMain.PRO_MCB_DATE = Convert.ToDateTime(m.PRO_MCB_DATE);
            newCabMain.PRO_SA_BRAND = m.PRO_SA_BRAND;
            newCabMain.PRO_SA_CAPASITY = m.PRO_SA_CAPASITY;
            newCabMain.PRO_SA_CONDITION = m.PRO_SA_CONDITION;
            newCabMain.PRO_AC_OHM = m.PRO_AC_OHM;
            newCabMain.PRO_AC_CONNECTION = m.PRO_AC_CONNECTION;
            newCabMain.RC_BRAND = m.RC_BRAND;
            newCabMain.RC_MODEL = m.RC_MODEL;
            newCabMain.RC_RATING = m.RC_RATING;
            newCabMain.RC_GMODULES = m.RC_GMODULES;
            newCabMain.RC_BMODULES = m.RC_BMODULES;
            newCabMain.RC_DATE = Convert.ToDateTime(m.RC_DATE);
            newCabMain.RC_VAC = m.RC_VAC;
            newCabMain.RC_VDC = m.RC_VDC;
            newCabMain.RC_PRESENT = m.RC_PRESENT;
            newCabMain.RC_MCB_BRAND = null;
            newCabMain.RC_MCB_CAPASITY = null;
            newCabMain.RC_SA_BRAND = m.RC_SA_BRAND;
            newCabMain.RC_SA_CAPASITY = m.RC_SA_CAPASITY;
            newCabMain.RC_SA_CONDITION = m.RC_SA_CONDITION;
            newCabMain.RC_SA_OHM = m.RC_SA_OHM;
            newCabMain.RC_SA_CONNECTION = m.RC_SA_CONNECTION;
            newCabMain.RC_LVD = m.RC_LVD;
            newCabMain.BATT_BRAND = m.BATT_BRAND;
            newCabMain.BATT_CAPASITY = m.BATT_CAPASITY;
            newCabMain.BATT_VOLT = m.BATT_VOLT;
            newCabMain.BATT_DATE = Convert.ToDateTime(m.BATT_DATE);
            newCabMain.BATT_CONDITION = m.BATT_CONDITION;
            newCabMain.AIRCOND_READING = m.AIRCOND_READING;
            newCabMain.AIRCOND_WORK = m.AIRCOND_WORK;
            newCabMain.AIRCOND_BROKE = null;
            newCabMain.AIRCOND_DOOR = m.AIRCOND_DOOR;
            newCabMain.ALARM = m.ALARM;
            newCabMain.CHECK_BY = m.CHECK_BY;
            newCabMain.TNB_METER = m.TNB_METER;
            newCabMain.CAB_INS_DATE = Convert.ToDateTime(m.CAB_INS_DATE);
            newCabMain.ALM_EXT = m.ALM_EXT;
            newCabMain.ALM_EXT_TO = m.ALM_EXT_TO;
            newCabMain.ALM_STATUS = m.ALM_STATUS;
            newCabMain.RC_WARRANTY_END = Convert.ToDateTime(m.RC_WARRANTY_END);
            newCabMain.BATT_EXIST = m.BATT_EXIST;
            newCabMain.BATT_WARRANTY_END = Convert.ToDateTime(m.BATT_WARRANTY_END);
            newCabMain.BATT_NO_CELL = m.BATT_NO_CELL;
            newCabMain.AIRCOND_IN_OUT = m.AIRCOND_IN_OUT;
            newCabMain.AIRCOND_HPOWER = m.AIRCOND_HPOWER;
            newCabMain.REMARKS = m.REMARKS;

            newCabMain.leftValues = m.leftValues;
            int groupId = 0;
            using (Entities ctxData = new Entities())
            {
                var queryGroup = (from p in ctxData.WV_USER
                                where p.FULL_NAME == m.CHECK_BY
                                select new { p.GROUPID }).Single();
                groupId = Convert.ToInt32(queryGroup.GROUPID);
            }
            

            if (newCabMain.PWR_CAB_ID != "" && newCabMain.TNB_METER != "" &&
                newCabMain.PRO_ELCB_AUTO_BRAND != "" && newCabMain.PRO_ELCB_BRAND != "" && newCabMain.PRO_ELCB_LOAC != "" && newCabMain.PRO_ELCB_CAPASITY != "" && newCabMain.PRO_MCB_BRAND != "" &&
                newCabMain.PRO_SA_BRAND != "" && newCabMain.PRO_SA_CAPASITY != "" && newCabMain.PRO_SA_CONDITION != "" && newCabMain.RC_BRAND != "" && newCabMain.RC_MODEL != "" && newCabMain.RC_RATING != ""
                && newCabMain.RC_GMODULES != "" && newCabMain.RC_BMODULES != "" && newCabMain.RC_VDC != "" && newCabMain.BATT_EXIST != "" && newCabMain.AIRCOND_IN_OUT != "")
            {
                if (newCabMain.BATT_EXIST == "YES" && newCabMain.BATT_BRAND != null && newCabMain.BATT_NO_CELL != null && newCabMain.BATT_NO_CELL != null && newCabMain.BATT_CONDITION != null && newCabMain.BATT_CAPASITY != null)
                {
                    System.Diagnostics.Debug.WriteLine("a");
                    newCabMain.STATUS = "COMPLETED";
                }
                else if (newCabMain.BATT_EXIST == "NO")
                {
                    newCabMain.STATUS = "COMPLETED";
                }
                else
                {
                    System.Diagnostics.Debug.WriteLine("c");
                    newCabMain.STATUS = "PENDING";
                }

                if (newCabMain.STATUS == "COMPLETED")
                {
                    if (newCabMain.AIRCOND_IN_OUT == "INDOOR" && newCabMain.AIRCOND_WORK != null)
                    {
                        newCabMain.STATUS = "COMPLETED";
                    }
                    else if (newCabMain.AIRCOND_IN_OUT == "OUTDOOR" && newCabMain.AIRCOND_WORK == null)
                    {
                        newCabMain.STATUS = "COMPLETED";
                    }
                    else
                    {
                        newCabMain.STATUS = "PENDING";
                    }
                }
            }
            else
            {
                newCabMain.STATUS = "PENDING";
            }

            if (newCabMain.STATUS == "COMPLETED")
            {
                if (groupId == 2 && newCabMain.CAB_INS_DATE == Convert.ToDateTime("1/1/0001 12:00:00 AM"))
                {
                    newCabMain.STATUS = "PENDING";
                }
                else { newCabMain.STATUS = "COMPLETED"; }
            }
            success = myWebService.AddCabInfo(newCabMain);
            selected = true;

            return Json(new
            {
                success = success,
                status = newCabMain.STATUS
            }, JsonRequestBehavior.AllowGet);
        }

        public ActionResult EditCabInfo(string id)
        {
            using (Entities ctxData = new Entities())
            {
                int data = Convert.ToInt32(id);

                var m = (from d in ctxData.WV_CAB_ADDINFO
                         where d.DATA_NO == data
                         select d).Single();

                var m1 = (from d in ctxData.OSP_FIB_FACILITIES
                         where d.G3E_FID == m.G3E_FID
                         select d).Single();

                CabInfoModel newCabMain = new CabInfoModel();
                newCabMain.G3E_FID = m.G3E_FID.ToString();
                newCabMain.PWR_CAB_ID = m.PWR_CAB_ID;
                newCabMain.PWR_CAB_CONDITION = m.PWR_CAB_CONDITION;
                newCabMain.EQUIP_CAB_CONDITION = m.EQUIP_CAB_CONDITION;
                newCabMain.PRO_ELCB_AUTO_BRAND = m.PRO_ELCB_AUTO_BRAND;
                newCabMain.PRO_ELCB_BRAND = m.PRO_ELCB_BRAND;
                newCabMain.PRO_ELCB_VAC = m.PRO_ELCB_VAC;
                newCabMain.PRO_ELCB_LOAC = m.PRO_ELCB_LOAC;
                newCabMain.PRO_ELCB_CAPASITY = m.PRO_ELCB_CAPASITY;
                newCabMain.PRO_ELCB_DATE = Convert.ToDateTime(m.PRO_ELCB_DATE);              
                newCabMain.PRO_MCB_BRAND = m.PRO_MCB_BRAND;
                newCabMain.PRO_MCB_CAPASITY = m.PRO_MCB_CAPASITY;
                newCabMain.PRO_MCB_DATE = Convert.ToDateTime(m.PRO_MCB_DATE);
                newCabMain.PRO_SA_BRAND = m.PRO_SA_BRAND;
                newCabMain.PRO_SA_CAPASITY = m.PRO_SA_CAPASITY;
                newCabMain.PRO_SA_CONDITION = m.PRO_SA_CONDITION;
                newCabMain.PRO_AC_OHM = m.PRO_AC_OHM;
                newCabMain.PRO_AC_CONNECTION = m.PRO_AC_CONNECTION;
                newCabMain.RC_BRAND = m.RC_BRAND;
                newCabMain.RC_MODEL = m.RC_MODEL;
                newCabMain.RC_RATING = m.RC_RATING;
                newCabMain.RC_GMODULES = m.RC_GMODULES;
                newCabMain.RC_BMODULES = m.RC_BMODULES;
                newCabMain.RC_DATE = Convert.ToDateTime(m.RC_DATE);
                newCabMain.RC_VAC = m.RC_VAC;
                newCabMain.RC_VDC = m.RC_VDC;
                newCabMain.RC_PRESENT = m.RC_PRESENT;
                //newCabMain.RC_MCB_BRAND = m.RC_MCB_BRAND;
                //newCabMain.RC_MCB_CAPASITY = m.RC_MCB_CAPASITY;
                newCabMain.RC_SA_BRAND = m.RC_SA_BRAND;
                newCabMain.RC_SA_CAPASITY = m.RC_SA_CAPASITY;
                newCabMain.RC_SA_CONDITION = m.RC_SA_CONDITION;
                newCabMain.RC_SA_OHM = m.RC_SA_OHM;
                newCabMain.RC_SA_CONNECTION = m.RC_SA_CONNECTION;
                newCabMain.RC_LVD = m.RC_LVD;
                newCabMain.BATT_BRAND = m.BATT_BRAND;
                newCabMain.BATT_CAPASITY = m.BATT_CAPASITY;
                newCabMain.BATT_VOLT = m.BATT_VOLT;
                newCabMain.BATT_DATE = Convert.ToDateTime(m.BATT_DATE);
                newCabMain.BATT_CONDITION = m.BATT_CONDITION;
                newCabMain.AIRCOND_READING = m.AIRCOND_READING;
                newCabMain.AIRCOND_WORK = m.AIRCOND_WORK;
                newCabMain.AIRCOND_BROKE = m.AIRCOND_BROKE;
                newCabMain.AIRCOND_DOOR = m.AIRCOND_DOOR;
                newCabMain.ALARM = m.ALARM;
                newCabMain.CHECK_BY = m.CHECK_BY;

                newCabMain.TNB_METER = m.TNB_METER;
                newCabMain.CAB_INS_DATE = Convert.ToDateTime(m.CAB_INS_DATE);
                newCabMain.ALM_EXT = m.ALM_EXT;
                newCabMain.ALM_EXT_TO = m.ALM_EXT_TO;
                newCabMain.ALM_STATUS = m.ALM_STATUS;
                newCabMain.RC_WARRANTY_END = Convert.ToDateTime(m.RC_WARRANTY_END);
                newCabMain.BATT_EXIST = m.BATT_EXIST;
                newCabMain.BATT_WARRANTY_END = Convert.ToDateTime(m.BATT_WARRANTY_END);
                newCabMain.BATT_NO_CELL = m.BATT_NO_CELL;
                newCabMain.AIRCOND_IN_OUT = m.AIRCOND_IN_OUT;
                newCabMain.AIRCOND_HPOWER = m.AIRCOND_HPOWER;
                newCabMain.REMARKS = m.REMARKS;
                newCabMain.DATA_NO = m.DATA_NO.ToString();
                newCabMain.CODE = m1.CODE;
                newCabMain.ZONE = m1.ZONE;
                newCabMain.BATT_EXTRA_LOCKING = m.BATT_EXTRA_LOCKING;

                //DateTime traplan = new DateTime(m1.INSTALL_YEAR);
                //string tarikhPlan  = traplan.Date() + "-" traplan.Month() + "-" + traplan.Year();
                newCabMain.INSTALL_YEAR = m1.INSTALL_YEAR.ToString();
                System.Diagnostics.Debug.WriteLine("c :" + newCabMain.RC_WARRANTY_END);
                System.Diagnostics.Debug.WriteLine("b :" + newCabMain.BATT_WARRANTY_END);
                
                int fid = Convert.ToInt32(m.G3E_FID);
                var n = (from d in ctxData.OSP_FIB_FACILITIES
                         where d.G3E_FID == fid
                         select d).Single();

                newCabMain.CONTRACTOR = n.CONTRACTOR;
                newCabMain.MANUFACTURER = n.MANUFACTURER;

                var m2 = (from d in ctxData.AG_OSP_FIB_FACILITIES
                          where d.G3E_FID == fid
                          select d).Single();

                decimal latitude = Convert.ToDecimal(m2.LATITUDE);
                decimal NewLat = Math.Truncate(1000000 * latitude) / 1000000;

                decimal longitude = Convert.ToDecimal(m2.LONGITUDE);
                decimal NewLong = Math.Truncate(1000000 * longitude) / 1000000;

                newCabMain.latlong = NewLat.ToString() + "/" + NewLong.ToString();

                List<SelectListItem> list4 = new List<SelectListItem>();
                list4.Add(new SelectListItem() { Text = null, Value = null });
                list4.Add(new SelectListItem() { Text = "OK", Value = "OK" });
                list4.Add(new SelectListItem() { Text = "NOT OK", Value = "NOT OK" });
                list4.Add(new SelectListItem() { Text = "N/A", Value = "N/A" });
                ViewBag.Condition = list4;

                List<SelectListItem> listPro = new List<SelectListItem>();
                listPro.Add(new SelectListItem() { Text = null, Value = null });
                listPro.Add(new SelectListItem() { Text = "OK", Value = "OK" });
                listPro.Add(new SelectListItem() { Text = "NOT OK", Value = "NOT OK" });
                listPro.Add(new SelectListItem() { Text = "N/A", Value = "N/A" });
                ViewBag.ConditionPro = listPro;

                List<SelectListItem> listRC_SA_Cond = new List<SelectListItem>();
                listRC_SA_Cond.Add(new SelectListItem() { Text = null, Value = null });
                listRC_SA_Cond.Add(new SelectListItem() { Text = "OK", Value = "OK" });
                listRC_SA_Cond.Add(new SelectListItem() { Text = "NOT OK", Value = "NOT OK" });
                listRC_SA_Cond.Add(new SelectListItem() { Text = "N/A", Value = "N/A" });
                ViewBag.ConditionRC_SA = listRC_SA_Cond;

                List<SelectListItem> list41 = new List<SelectListItem>();
                list41.Add(new SelectListItem() { Text = "", Value = "" });
                list41.Add(new SelectListItem() { Text = "YES", Value = "YES" });
                list41.Add(new SelectListItem() { Text = "NO", Value = "NO" });
                ViewBag.Condition1 = list41;

                List<SelectListItem> list411 = new List<SelectListItem>();
                list411.Add(new SelectListItem() { Text = "", Value = "" });
                list411.Add(new SelectListItem() { Text = "GOOD", Value = "GOOD" });
                list411.Add(new SelectListItem() { Text = "BAD", Value = "BAD" });
                ViewBag.Condition2 = list411;

                List<SelectListItem> listELCB = new List<SelectListItem>();
                var elcb = (from d in ctxData.WV_CAB_ELCB_BRAND
                            select new { d.BRAND });

                listELCB.Add(new SelectListItem() { Text = "", Value = "" });
                listELCB.Add(new SelectListItem() { Text = "N/A", Value = "N/A" });

                foreach (var a in elcb.Distinct().OrderBy(it => it.BRAND))
                {
                    listELCB.Add(new SelectListItem() { Text = a.BRAND, Value = a.BRAND });
                }
                listELCB.Add(new SelectListItem() { Text = "OTHERS", Value = "OTHERS" });
                ViewBag.ELCBRAND = listELCB;

                List<SelectListItem> listMC = new List<SelectListItem>();
                var MC = (from d in ctxData.WV_CAB_MCB_BRAND
                          select new { d.BRAND });

                listMC.Add(new SelectListItem() { Text = "", Value = "" });
                listMC.Add(new SelectListItem() { Text = "N/A", Value = "N/A" });
                foreach (var a in MC.Distinct().OrderBy(it => it.BRAND))
                {
                    listMC.Add(new SelectListItem() { Text = a.BRAND, Value = a.BRAND });
                }
                listMC.Add(new SelectListItem() { Text = "OTHERS", Value = "OTHERS" });
                ViewBag.MCBRAND = listMC;

                List<SelectListItem> listSA = new List<SelectListItem>();
                var SA = (from d in ctxData.WV_CAB_SA_BRAND
                          select new { d.BRAND });

                listSA.Add(new SelectListItem() { Text = "", Value = "" });
                listSA.Add(new SelectListItem() { Text = "N/A", Value = "N/A" });
                foreach (var a in SA.Distinct().OrderBy(it => it.BRAND))
                {
                    listSA.Add(new SelectListItem() { Text = a.BRAND, Value = a.BRAND });
                }
                listSA.Add(new SelectListItem() { Text = "OTHERS", Value = "OTHERS" });
                ViewBag.SABRAND = listSA;

                var RC = (from d in ctxData.WV_CAB_RC_BRAND
                          group d by d.BRAND into newGroup
                          orderby newGroup.Key
                          select newGroup);


                List<SelectListItem> listRCB = new List<SelectListItem>();
                listRCB.Add(new SelectListItem() { Text = "", Value = "" });
                listRCB.Add(new SelectListItem() { Text = "N/A", Value = "N/A" });
                foreach (var a in RC)
                {
                    listRCB.Add(new SelectListItem() { Text = a.Key, Value = a.Key });
                }
                listRCB.Add(new SelectListItem() { Text = "OTHERS", Value = "OTHERS" });
                ViewBag.RCBRAND = listRCB;

                List<SelectListItem> list3 = new List<SelectListItem>();
                list3.Add(new SelectListItem() { Text = m.RC_MODEL, Value = m.RC_MODEL });
                ViewBag.RCMODEL = list3;

                List<SelectListItem> listInOut = new List<SelectListItem>();
                listInOut.Add(new SelectListItem() { Text = "", Value = "" });
                listInOut.Add(new SelectListItem() { Text = "INDOOR", Value = "INDOOR" });
                listInOut.Add(new SelectListItem() { Text = "OUTDOOR", Value = "OUTDOOR" });
                ViewBag.IN_OUT = listInOut;
                
                List<SelectListItem> listAlarmExt = new List<SelectListItem>();
                listAlarmExt.Add(new SelectListItem() { Text = "", Value = "" });
                listAlarmExt.Add(new SelectListItem() { Text = "N/A", Value = "N/A" });
                listAlarmExt.Add(new SelectListItem() { Text = "CPAMS", Value = "CPAMS" });
                listAlarmExt.Add(new SelectListItem() { Text = "FUJITSU", Value = "FUJITSU" });
                listAlarmExt.Add(new SelectListItem() { Text = "EMS", Value = "EMS" });
                ViewBag.ALMEXT = listAlarmExt;

                List<SelectListItem> listAR = new List<SelectListItem>();
                var AR = (from d in ctxData.WV_CAB_AR_BRAND
                          select new { d.BRAND });

                listAR.Add(new SelectListItem() { Text = "", Value = "" });
                listAR.Add(new SelectListItem() { Text = "N/A", Value = "N/A" });
                foreach (var a in AR.Distinct().OrderBy(it => it.BRAND))
                {
                    listAR.Add(new SelectListItem() { Text = a.BRAND, Value = a.BRAND });
                }
                listAR.Add(new SelectListItem() { Text = "OTHERS", Value = "OTHERS" });
                ViewBag.ARBRAND = listAR;

                string[] exch;
                exch = m.PWR_CAB_ID.Split('_');
                string ech = exch[1];

                List<SelectListItem> list21 = new List<SelectListItem>();
                var PFCab = (from d in ctxData.OSP_FIB_FACILITIES
                             where d.EXCHANGE.Trim() == ech
                             select new { d.CODE });
                list21.Add(new SelectListItem() { Text = exch[2], Value = exch[2] });
                foreach (var i in PFCab.Distinct().OrderBy(it => it.CODE))
                {
                    list21.Add(new SelectListItem() { Text = i.CODE, Value = i.CODE });
                }
                ViewBag.Pwr_Cab = list21;

                newCabMain.EXC_ABB = ech;
                string dataRF = exch[2];

                var DCCab = (from fx in ctxData.WV_CAB_SHARE
                             where fx.NE_ID == m.G3E_FID 
                             select new { fx.DC_ID }).Single();

                List<SelectListItem> list32 = new List<SelectListItem>();
                var LFCab = (from d in ctxData.OSP_FIB_FACILITIES
                             join fx in ctxData.WV_CAB_SHARE on d.G3E_FID equals fx.NE_ID
                             where d.EXCHANGE.Trim() == ech
                             && d.CODE != m1.CODE //&& d.G3E_FID != m.G3E_FID 
                             && fx.DC_ID == DCCab.DC_ID
                             select new { d.CODE, d.G3E_FID }); 
                //list31.Add(new SelectListItem() { Text = exch[2], Value = exch[2] });
                foreach (var i in LFCab.Distinct().OrderBy(it => it.CODE))
                {
                    list32.Add(new SelectListItem() { Text = i.CODE, Value = i.G3E_FID.ToString(), Selected = true });
                }
                ViewBag.LF_Cab = list32;



                //-------update 12/02/2015----------------------------------------------------------------------------------------
                List<SelectListItem> list31 = new List<SelectListItem>();
                var RFCab = (from d in ctxData.OSP_FIB_FACILITIES
                             //join fx in ctxData.WV_CAB_SHARE on d.G3E_FID equals fx.NE_ID
                             where d.EXCHANGE.Trim() == ech
                                 //&& fx.DC_ID != DCCab.DC_ID
                             && d.CODE != m1.CODE
                             //&& d.FACILITIES == 1
                             select new { d.CODE, d.G3E_FID });
                //list31.Add(new SelectListItem() { Text = exch[2], Value = exch[2] });
                foreach (var i in RFCab.Distinct().OrderBy(it => it.CODE))
                {
                    list31.Add(new SelectListItem() { Text = i.CODE, Value = i.G3E_FID.ToString() });
                }
                ViewBag.RF_Cab = list31;

                return View(newCabMain);
            }
        }

        [HttpPost]
        public ActionResult EditCabInfo(CabInfoModel m)
        {
            bool success = true;
            WebView.WebService._base myWebService;
            myWebService = new WebService._base();

            bool selected = false;
            m.PWR_CAB_ID = "PC_" + m.EXC_ABB + "_" + m.PWR_CAB_ID;

            WebService._base.CabAddInfo newCabMain = new WebService._base.CabAddInfo();
            newCabMain.G3E_FID = m.G3E_FID;
            newCabMain.PWR_CAB_ID = m.PWR_CAB_ID;
            newCabMain.PWR_CAB_CONDITION = m.PWR_CAB_CONDITION;
            newCabMain.EQUIP_CAB_CONDITION = m.EQUIP_CAB_CONDITION;
            newCabMain.PRO_ELCB_AUTO_BRAND = m.PRO_ELCB_AUTO_BRAND;
            newCabMain.PRO_ELCB_BRAND = m.PRO_ELCB_BRAND;
            newCabMain.PRO_ELCB_VAC = m.PRO_ELCB_VAC;
            newCabMain.PRO_ELCB_LOAC = m.PRO_ELCB_LOAC;
            newCabMain.PRO_ELCB_CAPASITY = m.PRO_ELCB_CAPASITY;
            newCabMain.PRO_ELCB_DATE = Convert.ToDateTime(m.PRO_ELCB_DATE);
            newCabMain.PRO_MCB_BRAND = m.PRO_MCB_BRAND;
            newCabMain.PRO_MCB_CAPASITY = m.PRO_MCB_CAPASITY;
            newCabMain.PRO_MCB_DATE = Convert.ToDateTime(m.PRO_MCB_DATE);
            newCabMain.PRO_SA_BRAND = m.PRO_SA_BRAND;
            newCabMain.PRO_SA_CAPASITY = m.PRO_SA_CAPASITY;
            newCabMain.PRO_SA_CONDITION = m.PRO_SA_CONDITION;
            newCabMain.PRO_AC_OHM = m.PRO_AC_OHM;
            newCabMain.PRO_AC_CONNECTION = m.PRO_AC_CONNECTION;
            newCabMain.RC_BRAND = m.RC_BRAND;
            newCabMain.RC_MODEL = m.RC_MODEL;
            newCabMain.RC_RATING = m.RC_RATING;
            newCabMain.RC_GMODULES = m.RC_GMODULES;
            newCabMain.RC_BMODULES = m.RC_BMODULES;
            newCabMain.RC_DATE = Convert.ToDateTime(m.RC_DATE);
            newCabMain.RC_VAC = m.RC_VAC;
            newCabMain.RC_VDC = m.RC_VDC;
            newCabMain.RC_PRESENT = m.RC_PRESENT;
            //--------------------------------------------
            newCabMain.RC_MCB_BRAND = null;
            newCabMain.RC_MCB_CAPASITY = null;
            //--------------------------------------------
            newCabMain.RC_SA_BRAND = m.RC_SA_BRAND;
            newCabMain.RC_SA_CAPASITY = m.RC_SA_CAPASITY;
            newCabMain.RC_SA_CONDITION = m.RC_SA_CONDITION;
            newCabMain.RC_SA_OHM = m.RC_SA_OHM;
            newCabMain.RC_SA_CONNECTION = m.RC_SA_CONNECTION;
            newCabMain.RC_LVD = m.RC_LVD;
            newCabMain.BATT_BRAND = m.BATT_BRAND;
            newCabMain.BATT_CAPASITY = m.BATT_CAPASITY;
            newCabMain.BATT_VOLT = m.BATT_VOLT;
            newCabMain.BATT_DATE = Convert.ToDateTime(m.BATT_DATE);
            newCabMain.BATT_CONDITION = m.BATT_CONDITION;
            newCabMain.AIRCOND_READING = m.AIRCOND_READING;
            newCabMain.AIRCOND_WORK = m.AIRCOND_WORK;
            //--------------------------------------------
            newCabMain.AIRCOND_BROKE = null;
            //--------------------------------------------
            newCabMain.AIRCOND_DOOR = m.AIRCOND_DOOR;
            newCabMain.ALARM = m.ALARM;
            newCabMain.CHECK_BY = m.CHECK_BY;

            newCabMain.TNB_METER = m.TNB_METER;
            newCabMain.CAB_INS_DATE = Convert.ToDateTime(m.CAB_INS_DATE);
            newCabMain.ALM_EXT = m.ALM_EXT;
            newCabMain.ALM_EXT_TO = m.ALM_EXT_TO;
            newCabMain.ALM_STATUS = m.ALM_STATUS;
            newCabMain.RC_WARRANTY_END = Convert.ToDateTime(m.RC_WARRANTY_END);
            newCabMain.BATT_EXIST = m.BATT_EXIST;
            newCabMain.BATT_WARRANTY_END = Convert.ToDateTime(m.BATT_WARRANTY_END);
            newCabMain.BATT_NO_CELL = m.BATT_NO_CELL;
            newCabMain.AIRCOND_IN_OUT = m.AIRCOND_IN_OUT;
            newCabMain.AIRCOND_HPOWER = m.AIRCOND_HPOWER;
            newCabMain.REMARKS = m.REMARKS;
            newCabMain.DATA_NO = m.DATA_NO;
            newCabMain.leftValues = m.leftValues;
            newCabMain.BATT_EXTRA_LOCKING = m.BATT_EXTRA_LOCKING;
            //System.Diagnostics.Debug.WriteLine("b :" + m.PRO_ELCB_BRAND);

            int groupId = 0;
            using (Entities ctxData = new Entities())
            {
                var queryGroup = (from p in ctxData.WV_USER
                                  where p.FULL_NAME == m.CHECK_BY
                                  select new { p.GROUPID }).Single();
                groupId = Convert.ToInt32(queryGroup.GROUPID);
            }

            if (newCabMain.PWR_CAB_ID != null && newCabMain.TNB_METER != null &&
                newCabMain.PRO_ELCB_AUTO_BRAND != null && newCabMain.PRO_ELCB_BRAND != null && newCabMain.PRO_ELCB_LOAC != null && newCabMain.PRO_ELCB_CAPASITY != null && newCabMain.PRO_MCB_BRAND != null && newCabMain.PRO_MCB_CAPASITY != null &&
                newCabMain.PRO_SA_BRAND != null && newCabMain.PRO_SA_CAPASITY != null && newCabMain.PRO_SA_CONDITION != null && newCabMain.RC_BRAND != null && newCabMain.RC_MODEL != null && newCabMain.RC_RATING != null
                && newCabMain.RC_GMODULES != null && newCabMain.RC_BMODULES != null && newCabMain.RC_VDC != null && newCabMain.BATT_EXIST != null && newCabMain.AIRCOND_IN_OUT != null)
            {
                //newCabMain.STATUS = "COMPLETED";
                if (newCabMain.BATT_EXIST == "YES" && newCabMain.BATT_BRAND != null && newCabMain.BATT_NO_CELL != null && newCabMain.BATT_NO_CELL != null && newCabMain.BATT_CONDITION != null && newCabMain.BATT_CAPASITY != null)
                {
                    System.Diagnostics.Debug.WriteLine("a");
                    newCabMain.STATUS = "COMPLETED";
                }
                else if (newCabMain.BATT_EXIST == "NO")
                {
                    System.Diagnostics.Debug.WriteLine("b");
                    newCabMain.STATUS = "COMPLETED"; 
                }
                else
                {
                    System.Diagnostics.Debug.WriteLine("c");
                    newCabMain.STATUS = "PENDING";
                }

                if (newCabMain.STATUS == "COMPLETED")
                {
                    if (newCabMain.AIRCOND_IN_OUT == "INDOOR" && newCabMain.AIRCOND_WORK != null)
                    {
                        newCabMain.STATUS = "COMPLETED";
                    }
                    else if (newCabMain.AIRCOND_IN_OUT == "OUTDOOR" && newCabMain.AIRCOND_WORK == null)
                    {
                        newCabMain.STATUS = "COMPLETED";
                    }
                    else
                    {
                        newCabMain.STATUS = "PENDING";
                    }
                }
            }
            else
            {
                newCabMain.STATUS = "PENDING";
            }

            if (newCabMain.STATUS == "COMPLETED")
            {
                DateTime aa = Convert.ToDateTime("1/1/0001 12:00:00 AM");
                if (groupId == 2 && newCabMain.CAB_INS_DATE == aa)
                {
                    newCabMain.STATUS = "PENDING";
                }
                else { newCabMain.STATUS = "COMPLETED"; }
            }

            success = myWebService.EditCabInfo(newCabMain);
            selected = true;


            return Json(new
            {
                success = success,
                status = newCabMain.STATUS
            }, JsonRequestBehavior.AllowGet);
            //if (success == true)
            //    return RedirectToAction("ListCabInfo");
            //else
            //    return RedirectToAction("NewSaveFail"); 
        }

        [HttpPost]
        public ActionResult PTTChange(string PTT_ID)
        {
            string CITY = "";
            string CityLists = "";
            // System.Diagnostics.Debug.WriteLine("CITY: " + Stateid);
            using (Entities ctxState = new Entities())
            {
                string allExc = "";
                var EXC = (from d in ctxState.WV_EXC_MAST
                           where d.PTT_ID.Trim() == PTT_ID.Trim()
                           select new { d.EXC_ABB });

                foreach (var i in EXC.Distinct().OrderBy(it => it.EXC_ABB))
                {
                    allExc = allExc + i.EXC_ABB + ":";
                }
                // System.Diagnostics.Debug.WriteLine("CITY: " + CityLists );
                return Json(new
                {
                    Success = true,
                    EXC = allExc
                }, JsonRequestBehavior.AllowGet);
            }
        }

        [HttpPost]
        public ActionResult EXCChange(string PTT_ID, string EXC_ABB)
        {
            string CITY = "";
            string CityLists = "";
            // System.Diagnostics.Debug.WriteLine("CITY: " + Stateid);
            using (Entities ctxState = new Entities())
            {
                string allCC = "";
                string allFC = "";
           
                var dzone = (from d in ctxState.WV_CAB_ZONE
                             where d.EXCHANGE.Trim() == EXC_ABB.Trim() && d.PTT.Trim() == PTT_ID.Trim()
                             select new { d.ZONE }).Single();

                var FCab = (from d in ctxState.OSP_FIB_FACILITIES
                            where d.EXCHANGE.Trim() == EXC_ABB.Trim()
                            && d.FACILITIES != 1
                            select new { d.G3E_FID, d.CODE });

                foreach (var i in FCab.Distinct().OrderBy(it => it.CODE))
                {
                    allFC = allFC + i.G3E_FID + "/" + i.CODE + ":";
                }

                var PFCab = (from d in ctxState.OSP_FIB_FACILITIES
                            where d.EXCHANGE.Trim() == EXC_ABB.Trim()
                            select new { d.G3E_FID, d.CODE });

                foreach (var i in PFCab.Distinct().OrderBy(it => it.CODE))
                {
                    allCC = allCC + i.G3E_FID + "/" + i.CODE + ":";
                }

                return Json(new
                {
                    Success = true,
                    allCC = allCC,
                    allFC = allFC,
                    zone = dzone.ZONE
                }, JsonRequestBehavior.AllowGet);
            }
        }

        [HttpPost]
        public ActionResult CabChange(int G3E_FID)
        {
            using (Entities ctxState = new Entities())
            {
                string allBRAND = "";
                string allCONTRACTOR = "";
                var QBRAND = (from d in ctxState.GC_ITFACE
                              where d.G3E_FID == G3E_FID
                              select new { d.ITFACE_TYPE, d.ITFACE_CLASS, d.MODEL }).Single();

                allBRAND = QBRAND.ITFACE_TYPE + ":" + QBRAND.ITFACE_CLASS + ":" + QBRAND.MODEL + ":";

                return Json(new
                {
                    Success = true,
                    BRAND = allBRAND
                }, JsonRequestBehavior.AllowGet);
            }
        }

        [HttpPost]
        public ActionResult MSANChange(int G3E_FID2, string EXC_ABB)
        {
            using (Entities ctxState = new Entities())
            {
                string allBRAND = "";
                string allCONTRACTOR = "";
                string allFC = "";
                string CallCab = "";
                var QBRAND = (from d in ctxState.OSP_FIB_FACILITIES
                              join fx in ctxState.AG_OSP_FIB_FACILITIES on d.G3E_FID equals fx.G3E_FID
                              where d.G3E_FID == G3E_FID2
                              select new { d.MANUFACTURER, d.CONTRACTOR, d.MODEL, d.INSTALL_YEAR, fx.LATITUDE, fx.LONGITUDE }).Single();
                 
                decimal latitude = Convert.ToDecimal(QBRAND.LATITUDE);
                decimal NewLat = Math.Truncate(1000 * latitude) / 1000;
                string NNLat = NewLat.ToString();
                latitude = Decimal.Round(latitude, 6);

                decimal longitude = Convert.ToDecimal(QBRAND.LONGITUDE);
                decimal NewLong = Math.Truncate(1000 * longitude) / 1000;
                string NNLong = NewLong.ToString();//string.Format("{0:N3}", QBRAND.LONGITUDE);
                longitude = Decimal.Round(longitude, 6);

                allBRAND = QBRAND.MANUFACTURER + "!" + QBRAND.CONTRACTOR + "!" + QBRAND.MODEL + "!" + QBRAND.INSTALL_YEAR + "!" + latitude + " / " + longitude + "!";

                var CheckCab = (from d in ctxState.AG_OSP_FIB_FACILITIES
                                join fx in ctxState.OSP_FIB_FACILITIES on d.G3E_FID equals fx.G3E_FID
                                where d.LONGITUDE.Contains(NNLong)
                                && d.LATITUDE.Contains(NNLat)
                                select new { d.G3E_FID, d.CODE, d.LATITUDE, d.LONGITUDE ,fx.PWR_CAB_ID });

                foreach (var ii in CheckCab.Distinct().OrderBy(it => it.CODE))
                {
                    CallCab += ii.G3E_FID + "/" + ii.CODE + "/" + (Math.Truncate(1000000 * Convert.ToDecimal(ii.LATITUDE)) / 1000000) + "/" + (Math.Truncate(1000000 * Convert.ToDecimal(ii.LONGITUDE)) / 1000000) + "/" + ii.PWR_CAB_ID + ":";
                }

                //-------update 12/02/2015----------------------------------------------------------------------------------------
                var FCab = (from d in ctxState.OSP_FIB_FACILITIES
                            where d.EXCHANGE.Trim() == EXC_ABB.Trim()
                            //&& d.FACILITIES != 1
                            select new { d.G3E_FID, d.CODE });

                foreach (var i in FCab.Distinct().OrderBy(it => it.CODE))
                {
                    allFC = allFC + i.G3E_FID + "/" + i.CODE + ":";
                }
                return Json(new
                {
                    Success = true,
                    BRAND = allBRAND,
                    allFC = allFC,
                    CallCab = CallCab
                }, JsonRequestBehavior.AllowGet);
            }
        }
        
        [HttpPost]
        public ActionResult PwrCabChange(string newPCD)
        {
            CabInfoModel newCabMain = new CabInfoModel();
            string PWR_CAB_ID = "";
            using (Entities ctxState = new Entities())
            {
                var queryHandover = (from a in ctxState.WV_CAB_ADDINFO
                                     where a.PWR_CAB_ID.Trim() == newPCD.Trim()
                                     select a).Count();
                
                if (queryHandover > 0)
                {
                    
                    var QBRAND = (from d in ctxState.WV_CAB_ADDINFO
                                  where d.PWR_CAB_ID.Trim() == newPCD.Trim()
                                  group d by new { 
                                       d.PWR_CAB_ID, d.PWR_CAB_CONDITION, d.EQUIP_CAB_CONDITION, d.CAB_INS_DATE, d.TNB_METER, d.PRO_AC_CONNECTION, d.PRO_AC_OHM, d.PRO_ELCB_AUTO_BRAND, d.PRO_ELCB_BRAND,
                                       d.PRO_ELCB_CAPASITY, d.PRO_ELCB_DATE, d.PRO_ELCB_LOAC, d.PRO_ELCB_VAC, d.PRO_MCB_BRAND, d.PRO_MCB_CAPASITY,
                                       d.PRO_MCB_DATE, d.PRO_SA_BRAND, d.PRO_SA_CAPASITY, d.PRO_SA_CONDITION
                                  } into fx
                                  select new
                                  {
                                      fx.Key.PWR_CAB_ID,
                                      fx.Key.PWR_CAB_CONDITION,
                                      fx.Key.CAB_INS_DATE,
                                      fx.Key.TNB_METER,
                                      fx.Key.EQUIP_CAB_CONDITION,
                                      fx.Key.PRO_AC_CONNECTION,
                                      fx.Key.PRO_AC_OHM,
                                      fx.Key.PRO_ELCB_AUTO_BRAND,
                                      fx.Key.PRO_ELCB_BRAND,
                                      fx.Key.PRO_ELCB_CAPASITY,
                                      fx.Key.PRO_ELCB_DATE,
                                      fx.Key.PRO_ELCB_LOAC,
                                      fx.Key.PRO_ELCB_VAC,
                                      fx.Key.PRO_MCB_BRAND,
                                      fx.Key.PRO_MCB_CAPASITY,
                                      fx.Key.PRO_MCB_DATE,
                                      fx.Key.PRO_SA_BRAND,
                                      fx.Key.PRO_SA_CAPASITY,
                                      fx.Key.PRO_SA_CONDITION
                                  });

                    foreach (var a in QBRAND)
                    {
                        //newCabMain.PWR_CAB_ID = QBRAND.PWR_CAB_ID;
                        if (a.PWR_CAB_ID != "") { PWR_CAB_ID = "1"; }
                        else { PWR_CAB_ID = ""; }
                        newCabMain.PWR_CAB_CONDITION = a.PWR_CAB_CONDITION;
                        newCabMain.EQUIP_CAB_CONDITION = a.EQUIP_CAB_CONDITION;
                        newCabMain.CAB_INS_DATE = Convert.ToDateTime(a.CAB_INS_DATE);
                        newCabMain.TNB_METER = a.TNB_METER;
                        newCabMain.PRO_AC_CONNECTION = a.PRO_AC_CONNECTION;
                        newCabMain.PRO_AC_OHM = a.PRO_AC_OHM;
                        newCabMain.PRO_ELCB_AUTO_BRAND = a.PRO_ELCB_AUTO_BRAND;
                        newCabMain.PRO_ELCB_BRAND = a.PRO_ELCB_BRAND;
                        newCabMain.PRO_ELCB_CAPASITY = a.PRO_ELCB_CAPASITY;
                        newCabMain.PRO_ELCB_DATE = Convert.ToDateTime(a.PRO_ELCB_DATE);
                        newCabMain.PRO_ELCB_LOAC = a.PRO_ELCB_LOAC;
                        newCabMain.PRO_ELCB_VAC = a.PRO_ELCB_VAC;
                        newCabMain.PRO_MCB_BRAND = a.PRO_MCB_BRAND;
                        newCabMain.PRO_MCB_CAPASITY = a.PRO_MCB_CAPASITY;
                        newCabMain.PRO_MCB_DATE = Convert.ToDateTime(a.PRO_MCB_DATE);
                        newCabMain.PRO_SA_BRAND = a.PRO_SA_BRAND;
                        newCabMain.PRO_SA_CAPASITY = a.PRO_SA_CAPASITY;
                        newCabMain.PRO_SA_CONDITION = a.PRO_SA_CONDITION;
                    }
                
                }
                //return View(newCabMain);
                return Json(new
                {
                    Success = true,
                    PWR_CAB_ID = PWR_CAB_ID,
                    PWR_CAB_CONDITION = newCabMain.PWR_CAB_CONDITION,
                    EQUIP_CAB_CONDITION = newCabMain.EQUIP_CAB_CONDITION,
                    CAB_INS_DATE = newCabMain.CAB_INS_DATE,
                    TNB_METER = newCabMain.TNB_METER,
                    PRO_AC_CONNECTION = newCabMain.PRO_AC_CONNECTION,
                    PRO_AC_OHM = newCabMain.PRO_AC_OHM,
                    PRO_ELCB_AUTO_BRAND = newCabMain.PRO_ELCB_AUTO_BRAND,
                    PRO_ELCB_BRAND = newCabMain.PRO_ELCB_BRAND,
                    PRO_ELCB_CAPASITY = newCabMain.PRO_ELCB_CAPASITY,
                    PRO_ELCB_DATE = newCabMain.PRO_ELCB_DATE,
                    PRO_ELCB_LOAC = newCabMain.PRO_ELCB_LOAC,
                    PRO_ELCB_VAC = newCabMain.PRO_ELCB_VAC,
                    PRO_MCB_BRAND = newCabMain.PRO_MCB_BRAND,
                    PRO_MCB_CAPASITY = newCabMain.PRO_MCB_CAPASITY,
                    PRO_MCB_DATE = newCabMain.PRO_MCB_DATE,
                    PRO_SA_BRAND = newCabMain.PRO_SA_BRAND,
                    PRO_SA_CAPASITY = newCabMain.PRO_SA_CAPASITY,
                    PRO_SA_CONDITION = newCabMain.PRO_SA_CONDITION
                }, JsonRequestBehavior.AllowGet);
            }
        }

        [HttpPost]
        public ActionResult RCBrandChange(string RC_BRAND)
        {
            using (Entities ctxState = new Entities())
            {
                string allBRAND = "";
                var QBRAND = (from d in ctxState.WV_CAB_RC_BRAND
                              where d.BRAND == RC_BRAND.ToUpper()
                              select new { d.MODEL });
                allBRAND += " :N/A:";
                foreach (var a in QBRAND.Distinct().OrderBy(it => it.MODEL))
                {
                    allBRAND += a.MODEL + ":";
                }
                return Json(new
                {
                    Success = true,
                    BRAND = allBRAND
                }, JsonRequestBehavior.AllowGet);
            }
        }

        [HttpPost]
        public ActionResult checkPwr(int G3E_FID)
        {
            CabInfoModel newCabMain = new CabInfoModel();
            System.Diagnostics.Debug.WriteLine("G3E_FID : " + G3E_FID);

            using (Entities ctxState = new Entities())
            {
                string allBRAND = "";
                var QBRAND = (from d in ctxState.WV_CAB_SHARE
                              where d.NE_ID == G3E_FID
                              select new { d.NE_ID, d.PC_ID });

                foreach (var a in QBRAND.Distinct().OrderBy(it => it.PC_ID))
                {
                    allBRAND += a.NE_ID + ":" + a.PC_ID + ":"; 
                }
                System.Diagnostics.Debug.WriteLine("allBRAND : " + allBRAND);
                if (allBRAND != "")
                {
                    var ShareNo = (from d in ctxState.WV_CAB_SHARE
                              where d.NE_ID == G3E_FID
                              select new { d.DC_ID }).Single();

                    var daa = (from d in ctxState.WV_CAB_ADDINFO
                               where d.G3E_FID == G3E_FID
                               select d);

                    foreach (var a in daa.Distinct().OrderBy(it => it.RC_BRAND))
                    {
                        newCabMain.RC_BRAND = a.RC_BRAND;
                        newCabMain.RC_MODEL = a.RC_MODEL;
                        newCabMain.RC_RATING = a.RC_RATING;
                        newCabMain.RC_GMODULES = a.RC_GMODULES;
                        newCabMain.RC_BMODULES = a.RC_BMODULES;
                        newCabMain.RC_DATE = Convert.ToDateTime(a.RC_DATE);
                        newCabMain.RC_VAC = a.RC_VAC;
                        newCabMain.RC_VDC = a.RC_VDC;
                        newCabMain.RC_PRESENT = a.RC_PRESENT;
                        newCabMain.RC_SA_BRAND = a.RC_SA_BRAND;
                        newCabMain.RC_SA_CAPASITY = a.RC_SA_CAPASITY;
                        newCabMain.RC_SA_CONDITION = a.RC_SA_CONDITION;
                        newCabMain.RC_SA_OHM = a.RC_SA_OHM;
                        newCabMain.RC_SA_CONNECTION = a.RC_SA_CONNECTION;
                        newCabMain.RC_LVD = a.RC_LVD;
                        newCabMain.BATT_BRAND = a.BATT_BRAND;
                        newCabMain.BATT_CAPASITY = a.BATT_CAPASITY;
                        newCabMain.BATT_VOLT = a.BATT_VOLT;
                        newCabMain.BATT_DATE = Convert.ToDateTime(a.BATT_DATE);
                        newCabMain.BATT_CONDITION = a.BATT_CONDITION;
                        newCabMain.AIRCOND_READING = a.AIRCOND_READING;
                        newCabMain.AIRCOND_WORK = a.AIRCOND_WORK;
                        newCabMain.AIRCOND_BROKE = a.AIRCOND_BROKE;
                        newCabMain.AIRCOND_DOOR = a.AIRCOND_DOOR;
                        newCabMain.ALARM = a.ALARM;
                        newCabMain.CHECK_BY = a.CHECK_BY;
                        newCabMain.BATT_EXTRA_LOCKING = a.BATT_EXTRA_LOCKING;
                        newCabMain.PRO_AC_CONNECTION = a.PRO_AC_CONNECTION;

                        newCabMain.TNB_METER = a.TNB_METER;
                        newCabMain.CAB_INS_DATE = Convert.ToDateTime(a.CAB_INS_DATE);
                        newCabMain.ALM_EXT = a.ALM_EXT;
                        newCabMain.ALM_EXT_TO = a.ALM_EXT_TO;
                        newCabMain.ALM_STATUS = a.ALM_STATUS;
                        newCabMain.RC_WARRANTY_END = Convert.ToDateTime(a.RC_WARRANTY_END);
                        newCabMain.BATT_EXIST = a.BATT_EXIST;
                        newCabMain.BATT_WARRANTY_END = Convert.ToDateTime(a.BATT_WARRANTY_END);
                        newCabMain.BATT_NO_CELL = a.BATT_NO_CELL;
                        newCabMain.AIRCOND_IN_OUT = a.AIRCOND_IN_OUT;
                        newCabMain.AIRCOND_HPOWER = a.AIRCOND_HPOWER;
                        newCabMain.REMARKS = a.REMARKS;
                    }
                    newCabMain.SHARE_NO = ShareNo.DC_ID.ToString();
                }
                return Json(new
                {
                    Success = true,
                    pwr = allBRAND,
                    RC_BRAND = newCabMain.RC_BRAND, RC_MODEL = newCabMain.RC_MODEL, RC_RATING = newCabMain.RC_RATING, RC_GMODULES = newCabMain.RC_GMODULES, RC_BMODULES = newCabMain.RC_BMODULES,
                    RC_DATE = newCabMain.RC_DATE, RC_VAC = newCabMain.RC_VAC, RC_VDC = newCabMain.RC_VDC, RC_PRESENT = newCabMain.RC_PRESENT, RC_SA_BRAND = newCabMain.RC_SA_BRAND,
                    RC_SA_CAPASITY = newCabMain.RC_SA_CAPASITY, RC_SA_CONDITION = newCabMain.RC_SA_CONDITION, RC_SA_OHM = newCabMain.RC_SA_OHM, RC_SA_CONNECTION = newCabMain.RC_SA_CONNECTION,
                    RC_LVD = newCabMain.RC_LVD, BATT_BRAND = newCabMain.BATT_BRAND, BATT_CAPASITY = newCabMain.BATT_CAPASITY, BATT_VOLT = newCabMain.BATT_VOLT, BATT_DATE = newCabMain.BATT_DATE,
                    BATT_CONDITION = newCabMain.BATT_CONDITION, AIRCOND_READING = newCabMain.AIRCOND_READING, AIRCOND_WORK = newCabMain.AIRCOND_WORK, AIRCOND_BROKE = newCabMain.AIRCOND_BROKE,
                    AIRCOND_DOOR = newCabMain.AIRCOND_DOOR, ALARM = newCabMain.ALARM, CHECK_BY = newCabMain.CHECK_BY, TNB_METER = newCabMain.TNB_METER, CAB_INS_DATE = newCabMain.CAB_INS_DATE,
                    ALM_EXT = newCabMain.ALM_EXT, ALM_EXT_TO = newCabMain.ALM_EXT_TO, ALM_STATUS = newCabMain.ALM_STATUS, RC_WARRANTY_END = newCabMain.RC_WARRANTY_END, BATT_EXIST = newCabMain.BATT_EXIST,
                    BATT_WARRANTY_END = newCabMain.BATT_WARRANTY_END, BATT_NO_CELL = newCabMain.BATT_NO_CELL, AIRCOND_IN_OUT = newCabMain.AIRCOND_IN_OUT, AIRCOND_HPOWER = newCabMain.AIRCOND_HPOWER,
                    REMARKS = newCabMain.REMARKS, SHARE_NO = newCabMain.SHARE_NO, BATT_EXTRA_LOCKING = newCabMain.BATT_EXTRA_LOCKING
                }, JsonRequestBehavior.AllowGet);
            }
        }

        [HttpPost]
        public ActionResult CheckCab(int G3E_FID)
        {
            using (Entities ctxState = new Entities())
            {

                string CallCab1 = "";
                string CallCab = "";
                var QBRAND = (from d in ctxState.OSP_FIB_FACILITIES
                              join fx in ctxState.AG_OSP_FIB_FACILITIES on d.G3E_FID equals fx.G3E_FID
                              where d.G3E_FID == G3E_FID
                              select new { d.MANUFACTURER, d.CONTRACTOR, d.MODEL, d.INSTALL_YEAR, fx.LATITUDE, fx.LONGITUDE }).Single();
                 
                decimal latitude = Convert.ToDecimal(QBRAND.LATITUDE);
                decimal NewLat = Math.Truncate(1000 * latitude) / 1000;
                string NNLat = NewLat.ToString();
                latitude = Decimal.Round(latitude, 6);

                decimal longitude = Convert.ToDecimal(QBRAND.LONGITUDE);
                decimal NewLong = Math.Truncate(1000 * longitude) / 1000;
                string NNLong = NewLong.ToString();//string.Format("{0:N3}", QBRAND.LONGITUDE);
                longitude = Decimal.Round(longitude, 6);

                var CheckCab = (from d in ctxState.AG_OSP_FIB_FACILITIES
                                join fx in ctxState.OSP_FIB_FACILITIES on d.G3E_FID equals fx.G3E_FID
                                where d.LONGITUDE.Contains(NNLong)
                                && d.LATITUDE.Contains(NNLat)
                                select new { d.G3E_FID, d.CODE, d.LATITUDE, d.LONGITUDE ,fx.PWR_CAB_ID });

                foreach (var ii in CheckCab.Distinct().OrderBy(it => it.CODE))
                {
                    CallCab += ii.G3E_FID + "/" + ii.CODE + "/" + (Math.Truncate(1000000 * Convert.ToDecimal(ii.LATITUDE)) / 1000000) + "/" + (Math.Truncate(1000000 * Convert.ToDecimal(ii.LONGITUDE)) / 1000000) + "/" + ii.PWR_CAB_ID + ":";
                }
                return Json(new
                {
                    Success = true,
                    CallCab = CallCab
                }, JsonRequestBehavior.AllowGet);
            }
        }
        
        public ActionResult DeleteCabInfo(string id)
        {
            bool success;
            
            int data = Convert.ToInt32(id);
            System.Diagnostics.Debug.WriteLine(data);

            Tools tool = new Tools();
            using (Entities ctxData = new Entities())
            {
                var dSHARE = (from d in ctxData.WV_CAB_ADDINFO
                             where d.DATA_NO == data 
                             select new { d.G3E_FID }).Single();

                string sqlCmdGlobalS = "DELETE FROM WV_CAB_SHARE WHERE NE_ID  ='" + dSHARE.G3E_FID + "'";
                success = tool.ExecuteSql(ctxData, sqlCmdGlobalS);

                string sqlCmdGlobal = "DELETE FROM WV_CAB_ADDINFO  WHERE DATA_NO  ='" + data + "'";
                success = tool.ExecuteSql(ctxData, sqlCmdGlobal);
            }

            return RedirectToAction("ListCabInfo");
        }

        public ActionResult NewSave()
        {
            return View();
        }

        [HttpPost]
        public ActionResult AddBrand(string ABrand, string targetTable)
        {
            bool success;

            //WV_CAB_ELCB_BRAND--------------
            //WV_CAB_MCB_BRAND---------------
            //WV_CAB_SA_BRAND
            //WV_CAB_RC_BRAND
            string allBRAND = "";
            Tools tool = new Tools();
            string sqlCmdGlobal = "INSERT INTO " + targetTable + " (BRAND) VALUES (UPPER('" + ABrand + "'))";
            System.Diagnostics.Debug.WriteLine("sqlCmdGlobal : " + sqlCmdGlobal);

            using (Entities ctxData = new Entities())
            {
                success = tool.ExecuteSql(ctxData, sqlCmdGlobal);
            }

            using (Entities ctxState = new Entities())
            {
                if (targetTable == "WV_CAB_ELCB_BRAND")
                {
                    allBRAND = " :N/A:";
                    var QBRAND = (from d in ctxState.WV_CAB_ELCB_BRAND
                                  select new { d.BRAND });

                    foreach (var a in QBRAND.Distinct().OrderBy(it => it.BRAND))
                    {
                        allBRAND += a.BRAND + ":";
                    }
                    allBRAND += "OTHERS:";
                }

                if (targetTable == "WV_CAB_MCB_BRAND")
                {
                    allBRAND = " :N/A:";
                    var QBRAND = (from d in ctxState.WV_CAB_MCB_BRAND
                                  select new { d.BRAND });

                    foreach (var a in QBRAND.Distinct().OrderBy(it => it.BRAND))
                    {
                        allBRAND += a.BRAND + ":";
                    }
                    allBRAND += "OTHERS:";
                }

                if (targetTable == "WV_CAB_SA_BRAND")
                {
                    allBRAND = " :N/A:";
                    var QBRAND = (from d in ctxState.WV_CAB_SA_BRAND
                                  select new { d.BRAND });

                    foreach (var a in QBRAND.Distinct().OrderBy(it => it.BRAND))
                    {
                        allBRAND += a.BRAND + ":";
                    }
                    allBRAND += "OTHERS:";
                }

                if (targetTable == "WV_CAB_AR_BRAND")
                {
                    allBRAND = " :N/A:";
                    var QBRAND = (from d in ctxState.WV_CAB_AR_BRAND
                                  select new { d.BRAND });

                    foreach (var a in QBRAND.Distinct().OrderBy(it => it.BRAND))
                    {
                        allBRAND += a.BRAND + ":";
                    }
                    allBRAND += "OTHERS:";
                }
            }

            return Json(new
            {
                BRAND = allBRAND,
                Success = true
            }, JsonRequestBehavior.AllowGet);
        }

        //ADD HERE FOR REPORT OPFI ENZ
        public ActionResult GetZoneList()
        {
            string listZone = "";
            using (Entities ctxData = new Entities())
            {

                var queryZone = from p in ctxData.WV_CAB_ZONE
                                select new { p.ZONE };

                foreach (var a in queryZone.Distinct().OrderBy(it => it.ZONE))
                {
                    listZone = listZone + a.ZONE + "|";

                }
            }


            return Json(new
            {
                Success = true,
                listZone = listZone
            }, JsonRequestBehavior.AllowGet); //

        }
        //END ADD HERE FOR REPORT OPFI ENZ

        [HttpPost]
        public ActionResult AddModel(string ABrand, string AModel, string targetTable)
        {
            bool success;

            //WV_CAB_ELCB_BRAND--------------
            //WV_CAB_MCB_BRAND---------------
            //WV_CAB_SA_BRAND
            //WV_CAB_RC_BRAND
            string allBRAND = "";

            Tools tool = new Tools();
            string sqlCmdGlobal = "INSERT INTO " + targetTable + " (BRAND, MODEL) VALUES (UPPER('" + ABrand + "'), UPPER('" + AModel + "'))";
            System.Diagnostics.Debug.WriteLine("sqlCmdGlobal : " + sqlCmdGlobal);

            using (Entities ctxData = new Entities())
            {
                success = tool.ExecuteSql(ctxData, sqlCmdGlobal);
            }

            using (Entities ctxState = new Entities())
            {

                    allBRAND = " :N/A:";
                    var QBRAND = (from d in ctxState.WV_CAB_RC_BRAND
                                  select new { d.BRAND });

                    foreach (var a in QBRAND.Distinct().OrderBy(it => it.BRAND))
                    {
                        allBRAND += a.BRAND + ":";
                    }
                    allBRAND += "OTHERS:";               
            }

            return Json(new
            {
                BRAND = allBRAND,
                Success = true
            }, JsonRequestBehavior.AllowGet);
        }

        public ActionResult ReportCabInfo(string searchKey, string pttID, string excID, int? page)
        {
            WebView.WebService._base myWebService;
            myWebService = new WebService._base();

            WebService._base.DataCabAddInfo CabInfo1 = new WebService._base.DataCabAddInfo();

            using (Entities ctxData = new Entities())
            {
                List<SelectListItem> list1 = new List<SelectListItem>();
                var query1 = from p in ctxData.WV_STATE_MAST
                             select new { p.STATE_NAME };
                list1.Add(new SelectListItem() { Text = "", Value = "" });
                foreach (var a in query1.Distinct().OrderBy(it => it.STATE_NAME))
                {
                    list1.Add(new SelectListItem() { Text = a.STATE_NAME, Value = a.STATE_NAME });
                }
                ViewBag.StateId = list1;

                List<SelectListItem> list2 = new List<SelectListItem>();
                var query3 = from p in ctxData.WV_EXC_MAST
                             select new { p.PTT_ID };
                list2.Add(new SelectListItem() { Text = "", Value = "" });
                foreach (var a in query3.Distinct().OrderBy(it => it.PTT_ID))
                {
                    list2.Add(new SelectListItem() { Text = a.PTT_ID, Value = a.PTT_ID });
                }
                ViewBag.PttId = list2;

                List<SelectListItem> list3 = new List<SelectListItem>();
                var query4 = from p in ctxData.WV_EXC_MAST
                             select new { p.EXC_ABB };
                list3.Add(new SelectListItem() { Text = "", Value = "" });
                foreach (var a in query4.Distinct().OrderBy(it => it.EXC_ABB))
                {
                    list3.Add(new SelectListItem() { Text = a.EXC_ABB, Value = a.EXC_ABB });
                }
                ViewBag.Exc = list3;

                List<SelectListItem> list4 = new List<SelectListItem>();
                var query5 = from p in ctxData.WV_CAB_ZONE
                             select new { p.ZONE };
                list4.Add(new SelectListItem() { Text = "", Value = "" });
                foreach (var a in query5.Distinct().OrderBy(it => it.ZONE))
                {
                    list4.Add(new SelectListItem() { Text = a.ZONE, Value = a.ZONE });
                }
                ViewBag.Zone = list4;

                List<SelectListItem> list5 = new List<SelectListItem>();
                var query6 = (from d in ctxData.WV_CAB_RC_BRAND
                              group d by d.BRAND into newGroup
                              orderby newGroup.Key
                              select newGroup);
                list5.Add(new SelectListItem() { Text = "", Value = "" });
                foreach (var a in query6)
                {
                    list5.Add(new SelectListItem() { Text = a.Key, Value = a.Key });
                }
                ViewBag.RC = list5;

                List<SelectListItem> list6 = new List<SelectListItem>();
                var query7 = from p in ctxData.WV_CAB_SA_BRAND //------------BATTERY BRAND
                             select new { p.BRAND };
                list6.Add(new SelectListItem() { Text = "", Value = "" });
                foreach (var a in query7.Distinct().OrderBy(it => it.BRAND))
                {
                    list6.Add(new SelectListItem() { Text = a.BRAND, Value = a.BRAND });
                }
                ViewBag.Batt = list6;

                List<SelectListItem> list8 = new List<SelectListItem>();
                var query8 = from p in ctxData.OSP_FIB_FACILITIES //------------BATTERY BRAND
                             select new { p.MODEL };
                list8.Add(new SelectListItem() { Text = "", Value = "" });
                foreach (var a in query8.Distinct().OrderBy(it => it.MODEL))
                {
                    list8.Add(new SelectListItem() { Text = a.MODEL, Value = a.MODEL });
                }
                ViewBag.Model = list8;

                List<SelectListItem> list9 = new List<SelectListItem>();
                var query9 = from p in ctxData.OSP_FIB_FACILITIES //------------BATTERY BRAND
                             select new { p.MANUFACTURER };
                list9.Add(new SelectListItem() { Text = "", Value = "" });
                foreach (var a in query9.Distinct().OrderBy(it => it.MANUFACTURER))
                {
                    list9.Add(new SelectListItem() { Text = a.MANUFACTURER, Value = a.MANUFACTURER });
                }
                ViewBag.Manufacturer = list9;


                List<SelectListItem> list10 = new List<SelectListItem>();

                var userInfo = (from c in ctxData.WV_USER
                                join fx in ctxData.WV_GROUP on c.GROUPID equals fx.GRP_ID
                                where c.USERNAME == User.Identity.Name
                                select new { fx.GRPNAME }).Single();

                if (userInfo.GRPNAME == "ADMIN")
                {
                    list10.Add(new SelectListItem() { Text = "ALL", Value = "ADMIN" });
                    list10.Add(new SelectListItem() { Text = "AND", Value = "AND" });
                    list10.Add(new SelectListItem() { Text = "ENZ", Value = "ENZ" });
                }
                else if (userInfo.GRPNAME == "AND")
                {
                    list10.Add(new SelectListItem() { Text = "AND", Value = "AND" });
                    list10.Add(new SelectListItem() { Text = "ALL", Value = "ADMIN" });
                    list10.Add(new SelectListItem() { Text = "ENZ", Value = "ENZ" });
                }
                else if (userInfo.GRPNAME == "ENZ")
                {
                    list10.Add(new SelectListItem() { Text = "AND", Value = "AND" });
                    list10.Add(new SelectListItem() { Text = "ALL", Value = "ADMIN" });
                    list10.Add(new SelectListItem() { Text = "ENZ", Value = "ENZ" });
                }
                ViewBag.GroupUser = list10;
            }

            List<SelectListItem> list11 = new List<SelectListItem>();
            ViewBag.listZone = list11;
            ViewBag.zoneID2 = list11;
            ViewBag.scGroup2 = list11;

            string input = "\\\\adsvr";

            string output = String.Format("http:{0}", input.Replace("\\d$\\x\\y", String.Empty).Replace("\\", "/"));




            ViewBag.output = output;

            //return View();
            int pageSize = 20;
            int pageNumber = (page ?? 1);
            return View();
        }
        //public ActionResult ReportCabInfo(string searchKey, string pttID, string excID, int? page)
        //{
        //    WebView.WebService._base myWebService;
        //    myWebService = new WebService._base();

        //    WebService._base.DataCabAddInfo CabInfo1 = new WebService._base.DataCabAddInfo();

        //    using (Entities ctxData = new Entities())
        //    {
        //        List<SelectListItem> list1 = new List<SelectListItem>();
        //        var query1 = from p in ctxData.WV_STATE_MAST
        //                     select new { p.STATE_NAME };
        //        list1.Add(new SelectListItem() { Text = "", Value = "" });
        //        foreach (var a in query1.Distinct().OrderBy(it => it.STATE_NAME))
        //        {
        //            list1.Add(new SelectListItem() { Text = a.STATE_NAME, Value = a.STATE_NAME });
        //        }
        //        ViewBag.StateId = list1;

        //        List<SelectListItem> list2 = new List<SelectListItem>();
        //        var query3 = from p in ctxData.WV_EXC_MAST
        //                     select new { p.PTT_ID };
        //        list2.Add(new SelectListItem() { Text = "", Value = "" });
        //        foreach (var a in query3.Distinct().OrderBy(it => it.PTT_ID))
        //        {
        //            list2.Add(new SelectListItem() { Text = a.PTT_ID, Value = a.PTT_ID });
        //        }
        //        ViewBag.PttId = list2;

        //        List<SelectListItem> list3 = new List<SelectListItem>();
        //        var query4 = from p in ctxData.WV_EXC_MAST
        //                     select new { p.EXC_ABB };
        //        list3.Add(new SelectListItem() { Text = "", Value = "" });
        //        foreach (var a in query4.Distinct().OrderBy(it => it.EXC_ABB))
        //        {
        //            list3.Add(new SelectListItem() { Text = a.EXC_ABB, Value = a.EXC_ABB });
        //        }
        //        ViewBag.Exc = list3;

        //        List<SelectListItem> list4 = new List<SelectListItem>();
        //        var query5 = from p in ctxData.WV_CAB_ZONE
        //                     select new { p.ZONE };
        //        list4.Add(new SelectListItem() { Text = "", Value = "" });
        //        foreach (var a in query5.Distinct().OrderBy(it => it.ZONE))
        //        {
        //            list4.Add(new SelectListItem() { Text = a.ZONE, Value = a.ZONE });
        //        }
        //        ViewBag.Zone = list4;

        //        List<SelectListItem> list5 = new List<SelectListItem>();
        //        var query6 = (from d in ctxData.WV_CAB_RC_BRAND
        //                      group d by d.BRAND into newGroup
        //                      orderby newGroup.Key
        //                      select newGroup);
        //        list5.Add(new SelectListItem() { Text = "", Value = "" });
        //        foreach (var a in query6)
        //        {
        //            list5.Add(new SelectListItem() { Text = a.Key, Value = a.Key });
        //        }
        //        ViewBag.RC = list5;

        //        List<SelectListItem> list6 = new List<SelectListItem>();
        //        var query7 = from p in ctxData.WV_CAB_SA_BRAND //------------BATTERY BRAND
        //                     select new { p.BRAND };
        //        list6.Add(new SelectListItem() { Text = "", Value = "" });
        //        foreach (var a in query7.Distinct().OrderBy(it => it.BRAND))
        //        {
        //            list6.Add(new SelectListItem() { Text = a.BRAND, Value = a.BRAND });
        //        }
        //        ViewBag.Batt = list6;

        //        List<SelectListItem> list8 = new List<SelectListItem>();
        //        var query8 = from p in ctxData.OSP_FIB_FACILITIES //------------BATTERY BRAND
        //                     select new { p.MODEL };
        //        list8.Add(new SelectListItem() { Text = "", Value = "" });
        //        foreach (var a in query8.Distinct().OrderBy(it => it.MODEL))
        //        {
        //            list8.Add(new SelectListItem() { Text = a.MODEL, Value = a.MODEL });
        //        }
        //        ViewBag.Model = list8;

        //        List<SelectListItem> list9 = new List<SelectListItem>();
        //        var query9 = from p in ctxData.OSP_FIB_FACILITIES //------------BATTERY BRAND
        //                     select new { p.MANUFACTURER };
        //        list9.Add(new SelectListItem() { Text = "", Value = "" });
        //        foreach (var a in query9.Distinct().OrderBy(it => it.MANUFACTURER))
        //        {
        //            list9.Add(new SelectListItem() { Text = a.MANUFACTURER, Value = a.MANUFACTURER });
        //        }
        //        ViewBag.Manufacturer = list9;


        //        List<SelectListItem> list10 = new List<SelectListItem>();
                
        //        var userInfo = (from c in ctxData.WV_USER
        //                        join fx in ctxData.WV_GROUP on c.GROUPID equals fx.GRP_ID
        //                        where c.USERNAME == User.Identity.Name
        //                        select new { fx.GRPNAME } ).Single();

        //        if (userInfo.GRPNAME == "ADMIN")
        //        {
        //            list10.Add(new SelectListItem() { Text = "ALL", Value = "ADMIN" });
        //            list10.Add(new SelectListItem() { Text = "AND", Value = "AND" });
        //            list10.Add(new SelectListItem() { Text = "ENZ", Value = "ENZ" });
        //        }
        //        else if (userInfo.GRPNAME == "AND")
        //        {
        //            list10.Add(new SelectListItem() { Text = "AND", Value = "AND" });
        //            list10.Add(new SelectListItem() { Text = "ALL", Value = "ADMIN" });
        //            list10.Add(new SelectListItem() { Text = "PRM", Value = "PRM" });
        //        }
        //        else if (userInfo.GRPNAME == "PRM")
        //        {
        //            list10.Add(new SelectListItem() { Text = "PRM", Value = "PRM" });
        //            list10.Add(new SelectListItem() { Text = "ALL", Value = "ADMIN" });
        //            list10.Add(new SelectListItem() { Text = "AND", Value = "AND" });
        //        }
        //        ViewBag.GroupUser = list10;

        //    }
        //    string input = "\\\\adsvr";

        //    string output = String.Format("http:{0}", input.Replace("\\d$\\x\\y", String.Empty).Replace("\\", "/"));

        //    List<SelectListItem> zone = new List<SelectListItem>();
        //    ViewBag.ZoneENZ = zone;

        //    ViewBag.output = output;

        //    //return View();
        //    int pageSize = 20;
        //    int pageNumber = (page ?? 1);
        //    return View();
        //}

        public ActionResult GetZoneENZ()
        {
            string listENZ = "";
            using (Entities ctxData = new Entities())
            {
                var queryZone = (from p in ctxData.WV_CAB_ZONE
                                 select new {p.ZONE});
                
                foreach (var a in queryZone.Distinct().OrderBy(it => it.ZONE))
                {
                    listENZ = listENZ + a.ZONE + "|";
                }
            }

            return Json(new
            {
                Success = true,
                listENZ = listENZ
            }, JsonRequestBehavior.AllowGet); //
        }

        //public ActionResult ReportFacilitiesSummary(string id, int? page)
        public ActionResult ReportFacilitiesSummary(string scGroup, string GroupZone, string id, int? page)
        {
            WebView.WebService._base myWebService;
            myWebService = new WebService._base();

            WebService._base.DataCabAddInfo CabInfo1 = new WebService._base.DataCabAddInfo();

            if (scGroup == "ADMIN")
            {
                CabInfo1 = myWebService.GetReportCabAddInfo(0, 50000);
            }
            else if (scGroup == "AND")
            {
                CabInfo1 = myWebService.GetReportCabAddInfoAND(0, 50000);
            }
            else if (scGroup == "ENZ")
            {
            //{
            //    CabInfo1 = myWebService.GetReportCabAddInfoPRM(0, 50000);
                //ADD HERE FOR OPFI REPORT ENZ
                if (GroupZone != "Select")
                {
                    CabInfo1 = myWebService.GetReportCabAddInfoPRMEXC(0, 50000, GroupZone);
                }
                else
                {
                    CabInfo1 = myWebService.GetReportCabAddInfoPRM(0, 50000, GroupZone);
                }
            }

            ViewBag.Page = 1;
            int pageSize = 20;
            int pageNumber = (page ?? 1);

            if (id == "EXCEL")
            {
                this.Response.AddHeader("Content-Disposition", "attachment; filename=SummaryFacilities.xls");
                this.Response.ContentType = "application/vnd.ms-excel";
            }

            return View(CabInfo1.DataCabAddInfoList.ToPagedList(pageNumber, pageSize));
        }

        public ActionResult ReportFacilitiesDetails(string searchKey, string pwrCab, string StateID, string zoneID, string pttID, string excID, string rcID, string status, string Model, string Manufacturer, string type, int? page)
        {
            WebView.WebService._base myWebService;
            myWebService = new WebService._base();

            WebService._base.DataCabAddInfoReport CabInfo1 = new WebService._base.DataCabAddInfoReport();
            
            if (searchKey != null || StateID != null || zoneID != null || pttID != null || pttID != null || pwrCab != null || rcID != null || status != null || Model != null || Manufacturer != null)
            {
                if (searchKey == "" && StateID == "" && zoneID == "" && pttID == "" && excID == "" && pwrCab == "" && rcID == "" && status == "" && Model == "" && Manufacturer == "")
                {
                    CabInfo1 = myWebService.GetReportCabAddInfoDetail(0, 100000, null, null, null, null, null, null, null, null, null, null);
                }
                else if (searchKey != "" || StateID != "" || zoneID != "" || pttID != "" || excID != "" || pwrCab != "" || rcID != "" || status != "" || Model != "" || Manufacturer != "")
                {
                    CabInfo1 = myWebService.GetReportCabAddInfoDetail(0, 100000, searchKey, StateID, zoneID, pttID, excID, pwrCab.ToUpper(), rcID, status, Model, Manufacturer);
                }
            }
            else
            {
                CabInfo1 = myWebService.GetReportCabAddInfoDetail(0, 100000, null, null, null, null, null, null, null, null, null, null);
            }
            //CabInfo1 = myWebService.GetReportCabAddInfoDetail(0, 100000, searchKey, StateID, zoneID, pttID, excID, pwrCab.ToUpper(), status, Model, Manufacturer);

            ViewBag.Page = 1;
            int pageSize = 100000;
            int pageNumber = (page ?? 1);

            if (type == "EXCEL")
            {
                this.Response.Clear();
                this.Response.Buffer = true;
                this.Response.AddHeader("Content-Disposition", "attachment; filename=SummaryFacilities.xls");
                //this.Response.ContentType = "application/vnd.ms-excel";
                Response.ContentEncoding = System.Text.Encoding.GetEncoding("UTF-8");
                Response.ContentType = "application/ms-excel";
                                //Response.Charset = "UTF-8"; 
                System.Globalization.CultureInfo myCItrad = new System.Globalization.CultureInfo("EN-US", true);
                System.IO.StringWriter oStringWriter = new System.IO.StringWriter(myCItrad);
                System.Web.UI.HtmlTextWriter oHtmlTextWriter = new System.Web.UI.HtmlTextWriter(oStringWriter);
            }

            return View(CabInfo1.DataCabAddInfoListReport.ToPagedList(pageNumber, pageSize));
        }

        public ActionResult ReportFacilitiesDetailsGroup(string searchKey, string pwrCab, string StateID, string zoneID, string pttID, string excID, string rcID, string status, string Model, string Manufacturer, string scGroup, string type, int? page)
        {
            WebView.WebService._base myWebService;
            myWebService = new WebService._base();

            WebService._base.DataCabAddInfoReport CabInfo1 = new WebService._base.DataCabAddInfoReport();

            if (searchKey != null || StateID != null || zoneID != null || pttID != null || pttID != null || pwrCab != null || rcID != null || status != null || Model != null || Manufacturer != null)
            {
                if (searchKey == "" && StateID == "" && zoneID == "" && pttID == "" && excID == "" && pwrCab == "" && rcID == "" && status == "" && Model == "" && Manufacturer == "")
                {
                    CabInfo1 = myWebService.GetReportCabAddInfoDetailGroup(0, 100000, null, null, null, null, null, null, null, null, null, null, scGroup);
                }
                else if (searchKey != "" || StateID != "" || zoneID != "" || pttID != "" || excID != "" || pwrCab != "" || rcID != "" || status != "" || Model != "" || Manufacturer != "")
                {
                    CabInfo1 = myWebService.GetReportCabAddInfoDetailGroup(0, 100000, searchKey, StateID, zoneID, pttID, excID, pwrCab.ToUpper(), rcID, status, Model, Manufacturer, scGroup);
                }
            }
            else
            {
                CabInfo1 = myWebService.GetReportCabAddInfoDetailGroup(0, 100000, null, null, null, null, null, null, null, null, null, null, scGroup);
            }
            //CabInfo1 = myWebService.GetReportCabAddInfoDetail(0, 100000, searchKey, StateID, zoneID, pttID, excID, pwrCab.ToUpper(), status, Model, Manufacturer);

            ViewBag.Page = 1;
            int pageSize = 100000;
            int pageNumber = (page ?? 1);

            if (type == "EXCEL")
            {
                this.Response.Clear();
                this.Response.Buffer = true;
                this.Response.AddHeader("Content-Disposition", "attachment; filename=SummaryFacilities.xls");
                //this.Response.ContentType = "application/vnd.ms-excel";
                Response.ContentEncoding = System.Text.Encoding.GetEncoding("UTF-8");
                Response.ContentType = "application/ms-excel";
                //Response.Charset = "UTF-8"; 
                System.Globalization.CultureInfo myCItrad = new System.Globalization.CultureInfo("EN-US", true);
                System.IO.StringWriter oStringWriter = new System.IO.StringWriter(myCItrad);
                System.Web.UI.HtmlTextWriter oHtmlTextWriter = new System.Web.UI.HtmlTextWriter(oStringWriter);
            }

            return View(CabInfo1.DataCabAddInfoListReport.ToPagedList(pageNumber, pageSize));
        }
        
        public ActionResult ReportFacilitiesSummaryEXC(string scGroup, string GroupZone, string id, int? page)
        {
            WebView.WebService._base myWebService;
            myWebService = new WebService._base();

            WebService._base.DataCabAddInfo CabInfo1 = new WebService._base.DataCabAddInfo();


            if (scGroup == "PRM" || scGroup=="ENZ")
            {
                CabInfo1 = myWebService.GetReportCabAddInfoPRMEXC(0, 50000, GroupZone);
            }

            ViewBag.Page = 1;
            int pageSize = 20;
            int pageNumber = (page ?? 1);

            if (id == "EXCEL")
            {
                this.Response.AddHeader("Content-Disposition", "attachment; filename=SummaryFacilities.xls");
                this.Response.ContentType = "application/vnd.ms-excel";
            }

            return View(CabInfo1.DataCabAddInfoList.ToPagedList(pageNumber, pageSize));
        }
        // END ADD HERE FOR OPFI REPORT ENZ

        //ADD HERE FOR OPFI REPORT ENZ
        public ActionResult ReportFacilitiesSummaryZONE(string scGroup, string GroupZone, string id, int? page)
        {
            WebView.WebService._base myWebService;
            myWebService = new WebService._base();

            WebService._base.DataCabAddInfo CabInfo1 = new WebService._base.DataCabAddInfo();

            CabInfo1 = myWebService.GetReportCabAddInfoPRM(0, 50000, GroupZone);

            ViewBag.Page = 1;
            int pageSize = 150;
            int pageNumber = (page ?? 1);

            if (id == "EXCEL")
            {
                this.Response.AddHeader("Content-Disposition", "attachment; filename=SummaryFacilities.xls");
                this.Response.ContentType = "application/vnd.ms-excel";
            }

            return View(CabInfo1.DataCabAddInfoList.ToPagedList(pageNumber, pageSize));
        }
        //END ADD HERE FOR OPFI REPORT ENZ
    }
}
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
using System.Net.Mail;
using WebView.Models;
using WebView.Library;
using System.IO;
using PagedList;

namespace WebView.Controllers
{
    public class UserSendController : Controller
    {
        //
        // GET: /UserSend/

        public ActionResult Index()
        {
            return View();
        }

        public ActionResult Send()
        {
            WebView.WebService._base myWebService;
            myWebService = new WebService._base();

            UserSendModel.ViewModel model = new UserSendModel.ViewModel { AvailableProducts = myWebService.getUser(0, 100), RequestedProducts = new List<UserSendModel.User>(), ListSend = new List<int>() };

            return View(model);
        }

        [HttpPost]
        public ActionResult Send(UserSendModel.ViewModel m)
        {
            WebView.WebService._base myWebService;
            myWebService = new WebService._base();

            List<int> list = new List<int>();
            for (int i = 0; i < m.RequestedSelected.Length; i++)
            {
                list.Add(m.RequestedSelected[i]);
            }
            UserSendModel.ViewModel model = new UserSendModel.ViewModel { AvailableProducts = myWebService.getUser(0, 100), RequestedProducts = new List<UserSendModel.User>(), ListSend = list };
            return View(model);

        }



    }
}

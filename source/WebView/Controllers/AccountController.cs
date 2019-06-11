using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Web;
using System.Web.Mvc;
using System.Web.Security;
using WebView.Library;
using WebView.Models.Account;
using PagedList;
using System.Net.Mail;
using System.Configuration;

namespace WebView.Controllers
{
    public class AccountController : Controller
    {

        private string visionael = ConfigurationManager.AppSettings.Get("VISIONAEL_URL");
        //
        // GET: /Account/LogIn
        public class Crypto
        {
            private static byte[] _salt = Encoding.ASCII.GetBytes("o6806642kbM7c5");
            private const string sharedSecret = "NEPS";

            /// <summary>
            /// Encrypt the given string using AES.  The string can be decrypted using 
            /// DecryptStringAES().  The sharedSecret parameters must match.
            /// </summary>
            /// <param name="plainText">The text to encrypt.</param>
            /// <param name="sharedSecret">A password used to generate a key for encryption.</param>

            public static string EncryptStringAES(string plainText)
            {
                if (string.IsNullOrEmpty(plainText))
                    return "";

                string outStr = null;                       // Encrypted string to return
                RijndaelManaged aesAlg = null;              // RijndaelManaged object used to encrypt the data.

                try
                {
                    // generate the key from the shared secret and the salt
                    Rfc2898DeriveBytes key = new Rfc2898DeriveBytes(sharedSecret, _salt);

                    // Create a RijndaelManaged object
                    // with the specified key and IV.
                    aesAlg = new RijndaelManaged();
                    aesAlg.Key = key.GetBytes(aesAlg.KeySize / 8);
                    aesAlg.IV = key.GetBytes(aesAlg.BlockSize / 8);

                    // Create a decrytor to perform the stream transform.
                    ICryptoTransform encryptor = aesAlg.CreateEncryptor(aesAlg.Key, aesAlg.IV);

                    // Create the streams used for encryption.
                    using (MemoryStream msEncrypt = new MemoryStream())
                    {
                        using (CryptoStream csEncrypt = new CryptoStream(msEncrypt, encryptor, CryptoStreamMode.Write))
                        {
                            using (StreamWriter swEncrypt = new StreamWriter(csEncrypt))
                            {

                                //Write all data to the stream.
                                swEncrypt.Write(plainText);
                            }
                        }
                        outStr = Convert.ToBase64String(msEncrypt.ToArray());
                    }
                }
                finally
                {
                    // Clear the RijndaelManaged object.
                    if (aesAlg != null)
                        aesAlg.Clear();
                }

                // Return the encrypted bytes from the memory stream.
                return outStr;
            }

            /// <summary>
            /// Decrypt the given string.  Assumes the string was encrypted using 
            /// EncryptStringAES(), using an identical sharedSecret.
            /// </summary>
            /// <param name="cipherText">The text to decrypt.</param>
            /// <param name="sharedSecret">A password used to generate a key for decryption.</param>
            public static string DecryptStringAES(string cipherText)
            {
                if (string.IsNullOrEmpty(cipherText))
                    return "";

                // Declare the RijndaelManaged object
                // used to decrypt the data.
                RijndaelManaged aesAlg = null;

                // Declare the string used to hold
                // the decrypted text.
                string plaintext = null;

                try
                {
                    // generate the key from the shared secret and the salt
                    Rfc2898DeriveBytes key = new Rfc2898DeriveBytes(sharedSecret, _salt);

                    // Create a RijndaelManaged object
                    // with the specified key and IV.
                    aesAlg = new RijndaelManaged();
                    aesAlg.Key = key.GetBytes(aesAlg.KeySize / 8);
                    aesAlg.IV = key.GetBytes(aesAlg.BlockSize / 8);

                    // Create a decrytor to perform the stream transform.
                    ICryptoTransform decryptor = aesAlg.CreateDecryptor(aesAlg.Key, aesAlg.IV);
                    // Create the streams used for decryption.                
                    byte[] bytes = Convert.FromBase64String(cipherText);
                    using (MemoryStream msDecrypt = new MemoryStream(bytes))
                    {
                        using (CryptoStream csDecrypt = new CryptoStream(msDecrypt, decryptor, CryptoStreamMode.Read))
                        {
                            using (StreamReader srDecrypt = new StreamReader(csDecrypt))

                                // Read the decrypted bytes from the decrypting stream
                                // and place them in a string.
                                plaintext = srDecrypt.ReadToEnd();
                        }
                    }
                }
                finally
                {
                    // Clear the RijndaelManaged object.
                    if (aesAlg != null)
                        aesAlg.Clear();
                }

                return plaintext;
            }
        }

        public ActionResult LogIn()
        {
            FormsAuthentication.SignOut();
            return View();//
        }

        //
        // POST: /Account/LogIn

        [HttpPost]
        public ActionResult LogIn(LogInModel model, string returnUrl)
        {
            WebView.WebService._base myWebService;
            myWebService = new WebService._base();

            UserModel user;

            string password = Crypto.EncryptStringAES(model.Password);

            model.UserName = model.UserName.ToUpper();
            if (ModelState.IsValid)
            {
                user = myWebService.ValidateUser(model.UserName, model.Password);
                //user = myWebService.ValidateUser(model.UserName, password);
                if (user.IsValidUser)
                {
                    ////##  Add Cookie Method 1 ##////////////////
                    FormsAuthentication.SetAuthCookie(model.UserName, model.RememberMe);

                    Session["UserInfo"] = user;

                    string role = "";
                    string uAuditTrail;
                    List<string> roles = new List<string>();
                    Tools tool = new Tools();
                    using (Entities ctxData = new Entities())
                    {
                        var queryUser = (from p in ctxData.WV_USER
                                         where p.USERNAME.ToUpper() == model.UserName.ToUpper() || p.USERNAME.ToLower() == model.UserName.ToLower()
                                         select p).Single();

                        uAuditTrail = queryUser.USERNAME;

                        var queryUserRole = (from fx in ctxData.WV_GROUP
                                             join fxx in ctxData.WV_GRP_ROLE on fx.GRPNAME equals fxx.GRPNAME
                                             where fx.GRP_ID == queryUser.GROUPID
                                             select fxx).Single();
                        user.FullName = queryUser.FULL_NAME;
                        role = queryUserRole.ROLENAME;
                    }

                    roles.Add(role);
                    string passPhrase = "preAuthpassword";
                    string username = User.Identity.Name;
                    string url = Encryptor.GetURL(passPhrase, "10.41.61.177", "8080", model.UserName, roles);
                    System.Diagnostics.Debug.WriteLine(url);

                    Session["ISP_INFO"] = visionael + "/nrm/vfd/MainPage.iface" + url;
                    Session["ISP_LOGIN_INFO"] = url;
                    Session["JASPER"] = "http://10.41.61.177:8080/jasperserver/j_spring_security_check?j_username=neps-user&j_password=nepsuser";

                    if (Url.IsLocalUrl(returnUrl) && returnUrl.Length > 1 && returnUrl.StartsWith("/")
                        && !returnUrl.StartsWith("//") && !returnUrl.StartsWith("/\\"))
                    {
                        return Redirect(returnUrl);
                    }
                    else
                    {
                        myWebService.AddUserAudit(uAuditTrail);
                        return RedirectToAction("Index", "Home");
                    }
                }
                else
                {
                    ModelState.AddModelError("", "The user name or password provided is incorrect.");
                }
            }

            // If we got this far, something failed, redisplay form
            return View(model);
        }

        class Encryptor
        {
            private static byte[] CreateKeyBytes(string passPhrase)
            {
                char[] keyChars = passPhrase.ToCharArray();
                byte[] keyByte = new byte[16];
                for (int i = 0; i < 16; i++)
                {
                    if (keyChars.Length > i)
                        keyByte[i] = (byte)keyChars[i];
                }

                if (keyChars.Length < 16)
                {
                    for (int i = keyChars.Length; i < 16; i++)
                        keyByte[i] = 0x00;
                }

                return keyByte;

            }

            public static string Encrypt(Encoding encoding, string strtoencrypt, string key, string iv, CipherMode mode, PaddingMode padding, int blocksize)
            {

                var mstream = new MemoryStream();
                using (var aes = new AesManaged())
                {
                    var keybytes = CreateKeyBytes(key);
                    aes.BlockSize = blocksize;
                    aes.KeySize = keybytes.Length * 8;
                    aes.Key = keybytes;

                    aes.IV = CreateKeyBytes(iv);
                    aes.Mode = mode;
                    aes.Padding = padding;


                    using (var cstream = new CryptoStream(mstream, aes.CreateEncryptor(aes.Key, aes.IV), CryptoStreamMode.Write))
                    {
                        var bytesToEncrypt = encoding.GetBytes(strtoencrypt);
                        cstream.Write(bytesToEncrypt, 0, bytesToEncrypt.Length);
                        cstream.FlushFinalBlock();
                    }

                }

                var encrypted = mstream.ToArray();
                return Convert.ToBase64String(encrypted);
            }


            public static string GetURL(string passPhrase, string hostname, string port, string username, List<string> roles)
            {

                string output = "";


                string rolesCommaSeparated = ListToCommaSeparated(roles);
                var key = passPhrase;
                var iv = "";

                string encryptedUsername = Encrypt(Encoding.ASCII, username, key, iv, CipherMode.CBC, PaddingMode.PKCS7, 128);
                string encryptedRoles = Encrypt(Encoding.ASCII, rolesCommaSeparated, key, iv, CipherMode.CBC, PaddingMode.PKCS7, 128);


                //output = string.Format("http://{0}:{1}/nrm?user={2}&roles={3}", hostname, port,
                //    HttpUtility.UrlEncode(encryptedUsername), HttpUtility.UrlEncode(encryptedRoles));

                output = string.Format("?user={2}&roles={3}", hostname, port,
                    HttpUtility.UrlEncode(encryptedUsername), HttpUtility.UrlEncode(encryptedRoles));

                return output;
            }

            private static string ListToCommaSeparated(List<string> roles)
            {
                string output = "";

                int i = 0;
                foreach (string role in roles)
                {
                    if (i++ > 0)
                        output += ",";
                    output += role;
                }
                return output;
            }




        }
        //
        // GET: /Account/LogOff

        public ActionResult LogOff()
        {
            FormsAuthentication.SignOut();

            return RedirectToAction("Index", "Home");
        }

        //
        // GET: /Account/Register

        public ActionResult Register()
        {
            List<SelectListItem> list = new List<SelectListItem>();
            using (Entities ctxData = new Entities())
            {
                var query = from p in ctxData.WV_GROUP
                            orderby p.FULL_NAME
                            select new { Text = p.FULL_NAME, Value = p.GRP_ID };

                foreach (var a in query)
                {
                    list.Add(new SelectListItem() { Text = a.Text, Value = a.Value.ToString() });
                }

                ViewBag.groupID = list;
            }
            List<SelectListItem> list2 = new List<SelectListItem>();
            using (Entities ctxData = new Entities())
            {
                var query1 = from p in ctxData.WV_ROLE
                             orderby p.ROLENAME
                             select new { Text = p.ROLENAME, Value = p.ROLENAME };

                foreach (var a in query1)
                {
                    list2.Add(new SelectListItem() { Text = a.Text, Value = a.Value });
                }

                ViewBag.roleID = list2;
            }
            List<SelectListItem> list3 = new List<SelectListItem>();
            list3.Add(new SelectListItem() { Text = "AUTO", Value = "AUTO" });
            list3.Add(new SelectListItem() { Text = "MANUAL", Value = "MANUAL" });
            list3.Add(new SelectListItem() { Text = "N/A", Value = "N/A" });

            ViewBag.handoverState = list3;

            List<SelectListItem> list4 = new List<SelectListItem>();
            list4.Add(new SelectListItem() { Text = "ALL", Value = "ALL" });
            using (Entities ctxData = new Entities())
            {
                var query3 = from p in ctxData.WV_EXC_MAST
                             orderby p.PTT_ID
                             select new { Text = p.PTT_ID, Value = p.PTT_ID };

                foreach (var a in query3.Distinct())
                {
                    list4.Add(new SelectListItem() { Text = a.Text, Value = a.Value });
                }

                ViewBag.pttID = list4.OrderBy(x => x.Value);
            }
            //REGION
            List<SelectListItem> list6 = new List<SelectListItem>();
            //list6.Add(new SelectListItem() { Text = "ALL", Value = "ALL" });
            using (Entities ctxData = new Entities())
            {
                var query4 = from p in ctxData.WV_REGION_MAST
                             orderby p.REGION_ID
                             select new { Text = p.REGION_ID, Value = p.REGION_ID };

                foreach (var a in query4.Distinct())
                {
                    list6.Add(new SelectListItem() { Text = a.Text, Value = a.Value });
                }

                ViewBag.Region = list6.OrderBy(x => x.Value);
            }
            //NATION
            List<SelectListItem> list7 = new List<SelectListItem>();
            list7.Add(new SelectListItem() { Text = "ALL", Value = "ALL" });
            using (Entities ctxData = new Entities())
            {
                var query5 = from p in ctxData.WV_REGION_MAST
                             orderby p.NATIONWIDE_GRP
                             select new { Text = p.NATIONWIDE_GRP, Value = p.NATIONWIDE_GRP };

                foreach (var a in query5.Distinct())
                {
                    list7.Add(new SelectListItem() { Text = a.Text, Value = a.Value });
                }

                ViewBag.Nation = list7.OrderBy(x => x.Value);
            }

            List<SelectListItem> list5 = new List<SelectListItem>();
            using (Entities ctxData = new Entities())
            {
                //var query3 = from p in ctxData.WV_PTT_EXC_MAST
                //             where p.EXC_ABB = "0"
                //             orderby p.EXC_ABB
                //             select new { Text = p.EXC_ABB, Value = p.EXC_ABB };

                //foreach (var a in query3.Distinct())
                //{
                //    list5.Add(new SelectListItem() { Text = a.Text, Value = a.Value });
                //}

                ViewBag.exchangeID = list5;
            }
            ViewBag.checkj = "TRUE";

            return View();
        }

        //
        // POST: /Account/Register

        [HttpPost]
        public ActionResult Register(RegisterModel model)
        {
            bool success = true;
            WebView.WebService._base myWebService;
            myWebService = new WebService._base();

            bool selected = false;

            WebService._base.UserMaintenance newUserMaintenance = new WebService._base.UserMaintenance();
            // encrypt the password
            //string password = Crypto.EncryptStringAES(model.PASSWORD);

            newUserMaintenance.USERNAME = model.USERNAME.ToUpper();
            newUserMaintenance.FULL_NAME = model.FULL_NAME;
            newUserMaintenance.PASSWORD = model.PASSWORD;
            //newUserMaintenance.PASSWORD = password;
            newUserMaintenance.GROUPID = model.GROUPID;
            newUserMaintenance.AREA = model.AREA;
            newUserMaintenance.RW_ACCESS = model.RW_ACCESS;
            newUserMaintenance.AUTORIZATION = model.AUTORIZATION;
            newUserMaintenance.NETWORK = model.NETWORK;
            newUserMaintenance.PTT_STATE = model.PTT_STATE;
            newUserMaintenance.HANDOVER = model.HANDOVER;
            newUserMaintenance.USER_ROLES = model.USER_ROLES;
            newUserMaintenance.EXC = model.EXC;
            newUserMaintenance.EMAIL = model.Email;
            newUserMaintenance.NO_TEL = model.NO_TEL;
            newUserMaintenance.REGION = model.REGIONID;
            newUserMaintenance.NATIONS = model.NATION;


            System.Diagnostics.Debug.WriteLine("ID 1: " + model.NATION);
            System.Diagnostics.Debug.WriteLine("ID 2: " + model.REGIONID);
            success = myWebService.AddUser(newUserMaintenance);
            System.Diagnostics.Debug.WriteLine(success);
            selected = true;

            // If we got this far, something failed, redisplay form
            List<SelectListItem> list = new List<SelectListItem>();
            using (Entities ctxData = new Entities())
            {
                var query = from p in ctxData.WV_GROUP
                            orderby p.FULL_NAME
                            select new { Text = p.FULL_NAME, Value = p.GRP_ID };

                foreach (var a in query)
                {
                    list.Add(new SelectListItem() { Text = a.Text, Value = a.Value.ToString() });
                }

                ViewBag.groupID = list;
            }
            List<SelectListItem> list2 = new List<SelectListItem>();
            using (Entities ctxData = new Entities())
            {
                var query1 = from p in ctxData.WV_ROLE
                             orderby p.ROLENAME
                             select new { Text = p.ROLENAME, Value = p.ROLENAME };

                foreach (var a in query1)
                {
                    list2.Add(new SelectListItem() { Text = a.Text, Value = a.Value });
                }

                ViewBag.roleID = list2;
            }
            List<SelectListItem> list3 = new List<SelectListItem>();
            list3.Add(new SelectListItem() { Text = "AUTO", Value = "AUTO" });
            list3.Add(new SelectListItem() { Text = "MANUAL", Value = "MANUAL" });
            list3.Add(new SelectListItem() { Text = "N/A", Value = "N/A" });

            ViewBag.handoverState = list3;
            List<SelectListItem> list4 = new List<SelectListItem>();
            list4.Add(new SelectListItem() { Text = "ALL", Value = "ALL" });
            using (Entities ctxData = new Entities())
            {
                var query3 = from p in ctxData.WV_PTT_EXC_MAST
                             orderby p.PTT_ID ascending
                             select new { Text = p.PTT_ID, Value = p.PTT_ID };

                foreach (var a in query3.Distinct())
                {
                    list4.Add(new SelectListItem() { Text = a.Text, Value = a.Value });
                }

                ViewBag.pttID = list4;
            }
            //if (ModelState.IsValid)
            //{
            // Attempt to register the user
            if (success == true)
            {
                //return RedirectToAction("NewSave?res=" + result);
                return RedirectToAction("NewSave");
            }
            else
            {
                return RedirectToAction("NewSaveFail"); // store to db failed.
            }
            //}
            //return View(model);
        }

        public ActionResult NewSave()
        {
            return View();
        }
        //
        // GET: /Account/ChangePassword

        //[Authorize]
        public ActionResult ChangePassword()
        {
            return View();
        }

        //
        // POST: /Account/ChangePassword

        //[Authorize]
        [HttpPost]
        public ActionResult ChangePassword(ChangePasswordModel model)
        {
            bool success = true;
            string oldPass = "";
            string crrUsr = "";

            WebView.WebService._base myWebService;
            myWebService = new WebService._base();
            using (Entities ctxData = new Entities())
            {
                var query = (from p in ctxData.WV_USER
                             where p.USERNAME == User.Identity.Name
                             select new { p.PASSWORD, p.USERNAME }).Single();

                oldPass = query.PASSWORD;
                crrUsr = query.USERNAME;
            }
            // encrypt the password
            string oldpassword = Crypto.EncryptStringAES(model.OldPassword);
            string password = Crypto.EncryptStringAES(model.NewPassword);
            System.Diagnostics.Debug.WriteLine(password);
            if (oldPass == model.OldPassword)
            {
                WebService._base.UserMaintenance newUser = new WebService._base.UserMaintenance();
                newUser.PASSWORD = model.NewPassword;
                success = myWebService.UpdateUserPass(newUser, crrUsr);
                if (success == true)
                {
                    return RedirectToAction("ChangePasswordSuccess");
                }
                else
                {
                    ModelState.AddModelError("", "The current password is incorrect or the new password is invalid.");
                }
            }
            else
            {
                ModelState.AddModelError("", "The current password is incorrect or the new password is invalid.");
            }
            // If we got this far, something failed, redisplay form
            return View(model);
        }

        //
        // GET: /Account/ChangePasswordSuccess

        public ActionResult ChangePasswordSuccess()
        {
            return View();
        }

        #region Status Codes
        private static string ErrorCodeToString(MembershipCreateStatus createStatus)
        {
            // See http://go.microsoft.com/fwlink/?LinkID=177550 for
            // a full list of status codes.
            switch (createStatus)
            {
                case MembershipCreateStatus.DuplicateUserName:
                    return "User name already exists. Please enter a different user name.";

                case MembershipCreateStatus.DuplicateEmail:
                    return "A user name for that e-mail address already exists. Please enter a different e-mail address.";

                case MembershipCreateStatus.InvalidPassword:
                    return "The password provided is invalid. Please enter a valid password value.";

                case MembershipCreateStatus.InvalidEmail:
                    return "The e-mail address provided is invalid. Please check the value and try again.";

                case MembershipCreateStatus.InvalidAnswer:
                    return "The password retrieval answer provided is invalid. Please check the value and try again.";

                case MembershipCreateStatus.InvalidQuestion:
                    return "The password retrieval question provided is invalid. Please check the value and try again.";

                case MembershipCreateStatus.InvalidUserName:
                    return "The user name provided is invalid. Please check the value and try again.";

                case MembershipCreateStatus.ProviderError:
                    return "The authentication provider returned an error. Please verify your entry and try again. If the problem persists, please contact your system administrator.";

                case MembershipCreateStatus.UserRejected:
                    return "The user creation request has been canceled. Please verify your entry and try again. If the problem persists, please contact your system administrator.";

                default:
                    return "An unknown error occurred. Please verify your entry and try again. If the problem persists, please contact your system administrator.";
            }
        }
        #endregion

        public ActionResult List(string searchKey, string ptt, int? page)
        {
            WebView.WebService._base myWebService;
            myWebService = new WebService._base();

            WebService._base.UserList ExcMain = new WebService._base.UserList();
            if (searchKey != null)
            {
                if (searchKey.Equals(""))
                {
                    if (ptt.Equals(""))
                    {
                        ExcMain = myWebService.GetUserMaintenanceList(0, 50000, null, null);
                        ViewBag.ptt2 = "";
                    }
                    else
                    {
                        ExcMain = myWebService.GetUserMaintenanceList(0, 50000, null, ptt);
                        ViewBag.ptt2 = ptt;
                    }
                }
                else
                {
                    if (ptt.Equals(""))
                    {
                        ExcMain = myWebService.GetUserMaintenanceList(0, 50000, searchKey, null);
                        ViewBag.ptt2 = "";
                        ViewBag.searchKey = searchKey;
                    }
                    else
                    {
                        ExcMain = myWebService.GetUserMaintenanceList(0, 50000, searchKey, ptt);
                        ViewBag.searchKey = searchKey;
                        ViewBag.ptt2 = ptt;
                    }
                }
            }
            else
            {
                if (ptt == null)
                {
                    ExcMain = myWebService.GetUserMaintenanceList(0, 50000, null, null);
                    ViewBag.searchKey = "";
                    ViewBag.ptt2 = "";
                }
                else
                {
                    ExcMain = myWebService.GetUserMaintenanceList(0, 50000, null, ptt);
                    ViewBag.searchKey = "";
                    ViewBag.ptt2 = ptt;
                }
            }

            ViewData["data10"] = ExcMain.UsrList;

            string input = "\\\\adsvr";
            //string input = "\\\\server\\d$\\x\\y\\z\\AAA";
            string output = String.Format("http:{0}", input.Replace("\\d$\\x\\y", String.Empty).Replace("\\", "/"));

            ViewBag.output = output;

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
                ViewBag.ptt = list;
            }

            //return View();
            int pageSize = 15;
            int pageNumber = (page ?? 1);
            return View(ExcMain.UsrList.ToPagedList(pageNumber, pageSize));
        }

        [HttpPost]
        public ActionResult ChangePass(string UserName, string Old_Pass, string New_Pass)
        {
            bool success = true;
            // System.Diagnostics.Debug.WriteLine("UPDATE!! :" + UserName + "Old Pass: " + Old_Pass + "new Pass : " + New_Pass);
            WebView.WebService._base myWebService;
            myWebService = new WebService._base();
            //string password = Crypto.EncryptStringAES(txtPASSWORD);

            WebService._base.UserMaintenance newUser = new WebService._base.UserMaintenance();

            newUser.PASSWORD = New_Pass;

            success = myWebService.UpdateNewPass(newUser, UserName);
            System.Diagnostics.Debug.WriteLine("P :" + success);
            return Json(new
            {
                Success = success
            }, JsonRequestBehavior.AllowGet);
        }

        //[HttpPost]
        //public ActionResult ForgotPasswordChange(string User)
        //{
        //    bool success = true;
        //    string Emails = "";
        //    System.Diagnostics.Debug.WriteLine("PPPPPPPPPPPPPPP");
        //    using (Entities ctxData = new Entities())
        //    {
        //        var query = (from p in ctxData.WV_USER
        //                     where p.USERNAME == User
        //                     select new { Value = p.EMAIL });

        //        foreach (var b in query.Distinct())
        //        {
        //            if (b != null)
        //            {
        //                Emails = Emails + b.Value;
        //            }
        //        }

        //        return Json(new
        //        {
        //            Email = Emails
        //        }, JsonRequestBehavior.AllowGet); //
        //    }
        //}     

        [HttpPost]
        public ActionResult GetDetails(string id,string group)
        {
            string fileListStr = "";
            string Exctype = "";
            string PuList = "";
            string PuListPTT = "";
            string PuListGrp = "";
            string GrpID = "";
            string Emails = "";

            using (Entities ctxData = new Entities())
            {
//<<<<<<< .mine
                var queryUser = (from p in ctxData.WV_USER
                                 where p.USERNAME == id
                                 select p).Single();
                
//=======
                //var queryUser = (from p in ctxData.WV_USER
                //                 where p.USERNAME == id
                //                 select p).Single();
                ////select new { Value = p.EMAIL });
//>>>>>>> .r1774

//<<<<<<< .mine
                Emails = queryUser.EMAIL;
//=======
                ////foreach (var b in queryUser.Distinct())
                ////{
                ////    if (b != null)
                ////    {
                ////        Emails = b.Value;
                ////        System.Diagnostics.Debug.WriteLine("Pizal" + Emails);
                ////    }
                ////}
                //Emails = queryUser.EMAIL;
//>>>>>>> .r1774

//<<<<<<< .mine
                //System.Diagnostics.Debug.WriteLine("Email " + Emails);
                //if (Emails == null)
                //{
                //    System.Diagnostics.Debug.WriteLine("Details List " + Emails);
//=======
                System.Diagnostics.Debug.WriteLine("Email " + Emails);
                //if (Emails == null)
                //{
                    System.Diagnostics.Debug.WriteLine("Details List " + Emails);
//>>>>>>> .r1774
                    var query = (from q in ctxData.WV_USER
                                 where q.USERNAME == id
                                 select q).Single();
                    GrpID = query.GROUPID;
                    Emails = query.EMAIL;
                    // System.Diagnostics.Debug.WriteLine("GRODUP : " + GrpID);
                    var query1 = (from q in ctxData.WV_GROUP
                                  where q.GRP_ID == GrpID.Trim()
                                  select q).Single();

                    //filter PTT
                    List<SelectListItem> listPTT = new List<SelectListItem>();

                    var queryPTT = from p in ctxData.WV_EXC_MAST
                                   select new { p.PTT_ID };
                    PuListPTT = PuListPTT + "ALL:  ALL|";
                    foreach (var a in queryPTT.Distinct().OrderBy(it => it.PTT_ID))
                    {
                        PuListPTT = PuListPTT + a.PTT_ID + ":  " + a.PTT_ID + "|";
                    }
                    ViewBag.pttID = listPTT;

                    //filter Group
                    List<SelectListItem> listGrp = new List<SelectListItem>();
                    var queryGrp = from p in ctxData.WV_GROUP
                                   select new { p.FULL_NAME, p.GRP_ID };

                    foreach (var a in queryGrp.Distinct().OrderBy(it => it.FULL_NAME))
                    {
                        PuListGrp = PuListGrp + a.GRP_ID + ":  " + a.FULL_NAME + "|";
                    }
                    ViewBag.grpID = PuListGrp;

                    //filter EXC
                    List<SelectListItem> list = new List<SelectListItem>();
                    var queryEXC = from p in ctxData.WV_EXC_MAST
                                   where p.PTT_ID.Trim() == query.PTT_STATE.Trim()
                                   select new { p.EXC_ABB };

                    foreach (var a in queryEXC.Distinct().OrderBy(it => it.EXC_ABB))
                    {
                        PuList = PuList + a.EXC_ABB + ":  " + a.EXC_ABB + "|";
                    }
                    ViewBag.exchangeID = list;

                    //string password = Crypto.DecryptStringAES(query.PASSWORD);
                    return Json(new
                    {
                        Success = true,
                        USERNAME = id,
                        FULL_NAME = query.FULL_NAME,
                        //PASSWORD = password,
                        PASSWORD = query.PASSWORD,
                        GROUPID = query.GROUPID,
                        AREA = query.AREA,
                        RW_ACCESS = query.RW_ACCESS,
                        AUTORIZATION = query.AUTORIZATION,
                        NETWORK = query.NETWORK,
                        PTT_STATE = query.PTT_STATE,
                        HANDOVER = query.HANDOVER,
                        USER_ROLES = query.USER_ROLES,
                        EXC = query.EXC,
                        EMAIL = query.EMAIL,
                        NO_TEL = query.NO_TEL,
                        PuList = PuList,
                        PuListPTT = PuListPTT,
                        PuListGrp = PuListGrp,
                        GRPNAME = query1.FULL_NAME,
                        Email = Emails,
                        GRPOLD = group,
                    }, JsonRequestBehavior.AllowGet); //
                //}
                //else
                //{
                //    return Json(new
                //    {
                //        Email = Emails,
                //    }, JsonRequestBehavior.AllowGet); //
                //}
            }
        }

        public ActionResult UpdateData(string txtUSERNAME, string txtFULL_NAME, string txtPASSWORD, string txtGROUPID, string txtAREA, string txtRW_ACCESS, string txtAUTORIZATION, string txtPTT_STATE, string txtHANDOVER, string txtUSER_ROLES, string txtEXC, string txtEMAIL, string txtNO_TEL, string targetGroup)
        {
            bool success = true;


            WebView.WebService._base myWebService;
            myWebService = new WebService._base();
            //string password = Crypto.EncryptStringAES(txtPASSWORD);
            System.Diagnostics.Debug.WriteLine("Not Update!" + txtUSERNAME);
            WebService._base.UserMaintenance newUser = new WebService._base.UserMaintenance();
            System.Diagnostics.Debug.WriteLine("UPDATE!! :" + txtUSERNAME);
            newUser.FULL_NAME = txtFULL_NAME;
            //newUser.PASSWORD = password;
            newUser.PASSWORD = txtPASSWORD;
            newUser.GROUPID = txtGROUPID;
            newUser.AREA = txtAREA;
            newUser.RW_ACCESS = txtRW_ACCESS;
            newUser.AUTORIZATION = txtAUTORIZATION;
            newUser.PTT_STATE = txtPTT_STATE;
            newUser.HANDOVER = txtHANDOVER;
            newUser.USER_ROLES = txtUSER_ROLES;
            newUser.EXC = txtEXC;
            newUser.EMAIL = txtEMAIL;
            newUser.NO_TEL = txtNO_TEL;

            success = myWebService.UpdateUser(newUser, txtUSERNAME,targetGroup);

            return Json(new
            {
                Success = success
            }, JsonRequestBehavior.AllowGet);
        }

        [HttpPost]
        public ActionResult DeleteData(string targetCanCode)
        {
            bool success;

            WebView.WebService._base myWebService;
            myWebService = new WebService._base();

            success = myWebService.DeleteUserData(targetCanCode);
            System.Diagnostics.Debug.WriteLine("USERNAME!! :" + targetCanCode);
            return Json(new
            {
                Success = success
            }, JsonRequestBehavior.AllowGet);
        }

        [HttpPost]
        public ActionResult forgotPass(string id)
        {
            string Emails = "";
            System.Diagnostics.Debug.WriteLine("FUCK" + id);
            using (Entities ctxData = new Entities())
            {
                var query = (from q in ctxData.WV_USER
                             where q.USERNAME == id
                             select q).Single();

                //foreach(var a in query.Distinct())
                //{
                //    Emails = Emails +  a.value;
                //}
                System.Diagnostics.Debug.WriteLine("FUCK 2" + query.EMAIL);
                return Json(new
                {
                    Success = true,
                    Emails = query.EMAIL
                    
                }, JsonRequestBehavior.AllowGet); //
            }
        }

        [HttpPost]
        public ActionResult EmailTest() // NIS Network Element
        {
            System.Diagnostics.Debug.WriteLine("Email test : 1 ok");
            string result;
            try
            {
                MailMessage msg = new MailMessage();
                msg.IsBodyHtml = true;
                msg.From = new MailAddress("neps@tm.com.my", "NEPS");
                msg.To.Add("pizalpigi@gmail.com.my");
                msg.Subject = "EMAIL FROM LOCAL ";
                msg.Body = "<h1>FILES DETAILS</h1>SCHEME NAME	: <br/><br/>DESCRIPTION : <br/><br/>REDMARK FILE NAME: .xml ";
                msg.Body += "<br/><br/> <h1>RNO DETAILS</h1> <br>";
                msg.Body += "RNO ID : " + User.Identity.Name + "<br/><br/>RNO EMAIL	: <br/><br/>RNO PHONE NUMBER: <br/><br/>Please log in to <a href='http://10.41.101.168/'>NEPS WEBVIEW  </a>to download the file.";
                msg.IsBodyHtml = true;
                SmtpClient emailClient = new SmtpClient("smtp.tm.com.my");
                emailClient.UseDefaultCredentials = true;
                emailClient.Port = 25;
                emailClient.EnableSsl = false;
                //emailClient.UseDefaultCredentials = false;
                emailClient.Send(msg);
                result = "OK";
                System.Diagnostics.Debug.WriteLine("Email test : 2 ok");
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

        [HttpPost]
        public ActionResult updataListData(string PTT_ID)
        {
            string PuList = "";
            string PuList2 = "";
            using (Entities ctxData = new Entities())
            {
                //filter PTT
                List<SelectListItem> list = new List<SelectListItem>();
                var queryEXC = from p in ctxData.WV_EXC_MAST
                               where p.PTT_ID.Trim() == PTT_ID.Trim()
                               select new { p.EXC_ABB };

                foreach (var a in queryEXC.Distinct().OrderBy(it => it.EXC_ABB))
                {
                    PuList = PuList + a.EXC_ABB + ":  " + a.EXC_ABB + "|";
                }
                ViewBag.exchangeID = list;
            }

            return Json(new
            {
                Success = true,
                PuList = PuList
            }, JsonRequestBehavior.AllowGet); //
        }

    }
}

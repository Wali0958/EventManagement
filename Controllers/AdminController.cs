using ClosedXML.Excel;
using DocumentFormat.OpenXml.Office2010.Excel;
using EventManagement.Models;
using QRCoder;
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing.Imaging;
using System.Drawing;
using System.EnterpriseServices.Internal;
using System.IO;
using System.Linq;
using System.Text;
using System.Web;
using System.Web.Mvc;
using static Microsoft.IO.RecyclableMemoryStreamManager;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Text.RegularExpressions;

namespace EventManagement.Controllers
{

    public class AdminController : Controller
    {
        Master_Helper masterHelper = new Master_Helper();
        DataSet ds = new DataSet();
        #region code master dashboard
        public ActionResult AdminDashboard()
        {
            Dashboard dash = new Dashboard();
            ds = masterHelper.FetchDashboard("DashboardCount", dash, Convert.ToInt16(Session["LoginId"]), Convert.ToInt32(Session["EventId"]));
            if (ds != null && ds.Tables[0].Rows.Count > 0)
            {
                var row = ds.Tables[0].Rows[0];
                dash.TotalEvent = Convert.ToInt32(row["TotalEvent"]);
                dash.TotalUser = Convert.ToInt32(row["TotalUser"]);
                dash.TotalBadge = Convert.ToInt32(row["TotalBadge"]);
                dash.TotalBadgeNo = Convert.ToInt32(row["TotalBadgeNo"]);
            }
            ds = masterHelper.FetchDashboard("ViewAllbadgeEvent", dash, Convert.ToInt16(Session["LoginId"]), Convert.ToInt32(Session["EventId"]));
            if (ds != null && ds.Tables[0].Rows.Count > 0)
            {
                dash.EventExcelUpload = CommonFunctions.ToList<mExcel>(ds.Tables[0]);
            }
            return View(dash);
        }
        #endregion
        #region Event Code

        string Generate4Char()
        {
            const string chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
            Random rnd = new Random();
            StringBuilder sb = new StringBuilder(6);

            for (int i = 0; i < 6; i++)
                sb.Append(chars[rnd.Next(chars.Length)]);

            return sb.ToString();
        }
        // GET: Admin
        public ActionResult MasterEvent()
        {
            EventModels modal = new EventModels();
            ds = masterHelper.FetchEvent("Get Date", modal, Convert.ToInt16(Session["LoginId"]));
            List<EventModelsList> list = new List<EventModelsList>();
            dynamic mevent = CommonFunctions.ToList<EventModelsList>(ds.Tables[0]);
            if (ds != null && ds.Tables[0].Rows.Count > 0)
            {
                foreach (var item in mevent)
                {
                    EventModelsList ev = new EventModelsList();
                    ev.EventId = item.EventId;
                    ev.EventName = item.EventName;
                    ev.EventStartOn = item.EventStartOn;
                    ev.EventEndOn = item.EventEndOn;
                    ev.EventLocation = item.EventLocation;
                    ev.EventHeadName = item.EventHeadName;
                    ev.ContactNo = item.ContactNo;
                    ev.EventDay = item.EventDay;
                    ev.EventBadgePhoto = item.EventBadgePhoto;
                    ev.BadgeHeight = item.BadgeHeight;
                    ev.BadgeWidth = item.BadgeWidth;
                    ev.IsActive = item.IsActive;
                    list.Add(ev);
                }
            }
            else
            {

            }
            return View(mevent);
        }

        [HttpPost]
        public ActionResult MasterEvent(EventModels modal, HttpPostedFileBase EventBadgePhoto,HttpPostedFileBase EventCertificatePhoto)
        {
            if (modal.EventName.Trim() != "" && modal.EventLocation.Trim() != null && modal.EventEndOn.ToString() != "01-01-2025" && modal.EventStartOn.ToString() != "01-01-2025"
                && modal.ContactNo.Trim() != "" && modal.EventHeadName.Trim() != "")
            {

                if (EventBadgePhoto != null && EventBadgePhoto.ContentLength > 0)
                {
                    try
                    {
                       
                        if (modal.PrintType != null && modal.PrintType.Count > 0)
                        {
                            modal.PrintTypeValue = string.Join(",", modal.PrintType);
                        }
                        else
                        {
                            TempData["Msg"] = "Please select at least one Print Type.";
                            return RedirectToAction("MasterEvent");
                        }
                        // Optional: Validate file type
                        string[] allowedExtensions = { ".jpg", ".png", ".jpeg" };
                        string fileExtension = Path.GetExtension(EventBadgePhoto.FileName.Trim()).ToLower();

                        if (!allowedExtensions.Contains(fileExtension))
                        {
                            TempData["Message"] = "File type not allowed.";
                            return RedirectToAction("MasterEvent"); // Change as needed
                        }
                        string folderPath = Server.MapPath("~/Content/EventBadge/"); // Make sure this folder exists
                        if (!Directory.Exists(folderPath))
                            Directory.CreateDirectory(folderPath);
                        // Set the file name and save path
                        string fileName = Path.GetFileName(DateTime.Now + "_" + Regex.Replace(modal.EventName, @"\s+", "") + "_" + Generate4Char()+ fileExtension);
                        string fullPath = Path.Combine(folderPath.Trim(), fileName.Trim());
                        // Save the file
                        EventBadgePhoto.SaveAs(fullPath.Trim());
                        modal.EventBadgePhoto = fileName.Trim();
                    }
                    catch (Exception ex)
                    {
                        TempData["Message"] = "ERROR: " + ex.Message.ToString();
                        return RedirectToAction("MasterEvent", "Admin");
                    }
                }
                modal.BadgeHeight = (modal.BadgeHeight * 96);
                modal.BadgeWidth = (modal.BadgeWidth * 96);
                if (EventCertificatePhoto != null && EventCertificatePhoto.ContentLength > 0)
                {
                    try
                    {

                        if (modal.PrintType != null && modal.PrintType.Count > 0)
                        {
                            modal.PrintTypeValue = string.Join(",", modal.PrintType);
                        }
                        else
                        {
                            TempData["Msg"] = "Please select at least one Print Type.";
                            return RedirectToAction("MasterEvent");
                        }
                        // Optional: Validate file type
                        string[] allowedExtensions = { ".jpg", ".png", ".jpeg" };
                        string fileExtension = Path.GetExtension(EventCertificatePhoto.FileName.Trim()).ToLower();

                        if (!allowedExtensions.Contains(fileExtension))
                        {
                            TempData["Message"] = "File type not allowed.";
                            return RedirectToAction("MasterEvent"); 
                        }
                        string folderPath = Server.MapPath("~/Content/EventBadge/"); 
                        if (!Directory.Exists(folderPath))
                            Directory.CreateDirectory(folderPath);
                        // Set the file name and save path
                        string fileName = Path.GetFileName(DateTime.Now + "_" + Regex.Replace(modal.EventName, @"\s+", "") + "_" + Generate4Char() + fileExtension);
                        string fullPath = Path.Combine(folderPath.Trim(), fileName.Trim());
                        EventCertificatePhoto.SaveAs(fullPath.Trim());
                        modal.EventCertificatePhoto = fileName.Trim();
                    }
                    catch (Exception ex)
                    {
                        TempData["Message"] = "ERROR: " + ex.Message.ToString();
                        return RedirectToAction("MasterEvent", "Admin");
                    }
                }
                modal.CertificateHeight = (modal.CertificateHeight * 96);
                modal.CertificateWidth = (modal.CertificateWidth * 96);
                int printTypeMask = modal.PrintType.Sum();
                modal.PrintTypeValue = printTypeMask.ToString();
                ds = masterHelper.FetchEvent("SaveEvent", modal, Convert.ToInt16(Session["LoginId"]));
                if (ds != null && ds.Tables[0].Rows.Count > 0)
                {
                    if (ds.Tables[0].Rows[0]["DB_STATUS"].ToString() == "S")
                    {
                        Session["EventId"] = ds.Tables[1].Rows[0]["EventID"].ToString();
                        TempData["Msg"] = ds.Tables[0].Rows[0]["DB_StatusMessage"].ToString();
                    }
                    else
                    {
                        TempData["Msg"] = ds.Tables[0].Rows[0]["DB_StatusMessage"].ToString();
                    }
                }
                else
                {
                    TempData["Msg"] = "Oops some error occured";
                }
            }
            else
            {
                TempData["Msg"] = "Please fill all detsils.";
            }
            return MasterEvent();
        }

        [HttpGet]
        public ActionResult DeleteEvent(int id)
        {
            if (id != 0)
            {
                EventModels modal = new EventModels();
                modal.EventId = id;
                ds = masterHelper.FetchEvent("DeleteEvent", modal, Convert.ToInt16(Session["LoginId"]));
                if (ds != null && ds.Tables[0].Rows.Count > 0)
                {
                    if (ds.Tables[0].Rows[0]["DB_STATUS"].ToString() == "S")
                    {
                        TempData["Msg"] = ds.Tables[0].Rows[0]["DB_StatusMessage"].ToString();
                    }
                    else
                    {
                        TempData["Msg"] = ds.Tables[0].Rows[0]["DB_StatusMessage"].ToString();
                    }
                }
                else
                {
                    TempData["Msg"] = "Oops some error occured";
                }
            }
            else
            {
                TempData["Msg"] = "Oops some error occured";
            }

            return RedirectToAction("MasterEvent", "Admin");
        }
        [HttpGet]
        public ActionResult DeactiveEvent(int id)
        {
            if (id != 0)
            {
                EventModels modal = new EventModels();
                modal.EventId = id;
                ds = masterHelper.FetchEvent("DeActiveEvent", modal, Convert.ToInt16(Session["LoginId"]));
                if (ds != null && ds.Tables[0].Rows.Count > 0)
                {
                    if (ds.Tables[0].Rows[0]["DB_STATUS"].ToString() == "S")
                    {
                        TempData["Msg"] = ds.Tables[0].Rows[0]["DB_StatusMessage"].ToString();
                    }
                    else
                    {
                        TempData["Msg"] = ds.Tables[0].Rows[0]["DB_StatusMessage"].ToString();
                    }
                }
                else
                {
                    TempData["Msg"] = "Oops some error occured";
                }
            }
            else
            { TempData["Msg"] = "Oops some error occured"; }
            return RedirectToAction("MasterEvent", "Admin");
        }
        [HttpGet]
        public ActionResult ActiveEvent(int id)
        {

            if (id != 0)
            {
                EventModels modal = new EventModels();
                modal.EventId = id;
                ds = masterHelper.FetchEvent("ActiveEvent", modal, Convert.ToInt16(Session["LoginId"]));
                if (ds != null && ds.Tables[0].Rows.Count > 0)
                {
                    if (ds.Tables[0].Rows[0]["DB_STATUS"].ToString() == "S")
                    {
                        TempData["Msg"] = ds.Tables[0].Rows[0]["DB_StatusMessage"].ToString();
                    }
                    else
                    {
                        TempData["Msg"] = ds.Tables[0].Rows[0]["DB_StatusMessage"].ToString();
                    }
                }
                else
                {
                    TempData["Msg"] = "Oops some error occured";
                }
            }
            else
            {
                TempData["Msg"] = "Oops some error occured";
            }
            return RedirectToAction("MasterEvent", "Admin");
        }
        #endregion
        #region master category
        public ActionResult MasterEventCategory()
        {
            EventCat EC = new EventCat();
            ds = masterHelper.FetchEventCategory("BindEevCatorCat", EC, Convert.ToInt16(Session["LoginId"]));
            if (ds != null && ds.Tables[0].Rows.Count > 0)
            {
                EC.EventDrop = CommonFunctions.ToList<EventDropdownList>(ds.Tables[0]);
            }
            if (ds != null && ds.Tables[1].Rows.Count > 0)
            {
                EC.EventCatAll = CommonFunctions.ToList<EventCatNew>(ds.Tables[1]);
            }
            return View(EC);
        }
        [HttpPost]
        public ActionResult MasterEventCategory(EventCat modal)
        {
            if (modal.CategoryName != null)
            {
                ds = masterHelper.FetchEventCategory("SaveCateogry", modal, Convert.ToInt16(Session["LoginId"]));
                if (ds != null && ds.Tables[0].Rows.Count > 0)
                {
                    if (ds.Tables[0].Rows[0]["DB_STATUS"].ToString() == "S")
                    {
                        TempData["Msg"] = ds.Tables[0].Rows[0]["DB_StatusMessage"].ToString();
                    }
                    else
                    {
                        TempData["Msg"] = ds.Tables[0].Rows[0]["DB_StatusMessage"].ToString();
                    }
                }
                else
                {
                    TempData["Msg"] = "Oops some error occured";
                }
            }
            else
            {
                TempData["Msg"] = "Please fill all detsils.";
            }
            return MasterEventCategory();
        }

        [HttpGet]
        public ActionResult DeleteEventCategory(int id)
        {
            if (id != 0)
            {
                EventCat modal = new EventCat();
                modal.CategoryId = id;
                ds = masterHelper.FetchEventCategory("DelCat", modal, Convert.ToInt16(Session["LoginId"]));
                if (ds != null && ds.Tables[0].Rows.Count > 0)
                {
                    if (ds.Tables[0].Rows[0]["DB_STATUS"].ToString() == "S")
                    {
                        TempData["Msg"] = ds.Tables[0].Rows[0]["DB_StatusMessage"].ToString();
                    }
                    else
                    {
                        TempData["Msg"] = ds.Tables[0].Rows[0]["DB_StatusMessage"].ToString();
                    }
                }
                else
                {
                    TempData["Msg"] = "Oops some error occured";
                }
            }
            else
            {
                TempData["Msg"] = "Oops some error occured";
            }
            return RedirectToAction("MasterEventCategory", "Admin");
        }
        [HttpGet]
        public ActionResult DeactiveEventCategory(int id)
        {
            if (id != 0)
            {
                EventCat modal = new EventCat();
                modal.CategoryId = id;
                ds = masterHelper.FetchEventCategory("updateDeaCat", modal, Convert.ToInt16(Session["LoginId"]));
                if (ds != null && ds.Tables[0].Rows.Count > 0)
                {
                    if (ds.Tables[0].Rows[0]["DB_STATUS"].ToString() == "S")
                    {
                        TempData["Msg"] = ds.Tables[0].Rows[0]["DB_StatusMessage"].ToString();
                    }
                    else
                    {
                        TempData["Msg"] = ds.Tables[0].Rows[0]["DB_StatusMessage"].ToString();
                    }
                }
                else
                {
                    TempData["Msg"] = "Oops some error occured";
                }
            }
            else
            { TempData["Msg"] = "Oops some error occured"; }

            return RedirectToAction("MasterEventCategory", "Admin");
        }
        [HttpGet]
        public ActionResult ActiveEventCategory(int id)
        {

            if (id != 0)
            {
                EventCat modal = new EventCat();
                modal.CategoryId = id;
                ds = masterHelper.FetchEventCategory("updateActiveCat", modal, Convert.ToInt16(Session["LoginId"]));
                if (ds != null && ds.Tables[0].Rows.Count > 0)
                {
                    if (ds.Tables[0].Rows[0]["DB_STATUS"].ToString() == "S")
                    {
                        TempData["Msg"] = ds.Tables[0].Rows[0]["DB_StatusMessage"].ToString();
                    }
                    else
                    {
                        TempData["Msg"] = ds.Tables[0].Rows[0]["DB_StatusMessage"].ToString();
                    }
                }
                else
                {
                    TempData["Msg"] = "Oops some error occured";
                }
            }
            else
            {
                TempData["Msg"] = "Oops some error occured";
            }
            return RedirectToAction("MasterEventCategory", "Admin");
        }

        #endregion    
        #region Add master usercode
        [HttpGet]
        public ActionResult AddMasterUser()
        {
            MasterUser EC = new MasterUser();
            ds = masterHelper.FetchUser("Get User", EC, Convert.ToInt16(Session["LoginId"]));
            if (ds != null && ds.Tables[0].Rows.Count > 0)
            {
                EC.EventDrop = CommonFunctions.ToList<EventDropdownList>(ds.Tables[0]);
            }
            if (ds != null && ds.Tables[0].Rows.Count > 0)
            {
                EC.URole = CommonFunctions.ToList<UserRoleModel>(ds.Tables[1]);
            }
            if (ds != null && ds.Tables[1].Rows.Count > 0)
            {
                EC.UserDisplay = CommonFunctions.ToList<masteruserall>(ds.Tables[2]);
            }
            return View(EC);
        }
        [HttpPost]
        public ActionResult AddMasterUser(MasterUser modal)
        {
            if (modal.UserEmail != null && modal.UserRole != 0 && modal.Password != "" && modal.EventId != 0)
            {
                ds = masterHelper.FetchUser("SaveUser", modal, Convert.ToInt16(Session["LoginId"]));
                if (ds != null && ds.Tables[0].Rows.Count > 0)
                {
                    if (ds.Tables[0].Rows[0]["DB_STATUS"].ToString() == "S")
                    {
                        TempData["Msg"] = ds.Tables[0].Rows[0]["DB_StatusMessage"].ToString();
                    }
                    else
                    {
                        TempData["Msg"] = ds.Tables[0].Rows[0]["DB_StatusMessage"].ToString();
                    }
                }
                else
                {
                    TempData["Msg"] = "Oops some error occured";
                }
            }
            else
            {
                TempData["Msg"] = "Please fill all detsils.";
            }
            return AddMasterUser();
        }

        [HttpGet]
        public ActionResult DeleteUser(int id)
        {
            if (id != 0)
            {
                MasterUser modal = new MasterUser();
                modal.LoginId = id;
                ds = masterHelper.FetchUser("DelUser", modal, Convert.ToInt16(Session["LoginId"]));
                if (ds != null && ds.Tables[0].Rows.Count > 0)
                {
                    if (ds.Tables[0].Rows[0]["DB_STATUS"].ToString() == "S")
                    {
                        TempData["Msg"] = ds.Tables[0].Rows[0]["DB_StatusMessage"].ToString();
                    }
                    else
                    {
                        TempData["Msg"] = ds.Tables[0].Rows[0]["DB_StatusMessage"].ToString();
                    }
                }
                else
                {
                    TempData["Msg"] = "Oops some error occured";
                }
            }
            else
            {
                TempData["Msg"] = "Oops some error occured";
            }
            return RedirectToAction("AddMasterUser", "Admin");
        }
        [HttpGet]
        public ActionResult DeactiveUser(int id)
        {
            if (id != 0)
            {
                MasterUser modal = new MasterUser();
                modal.LoginId = id;
                ds = masterHelper.FetchUser("updateDeaUser", modal, Convert.ToInt16(Session["LoginId"]));
                if (ds != null && ds.Tables[0].Rows.Count > 0)
                {
                    if (ds.Tables[0].Rows[0]["DB_STATUS"].ToString() == "S")
                    {
                        TempData["Msg"] = ds.Tables[0].Rows[0]["DB_StatusMessage"].ToString();
                    }
                    else
                    {
                        TempData["Msg"] = ds.Tables[0].Rows[0]["DB_StatusMessage"].ToString();
                    }
                }
                else
                {
                    TempData["Msg"] = "Oops some error occured";
                }
            }
            else
            { TempData["Msg"] = "Oops some error occured"; }


            return RedirectToAction("AddMasterUser", "Admin");
        }
        [HttpGet]
        public ActionResult ActiveUser(int id)
        {

            if (id != 0)
            {
                MasterUser modal = new MasterUser();
                modal.LoginId = id;
                ds = masterHelper.FetchUser("updateActiUser", modal, Convert.ToInt16(Session["LoginId"]));
                if (ds != null && ds.Tables[0].Rows.Count > 0)
                {
                    if (ds.Tables[0].Rows[0]["DB_STATUS"].ToString() == "S")
                    {
                        TempData["Msg"] = ds.Tables[0].Rows[0]["DB_StatusMessage"].ToString();
                    }
                    else
                    {
                        TempData["Msg"] = ds.Tables[0].Rows[0]["DB_StatusMessage"].ToString();
                    }
                }
                else
                {
                    TempData["Msg"] = "Oops some error occured";
                }
            }
            else
            {
                TempData["Msg"] = "Oops some error occured";
            }
            return RedirectToAction("AddMasterUser", "Admin");
        }
        #endregion
        #region column set master
        public ActionResult MasterShowColEvent()
        {
            EventSettingColumns EC = new EventSettingColumns();
            ds = masterHelper.FetchEventColumns("Get Date columns", EC, Convert.ToInt16(Session["LoginId"]));
            if (ds != null && ds.Tables[0].Rows.Count > 0)
            {
                EC.EventDrop = CommonFunctions.ToList<EventDropdownList>(ds.Tables[0]);
            }
            if (ds != null && ds.Tables[1].Rows.Count > 0)
            {
                EC.EventcolumnsDrop = CommonFunctions.ToList<EventSettingColumns>(ds.Tables[1]);
            }
            return View(EC);
        }

        [HttpPost]
        public ActionResult MasterShowColEvent(EventSettingColumns modal)
        {
            if (modal.EventId != 0 && modal.RName != false)
            {
                ds = masterHelper.FetchEventColumns("SaveEventColumns", modal, Convert.ToInt16(Session["LoginId"]));
                if (ds != null && ds.Tables[0].Rows.Count > 0)
                {
                    if (ds.Tables[0].Rows[0]["DB_STATUS"].ToString() == "S")
                    {
                        TempData["Msg"] = ds.Tables[0].Rows[0]["DB_StatusMessage"].ToString();
                    }
                    else
                    {
                        TempData["Msg"] = ds.Tables[0].Rows[0]["DB_StatusMessage"].ToString();
                    }
                }
                else
                {
                    TempData["Msg"] = "Oops some error occured";
                }
            }
            else
            {
                TempData["Msg"] = "Please fill all details.";
            }
            return RedirectToAction("MasterShowColEvent", "Admin");
        }

        [HttpGet]
        public ActionResult DeleteEventcolumns(int id, int eventid)
        {
            if (id != 0)
            {
                EventSettingColumns modal = new EventSettingColumns();
                modal.ColumnID = id;
                modal.EventId = eventid;
                ds = masterHelper.FetchEventColumns("DeleteEventcolumns", modal, Convert.ToInt16(Session["LoginId"]));
                if (ds != null && ds.Tables[0].Rows.Count > 0)
                {
                    if (ds.Tables[0].Rows[0]["DB_STATUS"].ToString() == "S")
                    {
                        TempData["Msg"] = ds.Tables[0].Rows[0]["DB_StatusMessage"].ToString();
                    }
                    else
                    {
                        TempData["Msg"] = ds.Tables[0].Rows[0]["DB_StatusMessage"].ToString();
                    }
                }
                else
                {
                    TempData["Msg"] = "Oops some error occured";
                }
            }
            else
            {
                TempData["Msg"] = "Oops some error occured";
            }
            return RedirectToAction("MasterShowColEvent", "Admin");
        }
        [HttpGet]
        public ActionResult DeactiveEventcolumns(int id)
        {
            if (id != 0)
            {
                EventSettingColumns modal = new EventSettingColumns();
                modal.ColumnID = id;
                ds = masterHelper.FetchEventColumns("DeActiveEventcolumns", modal, Convert.ToInt16(Session["LoginId"]));
                if (ds != null && ds.Tables[0].Rows.Count > 0)
                {
                    if (ds.Tables[0].Rows[0]["DB_STATUS"].ToString() == "S")
                    {
                        TempData["Msg"] = ds.Tables[0].Rows[0]["DB_StatusMessage"].ToString();
                    }
                    else
                    {
                        TempData["Msg"] = ds.Tables[0].Rows[0]["DB_StatusMessage"].ToString();
                    }
                }
                else
                {
                    TempData["Msg"] = "Oops some error occured";
                }
            }
            else
            { TempData["Msg"] = "Oops some error occured"; }


            return RedirectToAction("MasterShowColEvent", "Admin");
        }
        [HttpGet]
        public ActionResult ActiveEventcolumns(int id)
        {

            if (id != 0)
            {
                EventSettingColumns modal = new EventSettingColumns();
                modal.ColumnID = id;
                ds = masterHelper.FetchEventColumns("ActiveEventcolumns", modal, Convert.ToInt16(Session["LoginId"]));
                if (ds != null && ds.Tables[0].Rows.Count > 0)
                {
                    if (ds.Tables[0].Rows[0]["DB_STATUS"].ToString() == "S")
                    {
                        TempData["Msg"] = ds.Tables[0].Rows[0]["DB_StatusMessage"].ToString();
                    }
                    else
                    {
                        TempData["Msg"] = ds.Tables[0].Rows[0]["DB_StatusMessage"].ToString();
                    }
                }
                else
                {
                    TempData["Msg"] = "Oops some error occured";
                }
            }
            else
            {
                TempData["Msg"] = "Oops some error occured";
            }
            return RedirectToAction("MasterShowColEvent", "Admin");
        }


        #endregion columns end
        #region column set master
        public ActionResult MasterShowColEventCertificate()
        {
            EventSettingColumns EC = new EventSettingColumns();
            EC.mType = "View";
            ds = masterHelper.FetchEventColumns("Get Date columns", EC, Convert.ToInt16(Session["LoginId"]));
            if (ds != null && ds.Tables[0].Rows.Count > 0)
            {
                EC.EventDrop = CommonFunctions.ToList<EventDropdownList>(ds.Tables[0]);
            }
            if (ds != null && ds.Tables[1].Rows.Count > 0)
            {
                EC.EventcolumnsDrop = CommonFunctions.ToList<EventSettingColumns>(ds.Tables[1]);
            }
            return View(EC);
        }

        [HttpPost]
        public ActionResult MasterShowColEventCertificate(EventSettingColumns modal)
        {
            if (modal.EventId != 0 && modal.RName != false)
            {
                ds = masterHelper.FetchEventColumns("SaveEventColumnsCertificate", modal, Convert.ToInt16(Session["LoginId"]));
                if (ds != null && ds.Tables[0].Rows.Count > 0)
                {
                    if (ds.Tables[0].Rows[0]["DB_STATUS"].ToString() == "S")
                    {
                        TempData["Msg"] = ds.Tables[0].Rows[0]["DB_StatusMessage"].ToString();
                    }
                    else
                    {
                        TempData["Msg"] = ds.Tables[0].Rows[0]["DB_StatusMessage"].ToString();
                    }
                }
                else
                {
                    TempData["Msg"] = "Oops some error occured";
                }
            }
            else
            {
                TempData["Msg"] = "Please fill all details.";
            }
            return RedirectToAction("MasterShowColEventCertificate", "Admin");
        }
        [HttpGet]
        public ActionResult DeleteEventcolumnsCertificate(int id, int eventid)
        {
            if (id != 0)
            {
                EventSettingColumns modal = new EventSettingColumns();
                modal.ColumnID = id;
                modal.EventId = eventid;
                ds = masterHelper.FetchEventColumns("DeleteEventcolumnsCertificate", modal, Convert.ToInt16(Session["LoginId"]));
                if (ds != null && ds.Tables[0].Rows.Count > 0)
                {
                    if (ds.Tables[0].Rows[0]["DB_STATUS"].ToString() == "S")
                    {
                        TempData["Msg"] = ds.Tables[0].Rows[0]["DB_StatusMessage"].ToString();
                    }
                    else
                    {
                        TempData["Msg"] = ds.Tables[0].Rows[0]["DB_StatusMessage"].ToString();
                    }
                }
                else
                {
                    TempData["Msg"] = "Oops some error occured";
                }
            }
            else
            {
                TempData["Msg"] = "Oops some error occured";
            }
            return RedirectToAction("MasterShowColEventCertificate", "Admin");
        }
        [HttpGet]
        public ActionResult DeactiveEventcolumnsCertificate(int id)
        {
            if (id != 0)
            {
                EventSettingColumns modal = new EventSettingColumns();
                modal.ColumnID = id;
                ds = masterHelper.FetchEventColumns("DeActiveEventcolumnsCertificate", modal, Convert.ToInt16(Session["LoginId"]));
                if (ds != null && ds.Tables[0].Rows.Count > 0)
                {
                    if (ds.Tables[0].Rows[0]["DB_STATUS"].ToString() == "S")
                    {
                        TempData["Msg"] = ds.Tables[0].Rows[0]["DB_StatusMessage"].ToString();
                    }
                    else
                    {
                        TempData["Msg"] = ds.Tables[0].Rows[0]["DB_StatusMessage"].ToString();
                    }
                }
                else
                {
                    TempData["Msg"] = "Oops some error occured";
                }
            }
            else
            { TempData["Msg"] = "Oops some error occured"; }


            return RedirectToAction("MasterShowColEventCertificate", "Admin");
        }
        [HttpGet]
        public ActionResult ActiveEventcolumnsCertificate(int id)
        {

            if (id != 0)
            {
                EventSettingColumns modal = new EventSettingColumns();
                modal.ColumnID = id;
                ds = masterHelper.FetchEventColumns("ActiveEventcolumnsCertificate", modal, Convert.ToInt16(Session["LoginId"]));
                if (ds != null && ds.Tables[0].Rows.Count > 0)
                {
                    if (ds.Tables[0].Rows[0]["DB_STATUS"].ToString() == "S")
                    {
                        TempData["Msg"] = ds.Tables[0].Rows[0]["DB_StatusMessage"].ToString();
                    }
                    else
                    {
                        TempData["Msg"] = ds.Tables[0].Rows[0]["DB_StatusMessage"].ToString();
                    }
                }
                else
                {
                    TempData["Msg"] = "Oops some error occured";
                }
            }
            else
            {
                TempData["Msg"] = "Oops some error occured";
            }
            return RedirectToAction("MasterShowColEventCertificate", "Admin");
        }


        #endregion columns end
        #region code MasterPrintSetupEvent
        public ActionResult MasterPrintSetupEvent()
        {
            EventPrintSetup EC = new EventPrintSetup();
            ds = masterHelper.FetchEventPrintSetup("Get Date PrintSetup", EC, Convert.ToInt16(Session["LoginId"]));
            if (ds != null && ds.Tables[0].Rows.Count > 0)
            {
                EC.EventDrop = CommonFunctions.ToList<EventDropdownList>(ds.Tables[0]);
            }
            return View(EC);
        }
        [HttpGet]
        public ActionResult GetColumnsByFilter(int EventId = 0)
        {
            StringBuilder sb = new StringBuilder();
            EventSettingCol2 model = new EventSettingCol2();
            try
            {
                model.EventId = EventId;
                ds = masterHelper.FetchEventPrintSetup2("Get Date PrintSetup", model, Convert.ToInt16(Session["LoginId"]));
                if (ds != null && ds.Tables[1].Rows.Count > 0)
                {
                    model.EventcolumnsDrop = CommonFunctions.ToList<EventSettingCol2>(ds.Tables[1]);
                }
                if (ds != null && ds.Tables[2].Rows.Count > 0)
                {
                    model.ResultSetting = CommonFunctions.ToList<EventPrintSet>(ds.Tables[2]);
                }
                //  string partial1Html = RenderPartialViewToString("_PrintSetupColumns", model.EventcolumnsDrop);
                var combinedModel = new EventSettingCol2
                {
                    EventcolumnsDrop = model.EventcolumnsDrop,
                    ResultSetting = model.ResultSetting
                };

                // Pass both to the partial view
                string partial1Html = RenderPartialViewToString("_PrintSetupColumns", combinedModel);
                string width = ds.Tables[0].Rows[0]["BadgeWidth"].ToString();   // e.g., "252"
                string height = ds.Tables[0].Rows[0]["BadgeHeight"].ToString(); // e.g., "161"
                string backgroundImage = "../Content/EventBadge/" + ds.Tables[0].Rows[0]["EventBadgePhoto"].ToString();

                string mstyle = $"width: {width}px; height: {height}px; background-image: url('{backgroundImage}'); background-size: cover; position: relative; border: 1px solid #000; page-break-after: always;";

                sb.Append("<div  style=\"" + mstyle + "\">");
                sb.Append("<div id='divprint' class='text-center canvas'>");
                sb.Append("<div id='canvas' class='text-center'>");
                // Add content as absolutely positioned inside the badge
                if (ds.Tables[2].Rows.Count > 0)
                {
                    //photo
                    for (int i = 0; ds.Tables[2].Rows.Count > i; i++)
                    {
                        for (int j = 0; ds.Tables[2].Rows.Count > j; j++)
                        {
                            string mstyleloopcode = $"margin-top: {ds.Tables[2].Rows[j]["MarginTop"].ToString().ToLower()}px; margin-bottom: {ds.Tables[2].Rows[j]["MarginBottom"].ToString().ToLower()}px; margin-left: {ds.Tables[2].Rows[j]["MarginLeft"].ToString().ToLower()}px; margin-right: {ds.Tables[2].Rows[j]["MarginRight"].ToString().ToLower()}px;";

                            if (ds.Tables[2].Rows[j]["columName"].ToString().ToLower() == "photo" && ds.Tables[2].Rows[i]["columName"].ToString().ToLower() == "photo")
                            {
                                sb.Append("<div class='draggable' style=\"" + mstyleloopcode + "\"><img src='../Content/BadgePhoto/Dummybadgephoto.jpg' height='" + ds.Tables[2].Rows[j]["TotalHeight"].ToString().ToLower() + "px' width=' " + ds.Tables[2].Rows[j]["TotalWidth"].ToString().ToLower() + "'px'></div>");
                            }
                        }
                        //name
                        for (int k = 0; ds.Tables[2].Rows.Count > k; k++)
                        {
                            string mstyleloopcode = $"margin-top: {ds.Tables[2].Rows[k]["MarginTop"].ToString().ToLower()}px; margin-bottom: {ds.Tables[2].Rows[k]["MarginBottom"].ToString().ToLower()}px; margin-left: {ds.Tables[2].Rows[k]["MarginLeft"].ToString().ToLower()}px; margin-right: {ds.Tables[2].Rows[k]["MarginRight"].ToString().ToLower()}px;font-size: {ds.Tables[2].Rows[k]["FontSize"].ToString().ToLower()}px; font: {ds.Tables[2].Rows[k]["FontName"].ToString().ToLower()};font-weight: {ds.Tables[2].Rows[k]["FontWeight"].ToString().ToLower()};";

                            if (ds.Tables[2].Rows[k]["columName"].ToString().ToLower() == "name" && ds.Tables[2].Rows[i]["columName"].ToString().ToLower() == "name")
                            {
                                sb.Append("<div class='draggable' style=\"" + mstyleloopcode + "\">Display Exhibitor Name</div>");
                            }
                        }
                        // designation
                        for (int l = 0; ds.Tables[2].Rows.Count > l; l++)
                        {
                            string mstyleloopcode = $"margin-top: {ds.Tables[2].Rows[l]["MarginTop"].ToString().ToLower()}px; margin-bottom: {ds.Tables[2].Rows[l]["MarginBottom"].ToString().ToLower()}px; margin-left: {ds.Tables[2].Rows[l]["MarginLeft"].ToString().ToLower()}px; margin-right: {ds.Tables[2].Rows[l]["MarginRight"].ToString().ToLower()}px;font-size: {ds.Tables[2].Rows[l]["FontSize"].ToString().ToLower()}px; font: {ds.Tables[2].Rows[l]["FontName"].ToString().ToLower()};font-weight: {ds.Tables[2].Rows[l]["FontWeight"].ToString().ToLower()};";

                            if (ds.Tables[2].Rows[l]["columName"].ToString().ToLower() == "designation" && ds.Tables[2].Rows[i]["columName"].ToString().ToLower() == "designation")
                            {
                                sb.Append("<div class='draggable' style=\"" + mstyleloopcode + "\">Display Designation Name</div>");
                            }
                        }
                        //company
                        for (int m = 0; ds.Tables[2].Rows.Count > m; m++)
                        {
                            string mstyleloopcode = $"margin-top: {ds.Tables[2].Rows[m]["MarginTop"].ToString().ToLower()}px; margin-bottom: {ds.Tables[2].Rows[m]["MarginBottom"].ToString().ToLower()}px; margin-left: {ds.Tables[2].Rows[m]["MarginLeft"].ToString().ToLower()}px; margin-right: {ds.Tables[2].Rows[m]["MarginRight"].ToString().ToLower()}px;font-size: {ds.Tables[2].Rows[m]["FontSize"].ToString().ToLower()}px; font: {ds.Tables[2].Rows[m]["FontName"].ToString().ToLower()};font-weight: {ds.Tables[2].Rows[m]["FontWeight"].ToString().ToLower()};";

                            if (ds.Tables[2].Rows[m]["columName"].ToString().ToLower() == "company" && ds.Tables[2].Rows[i]["columName"].ToString().ToLower() == "company")
                            {
                                sb.Append("<div class='draggable' style=\"" + mstyleloopcode + "\">Display Company Name</div>");
                            }
                        }
                        //mobile
                        for (int p = 0; ds.Tables[2].Rows.Count > p; p++)
                        {
                            string mstyleloopcode = $"margin-top: {ds.Tables[2].Rows[p]["MarginTop"].ToString().ToLower()}px; margin-bottom: {ds.Tables[2].Rows[p]["MarginBottom"].ToString().ToLower()}px; margin-left: {ds.Tables[2].Rows[p]["MarginLeft"].ToString().ToLower()}px; margin-right: {ds.Tables[2].Rows[p]["MarginRight"].ToString().ToLower()}px;font-size: {ds.Tables[2].Rows[p]["FontSize"].ToString().ToLower()}px; font: {ds.Tables[2].Rows[p]["FontName"].ToString().ToLower()};font-weight: {ds.Tables[2].Rows[p]["FontWeight"].ToString().ToLower()};";

                            if (ds.Tables[2].Rows[p]["columName"].ToString().ToLower() == "mobile" && ds.Tables[2].Rows[i]["columName"].ToString().ToLower() == "mobile")
                            {
                                sb.Append("<div class='draggable' style=\"" + mstyleloopcode + "\">Display Mobile Number</div>");
                            }
                        }
                        //email
                        for (int q = 0; ds.Tables[2].Rows.Count > q; q++)
                        {
                            string mstyleloopcode = $"margin-top: {ds.Tables[2].Rows[q]["MarginTop"].ToString().ToLower()}px; margin-bottom: {ds.Tables[2].Rows[q]["MarginBottom"].ToString().ToLower()}px; margin-left: {ds.Tables[2].Rows[q]["MarginLeft"].ToString().ToLower()}px; margin-right: {ds.Tables[2].Rows[q]["MarginRight"].ToString().ToLower()}px;font-size: {ds.Tables[2].Rows[q]["FontSize"].ToString().ToLower()}px; font: {ds.Tables[2].Rows[q]["FontName"].ToString().ToLower()};font-weight: {ds.Tables[2].Rows[q]["FontWeight"].ToString().ToLower()};";

                            if (ds.Tables[2].Rows[q]["columName"].ToString().ToLower() == "email" && ds.Tables[2].Rows[i]["columName"].ToString().ToLower() == "email")
                            {
                                sb.Append("<div class='draggable' style=\"" + mstyleloopcode + "\">Display Email Name</div>");
                            }
                        }
                        //barcode
                        for (int n = 0; ds.Tables[2].Rows.Count > n; n++)
                        {
                            string mstyleloopcode = $"margin-top: {ds.Tables[2].Rows[n]["MarginTop"].ToString().ToLower()}px; margin-bottom: {ds.Tables[2].Rows[n]["MarginBottom"].ToString().ToLower()}px; margin-left: {ds.Tables[2].Rows[n]["MarginLeft"].ToString().ToLower()}px; margin-right: {ds.Tables[2].Rows[n]["MarginRight"].ToString().ToLower()}px;";

                            if (ds.Tables[2].Rows[n]["columName"].ToString().ToLower() == "barcode" && ds.Tables[2].Rows[i]["columName"].ToString().ToLower() == "barcode")
                            {
                                sb.Append("<div class='draggable' style=\"" + mstyleloopcode + "\"><img src='../Content/QRCodesScanner/DummyQR.png' height='" + ds.Tables[2].Rows[n]["TotalHeight"].ToString().ToLower() + "px' width='" + ds.Tables[2].Rows[n]["TotalWidth"].ToString().ToLower() + "'px'></div>");
                            }
                        }
                        //category
                        for (int o = 0; ds.Tables[2].Rows.Count > o; o++)
                        {
                            string mstyleloopcode = $"margin-top: {ds.Tables[2].Rows[o]["MarginTop"].ToString().ToLower()}px; margin-bottom: {ds.Tables[2].Rows[o]["MarginBottom"].ToString().ToLower()}px; margin-left: {ds.Tables[2].Rows[o]["MarginLeft"].ToString().ToLower()}px; margin-right: {ds.Tables[2].Rows[o]["MarginRight"].ToString().ToLower()}px;font-size: {ds.Tables[2].Rows[o]["FontSize"].ToString().ToLower()}px; font: {ds.Tables[2].Rows[o]["FontName"].ToString().ToLower()};font-weight: {ds.Tables[2].Rows[o]["FontWeight"].ToString().ToLower()};";

                            if (ds.Tables[2].Rows[o]["columName"].ToString().ToLower() == "category" && ds.Tables[2].Rows[i]["columName"].ToString().ToLower() == "category")
                            {
                                sb.Append("<div class='draggable' style=\"" + mstyleloopcode + "\" id='mcat'>Display Category Name</div>");
                            }
                        }
                    }
                }
                else
                {
                    if (ds.Tables[1].Rows[0]["Photo"].ToString().ToLower() == "true")
                    {
                        sb.Append("<div id='divimage' class='draggable'><img src='../Content/BadgePhoto/Dummybadgephoto.jpg' height='75px' width='75px'></div>");
                    }
                    if (ds.Tables[1].Rows[0]["RName"].ToString().ToLower() == "true")
                    {
                        sb.Append("<div  id='divname'>Display Exhibitor Name</div>");
                    }
                    if (ds.Tables[1].Rows[0]["Designation"].ToString().ToLower() == "true")
                    {
                        sb.Append("<div  id='divdesignation'>Designation Name</div>");
                    }
                    if (ds.Tables[1].Rows[0]["Company"].ToString().ToLower() == "true")
                    {
                        sb.Append("<div  id='divcompany'>Company Name</div>");
                    }
                    if (ds.Tables[1].Rows[0]["Mobile"].ToString().ToLower() == "true")
                    {
                        sb.Append("<div  id='divMobile'>Mobile Number</div>");
                    }
                    if (ds.Tables[1].Rows[0]["Email"].ToString().ToLower() == "true")
                    {
                        sb.Append("<div  id='divEmail'>Email Id</div>");
                    }
                    if (ds.Tables[1].Rows[0]["BarCode"].ToString().ToLower() == "true")
                    {
                        sb.Append("<div  id='divbarcode'><img src='../Content/QRCodesScanner/DummyQR.png' height='75px' width='75px'></div>");
                    }
                    if (ds.Tables[1].Rows[0]["Category"].ToString().ToLower() == "true")
                    {
                        sb.Append("<div  id='divcartegory'>Category Name</div>");
                    }
                }
                sb.Append("</div>");
                sb.Append("</div>");
                sb.Append("</div>");
                string mBadge = sb.ToString();
                return Json(new { partial1 = partial1Html, partial2 = mBadge }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                TempData["Msg"] = ex + " Oops some error occured. Redirected to login page";
                return new HttpStatusCodeResult(500, "Internal Error: " + ex.Message);
            }
        }
        // Helper method to render partial view as string (same as previous)
        protected string RenderPartialViewToString(string viewName, object model)
        {
            if (string.IsNullOrEmpty(viewName))
                viewName = ControllerContext.RouteData.GetRequiredString("action");

            ViewData.Model = model;

            using (var sw = new StringWriter())
            {
                var viewResult = ViewEngines.Engines.FindPartialView(ControllerContext, viewName);
                var viewContext = new ViewContext(ControllerContext, viewResult.View, ViewData, TempData, sw);
                viewResult.View.Render(viewContext, sw);
                viewResult.ViewEngine.ReleaseView(ControllerContext, viewResult.View);
                return sw.GetStringBuilder().ToString();
            }
        }

        [HttpPost]
        public ActionResult MasterPrintSetupEvent(string[] PrintId, string[] ColumName, string[] MarginTop, string[] MarginBottom, string[] MarginLeft,
           string[] MarginRight, string[] TotalHeight, string[] TotalWidth, string[] FontSize,
           string[] FontName, string[] FontWeight, string[] Position, int EventId, EventPrintSet model)
        {
            try
            {
                List<EventPrintSetInner> list = new List<EventPrintSetInner>();
                model.EventId = EventId;
                for (int i = 0; i < ColumName.Length; i++)
                {
                    if (string.IsNullOrWhiteSpace(ColumName[i]) || string.IsNullOrWhiteSpace(MarginTop[i]) || string.IsNullOrWhiteSpace(MarginBottom[i]) || string.IsNullOrWhiteSpace(MarginLeft[i]) ||
                        string.IsNullOrWhiteSpace(MarginRight[i]) || (ColumName[i] != "Photo" && ColumName[i] != "BarCode" && (string.IsNullOrWhiteSpace(FontSize[i]) || string.IsNullOrWhiteSpace(FontName[i]) ||
                         string.IsNullOrWhiteSpace(FontWeight[i]))) || string.IsNullOrWhiteSpace(Position[i]))
                    {
                        TempData["Msg"] = "All fields are required. Please ensure no row is left blank.";
                        return RedirectToAction("MasterPrintSetupEvent", "Admin");
                    }
                    else
                    {

                        EventPrintSetInner pis = new EventPrintSetInner();
                        pis.PrintId = string.IsNullOrWhiteSpace(PrintId[i]) ? 0 : Convert.ToInt32(PrintId[i]);
                        pis.EventId = EventId;
                        pis.ColumName = ColumName[i];
                        pis.MarginTop = Convert.ToInt32(MarginTop[i]);
                        pis.MarginBottom = Convert.ToInt32(MarginBottom[i]);
                        pis.MarginLeft = Convert.ToInt32(MarginLeft[i]);
                        pis.MarginRight = Convert.ToInt32(MarginRight[i]);

                        if (ColumName[i] == "Photo" || ColumName[i] == "BarCode")
                        {
                            string mhasphoto = "";
                            if (ColumName[i] == "Photo")
                            {
                                pis.FontSize = 0.0m;
                                pis.FontName = null;
                                pis.FontWeight = null;
                                pis.TotalHeight = Convert.ToInt32(TotalHeight[0]);
                                pis.TotalWidth = Convert.ToInt32(TotalWidth[0]);
                                mhasphoto = "Yes";
                            }
                            if (ColumName[i] == "BarCode")
                            {
                                pis.FontSize = 0.0m;
                                pis.FontName = null;
                                pis.FontWeight = null;
                                if (mhasphoto == "Yes")
                                {
                                    pis.TotalHeight = Convert.ToInt32(TotalHeight[1]);
                                    pis.TotalWidth = Convert.ToInt32(TotalWidth[1]);
                                }
                                else
                                {
                                    pis.TotalHeight = Convert.ToInt32(TotalHeight[0]);
                                    pis.TotalWidth = Convert.ToInt32(TotalWidth[0]);
                                }
                            }
                        }
                        else
                        {
                            pis.FontSize = Convert.ToDecimal(FontSize[i]);
                            pis.FontName = FontName[i];
                            pis.FontWeight = FontWeight[i];
                            pis.TotalHeight = 0;
                            pis.TotalWidth = 0;
                        }
                        pis.Position = Position[i];
                        list.Add(pis);
                    }
                }
                model.tbl4EventPrint = list;
                ds = masterHelper.FetchEventPrintSetupSave("SavePrintSetup", model, Convert.ToInt16(Session["LoginId"]));
                if (ds != null && ds.Tables.Count > 0 && ds.Tables[0].Rows.Count > 0)
                {
                    if (ds.Tables[0].Rows[0]["DB_STATUS"].ToString().Trim() == "S")
                    {
                        TempData["Msg"] = ds.Tables[0].Rows[0]["DB_StatusMessage"].ToString().Trim();
                        return RedirectToAction("MasterPrintSetupEvent", "Admin");
                    }
                    else
                    {
                        TempData["Msg"] = ds.Tables[0].Rows[0]["DB_StatusMessage"].ToString().Trim();
                        return RedirectToAction("MasterPrintSetupEvent", "Admin");
                    }
                }
                else
                {
                    TempData["Msg"] = ds.Tables[0].Rows[0]["DB_StatusMessage"].ToString().Trim();
                    return RedirectToAction("MasterPrintSetupEvent", "Admin");
                }
            }
            catch (Exception ex)
            {
                TempData["Msg"] = ex.ToString() + "!!! Something went wrong !!!. Please contact admin. ";
                return RedirectToAction("MasterPrintSetupEvent", "Admin");
            }

        }

        [HttpGet]
        public ActionResult DeleteMasterPrintSetupEvent(int id)
        {
            if (id != 0)
            {
                EventPrintSetup modal = new EventPrintSetup();
                modal.PrintId = id;
                ds = masterHelper.FetchEventPrintSetup("DeletePrintSetup", modal, Convert.ToInt16(Session["LoginId"]));
                if (ds != null && ds.Tables[0].Rows.Count > 0)
                {
                    if (ds.Tables[0].Rows[0]["DB_STATUS"].ToString() == "S")
                    {
                        TempData["Msg"] = ds.Tables[0].Rows[0]["DB_StatusMessage"].ToString();
                    }
                    else
                    {
                        TempData["Msg"] = ds.Tables[0].Rows[0]["DB_StatusMessage"].ToString();
                    }
                }
                else
                {
                    TempData["Msg"] = "Oops some error occured";
                }
            }
            else
            {
                TempData["Msg"] = "Oops some error occured";
            }
            return RedirectToAction("MasterPrintSetupEvent", "Admin");
        }
        [HttpGet]
        public ActionResult DeactiveMasterPrintSetupEvent(int id)
        {
            if (id != 0)
            {
                EventPrintSetup modal = new EventPrintSetup();
                modal.PrintId = id;
                ds = masterHelper.FetchEventPrintSetup("DeActivePrintSetup", modal, Convert.ToInt16(Session["LoginId"]));
                if (ds != null && ds.Tables[0].Rows.Count > 0)
                {
                    if (ds.Tables[0].Rows[0]["DB_STATUS"].ToString() == "S")
                    {
                        TempData["Msg"] = ds.Tables[0].Rows[0]["DB_StatusMessage"].ToString();
                    }
                    else
                    {
                        TempData["Msg"] = ds.Tables[0].Rows[0]["DB_StatusMessage"].ToString();
                    }
                }
                else
                {
                    TempData["Msg"] = "Oops some error occured";
                }
            }
            else
            { TempData["Msg"] = "Oops some error occured"; }


            return RedirectToAction("MasterPrintSetupEvent", "Admin");
        }
        [HttpGet]
        public ActionResult ActiveMasterPrintSetupEvent(int id)
        {

            if (id != 0)
            {
                EventPrintSetup modal = new EventPrintSetup();
                modal.PrintId = id;
                ds = masterHelper.FetchEventPrintSetup("ActivePrintSetup", modal, Convert.ToInt16(Session["LoginId"]));
                if (ds != null && ds.Tables[0].Rows.Count > 0)
                {
                    if (ds.Tables[0].Rows[0]["DB_STATUS"].ToString() == "S")
                    {
                        TempData["Msg"] = ds.Tables[0].Rows[0]["DB_StatusMessage"].ToString();
                    }
                    else
                    {
                        TempData["Msg"] = ds.Tables[0].Rows[0]["DB_StatusMessage"].ToString();
                    }
                }
                else
                {
                    TempData["Msg"] = "Oops some error occured";
                }
            }
            else
            {
                TempData["Msg"] = "Oops some error occured";
            }
            return RedirectToAction("MasterPrintSetupEvent", "Admin");
        }
        #endregion
        #region code MasterPrintSetupEventCertificate
        public ActionResult MasterPrintSetupEventCertificate()
        {
            EventPrintSetup EC = new EventPrintSetup();
            ds = masterHelper.FetchEventPrintSetup("Get Date PrintSetupCertificate", EC, Convert.ToInt16(Session["LoginId"]));
            if (ds != null && ds.Tables[0].Rows.Count > 0)
            {
                EC.EventDrop = CommonFunctions.ToList<EventDropdownList>(ds.Tables[0]);
            }
            return View(EC);
        }
        [HttpGet]
        public ActionResult GetColumnsByFilterCertificate(int EventId = 0)
        {
            StringBuilder sb = new StringBuilder();
            EventSettingCol2 model = new EventSettingCol2();
            try
            {
                model.EventId = EventId;
                ds = masterHelper.FetchEventPrintSetup2("Get Date PrintSetupCertificate", model, Convert.ToInt16(Session["LoginId"]));
                if (ds != null && ds.Tables[1].Rows.Count > 0)
                {
                    model.EventcolumnsDrop = CommonFunctions.ToList<EventSettingCol2>(ds.Tables[1]);
                }
                if (ds != null && ds.Tables[2].Rows.Count > 0)
                {
                    model.ResultSetting = CommonFunctions.ToList<EventPrintSet>(ds.Tables[2]);
                }
                //  string partial1Html = RenderPartialViewToString("_PrintSetupColumns", model.EventcolumnsDrop);
                var combinedModel = new EventSettingCol2
                {
                    EventcolumnsDrop = model.EventcolumnsDrop,
                    ResultSetting = model.ResultSetting
                };

                // Pass both to the partial view
                string partial1Html = RenderPartialViewToString("_PrintSetupColumnsCertificate", combinedModel);
                string width = ds.Tables[0].Rows[0]["BadgeWidth"].ToString();   // e.g., "252"
                string height = ds.Tables[0].Rows[0]["BadgeHeight"].ToString(); // e.g., "161"
                string backgroundImage = "../Content/EventBadge/" + ds.Tables[0].Rows[0]["EventBadgePhoto"].ToString();

                string mstyle = $"width: {width}px; height: {height}px; background-image: url('{backgroundImage}'); background-size: cover; position: relative; border: 1px solid #000; page-break-after: always;";

                sb.Append("<div  style=\"" + mstyle + "\">");
                sb.Append("<div id='divprint' class='text-center'>");
                // Add content as absolutely positioned inside the badge
                if (ds.Tables[2].Rows.Count > 0)
                {
                    //photo
                    for (int i = 0; ds.Tables[2].Rows.Count > i; i++)
                    {
                        //barcode
                        for (int n = 0; ds.Tables[2].Rows.Count > n; n++)
                        {
                            string mstyleloopcode = $"margin-top: {ds.Tables[2].Rows[n]["MarginTop"].ToString().ToLower()}px; margin-bottom: {ds.Tables[2].Rows[n]["MarginBottom"].ToString().ToLower()}px; margin-left: {ds.Tables[2].Rows[n]["MarginLeft"].ToString().ToLower()}px; margin-right: {ds.Tables[2].Rows[n]["MarginRight"].ToString().ToLower()}px;";

                            if (ds.Tables[2].Rows[n]["columName"].ToString().ToLower() == "barcode" && ds.Tables[2].Rows[i]["columName"].ToString().ToLower() == "barcode")
                            {
                                sb.Append("<div style=\"" + mstyleloopcode + "\"><img src='../Content/QRCodesScanner/DummyQR.png' height='" + ds.Tables[2].Rows[n]["TotalHeight"].ToString().ToLower() + "px' width='" + ds.Tables[2].Rows[n]["TotalWidth"].ToString().ToLower() + "'px'></div>");
                            }
                        }
                        //name
                        for (int k = 0; ds.Tables[2].Rows.Count > k; k++)
                        {
                            string mstyleloopcode = $"margin-top: {ds.Tables[2].Rows[k]["MarginTop"].ToString().ToLower()}px; margin-bottom: {ds.Tables[2].Rows[k]["MarginBottom"].ToString().ToLower()}px; margin-left: {ds.Tables[2].Rows[k]["MarginLeft"].ToString().ToLower()}px; margin-right: {ds.Tables[2].Rows[k]["MarginRight"].ToString().ToLower()}px;font-size: {ds.Tables[2].Rows[k]["FontSize"].ToString().ToLower()}px; font: {ds.Tables[2].Rows[k]["FontName"].ToString().ToLower()};font-weight: {ds.Tables[2].Rows[k]["FontWeight"].ToString().ToLower()};";

                            if (ds.Tables[2].Rows[k]["columName"].ToString().ToLower() == "name" && ds.Tables[2].Rows[i]["columName"].ToString().ToLower() == "name")
                            {
                                sb.Append("<div style=\"" + mstyleloopcode + "\">Display Certificate Name</div>");
                            }
                        }
                        // designation
                        for (int l = 0; ds.Tables[2].Rows.Count > l; l++)
                        {
                            string mstyleloopcode = $"margin-top: {ds.Tables[2].Rows[l]["MarginTop"].ToString().ToLower()}px; margin-bottom: {ds.Tables[2].Rows[l]["MarginBottom"].ToString().ToLower()}px; margin-left: {ds.Tables[2].Rows[l]["MarginLeft"].ToString().ToLower()}px; margin-right: {ds.Tables[2].Rows[l]["MarginRight"].ToString().ToLower()}px;font-size: {ds.Tables[2].Rows[l]["FontSize"].ToString().ToLower()}px; font: {ds.Tables[2].Rows[l]["FontName"].ToString().ToLower()};font-weight: {ds.Tables[2].Rows[l]["FontWeight"].ToString().ToLower()};";

                            if (ds.Tables[2].Rows[l]["columName"].ToString().ToLower() == "designation" && ds.Tables[2].Rows[i]["columName"].ToString().ToLower() == "designation")
                            {
                                sb.Append("<div style=\"" + mstyleloopcode + "\">Display Certificate Designation Name</div>");
                            }
                        }
                        //company
                        for (int m = 0; ds.Tables[2].Rows.Count > m; m++)
                        {
                            string mstyleloopcode = $"margin-top: {ds.Tables[2].Rows[m]["MarginTop"].ToString().ToLower()}px; margin-bottom: {ds.Tables[2].Rows[m]["MarginBottom"].ToString().ToLower()}px; margin-left: {ds.Tables[2].Rows[m]["MarginLeft"].ToString().ToLower()}px; margin-right: {ds.Tables[2].Rows[m]["MarginRight"].ToString().ToLower()}px;font-size: {ds.Tables[2].Rows[m]["FontSize"].ToString().ToLower()}px; font: {ds.Tables[2].Rows[m]["FontName"].ToString().ToLower()};font-weight: {ds.Tables[2].Rows[m]["FontWeight"].ToString().ToLower()};";

                            if (ds.Tables[2].Rows[m]["columName"].ToString().ToLower() == "company" && ds.Tables[2].Rows[i]["columName"].ToString().ToLower() == "company")
                            {
                                sb.Append("<div style=\"" + mstyleloopcode + "\">Display Certificate Company Name</div>");
                            }
                        }
                                            
                    }
                }
                else
                {
                    //if (ds.Tables[1].Rows[0]["Photo"].ToString().ToLower() == "true")
                    //{
                    //    sb.Append("<div id='divimage' class='draggable'><img src='../Content/BadgePhoto/Dummybadgephoto.jpg' height='75px' width='75px'></div>");
                    //}
                    if (ds.Tables[1].Rows[0]["BarCode"].ToString().ToLower() == "true")
                    {
                        sb.Append("<div  id='divbarcode'><img src='../Content/QRCodesScanner/DummyQR.png' height='75px' width='75px'></div>");
                    }
                    if (ds.Tables[1].Rows[0]["RName"].ToString().ToLower() == "true")
                    {
                        sb.Append("<div  id='divname'>Display Certificate Exhibitor Name</div>");
                    }
                    if (ds.Tables[1].Rows[0]["Designation"].ToString().ToLower() == "true")
                    {
                        sb.Append("<div  id='divdesignation'>Certificate Designation Name</div>");
                    }
                    if (ds.Tables[1].Rows[0]["Company"].ToString().ToLower() == "true")
                    {
                        sb.Append("<div  id='divcompany'>Certificate Company Name</div>");
                    }
                }
                sb.Append("</div>");
                sb.Append("</div>");
                string mBadge = sb.ToString();
                return Json(new { partial1 = partial1Html, partial2 = mBadge }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                TempData["Msg"] = ex + " Oops some error occured. Redirected to login page";
                return new HttpStatusCodeResult(500, "Internal Error: " + ex.Message);
            }
        }
        // Helper method to render partial view as string (same as previous)
        protected string RenderPartialViewToStringCertificate(string viewName, object model)
        {
            if (string.IsNullOrEmpty(viewName))
                viewName = ControllerContext.RouteData.GetRequiredString("action");

            ViewData.Model = model;

            using (var sw = new StringWriter())
            {
                var viewResult = ViewEngines.Engines.FindPartialView(ControllerContext, viewName);
                var viewContext = new ViewContext(ControllerContext, viewResult.View, ViewData, TempData, sw);
                viewResult.View.Render(viewContext, sw);
                viewResult.ViewEngine.ReleaseView(ControllerContext, viewResult.View);
                return sw.GetStringBuilder().ToString();
            }
        }

        [HttpPost]
        public ActionResult MasterPrintSetupEventCertifcate(string[] PrintId, string[] ColumName, string[] MarginTop, string[] MarginBottom, string[] MarginLeft,
           string[] MarginRight, string[] TotalHeight, string[] TotalWidth, string[] FontSize,
           string[] FontName, string[] FontWeight, string[] Position, int EventId, EventPrintSet model)
        {
            try
            {
                List<EventPrintSetInner> list = new List<EventPrintSetInner>();
                model.EventId = EventId;
                for (int i = 0; i < ColumName.Length; i++)
                {
                    if (string.IsNullOrWhiteSpace(ColumName[i]) || string.IsNullOrWhiteSpace(MarginTop[i]) || string.IsNullOrWhiteSpace(MarginBottom[i]) || string.IsNullOrWhiteSpace(MarginLeft[i]) ||
                        string.IsNullOrWhiteSpace(MarginRight[i]) || (ColumName[i] != "Photo" && ColumName[i] != "BarCode" && (string.IsNullOrWhiteSpace(FontSize[i]) || string.IsNullOrWhiteSpace(FontName[i]) ||
                         string.IsNullOrWhiteSpace(FontWeight[i]))) || string.IsNullOrWhiteSpace(Position[i]))
                    {
                        TempData["Msg"] = "All fields are required. Please ensure no row is left blank.";
                        return RedirectToAction("MasterPrintSetupEventCertificate", "Admin");
                    }
                    else
                    {

                        EventPrintSetInner pis = new EventPrintSetInner();
                        pis.PrintId = string.IsNullOrWhiteSpace(PrintId[i]) ? 0 : Convert.ToInt32(PrintId[i]);
                        pis.EventId = EventId;
                        pis.ColumName = ColumName[i];
                        pis.MarginTop = Convert.ToInt32(MarginTop[i]);
                        pis.MarginBottom = Convert.ToInt32(MarginBottom[i]);
                        pis.MarginLeft = Convert.ToInt32(MarginLeft[i]);
                        pis.MarginRight = Convert.ToInt32(MarginRight[i]);

                        if (ColumName[i] == "Photo" || ColumName[i] == "BarCode")
                        {
                            string mhasphoto = "";
                            if (ColumName[i] == "Photo")
                            {
                                pis.FontSize = 0.0m;
                                pis.FontName = null;
                                pis.FontWeight = null;
                                pis.TotalHeight = Convert.ToInt32(TotalHeight[0]);
                                pis.TotalWidth = Convert.ToInt32(TotalWidth[0]);
                                mhasphoto = "Yes";
                            }
                            if (ColumName[i] == "BarCode")
                            {
                                pis.FontSize = 0.0m;
                                pis.FontName = null;
                                pis.FontWeight = null;
                                if (mhasphoto == "Yes")
                                {
                                    pis.TotalHeight = Convert.ToInt32(TotalHeight[1]);
                                    pis.TotalWidth = Convert.ToInt32(TotalWidth[1]);
                                }
                                else
                                {
                                    pis.TotalHeight = Convert.ToInt32(TotalHeight[0]);
                                    pis.TotalWidth = Convert.ToInt32(TotalWidth[0]);
                                }
                            }
                        }
                        else
                        {
                            pis.FontSize = Convert.ToDecimal(FontSize[i]);
                            pis.FontName = FontName[i];
                            pis.FontWeight = FontWeight[i];
                            pis.TotalHeight = 0;
                            pis.TotalWidth = 0;
                        }
                        pis.Position = Position[i];
                        list.Add(pis);
                    }
                }
                model.tbl4EventPrint = list;
                ds = masterHelper.FetchEventPrintSetupSave("SavePrintSetupCertificate", model, Convert.ToInt16(Session["LoginId"]));
                if (ds != null && ds.Tables.Count > 0 && ds.Tables[0].Rows.Count > 0)
                {
                    if (ds.Tables[0].Rows[0]["DB_STATUS"].ToString().Trim() == "S")
                    {
                        TempData["Msg"] = ds.Tables[0].Rows[0]["DB_StatusMessage"].ToString().Trim();
                        return RedirectToAction("MasterPrintSetupEventCertificate", "Admin");
                    }
                    else
                    {
                        TempData["Msg"] = ds.Tables[0].Rows[0]["DB_StatusMessage"].ToString().Trim();
                        return RedirectToAction("MasterPrintSetupEventCertificate", "Admin");
                    }
                }
                else
                {
                    TempData["Msg"] = ds.Tables[0].Rows[0]["DB_StatusMessage"].ToString().Trim();
                    return RedirectToAction("MasterPrintSetupEventCertificate", "Admin");
                }
            }
            catch (Exception ex)
            {
                TempData["Msg"] = ex.ToString() + "!!! Something went wrong !!!. Please contact admin. ";
                return RedirectToAction("MasterPrintSetupEventCertificate", "Admin");
            }

        }

        [HttpGet]
        public ActionResult DeleteMasterPrintSetupEventCertifcate(int id)
        {
            if (id != 0)
            {
                EventPrintSetup modal = new EventPrintSetup();
                modal.PrintId = id;
                ds = masterHelper.FetchEventPrintSetup("DeletePrintSetupCertificate", modal, Convert.ToInt16(Session["LoginId"]));
                if (ds != null && ds.Tables[0].Rows.Count > 0)
                {
                    if (ds.Tables[0].Rows[0]["DB_STATUS"].ToString() == "S")
                    {
                        TempData["Msg"] = ds.Tables[0].Rows[0]["DB_StatusMessage"].ToString();
                    }
                    else
                    {
                        TempData["Msg"] = ds.Tables[0].Rows[0]["DB_StatusMessage"].ToString();
                    }
                }
                else
                {
                    TempData["Msg"] = "Oops some error occured";
                }
            }
            else
            {
                TempData["Msg"] = "Oops some error occured";
            }
            return RedirectToAction("MasterPrintSetupEventCertificate", "Admin");
        }
        [HttpGet]
        public ActionResult DeactiveMasterPrintSetupEventCertifcate(int id)
        {
            if (id != 0)
            {
                EventPrintSetup modal = new EventPrintSetup();
                modal.PrintId = id;
                ds = masterHelper.FetchEventPrintSetup("DeActivePrintSetupCertificate", modal, Convert.ToInt16(Session["LoginId"]));
                if (ds != null && ds.Tables[0].Rows.Count > 0)
                {
                    if (ds.Tables[0].Rows[0]["DB_STATUS"].ToString() == "S")
                    {
                        TempData["Msg"] = ds.Tables[0].Rows[0]["DB_StatusMessage"].ToString();
                    }
                    else
                    {
                        TempData["Msg"] = ds.Tables[0].Rows[0]["DB_StatusMessage"].ToString();
                    }
                }
                else
                {
                    TempData["Msg"] = "Oops some error occured";
                }
            }
            else
            { TempData["Msg"] = "Oops some error occured"; }


            return RedirectToAction("MasterPrintSetupEventCertificate", "Admin");
        }
        [HttpGet]
        public ActionResult ActiveMasterPrintSetupEventCertifcate(int id)
        {

            if (id != 0)
            {
                EventPrintSetup modal = new EventPrintSetup();
                modal.PrintId = id;
                ds = masterHelper.FetchEventPrintSetup("ActivePrintSetupCertificate", modal, Convert.ToInt16(Session["LoginId"]));
                if (ds != null && ds.Tables[0].Rows.Count > 0)
                {
                    if (ds.Tables[0].Rows[0]["DB_STATUS"].ToString() == "S")
                    {
                        TempData["Msg"] = ds.Tables[0].Rows[0]["DB_StatusMessage"].ToString();
                    }
                    else
                    {
                        TempData["Msg"] = ds.Tables[0].Rows[0]["DB_StatusMessage"].ToString();
                    }
                }
                else
                {
                    TempData["Msg"] = "Oops some error occured";
                }
            }
            else
            {
                TempData["Msg"] = "Oops some error occured";
            }
            return RedirectToAction("MasterPrintSetupEventCertificate", "Admin");
        }
        #endregion
        #region master excel 
        public ActionResult MasterExcel()
        {
            Excelupload EC = new Excelupload();
            ds = masterHelper.FetchExcelUpload("ViewExcelEvent", EC, Convert.ToInt16(Session["LoginId"]));
            if (ds != null && ds.Tables[0].Rows.Count > 0)
            {
                EC.EventDrop = CommonFunctions.ToList<EventDropdownList>(ds.Tables[0]);
            }
            if (ds != null && ds.Tables[1].Rows.Count > 0)
            {
                EC.EventExcelUpload = CommonFunctions.ToList<mExcel>(ds.Tables[1]);
            }
            if (ds != null && ds.Tables[2].Rows.Count > 0)
            {
                EC.EventExcelUploadListwht = CommonFunctions.ToList<mExcelListWht>(ds.Tables[2]);
            }
            return View(EC);
        }
        [HttpPost]
        public ActionResult MasterExcel(HttpPostedFileBase excelfile, Excelupload modal)
        {
            if (excelfile != null && excelfile.ContentLength > 0)
            {
                using (var workbook = new XLWorkbook(excelfile.InputStream))
                {
                    List<mExcelUploadColumns> list = new List<mExcelUploadColumns>();
                    var worksheet = workbook.Worksheet(1); // first sheet
                    var rows = worksheet.RangeUsed().RowsUsed().Skip(1); // s

                    foreach (var row in rows) // Assuming first row has headers
                    {
                        mExcelUploadColumns mexcel = new mExcelUploadColumns();
                        {
                            string a = row.Cell(1).Value.ToString();

                            mexcel.Category = row.Cell(1).Value.ToString().Trim();
                            mexcel.QrCodeNo = row.Cell(2).Value.ToString().Trim();
                            mexcel.NAME = row.Cell(3).Value.ToString().Trim();
                            mexcel.Designation = row.Cell(4).Value.ToString().Trim();
                            mexcel.Company = row.Cell(5).Value.ToString().Trim();
                            mexcel.Mobile = row.Cell(6).Value.ToString().Trim();
                            mexcel.Email = row.Cell(7).Value.ToString().Trim();
                            mexcel.Country = row.Cell(8).Value.ToString().Trim();
                            if (row.Cell(9).Value.ToString().ToLower() == "paid"|| row.Cell(9).Value.ToString().ToLower() == "yes")
                            {
                                mexcel.PaymentStatus = "true";
                            }
                            else
                            {
                                mexcel.PaymentStatus = "false";
                            }
                            mexcel.PaymentOn = Convert.ToDateTime(row.Cell(10).Value);
                            mexcel.PaymentRemarks = row.Cell(11).Value.ToString().Trim();
                            list.Add(mexcel);
                        }
                    }
                    modal.mExcelUploadListcolumns = list;
                    if (modal.EventId != 0)
                    {
                        ds = masterHelper.FetchExcelUpload("ExcelCode", modal, Convert.ToInt16(Session["LoginId"]));
                        if (ds != null && ds.Tables[0].Rows.Count > 0)
                        {
                            if (ds.Tables[0].Rows[0]["DB_STATUS"].ToString() == "S")
                            {
                                try
                                {
                                    for (int i = 0; ds.Tables[1].Rows.Count > i; i++)
                                    {
                                        using (QRCodeGenerator qrGenerator = new QRCodeGenerator())
                                        using (QRCodeData qrCodeData = qrGenerator.CreateQrCode(ds.Tables[1].Rows[i]["BPathCode"].ToString(), QRCodeGenerator.ECCLevel.Q))
                                        using (QRCode qrCode = new QRCode(qrCodeData))
                                        using (Bitmap qrCodeImage = qrCode.GetGraphic(20))
                                        {
                                            string folderPath = Server.MapPath("~/Content/QRCodesScanner/" + Regex.Replace(ds.Tables[1].Rows[i]["EventName"].ToString(), @"\s+", "") );
                                            if (!Directory.Exists(folderPath.Trim()))
                                                Directory.CreateDirectory(folderPath.Trim());
                                            string fileName = $"{ds.Tables[1].Rows[i]["BPathCode"].ToString() + "_" + Regex.Replace(ds.Tables[1].Rows[i]["EventName"].ToString(), @"\s+", "")}.png";
                                            string fullPath = Path.Combine(folderPath.Trim(), fileName.Trim());
                                            qrCodeImage.Save(fullPath.Trim(), ImageFormat.Png);
                                            string BPath = fileName.Trim();
                                            int BadgeId = Convert.ToInt32(ds.Tables[1].Rows[i]["BadgeId"].ToString());
                                            masterHelper.SaveExcelpathqr("Update PathQR Excel", BadgeId, Convert.ToInt16(ds.Tables[1].Rows[i]["EventId"]), BPath, Convert.ToInt16(Session["LoginId"]));
                                        }
                                    }
                                }
                                catch (Exception ex)
                                {
                                    TempData["Msg"] = ex.Message + " Oops some error occured";
                                    return Json(new { success = false, message = TempData["Msg"] });
                                }



                                TempData["Msg"] = ds.Tables[0].Rows[0]["DB_StatusMessage"].ToString();
                            }
                            else
                            {
                                TempData["Msg"] = ds.Tables[0].Rows[0]["DB_StatusMessage"].ToString();
                            }
                        }
                        else
                        {
                            TempData["Msg"] = "Oops some error occured";
                        }
                    }
                    else
                    {
                        TempData["Msg"] = "Please fill all detsils.";
                    }
                }
            }
            else
            {
                TempData["Msg"] = "Please fill valid excel file.";
            }
            return RedirectToAction("MasterExcel");
        }

        [HttpGet]
        public ActionResult DeleteMasrterExcel(int id)
        {
            if (id != 0)
            {
                Excelupload modal = new Excelupload();
                modal.EventId = id;
                ds = masterHelper.FetchExcelUpload("DelExcel", modal, Convert.ToInt16(Session["LoginId"]));
                if (ds != null && ds.Tables[0].Rows.Count > 0)
                {
                    if (ds.Tables[0].Rows[0]["DB_STATUS"].ToString() == "S")
                    {
                        TempData["Msg"] = ds.Tables[0].Rows[0]["DB_StatusMessage"].ToString();
                    }
                    else
                    {
                        TempData["Msg"] = ds.Tables[0].Rows[0]["DB_StatusMessage"].ToString();
                    }
                }
                else
                {
                    TempData["Msg"] = "Oops some error occured";
                }
            }
            else
            {
                TempData["Msg"] = "Oops some error occured";
            }
            return RedirectToAction("MasterExcel", "Admin");
        }
        #endregion
        #region master badge
        public ActionResult MasterBadge()
        {
            allbadgebyevent EC = new allbadgebyevent();
            ds = masterHelper.FetchBadgeEvent("ViewAllbadgeEvent", EC, Convert.ToInt16(Session["LoginId"]));
            if (ds != null && ds.Tables[0].Rows.Count > 0)
            {
                EC.EventExcelUpload = CommonFunctions.ToList<mExcel>(ds.Tables[0]);
            }
            return View(EC);
        }
        [HttpGet]
        public ActionResult DeleteMasrterBadge(int id)
        {
            if (id != 0)
            {
                Excelupload modal = new Excelupload();
                modal.EventId = id;
                ds = masterHelper.FetchExcelUpload("DelBadgeEvent", modal, Convert.ToInt16(Session["LoginId"]));
                if (ds != null && ds.Tables[0].Rows.Count > 0)
                {
                    if (ds.Tables[0].Rows[0]["DB_STATUS"].ToString() == "S")
                    {
                        TempData["Msg"] = ds.Tables[0].Rows[0]["DB_StatusMessage"].ToString();
                    }
                    else
                    {
                        TempData["Msg"] = ds.Tables[0].Rows[0]["DB_StatusMessage"].ToString();
                    }
                }
                else
                {
                    TempData["Msg"] = "Oops some error occured";
                }
            }
            else
            {
                TempData["Msg"] = "Oops some error occured";
            }
            return RedirectToAction("MasterBadge", "Admin");
        }

        public ActionResult DeleteBadge(int TokenId)
        {
            if (TokenId != 0)
            {
                allbadgebyevent modal = new allbadgebyevent();
                modal.BadgeId = TokenId;
                ds = masterHelper.FetchBadgeEvent("DelPerBadge", modal, Convert.ToInt16(Session["LoginId"]));
                if (ds != null && ds.Tables[0].Rows.Count > 0)
                {
                    if (ds.Tables[0].Rows[0]["DB_STATUS"].ToString() == "S")
                    {
                        TempData["Msg"] = ds.Tables[0].Rows[0]["DB_StatusMessage"].ToString();
                    }
                    else
                    {
                        TempData["Msg"] = ds.Tables[0].Rows[0]["DB_StatusMessage"].ToString();
                    }
                }
                else
                {
                    TempData["Msg"] = "Oops some error occured";
                }
            }
            else
            {
                TempData["Msg"] = "Oops some error occured";
            }
            return RedirectToAction("MasterBadge", "Admin");
        }

        public string GetAllBadges(int EventId = 0)
        {

            allbadgebyevent modal = new allbadgebyevent();
            modal.EventId = EventId;
            ds = masterHelper.FetchBadgeEvent("GetAllBadgeofEvent", modal, Convert.ToInt16(Session["LoginId"]));
            StringBuilder sb = new StringBuilder();
            if (ds != null && ds.Tables.Count > 0 && ds.Tables[0].Rows.Count > 0)
            {
                for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                {
                    sb.Append("<tr>");
                    sb.Append("<td>" + (i + 1) + "&nbsp&nbsp;<a href='/Admin/DeleteBadge?TokenId=" + ds.Tables[0].Rows[i]["BadgeId"] + "' class='btn btn-primary btn-sm btn-with-icon'><span class='icon-paswrd-visible' onclick ='return confirm('Are you sure you want to delete all badge from this event record from database.?')'></span><span>Delete Badge</span></a></td>");
                    sb.Append("<td>" + ds.Tables[0].Rows[i]["Category"] + "</td>");
                    sb.Append("<td>" + ds.Tables[0].Rows[i]["SpecialNo"] + "</td>");
                    sb.Append("<td>" + ds.Tables[0].Rows[i]["Name"] + "</td>");
                    sb.Append("<td>" + ds.Tables[0].Rows[i]["Designation"] + "</td>");
                    sb.Append("<td>" + ds.Tables[0].Rows[i]["Company"] + "</td>");
                    sb.Append("<td>" + ds.Tables[0].Rows[i]["Mobile"] + "</td>");
                    sb.Append("<td>" + ds.Tables[0].Rows[i]["Country"] + "</td>");
                    sb.Append("<td>" + ds.Tables[0].Rows[i]["IDType"] + "</td>");
                    sb.Append("<td>" + ds.Tables[0].Rows[i]["IDNumber"] + "</td>");
                    sb.Append("<td>" + ds.Tables[0].Rows[i]["Photo"] + "</td>");
                    sb.Append("<td>" + ds.Tables[0].Rows[i]["IsPrint"] + "</td>");
                    sb.Append("<td>" + ds.Tables[0].Rows[i]["PrintOn"] + "</td>");
                    sb.Append("<td>" + ds.Tables[0].Rows[i]["IsReprint"] + "</td>");
                    sb.Append("<td>" + ds.Tables[0].Rows[i]["ReprintOn"] + "</td>");
                    sb.Append("<td>" + ds.Tables[0].Rows[i]["Remarks"] + "</td>");
                    sb.Append("<td>" + ds.Tables[0].Rows[i]["DataInsertby"] + "</td>");
                    sb.Append("<td>" + ds.Tables[0].Rows[i]["IsActive"] + "</td>");
                    sb.Append("<td>" + ds.Tables[0].Rows[i]["IsVisit"] + "</td>");
                    sb.Append("<td>" + ds.Tables[0].Rows[i]["VisitOn"] + "</td>");
                    sb.Append("<td>" + ds.Tables[0].Rows[i]["KitCollect"] + "</td>");
                    sb.Append("<td>" + ds.Tables[0].Rows[i]["CollectOn"] + "</td>");
                    sb.Append("<td>" + ds.Tables[0].Rows[i]["CollectBy"] + "</td>");
                    sb.Append("</tr>");
                }
            }
            else
            {
                sb.Append("<tr><td colspan='23'>No Records Found</td></tr>");
            }
            return sb.ToString();
        }


        public Action SaveMailContent(MailContnet modal)
        {
            if (modal.EventId != 0 && modal.MailContent.Trim() != null)
            {
                modal.EventId = modal.EventId;
                modal.MailContent = modal.MailContent;
                ds = masterHelper.FetchEmailContnet("EmailContnet", modal);
                if (ds != null && ds.Tables[0].Rows.Count > 0)
                {
                    if (ds.Tables[0].Rows[0]["DB_STATUS"].ToString() == "S")
                    {
                        Session["EventId"] = ds.Tables[1].Rows[0]["EventID"].ToString();
                        TempData["Msg"] = ds.Tables[0].Rows[0]["DB_StatusMessage"].ToString();
                    }
                    else
                    {
                        TempData["Msg"] = ds.Tables[0].Rows[0]["DB_StatusMessage"].ToString();
                    }
                }
                else
                {
                    TempData["Msg"] = "Oops some error occured";
                }
            }
            else
            {
                TempData["Msg"] = "Please fill all detsils.";
            }
            return null;// MasterEvent();
        }



        #endregion
        #region Bulk QR Code
        public ActionResult BulkQR()
        {
            BulkQR EC = new BulkQR();
            ds = masterHelper.FetchQrCodes("GetCatById", EC, Convert.ToInt16(Session["LoginId"]));
            if (ds != null && ds.Tables[0].Rows.Count > 0)
            {
                EC.EventDrop = CommonFunctions.ToList<EventDropdownList>(ds.Tables[0]);
            }
            return View(EC);
        }
        [HttpPost]
        public string GetCategoryByEvent(int EventId)
        {
            BulkQR EC = new BulkQR();
            EC.EventId = EventId;
            ds = masterHelper.FetchQrCodes("GetCatById", EC, Convert.ToInt16(Session["LoginId"]));
            StringBuilder sb = new StringBuilder();
            if (ds.Tables.Count > 0 && ds.Tables[1].Rows.Count > 0 && ds.Tables[1].Rows[0][1] != DBNull.Value)
            {
                for (int i = 0; i < ds.Tables[1].Rows.Count; i++)
                {

                    sb.Append("<option value=" + ds.Tables[1].Rows[i]["CategoryName"] + ">" + ds.Tables[1].Rows[i]["CategoryName"] + "</option>");
                }
            }
            else
            {
                sb.Append("<option value=" + 0 + ">" + "Select" + "</option>");
            }
            return sb.ToString();
        }
        [HttpPost]
        public ActionResult BulkQR(BulkQR model)
        {
            try
            {
                if (model.EventId != 0 && model.CategoryName != "" && model.Company != "" && model.Total != 0)
                {
                    int mcount = model.Total;
                    for (int i = 0; mcount > i; i++)
                    {
                        ds = masterHelper.FetchQrCodes("SaveBulkQR", model, Convert.ToInt16(Session["LoginId"]));
                        if (ds != null && ds.Tables[0].Rows.Count > 0)
                        {
                            if (ds.Tables[1].Rows[0]["DB_STATUS"].ToString() == "S")
                            {
                                try
                                {
                                    using (QRCodeGenerator qrGenerator = new QRCodeGenerator())
                                    using (QRCodeData qrCodeData = qrGenerator.CreateQrCode(ds.Tables[0].Rows[0]["LastInsertedID"].ToString(), QRCodeGenerator.ECCLevel.Q))
                                    using (QRCode qrCode = new QRCode(qrCodeData))
                                    using (Bitmap qrCodeImage = qrCode.GetGraphic(20))
                                    {
                                        string folderPath = Server.MapPath("~/Content/QRCodesScanner/" + Session["EventName"].ToString());
                                        if (!Directory.Exists(folderPath))
                                            Directory.CreateDirectory(folderPath);
                                        string fileName = $"{ds.Tables[0].Rows[0]["LastInsertedID"].ToString() + "_" + Session["EventName"].ToString()}.png";
                                        string fullPath = Path.Combine(folderPath, fileName);
                                        qrCodeImage.Save(fullPath, ImageFormat.Png);
                                        model.Company = fileName;
                                        model.BadgeId = Convert.ToInt32(ds.Tables[0].Rows[0]["LastInsertedID"].ToString());
                                        ds = masterHelper.FetchQrCodes("UpdateQrCodePath", model, Convert.ToInt16(Session["LoginId"]));
                                    }
                                }
                                catch (Exception ex)
                                {
                                    TempData["Msg"] = ex.Message + " Oops some error occured";
                                }
                            }
                            else
                            {
                                TempData["Msg"] = ds.Tables[1].Rows[0]["DB_StatusMessage"].ToString();
                            }
                        }
                        else
                        {
                            TempData["Msg"] = "Oops some error occured";
                        }
                    }
                    TempData["Msg"] = "Qr Code generated successfully.";
                }
                else
                {
                    TempData["Msg"] = "Fill Valid details";
                }
                return RedirectToAction("BulkQR", "Admin");
            }
            catch (Exception ex)
            {
                TempData["Msg"] = ex.Message + " Fill Valid details";
                return RedirectToAction("BulkQR", "Admin");
            }
        }


        #endregion
        #region Code for KioskBg
        public ActionResult KioskSetting()
        {
            Kiosk EC = new Kiosk();
            ds = masterHelper.FetchEventKiosk("Kiosk", EC, Convert.ToInt16(Session["LoginId"]));
            if (ds != null && ds.Tables[0].Rows.Count > 0)
            {
                EC.EventDrop = CommonFunctions.ToList<EventDropdownList>(ds.Tables[0]);
            }
            if (ds != null && ds.Tables[1].Rows.Count > 0)
            {
                EC.KioskList = CommonFunctions.ToList<Kiosk>(ds.Tables[1]);
            }
            return View(EC);
        }
        [HttpPost]
        public ActionResult AddKioskSetting(Kiosk modal, HttpPostedFileBase kioskbg)
        {
            if (Convert.ToInt64(modal.EventId) != 0 && modal.Title.Trim() != null && kioskbg != null)
            {
                if (kioskbg != null && kioskbg.ContentLength > 0)
                {
                    try
                    {
                        string[] allowedExtensions = { ".jpg", ".png", ".jpeg" };
                        string fileExtension = Path.GetExtension(kioskbg.FileName.Trim()).ToLower();
                        if (!allowedExtensions.Contains(fileExtension))
                        {
                            TempData["Message"] = "File type not allowed.";
                            return RedirectToAction("KioskSetting");
                        }
                        string folderPath = Server.MapPath("~/Content/Kiosk/");
                        if (!Directory.Exists(folderPath))
                            Directory.CreateDirectory(folderPath);
                        string fileName = Path.GetFileName(DateTime.Now + "_" + modal.EventName + "_" + kioskbg.FileName);
                        string fullPath = Path.Combine(folderPath.Trim(), fileName.Trim());
                        kioskbg.SaveAs(fullPath.Trim());
                        modal.KioskBg = fileName.Trim();
                    }
                    catch (Exception ex)
                    {
                        TempData["Message"] = "ERROR: " + ex.Message.ToString();
                        return RedirectToAction("KioskSetting", "Admin");
                    }
                }
                ds = masterHelper.FetchEventKiosk("SaveKiosk", modal, Convert.ToInt16(Session["LoginId"]));
                if (ds != null && ds.Tables[0].Rows.Count > 0)
                {
                    if (ds.Tables[0].Rows[0]["DB_STATUS"].ToString() == "S")
                    {
                        TempData["Msg"] = ds.Tables[0].Rows[0]["DB_StatusMessage"].ToString();
                    }
                    else
                    {
                        TempData["Msg"] = ds.Tables[0].Rows[0]["DB_StatusMessage"].ToString();
                    }
                }
                else
                {
                    TempData["Msg"] = "Oops some error occured";
                }
            }
            else
            {
                TempData["Msg"] = "Please fill all detsils.";
            }
            return RedirectToAction("KioskSetting", "Admin");
        }
        [HttpGet]
        public ActionResult DeleteKiosk(int id)
        {
            if (id != 0)
            {
                Kiosk modal = new Kiosk();
                modal.EventId = id;
                ds = masterHelper.FetchEventKiosk("DeleteKiosk", modal, Convert.ToInt16(Session["LoginId"]));
                if (ds != null && ds.Tables[0].Rows.Count > 0)
                {
                    if (ds.Tables[0].Rows[0]["DB_STATUS"].ToString() == "S")
                    {
                        TempData["Msg"] = ds.Tables[0].Rows[0]["DB_StatusMessage"].ToString();
                    }
                    else
                    {
                        TempData["Msg"] = ds.Tables[0].Rows[0]["DB_StatusMessage"].ToString();
                    }
                }
                else
                {
                    TempData["Msg"] = "Oops some error occured";
                }
            }
            else
            {
                TempData["Msg"] = "Oops some error occured";
            }

            return RedirectToAction("KioskSetting", "Admin");
        }
        [HttpGet]
        public ActionResult DeactiveKiosk(int id)
        {
            if (id != 0)
            {
                Kiosk modal = new Kiosk();
                modal.EventId = id;
                ds = masterHelper.FetchEventKiosk("DeActiveKiosk", modal, Convert.ToInt16(Session["LoginId"]));
                if (ds != null && ds.Tables[0].Rows.Count > 0)
                {
                    if (ds.Tables[0].Rows[0]["DB_STATUS"].ToString() == "S")
                    {
                        TempData["Msg"] = ds.Tables[0].Rows[0]["DB_StatusMessage"].ToString();
                    }
                    else
                    {
                        TempData["Msg"] = ds.Tables[0].Rows[0]["DB_StatusMessage"].ToString();
                    }
                }
                else
                {
                    TempData["Msg"] = "Oops some error occured";
                }
            }
            else
            { TempData["Msg"] = "Oops some error occured"; }
            return RedirectToAction("KioskSetting", "Admin");
        }
        [HttpGet]
        public ActionResult ActiveKiosk(int id)
        {
            if (id != 0)
            {
                Kiosk modal = new Kiosk();
                modal.EventId = id;
                ds = masterHelper.FetchEventKiosk("ActiveKiosk", modal, Convert.ToInt16(Session["LoginId"]));
                if (ds != null && ds.Tables[0].Rows.Count > 0)
                {
                    if (ds.Tables[0].Rows[0]["DB_STATUS"].ToString() == "S")
                    {
                        TempData["Msg"] = ds.Tables[0].Rows[0]["DB_StatusMessage"].ToString();
                    }
                    else
                    {
                        TempData["Msg"] = ds.Tables[0].Rows[0]["DB_StatusMessage"].ToString();
                    }
                }
                else
                {
                    TempData["Msg"] = "Oops some error occured";
                }
            }
            else
            {
                TempData["Msg"] = "Oops some error occured";
            }
            return RedirectToAction("KioskSetting", "Admin");
        }
        #endregion
    }
}


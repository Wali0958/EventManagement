using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Reflection;
using System.Web;
using System.Web.UI.WebControls;

namespace EventManagement.Models
{
    public class Master_Helper
    {

        string con_str = ConfigurationManager.ConnectionStrings["ConnectionStr"].ConnectionString;
        SqlCommand cmd = new SqlCommand();
        SqlDataAdapter adp = new SqlDataAdapter();


        public DataSet FetchCountData(string Action,int EventId)
        {
            DataSet ds = new DataSet();
            SqlConnection con = new SqlConnection(con_str);
            con.Open();
            try
            {
                cmd = new SqlCommand("usp_MasterDashboard", con);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@pAction", Action);
                cmd.Parameters.AddWithValue("@pEventId", EventId);
                adp = new SqlDataAdapter(cmd);
                adp.Fill(ds);
                return ds;
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                if (con != null)
                {
                    con.Close();
                    con.Dispose();
                }
            }
        }
        public DataSet FetchLogin(string Action, LoginModels Login)
        {
            DataSet ds = new DataSet();
            SqlConnection con = new SqlConnection(con_str);
            con.Open();
            try
            {
                cmd = new SqlCommand("usp_Login", con);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@pAction", Action);
                cmd.Parameters.AddWithValue("@pUserName", Login.UserName);
                cmd.Parameters.AddWithValue("@pPassword", Login.Pass);
                adp = new SqlDataAdapter(cmd);
                adp.Fill(ds);
                return ds;
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                if (con != null)
                {
                    con.Close();
                    con.Dispose();
                }
            }
        }
        public DataSet FetchEvent(string Action, EventModels Event, int LoginId)
        {
            DataSet ds = new DataSet();
            SqlConnection con = new SqlConnection(con_str);
            con.Open();
            try
            {
                cmd = new SqlCommand("usp_MasterEvent", con);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@pAction", Action);
                cmd.Parameters.AddWithValue("@pEventId", Event.EventId);
                cmd.Parameters.AddWithValue("@pEventName", Event.EventName);
                cmd.Parameters.AddWithValue("@pEventStartOn", Event.EventStartOn);
                cmd.Parameters.AddWithValue("@pEventEndOn", Event.EventEndOn);
                cmd.Parameters.AddWithValue("@pEventLocation", Event.EventLocation);
                cmd.Parameters.AddWithValue("@pEventHeadName", Event.EventHeadName);
                cmd.Parameters.AddWithValue("@pContactNo", Event.ContactNo);
                cmd.Parameters.AddWithValue("@pEventDay", Event.EventDay);
                cmd.Parameters.AddWithValue("@pEventBadgePhoto", Event.EventBadgePhoto);
                cmd.Parameters.AddWithValue("@pBadgeHeight", Event.BadgeHeight);
                cmd.Parameters.AddWithValue("@pBadgeWidth", Event.BadgeWidth);
                cmd.Parameters.AddWithValue("@pPrintType", Event.PrintType);
                cmd.Parameters.AddWithValue("@pIsActive", Event.IsActive);
                cmd.Parameters.AddWithValue("@pLoginId", LoginId);
                cmd.Parameters.AddWithValue("@pLoginIP", HttpContext.Current.Request.UserHostAddress);
                adp = new SqlDataAdapter(cmd);
                adp.Fill(ds);
                return ds;
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                if (con != null)
                {
                    con.Close();
                    con.Dispose();
                }
            }
        }
        public DataSet FetchEventCategory(string Action, EventCat EC, int LoginId)
        {
            DataSet ds = new DataSet();
            SqlConnection con = new SqlConnection(con_str);
            con.Open();
            try
            {
                cmd = new SqlCommand("usp_MasterEvent", con);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@pAction", Action);
                cmd.Parameters.AddWithValue("@pCategoryId", EC.CategoryId);
                cmd.Parameters.AddWithValue("@pEventId", EC.EventId);
                cmd.Parameters.AddWithValue("@pCategoryName", EC.CategoryName);
                cmd.Parameters.AddWithValue("@pCategroySubName", EC.CategroySubName);
                cmd.Parameters.AddWithValue("@pIsKitAllow", EC.IsKitAllow);
                cmd.Parameters.AddWithValue("@pIsPaymentAllow", EC.IsPaymentAllow);
                cmd.Parameters.AddWithValue("@pIsActive", EC.IsActive);
                cmd.Parameters.AddWithValue("@pLoginId", LoginId);
                cmd.Parameters.AddWithValue("@pLoginIP", HttpContext.Current.Request.UserHostAddress);
                adp = new SqlDataAdapter(cmd);
                adp.Fill(ds);
                return ds;
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                if (con != null)
                {
                    con.Close();
                    con.Dispose();
                }
            }
        }

        public DataSet FetchUser(string Action, MasterUser mUser, int LoginId)
        {
            DataSet ds = new DataSet();
            SqlConnection con = new SqlConnection(con_str);
            con.Open();
            try
            {
                cmd = new SqlCommand("usp_Login", con);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@pAction", Action);
                cmd.Parameters.AddWithValue("@pEventId", mUser.EventId);
                cmd.Parameters.AddWithValue("@pUserName", mUser.UserEmail);
                cmd.Parameters.AddWithValue("@pRole", mUser.UserRole);
                cmd.Parameters.AddWithValue("@pPassword", mUser.Password);
                cmd.Parameters.AddWithValue("@pIsActive", mUser.IsActive);
                cmd.Parameters.AddWithValue("@pLoginId", mUser.LoginId);
                cmd.Parameters.AddWithValue("@pLoginIP", HttpContext.Current.Request.UserHostAddress);
                adp = new SqlDataAdapter(cmd);
                adp.Fill(ds);
                return ds;
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                if (con != null)
                {
                    con.Close();
                    con.Dispose();
                }
            }
        }
        public DataSet FetchEventColumns(string Action, EventSettingColumns mEColumn, int LoginId)
        {
            DataSet ds = new DataSet();
            SqlConnection con = new SqlConnection(con_str);
            con.Open();
            try
            {
                cmd = new SqlCommand("usp_MasterEventColumns", con);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@pAction", Action);
                cmd.Parameters.AddWithValue("@pColumnID", mEColumn.ColumnID);
                cmd.Parameters.AddWithValue("@pEventId", mEColumn.EventId);
                cmd.Parameters.AddWithValue("@pCategory", mEColumn.Category);
                cmd.Parameters.AddWithValue("@pSpecialNo", mEColumn.SpecialNo);
                cmd.Parameters.AddWithValue("@pName", mEColumn.RName);
                cmd.Parameters.AddWithValue("@pDesignation", mEColumn.Designation);
                cmd.Parameters.AddWithValue("@pCompany", mEColumn.Company);
                cmd.Parameters.AddWithValue("@pMobile", mEColumn.Mobile);
                cmd.Parameters.AddWithValue("@pEmail", mEColumn.Email);
                cmd.Parameters.AddWithValue("@pCountry", mEColumn.Country);
                cmd.Parameters.AddWithValue("@pIdType", mEColumn.IdType);
                cmd.Parameters.AddWithValue("@pIdNumber", mEColumn.IdNumber);
                cmd.Parameters.AddWithValue("@pPhoto", mEColumn.Photo);
                cmd.Parameters.AddWithValue("@pBarCode", mEColumn.BarCode);
                cmd.Parameters.AddWithValue("@pRemarks", mEColumn.Remarks);
                cmd.Parameters.AddWithValue("@pLoginId", LoginId);
                cmd.Parameters.AddWithValue("@pLoginIP", HttpContext.Current.Request.UserHostAddress);
                adp = new SqlDataAdapter(cmd);
                adp.Fill(ds);
                return ds;
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                if (con != null)
                {
                    con.Close();
                    con.Dispose();
                }
            }
        }
        public DataSet FetchDashboard(string Action, Dashboard dash, int LoginId, int EventId)
        {
            DataSet ds = new DataSet();
            SqlConnection con = new SqlConnection(con_str);
            con.Open();
            try
            {
                cmd = new SqlCommand("usp_MasterDashboard", con);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@pAction", Action);
                cmd.Parameters.AddWithValue("@pId", dash.Id);
                cmd.Parameters.AddWithValue("@pIsActive", dash.IsActive);
                cmd.Parameters.AddWithValue("@pLoginId", LoginId);
                cmd.Parameters.AddWithValue("@pEventId", EventId);
                cmd.Parameters.AddWithValue("@pLoginIP", HttpContext.Current.Request.UserHostAddress);
                adp = new SqlDataAdapter(cmd);
                adp.Fill(ds);
                return ds;
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                if (con != null)
                {
                    con.Close();
                    con.Dispose();
                }
            }
        }
        public DataSet FetchEventPrintSetup(string Action, EventPrintSetup EPS, int LoginId)
        {
            DataSet ds = new DataSet();
            SqlConnection con = new SqlConnection(con_str);
            con.Open();
            try
            {
                cmd = new SqlCommand("usp_MasterPrintSetup", con);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@pAction", Action);
                cmd.Parameters.AddWithValue("@pPrintId", EPS.PrintId);
                cmd.Parameters.AddWithValue("@pEventId", EPS.EventId);
                cmd.Parameters.AddWithValue("@pLoginId", LoginId);
                cmd.Parameters.AddWithValue("@pLoginIP", HttpContext.Current.Request.UserHostAddress);
                adp = new SqlDataAdapter(cmd);
                adp.Fill(ds);
                return ds;
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                if (con != null)
                {
                    con.Close();
                    con.Dispose();
                }
            }
        }
        public DataSet FetchEventPrintSetup2(string Action, EventSettingCol2 EPS, int LoginId)
        {
            DataSet ds = new DataSet();
            SqlConnection con = new SqlConnection(con_str);
            con.Open();
            try
            {
                cmd = new SqlCommand("usp_MasterPrintSetup", con);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@pAction", Action);
                cmd.Parameters.AddWithValue("@pEventId", EPS.EventId);
                cmd.Parameters.AddWithValue("@pLoginId", LoginId);
                cmd.Parameters.AddWithValue("@pLoginIP", HttpContext.Current.Request.UserHostAddress);
                adp = new SqlDataAdapter(cmd);
                adp.Fill(ds);
                return ds;
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                if (con != null)
                {
                    con.Close();
                    con.Dispose();
                }
            }
        }

        public DataSet FetchEventPrintSetupSave(string Action, EventPrintSet EPS, int LoginId)
        {
            DataSet ds = new DataSet();
            SqlConnection con = new SqlConnection(con_str);
            con.Open();
            try
            {
                cmd = new SqlCommand("usp_MasterPrintSetup", con);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@pAction", Action);
                cmd.Parameters.AddWithValue("@pEventId", EPS.EventId);
                cmd.Parameters.AddWithValue("@pLoginId", LoginId);
                cmd.Parameters.AddWithValue("@pLoginIP", HttpContext.Current.Request.UserHostAddress);
                if (EPS.tbl4EventPrint != null)
                {
                    cmd.Parameters.AddWithValue("@ptbl4EventPrint", CommonFunctions.ToDataTable(EPS.tbl4EventPrint));
                }
                adp = new SqlDataAdapter(cmd);
                adp.Fill(ds);
                return ds;
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                if (con != null)
                {
                    con.Close();
                    con.Dispose();
                }
            }
        }
        public DataSet FetchExcelUpload(string Action, Excelupload Exx, int LoginId)
        {
            DataSet ds = new DataSet();
            SqlConnection con = new SqlConnection(con_str);
            con.Open();
            try
            {
                cmd = new SqlCommand("usp_EventExcelBadgeReport", con);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@pAction", Action);
                cmd.Parameters.AddWithValue("@pEventId", Exx.EventId);
                cmd.Parameters.AddWithValue("@pLoginId", LoginId);
                cmd.Parameters.AddWithValue("@pLoginIP", HttpContext.Current.Request.UserHostAddress);
                if (Exx.mExcelUploadListcolumns != null)
                {
                    cmd.Parameters.AddWithValue("@pTbl4Excel", CommonFunctions.ToDataTable(Exx.mExcelUploadListcolumns));
                }
                adp = new SqlDataAdapter(cmd);
                adp.Fill(ds);
                return ds;
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                if (con != null)
                {
                    con.Close();
                    con.Dispose();
                }
            }
        }
        public DataSet FetchBadgeEvent(string Action, allbadgebyevent Exx, int LoginId)
        {
            DataSet ds = new DataSet();
            SqlConnection con = new SqlConnection(con_str);
            con.Open();
            try
            {
                cmd = new SqlCommand("usp_EventExcelBadgeReport", con);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@pAction", Action);
                cmd.Parameters.AddWithValue("@pEventId", Exx.EventId);
                cmd.Parameters.AddWithValue("@pBadgeId", Exx.BadgeId);
                cmd.Parameters.AddWithValue("@pLoginId", LoginId);
                cmd.Parameters.AddWithValue("@pLoginIP", HttpContext.Current.Request.UserHostAddress);
                adp = new SqlDataAdapter(cmd);
                adp.Fill(ds);
                return ds;
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                if (con != null)
                {
                    con.Close();
                    con.Dispose();
                }
            }
        }

        public DataSet FetchQrCodes(string Action, BulkQR Exx, int LoginId)
        {
            DataSet ds = new DataSet();
            SqlConnection con = new SqlConnection(con_str);
            con.Open();
            try
            {
                cmd = new SqlCommand("usp_EventExcelBadgeReport", con);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@pAction", Action);
                cmd.Parameters.AddWithValue("@pEventId", Exx.EventId);
                cmd.Parameters.AddWithValue("@pBadgeId", Exx.BadgeId);
                cmd.Parameters.AddWithValue("@pCategory", Exx.CategoryName);
                cmd.Parameters.AddWithValue("@pCompany", Exx.Company);
                cmd.Parameters.AddWithValue("@pLoginId", LoginId);
                cmd.Parameters.AddWithValue("@pLoginIP", HttpContext.Current.Request.UserHostAddress);
                adp = new SqlDataAdapter(cmd);
                adp.Fill(ds);
                return ds;
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                if (con != null)
                {
                    con.Close();
                    con.Dispose();
                }
            }
        }

        public DataSet FetchHomePage(string Action, GetPrintBadge mbadgemodel, int LoginId)
        {
            DataSet ds = new DataSet();
            SqlConnection con = new SqlConnection(con_str);
            con.Open();
            try
            {
                cmd = new SqlCommand("usp_EventExcelBadgeReport", con);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@pAction", Action);
                cmd.Parameters.AddWithValue("@pEventId", mbadgemodel.EventId);
                cmd.Parameters.AddWithValue("@pBadgeId", mbadgemodel.BadgeId);
                cmd.Parameters.AddWithValue("@pCompany", mbadgemodel.Company);
                cmd.Parameters.AddWithValue("@pCategory", mbadgemodel.Category);
                cmd.Parameters.AddWithValue("@pIsActive", mbadgemodel.IsActive);
                cmd.Parameters.AddWithValue("@pIsPrint", mbadgemodel.IsPrint);
                cmd.Parameters.AddWithValue("@pSearcText", mbadgemodel.SearchText);
                cmd.Parameters.AddWithValue("@pLoginId", LoginId);
                cmd.Parameters.AddWithValue("@pLoginIP", HttpContext.Current.Request.UserHostAddress);
                adp = new SqlDataAdapter(cmd);
                adp.Fill(ds);
                return ds;
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                if (con != null)
                {
                    con.Close();
                    con.Dispose();
                }
            }
        }
        public DataSet FetchAddBadgePrintNew(string Action, AddBadgeCode mbadgemodel, int LoginId)
        {
            DataSet ds = new DataSet();
            SqlConnection con = new SqlConnection(con_str);
            con.Open();
            try
            {
                cmd = new SqlCommand("usp_MasterBadgeSavePrint", con);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@pAction", Action);
                cmd.Parameters.AddWithValue("@pEventId", mbadgemodel.EventId);
                cmd.Parameters.AddWithValue("@pBadgeId", mbadgemodel.BadgeId);
                cmd.Parameters.AddWithValue("@pIsActive", mbadgemodel.IsActive);
                cmd.Parameters.AddWithValue("@pLoginId", LoginId);
                cmd.Parameters.AddWithValue("@pLoginIP", HttpContext.Current.Request.UserHostAddress);
                adp = new SqlDataAdapter(cmd);
                adp.Fill(ds);
                return ds;
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                if (con != null)
                {
                    con.Close();
                    con.Dispose();
                }
            }
        }

        public DataSet SaveMasterBadgeOnebyOne(string Action, allbadgebyevent mbadgemodel, int LoginId)
        {
            DataSet ds = new DataSet();
            SqlConnection con = new SqlConnection(con_str);
            con.Open();
            try
            {
                cmd = new SqlCommand("usp_MasterBadgeSavePrint", con);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@pAction", Action);
                cmd.Parameters.AddWithValue("@pEventId", mbadgemodel.EventId);
                cmd.Parameters.AddWithValue("@pBadgeId", mbadgemodel.BadgeId);
                cmd.Parameters.AddWithValue("@pCategory", mbadgemodel.Category);
                cmd.Parameters.AddWithValue("@pSpecialNo", mbadgemodel.SpecialNo);
                cmd.Parameters.AddWithValue("@pName", mbadgemodel.Name);
                cmd.Parameters.AddWithValue("@pDesignation", mbadgemodel.Designation);
                cmd.Parameters.AddWithValue("@pCompany", mbadgemodel.Company);
                cmd.Parameters.AddWithValue("@pMobile", mbadgemodel.Mobile);
                cmd.Parameters.AddWithValue("@pEMail", mbadgemodel.Email);
                cmd.Parameters.AddWithValue("@pCountry", mbadgemodel.Country);
                cmd.Parameters.AddWithValue("@pIDType", mbadgemodel.IDType);
                cmd.Parameters.AddWithValue("@pIDNumber", mbadgemodel.IDNumber);
                cmd.Parameters.AddWithValue("@pPhoto", mbadgemodel.Photo);
                cmd.Parameters.AddWithValue("@pIsActive", mbadgemodel.IsActive);
                cmd.Parameters.AddWithValue("@pBpath", mbadgemodel.BPath);
                cmd.Parameters.AddWithValue("@pLoginId", LoginId);
                cmd.Parameters.AddWithValue("@pLoginIP", HttpContext.Current.Request.UserHostAddress);
                adp = new SqlDataAdapter(cmd);
                adp.Fill(ds);
                return ds;
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                if (con != null)
                {
                    con.Close();
                    con.Dispose();
                }
            }
        }

        public DataSet SaveExcelpathqr(string Action,int BadgeId, int EventId, string path, int LoginId)
        {
            DataSet ds = new DataSet();
            SqlConnection con = new SqlConnection(con_str);
            con.Open();
            try
            {
                cmd = new SqlCommand("usp_EventExcelBadgeReport", con);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@pAction", Action);
                cmd.Parameters.AddWithValue("@pEventId", EventId);
                cmd.Parameters.AddWithValue("@pBadgeId", BadgeId);
                cmd.Parameters.AddWithValue("@pCategory", path); 
                cmd.Parameters.AddWithValue("@pLoginId", LoginId);
                cmd.Parameters.AddWithValue("@pLoginIP", HttpContext.Current.Request.UserHostAddress);
                adp = new SqlDataAdapter(cmd);
                adp.Fill(ds);
                return ds;
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                if (con != null)
                {
                    con.Close();
                    con.Dispose();
                }
            }
        }



        public DataSet BadgePrintDataBulk(string Action, int EventId, int BadgeId, int LoginId)
        {
            DataSet ds = new DataSet();
            SqlConnection con = new SqlConnection(con_str);
            con.Open();
            try
            {
                cmd = new SqlCommand("usp_MasterBadgeSavePrint", con);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@pAction", Action);
                cmd.Parameters.AddWithValue("@pEventId", EventId);
                cmd.Parameters.AddWithValue("@pBadgeId", BadgeId);
                cmd.Parameters.AddWithValue("@pLoginId", LoginId);
                cmd.Parameters.AddWithValue("@pLoginIP", HttpContext.Current.Request.UserHostAddress);
                adp = new SqlDataAdapter(cmd);
                adp.Fill(ds);
                return ds;
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                if (con != null)
                {
                    con.Close();
                    con.Dispose();
                }
            }
        }

        public DataSet BadgePrintDataBulk2(string Action, int EventId, int BadgeId, int LoginId, string kitstatus, string CollectBy, string Payment, string commentsname)
        {
            DataSet ds = new DataSet();
            SqlConnection con = new SqlConnection(con_str);
            con.Open();
            try
            {
                cmd = new SqlCommand("usp_MasterBadgeSavePrint", con);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@pAction", Action);
                cmd.Parameters.AddWithValue("@pEventId", EventId);
                cmd.Parameters.AddWithValue("@pBadgeId", BadgeId);
                cmd.Parameters.AddWithValue("@pKitstatus", kitstatus);
                cmd.Parameters.AddWithValue("@pCollectBy", CollectBy);
                cmd.Parameters.AddWithValue("@pPayment", Payment);
                cmd.Parameters.AddWithValue("@pRemarks", commentsname);
                cmd.Parameters.AddWithValue("@pLoginId", LoginId);
                cmd.Parameters.AddWithValue("@pLoginIP", HttpContext.Current.Request.UserHostAddress);
                adp = new SqlDataAdapter(cmd);
                adp.Fill(ds);
                return ds;
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                if (con != null)
                {
                    con.Close();
                    con.Dispose();
                }
            }
        }

        public DataSet FetchIdByBatch(string Action, BadgeDetailsScanning mbadgemodel, int LoginId)
        {
            DataSet ds = new DataSet();
            SqlConnection con = new SqlConnection(con_str);
            con.Open();
            try
            {
                cmd = new SqlCommand("usp_MasterBadgeSavePrint", con);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@pAction", Action);
                cmd.Parameters.AddWithValue("@pEventId", mbadgemodel.EventId);
                cmd.Parameters.AddWithValue("@pBadgeId", mbadgemodel.BadgeId);
                cmd.Parameters.AddWithValue("@pSearchText", mbadgemodel.SearchText);
                cmd.Parameters.AddWithValue("@pIsActive", mbadgemodel.IsActive);
                cmd.Parameters.AddWithValue("@pLoginId", LoginId);
                cmd.Parameters.AddWithValue("@pLoginIP", HttpContext.Current.Request.UserHostAddress);
                adp = new SqlDataAdapter(cmd);
                adp.Fill(ds);
                return ds;
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                if (con != null)
                {
                    con.Close();
                    con.Dispose();
                }
            }
        }

        public DataSet PrintBadgeCode(string Action, int BadgeId, string CollectBy, string remarks,
            int EventId,string multiprint)
        {
            DataSet ds = new DataSet();
            SqlConnection con = new SqlConnection(con_str);
            con.Open();
            try
            {
                cmd = new SqlCommand("usp_MasterPrintSetup", con);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@pAction", Action);
                cmd.Parameters.AddWithValue("@pEventId", EventId);
                cmd.Parameters.AddWithValue("@pBadgeId", BadgeId);
                cmd.Parameters.AddWithValue("@pCollectby", CollectBy);
                cmd.Parameters.AddWithValue("@pRemarks", remarks);
                cmd.Parameters.AddWithValue("@pMultiprint", multiprint);

                adp = new SqlDataAdapter(cmd);
                adp.Fill(ds);
                return ds;
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                if (con != null)
                {
                    con.Close();
                    con.Dispose();
                }
            }
        }

        public DataSet SearchData(string Action, string text, int EventId, int loginid)
        {
            DataSet ds = new DataSet();
            SqlConnection con = new SqlConnection(con_str);
            con.Open();
            try
            {
                cmd = new SqlCommand("usp_MasterBadgeSavePrint", con);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@pAction", Action);
                cmd.Parameters.AddWithValue("@pSearchText", text);
                cmd.Parameters.AddWithValue("@pEventId", EventId);
                cmd.Parameters.AddWithValue("@pLoginId", loginid);
                adp = new SqlDataAdapter(cmd);
                adp.Fill(ds);
                return ds;
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                if (con != null)
                {
                    con.Close();
                    con.Dispose();
                }
            }
        }

        public DataSet BadgeUpdate(string Action,  int BadgeId, string Name, string Company, string Designation)
        {
            DataSet ds = new DataSet();
            SqlConnection con = new SqlConnection(con_str);
            con.Open();
            try
            {
                cmd = new SqlCommand("usp_MasterBadgeSavePrint", con);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@pAction", Action);
                cmd.Parameters.AddWithValue("@pBadgeId", BadgeId);
                cmd.Parameters.AddWithValue("@pName", Name);
                cmd.Parameters.AddWithValue("@pCompany", Company);
                cmd.Parameters.AddWithValue("@pDesignation", Designation);
                adp = new SqlDataAdapter(cmd);
                adp.Fill(ds);
                return ds;
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                if (con != null)
                {
                    con.Close();
                    con.Dispose();
                }
            }
        }


        public DataSet FetchEventKiosk(string Action, Kiosk EC, int LoginId)
        {
            DataSet ds = new DataSet();
            SqlConnection con = new SqlConnection(con_str);
            con.Open();
            try
            {
                cmd = new SqlCommand("usp_MasterEvent", con);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@pAction", Action);
                cmd.Parameters.AddWithValue("@pKioskId", EC.KioskId);
                cmd.Parameters.AddWithValue("@pEventId", EC.EventId);
                cmd.Parameters.AddWithValue("@pKioskBg", EC.KioskBg);
                cmd.Parameters.AddWithValue("@pTitle", EC.Title);
                cmd.Parameters.AddWithValue("@pIsActive", EC.IsActive);
                cmd.Parameters.AddWithValue("@pLoginId", LoginId);
                cmd.Parameters.AddWithValue("@pLoginIP", HttpContext.Current.Request.UserHostAddress);
                adp = new SqlDataAdapter(cmd);
                adp.Fill(ds);
                return ds;
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                if (con != null)
                {
                    con.Close();
                    con.Dispose();
                }
            }
        }




    }
}
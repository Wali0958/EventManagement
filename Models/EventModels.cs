using DocumentFormat.OpenXml.ExtendedProperties;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Web;
using static ClosedXML.Excel.XLPredefinedFormat;

namespace EventManagement.Models
{
    #region ForMaster Event
    public class EventModels
    {
        public int EventId { get; set; }

        [Required(ErrorMessage = "EventName is required.")]
        public string EventName { get; set; }

        [Required(ErrorMessage = "Start on is required.")]
        public System.DateTime? EventStartOn { get; set; }

        [Required(ErrorMessage = "End on is required.")]
        public System.DateTime? EventEndOn { get; set; }

        [Required(ErrorMessage = "Location is required.")]
        public string EventLocation { get; set; }

        [Required(ErrorMessage = "Head name is required.")]
        public string EventHeadName { get; set; }

        [Required(ErrorMessage = "Mobile no is required.")]
        public string ContactNo { get; set; }
        public string EventDay { get; set; }
        public string EventBadgePhoto { get; set; }
        public decimal BadgeHeight { get; set; }
        public decimal BadgeWidth { get; set; }
        public int PrintType { get; set; }
        public bool IsActive { get; set; }
        public List<EventModelsList> mList { get; set; }
    }
    public class EventModelsList
    {
        public int EventId { get; set; }
        public string EventName { get; set; }
        public System.DateTime? EventStartOn { get; set; }
        public System.DateTime? EventEndOn { get; set; }
        public string EventLocation { get; set; }
        public string EventHeadName { get; set; }
        public string ContactNo { get; set; }
        public string EventDay { get; set; }
        public string EventBadgePhoto { get; set; }
        public decimal BadgeHeight { get; set; }
        public decimal BadgeWidth { get; set; }
        public int PrintType { get; set; }
        public bool IsActive { get; set; }
    }
    public class EventDropdownList
    {
        public int EventId { get; set; }
        public string EventName { get; set; }
    }
    #endregion
    #region master event category
    public class EventCat
    {
        public int CategoryId { get; set; }
        public int EventId { get; set; }
        public string CategoryName { get; set; }
        public string CategroySubName { get; set; }
        public bool IsActive { get; set; }
        public bool IsKitAllow { get; set; }
        public bool IsPaymentAllow { get; set; }
      
        public List<EventCatNew> EventCatAll { get; set; }
        public List<EventDropdownList> EventDrop { get; set; }
    }
    public class EventCatNew
    {
        public int CategoryId { get; set; }
        public int EventId { get; set; }
        public string CategoryName { get; set; }
        public string CategroySubName { get; set; }
        public bool IsActive { get; set; }
        public bool IsKitAllow { get; set; }
        public bool IsPaymentAllow { get; set; }
       
        public string EventName { get; set; }
    }

    public class EventCountry
    {
        public int CountryId { get; set; }
        public string CountryName { get; set; }
        public bool IsActive { get; set; }
    }
    #endregion
    #region master event user
    public class MasterUser
    {
        public int LoginId { get; set; }
        public int UserRole { get; set; }
        public int EventId { get; set; }
        public string UserEmail { get; set; }
        public string Password { get; set; }
        public bool IsActive { get; set; }
        public List<EventDropdownList> EventDrop { get; set; }
        public List<UserRoleModel> URole { get; set; }
        public List<masteruserall> UserDisplay { get; set; }
    }

    public class masteruserall
    {
        public int LoginId { get; set; }
        public string UserEmail { get; set; }
        public string RoleType { get; set; }
        public string RoleAbbr { get; set; }
        public string EventName { get; set; }
        public string Password { get; set; }
        public bool IsActive { get; set; }
    }
    public class UserRoleModel
    {
        public int UserRole { get; set; }
        public string RoleType { get; set; }
        public string RoleAbbr { get; set; }
    }
    #endregion
    #region event setting columns
    public class EventSettingColumns
    {
        public long ColumnID { get; set; }
        public int EventId { get; set; }
        public string EventName { get; set; } = "";
        public bool Category { get; set; }
        public bool SpecialNo { get; set; }
        public bool RName { get; set; }
        public bool Designation { get; set; }
        public bool Company { get; set; }
        public bool Mobile { get; set; }
        public bool Email { get; set; }
        public bool Country { get; set; }
        public bool IdType { get; set; }
        public bool IdNumber { get; set; }
        public bool Photo { get; set; }
        public bool BarCode { get; set; }
        public bool IsActive { get; set; }
        public bool Remarks { get; set; }

        public List<EventDropdownList> EventDrop { get; set; }
        public List<EventSettingColumns> EventcolumnsDrop { get; set; }
    }
    #endregion
    #region master event print setup
    public class EventPrintSetup
    {
        public int PrintId { get; set; }
        public int MarginTop { get; set; }
        public int MarginBottom { get; set; }
        public int MarginLeft { get; set; }
        public int MarginRight { get; set; }
        public int TotalHeight { get; set; }
        public int TotalWidth { get; set; }
        public int EventId { get; set; }
        public bool IsActive { get; set; }
        public List<EventDropdownList> EventDrop { get; set; }
        public List<EventSettingCol2> EventcolumnsDrop { get; set; }

    }
    public class EventSettingCol2
    {
        public int EventId { get; set; }
        public bool Category { get; set; }
        public bool RName { get; set; }
        public bool Designation { get; set; }
        public bool Company { get; set; }
        public bool Mobile { get; set; }
        public bool Email { get; set; }
        public bool Country { get; set; }
        public bool IdType { get; set; }
        public bool IdNumber { get; set; }
        public bool Photo { get; set; }
        public bool BarCode { get; set; }
        public bool Remarks { get; set; }
        public bool SpecialNo { get; set; }
        public bool IsActive { get; set; }
        public List<EventSettingCol2> EventcolumnsDrop { get; set; }
        public List<EventPrintSet> ResultSetting { get; set; }
    }

    public class EventPrintSet
    {
        public int PrintId { get; set; }
        public int EventId { get; set; }
        public string ColumName { get; set; }
        public int MarginTop { get; set; }
        public int MarginBottom { get; set; }
        public int MarginLeft { get; set; }
        public int MarginRight { get; set; }
        public int TotalHeight { get; set; }
        public int TotalWidth { get; set; }
        public Decimal FontSize { get; set; }
        public string FontName { get; set; }
        public string FontWeight { get; set; }
        public int Position { get; set; }
        public bool IsActive { get; set; }
        public List<EventPrintSetInner> tbl4EventPrint { get; set; }
    }

    public class EventPrintSetInner
    {
        public int PrintId { get; set; }
        public int EventId { get; set; }
        public string ColumName { get; set; }
        public int MarginTop { get; set; }
        public int MarginBottom { get; set; }
        public int MarginLeft { get; set; }
        public int MarginRight { get; set; }
        public int TotalHeight { get; set; }
        public int TotalWidth { get; set; }
        public Decimal FontSize { get; set; }
        public string FontName { get; set; }
        public string FontWeight { get; set; }
        public string Position { get; set; }
    }


    #endregion
    #region masterexcel
    public class Excelupload
    {
        public int EventId { get; set; }
        public string EventName { get; set; }
        public int TotalExcel { get; set; }
        public List<EventDropdownList> EventDrop { get; set; }
        public List<mExcel> EventExcelUpload { get; set; }
        public List<mExcelUploadColumns> mExcelUploadListcolumns { get; set; }
        public List<mExcelListWht> EventExcelUploadListwht { get; set; }
    }
    public class mExcel
    {
        public int EventId { get; set; }
        public string EventName { get; set; }
        public int TotalExcel { get; set; }
        public System.DateTime EventDate { get; set; }
    }

    public class mExcelListWht
    {
        public int EventId { get; set; }
        public string Name { get; set; }
        public string Company { get; set; }
        public string Designation { get; set; }
        public string EventName { get; set; }
    }

    public class mExcelUploadColumns
    {
        public string Category { get; set; }
        public string QrCodeNo { get; set; }
        public string NAME { get; set; }
        public string Designation { get; set; }
        public string Company { get; set; }
        public string Mobile { get; set; }
        public string Email { get; set; }
        public string Country { get; set; }
        public string PaymentStatus { get; set; }
        public System.DateTime PaymentOn { get; set; }
        public string PaymentRemarks { get; set; }
        //public string IDType { get; set; }
        //public string IDNumber { get; set; }
    }
    #endregion
    #region master badge delete section
    public class allbadgebyevent
    {

        public int BadgeId { get; set; }
        public int EventId { get; set; }
        public string Category { get; set; }
        public string SpecialNo { get; set; }
        public string Name { get; set; }
        public string Designation { get; set; }
        public string Company { get; set; }
        public string Mobile { get; set; }
        public string Email { get; set; }
        public string Country { get; set; }
        public string IDType { get; set; }
        public string IDNumber { get; set; }
        public string Photo { get; set; }
        public string BPath { get; set; }
        public bool IsPrint { get; set; }
        public System.DateTime PrintOn { get; set; }
        public bool IsReprint { get; set; }
        public System.DateTime ReprintOn { get; set; }
        public string Remarks { get; set; }
        public bool DataInsertby { get; set; }
        public bool IsActive { get; set; }
        public bool IsVisit { get; set; }
        public System.DateTime VisitOn { get; set; }
        public bool KitCollect { get; set; }
        public System.DateTime CollectOn { get; set; }
        public int CollectBy { get; set; }
        public List<mExcel> EventExcelUpload { get; set; }
    }

    #endregion
    #region master Qr
    public class BulkQR
    {
        public int EventId { get; set; }
        public string CategoryName { get; set; }
        public string Company { get; set; }
        public int BadgeId { get; set; }
        public string EventName { get; set; }
        public int Total { get; set; }
        public List<EventDropdownList> EventDrop { get; set; }
        public List<EventCatNew> EventCatAll { get; set; }

    }
    #endregion

    #region Kiosk
    #region master event category
    public class Kiosk
    {
        public Int64 KioskId { get; set; }
        public int EventId { get; set; }
        public string KioskBg { get; set; }
        public string Title { get; set; }
        public bool IsActive { get; set; }
        public string EventName { get; set; }

        public List<Kiosk> KioskList { get; set; }
        public List<EventDropdownList> EventDrop { get; set; }
    } 
    #endregion
    #endregion

}
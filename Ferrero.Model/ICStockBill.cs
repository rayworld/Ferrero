using System;
namespace Ferrero.Model
{
    /// <summary>
    /// ICStockBill:实体类(属性说明自动提取数据库字段的描述信息)
    /// </summary>
    [Serializable]
    public partial class ICStockBill
    {
        public ICStockBill()
        { }
        #region Model
        private string _fbrno;
        private int _finterid;
        private int _ftrantype;
        private DateTime _fdate;
        private string _fbillno;
        private string _fuse;
        private string _fnote;
        private int? _fdcstockid;
        private int? _fscstockid;
        private int? _fdeptid;
        private int? _fempid;
        private int _fsupplyid;
        private int? _fposterid;
        private int? _fcheckerid;
        private int? _ffmanagerid;
        private int? _fsmanagerid;
        private int? _fbillerid;
        private int? _freturnbillinterid;
        private string _fscbillno;
        private int _fhookinterid = 0;
        private int _fvchinterid;
        private int _fposted = 0;
        private int? _fcheckselect = 0;
        private int? _fcurrencyid;
        private int? _fsalestyle;
        private int? _facctid;
        private int _frob = 1;
        private string _frscbillno;
        private int _fstatus = 0;
        private bool _fupstockwhensave = false;
        private bool _fcancellation = false;
        private int _forgbillinterid = 0;
        private int? _fbilltypeid;
        private int? _fpostyle;
        private int? _fmultichecklevel1;
        private int? _fmultichecklevel2;
        private int? _fmultichecklevel3;
        private int? _fmultichecklevel4;
        private int? _fmultichecklevel5;
        private int? _fmultichecklevel6;
        private DateTime? _fmulticheckdate1;
        private DateTime? _fmulticheckdate2;
        private DateTime? _fmulticheckdate3;
        private DateTime? _fmulticheckdate4;
        private DateTime? _fmulticheckdate5;
        private DateTime? _fmulticheckdate6;
        private int? _fcurchecklevel;
        private int? _ftaskid;
        private int? _fresourceid;
        private bool _fbackflushed = false;
        private int _fwbinterid = 0;
        private int _ftranstatus = 0;
        private int? _fzpbillinterid;
        private int? _frelatebrid;
        private int _fpurposeid = 0;
        private Guid _fuuid;
        private int _frelateinvoiceid = 0;
        private DateTime? _foperdate;
        private int? _fimport = 0;
        private int? _fsystemtype = 0;
        private int _fmarketingstyle = 12530;
        private int _fpaybillid = 0;
        private DateTime? _fcheckdate;
        private string _fexplanation = "";
        private string _ffetchadd = "";
        private DateTime? _ffetchdate;
        private int _fmanagerid = 0;
        private int _freftype = 0;
        private int _fseltrantype = 0;
        private int _fchildren = 0;
        private int _fhookstatus = 0;
        private int _factpricevchtplid = 0;
        private int _fplanpricevchtplid = 0;
        private int _fprocid = 0;
        private int _factualvchtplid = 0;
        private int _fplanvchtplid = 0;
        private int? _fbrid;
        private int _fvipcardid = 0;
        private decimal _fvipscore = 0M;
        private decimal _fholisticdiscountrate = 0M;
        private string _fposname = "";
        private int _fworkshiftid = 0;
        private int _fcussentacctid = 0;
        private bool _fzangucount = false;
        private string _fpoordbillno;
        private int _flssrcinterid = 0;
        private DateTime? _fsettledate;
        private int? _fmanagetype;
        private string _fheadselfb0146;
        private string _fheadselfb0147;
        private string _fheadselfb0148;
        private string _fheadselfb0149;
        private string _fheadselfb0150;
        private string _fheadselfb0151;
        private string _fheadselfb0152;
        /// <summary>
        /// 
        /// </summary>
        public string FBrNo
        {
            set { _fbrno = value; }
            get { return _fbrno; }
        }
        /// <summary>
        /// 
        /// </summary>
        public int FInterID
        {
            set { _finterid = value; }
            get { return _finterid; }
        }
        /// <summary>
        /// 
        /// </summary>
        public int FTranType
        {
            set { _ftrantype = value; }
            get { return _ftrantype; }
        }
        /// <summary>
        /// 
        /// </summary>
        public DateTime FDate
        {
            set { _fdate = value; }
            get { return _fdate; }
        }
        /// <summary>
        /// 
        /// </summary>
        public string FBillNo
        {
            set { _fbillno = value; }
            get { return _fbillno; }
        }
        /// <summary>
        /// 
        /// </summary>
        public string FUse
        {
            set { _fuse = value; }
            get { return _fuse; }
        }
        /// <summary>
        /// 
        /// </summary>
        public string FNote
        {
            set { _fnote = value; }
            get { return _fnote; }
        }
        /// <summary>
        /// 
        /// </summary>
        public int? FDCStockID
        {
            set { _fdcstockid = value; }
            get { return _fdcstockid; }
        }
        /// <summary>
        /// 
        /// </summary>
        public int? FSCStockID
        {
            set { _fscstockid = value; }
            get { return _fscstockid; }
        }
        /// <summary>
        /// 
        /// </summary>
        public int? FDeptID
        {
            set { _fdeptid = value; }
            get { return _fdeptid; }
        }
        /// <summary>
        /// 
        /// </summary>
        public int? FEmpID
        {
            set { _fempid = value; }
            get { return _fempid; }
        }
        /// <summary>
        /// 
        /// </summary>
        public int FSupplyID
        {
            set { _fsupplyid = value; }
            get { return _fsupplyid; }
        }
        /// <summary>
        /// 
        /// </summary>
        public int? FPosterID
        {
            set { _fposterid = value; }
            get { return _fposterid; }
        }
        /// <summary>
        /// 
        /// </summary>
        public int? FCheckerID
        {
            set { _fcheckerid = value; }
            get { return _fcheckerid; }
        }
        /// <summary>
        /// 
        /// </summary>
        public int? FFManagerID
        {
            set { _ffmanagerid = value; }
            get { return _ffmanagerid; }
        }
        /// <summary>
        /// 
        /// </summary>
        public int? FSManagerID
        {
            set { _fsmanagerid = value; }
            get { return _fsmanagerid; }
        }
        /// <summary>
        /// 
        /// </summary>
        public int? FBillerID
        {
            set { _fbillerid = value; }
            get { return _fbillerid; }
        }
        /// <summary>
        /// 
        /// </summary>
        public int? FReturnBillInterID
        {
            set { _freturnbillinterid = value; }
            get { return _freturnbillinterid; }
        }
        /// <summary>
        /// 
        /// </summary>
        public string FSCBillNo
        {
            set { _fscbillno = value; }
            get { return _fscbillno; }
        }
        /// <summary>
        /// 
        /// </summary>
        public int FHookInterID
        {
            set { _fhookinterid = value; }
            get { return _fhookinterid; }
        }
        /// <summary>
        /// 
        /// </summary>
        public int FVchInterID
        {
            set { _fvchinterid = value; }
            get { return _fvchinterid; }
        }
        /// <summary>
        /// 
        /// </summary>
        public int FPosted
        {
            set { _fposted = value; }
            get { return _fposted; }
        }
        /// <summary>
        /// 
        /// </summary>
        public int? FCheckSelect
        {
            set { _fcheckselect = value; }
            get { return _fcheckselect; }
        }
        /// <summary>
        /// 
        /// </summary>
        public int? FCurrencyID
        {
            set { _fcurrencyid = value; }
            get { return _fcurrencyid; }
        }
        /// <summary>
        /// 
        /// </summary>
        public int? FSaleStyle
        {
            set { _fsalestyle = value; }
            get { return _fsalestyle; }
        }
        /// <summary>
        /// 
        /// </summary>
        public int? FAcctID
        {
            set { _facctid = value; }
            get { return _facctid; }
        }
        /// <summary>
        /// 
        /// </summary>
        public int FROB
        {
            set { _frob = value; }
            get { return _frob; }
        }
        /// <summary>
        /// 
        /// </summary>
        public string FRSCBillNo
        {
            set { _frscbillno = value; }
            get { return _frscbillno; }
        }
        /// <summary>
        /// 
        /// </summary>
        public int FStatus
        {
            set { _fstatus = value; }
            get { return _fstatus; }
        }
        /// <summary>
        /// 
        /// </summary>
        public bool FUpStockWhenSave
        {
            set { _fupstockwhensave = value; }
            get { return _fupstockwhensave; }
        }
        /// <summary>
        /// 
        /// </summary>
        public bool FCancellation
        {
            set { _fcancellation = value; }
            get { return _fcancellation; }
        }
        /// <summary>
        /// 
        /// </summary>
        public int FOrgBillInterID
        {
            set { _forgbillinterid = value; }
            get { return _forgbillinterid; }
        }
        /// <summary>
        /// 
        /// </summary>
        public int? FBillTypeID
        {
            set { _fbilltypeid = value; }
            get { return _fbilltypeid; }
        }
        /// <summary>
        /// 
        /// </summary>
        public int? FPOStyle
        {
            set { _fpostyle = value; }
            get { return _fpostyle; }
        }
        /// <summary>
        /// 
        /// </summary>
        public int? FMultiCheckLevel1
        {
            set { _fmultichecklevel1 = value; }
            get { return _fmultichecklevel1; }
        }
        /// <summary>
        /// 
        /// </summary>
        public int? FMultiCheckLevel2
        {
            set { _fmultichecklevel2 = value; }
            get { return _fmultichecklevel2; }
        }
        /// <summary>
        /// 
        /// </summary>
        public int? FMultiCheckLevel3
        {
            set { _fmultichecklevel3 = value; }
            get { return _fmultichecklevel3; }
        }
        /// <summary>
        /// 
        /// </summary>
        public int? FMultiCheckLevel4
        {
            set { _fmultichecklevel4 = value; }
            get { return _fmultichecklevel4; }
        }
        /// <summary>
        /// 
        /// </summary>
        public int? FMultiCheckLevel5
        {
            set { _fmultichecklevel5 = value; }
            get { return _fmultichecklevel5; }
        }
        /// <summary>
        /// 
        /// </summary>
        public int? FMultiCheckLevel6
        {
            set { _fmultichecklevel6 = value; }
            get { return _fmultichecklevel6; }
        }
        /// <summary>
        /// 
        /// </summary>
        public DateTime? FMultiCheckDate1
        {
            set { _fmulticheckdate1 = value; }
            get { return _fmulticheckdate1; }
        }
        /// <summary>
        /// 
        /// </summary>
        public DateTime? FMultiCheckDate2
        {
            set { _fmulticheckdate2 = value; }
            get { return _fmulticheckdate2; }
        }
        /// <summary>
        /// 
        /// </summary>
        public DateTime? FMultiCheckDate3
        {
            set { _fmulticheckdate3 = value; }
            get { return _fmulticheckdate3; }
        }
        /// <summary>
        /// 
        /// </summary>
        public DateTime? FMultiCheckDate4
        {
            set { _fmulticheckdate4 = value; }
            get { return _fmulticheckdate4; }
        }
        /// <summary>
        /// 
        /// </summary>
        public DateTime? FMultiCheckDate5
        {
            set { _fmulticheckdate5 = value; }
            get { return _fmulticheckdate5; }
        }
        /// <summary>
        /// 
        /// </summary>
        public DateTime? FMultiCheckDate6
        {
            set { _fmulticheckdate6 = value; }
            get { return _fmulticheckdate6; }
        }
        /// <summary>
        /// 
        /// </summary>
        public int? FCurCheckLevel
        {
            set { _fcurchecklevel = value; }
            get { return _fcurchecklevel; }
        }
        /// <summary>
        /// 
        /// </summary>
        public int? FTaskID
        {
            set { _ftaskid = value; }
            get { return _ftaskid; }
        }
        /// <summary>
        /// 
        /// </summary>
        public int? FResourceID
        {
            set { _fresourceid = value; }
            get { return _fresourceid; }
        }
        /// <summary>
        /// 
        /// </summary>
        public bool FBackFlushed
        {
            set { _fbackflushed = value; }
            get { return _fbackflushed; }
        }
        /// <summary>
        /// 
        /// </summary>
        public int FWBInterID
        {
            set { _fwbinterid = value; }
            get { return _fwbinterid; }
        }
        /// <summary>
        /// 
        /// </summary>
        public int FTranStatus
        {
            set { _ftranstatus = value; }
            get { return _ftranstatus; }
        }
        /// <summary>
        /// 
        /// </summary>
        public int? FZPBillInterID
        {
            set { _fzpbillinterid = value; }
            get { return _fzpbillinterid; }
        }
        /// <summary>
        /// 
        /// </summary>
        public int? FRelateBrID
        {
            set { _frelatebrid = value; }
            get { return _frelatebrid; }
        }
        /// <summary>
        /// 
        /// </summary>
        public int FPurposeID
        {
            set { _fpurposeid = value; }
            get { return _fpurposeid; }
        }
        /// <summary>
        /// 
        /// </summary>
        public Guid FUUID
        {
            set { _fuuid = value; }
            get { return _fuuid; }
        }
        /// <summary>
        /// 
        /// </summary>
        public int FRelateInvoiceID
        {
            set { _frelateinvoiceid = value; }
            get { return _frelateinvoiceid; }
        }
        /// <summary>
        /// 
        /// </summary>
        public DateTime? FOperDate
        {
            set { _foperdate = value; }
            get { return _foperdate; }
        }
        /// <summary>
        /// 
        /// </summary>
        public int? FImport
        {
            set { _fimport = value; }
            get { return _fimport; }
        }
        /// <summary>
        /// 
        /// </summary>
        public int? FSystemType
        {
            set { _fsystemtype = value; }
            get { return _fsystemtype; }
        }
        /// <summary>
        /// 
        /// </summary>
        public int FMarketingStyle
        {
            set { _fmarketingstyle = value; }
            get { return _fmarketingstyle; }
        }
        /// <summary>
        /// 
        /// </summary>
        public int FPayBillID
        {
            set { _fpaybillid = value; }
            get { return _fpaybillid; }
        }
        /// <summary>
        /// 
        /// </summary>
        public DateTime? FCheckDate
        {
            set { _fcheckdate = value; }
            get { return _fcheckdate; }
        }
        /// <summary>
        /// 
        /// </summary>
        public string FExplanation
        {
            set { _fexplanation = value; }
            get { return _fexplanation; }
        }
        /// <summary>
        /// 
        /// </summary>
        public string FFetchAdd
        {
            set { _ffetchadd = value; }
            get { return _ffetchadd; }
        }
        /// <summary>
        /// 
        /// </summary>
        public DateTime? FFetchDate
        {
            set { _ffetchdate = value; }
            get { return _ffetchdate; }
        }
        /// <summary>
        /// 
        /// </summary>
        public int FManagerID
        {
            set { _fmanagerid = value; }
            get { return _fmanagerid; }
        }
        /// <summary>
        /// 
        /// </summary>
        public int FRefType
        {
            set { _freftype = value; }
            get { return _freftype; }
        }
        /// <summary>
        /// 
        /// </summary>
        public int FSelTranType
        {
            set { _fseltrantype = value; }
            get { return _fseltrantype; }
        }
        /// <summary>
        /// 
        /// </summary>
        public int FChildren
        {
            set { _fchildren = value; }
            get { return _fchildren; }
        }
        /// <summary>
        /// 
        /// </summary>
        public int FHookStatus
        {
            set { _fhookstatus = value; }
            get { return _fhookstatus; }
        }
        /// <summary>
        /// 
        /// </summary>
        public int FActPriceVchTplID
        {
            set { _factpricevchtplid = value; }
            get { return _factpricevchtplid; }
        }
        /// <summary>
        /// 
        /// </summary>
        public int FPlanPriceVchTplID
        {
            set { _fplanpricevchtplid = value; }
            get { return _fplanpricevchtplid; }
        }
        /// <summary>
        /// 
        /// </summary>
        public int FProcID
        {
            set { _fprocid = value; }
            get { return _fprocid; }
        }
        /// <summary>
        /// 
        /// </summary>
        public int FActualVchTplID
        {
            set { _factualvchtplid = value; }
            get { return _factualvchtplid; }
        }
        /// <summary>
        /// 
        /// </summary>
        public int FPlanVchTplID
        {
            set { _fplanvchtplid = value; }
            get { return _fplanvchtplid; }
        }
        /// <summary>
        /// 
        /// </summary>
        public int? FBrID
        {
            set { _fbrid = value; }
            get { return _fbrid; }
        }
        /// <summary>
        /// 
        /// </summary>
        public int FVIPCardID
        {
            set { _fvipcardid = value; }
            get { return _fvipcardid; }
        }
        /// <summary>
        /// 
        /// </summary>
        public decimal FVIPScore
        {
            set { _fvipscore = value; }
            get { return _fvipscore; }
        }
        /// <summary>
        /// 
        /// </summary>
        public decimal FHolisticDiscountRate
        {
            set { _fholisticdiscountrate = value; }
            get { return _fholisticdiscountrate; }
        }
        /// <summary>
        /// 
        /// </summary>
        public string FPOSName
        {
            set { _fposname = value; }
            get { return _fposname; }
        }
        /// <summary>
        /// 
        /// </summary>
        public int FWorkShiftId
        {
            set { _fworkshiftid = value; }
            get { return _fworkshiftid; }
        }
        /// <summary>
        /// 
        /// </summary>
        public int FCussentAcctID
        {
            set { _fcussentacctid = value; }
            get { return _fcussentacctid; }
        }
        /// <summary>
        /// 
        /// </summary>
        public bool FZanGuCount
        {
            set { _fzangucount = value; }
            get { return _fzangucount; }
        }
        /// <summary>
        /// 
        /// </summary>
        public string FPOOrdBillNo
        {
            set { _fpoordbillno = value; }
            get { return _fpoordbillno; }
        }
        /// <summary>
        /// 
        /// </summary>
        public int FLSSrcInterID
        {
            set { _flssrcinterid = value; }
            get { return _flssrcinterid; }
        }
        /// <summary>
        /// 
        /// </summary>
        public DateTime? FSettleDate
        {
            set { _fsettledate = value; }
            get { return _fsettledate; }
        }
        /// <summary>
        /// 
        /// </summary>
        public int? FManageType
        {
            set { _fmanagetype = value; }
            get { return _fmanagetype; }
        }
        /// <summary>
        /// 
        /// </summary>
        public string FHeadSelfB0146
        {
            set { _fheadselfb0146 = value; }
            get { return _fheadselfb0146; }
        }
        /// <summary>
        /// 
        /// </summary>
        public string FHeadSelfB0147
        {
            set { _fheadselfb0147 = value; }
            get { return _fheadselfb0147; }
        }
        /// <summary>
        /// 
        /// </summary>
        public string FHeadSelfB0148
        {
            set { _fheadselfb0148 = value; }
            get { return _fheadselfb0148; }
        }
        /// <summary>
        /// 
        /// </summary>
        public string FHeadSelfB0149
        {
            set { _fheadselfb0149 = value; }
            get { return _fheadselfb0149; }
        }
        /// <summary>
        /// 
        /// </summary>
        public string FHeadSelfB0150
        {
            set { _fheadselfb0150 = value; }
            get { return _fheadselfb0150; }
        }
        /// <summary>
        /// 
        /// </summary>
        public string FHeadSelfB0151
        {
            set { _fheadselfb0151 = value; }
            get { return _fheadselfb0151; }
        }

        /// <summary>
        /// 
        /// </summary>
        public string FHeadSelfB0152
        {
            set { _fheadselfb0152 = value; }
            get { return _fheadselfb0152; }
        }
        #endregion Model

    }
}



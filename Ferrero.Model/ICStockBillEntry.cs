using System;
namespace Ferrero.Model
{
    /// <summary>
    /// ICStockBillEntry:实体类(属性说明自动提取数据库字段的描述信息)
    /// </summary>
    [Serializable]
    public partial class ICStockBillEntry
    {
        public ICStockBillEntry()
        { }
        #region Model
        private string _fbrno;
        private int _finterid;
        private int _fentryid;
        private int _fitemid;
        private decimal _fqtymust = 0M;
        private decimal _fqty = 0M;
        private decimal _fprice = 0M;
        private string _fbatchno = "";
        private decimal _famount = 0M;
        private string _fnote;
        private int? _fscbillinterid;
        private string _fscbillno;
        private int _funitid = 0;
        private decimal _fauxprice = 0M;
        private decimal _fauxqty = 0M;
        private decimal _fauxqtymust = 0M;
        private decimal _fqtyactual = 0M;
        private decimal _fauxqtyactual = 0M;
        private decimal _fplanprice = 0M;
        private decimal _fauxplanprice = 0M;
        private int _fsourceentryid = 0;
        private decimal _fcommitqty = 0M;
        private decimal _fauxcommitqty = 0M;
        private DateTime? _fkfdate;
        private int _fkfperiod = 0;
        private int? _fdcspid;
        private int? _fscspid;
        private decimal _fconsignprice = 0M;
        private decimal _fconsignamount = 0M;
        private decimal _fprocesscost = 0M;
        private decimal _fmaterialcost = 0M;
        private decimal _ftaxamount = 0M;
        private string _fmapnumber = "";
        private string _fmapname = "";
        private int _forgbillentryid = 0;
        private int? _foperid = 0;
        private decimal _fplanamount = 0M;
        private decimal _fprocessprice = 0M;
        private decimal _ftaxrate = 0M;
        private int? _fsnlistid;
        private decimal _famtref = 0M;
        private int _fauxpropid = 0;
        private decimal _fcost = 0M;
        private decimal _fpriceref = 0M;
        private decimal _fauxpriceref = 0M;
        private DateTime? _ffetchdate;
        private decimal _fqtyinvoice = 0M;
        private decimal _fqtyinvoicebase = 0M;
        private decimal _funitcost = 0M;
        private decimal _fseccoefficient = 0M;
        private decimal _fsecqty = 0M;
        private decimal _fseccommitqty = 0M;
        private int _fsourcetrantype = 0;
        private int _fsourceinterid = 0;
        private string _fsourcebillno = "";
        private int _fcontractinterid = 0;
        private int _fcontractentryid = 0;
        private string _fcontractbillno = "";
        private string _ficmobillno = "";
        private int _ficmointerid = 0;
        private int _fppbomentryid = 0;
        private int _forderinterid = 0;
        private int _forderentryid = 0;
        private string _forderbillno = "";
        private decimal _fallhookqty = 0M;
        private decimal _fallhookamount = 0M;
        private decimal _fcurrenthookqty = 0M;
        private decimal _fcurrenthookamount = 0M;
        private decimal _fstdallhookamount = 0M;
        private decimal _fstdcurrenthookamount = 0M;
        private int _fscstockid = 0;
        private int _fdcstockid = 0;
        private DateTime? _fperioddate;
        private int _fcostobjgroupid = 0;
        private int _fcostobjid = 0;
        private int _fdetailid;
        private decimal _fmaterialcostprice = 0M;
        private int _freproducetype = 0;
        private int _fbominterid = 0;
        private decimal _fdiscountrate = 0M;
        private decimal _fdiscountamount = 0M;
        private int _fsepcialsaleid = 0;
        private decimal _foutcommitqty = 0M;
        private decimal _foutseccommitqty = 0M;
        private decimal _fdbcommitqty = 0M;
        private decimal _fdbseccommitqty = 0M;
        private decimal _fauxqtyinvoice = 0M;
        private int _fopersn = 0;
        private int _fcheckstatus = 0;
        private decimal? _fsplitsecqty;
        private int _finstockid = 0;
        private decimal _fsalecommitqty = 0M;
        private decimal _fsaleseccommitqty = 0M;
        private decimal _fsaleauxcommitqty = 0M;
        private int _fselectedprocid = 0;
        private decimal _fvwinstockqty = 0M;
        private decimal _fauxvwinstockqty = 0M;
        private decimal _fsecvwinstockqty = 0M;
        private decimal _fsecinvoiceqty = 0M;
        private int _fcostcenterid = 0;
        private decimal _fentryselfb0174;
        private string _fentryselfb0158;
        private string _fentryselfb0160;
        private string _fentryselfb0161;
        private string _fentryselfb0164;
        private string _fentryselfb0165;
        private string _fentryselfb0166;
        private string _fentryselfb0167;
        private string _fentryselfb0168;
        private string _fentryselfb0169;
        private string _fentryselfb0171;
        private decimal? _fentryselfb0170;
        private decimal? _fentryselfb0162;
        private decimal? _fentryselfb0172;
        private decimal? _fentryselfb0173;
        private decimal? _fentryselfb0157;
        private decimal? _fentryselfb0163;
        private decimal? _fentryselfb0159;
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
        public int FEntryID
        {
            set { _fentryid = value; }
            get { return _fentryid; }
        }
        /// <summary>
        /// 
        /// </summary>
        public int FItemID
        {
            set { _fitemid = value; }
            get { return _fitemid; }
        }
        /// <summary>
        /// 
        /// </summary>
        public decimal FQtyMust
        {
            set { _fqtymust = value; }
            get { return _fqtymust; }
        }
        /// <summary>
        /// 
        /// </summary>
        public decimal FQty
        {
            set { _fqty = value; }
            get { return _fqty; }
        }
        /// <summary>
        /// 
        /// </summary>
        public decimal FPrice
        {
            set { _fprice = value; }
            get { return _fprice; }
        }
        /// <summary>
        /// 
        /// </summary>
        public string FBatchNo
        {
            set { _fbatchno = value; }
            get { return _fbatchno; }
        }
        /// <summary>
        /// 
        /// </summary>
        public decimal FAmount
        {
            set { _famount = value; }
            get { return _famount; }
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
        public int? FSCBillInterID
        {
            set { _fscbillinterid = value; }
            get { return _fscbillinterid; }
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
        public int FUnitID
        {
            set { _funitid = value; }
            get { return _funitid; }
        }
        /// <summary>
        /// 
        /// </summary>
        public decimal FAuxPrice
        {
            set { _fauxprice = value; }
            get { return _fauxprice; }
        }
        /// <summary>
        /// 
        /// </summary>
        public decimal FAuxQty
        {
            set { _fauxqty = value; }
            get { return _fauxqty; }
        }
        /// <summary>
        /// 
        /// </summary>
        public decimal FAuxQtyMust
        {
            set { _fauxqtymust = value; }
            get { return _fauxqtymust; }
        }
        /// <summary>
        /// 
        /// </summary>
        public decimal FQtyActual
        {
            set { _fqtyactual = value; }
            get { return _fqtyactual; }
        }
        /// <summary>
        /// 
        /// </summary>
        public decimal FAuxQtyActual
        {
            set { _fauxqtyactual = value; }
            get { return _fauxqtyactual; }
        }
        /// <summary>
        /// 
        /// </summary>
        public decimal FPlanPrice
        {
            set { _fplanprice = value; }
            get { return _fplanprice; }
        }
        /// <summary>
        /// 
        /// </summary>
        public decimal FAuxPlanPrice
        {
            set { _fauxplanprice = value; }
            get { return _fauxplanprice; }
        }
        /// <summary>
        /// 
        /// </summary>
        public int FSourceEntryID
        {
            set { _fsourceentryid = value; }
            get { return _fsourceentryid; }
        }
        /// <summary>
        /// 
        /// </summary>
        public decimal FCommitQty
        {
            set { _fcommitqty = value; }
            get { return _fcommitqty; }
        }
        /// <summary>
        /// 
        /// </summary>
        public decimal FAuxCommitQty
        {
            set { _fauxcommitqty = value; }
            get { return _fauxcommitqty; }
        }
        /// <summary>
        /// 
        /// </summary>
        public DateTime? FKFDate
        {
            set { _fkfdate = value; }
            get { return _fkfdate; }
        }
        /// <summary>
        /// 
        /// </summary>
        public int FKFPeriod
        {
            set { _fkfperiod = value; }
            get { return _fkfperiod; }
        }
        /// <summary>
        /// 
        /// </summary>
        public int? FDCSPID
        {
            set { _fdcspid = value; }
            get { return _fdcspid; }
        }
        /// <summary>
        /// 
        /// </summary>
        public int? FSCSPID
        {
            set { _fscspid = value; }
            get { return _fscspid; }
        }
        /// <summary>
        /// 
        /// </summary>
        public decimal FConsignPrice
        {
            set { _fconsignprice = value; }
            get { return _fconsignprice; }
        }
        /// <summary>
        /// 
        /// </summary>
        public decimal FConsignAmount
        {
            set { _fconsignamount = value; }
            get { return _fconsignamount; }
        }
        /// <summary>
        /// 
        /// </summary>
        public decimal FProcessCost
        {
            set { _fprocesscost = value; }
            get { return _fprocesscost; }
        }
        /// <summary>
        /// 
        /// </summary>
        public decimal FMaterialCost
        {
            set { _fmaterialcost = value; }
            get { return _fmaterialcost; }
        }
        /// <summary>
        /// 
        /// </summary>
        public decimal FTaxAmount
        {
            set { _ftaxamount = value; }
            get { return _ftaxamount; }
        }
        /// <summary>
        /// 
        /// </summary>
        public string FMapNumber
        {
            set { _fmapnumber = value; }
            get { return _fmapnumber; }
        }
        /// <summary>
        /// 
        /// </summary>
        public string FMapName
        {
            set { _fmapname = value; }
            get { return _fmapname; }
        }
        /// <summary>
        /// 
        /// </summary>
        public int FOrgBillEntryID
        {
            set { _forgbillentryid = value; }
            get { return _forgbillentryid; }
        }
        /// <summary>
        /// 
        /// </summary>
        public int? FOperID
        {
            set { _foperid = value; }
            get { return _foperid; }
        }
        /// <summary>
        /// 
        /// </summary>
        public decimal FPlanAmount
        {
            set { _fplanamount = value; }
            get { return _fplanamount; }
        }
        /// <summary>
        /// 
        /// </summary>
        public decimal FProcessPrice
        {
            set { _fprocessprice = value; }
            get { return _fprocessprice; }
        }
        /// <summary>
        /// 
        /// </summary>
        public decimal FTaxRate
        {
            set { _ftaxrate = value; }
            get { return _ftaxrate; }
        }
        /// <summary>
        /// 
        /// </summary>
        public int? FSnListID
        {
            set { _fsnlistid = value; }
            get { return _fsnlistid; }
        }
        /// <summary>
        /// 
        /// </summary>
        public decimal FAmtRef
        {
            set { _famtref = value; }
            get { return _famtref; }
        }
        /// <summary>
        /// 
        /// </summary>
        public int FAuxPropID
        {
            set { _fauxpropid = value; }
            get { return _fauxpropid; }
        }
        /// <summary>
        /// 
        /// </summary>
        public decimal FCost
        {
            set { _fcost = value; }
            get { return _fcost; }
        }
        /// <summary>
        /// 
        /// </summary>
        public decimal FPriceRef
        {
            set { _fpriceref = value; }
            get { return _fpriceref; }
        }
        /// <summary>
        /// 
        /// </summary>
        public decimal FAuxPriceRef
        {
            set { _fauxpriceref = value; }
            get { return _fauxpriceref; }
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
        public decimal FQtyInvoice
        {
            set { _fqtyinvoice = value; }
            get { return _fqtyinvoice; }
        }
        /// <summary>
        /// 
        /// </summary>
        public decimal FQtyInvoiceBase
        {
            set { _fqtyinvoicebase = value; }
            get { return _fqtyinvoicebase; }
        }
        /// <summary>
        /// 
        /// </summary>
        public decimal FUnitCost
        {
            set { _funitcost = value; }
            get { return _funitcost; }
        }
        /// <summary>
        /// 
        /// </summary>
        public decimal FSecCoefficient
        {
            set { _fseccoefficient = value; }
            get { return _fseccoefficient; }
        }
        /// <summary>
        /// 
        /// </summary>
        public decimal FSecQty
        {
            set { _fsecqty = value; }
            get { return _fsecqty; }
        }
        /// <summary>
        /// 
        /// </summary>
        public decimal FSecCommitQty
        {
            set { _fseccommitqty = value; }
            get { return _fseccommitqty; }
        }
        /// <summary>
        /// 
        /// </summary>
        public int FSourceTranType
        {
            set { _fsourcetrantype = value; }
            get { return _fsourcetrantype; }
        }
        /// <summary>
        /// 
        /// </summary>
        public int FSourceInterId
        {
            set { _fsourceinterid = value; }
            get { return _fsourceinterid; }
        }
        /// <summary>
        /// 
        /// </summary>
        public string FSourceBillNo
        {
            set { _fsourcebillno = value; }
            get { return _fsourcebillno; }
        }
        /// <summary>
        /// 
        /// </summary>
        public int FContractInterID
        {
            set { _fcontractinterid = value; }
            get { return _fcontractinterid; }
        }
        /// <summary>
        /// 
        /// </summary>
        public int FContractEntryID
        {
            set { _fcontractentryid = value; }
            get { return _fcontractentryid; }
        }
        /// <summary>
        /// 
        /// </summary>
        public string FContractBillNo
        {
            set { _fcontractbillno = value; }
            get { return _fcontractbillno; }
        }
        /// <summary>
        /// 
        /// </summary>
        public string FICMOBillNo
        {
            set { _ficmobillno = value; }
            get { return _ficmobillno; }
        }
        /// <summary>
        /// 
        /// </summary>
        public int FICMOInterID
        {
            set { _ficmointerid = value; }
            get { return _ficmointerid; }
        }
        /// <summary>
        /// 
        /// </summary>
        public int FPPBomEntryID
        {
            set { _fppbomentryid = value; }
            get { return _fppbomentryid; }
        }
        /// <summary>
        /// 
        /// </summary>
        public int FOrderInterID
        {
            set { _forderinterid = value; }
            get { return _forderinterid; }
        }
        /// <summary>
        /// 
        /// </summary>
        public int FOrderEntryID
        {
            set { _forderentryid = value; }
            get { return _forderentryid; }
        }
        /// <summary>
        /// 
        /// </summary>
        public string FOrderBillNo
        {
            set { _forderbillno = value; }
            get { return _forderbillno; }
        }
        /// <summary>
        /// 
        /// </summary>
        public decimal FAllHookQTY
        {
            set { _fallhookqty = value; }
            get { return _fallhookqty; }
        }
        /// <summary>
        /// 
        /// </summary>
        public decimal FAllHookAmount
        {
            set { _fallhookamount = value; }
            get { return _fallhookamount; }
        }
        /// <summary>
        /// 
        /// </summary>
        public decimal FCurrentHookQTY
        {
            set { _fcurrenthookqty = value; }
            get { return _fcurrenthookqty; }
        }
        /// <summary>
        /// 
        /// </summary>
        public decimal FCurrentHookAmount
        {
            set { _fcurrenthookamount = value; }
            get { return _fcurrenthookamount; }
        }
        /// <summary>
        /// 
        /// </summary>
        public decimal FStdAllHookAmount
        {
            set { _fstdallhookamount = value; }
            get { return _fstdallhookamount; }
        }
        /// <summary>
        /// 
        /// </summary>
        public decimal FStdCurrentHookAmount
        {
            set { _fstdcurrenthookamount = value; }
            get { return _fstdcurrenthookamount; }
        }
        /// <summary>
        /// 
        /// </summary>
        public int FSCStockID
        {
            set { _fscstockid = value; }
            get { return _fscstockid; }
        }
        /// <summary>
        /// 
        /// </summary>
        public int FDCStockID
        {
            set { _fdcstockid = value; }
            get { return _fdcstockid; }
        }
        /// <summary>
        /// 
        /// </summary>
        public DateTime? FPeriodDate
        {
            set { _fperioddate = value; }
            get { return _fperioddate; }
        }
        /// <summary>
        /// 
        /// </summary>
        public int FCostObjGroupID
        {
            set { _fcostobjgroupid = value; }
            get { return _fcostobjgroupid; }
        }
        /// <summary>
        /// 
        /// </summary>
        public int FCostOBJID
        {
            set { _fcostobjid = value; }
            get { return _fcostobjid; }
        }
        /// <summary>
        /// 
        /// </summary>
        public int FDetailID
        {
            set { _fdetailid = value; }
            get { return _fdetailid; }
        }
        /// <summary>
        /// 
        /// </summary>
        public decimal FMaterialCostPrice
        {
            set { _fmaterialcostprice = value; }
            get { return _fmaterialcostprice; }
        }
        /// <summary>
        /// 
        /// </summary>
        public int FReProduceType
        {
            set { _freproducetype = value; }
            get { return _freproducetype; }
        }
        /// <summary>
        /// 
        /// </summary>
        public int FBomInterID
        {
            set { _fbominterid = value; }
            get { return _fbominterid; }
        }
        /// <summary>
        /// 
        /// </summary>
        public decimal FDiscountRate
        {
            set { _fdiscountrate = value; }
            get { return _fdiscountrate; }
        }
        /// <summary>
        /// 
        /// </summary>
        public decimal FDiscountAmount
        {
            set { _fdiscountamount = value; }
            get { return _fdiscountamount; }
        }
        /// <summary>
        /// 
        /// </summary>
        public int FSepcialSaleId
        {
            set { _fsepcialsaleid = value; }
            get { return _fsepcialsaleid; }
        }
        /// <summary>
        /// 
        /// </summary>
        public decimal FOutCommitQty
        {
            set { _foutcommitqty = value; }
            get { return _foutcommitqty; }
        }
        /// <summary>
        /// 
        /// </summary>
        public decimal FOutSecCommitQty
        {
            set { _foutseccommitqty = value; }
            get { return _foutseccommitqty; }
        }
        /// <summary>
        /// 
        /// </summary>
        public decimal FDBCommitQty
        {
            set { _fdbcommitqty = value; }
            get { return _fdbcommitqty; }
        }
        /// <summary>
        /// 
        /// </summary>
        public decimal FDBSecCommitQty
        {
            set { _fdbseccommitqty = value; }
            get { return _fdbseccommitqty; }
        }
        /// <summary>
        /// 
        /// </summary>
        public decimal FAuxQtyInvoice
        {
            set { _fauxqtyinvoice = value; }
            get { return _fauxqtyinvoice; }
        }
        /// <summary>
        /// 
        /// </summary>
        public int FOperSN
        {
            set { _fopersn = value; }
            get { return _fopersn; }
        }
        /// <summary>
        /// 
        /// </summary>
        public int FCheckStatus
        {
            set { _fcheckstatus = value; }
            get { return _fcheckstatus; }
        }
        /// <summary>
        /// 
        /// </summary>
        public decimal? FSplitSecQty
        {
            set { _fsplitsecqty = value; }
            get { return _fsplitsecqty; }
        }
        /// <summary>
        /// 
        /// </summary>
        public int FInStockID
        {
            set { _finstockid = value; }
            get { return _finstockid; }
        }
        /// <summary>
        /// 
        /// </summary>
        public decimal FSaleCommitQty
        {
            set { _fsalecommitqty = value; }
            get { return _fsalecommitqty; }
        }
        /// <summary>
        /// 
        /// </summary>
        public decimal FSaleSecCommitQty
        {
            set { _fsaleseccommitqty = value; }
            get { return _fsaleseccommitqty; }
        }
        /// <summary>
        /// 
        /// </summary>
        public decimal FSaleAuxCommitQty
        {
            set { _fsaleauxcommitqty = value; }
            get { return _fsaleauxcommitqty; }
        }
        /// <summary>
        /// 
        /// </summary>
        public int FSelectedProcID
        {
            set { _fselectedprocid = value; }
            get { return _fselectedprocid; }
        }
        /// <summary>
        /// 
        /// </summary>
        public decimal FVWInStockQty
        {
            set { _fvwinstockqty = value; }
            get { return _fvwinstockqty; }
        }
        /// <summary>
        /// 
        /// </summary>
        public decimal FAuxVWInStockQty
        {
            set { _fauxvwinstockqty = value; }
            get { return _fauxvwinstockqty; }
        }
        /// <summary>
        /// 
        /// </summary>
        public decimal FSecVWInStockQty
        {
            set { _fsecvwinstockqty = value; }
            get { return _fsecvwinstockqty; }
        }
        /// <summary>
        /// 
        /// </summary>
        public decimal FSecInvoiceQty
        {
            set { _fsecinvoiceqty = value; }
            get { return _fsecinvoiceqty; }
        }
        /// <summary>
        /// 
        /// </summary>
        public int FCostCenterID
        {
            set { _fcostcenterid = value; }
            get { return _fcostcenterid; }
        }
        /// <summary>
        /// 
        /// </summary>
        public decimal FEntrySelfB0174
        {
            set { _fentryselfb0174 = value; }
            get { return _fentryselfb0174; }
        }
        /// <summary>
        /// 
        /// </summary>
        public string FEntrySelfB0158
        {
            set { _fentryselfb0158 = value; }
            get { return _fentryselfb0158; }
        }
        /// <summary>
        /// 
        /// </summary>
        public string FEntrySelfB0160
        {
            set { _fentryselfb0160 = value; }
            get { return _fentryselfb0160; }
        }
        /// <summary>
        /// 
        /// </summary>
        public string FEntrySelfB0161
        {
            set { _fentryselfb0161 = value; }
            get { return _fentryselfb0161; }
        }
        /// <summary>
        /// 
        /// </summary>
        public string FEntrySelfB0164
        {
            set { _fentryselfb0164 = value; }
            get { return _fentryselfb0164; }
        }
        /// <summary>
        /// 
        /// </summary>
        public string FEntrySelfB0165
        {
            set { _fentryselfb0165 = value; }
            get { return _fentryselfb0165; }
        }
        /// <summary>
        /// 
        /// </summary>
        public string FEntrySelfB0166
        {
            set { _fentryselfb0166 = value; }
            get { return _fentryselfb0166; }
        }
        /// <summary>
        /// 
        /// </summary>
        public string FEntrySelfB0167
        {
            set { _fentryselfb0167 = value; }
            get { return _fentryselfb0167; }
        }
        /// <summary>
        /// 
        /// </summary>
        public string FEntrySelfB0168
        {
            set { _fentryselfb0168 = value; }
            get { return _fentryselfb0168; }
        }
        /// <summary>
        /// 
        /// </summary>
        public string FEntrySelfB0169
        {
            set { _fentryselfb0169 = value; }
            get { return _fentryselfb0169; }
        }
        /// <summary>
        /// 
        /// </summary>
        public string FEntrySelfB0171
        {
            set { _fentryselfb0171 = value; }
            get { return _fentryselfb0171; }
        }
        /// <summary>
        /// 
        /// </summary>
        public decimal? FEntrySelfB0170
        {
            set { _fentryselfb0170 = value; }
            get { return _fentryselfb0170; }
        }
        /// <summary>
        /// 
        /// </summary>
        public decimal? FEntrySelfB0162
        {
            set { _fentryselfb0162 = value; }
            get { return _fentryselfb0162; }
        }
        /// <summary>
        /// 
        /// </summary>
        public decimal? FEntrySelfB0172
        {
            set { _fentryselfb0172 = value; }
            get { return _fentryselfb0172; }
        }
        /// <summary>
        /// 
        /// </summary>
        public decimal? FEntrySelfB0173
        {
            set { _fentryselfb0173 = value; }
            get { return _fentryselfb0173; }
        }
        /// <summary>
        /// 
        /// </summary>
        public decimal? FEntrySelfB0157
        {
            set { _fentryselfb0157 = value; }
            get { return _fentryselfb0157; }
        }
        /// <summary>
        /// 
        /// </summary>
        public decimal? FEntrySelfB0163
        {
            set { _fentryselfb0163 = value; }
            get { return _fentryselfb0163; }
        }
        /// <summary>
        /// 
        /// </summary>
        public decimal? FEntrySelfB0159
        {
            set { _fentryselfb0159 = value; }
            get { return _fentryselfb0159; }
        }
        #endregion Model

    }
}


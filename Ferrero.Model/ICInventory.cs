using System;
namespace Ferrero.Model
{
    /// <summary>
    /// ICInventory:实体类(属性说明自动提取数据库字段的描述信息)
    /// </summary>
    [Serializable]
    public partial class ICInventory
    {
        public ICInventory()
        { }
        #region Model
        private string _fbrno;
        private int _fitemid;
        private string _fbatchno = "";
        private int _fstockid;
        private decimal _fqty = 0M;
        private decimal _fbal = 0M;
        private int _fstockplaceid = 0;
        private int _fkfperiod = 0;
        private string _fkfdate = "";
        private decimal _fqtylock = 0M;
        private int _fauxpropid = 0;
        private decimal _fsecqty = 0M;
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
        public int FItemID
        {
            set { _fitemid = value; }
            get { return _fitemid; }
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
        public int FStockID
        {
            set { _fstockid = value; }
            get { return _fstockid; }
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
        public decimal FBal
        {
            set { _fbal = value; }
            get { return _fbal; }
        }
        /// <summary>
        /// 
        /// </summary>
        public int FStockPlaceID
        {
            set { _fstockplaceid = value; }
            get { return _fstockplaceid; }
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
        public string FKFDate
        {
            set { _fkfdate = value; }
            get { return _fkfdate; }
        }
        /// <summary>
        /// 
        /// </summary>
        public decimal FQtyLock
        {
            set { _fqtylock = value; }
            get { return _fqtylock; }
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
        public decimal FSecQty
        {
            set { _fsecqty = value; }
            get { return _fsecqty; }
        }
        #endregion Model

    }
}


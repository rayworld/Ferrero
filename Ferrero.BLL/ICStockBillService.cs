using System;
using System.Data;
using System.Collections.Generic;
using Ferrero.Model;

namespace Ferrero.BLL
{
    /// <summary>
    /// ICStockBill
    /// </summary>
    public partial class ICStockBillService
    {
        private readonly Ferrero.DAL.ICStockBillRepository dal = new Ferrero.DAL.ICStockBillRepository();
        public ICStockBillService()
        { }

        #region  BasicMethod

        /// <summary>
        /// </summary>
        public bool Add(string sConnectionString, Ferrero.Model.ICStockBill model)
        {
            return dal.Add(sConnectionString, model);
        }

        ///// <summary>
        ///// 得到一个对象实体
        ///// </summary>
        //public Ferrero.Model.ICStockBill GetModel(string sConnectionString,int FInterID, string FBillNo, int FRelateInvoiceID)
        //{

        //    return dal.GetModel(sConnectionString,FInterID, FBillNo, FRelateInvoiceID);
        //}

        ///// <summary>
        ///// 获得数据列表
        ///// </summary>
        //public DataSet GetList(string sConnectionString, string strWhere)
        //{
        //    return dal.GetList(sConnectionString, strWhere);
        //}
        ///// <summary>
        ///// 获得前几行数据
        ///// </summary>
        //public DataSet GetList(string sConnectionString,int Top, string strWhere, string filedOrder)
        //{
        //    return dal.GetList(sConnectionString,Top, strWhere, filedOrder);
        //}
        ///// <summary>
        ///// 获得数据列表
        ///// </summary>
        //public List<Ferrero.Model.ICStockBill> GetModelList(string sConnectionString,string strWhere)
        //{
        //    DataSet ds = dal.GetList(sConnectionString, strWhere);
        //    return DataTableToList(ds.Tables[0]);
        //}
        ///// <summary>
        ///// 获得数据列表
        ///// </summary>
        //public List<Ferrero.Model.ICStockBill> DataTableToList(DataTable dt)
        //{
        //    List<Ferrero.Model.ICStockBill> modelList = new List<Ferrero.Model.ICStockBill>();
        //    int rowsCount = dt.Rows.Count;
        //    if (rowsCount > 0)
        //    {
        //        Ferrero.Model.ICStockBill model;
        //        for (int n = 0; n < rowsCount; n++)
        //        {
        //            model = dal.DataRowToModel(dt.Rows[n]);
        //            if (model != null)
        //            {
        //                modelList.Add(model);
        //            }
        //        }
        //    }
        //    return modelList;
        //}

        ///// <summary>
        ///// 获得数据列表
        ///// </summary>
        //public DataSet GetAllList(string sConnectionString)
        //{
        //    return GetList(sConnectionString,"");
        //}

        ///// <summary>
        ///// 分页获取数据列表
        ///// </summary>
        //public int GetRecordCount(string sConnectionString,string strWhere)
        //{
        //    return dal.GetRecordCount(sConnectionString,strWhere);
        //}
        ///// <summary>
        ///// 分页获取数据列表
        ///// </summary>
        //public DataSet GetListByPage(string sConnectionString, string strWhere, string orderby, int startIndex, int endIndex)
        //{
        //    return dal.GetListByPage(sConnectionString,strWhere, orderby, startIndex, endIndex);
        //}
        ///// <summary>
        ///// 分页获取数据列表
        ///// </summary>
        ////public DataSet GetList(int PageSize,int PageIndex,string strWhere)
        ////{
        ////return dal.GetList(PageSize,PageIndex,strWhere);
        ////}

        #endregion  BasicMethod

        #region  ExtensionMethod

        /// <summary>
        /// 按日期和单据类型删除数据
        /// </summary>
        public int Delete(string connectionName, string minDate, string maxDate, int tranType)
        {
            return dal.Delete(connectionName, minDate, maxDate, tranType);
        }

        /// <summary>
        /// 得到当前的内联ID号
        /// </summary>
        /// <param name="connectionName"></param>
        /// <returns></returns>
        public int GetMaxFInterID(string connectionName)
        {
            return dal.GetMaxFInterID(connectionName);
        }

        /// <summary>
        /// 通过仓库名称得到仓库的编号
        /// </summary>
        /// <param name="connectionName"></param>
        /// <param name="stockName"></param>
        /// <returns></returns>
        public int GetStockId(string connectionName, string stockName)
        {
            return dal.GetStockID(connectionName, stockName);
        }

        /// <summary>
        /// 更新当前的内联ID号
        /// </summary>
        /// <param name="connectionName"></param>
        /// <returns></returns>
        public bool UpdateFInterID(string connectionName)
        {
            return dal.UpdateFInterID(connectionName);
        }

        /// <summary>
        /// 更新当前的单据号
        /// </summary>
        /// <param name="connectionName"></param>
        /// <param name="fBillNum"></param>
        /// <param name="fDesc"></param>
        /// <param name="fTranType"></param>
        /// <returns></returns>
        public bool UpdateFBillNo(string connectionName, int fBillNum, string fDesc, int fTranType)
        {
            return dal.UpdateFBillNo(connectionName, fBillNum, fDesc, fTranType);
        }


        /// <summary>
        /// 通过用户名称得到用户的编号
        /// </summary>
        /// <param name="sConnectionName"></param>
        /// <param name="sUserName"></param>
        /// <returns></returns>
        public int getUserIDByName(string sConnectionName, string sUserName)
        {
            return dal.GetUserIDByName(sConnectionName, sUserName);
        }

        #endregion  ExtensionMethod
    }
}


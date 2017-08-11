using System;
using System.Data;
using System.Collections.Generic;
using Ferrero.Model;
namespace Ferrero.BLL
{
    /// <summary>
    /// ICStockBillEntry
    /// </summary>
    public partial class ICStockBillEntryService
    {
        private readonly Ferrero.DAL.ICStockBillEntryRepository  dal = new Ferrero.DAL.ICStockBillEntryRepository();
        public ICStockBillEntryService()
        { }
        #region  BasicMethod
        /// <summary>
        /// 增加一条数据
        /// </summary>
        public bool Add(string sConnectionString, Ferrero.Model.ICStockBillEntry model)
        {
            return dal.Add(sConnectionString,model);
        }

        ///// <summary>
        ///// 得到一个对象实体
        ///// </summary>
        //public Ferrero.Model.ICStockBillEntry GetModel(string sConnectionString,int FDetailID)
        //{

        //    return dal.GetModel(sConnectionString,FDetailID);
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
        //public DataSet GetList(string  sConnectionString,int Top, string strWhere, string filedOrder)
        //{
        //    return dal.GetList(sConnectionString,Top, strWhere, filedOrder);
        //}
        ///// <summary>
        ///// 获得数据列表
        ///// </summary>
        //public List<Ferrero.Model.ICStockBillEntry> GetModelList(string sConnectionString,string strWhere)
        //{
        //    DataSet ds = dal.GetList(sConnectionString,strWhere);
        //    return DataTableToList(ds.Tables[0]);
        //}
        ///// <summary>
        ///// 获得数据列表
        ///// </summary>
        //public List<Ferrero.Model.ICStockBillEntry> DataTableToList(DataTable dt)
        //{
        //    List<Ferrero.Model.ICStockBillEntry> modelList = new List<Ferrero.Model.ICStockBillEntry>();
        //    int rowsCount = dt.Rows.Count;
        //    if (rowsCount > 0)
        //    {
        //        Ferrero.Model.ICStockBillEntry model;
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
        //    return dal.GetRecordCount(sConnectionString, strWhere);
        //}
        ///// <summary>
        ///// 分页获取数据列表
        ///// </summary>
        //public DataSet GetListByPage(string sConnectionString, string strWhere, string orderby, int startIndex, int endIndex)
        //{
        //    return dal.GetListByPage(sConnectionString, strWhere, orderby, startIndex, endIndex);
        //}
        /// <summary>
        /// 分页获取数据列表
        /// </summary>
        //public DataSet GetList(int PageSize,int PageIndex,string strWhere)
        //{
        //return dal.GetList(PageSize,PageIndex,strWhere);
        //}

        #endregion  BasicMethod

        #region  ExtensionMethod

        /// <summary>
        /// 通过商品代码得到商品编号
        /// </summary>
        /// <param name="connectionName"></param>
        /// <param name="fNumber">商品代码</param>
        /// <returns></returns>
        public int GetfItemId(string connectionName, string fNumber)
        {
            return dal.GetfItemId(connectionName, fNumber);
        }

        /// <summary>
        /// 通过单位名称得到单位编号
        /// </summary>
        /// <param name="connectionName"></param>
        /// <param name="fName"></param>
        /// <returns></returns>
        public int GetUnitID(string connectionName, string fName)
        {
            return dal.GetUnitID(connectionName, fName);
        }

        /// <summary>
        /// 通过供应商代码得到供应商编号
        /// </summary>
        /// <param name="connectionName"></param>
        /// <param name="fNumber">供应商代码</param>
        /// <returns></returns>
        public int GetSupplyID(string connectionName, string fNumber)
        {
            return dal.GetSupplyID(connectionName, fNumber);
        }

        /// <summary>
        /// 检查供应商代码和名称是否匹配
        /// </summary>
        /// <param name="connectionName"></param>
        /// <param name="fNumber"></param>
        /// <param name="fName"></param>
        /// <returns></returns>
        public bool CheckSupplyID(string connectionName, string fNumber, string fName)
        {
            return dal.CheckSupplyID(connectionName, fNumber, fName);
        }

        /// <summary>
        /// 检查商品代码和名称是否匹配
        /// </summary>
        /// <param name="connectionName"></param>
        /// <param name="fNumber"></param>
        /// <param name="fName"></param>
        /// <returns></returns>
        public bool checkProductID(string connectionName, string fNumber, string fName)
        {
            return dal.checkProductID(connectionName, fNumber, fName);
        }

        /// <summary>
        /// 读取SKU表中的大客户单价
        /// </summary>
        /// <param name="connectionName">分公司数据库</param>
        /// <param name="sLongCode">产品长代码</param>
        /// <returns></returns>
        public decimal getSpecialUnitPriceByLongCode(string connectionName, string sLongCode)
        {
            return dal.getSpecialUnitPriceByLongCode(connectionName, sLongCode);
        }

        #endregion  ExtensionMethod
    }
}


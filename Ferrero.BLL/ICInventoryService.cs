using System;
using System.Data;
using System.Collections.Generic;
using Ferrero.Model;

namespace Ferrero.BLL
{
	/// <summary>
	/// ICInventory
	/// </summary>
	public partial class ICInventoryService
	{
		private readonly Ferrero.DAL.ICInventoryRepository  dal=new Ferrero.DAL.ICInventoryRepository();
        public ICInventoryService()
		{}
		#region  BasicMethod
		/// <summary>
		/// 增加一条数据
		/// </summary>
		public bool Add(string sConnectionString,Ferrero.Model.ICInventory model)
		{
			return dal.Add(sConnectionString, model);
		}

		/// <summary>
		/// 得到一个对象实体
		/// </summary>
		public Ferrero.Model.ICInventory GetModel(string sConnectionString, string FBrNo,int FAuxPropID,int FItemID,string FBatchNo,int FStockID,int FStockPlaceID,int FKFPeriod,string FKFDate)
		{
			
			return dal.GetModel(sConnectionString,FBrNo,FAuxPropID,FItemID,FBatchNo,FStockID,FStockPlaceID,FKFPeriod,FKFDate);
		}

		/// <summary>
		/// 获得数据列表
		/// </summary>
		public DataSet GetList(string sConnectionString, string strWhere)
		{
			return dal.GetList(sConnectionString, strWhere);
		}
		/// <summary>
		/// 获得前几行数据
		/// </summary>
		public DataSet GetList(string sConnectionString, int Top,string strWhere,string filedOrder)
		{
			return dal.GetList(sConnectionString,Top,strWhere,filedOrder);
		}
		/// <summary>
		/// 获得数据列表
		/// </summary>
		public List<Ferrero.Model.ICInventory> GetModelList(string sConnectionString, string strWhere)
		{
			DataSet ds = dal.GetList(sConnectionString,strWhere);
			return DataTableToList(ds.Tables[0]);
		}
		/// <summary>
		/// 获得数据列表
		/// </summary>
		public List<Ferrero.Model.ICInventory> DataTableToList(DataTable dt)
		{
			List<Ferrero.Model.ICInventory> modelList = new List<Ferrero.Model.ICInventory>();
			int rowsCount = dt.Rows.Count;
			if (rowsCount > 0)
			{
				Ferrero.Model.ICInventory model;
				for (int n = 0; n < rowsCount; n++)
				{
					model = dal.DataRowToModel(dt.Rows[n]);
					if (model != null)
					{
						modelList.Add(model);
					}
				}
			}
			return modelList;
		}

		/// <summary>
		/// 获得数据列表
		/// </summary>
		public DataSet GetAllList(string sConnectionString)
		{
			return GetList(sConnectionString,"");
		}

		/// <summary>
		/// 分页获取数据列表
		/// </summary>
		public int GetRecordCount(string sConnectionString, string strWhere)
		{
			return dal.GetRecordCount(sConnectionString,strWhere);
		}
		/// <summary>
		/// 分页获取数据列表
		/// </summary>
		public DataSet GetListByPage(string sConnectionString, string strWhere, string orderby, int startIndex, int endIndex)
		{
			return dal.GetListByPage(sConnectionString, strWhere,  orderby,  startIndex,  endIndex);
		}
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
        /// 
        /// </summary>
        /// <param name="sConnectionString"></param>
        /// <returns></returns>
        public bool DeleteAll(string sConnectionString)
        {
            return dal.DeleteAll(sConnectionString);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sConnectionString"></param>
        /// <param name="sTableName"></param>
        /// <param name="sExcelVal"></param>
        /// <param name="sRelatColumn"></param>
        /// <param name="sValue"></param>
        /// <returns></returns>
        public object getVal(string sConnectionString, string sTableName, string sExcelVal, string sRelatColumn, string sValue)
        {
            return dal.getVal(sConnectionString, sTableName, sExcelVal, sRelatColumn, sValue);
        }
        #endregion  ExtensionMethod
	}
}
using System;
using System.Data;
using System.Text;
using System.Data.SqlClient;

using Ray.Framework.DBUtility;//Please add references

namespace Ferrero.DAL
{
	/// <summary>
	/// 数据访问类:ICInventory
	/// </summary>
	public partial class ICInventoryRepository
	{
        
		public ICInventoryRepository()
		{}
		#region  BasicMethod

		/// <summary>
		/// 增加一条数据
		/// </summary>
		public bool Add(string sConnectionName, Ferrero.Model.ICInventory model)
		{
			StringBuilder strSql=new StringBuilder();
			strSql.Append("insert into ICInventory(");
			strSql.Append("FBrNo,FItemID,FBatchNo,FStockID,FQty,FBal,FStockPlaceID,FKFPeriod,FKFDate,FQtyLock,FAuxPropID,FSecQty)");
			strSql.Append(" values (");
			strSql.Append("@FBrNo,@FItemID,@FBatchNo,@FStockID,@FQty,@FBal,@FStockPlaceID,@FKFPeriod,@FKFDate,@FQtyLock,@FAuxPropID,@FSecQty)");
			SqlParameter[] parameters = {
					new SqlParameter("@FBrNo", SqlDbType.VarChar,10),
					new SqlParameter("@FItemID", SqlDbType.Int,4),
					new SqlParameter("@FBatchNo", SqlDbType.VarChar,200),
					new SqlParameter("@FStockID", SqlDbType.Int,4),
					new SqlParameter("@FQty", SqlDbType.Decimal,13),
					new SqlParameter("@FBal", SqlDbType.Decimal,13),
					new SqlParameter("@FStockPlaceID", SqlDbType.Int,4),
					new SqlParameter("@FKFPeriod", SqlDbType.Int,4),
					new SqlParameter("@FKFDate", SqlDbType.VarChar,20),
					new SqlParameter("@FQtyLock", SqlDbType.Decimal,13),
					new SqlParameter("@FAuxPropID", SqlDbType.Int,4),
					new SqlParameter("@FSecQty", SqlDbType.Decimal,13)};
			parameters[0].Value = model.FBrNo;
			parameters[1].Value = model.FItemID;
			parameters[2].Value = model.FBatchNo;
			parameters[3].Value = model.FStockID;
			parameters[4].Value = model.FQty;
			parameters[5].Value = model.FBal;
			parameters[6].Value = model.FStockPlaceID;
			parameters[7].Value = model.FKFPeriod;
			parameters[8].Value = model.FKFDate;
			parameters[9].Value = model.FQtyLock;
			parameters[10].Value = model.FAuxPropID;
			parameters[11].Value = model.FSecQty;

            int rows = SqlHelper.ExecuteSql(SqlHelper.GetConnectionString(sConnectionName), strSql.ToString(), parameters);
			if (rows > 0)
			{
				return true;
			}
			else
			{
				return false;
			}
		}

		/// <summary>
		/// 得到一个对象实体
		/// </summary>
        public Ferrero.Model.ICInventory GetModel(string sConnectionName, string FBrNo, int FAuxPropID, int FItemID, string FBatchNo, int FStockID, int FStockPlaceID, int FKFPeriod, string FKFDate)
		{
			
			StringBuilder strSql=new StringBuilder();
			strSql.Append("select  top 1 FBrNo,FItemID,FBatchNo,FStockID,FQty,FBal,FStockPlaceID,FKFPeriod,FKFDate,FQtyLock,FAuxPropID,FSecQty from ICInventory ");
			strSql.Append(" where FBrNo=@FBrNo and FAuxPropID=@FAuxPropID and FItemID=@FItemID and FBatchNo=@FBatchNo and FStockID=@FStockID and FStockPlaceID=@FStockPlaceID and FKFPeriod=@FKFPeriod and FKFDate=@FKFDate ");
			SqlParameter[] parameters = {
					new SqlParameter("@FBrNo", SqlDbType.VarChar,10),
					new SqlParameter("@FAuxPropID", SqlDbType.Int,4),
					new SqlParameter("@FItemID", SqlDbType.Int,4),
					new SqlParameter("@FBatchNo", SqlDbType.VarChar,200),
					new SqlParameter("@FStockID", SqlDbType.Int,4),
					new SqlParameter("@FStockPlaceID", SqlDbType.Int,4),
					new SqlParameter("@FKFPeriod", SqlDbType.Int,4),
					new SqlParameter("@FKFDate", SqlDbType.VarChar,20)			};
			parameters[0].Value = FBrNo;
			parameters[1].Value = FAuxPropID;
			parameters[2].Value = FItemID;
			parameters[3].Value = FBatchNo;
			parameters[4].Value = FStockID;
			parameters[5].Value = FStockPlaceID;
			parameters[6].Value = FKFPeriod;
			parameters[7].Value = FKFDate;

			Ferrero.Model.ICInventory model=new Ferrero.Model.ICInventory();
            DataSet ds = SqlHelper.Query(SqlHelper.GetConnectionString(sConnectionName),strSql.ToString(), parameters);
			if(ds.Tables[0].Rows.Count>0)
			{
				return DataRowToModel(ds.Tables[0].Rows[0]);
			}
			else
			{
				return null;
			}
		}


		/// <summary>
		/// 得到一个对象实体
		/// </summary>
		public Ferrero.Model.ICInventory DataRowToModel(DataRow row)
		{
			Ferrero.Model.ICInventory model=new Ferrero.Model.ICInventory();
			if (row != null)
			{
				if(row["FBrNo"]!=null)
				{
					model.FBrNo=row["FBrNo"].ToString();
				}
				if(row["FItemID"]!=null && row["FItemID"].ToString()!="")
				{
					model.FItemID=int.Parse(row["FItemID"].ToString());
				}
				if(row["FBatchNo"]!=null)
				{
					model.FBatchNo=row["FBatchNo"].ToString();
				}
				if(row["FStockID"]!=null && row["FStockID"].ToString()!="")
				{
					model.FStockID=int.Parse(row["FStockID"].ToString());
				}
				if(row["FQty"]!=null && row["FQty"].ToString()!="")
				{
					model.FQty=decimal.Parse(row["FQty"].ToString());
				}
				if(row["FBal"]!=null && row["FBal"].ToString()!="")
				{
					model.FBal=decimal.Parse(row["FBal"].ToString());
				}
				if(row["FStockPlaceID"]!=null && row["FStockPlaceID"].ToString()!="")
				{
					model.FStockPlaceID=int.Parse(row["FStockPlaceID"].ToString());
				}
				if(row["FKFPeriod"]!=null && row["FKFPeriod"].ToString()!="")
				{
					model.FKFPeriod=int.Parse(row["FKFPeriod"].ToString());
				}
				if(row["FKFDate"]!=null)
				{
					model.FKFDate=row["FKFDate"].ToString();
				}
				if(row["FQtyLock"]!=null && row["FQtyLock"].ToString()!="")
				{
					model.FQtyLock=decimal.Parse(row["FQtyLock"].ToString());
				}
				if(row["FAuxPropID"]!=null && row["FAuxPropID"].ToString()!="")
				{
					model.FAuxPropID=int.Parse(row["FAuxPropID"].ToString());
				}
				if(row["FSecQty"]!=null && row["FSecQty"].ToString()!="")
				{
					model.FSecQty=decimal.Parse(row["FSecQty"].ToString());
				}
			}
			return model;
		}

		/// <summary>
		/// 获得数据列表
		/// </summary>
        public DataSet GetList(string sConnectionName, string strWhere)
		{
			StringBuilder strSql=new StringBuilder();
			strSql.Append("select FBrNo,FItemID,FBatchNo,FStockID,FQty,FBal,FStockPlaceID,FKFPeriod,FKFDate,FQtyLock,FAuxPropID,FSecQty ");
			strSql.Append(" FROM ICInventory ");
			if(strWhere.Trim()!="")
			{
				strSql.Append(" where "+strWhere);
			}
            return SqlHelper.Query(SqlHelper.GetConnectionString(sConnectionName),  strSql.ToString());
		}

		/// <summary>
		/// 获得前几行数据
		/// </summary>
        public DataSet GetList(string sConnectionName, int Top, string strWhere, string filedOrder)
		{
			StringBuilder strSql=new StringBuilder();
			strSql.Append("select ");
			if(Top>0)
			{
				strSql.Append(" top "+Top.ToString());
			}
			strSql.Append(" FBrNo,FItemID,FBatchNo,FStockID,FQty,FBal,FStockPlaceID,FKFPeriod,FKFDate,FQtyLock,FAuxPropID,FSecQty ");
			strSql.Append(" FROM ICInventory ");
			if(strWhere.Trim()!="")
			{
				strSql.Append(" where "+strWhere);
			}
			strSql.Append(" order by " + filedOrder);
            return SqlHelper.Query(SqlHelper.GetConnectionString(sConnectionName), strSql.ToString());
		}

		/// <summary>
		/// 获取记录总数
		/// </summary>
        public int GetRecordCount(string sConnectionName, string strWhere)
		{
			StringBuilder strSql=new StringBuilder();
			strSql.Append("select count(1) FROM ICInventory ");
			if(strWhere.Trim()!="")
			{
				strSql.Append(" where "+strWhere);
			}
            object obj = SqlHelper.GetSingle(SqlHelper.GetConnectionString(sConnectionName), strSql.ToString());
			return obj != null ? Convert.ToInt32(obj):0;

		}
		/// <summary>
		/// 分页获取数据列表
		/// </summary>
        public DataSet GetListByPage(string sConnectionName, string strWhere, string orderby, int startIndex, int endIndex)
		{
			StringBuilder strSql=new StringBuilder();
			strSql.Append("SELECT * FROM ( ");
			strSql.Append(" SELECT ROW_NUMBER() OVER (");
			if (!string.IsNullOrEmpty(orderby.Trim()))
			{
				strSql.Append("order by T." + orderby );
			}
			else
			{
				strSql.Append("order by T.FKFDate desc");
			}
			strSql.Append(")AS Row, T.*  from ICInventory T ");
			if (!string.IsNullOrEmpty(strWhere.Trim()))
			{
				strSql.Append(" WHERE " + strWhere);
			}
			strSql.Append(" ) TT");
			strSql.AppendFormat(" WHERE TT.Row between {0} and {1}", startIndex, endIndex);
            return SqlHelper.Query(SqlHelper.GetConnectionString(sConnectionName),  strSql.ToString());
		}

		/*
		/// <summary>
		/// 分页获取数据列表
		/// </summary>
		public DataSet GetList(int PageSize,int PageIndex,string strWhere)
		{
			SqlParameter[] parameters = {
					new SqlParameter("@tblName", SqlDbType.VarChar, 255),
					new SqlParameter("@fldName", SqlDbType.VarChar, 255),
					new SqlParameter("@PageSize", SqlDbType.Int),
					new SqlParameter("@PageIndex", SqlDbType.Int),
					new SqlParameter("@IsReCount", SqlDbType.Bit),
					new SqlParameter("@OrderType", SqlDbType.Bit),
					new SqlParameter("@strWhere", SqlDbType.VarChar,1000),
					};
			parameters[0].Value = "ICInventory";
			parameters[1].Value = "FKFDate";
			parameters[2].Value = PageSize;
			parameters[3].Value = PageIndex;
			parameters[4].Value = 0;
			parameters[5].Value = 0;
			parameters[6].Value = strWhere;	
			return SqlHelper.RunProcedure("UP_GetRecordByPage",parameters,"ds");
		}*/

		#endregion  BasicMethod
        #region  ExtensionMethod
        /// <summary>
        /// 删除所有历史数据
        /// </summary>
        public bool DeleteAll(string sConnectionName)
        {

            StringBuilder strSql = new StringBuilder();
            strSql.Append("delete from ICInventory ");

            int rows = SqlHelper.ExecuteSql(SqlHelper.GetConnectionString(sConnectionName), strSql.ToString(), null);
            if (rows > 0)
            {
                return true;
            }
            else
            {
                return false;
            }
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
        public object getVal(string sConnectionName, string sTableName, string sExcelVal, string sRelatColumn, string sValue)
        {
            StringBuilder strSql = new StringBuilder();
            strSql.Append("Select " + sValue + " FROM " + sTableName + " ");
            strSql.Append(" Where " + sRelatColumn + " = '" + sExcelVal + "'");

            object obj = SqlHelper.GetSingle(SqlHelper.GetConnectionString(sConnectionName), strSql.ToString());
            return obj;
        }
        #endregion  ExtensionMethod
	}
}
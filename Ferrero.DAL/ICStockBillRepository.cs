using System;
using System.Data;
using System.Text;
using System.Data.SqlClient;
using Ray.Framework.DBUtility;//Please add references

namespace Ferrero.DAL
{
    /// <summary>
    /// 数据访问类:ICStockBill
    /// </summary>
    public partial class ICStockBillRepository
    {
        public ICStockBillRepository()
        { }

        #region  BasicMethod

        /// <summary>
        /// 增加一条数据
        /// </summary>
        public bool Add(string sConnectionName, Ferrero.Model.ICStockBill model)
        {
            StringBuilder strSql = new StringBuilder();
            if (sConnectionName.ToLower() == "wuhan")
            {
                strSql.Append("insert into ICStockBill(");
                strSql.Append("FBrNo,FInterID,FTranType,FDate,FBillNo,FUse,FNote,FDCStockID,FSCStockID,FDeptID,FEmpID,FSupplyID,FPosterID,FCheckerID,FFManagerID,FSManagerID,FBillerID,FReturnBillInterID,FSCBillNo,FHookInterID,FVchInterID,FPosted,FCheckSelect,FCurrencyID,FSaleStyle,FAcctID,FROB,FRSCBillNo,FStatus,FUpStockWhenSave,FCancellation,FOrgBillInterID,FBillTypeID,FPOStyle,FMultiCheckLevel1,FMultiCheckLevel2,FMultiCheckLevel3,FMultiCheckLevel4,FMultiCheckLevel5,FMultiCheckLevel6,FMultiCheckDate1,FMultiCheckDate2,FMultiCheckDate3,FMultiCheckDate4,FMultiCheckDate5,FMultiCheckDate6,FCurCheckLevel,FTaskID,FResourceID,FBackFlushed,FWBInterID,FTranStatus,FZPBillInterID,FRelateBrID,FPurposeID,FUUID,FRelateInvoiceID,FImport,FSystemType,FMarketingStyle,FPayBillID,FCheckDate,FExplanation,FFetchAdd,FFetchDate,FManagerID,FRefType,FSelTranType,FChildren,FHookStatus,FActPriceVchTplID,FPlanPriceVchTplID,FProcID,FActualVchTplID,FPlanVchTplID,FBrID,FVIPCardID,FVIPScore,FHolisticDiscountRate,FPOSName,FWorkShiftId,FCussentAcctID,FZanGuCount,FPOOrdBillNo,FLSSrcInterID,FSettleDate,FManageType,FHeadSelfB0146,FHeadSelfB0147,FHeadSelfB0148,FHeadSelfB0149,FHeadSelfB0150,FHeadSelfB0151)");
                strSql.Append(" values (");
                strSql.Append("@FBrNo,@FInterID,@FTranType,@FDate,@FBillNo,@FUse,@FNote,@FDCStockID,@FSCStockID,@FDeptID,@FEmpID,@FSupplyID,@FPosterID,@FCheckerID,@FFManagerID,@FSManagerID,@FBillerID,@FReturnBillInterID,@FSCBillNo,@FHookInterID,@FVchInterID,@FPosted,@FCheckSelect,@FCurrencyID,@FSaleStyle,@FAcctID,@FROB,@FRSCBillNo,@FStatus,@FUpStockWhenSave,@FCancellation,@FOrgBillInterID,@FBillTypeID,@FPOStyle,@FMultiCheckLevel1,@FMultiCheckLevel2,@FMultiCheckLevel3,@FMultiCheckLevel4,@FMultiCheckLevel5,@FMultiCheckLevel6,@FMultiCheckDate1,@FMultiCheckDate2,@FMultiCheckDate3,@FMultiCheckDate4,@FMultiCheckDate5,@FMultiCheckDate6,@FCurCheckLevel,@FTaskID,@FResourceID,@FBackFlushed,@FWBInterID,@FTranStatus,@FZPBillInterID,@FRelateBrID,@FPurposeID,@FUUID,@FRelateInvoiceID,@FImport,@FSystemType,@FMarketingStyle,@FPayBillID,@FCheckDate,@FExplanation,@FFetchAdd,@FFetchDate,@FManagerID,@FRefType,@FSelTranType,@FChildren,@FHookStatus,@FActPriceVchTplID,@FPlanPriceVchTplID,@FProcID,@FActualVchTplID,@FPlanVchTplID,@FBrID,@FVIPCardID,@FVIPScore,@FHolisticDiscountRate,@FPOSName,@FWorkShiftId,@FCussentAcctID,@FZanGuCount,@FPOOrdBillNo,@FLSSrcInterID,@FSettleDate,@FManageType,@FHeadSelfB0146,@FHeadSelfB0147,@FHeadSelfB0148,@FHeadSelfB0149,@FHeadSelfB0150,@FHeadSelfB0151)");

                SqlParameter[] parameters = {
					new SqlParameter("@FBrNo", SqlDbType.VarChar,10),
					new SqlParameter("@FInterID", SqlDbType.Int,4),
					new SqlParameter("@FTranType", SqlDbType.SmallInt,2),
					new SqlParameter("@FDate", SqlDbType.DateTime),
					new SqlParameter("@FBillNo", SqlDbType.NVarChar,255),
					new SqlParameter("@FUse", SqlDbType.VarChar,255),
					new SqlParameter("@FNote", SqlDbType.VarChar,255),
					new SqlParameter("@FDCStockID", SqlDbType.Int,4),
					new SqlParameter("@FSCStockID", SqlDbType.Int,4),
					new SqlParameter("@FDeptID", SqlDbType.Int,4),
					new SqlParameter("@FEmpID", SqlDbType.Int,4),
					new SqlParameter("@FSupplyID", SqlDbType.Int,4),
					new SqlParameter("@FPosterID", SqlDbType.Int,4),
					new SqlParameter("@FCheckerID", SqlDbType.Int,4),
					new SqlParameter("@FFManagerID", SqlDbType.Int,4),
					new SqlParameter("@FSManagerID", SqlDbType.Int,4),
					new SqlParameter("@FBillerID", SqlDbType.Int,4),
					new SqlParameter("@FReturnBillInterID", SqlDbType.Int,4),
					new SqlParameter("@FSCBillNo", SqlDbType.VarChar,30),
					new SqlParameter("@FHookInterID", SqlDbType.Int,4),
					new SqlParameter("@FVchInterID", SqlDbType.Int,4),
					new SqlParameter("@FPosted", SqlDbType.SmallInt,2),
					new SqlParameter("@FCheckSelect", SqlDbType.SmallInt,2),
					new SqlParameter("@FCurrencyID", SqlDbType.Int,4),
					new SqlParameter("@FSaleStyle", SqlDbType.Int,4),
					new SqlParameter("@FAcctID", SqlDbType.Int,4),
					new SqlParameter("@FROB", SqlDbType.SmallInt,2),
					new SqlParameter("@FRSCBillNo", SqlDbType.VarChar,30),
					new SqlParameter("@FStatus", SqlDbType.SmallInt,2),
					new SqlParameter("@FUpStockWhenSave", SqlDbType.Bit,1),
					new SqlParameter("@FCancellation", SqlDbType.Bit,1),
					new SqlParameter("@FOrgBillInterID", SqlDbType.Int,4),
					new SqlParameter("@FBillTypeID", SqlDbType.Int,4),
					new SqlParameter("@FPOStyle", SqlDbType.Int,4),
					new SqlParameter("@FMultiCheckLevel1", SqlDbType.Int,4),
					new SqlParameter("@FMultiCheckLevel2", SqlDbType.Int,4),
					new SqlParameter("@FMultiCheckLevel3", SqlDbType.Int,4),
					new SqlParameter("@FMultiCheckLevel4", SqlDbType.Int,4),
					new SqlParameter("@FMultiCheckLevel5", SqlDbType.Int,4),
					new SqlParameter("@FMultiCheckLevel6", SqlDbType.Int,4),
					new SqlParameter("@FMultiCheckDate1", SqlDbType.DateTime),
					new SqlParameter("@FMultiCheckDate2", SqlDbType.DateTime),
					new SqlParameter("@FMultiCheckDate3", SqlDbType.DateTime),
					new SqlParameter("@FMultiCheckDate4", SqlDbType.DateTime),
					new SqlParameter("@FMultiCheckDate5", SqlDbType.DateTime),
					new SqlParameter("@FMultiCheckDate6", SqlDbType.DateTime),
					new SqlParameter("@FCurCheckLevel", SqlDbType.Int,4),
					new SqlParameter("@FTaskID", SqlDbType.Int,4),
					new SqlParameter("@FResourceID", SqlDbType.Int,4),
					new SqlParameter("@FBackFlushed", SqlDbType.Bit,1),
					new SqlParameter("@FWBInterID", SqlDbType.Int,4),
					new SqlParameter("@FTranStatus", SqlDbType.Int,4),
					new SqlParameter("@FZPBillInterID", SqlDbType.Int,4),
					new SqlParameter("@FRelateBrID", SqlDbType.Int,4),
					new SqlParameter("@FPurposeID", SqlDbType.Int,4),
					new SqlParameter("@FUUID", SqlDbType.UniqueIdentifier,16),
					new SqlParameter("@FRelateInvoiceID", SqlDbType.Int,4),
					///new SqlParameter("@FOperDate", SqlDbType.Timestamp,8),
					new SqlParameter("@FImport", SqlDbType.Int,4),
					new SqlParameter("@FSystemType", SqlDbType.Int,4),
					new SqlParameter("@FMarketingStyle", SqlDbType.Int,4),
					new SqlParameter("@FPayBillID", SqlDbType.Int,4),
					new SqlParameter("@FCheckDate", SqlDbType.DateTime),
					new SqlParameter("@FExplanation", SqlDbType.NVarChar,255),
					new SqlParameter("@FFetchAdd", SqlDbType.NVarChar,255),
					new SqlParameter("@FFetchDate", SqlDbType.DateTime),
					new SqlParameter("@FManagerID", SqlDbType.Int,4),
					new SqlParameter("@FRefType", SqlDbType.Int,4),
					new SqlParameter("@FSelTranType", SqlDbType.Int,4),
					new SqlParameter("@FChildren", SqlDbType.Int,4),
					new SqlParameter("@FHookStatus", SqlDbType.SmallInt,2),
					new SqlParameter("@FActPriceVchTplID", SqlDbType.Int,4),
					new SqlParameter("@FPlanPriceVchTplID", SqlDbType.Int,4),
					new SqlParameter("@FProcID", SqlDbType.Int,4),
					new SqlParameter("@FActualVchTplID", SqlDbType.Int,4),
					new SqlParameter("@FPlanVchTplID", SqlDbType.Int,4),
					new SqlParameter("@FBrID", SqlDbType.Int,4),
					new SqlParameter("@FVIPCardID", SqlDbType.Int,4),
					new SqlParameter("@FVIPScore", SqlDbType.Decimal,13),
					new SqlParameter("@FHolisticDiscountRate", SqlDbType.Decimal,13),
					new SqlParameter("@FPOSName", SqlDbType.VarChar,255),
					new SqlParameter("@FWorkShiftId", SqlDbType.Int,4),
					new SqlParameter("@FCussentAcctID", SqlDbType.Int,4),
					new SqlParameter("@FZanGuCount", SqlDbType.Bit,1),
					new SqlParameter("@FPOOrdBillNo", SqlDbType.NVarChar,255),
					new SqlParameter("@FLSSrcInterID", SqlDbType.Int,4),
					new SqlParameter("@FSettleDate", SqlDbType.DateTime),
					new SqlParameter("@FManageType", SqlDbType.Int,4),
					new SqlParameter("@FHeadSelfB0146", SqlDbType.VarChar,255),
					new SqlParameter("@FHeadSelfB0147", SqlDbType.VarChar,255),
					new SqlParameter("@FHeadSelfB0148", SqlDbType.VarChar,255),
					new SqlParameter("@FHeadSelfB0149", SqlDbType.VarChar,255),
					new SqlParameter("@FHeadSelfB0150", SqlDbType.VarChar,255),
					new SqlParameter("@FHeadSelfB0151", SqlDbType.VarChar,255)};
                parameters[0].Value = "0";
                parameters[1].Value = model.FInterID;
                parameters[2].Value = model.FTranType;
                parameters[3].Value = model.FDate;
                parameters[4].Value = model.FBillNo;
                parameters[5].Value = DBNull.Value;
                parameters[6].Value = DBNull.Value;
                parameters[7].Value = DBNull.Value;
                parameters[8].Value = DBNull.Value;
                parameters[9].Value = 0;
                parameters[10].Value = 0;
                parameters[11].Value = model.FSupplyID;
                parameters[12].Value = DBNull.Value;
                parameters[13].Value = DBNull.Value;
                parameters[14].Value = model.FFManagerID;
                parameters[15].Value = model.FSManagerID;
                parameters[16].Value = model.FBillerID;
                parameters[17].Value = DBNull.Value;
                parameters[18].Value = DBNull.Value;
                parameters[19].Value = 0;
                parameters[20].Value = 0;
                parameters[21].Value = 0;
                parameters[22].Value = 0;
                parameters[23].Value = DBNull.Value;

                parameters[25].Value = DBNull.Value;
                parameters[26].Value = model.FROB;
                parameters[27].Value = DBNull.Value;
                parameters[28].Value = 0;
                parameters[29].Value = false;
                parameters[30].Value = false;
                parameters[31].Value = 0;
                parameters[32].Value = DBNull.Value;

                parameters[34].Value = DBNull.Value;
                parameters[35].Value = DBNull.Value;
                parameters[36].Value = DBNull.Value;
                parameters[37].Value = DBNull.Value;
                parameters[38].Value = DBNull.Value;
                parameters[39].Value = DBNull.Value;
                parameters[40].Value = DBNull.Value;
                parameters[41].Value = DBNull.Value;
                parameters[42].Value = DBNull.Value;
                parameters[43].Value = DBNull.Value;
                parameters[44].Value = DBNull.Value;
                parameters[45].Value = DBNull.Value;
                parameters[46].Value = DBNull.Value;
                parameters[47].Value = DBNull.Value;
                parameters[48].Value = DBNull.Value;
                parameters[49].Value = false;
                parameters[50].Value = 0;
                parameters[51].Value = 0;
                parameters[52].Value = DBNull.Value;
                parameters[53].Value = 0;
                parameters[54].Value = 0;
                parameters[55].Value = Guid.NewGuid();
                parameters[56].Value = 0;
                parameters[57].Value = 0;
                parameters[58].Value = 0;
                parameters[59].Value = 12530;
                parameters[60].Value = 0;
                parameters[61].Value = DBNull.Value;
                parameters[62].Value = "";
                parameters[63].Value = "";
                parameters[64].Value = DBNull.Value;
                parameters[65].Value = 0;
                parameters[66].Value = 0;
                parameters[67].Value = 0;
                parameters[68].Value = 0;
                parameters[69].Value = 0;
                parameters[70].Value = 0;
                parameters[71].Value = 0;
                parameters[72].Value = 0;
                parameters[73].Value = 0;
                parameters[74].Value = 0;
                parameters[75].Value = 0;
                parameters[76].Value = 0;
                parameters[77].Value = 0;
                parameters[78].Value = 0;
                parameters[79].Value = "";
                parameters[80].Value = 0;
                parameters[81].Value = 0;
                parameters[82].Value = false;
                parameters[83].Value = "";
                parameters[84].Value = 0;
                parameters[85].Value = model.FDate;
                parameters[86].Value = 0;
                if (model.FTranType == 21)
                {
                    parameters[24].Value = 101;
                    parameters[33].Value = DBNull.Value;//POStyle
                    parameters[87].Value = model.FHeadSelfB0146;
                    parameters[88].Value = model.FHeadSelfB0147;
                    parameters[89].Value = model.FHeadSelfB0148;
                    parameters[90].Value = model.FHeadSelfB0149;
                    parameters[91].Value = model.FHeadSelfB0150;
                    parameters[92].Value = model.FHeadSelfB0151;
                }
                else
                {
                    parameters[24].Value = DBNull.Value;
                    parameters[33].Value = 252;//POStyle
                    parameters[87].Value = DBNull.Value;
                    parameters[88].Value = DBNull.Value;
                    parameters[89].Value = DBNull.Value;
                    parameters[90].Value = DBNull.Value;
                    parameters[91].Value = DBNull.Value;
                    parameters[92].Value = DBNull.Value;
                }

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
            else
            {
                strSql.Append("insert into ICStockBill(");
                strSql.Append("FBrNo,FInterID,FTranType,FDate,FBillNo,FUse,FNote,FDCStockID,FSCStockID,FDeptID,FEmpID,FSupplyID,FPosterID,FCheckerID,FFManagerID,FSManagerID,FBillerID,FReturnBillInterID,FSCBillNo,FHookInterID,FVchInterID,FPosted,FCheckSelect,FCurrencyID,FSaleStyle,FAcctID,FROB,FRSCBillNo,FStatus,FUpStockWhenSave,FCancellation,FOrgBillInterID,FBillTypeID,FPOStyle,FMultiCheckLevel1,FMultiCheckLevel2,FMultiCheckLevel3,FMultiCheckLevel4,FMultiCheckLevel5,FMultiCheckLevel6,FMultiCheckDate1,FMultiCheckDate2,FMultiCheckDate3,FMultiCheckDate4,FMultiCheckDate5,FMultiCheckDate6,FCurCheckLevel,FTaskID,FResourceID,FBackFlushed,FWBInterID,FTranStatus,FZPBillInterID,FRelateBrID,FPurposeID,FUUID,FRelateInvoiceID,FImport,FSystemType,FMarketingStyle,FPayBillID,FCheckDate,FExplanation,FFetchAdd,FFetchDate,FManagerID,FRefType,FSelTranType,FChildren,FHookStatus,FActPriceVchTplID,FPlanPriceVchTplID,FProcID,FActualVchTplID,FPlanVchTplID,FBrID,FVIPCardID,FVIPScore,FHolisticDiscountRate,FPOSName,FWorkShiftId,FCussentAcctID,FZanGuCount,FPOOrdBillNo,FLSSrcInterID,FSettleDate,FManageType,FHeadSelfB0146)");
                strSql.Append(" values (");
                strSql.Append("@FBrNo,@FInterID,@FTranType,@FDate,@FBillNo,@FUse,@FNote,@FDCStockID,@FSCStockID,@FDeptID,@FEmpID,@FSupplyID,@FPosterID,@FCheckerID,@FFManagerID,@FSManagerID,@FBillerID,@FReturnBillInterID,@FSCBillNo,@FHookInterID,@FVchInterID,@FPosted,@FCheckSelect,@FCurrencyID,@FSaleStyle,@FAcctID,@FROB,@FRSCBillNo,@FStatus,@FUpStockWhenSave,@FCancellation,@FOrgBillInterID,@FBillTypeID,@FPOStyle,@FMultiCheckLevel1,@FMultiCheckLevel2,@FMultiCheckLevel3,@FMultiCheckLevel4,@FMultiCheckLevel5,@FMultiCheckLevel6,@FMultiCheckDate1,@FMultiCheckDate2,@FMultiCheckDate3,@FMultiCheckDate4,@FMultiCheckDate5,@FMultiCheckDate6,@FCurCheckLevel,@FTaskID,@FResourceID,@FBackFlushed,@FWBInterID,@FTranStatus,@FZPBillInterID,@FRelateBrID,@FPurposeID,@FUUID,@FRelateInvoiceID,@FImport,@FSystemType,@FMarketingStyle,@FPayBillID,@FCheckDate,@FExplanation,@FFetchAdd,@FFetchDate,@FManagerID,@FRefType,@FSelTranType,@FChildren,@FHookStatus,@FActPriceVchTplID,@FPlanPriceVchTplID,@FProcID,@FActualVchTplID,@FPlanVchTplID,@FBrID,@FVIPCardID,@FVIPScore,@FHolisticDiscountRate,@FPOSName,@FWorkShiftId,@FCussentAcctID,@FZanGuCount,@FPOOrdBillNo,@FLSSrcInterID,@FSettleDate,@FManageType,@FHeadSelfB0146)");

                SqlParameter[] parameters = {
					new SqlParameter("@FBrNo", SqlDbType.VarChar,10),
					new SqlParameter("@FInterID", SqlDbType.Int,4),
					new SqlParameter("@FTranType", SqlDbType.SmallInt,2),
					new SqlParameter("@FDate", SqlDbType.DateTime),
					new SqlParameter("@FBillNo", SqlDbType.NVarChar,255),
					new SqlParameter("@FUse", SqlDbType.VarChar,255),
					new SqlParameter("@FNote", SqlDbType.VarChar,255),
					new SqlParameter("@FDCStockID", SqlDbType.Int,4),
					new SqlParameter("@FSCStockID", SqlDbType.Int,4),
					new SqlParameter("@FDeptID", SqlDbType.Int,4),
					new SqlParameter("@FEmpID", SqlDbType.Int,4),
					new SqlParameter("@FSupplyID", SqlDbType.Int,4),
					new SqlParameter("@FPosterID", SqlDbType.Int,4),
					new SqlParameter("@FCheckerID", SqlDbType.Int,4),
					new SqlParameter("@FFManagerID", SqlDbType.Int,4),
					new SqlParameter("@FSManagerID", SqlDbType.Int,4),
					new SqlParameter("@FBillerID", SqlDbType.Int,4),
					new SqlParameter("@FReturnBillInterID", SqlDbType.Int,4),
					new SqlParameter("@FSCBillNo", SqlDbType.VarChar,30),
					new SqlParameter("@FHookInterID", SqlDbType.Int,4),
					new SqlParameter("@FVchInterID", SqlDbType.Int,4),
					new SqlParameter("@FPosted", SqlDbType.SmallInt,2),
					new SqlParameter("@FCheckSelect", SqlDbType.SmallInt,2),
					new SqlParameter("@FCurrencyID", SqlDbType.Int,4),
					new SqlParameter("@FSaleStyle", SqlDbType.Int,4),
					new SqlParameter("@FAcctID", SqlDbType.Int,4),
					new SqlParameter("@FROB", SqlDbType.SmallInt,2),
					new SqlParameter("@FRSCBillNo", SqlDbType.VarChar,30),
					new SqlParameter("@FStatus", SqlDbType.SmallInt,2),
					new SqlParameter("@FUpStockWhenSave", SqlDbType.Bit,1),
					new SqlParameter("@FCancellation", SqlDbType.Bit,1),
					new SqlParameter("@FOrgBillInterID", SqlDbType.Int,4),
					new SqlParameter("@FBillTypeID", SqlDbType.Int,4),
					new SqlParameter("@FPOStyle", SqlDbType.Int,4),
					new SqlParameter("@FMultiCheckLevel1", SqlDbType.Int,4),
					new SqlParameter("@FMultiCheckLevel2", SqlDbType.Int,4),
					new SqlParameter("@FMultiCheckLevel3", SqlDbType.Int,4),
					new SqlParameter("@FMultiCheckLevel4", SqlDbType.Int,4),
					new SqlParameter("@FMultiCheckLevel5", SqlDbType.Int,4),
					new SqlParameter("@FMultiCheckLevel6", SqlDbType.Int,4),
					new SqlParameter("@FMultiCheckDate1", SqlDbType.DateTime),
					new SqlParameter("@FMultiCheckDate2", SqlDbType.DateTime),
					new SqlParameter("@FMultiCheckDate3", SqlDbType.DateTime),
					new SqlParameter("@FMultiCheckDate4", SqlDbType.DateTime),
					new SqlParameter("@FMultiCheckDate5", SqlDbType.DateTime),
					new SqlParameter("@FMultiCheckDate6", SqlDbType.DateTime),
					new SqlParameter("@FCurCheckLevel", SqlDbType.Int,4),
					new SqlParameter("@FTaskID", SqlDbType.Int,4),
					new SqlParameter("@FResourceID", SqlDbType.Int,4),
					new SqlParameter("@FBackFlushed", SqlDbType.Bit,1),
					new SqlParameter("@FWBInterID", SqlDbType.Int,4),
					new SqlParameter("@FTranStatus", SqlDbType.Int,4),
					new SqlParameter("@FZPBillInterID", SqlDbType.Int,4),
					new SqlParameter("@FRelateBrID", SqlDbType.Int,4),
					new SqlParameter("@FPurposeID", SqlDbType.Int,4),
					new SqlParameter("@FUUID", SqlDbType.UniqueIdentifier,16),
					new SqlParameter("@FRelateInvoiceID", SqlDbType.Int,4),
					///new SqlParameter("@FOperDate", SqlDbType.Timestamp,8),
					new SqlParameter("@FImport", SqlDbType.Int,4),
					new SqlParameter("@FSystemType", SqlDbType.Int,4),
					new SqlParameter("@FMarketingStyle", SqlDbType.Int,4),
					new SqlParameter("@FPayBillID", SqlDbType.Int,4),
					new SqlParameter("@FCheckDate", SqlDbType.DateTime),
					new SqlParameter("@FExplanation", SqlDbType.NVarChar,255),
					new SqlParameter("@FFetchAdd", SqlDbType.NVarChar,255),
					new SqlParameter("@FFetchDate", SqlDbType.DateTime),
					new SqlParameter("@FManagerID", SqlDbType.Int,4),
					new SqlParameter("@FRefType", SqlDbType.Int,4),
					new SqlParameter("@FSelTranType", SqlDbType.Int,4),
					new SqlParameter("@FChildren", SqlDbType.Int,4),
					new SqlParameter("@FHookStatus", SqlDbType.SmallInt,2),
					new SqlParameter("@FActPriceVchTplID", SqlDbType.Int,4),
					new SqlParameter("@FPlanPriceVchTplID", SqlDbType.Int,4),
					new SqlParameter("@FProcID", SqlDbType.Int,4),
					new SqlParameter("@FActualVchTplID", SqlDbType.Int,4),
					new SqlParameter("@FPlanVchTplID", SqlDbType.Int,4),
					new SqlParameter("@FBrID", SqlDbType.Int,4),
					new SqlParameter("@FVIPCardID", SqlDbType.Int,4),
					new SqlParameter("@FVIPScore", SqlDbType.Decimal,13),
					new SqlParameter("@FHolisticDiscountRate", SqlDbType.Decimal,13),
					new SqlParameter("@FPOSName", SqlDbType.VarChar,255),
					new SqlParameter("@FWorkShiftId", SqlDbType.Int,4),
					new SqlParameter("@FCussentAcctID", SqlDbType.Int,4),
					new SqlParameter("@FZanGuCount", SqlDbType.Bit,1),
					new SqlParameter("@FPOOrdBillNo", SqlDbType.NVarChar,255),
					new SqlParameter("@FLSSrcInterID", SqlDbType.Int,4),
					new SqlParameter("@FSettleDate", SqlDbType.DateTime),
					new SqlParameter("@FManageType", SqlDbType.Int,4),
					new SqlParameter("@FHeadSelfB0146", SqlDbType.VarChar,255)};
                parameters[0].Value = "0";
                parameters[1].Value = model.FInterID;
                parameters[2].Value = model.FTranType;
                parameters[3].Value = model.FDate;
                parameters[4].Value = model.FBillNo;
                parameters[5].Value = DBNull.Value;
                parameters[6].Value = DBNull.Value;
                parameters[7].Value = DBNull.Value;
                parameters[8].Value = DBNull.Value;
                parameters[9].Value = 0;
                parameters[10].Value = 0;
                parameters[11].Value = model.FSupplyID;
                parameters[12].Value = DBNull.Value;
                parameters[13].Value = DBNull.Value;
                parameters[14].Value = model.FFManagerID;
                parameters[15].Value = model.FSManagerID;
                parameters[16].Value = model.FBillerID;
                parameters[17].Value = DBNull.Value;
                parameters[18].Value = DBNull.Value;
                parameters[19].Value = 0;
                parameters[20].Value = 0;
                parameters[21].Value = 0;
                parameters[22].Value = 0;
                parameters[23].Value = DBNull.Value;

                parameters[25].Value = DBNull.Value;
                parameters[26].Value = model.FROB;
                parameters[27].Value = DBNull.Value;
                parameters[28].Value = 0;
                parameters[29].Value = false;
                parameters[30].Value = false;
                parameters[31].Value = 0;
                parameters[32].Value = DBNull.Value;

                parameters[34].Value = DBNull.Value;
                parameters[35].Value = DBNull.Value;
                parameters[36].Value = DBNull.Value;
                parameters[37].Value = DBNull.Value;
                parameters[38].Value = DBNull.Value;
                parameters[39].Value = DBNull.Value;
                parameters[40].Value = DBNull.Value;
                parameters[41].Value = DBNull.Value;
                parameters[42].Value = DBNull.Value;
                parameters[43].Value = DBNull.Value;
                parameters[44].Value = DBNull.Value;
                parameters[45].Value = DBNull.Value;
                parameters[46].Value = DBNull.Value;
                parameters[47].Value = DBNull.Value;
                parameters[48].Value = DBNull.Value;
                parameters[49].Value = false;
                parameters[50].Value = 0;
                parameters[51].Value = 0;
                parameters[52].Value = DBNull.Value;
                parameters[53].Value = 0;
                parameters[54].Value = 0;
                parameters[55].Value = Guid.NewGuid();
                parameters[56].Value = 0;
                parameters[57].Value = 0;
                parameters[58].Value = 0;
                parameters[59].Value = 12530;
                parameters[60].Value = 0;
                parameters[61].Value = DBNull.Value;
                parameters[62].Value = "";
                parameters[63].Value = "";
                parameters[64].Value = DBNull.Value;
                parameters[65].Value = 0;
                parameters[66].Value = 0;
                parameters[67].Value = 0;
                parameters[68].Value = 0;
                parameters[69].Value = 0;
                parameters[70].Value = 0;
                parameters[71].Value = 0;
                parameters[72].Value = 0;
                parameters[73].Value = 0;
                parameters[74].Value = 0;
                parameters[75].Value = 0;
                parameters[76].Value = 0;
                parameters[77].Value = 0;
                parameters[78].Value = 0;
                parameters[79].Value = "";
                parameters[80].Value = 0;
                parameters[81].Value = 0;
                parameters[82].Value = false;
                parameters[83].Value = "";
                parameters[84].Value = 0;
                parameters[85].Value = model.FDate;
                parameters[86].Value = 0;
                if (model.FTranType == 21)
                {
                    parameters[24].Value = 101;
                    parameters[33].Value = DBNull.Value;//POStyle
                    parameters[87].Value = model.FHeadSelfB0146;

                }
                else
                {
                    parameters[24].Value = DBNull.Value;
                    parameters[33].Value = 252;//POStyle
                    parameters[87].Value = DBNull.Value;
                }

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
        }

        ///// <summary>
        ///// 得到一个对象实体
        ///// </summary>
        //public Ferrero.Model.ICStockBill GetModel(string sConnectionName, int FInterID, string FBillNo, int FRelateInvoiceID)
        //{

        //    StringBuilder strSql = new StringBuilder();
        //    strSql.Append("select  top 1 FBrNo,FInterID,FTranType,FDate,FBillNo,FUse,FNote,FDCStockID,FSCStockID,FDeptID,FEmpID,FSupplyID,FPosterID,FCheckerID,FFManagerID,FSManagerID,FBillerID,FReturnBillInterID,FSCBillNo,FHookInterID,FVchInterID,FPosted,FCheckSelect,FCurrencyID,FSaleStyle,FAcctID,FROB,FRSCBillNo,FStatus,FUpStockWhenSave,FCancellation,FOrgBillInterID,FBillTypeID,FPOStyle,FMultiCheckLevel1,FMultiCheckLevel2,FMultiCheckLevel3,FMultiCheckLevel4,FMultiCheckLevel5,FMultiCheckLevel6,FMultiCheckDate1,FMultiCheckDate2,FMultiCheckDate3,FMultiCheckDate4,FMultiCheckDate5,FMultiCheckDate6,FCurCheckLevel,FTaskID,FResourceID,FBackFlushed,FWBInterID,FTranStatus,FZPBillInterID,FRelateBrID,FPurposeID,FUUID,FRelateInvoiceID,FImport,FSystemType,FMarketingStyle,FPayBillID,FCheckDate,FExplanation,FFetchAdd,FFetchDate,FManagerID,FRefType,FSelTranType,FChildren,FHookStatus,FActPriceVchTplID,FPlanPriceVchTplID,FProcID,FActualVchTplID,FPlanVchTplID,FBrID,FVIPCardID,FVIPScore,FHolisticDiscountRate,FPOSName,FWorkShiftId,FCussentAcctID,FZanGuCount,FPOOrdBillNo,FLSSrcInterID,FSettleDate,FManageType,FHeadSelfB0146,FHeadSelfB0147,FHeadSelfB0148,FHeadSelfB0149,FHeadSelfB0150,FHeadSelfB0151 from ICStockBill ");
        //    strSql.Append(" where FInterID=@FInterID and FBillNo=@FBillNo and FRelateInvoiceID=@FRelateInvoiceID ");
        //    SqlParameter[] parameters = {
        //            new SqlParameter("@FInterID", SqlDbType.Int,4),
        //            new SqlParameter("@FBillNo", SqlDbType.NVarChar,255),
        //            new SqlParameter("@FRelateInvoiceID", SqlDbType.Int,4)			};
        //    parameters[0].Value = FInterID;
        //    parameters[1].Value = FBillNo;
        //    parameters[2].Value = FRelateInvoiceID;

        //    Ferrero.Model.ICStockBill model = new Ferrero.Model.ICStockBill();
        //    DataSet ds = SqlHelper.ExecuteDataSet(SqlHelper.GetConnectionString(sConnectionName), strSql.ToString(), parameters);
        //    if (ds.Tables[0].Rows.Count > 0)
        //    {
        //        return DataRowToModel(ds.Tables[0].Rows[0]);
        //    }
        //    else
        //    {
        //        return null;
        //    }
        //}


        ///// <summary>
        ///// 得到一个对象实体
        ///// </summary>
        //public Ferrero.Model.ICStockBill DataRowToModel(DataRow row)
        //{
        //    Ferrero.Model.ICStockBill model = new Ferrero.Model.ICStockBill();
        //    if (row != null)
        //    {
        //        if (row["FBrNo"] != null)
        //        {
        //            model.FBrNo = row["FBrNo"].ToString();
        //        }
        //        if (row["FInterID"] != null && row["FInterID"].ToString() != "")
        //        {
        //            model.FInterID = int.Parse(row["FInterID"].ToString());
        //        }
        //        if (row["FTranType"] != null && row["FTranType"].ToString() != "")
        //        {
        //            model.FTranType = int.Parse(row["FTranType"].ToString());
        //        }
        //        if (row["FDate"] != null && row["FDate"].ToString() != "")
        //        {
        //            model.FDate = DateTime.Parse(row["FDate"].ToString());
        //        }
        //        if (row["FBillNo"] != null)
        //        {
        //            model.FBillNo = row["FBillNo"].ToString();
        //        }
        //        if (row["FUse"] != null)
        //        {
        //            model.FUse = row["FUse"].ToString();
        //        }
        //        if (row["FNote"] != null)
        //        {
        //            model.FNote = row["FNote"].ToString();
        //        }
        //        if (row["FDCStockID"] != null && row["FDCStockID"].ToString() != "")
        //        {
        //            model.FDCStockID = int.Parse(row["FDCStockID"].ToString());
        //        }
        //        if (row["FSCStockID"] != null && row["FSCStockID"].ToString() != "")
        //        {
        //            model.FSCStockID = int.Parse(row["FSCStockID"].ToString());
        //        }
        //        if (row["FDeptID"] != null && row["FDeptID"].ToString() != "")
        //        {
        //            model.FDeptID = int.Parse(row["FDeptID"].ToString());
        //        }
        //        if (row["FEmpID"] != null && row["FEmpID"].ToString() != "")
        //        {
        //            model.FEmpID = int.Parse(row["FEmpID"].ToString());
        //        }
        //        if (row["FSupplyID"] != null && row["FSupplyID"].ToString() != "")
        //        {
        //            model.FSupplyID = int.Parse(row["FSupplyID"].ToString());
        //        }
        //        if (row["FPosterID"] != null && row["FPosterID"].ToString() != "")
        //        {
        //            model.FPosterID = int.Parse(row["FPosterID"].ToString());
        //        }
        //        if (row["FCheckerID"] != null && row["FCheckerID"].ToString() != "")
        //        {
        //            model.FCheckerID = int.Parse(row["FCheckerID"].ToString());
        //        }
        //        if (row["FFManagerID"] != null && row["FFManagerID"].ToString() != "")
        //        {
        //            model.FFManagerID = int.Parse(row["FFManagerID"].ToString());
        //        }
        //        if (row["FSManagerID"] != null && row["FSManagerID"].ToString() != "")
        //        {
        //            model.FSManagerID = int.Parse(row["FSManagerID"].ToString());
        //        }
        //        if (row["FBillerID"] != null && row["FBillerID"].ToString() != "")
        //        {
        //            model.FBillerID = int.Parse(row["FBillerID"].ToString());
        //        }
        //        if (row["FReturnBillInterID"] != null && row["FReturnBillInterID"].ToString() != "")
        //        {
        //            model.FReturnBillInterID = int.Parse(row["FReturnBillInterID"].ToString());
        //        }
        //        if (row["FSCBillNo"] != null)
        //        {
        //            model.FSCBillNo = row["FSCBillNo"].ToString();
        //        }
        //        if (row["FHookInterID"] != null && row["FHookInterID"].ToString() != "")
        //        {
        //            model.FHookInterID = int.Parse(row["FHookInterID"].ToString());
        //        }
        //        if (row["FVchInterID"] != null && row["FVchInterID"].ToString() != "")
        //        {
        //            model.FVchInterID = int.Parse(row["FVchInterID"].ToString());
        //        }
        //        if (row["FPosted"] != null && row["FPosted"].ToString() != "")
        //        {
        //            model.FPosted = int.Parse(row["FPosted"].ToString());
        //        }
        //        if (row["FCheckSelect"] != null && row["FCheckSelect"].ToString() != "")
        //        {
        //            model.FCheckSelect = int.Parse(row["FCheckSelect"].ToString());
        //        }
        //        if (row["FCurrencyID"] != null && row["FCurrencyID"].ToString() != "")
        //        {
        //            model.FCurrencyID = int.Parse(row["FCurrencyID"].ToString());
        //        }
        //        if (row["FSaleStyle"] != null && row["FSaleStyle"].ToString() != "")
        //        {
        //            model.FSaleStyle = int.Parse(row["FSaleStyle"].ToString());
        //        }
        //        if (row["FAcctID"] != null && row["FAcctID"].ToString() != "")
        //        {
        //            model.FAcctID = int.Parse(row["FAcctID"].ToString());
        //        }
        //        if (row["FROB"] != null && row["FROB"].ToString() != "")
        //        {
        //            model.FROB = int.Parse(row["FROB"].ToString());
        //        }
        //        if (row["FRSCBillNo"] != null)
        //        {
        //            model.FRSCBillNo = row["FRSCBillNo"].ToString();
        //        }
        //        if (row["FStatus"] != null && row["FStatus"].ToString() != "")
        //        {
        //            model.FStatus = int.Parse(row["FStatus"].ToString());
        //        }
        //        if (row["FUpStockWhenSave"] != null && row["FUpStockWhenSave"].ToString() != "")
        //        {
        //            if ((row["FUpStockWhenSave"].ToString() == "1") || (row["FUpStockWhenSave"].ToString().ToLower() == "true"))
        //            {
        //                model.FUpStockWhenSave = true;
        //            }
        //            else
        //            {
        //                model.FUpStockWhenSave = false;
        //            }
        //        }
        //        if (row["FCancellation"] != null && row["FCancellation"].ToString() != "")
        //        {
        //            if ((row["FCancellation"].ToString() == "1") || (row["FCancellation"].ToString().ToLower() == "true"))
        //            {
        //                model.FCancellation = true;
        //            }
        //            else
        //            {
        //                model.FCancellation = false;
        //            }
        //        }
        //        if (row["FOrgBillInterID"] != null && row["FOrgBillInterID"].ToString() != "")
        //        {
        //            model.FOrgBillInterID = int.Parse(row["FOrgBillInterID"].ToString());
        //        }
        //        if (row["FBillTypeID"] != null && row["FBillTypeID"].ToString() != "")
        //        {
        //            model.FBillTypeID = int.Parse(row["FBillTypeID"].ToString());
        //        }
        //        if (row["FPOStyle"] != null && row["FPOStyle"].ToString() != "")
        //        {
        //            model.FPOStyle = int.Parse(row["FPOStyle"].ToString());
        //        }
        //        if (row["FMultiCheckLevel1"] != null && row["FMultiCheckLevel1"].ToString() != "")
        //        {
        //            model.FMultiCheckLevel1 = int.Parse(row["FMultiCheckLevel1"].ToString());
        //        }
        //        if (row["FMultiCheckLevel2"] != null && row["FMultiCheckLevel2"].ToString() != "")
        //        {
        //            model.FMultiCheckLevel2 = int.Parse(row["FMultiCheckLevel2"].ToString());
        //        }
        //        if (row["FMultiCheckLevel3"] != null && row["FMultiCheckLevel3"].ToString() != "")
        //        {
        //            model.FMultiCheckLevel3 = int.Parse(row["FMultiCheckLevel3"].ToString());
        //        }
        //        if (row["FMultiCheckLevel4"] != null && row["FMultiCheckLevel4"].ToString() != "")
        //        {
        //            model.FMultiCheckLevel4 = int.Parse(row["FMultiCheckLevel4"].ToString());
        //        }
        //        if (row["FMultiCheckLevel5"] != null && row["FMultiCheckLevel5"].ToString() != "")
        //        {
        //            model.FMultiCheckLevel5 = int.Parse(row["FMultiCheckLevel5"].ToString());
        //        }
        //        if (row["FMultiCheckLevel6"] != null && row["FMultiCheckLevel6"].ToString() != "")
        //        {
        //            model.FMultiCheckLevel6 = int.Parse(row["FMultiCheckLevel6"].ToString());
        //        }
        //        if (row["FMultiCheckDate1"] != null && row["FMultiCheckDate1"].ToString() != "")
        //        {
        //            model.FMultiCheckDate1 = DateTime.Parse(row["FMultiCheckDate1"].ToString());
        //        }
        //        if (row["FMultiCheckDate2"] != null && row["FMultiCheckDate2"].ToString() != "")
        //        {
        //            model.FMultiCheckDate2 = DateTime.Parse(row["FMultiCheckDate2"].ToString());
        //        }
        //        if (row["FMultiCheckDate3"] != null && row["FMultiCheckDate3"].ToString() != "")
        //        {
        //            model.FMultiCheckDate3 = DateTime.Parse(row["FMultiCheckDate3"].ToString());
        //        }
        //        if (row["FMultiCheckDate4"] != null && row["FMultiCheckDate4"].ToString() != "")
        //        {
        //            model.FMultiCheckDate4 = DateTime.Parse(row["FMultiCheckDate4"].ToString());
        //        }
        //        if (row["FMultiCheckDate5"] != null && row["FMultiCheckDate5"].ToString() != "")
        //        {
        //            model.FMultiCheckDate5 = DateTime.Parse(row["FMultiCheckDate5"].ToString());
        //        }
        //        if (row["FMultiCheckDate6"] != null && row["FMultiCheckDate6"].ToString() != "")
        //        {
        //            model.FMultiCheckDate6 = DateTime.Parse(row["FMultiCheckDate6"].ToString());
        //        }
        //        if (row["FCurCheckLevel"] != null && row["FCurCheckLevel"].ToString() != "")
        //        {
        //            model.FCurCheckLevel = int.Parse(row["FCurCheckLevel"].ToString());
        //        }
        //        if (row["FTaskID"] != null && row["FTaskID"].ToString() != "")
        //        {
        //            model.FTaskID = int.Parse(row["FTaskID"].ToString());
        //        }
        //        if (row["FResourceID"] != null && row["FResourceID"].ToString() != "")
        //        {
        //            model.FResourceID = int.Parse(row["FResourceID"].ToString());
        //        }
        //        if (row["FBackFlushed"] != null && row["FBackFlushed"].ToString() != "")
        //        {
        //            if ((row["FBackFlushed"].ToString() == "1") || (row["FBackFlushed"].ToString().ToLower() == "true"))
        //            {
        //                model.FBackFlushed = true;
        //            }
        //            else
        //            {
        //                model.FBackFlushed = false;
        //            }
        //        }
        //        if (row["FWBInterID"] != null && row["FWBInterID"].ToString() != "")
        //        {
        //            model.FWBInterID = int.Parse(row["FWBInterID"].ToString());
        //        }
        //        if (row["FTranStatus"] != null && row["FTranStatus"].ToString() != "")
        //        {
        //            model.FTranStatus = int.Parse(row["FTranStatus"].ToString());
        //        }
        //        if (row["FZPBillInterID"] != null && row["FZPBillInterID"].ToString() != "")
        //        {
        //            model.FZPBillInterID = int.Parse(row["FZPBillInterID"].ToString());
        //        }
        //        if (row["FRelateBrID"] != null && row["FRelateBrID"].ToString() != "")
        //        {
        //            model.FRelateBrID = int.Parse(row["FRelateBrID"].ToString());
        //        }
        //        if (row["FPurposeID"] != null && row["FPurposeID"].ToString() != "")
        //        {
        //            model.FPurposeID = int.Parse(row["FPurposeID"].ToString());
        //        }
        //        if (row["FUUID"] != null && row["FUUID"].ToString() != "")
        //        {
        //            model.FUUID = new Guid(row["FUUID"].ToString());
        //        }
        //        if (row["FRelateInvoiceID"] != null && row["FRelateInvoiceID"].ToString() != "")
        //        {
        //            model.FRelateInvoiceID = int.Parse(row["FRelateInvoiceID"].ToString());
        //        }
        //        if (row["FOperDate"] != null && row["FOperDate"].ToString() != "")
        //        {
        //            model.FOperDate = DateTime.Parse(row["FOperDate"].ToString());
        //        }
        //        if (row["FImport"] != null && row["FImport"].ToString() != "")
        //        {
        //            model.FImport = int.Parse(row["FImport"].ToString());
        //        }
        //        if (row["FSystemType"] != null && row["FSystemType"].ToString() != "")
        //        {
        //            model.FSystemType = int.Parse(row["FSystemType"].ToString());
        //        }
        //        if (row["FMarketingStyle"] != null && row["FMarketingStyle"].ToString() != "")
        //        {
        //            model.FMarketingStyle = int.Parse(row["FMarketingStyle"].ToString());
        //        }
        //        if (row["FPayBillID"] != null && row["FPayBillID"].ToString() != "")
        //        {
        //            model.FPayBillID = int.Parse(row["FPayBillID"].ToString());
        //        }
        //        if (row["FCheckDate"] != null && row["FCheckDate"].ToString() != "")
        //        {
        //            model.FCheckDate = DateTime.Parse(row["FCheckDate"].ToString());
        //        }
        //        if (row["FExplanation"] != null)
        //        {
        //            model.FExplanation = row["FExplanation"].ToString();
        //        }
        //        if (row["FFetchAdd"] != null)
        //        {
        //            model.FFetchAdd = row["FFetchAdd"].ToString();
        //        }
        //        if (row["FFetchDate"] != null && row["FFetchDate"].ToString() != "")
        //        {
        //            model.FFetchDate = DateTime.Parse(row["FFetchDate"].ToString());
        //        }
        //        if (row["FManagerID"] != null && row["FManagerID"].ToString() != "")
        //        {
        //            model.FManagerID = int.Parse(row["FManagerID"].ToString());
        //        }
        //        if (row["FRefType"] != null && row["FRefType"].ToString() != "")
        //        {
        //            model.FRefType = int.Parse(row["FRefType"].ToString());
        //        }
        //        if (row["FSelTranType"] != null && row["FSelTranType"].ToString() != "")
        //        {
        //            model.FSelTranType = int.Parse(row["FSelTranType"].ToString());
        //        }
        //        if (row["FChildren"] != null && row["FChildren"].ToString() != "")
        //        {
        //            model.FChildren = int.Parse(row["FChildren"].ToString());
        //        }
        //        if (row["FHookStatus"] != null && row["FHookStatus"].ToString() != "")
        //        {
        //            model.FHookStatus = int.Parse(row["FHookStatus"].ToString());
        //        }
        //        if (row["FActPriceVchTplID"] != null && row["FActPriceVchTplID"].ToString() != "")
        //        {
        //            model.FActPriceVchTplID = int.Parse(row["FActPriceVchTplID"].ToString());
        //        }
        //        if (row["FPlanPriceVchTplID"] != null && row["FPlanPriceVchTplID"].ToString() != "")
        //        {
        //            model.FPlanPriceVchTplID = int.Parse(row["FPlanPriceVchTplID"].ToString());
        //        }
        //        if (row["FProcID"] != null && row["FProcID"].ToString() != "")
        //        {
        //            model.FProcID = int.Parse(row["FProcID"].ToString());
        //        }
        //        if (row["FActualVchTplID"] != null && row["FActualVchTplID"].ToString() != "")
        //        {
        //            model.FActualVchTplID = int.Parse(row["FActualVchTplID"].ToString());
        //        }
        //        if (row["FPlanVchTplID"] != null && row["FPlanVchTplID"].ToString() != "")
        //        {
        //            model.FPlanVchTplID = int.Parse(row["FPlanVchTplID"].ToString());
        //        }
        //        if (row["FBrID"] != null && row["FBrID"].ToString() != "")
        //        {
        //            model.FBrID = int.Parse(row["FBrID"].ToString());
        //        }
        //        if (row["FVIPCardID"] != null && row["FVIPCardID"].ToString() != "")
        //        {
        //            model.FVIPCardID = int.Parse(row["FVIPCardID"].ToString());
        //        }
        //        if (row["FVIPScore"] != null && row["FVIPScore"].ToString() != "")
        //        {
        //            model.FVIPScore = decimal.Parse(row["FVIPScore"].ToString());
        //        }
        //        if (row["FHolisticDiscountRate"] != null && row["FHolisticDiscountRate"].ToString() != "")
        //        {
        //            model.FHolisticDiscountRate = decimal.Parse(row["FHolisticDiscountRate"].ToString());
        //        }
        //        if (row["FPOSName"] != null)
        //        {
        //            model.FPOSName = row["FPOSName"].ToString();
        //        }
        //        if (row["FWorkShiftId"] != null && row["FWorkShiftId"].ToString() != "")
        //        {
        //            model.FWorkShiftId = int.Parse(row["FWorkShiftId"].ToString());
        //        }
        //        if (row["FCussentAcctID"] != null && row["FCussentAcctID"].ToString() != "")
        //        {
        //            model.FCussentAcctID = int.Parse(row["FCussentAcctID"].ToString());
        //        }
        //        if (row["FZanGuCount"] != null && row["FZanGuCount"].ToString() != "")
        //        {
        //            if ((row["FZanGuCount"].ToString() == "1") || (row["FZanGuCount"].ToString().ToLower() == "true"))
        //            {
        //                model.FZanGuCount = true;
        //            }
        //            else
        //            {
        //                model.FZanGuCount = false;
        //            }
        //        }
        //        if (row["FPOOrdBillNo"] != null)
        //        {
        //            model.FPOOrdBillNo = row["FPOOrdBillNo"].ToString();
        //        }
        //        if (row["FLSSrcInterID"] != null && row["FLSSrcInterID"].ToString() != "")
        //        {
        //            model.FLSSrcInterID = int.Parse(row["FLSSrcInterID"].ToString());
        //        }
        //        if (row["FSettleDate"] != null && row["FSettleDate"].ToString() != "")
        //        {
        //            model.FSettleDate = DateTime.Parse(row["FSettleDate"].ToString());
        //        }
        //        if (row["FManageType"] != null && row["FManageType"].ToString() != "")
        //        {
        //            model.FManageType = int.Parse(row["FManageType"].ToString());
        //        }
        //        if (row["FHeadSelfB0146"] != null)
        //        {
        //            model.FHeadSelfB0146 = row["FHeadSelfB0146"].ToString();
        //        }
        //        if (row["FHeadSelfB0147"] != null)
        //        {
        //            model.FHeadSelfB0147 = row["FHeadSelfB0147"].ToString();
        //        }
        //        if (row["FHeadSelfB0148"] != null)
        //        {
        //            model.FHeadSelfB0148 = row["FHeadSelfB0148"].ToString();
        //        }
        //        if (row["FHeadSelfB0149"] != null)
        //        {
        //            model.FHeadSelfB0149 = row["FHeadSelfB0149"].ToString();
        //        }
        //        if (row["FHeadSelfB0150"] != null)
        //        {
        //            model.FHeadSelfB0150 = row["FHeadSelfB0150"].ToString();
        //        }
        //        if (row["FHeadSelfB0151"] != null)
        //        {
        //            model.FHeadSelfB0151 = row["FHeadSelfB0151"].ToString();
        //        }
        //    }
        //    return model;
        //}

        ///// <summary>
        ///// 获得数据列表
        ///// </summary>
        //public DataSet GetList(string sConnectionName, string strWhere)
        //{
        //    StringBuilder strSql = new StringBuilder();
        //    strSql.Append("select FBrNo,FInterID,FTranType,FDate,FBillNo,FUse,FNote,FDCStockID,FSCStockID,FDeptID,FEmpID,FSupplyID,FPosterID,FCheckerID,FFManagerID,FSManagerID,FBillerID,FReturnBillInterID,FSCBillNo,FHookInterID,FVchInterID,FPosted,FCheckSelect,FCurrencyID,FSaleStyle,FAcctID,FROB,FRSCBillNo,FStatus,FUpStockWhenSave,FCancellation,FOrgBillInterID,FBillTypeID,FPOStyle,FMultiCheckLevel1,FMultiCheckLevel2,FMultiCheckLevel3,FMultiCheckLevel4,FMultiCheckLevel5,FMultiCheckLevel6,FMultiCheckDate1,FMultiCheckDate2,FMultiCheckDate3,FMultiCheckDate4,FMultiCheckDate5,FMultiCheckDate6,FCurCheckLevel,FTaskID,FResourceID,FBackFlushed,FWBInterID,FTranStatus,FZPBillInterID,FRelateBrID,FPurposeID,FUUID,FRelateInvoiceID,FOperDate,FImport,FSystemType,FMarketingStyle,FPayBillID,FCheckDate,FExplanation,FFetchAdd,FFetchDate,FManagerID,FRefType,FSelTranType,FChildren,FHookStatus,FActPriceVchTplID,FPlanPriceVchTplID,FProcID,FActualVchTplID,FPlanVchTplID,FBrID,FVIPCardID,FVIPScore,FHolisticDiscountRate,FPOSName,FWorkShiftId,FCussentAcctID,FZanGuCount,FPOOrdBillNo,FLSSrcInterID,FSettleDate,FManageType,FHeadSelfB0146,FHeadSelfB0147,FHeadSelfB0148,FHeadSelfB0149,FHeadSelfB0150,FHeadSelfB0151 ");
        //    strSql.Append(" FROM ICStockBill ");
        //    if (strWhere.Trim() != "")
        //    {
        //        strSql.Append(" where " + strWhere);
        //    }
        //    return SqlHelper.ExecuteDataSet(SqlHelper.GetConnectionString(sConnectionName), strSql.ToString());
        //}

        ///// <summary>
        ///// 获得前几行数据
        ///// </summary>
        //public DataSet GetList(string sConnectionName, int Top, string strWhere, string filedOrder)
        //{
        //    StringBuilder strSql = new StringBuilder();
        //    strSql.Append("select ");
        //    if (Top > 0)
        //    {
        //        strSql.Append(" top " + Top.ToString());
        //    }
        //    strSql.Append(" FBrNo,FInterID,FTranType,FDate,FBillNo,FUse,FNote,FDCStockID,FSCStockID,FDeptID,FEmpID,FSupplyID,FPosterID,FCheckerID,FFManagerID,FSManagerID,FBillerID,FReturnBillInterID,FSCBillNo,FHookInterID,FVchInterID,FPosted,FCheckSelect,FCurrencyID,FSaleStyle,FAcctID,FROB,FRSCBillNo,FStatus,FUpStockWhenSave,FCancellation,FOrgBillInterID,FBillTypeID,FPOStyle,FMultiCheckLevel1,FMultiCheckLevel2,FMultiCheckLevel3,FMultiCheckLevel4,FMultiCheckLevel5,FMultiCheckLevel6,FMultiCheckDate1,FMultiCheckDate2,FMultiCheckDate3,FMultiCheckDate4,FMultiCheckDate5,FMultiCheckDate6,FCurCheckLevel,FTaskID,FResourceID,FBackFlushed,FWBInterID,FTranStatus,FZPBillInterID,FRelateBrID,FPurposeID,FUUID,FRelateInvoiceID,FOperDate,FImport,FSystemType,FMarketingStyle,FPayBillID,FCheckDate,FExplanation,FFetchAdd,FFetchDate,FManagerID,FRefType,FSelTranType,FChildren,FHookStatus,FActPriceVchTplID,FPlanPriceVchTplID,FProcID,FActualVchTplID,FPlanVchTplID,FBrID,FVIPCardID,FVIPScore,FHolisticDiscountRate,FPOSName,FWorkShiftId,FCussentAcctID,FZanGuCount,FPOOrdBillNo,FLSSrcInterID,FSettleDate,FManageType,FHeadSelfB0146,FHeadSelfB0147,FHeadSelfB0148,FHeadSelfB0149,FHeadSelfB0150,FHeadSelfB0151 ");
        //    strSql.Append(" FROM ICStockBill ");
        //    if (strWhere.Trim() != "")
        //    {
        //        strSql.Append(" where " + strWhere);
        //    }
        //    strSql.Append(" order by " + filedOrder);
        //    return SqlHelper.ExecuteDataSet(SqlHelper.GetConnectionString(sConnectionName), strSql.ToString());
        //}

        ///// <summary>
        ///// 获取记录总数
        ///// </summary>
        //public int GetRecordCount(string sConnectionName, string strWhere)
        //{
        //    StringBuilder strSql = new StringBuilder();
        //    strSql.Append("select count(1) FROM ICStockBill ");
        //    if (strWhere.Trim() != "")
        //    {
        //        strSql.Append(" where " + strWhere);
        //    }
        //    object obj = SqlHelper.GetSingle(SqlHelper.GetConnectionString(sConnectionName), strSql.ToString());
        //    if (obj == null)
        //    {
        //        return 0;
        //    }
        //    else
        //    {
        //        return Convert.ToInt32(obj);
        //    }
        //}

        ///// <summary>
        ///// 分页获取数据列表
        ///// </summary>
        //public DataSet GetListByPage(string sConnectionName, string strWhere, string orderby, int startIndex, int endIndex)
        //{
        //    StringBuilder strSql = new StringBuilder();
        //    strSql.Append("SELECT * FROM ( ");
        //    strSql.Append(" SELECT ROW_NUMBER() OVER (");
        //    if (!string.IsNullOrEmpty(orderby.Trim()))
        //    {
        //        strSql.Append("order by T." + orderby);
        //    }
        //    else
        //    {
        //        strSql.Append("order by T.FRelateInvoiceID desc");
        //    }
        //    strSql.Append(")AS Row, T.*  from ICStockBill T ");
        //    if (!string.IsNullOrEmpty(strWhere.Trim()))
        //    {
        //        strSql.Append(" WHERE " + strWhere);
        //    }
        //    strSql.Append(" ) TT");
        //    strSql.AppendFormat(" WHERE TT.Row between {0} and {1}", startIndex, endIndex);
        //    return SqlHelper.ExecuteDataSet(SqlHelper.GetConnectionString(sConnectionName), strSql.ToString());
        //}

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
            parameters[0].Value = "ICStockBill";
            parameters[1].Value = "FRelateInvoiceID";
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
        /// 按日期和单据类型删除数据
        /// </summary>
        /// <param name="connectionName"></param>
        /// <param name="minDate"></param>
        /// <param name="maxDate"></param>
        /// <param name="tranType"></param>
        /// <returns></returns>
        public int Delete(string connectionName, string minDate, string maxDate, int tranType)
        {
            StringBuilder strSql = new StringBuilder();
            strSql.Append(" delete dbo.ICStockBill from dbo.ICStockBill inner join dbo.ICStockBillEntry on ICStockBill.finterid = ICStockBillEntry.FInterID  ");
            strSql.Append(" Where ICStockBill.FDate >= @minDate and ICStockBill.FDate <= @maxDate and ICStockBill.fTranType =@trantype ");
            SqlParameter[] parameters = {
					new SqlParameter("@minDate", SqlDbType.NVarChar,20),
					new SqlParameter("@maxDate", SqlDbType.NVarChar,20),
                    new SqlParameter ("@trantype",SqlDbType .Int)};
            parameters[0].Value = minDate;
            parameters[1].Value = maxDate;
            parameters[2].Value = tranType;

            int rows = SqlHelper.ExecuteSql(SqlHelper.GetConnectionString(connectionName), strSql.ToString(), parameters);
            if (rows > 0)
            {
                return rows;
            }
            else
            {
                return 0;
            }
        }

        /// <summary>
        /// 得到当前的内联ID号
        /// </summary>
        /// <param name="connectionName"></param>
        /// <returns></returns>
        public int GetMaxFInterID(string connectionName)
        {
            StringBuilder strSql = new StringBuilder();
            strSql.Append(" Select FMaxNum From ICMaxNum Where FTableName = 'ICStockBill' ");
            //return int.Parse(SqlHelper.GetSingle(SqlHelper.GetConnectionString(connectionName), strSql.ToString(), null).ToString());
            object obj = SqlHelper.GetSingle(SqlHelper.GetConnectionString(connectionName), strSql.ToString(), null);
            return obj != null ? int.Parse(obj.ToString()) : 0;

        }

        /// <summary>
        /// 通过仓库名称得到仓库的编号
        /// </summary>
        public int GetStockID(string connectionName, string stockName)
        {
            StringBuilder strSql = new StringBuilder();
            strSql.Append(String.Format(" select FItemID from t_stock where FName = '{0}' ", stockName));
             object obj = SqlHelper.GetSingle(SqlHelper.GetConnectionString(connectionName), strSql.ToString(), null);
            return obj != null  ? int.Parse(obj.ToString()) : 0;
        }

        /// <summary>
        /// 更新当前的内联ID号
        /// </summary>
        public bool UpdateFInterID(string connectionName)
        {
            StringBuilder strSql = new StringBuilder();
            strSql.Append(" Update ICMaxNum set FMaxNum = FMaxNum + 1 where FTableName = 'ICStockBill' ");
            int rows = SqlHelper.ExecuteSql(SqlHelper.GetConnectionString(connectionName), strSql.ToString(), null);
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
        /// 更新当前的单据号
        /// </summary>
        /// <param name="connectionName"></param>
        /// <param name="fBillNum"></param>
        /// <param name="FDesc"></param>
        /// <param name="fTranType"></param>
        /// <returns></returns>
        public bool UpdateFBillNo(string connectionName, int fBillNum, string FDesc, int fTranType)
        {
            StringBuilder strSql = new StringBuilder();
            if (fTranType == 1)
            {
                strSql.Append(String.Format(" Update ICBillNo Set FCurNo = {0} , FDesc = '{1}' Where FBillName = '外购入库单据' ", fBillNum, FDesc));
            }
            else
            {
                strSql.Append(String.Format(" Update ICBillNo Set FCurNo = {0} , FDesc = '{1}' Where FBillName = '销售出库单据' ", fBillNum, FDesc));

            }
            int rows = SqlHelper.ExecuteSql(SqlHelper.GetConnectionString(connectionName), strSql.ToString(), null);
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
        /// 通过用户名称得到用户的编号
        /// </summary>
        /// <param name="connectionName"></param>
        /// <param name="UserName"></param>
        /// <returns></returns>
        public int GetUserIDByName(string connectionName, string UserName)
        {
            StringBuilder strSql = new StringBuilder();
            strSql.Append(String.Format(" Select FUserID From t_User Where FName = '{0}' ", UserName));
            object obj = SqlHelper.GetSingle(SqlHelper.GetConnectionString(connectionName), strSql.ToString(), null).ToString();
            return obj != null ? int.Parse(obj.ToString()) : 0;
        }
        #endregion  ExtensionMethod
    }
}


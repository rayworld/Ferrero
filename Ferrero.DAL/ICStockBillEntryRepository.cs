using System;
using System.Data;
using System.Text;
using System.Data.SqlClient;
using Ray.Framework.DBUtility;//Please add references

namespace Ferrero.DAL
{
    /// <summary>
    /// 数据访问类:ICStockBillEntry
    /// </summary>
    public partial class ICStockBillEntryRepository
    {
        public ICStockBillEntryRepository()
        { }

        #region  BasicMethod

        /// <summary>
        /// 增加一条数据
        /// </summary>
        public bool Add(string sConnectionName, Ferrero.Model.ICStockBillEntry model)
        {
            StringBuilder strSql = new StringBuilder();
            if (sConnectionName.ToLower() == "wuhan")
            {
                strSql.Append("insert into ICStockBillEntry(");
                strSql.Append("FBrNo,FInterID,FEntryID,FItemID,FQtyMust,FQty,FPrice,FBatchNo,FAmount,FNote,FSCBillInterID,FSCBillNo,FUnitID,FAuxPrice,FAuxQty,FAuxQtyMust,FQtyActual,FAuxQtyActual,FPlanPrice,FAuxPlanPrice,FSourceEntryID,FCommitQty,FAuxCommitQty,FKFDate,FKFPeriod,FDCSPID,FSCSPID,FConsignPrice,FConsignAmount,FProcessCost,FMaterialCost,FTaxAmount,FMapNumber,FMapName,FOrgBillEntryID,FOperID,FPlanAmount,FProcessPrice,FTaxRate,FSnListID,FAmtRef,FAuxPropID,FCost,FPriceRef,FAuxPriceRef,FFetchDate,FQtyInvoice,FQtyInvoiceBase,FUnitCost,FSecCoefficient,FSecQty,FSecCommitQty,FSourceTranType,FSourceInterId,FSourceBillNo,FContractInterID,FContractEntryID,FContractBillNo,FICMOBillNo,FICMOInterID,FPPBomEntryID,FOrderInterID,FOrderEntryID,FOrderBillNo,FAllHookQTY,FAllHookAmount,FCurrentHookQTY,FCurrentHookAmount,FStdAllHookAmount,FStdCurrentHookAmount,FSCStockID,FDCStockID,FPeriodDate,FCostObjGroupID,FCostOBJID,FMaterialCostPrice,FReProduceType,FBomInterID,FDiscountRate,FDiscountAmount,FSepcialSaleId,FOutCommitQty,FOutSecCommitQty,FDBCommitQty,FDBSecCommitQty,FAuxQtyInvoice,FOperSN,FCheckStatus,FSplitSecQty,FInStockID,FSaleCommitQty,FSaleSecCommitQty,FSaleAuxCommitQty,FSelectedProcID,FVWInStockQty,FAuxVWInStockQty,FSecVWInStockQty,FSecInvoiceQty,FCostCenterID,FEntrySelfB0158,FEntrySelfB0160,FEntrySelfB0161,FEntrySelfB0164,FEntrySelfB0165,FEntrySelfB0166,FEntrySelfB0167,FEntrySelfB0168,FEntrySelfB0169,FEntrySelfB0171,FEntrySelfB0170,FEntrySelfB0162,FEntrySelfB0172,FEntrySelfB0173,FEntrySelfB0157,FEntrySelfB0163,FEntrySelfB0159)");
                strSql.Append(" values (");
                strSql.Append("@FBrNo,@FInterID,@FEntryID,@FItemID,@FQtyMust,@FQty,@FPrice,@FBatchNo,@FAmount,@FNote,@FSCBillInterID,@FSCBillNo,@FUnitID,@FAuxPrice,@FAuxQty,@FAuxQtyMust,@FQtyActual,@FAuxQtyActual,@FPlanPrice,@FAuxPlanPrice,@FSourceEntryID,@FCommitQty,@FAuxCommitQty,@FKFDate,@FKFPeriod,@FDCSPID,@FSCSPID,@FConsignPrice,@FConsignAmount,@FProcessCost,@FMaterialCost,@FTaxAmount,@FMapNumber,@FMapName,@FOrgBillEntryID,@FOperID,@FPlanAmount,@FProcessPrice,@FTaxRate,@FSnListID,@FAmtRef,@FAuxPropID,@FCost,@FPriceRef,@FAuxPriceRef,@FFetchDate,@FQtyInvoice,@FQtyInvoiceBase,@FUnitCost,@FSecCoefficient,@FSecQty,@FSecCommitQty,@FSourceTranType,@FSourceInterId,@FSourceBillNo,@FContractInterID,@FContractEntryID,@FContractBillNo,@FICMOBillNo,@FICMOInterID,@FPPBomEntryID,@FOrderInterID,@FOrderEntryID,@FOrderBillNo,@FAllHookQTY,@FAllHookAmount,@FCurrentHookQTY,@FCurrentHookAmount,@FStdAllHookAmount,@FStdCurrentHookAmount,@FSCStockID,@FDCStockID,@FPeriodDate,@FCostObjGroupID,@FCostOBJID,@FMaterialCostPrice,@FReProduceType,@FBomInterID,@FDiscountRate,@FDiscountAmount,@FSepcialSaleId,@FOutCommitQty,@FOutSecCommitQty,@FDBCommitQty,@FDBSecCommitQty,@FAuxQtyInvoice,@FOperSN,@FCheckStatus,@FSplitSecQty,@FInStockID,@FSaleCommitQty,@FSaleSecCommitQty,@FSaleAuxCommitQty,@FSelectedProcID,@FVWInStockQty,@FAuxVWInStockQty,@FSecVWInStockQty,@FSecInvoiceQty,@FCostCenterID,@FEntrySelfB0158,@FEntrySelfB0160,@FEntrySelfB0161,@FEntrySelfB0164,@FEntrySelfB0165,@FEntrySelfB0166,@FEntrySelfB0167,@FEntrySelfB0168,@FEntrySelfB0169,@FEntrySelfB0171,@FEntrySelfB0170,@FEntrySelfB0162,@FEntrySelfB0172,@FEntrySelfB0173,@FEntrySelfB0157,@FEntrySelfB0163,@FEntrySelfB0159)");
                strSql.Append(";select @@IDENTITY");
                SqlParameter[] parameters = {
					new SqlParameter("@FBrNo", SqlDbType.VarChar,10),
					new SqlParameter("@FInterID", SqlDbType.Int,4),
					new SqlParameter("@FEntryID", SqlDbType.Int,4),
					new SqlParameter("@FItemID", SqlDbType.Int,4),
					new SqlParameter("@FQtyMust", SqlDbType.Decimal,13),
					new SqlParameter("@FQty", SqlDbType.Decimal,13),
					new SqlParameter("@FPrice", SqlDbType.Decimal,13),
					new SqlParameter("@FBatchNo", SqlDbType.VarChar,200),
					new SqlParameter("@FAmount", SqlDbType.Decimal,13),
					new SqlParameter("@FNote", SqlDbType.VarChar,255),
					new SqlParameter("@FSCBillInterID", SqlDbType.Int,4),
					new SqlParameter("@FSCBillNo", SqlDbType.VarChar,30),
					new SqlParameter("@FUnitID", SqlDbType.Int,4),
					new SqlParameter("@FAuxPrice", SqlDbType.Decimal,13),
					new SqlParameter("@FAuxQty", SqlDbType.Decimal,13),
					new SqlParameter("@FAuxQtyMust", SqlDbType.Decimal,13),
					new SqlParameter("@FQtyActual", SqlDbType.Decimal,13),
					new SqlParameter("@FAuxQtyActual", SqlDbType.Decimal,13),
					new SqlParameter("@FPlanPrice", SqlDbType.Decimal,13),
					new SqlParameter("@FAuxPlanPrice", SqlDbType.Decimal,13),
					new SqlParameter("@FSourceEntryID", SqlDbType.Int,4),
					new SqlParameter("@FCommitQty", SqlDbType.Decimal,13),
					new SqlParameter("@FAuxCommitQty", SqlDbType.Decimal,13),
					new SqlParameter("@FKFDate", SqlDbType.DateTime),
					new SqlParameter("@FKFPeriod", SqlDbType.Int,4),
					new SqlParameter("@FDCSPID", SqlDbType.Int,4),
					new SqlParameter("@FSCSPID", SqlDbType.Int,4),
					new SqlParameter("@FConsignPrice", SqlDbType.Decimal,13),
					new SqlParameter("@FConsignAmount", SqlDbType.Decimal,13),
					new SqlParameter("@FProcessCost", SqlDbType.Decimal,13),
					new SqlParameter("@FMaterialCost", SqlDbType.Decimal,13),
					new SqlParameter("@FTaxAmount", SqlDbType.Decimal,13),
					new SqlParameter("@FMapNumber", SqlDbType.VarChar,80),
					new SqlParameter("@FMapName", SqlDbType.NVarChar,256),
					new SqlParameter("@FOrgBillEntryID", SqlDbType.Int,4),
					new SqlParameter("@FOperID", SqlDbType.Int,4),
					new SqlParameter("@FPlanAmount", SqlDbType.Decimal,13),
					new SqlParameter("@FProcessPrice", SqlDbType.Decimal,13),
					new SqlParameter("@FTaxRate", SqlDbType.Decimal,13),
					new SqlParameter("@FSnListID", SqlDbType.Int,4),
					new SqlParameter("@FAmtRef", SqlDbType.Decimal,13),
					new SqlParameter("@FAuxPropID", SqlDbType.Int,4),
					new SqlParameter("@FCost", SqlDbType.Decimal,13),
					new SqlParameter("@FPriceRef", SqlDbType.Decimal,13),
					new SqlParameter("@FAuxPriceRef", SqlDbType.Decimal,13),
					new SqlParameter("@FFetchDate", SqlDbType.DateTime),
					new SqlParameter("@FQtyInvoice", SqlDbType.Decimal,13),
					new SqlParameter("@FQtyInvoiceBase", SqlDbType.Decimal,13),
					new SqlParameter("@FUnitCost", SqlDbType.Decimal,13),
					new SqlParameter("@FSecCoefficient", SqlDbType.Decimal,13),
					new SqlParameter("@FSecQty", SqlDbType.Decimal,13),
					new SqlParameter("@FSecCommitQty", SqlDbType.Decimal,13),
					new SqlParameter("@FSourceTranType", SqlDbType.Int,4),
					new SqlParameter("@FSourceInterId", SqlDbType.Int,4),
					new SqlParameter("@FSourceBillNo", SqlDbType.NVarChar,255),
					new SqlParameter("@FContractInterID", SqlDbType.Int,4),
					new SqlParameter("@FContractEntryID", SqlDbType.Int,4),
					new SqlParameter("@FContractBillNo", SqlDbType.NVarChar,255),
					new SqlParameter("@FICMOBillNo", SqlDbType.NVarChar,255),
					new SqlParameter("@FICMOInterID", SqlDbType.Int,4),
					new SqlParameter("@FPPBomEntryID", SqlDbType.Int,4),
					new SqlParameter("@FOrderInterID", SqlDbType.Int,4),
					new SqlParameter("@FOrderEntryID", SqlDbType.Int,4),
					new SqlParameter("@FOrderBillNo", SqlDbType.NVarChar,255),
					new SqlParameter("@FAllHookQTY", SqlDbType.Decimal,13),
					new SqlParameter("@FAllHookAmount", SqlDbType.Decimal,13),
					new SqlParameter("@FCurrentHookQTY", SqlDbType.Decimal,13),
					new SqlParameter("@FCurrentHookAmount", SqlDbType.Decimal,13),
					new SqlParameter("@FStdAllHookAmount", SqlDbType.Decimal,13),
					new SqlParameter("@FStdCurrentHookAmount", SqlDbType.Decimal,13),
					new SqlParameter("@FSCStockID", SqlDbType.Int,4),
					new SqlParameter("@FDCStockID", SqlDbType.Int,4),
					new SqlParameter("@FPeriodDate", SqlDbType.DateTime),
					new SqlParameter("@FCostObjGroupID", SqlDbType.Int,4),
					new SqlParameter("@FCostOBJID", SqlDbType.Int,4),
					new SqlParameter("@FMaterialCostPrice", SqlDbType.Decimal,13),
					new SqlParameter("@FReProduceType", SqlDbType.Int,4),
					new SqlParameter("@FBomInterID", SqlDbType.Int,4),
					new SqlParameter("@FDiscountRate", SqlDbType.Decimal,13),
					new SqlParameter("@FDiscountAmount", SqlDbType.Decimal,13),
					new SqlParameter("@FSepcialSaleId", SqlDbType.Int,4),
					new SqlParameter("@FOutCommitQty", SqlDbType.Decimal,13),
					new SqlParameter("@FOutSecCommitQty", SqlDbType.Decimal,13),
					new SqlParameter("@FDBCommitQty", SqlDbType.Decimal,13),
					new SqlParameter("@FDBSecCommitQty", SqlDbType.Decimal,13),
					new SqlParameter("@FAuxQtyInvoice", SqlDbType.Decimal,13),
					new SqlParameter("@FOperSN", SqlDbType.Int,4),
					new SqlParameter("@FCheckStatus", SqlDbType.SmallInt,2),
					new SqlParameter("@FSplitSecQty", SqlDbType.Decimal,13),
					new SqlParameter("@FInStockID", SqlDbType.Int,4),
					new SqlParameter("@FSaleCommitQty", SqlDbType.Decimal,13),
					new SqlParameter("@FSaleSecCommitQty", SqlDbType.Decimal,13),
					new SqlParameter("@FSaleAuxCommitQty", SqlDbType.Decimal,13),
					new SqlParameter("@FSelectedProcID", SqlDbType.Int,4),
					new SqlParameter("@FVWInStockQty", SqlDbType.Decimal,13),
					new SqlParameter("@FAuxVWInStockQty", SqlDbType.Decimal,13),
					new SqlParameter("@FSecVWInStockQty", SqlDbType.Decimal,13),
					new SqlParameter("@FSecInvoiceQty", SqlDbType.Decimal,13),
					new SqlParameter("@FCostCenterID", SqlDbType.Int,4),
					new SqlParameter("@FEntrySelfB0158", SqlDbType.VarChar,255),
					new SqlParameter("@FEntrySelfB0160", SqlDbType.VarChar,255),
					new SqlParameter("@FEntrySelfB0161", SqlDbType.VarChar,255),
					new SqlParameter("@FEntrySelfB0164", SqlDbType.VarChar,255),
					new SqlParameter("@FEntrySelfB0165", SqlDbType.VarChar,255),
					new SqlParameter("@FEntrySelfB0166", SqlDbType.VarChar,255),
					new SqlParameter("@FEntrySelfB0167", SqlDbType.VarChar,255),
					new SqlParameter("@FEntrySelfB0168", SqlDbType.VarChar,255),
					new SqlParameter("@FEntrySelfB0169", SqlDbType.VarChar,255),
					new SqlParameter("@FEntrySelfB0171", SqlDbType.VarChar,255),
					new SqlParameter("@FEntrySelfB0170", SqlDbType.Decimal,13),
					new SqlParameter("@FEntrySelfB0162", SqlDbType.Decimal,13),
					new SqlParameter("@FEntrySelfB0172", SqlDbType.Decimal,13),
					new SqlParameter("@FEntrySelfB0173", SqlDbType.Decimal,13),
					new SqlParameter("@FEntrySelfB0157", SqlDbType.Decimal,13),
					new SqlParameter("@FEntrySelfB0163", SqlDbType.Decimal,13),
					new SqlParameter("@FEntrySelfB0159", SqlDbType.Decimal,13)};
                parameters[0].Value = "0";
                parameters[1].Value = model.FInterID;
                parameters[2].Value = model.FEntryID;
                parameters[3].Value = model.FItemID;
                parameters[4].Value = 0;
                parameters[5].Value = model.FQty;
                parameters[6].Value = model.FPrice;
                parameters[7].Value = model.FBatchNo;
                parameters[8].Value = model.FAmount;
                parameters[9].Value = "";
                parameters[10].Value = DBNull.Value;
                parameters[11].Value = DBNull.Value;
                parameters[12].Value = model.FUnitID;
                parameters[13].Value = model.FAuxPrice;
                parameters[14].Value = model.FAuxQty;
                parameters[15].Value = 0;
                parameters[16].Value = 0;
                parameters[17].Value = 0;
                parameters[18].Value = 0;
                parameters[19].Value = 0;
                parameters[20].Value = 0;
                parameters[21].Value = 0;
                parameters[22].Value = 0;
                parameters[23].Value = DBNull.Value;
                parameters[24].Value = model.FKFPeriod;
                parameters[25].Value = 0;
                parameters[26].Value = DBNull.Value;
                parameters[27].Value = model.FConsignPrice;
                parameters[28].Value = model.FConsignAmount;
                parameters[29].Value = 0;
                parameters[30].Value = 0;
                parameters[31].Value = 0;
                parameters[32].Value = "";
                parameters[33].Value = "";
                parameters[34].Value = 0;
                parameters[35].Value = 0;
                parameters[36].Value = 0;
                parameters[37].Value = 0;
                parameters[38].Value = 0;
                parameters[39].Value = 0;
                parameters[40].Value = 0;
                parameters[41].Value = 0;
                parameters[42].Value = 0;
                parameters[43].Value = 0;
                parameters[44].Value = 0;
                parameters[45].Value = DBNull.Value;
                parameters[46].Value = 0;
                parameters[47].Value = 0;
                parameters[48].Value = 0;
                parameters[49].Value = 0;
                parameters[50].Value = 0;
                parameters[51].Value = 0;
                parameters[52].Value = 0;
                parameters[53].Value = 0;
                parameters[54].Value = "";
                parameters[55].Value = 0;
                parameters[56].Value = 0;
                parameters[57].Value = "";
                parameters[58].Value = "";
                parameters[59].Value = 0;
                parameters[60].Value = 0;
                parameters[61].Value = 0;
                parameters[62].Value = 0;
                parameters[63].Value = "";
                parameters[64].Value = 0;
                parameters[65].Value = 0;
                parameters[66].Value = 0;
                parameters[67].Value = 0;
                parameters[68].Value = 0;
                parameters[69].Value = 0;
                parameters[70].Value = 0;
                parameters[71].Value = model.FDCStockID;
                parameters[72].Value = DBNull.Value;
                parameters[73].Value = 0;
                parameters[74].Value = 0;
                parameters[75].Value = 0;
                parameters[76].Value = 0;
                parameters[77].Value = 0;
                parameters[78].Value = 0;
                parameters[79].Value = 0;
                parameters[80].Value = 0;
                parameters[81].Value = 0;
                parameters[82].Value = 0;
                parameters[83].Value = 0;
                parameters[84].Value = 0;
                parameters[85].Value = 0;
                parameters[86].Value = 0;
                parameters[87].Value = 0;
                parameters[88].Value = DBNull.Value;
                parameters[89].Value = 0;
                parameters[90].Value = 0;
                parameters[91].Value = 0;
                parameters[92].Value = 0;
                parameters[93].Value = 0;
                parameters[94].Value = 0;
                parameters[95].Value = 0;
                parameters[96].Value = 0;
                parameters[97].Value = 0;
                parameters[98].Value = 0;
                parameters[99].Value = model.FEntrySelfB0158;
                parameters[100].Value = model.FEntrySelfB0160;
                parameters[101].Value = model.FEntrySelfB0161;
                parameters[102].Value = model.FEntrySelfB0164;
                parameters[103].Value = model.FEntrySelfB0165;
                parameters[104].Value = model.FEntrySelfB0166;
                parameters[105].Value = model.FEntrySelfB0167;
                parameters[106].Value = model.FEntrySelfB0168;
                parameters[107].Value = model.FEntrySelfB0169;
                parameters[108].Value = model.FEntrySelfB0171;
                parameters[109].Value = model.FEntrySelfB0170;
                parameters[110].Value = model.FEntrySelfB0162;
                parameters[111].Value = model.FEntrySelfB0172;
                parameters[112].Value = model.FEntrySelfB0173;
                parameters[113].Value = model.FEntrySelfB0157;
                parameters[114].Value = model.FEntrySelfB0163;
                parameters[115].Value = model.FEntrySelfB0159;

                int rows = SqlHelper.ExecuteSql(SqlHelper.GetConnectionString(sConnectionName),strSql.ToString(), parameters);
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
                strSql.Append("insert into ICStockBillEntry(");
                strSql.Append("FBrNo,FInterID,FEntryID,FItemID,FQtyMust,FQty,FPrice,FBatchNo,FAmount,FNote,FSCBillInterID,FSCBillNo,FUnitID,FAuxPrice,FAuxQty,FAuxQtyMust,FQtyActual,FAuxQtyActual,FPlanPrice,FAuxPlanPrice,FSourceEntryID,FCommitQty,FAuxCommitQty,FKFDate,FKFPeriod,FDCSPID,FSCSPID,FConsignPrice,FConsignAmount,FProcessCost,FMaterialCost,FTaxAmount,FMapNumber,FMapName,FOrgBillEntryID,FOperID,FPlanAmount,FProcessPrice,FTaxRate,FSnListID,FAmtRef,FAuxPropID,FCost,FPriceRef,FAuxPriceRef,FFetchDate,FQtyInvoice,FQtyInvoiceBase,FUnitCost,FSecCoefficient,FSecQty,FSecCommitQty,FSourceTranType,FSourceInterId,FSourceBillNo,FContractInterID,FContractEntryID,FContractBillNo,FICMOBillNo,FICMOInterID,FPPBomEntryID,FOrderInterID,FOrderEntryID,FOrderBillNo,FAllHookQTY,FAllHookAmount,FCurrentHookQTY,FCurrentHookAmount,FStdAllHookAmount,FStdCurrentHookAmount,FSCStockID,FDCStockID,FPeriodDate,FCostObjGroupID,FCostOBJID,FMaterialCostPrice,FReProduceType,FBomInterID,FDiscountRate,FDiscountAmount,FSepcialSaleId,FOutCommitQty,FOutSecCommitQty,FDBCommitQty,FDBSecCommitQty,FAuxQtyInvoice,FOperSN,FCheckStatus,FSplitSecQty,FInStockID,FSaleCommitQty,FSaleSecCommitQty,FSaleAuxCommitQty,FSelectedProcID,FVWInStockQty,FAuxVWInStockQty,FSecVWInStockQty,FSecInvoiceQty,FCostCenterID)");
                strSql.Append(" values (");
                strSql.Append("@FBrNo,@FInterID,@FEntryID,@FItemID,@FQtyMust,@FQty,@FPrice,@FBatchNo,@FAmount,@FNote,@FSCBillInterID,@FSCBillNo,@FUnitID,@FAuxPrice,@FAuxQty,@FAuxQtyMust,@FQtyActual,@FAuxQtyActual,@FPlanPrice,@FAuxPlanPrice,@FSourceEntryID,@FCommitQty,@FAuxCommitQty,@FKFDate,@FKFPeriod,@FDCSPID,@FSCSPID,@FConsignPrice,@FConsignAmount,@FProcessCost,@FMaterialCost,@FTaxAmount,@FMapNumber,@FMapName,@FOrgBillEntryID,@FOperID,@FPlanAmount,@FProcessPrice,@FTaxRate,@FSnListID,@FAmtRef,@FAuxPropID,@FCost,@FPriceRef,@FAuxPriceRef,@FFetchDate,@FQtyInvoice,@FQtyInvoiceBase,@FUnitCost,@FSecCoefficient,@FSecQty,@FSecCommitQty,@FSourceTranType,@FSourceInterId,@FSourceBillNo,@FContractInterID,@FContractEntryID,@FContractBillNo,@FICMOBillNo,@FICMOInterID,@FPPBomEntryID,@FOrderInterID,@FOrderEntryID,@FOrderBillNo,@FAllHookQTY,@FAllHookAmount,@FCurrentHookQTY,@FCurrentHookAmount,@FStdAllHookAmount,@FStdCurrentHookAmount,@FSCStockID,@FDCStockID,@FPeriodDate,@FCostObjGroupID,@FCostOBJID,@FMaterialCostPrice,@FReProduceType,@FBomInterID,@FDiscountRate,@FDiscountAmount,@FSepcialSaleId,@FOutCommitQty,@FOutSecCommitQty,@FDBCommitQty,@FDBSecCommitQty,@FAuxQtyInvoice,@FOperSN,@FCheckStatus,@FSplitSecQty,@FInStockID,@FSaleCommitQty,@FSaleSecCommitQty,@FSaleAuxCommitQty,@FSelectedProcID,@FVWInStockQty,@FAuxVWInStockQty,@FSecVWInStockQty,@FSecInvoiceQty,@FCostCenterID)");
                strSql.Append(";select @@IDENTITY");
                SqlParameter[] parameters = {
					new SqlParameter("@FBrNo", SqlDbType.VarChar,10),
					new SqlParameter("@FInterID", SqlDbType.Int,4),
					new SqlParameter("@FEntryID", SqlDbType.Int,4),
					new SqlParameter("@FItemID", SqlDbType.Int,4),
					new SqlParameter("@FQtyMust", SqlDbType.Decimal,13),
					new SqlParameter("@FQty", SqlDbType.Decimal,13),
					new SqlParameter("@FPrice", SqlDbType.Decimal,13),
					new SqlParameter("@FBatchNo", SqlDbType.VarChar,200),
					new SqlParameter("@FAmount", SqlDbType.Decimal,13),
					new SqlParameter("@FNote", SqlDbType.VarChar,255),
					new SqlParameter("@FSCBillInterID", SqlDbType.Int,4),
					new SqlParameter("@FSCBillNo", SqlDbType.VarChar,30),
					new SqlParameter("@FUnitID", SqlDbType.Int,4),
					new SqlParameter("@FAuxPrice", SqlDbType.Decimal,13),
					new SqlParameter("@FAuxQty", SqlDbType.Decimal,13),
					new SqlParameter("@FAuxQtyMust", SqlDbType.Decimal,13),
					new SqlParameter("@FQtyActual", SqlDbType.Decimal,13),
					new SqlParameter("@FAuxQtyActual", SqlDbType.Decimal,13),
					new SqlParameter("@FPlanPrice", SqlDbType.Decimal,13),
					new SqlParameter("@FAuxPlanPrice", SqlDbType.Decimal,13),
					new SqlParameter("@FSourceEntryID", SqlDbType.Int,4),
					new SqlParameter("@FCommitQty", SqlDbType.Decimal,13),
					new SqlParameter("@FAuxCommitQty", SqlDbType.Decimal,13),
					new SqlParameter("@FKFDate", SqlDbType.DateTime),
					new SqlParameter("@FKFPeriod", SqlDbType.Int,4),
					new SqlParameter("@FDCSPID", SqlDbType.Int,4),
					new SqlParameter("@FSCSPID", SqlDbType.Int,4),
					new SqlParameter("@FConsignPrice", SqlDbType.Decimal,13),
					new SqlParameter("@FConsignAmount", SqlDbType.Decimal,13),
					new SqlParameter("@FProcessCost", SqlDbType.Decimal,13),
					new SqlParameter("@FMaterialCost", SqlDbType.Decimal,13),
					new SqlParameter("@FTaxAmount", SqlDbType.Decimal,13),
					new SqlParameter("@FMapNumber", SqlDbType.VarChar,80),
					new SqlParameter("@FMapName", SqlDbType.NVarChar,256),
					new SqlParameter("@FOrgBillEntryID", SqlDbType.Int,4),
					new SqlParameter("@FOperID", SqlDbType.Int,4),
					new SqlParameter("@FPlanAmount", SqlDbType.Decimal,13),
					new SqlParameter("@FProcessPrice", SqlDbType.Decimal,13),
					new SqlParameter("@FTaxRate", SqlDbType.Decimal,13),
					new SqlParameter("@FSnListID", SqlDbType.Int,4),
					new SqlParameter("@FAmtRef", SqlDbType.Decimal,13),
					new SqlParameter("@FAuxPropID", SqlDbType.Int,4),
					new SqlParameter("@FCost", SqlDbType.Decimal,13),
					new SqlParameter("@FPriceRef", SqlDbType.Decimal,13),
					new SqlParameter("@FAuxPriceRef", SqlDbType.Decimal,13),
					new SqlParameter("@FFetchDate", SqlDbType.DateTime),
					new SqlParameter("@FQtyInvoice", SqlDbType.Decimal,13),
					new SqlParameter("@FQtyInvoiceBase", SqlDbType.Decimal,13),
					new SqlParameter("@FUnitCost", SqlDbType.Decimal,13),
					new SqlParameter("@FSecCoefficient", SqlDbType.Decimal,13),
					new SqlParameter("@FSecQty", SqlDbType.Decimal,13),
					new SqlParameter("@FSecCommitQty", SqlDbType.Decimal,13),
					new SqlParameter("@FSourceTranType", SqlDbType.Int,4),
					new SqlParameter("@FSourceInterId", SqlDbType.Int,4),
					new SqlParameter("@FSourceBillNo", SqlDbType.NVarChar,255),
					new SqlParameter("@FContractInterID", SqlDbType.Int,4),
					new SqlParameter("@FContractEntryID", SqlDbType.Int,4),
					new SqlParameter("@FContractBillNo", SqlDbType.NVarChar,255),
					new SqlParameter("@FICMOBillNo", SqlDbType.NVarChar,255),
					new SqlParameter("@FICMOInterID", SqlDbType.Int,4),
					new SqlParameter("@FPPBomEntryID", SqlDbType.Int,4),
					new SqlParameter("@FOrderInterID", SqlDbType.Int,4),
					new SqlParameter("@FOrderEntryID", SqlDbType.Int,4),
					new SqlParameter("@FOrderBillNo", SqlDbType.NVarChar,255),
					new SqlParameter("@FAllHookQTY", SqlDbType.Decimal,13),
					new SqlParameter("@FAllHookAmount", SqlDbType.Decimal,13),
					new SqlParameter("@FCurrentHookQTY", SqlDbType.Decimal,13),
					new SqlParameter("@FCurrentHookAmount", SqlDbType.Decimal,13),
					new SqlParameter("@FStdAllHookAmount", SqlDbType.Decimal,13),
					new SqlParameter("@FStdCurrentHookAmount", SqlDbType.Decimal,13),
					new SqlParameter("@FSCStockID", SqlDbType.Int,4),
					new SqlParameter("@FDCStockID", SqlDbType.Int,4),
					new SqlParameter("@FPeriodDate", SqlDbType.DateTime),
					new SqlParameter("@FCostObjGroupID", SqlDbType.Int,4),
					new SqlParameter("@FCostOBJID", SqlDbType.Int,4),
					new SqlParameter("@FMaterialCostPrice", SqlDbType.Decimal,13),
					new SqlParameter("@FReProduceType", SqlDbType.Int,4),
					new SqlParameter("@FBomInterID", SqlDbType.Int,4),
					new SqlParameter("@FDiscountRate", SqlDbType.Decimal,13),
					new SqlParameter("@FDiscountAmount", SqlDbType.Decimal,13),
					new SqlParameter("@FSepcialSaleId", SqlDbType.Int,4),
					new SqlParameter("@FOutCommitQty", SqlDbType.Decimal,13),
					new SqlParameter("@FOutSecCommitQty", SqlDbType.Decimal,13),
					new SqlParameter("@FDBCommitQty", SqlDbType.Decimal,13),
					new SqlParameter("@FDBSecCommitQty", SqlDbType.Decimal,13),
					new SqlParameter("@FAuxQtyInvoice", SqlDbType.Decimal,13),
					new SqlParameter("@FOperSN", SqlDbType.Int,4),
					new SqlParameter("@FCheckStatus", SqlDbType.SmallInt,2),
					new SqlParameter("@FSplitSecQty", SqlDbType.Decimal,13),
					new SqlParameter("@FInStockID", SqlDbType.Int,4),
					new SqlParameter("@FSaleCommitQty", SqlDbType.Decimal,13),
					new SqlParameter("@FSaleSecCommitQty", SqlDbType.Decimal,13),
					new SqlParameter("@FSaleAuxCommitQty", SqlDbType.Decimal,13),
					new SqlParameter("@FSelectedProcID", SqlDbType.Int,4),
					new SqlParameter("@FVWInStockQty", SqlDbType.Decimal,13),
					new SqlParameter("@FAuxVWInStockQty", SqlDbType.Decimal,13),
					new SqlParameter("@FSecVWInStockQty", SqlDbType.Decimal,13),
					new SqlParameter("@FSecInvoiceQty", SqlDbType.Decimal,13),
					new SqlParameter("@FCostCenterID", SqlDbType.Int,4)};
                parameters[0].Value = "0";
                parameters[1].Value = model.FInterID;
                parameters[2].Value = model.FEntryID;
                parameters[3].Value = model.FItemID;
                parameters[4].Value = 0;
                parameters[5].Value = model.FQty;
                parameters[6].Value = model.FPrice;
                parameters[7].Value = model.FBatchNo;
                parameters[8].Value = model.FAmount;
                parameters[9].Value = "";
                parameters[10].Value = DBNull.Value;
                parameters[11].Value = DBNull.Value;
                parameters[12].Value = model.FUnitID;
                parameters[13].Value = model.FAuxPrice;
                parameters[14].Value = model.FAuxQty;
                parameters[15].Value = 0;
                parameters[16].Value = 0;
                parameters[17].Value = 0;
                parameters[18].Value = 0;
                parameters[19].Value = 0;
                parameters[20].Value = 0;
                parameters[21].Value = 0;
                parameters[22].Value = 0;
                parameters[23].Value = DBNull.Value;
                parameters[24].Value = model.FKFPeriod;
                parameters[25].Value = 0;
                parameters[26].Value = DBNull.Value;
                parameters[27].Value = model.FConsignPrice;
                parameters[28].Value = model.FConsignAmount;
                parameters[29].Value = 0;
                parameters[30].Value = 0;
                parameters[31].Value = 0;
                parameters[32].Value = "";
                parameters[33].Value = "";
                parameters[34].Value = 0;
                parameters[35].Value = 0;
                parameters[36].Value = 0;
                parameters[37].Value = 0;
                parameters[38].Value = 0;
                parameters[39].Value = 0;
                parameters[40].Value = 0;
                parameters[41].Value = 0;
                parameters[42].Value = 0;
                parameters[43].Value = 0;
                parameters[44].Value = 0;
                parameters[45].Value = DBNull.Value;
                parameters[46].Value = 0;
                parameters[47].Value = 0;
                parameters[48].Value = 0;
                parameters[49].Value = 0;
                parameters[50].Value = 0;
                parameters[51].Value = 0;
                parameters[52].Value = 0;
                parameters[53].Value = 0;
                parameters[54].Value = "";
                parameters[55].Value = 0;
                parameters[56].Value = 0;
                parameters[57].Value = "";
                parameters[58].Value = "";
                parameters[59].Value = 0;
                parameters[60].Value = 0;
                parameters[61].Value = 0;
                parameters[62].Value = 0;
                parameters[63].Value = "";
                parameters[64].Value = 0;
                parameters[65].Value = 0;
                parameters[66].Value = 0;
                parameters[67].Value = 0;
                parameters[68].Value = 0;
                parameters[69].Value = 0;
                parameters[70].Value = 0;
                parameters[71].Value = model.FDCStockID;
                parameters[72].Value = DBNull.Value;
                parameters[73].Value = 0;
                parameters[74].Value = 0;
                parameters[75].Value = 0;
                parameters[76].Value = 0;
                parameters[77].Value = 0;
                parameters[78].Value = 0;
                parameters[79].Value = 0;
                parameters[80].Value = 0;
                parameters[81].Value = 0;
                parameters[82].Value = 0;
                parameters[83].Value = 0;
                parameters[84].Value = 0;
                parameters[85].Value = 0;
                parameters[86].Value = 0;
                parameters[87].Value = 0;
                parameters[88].Value = DBNull.Value;
                parameters[89].Value = 0;
                parameters[90].Value = 0;
                parameters[91].Value = 0;
                parameters[92].Value = 0;
                parameters[93].Value = 0;
                parameters[94].Value = 0;
                parameters[95].Value = 0;
                parameters[96].Value = 0;
                parameters[97].Value = 0;
                parameters[98].Value = 0;


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
        //public Ferrero.Model.ICStockBillEntry GetModel(string sConnectionName, int FDetailID)
        //{

        //    StringBuilder strSql = new StringBuilder();
        //    strSql.Append("select  top 1 FBrNo,FInterID,FEntryID,FItemID,FQtyMust,FQty,FPrice,FBatchNo,FAmount,FNote,FSCBillInterID,FSCBillNo,FUnitID,FAuxPrice,FAuxQty,FAuxQtyMust,FQtyActual,FAuxQtyActual,FPlanPrice,FAuxPlanPrice,FSourceEntryID,FCommitQty,FAuxCommitQty,FKFDate,FKFPeriod,FDCSPID,FSCSPID,FConsignPrice,FConsignAmount,FProcessCost,FMaterialCost,FTaxAmount,FMapNumber,FMapName,FOrgBillEntryID,FOperID,FPlanAmount,FProcessPrice,FTaxRate,FSnListID,FAmtRef,FAuxPropID,FCost,FPriceRef,FAuxPriceRef,FFetchDate,FQtyInvoice,FQtyInvoiceBase,FUnitCost,FSecCoefficient,FSecQty,FSecCommitQty,FSourceTranType,FSourceInterId,FSourceBillNo,FContractInterID,FContractEntryID,FContractBillNo,FICMOBillNo,FICMOInterID,FPPBomEntryID,FOrderInterID,FOrderEntryID,FOrderBillNo,FAllHookQTY,FAllHookAmount,FCurrentHookQTY,FCurrentHookAmount,FStdAllHookAmount,FStdCurrentHookAmount,FSCStockID,FDCStockID,FPeriodDate,FCostObjGroupID,FCostOBJID,FDetailID,FMaterialCostPrice,FReProduceType,FBomInterID,FDiscountRate,FDiscountAmount,FSepcialSaleId,FOutCommitQty,FOutSecCommitQty,FDBCommitQty,FDBSecCommitQty,FAuxQtyInvoice,FOperSN,FCheckStatus,FSplitSecQty,FInStockID,FSaleCommitQty,FSaleSecCommitQty,FSaleAuxCommitQty,FSelectedProcID,FVWInStockQty,FAuxVWInStockQty,FSecVWInStockQty,FSecInvoiceQty,FCostCenterID,FEntrySelfB0158,FEntrySelfB0160,FEntrySelfB0161,FEntrySelfB0164,FEntrySelfB0165,FEntrySelfB0166,FEntrySelfB0167,FEntrySelfB0168,FEntrySelfB0169,FEntrySelfB0171,FEntrySelfB0170,FEntrySelfB0162,FEntrySelfB0172,FEntrySelfB0173,FEntrySelfB0157,FEntrySelfB0163,FEntrySelfB0159 from ICStockBillEntry ");
        //    strSql.Append(" where FDetailID=@FDetailID");
        //    SqlParameter[] parameters = {
        //            new SqlParameter("@FDetailID", SqlDbType.Int,4)
        //    };
        //    parameters[0].Value = FDetailID;

        //    Ferrero.Model.ICStockBillEntry model = new Ferrero.Model.ICStockBillEntry();
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
        //public Ferrero.Model.ICStockBillEntry DataRowToModel(DataRow row)
        //{
        //    Ferrero.Model.ICStockBillEntry model = new Ferrero.Model.ICStockBillEntry();
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
        //        if (row["FEntryID"] != null && row["FEntryID"].ToString() != "")
        //        {
        //            model.FEntryID = int.Parse(row["FEntryID"].ToString());
        //        }
        //        if (row["FItemID"] != null && row["FItemID"].ToString() != "")
        //        {
        //            model.FItemID = int.Parse(row["FItemID"].ToString());
        //        }
        //        if (row["FQtyMust"] != null && row["FQtyMust"].ToString() != "")
        //        {
        //            model.FQtyMust = decimal.Parse(row["FQtyMust"].ToString());
        //        }
        //        if (row["FQty"] != null && row["FQty"].ToString() != "")
        //        {
        //            model.FQty = decimal.Parse(row["FQty"].ToString());
        //        }
        //        if (row["FPrice"] != null && row["FPrice"].ToString() != "")
        //        {
        //            model.FPrice = decimal.Parse(row["FPrice"].ToString());
        //        }
        //        if (row["FBatchNo"] != null)
        //        {
        //            model.FBatchNo = row["FBatchNo"].ToString();
        //        }
        //        if (row["FAmount"] != null && row["FAmount"].ToString() != "")
        //        {
        //            model.FAmount = decimal.Parse(row["FAmount"].ToString());
        //        }
        //        if (row["FNote"] != null)
        //        {
        //            model.FNote = row["FNote"].ToString();
        //        }
        //        if (row["FSCBillInterID"] != null && row["FSCBillInterID"].ToString() != "")
        //        {
        //            model.FSCBillInterID = int.Parse(row["FSCBillInterID"].ToString());
        //        }
        //        if (row["FSCBillNo"] != null)
        //        {
        //            model.FSCBillNo = row["FSCBillNo"].ToString();
        //        }
        //        if (row["FUnitID"] != null && row["FUnitID"].ToString() != "")
        //        {
        //            model.FUnitID = int.Parse(row["FUnitID"].ToString());
        //        }
        //        if (row["FAuxPrice"] != null && row["FAuxPrice"].ToString() != "")
        //        {
        //            model.FAuxPrice = decimal.Parse(row["FAuxPrice"].ToString());
        //        }
        //        if (row["FAuxQty"] != null && row["FAuxQty"].ToString() != "")
        //        {
        //            model.FAuxQty = decimal.Parse(row["FAuxQty"].ToString());
        //        }
        //        if (row["FAuxQtyMust"] != null && row["FAuxQtyMust"].ToString() != "")
        //        {
        //            model.FAuxQtyMust = decimal.Parse(row["FAuxQtyMust"].ToString());
        //        }
        //        if (row["FQtyActual"] != null && row["FQtyActual"].ToString() != "")
        //        {
        //            model.FQtyActual = decimal.Parse(row["FQtyActual"].ToString());
        //        }
        //        if (row["FAuxQtyActual"] != null && row["FAuxQtyActual"].ToString() != "")
        //        {
        //            model.FAuxQtyActual = decimal.Parse(row["FAuxQtyActual"].ToString());
        //        }
        //        if (row["FPlanPrice"] != null && row["FPlanPrice"].ToString() != "")
        //        {
        //            model.FPlanPrice = decimal.Parse(row["FPlanPrice"].ToString());
        //        }
        //        if (row["FAuxPlanPrice"] != null && row["FAuxPlanPrice"].ToString() != "")
        //        {
        //            model.FAuxPlanPrice = decimal.Parse(row["FAuxPlanPrice"].ToString());
        //        }
        //        if (row["FSourceEntryID"] != null && row["FSourceEntryID"].ToString() != "")
        //        {
        //            model.FSourceEntryID = int.Parse(row["FSourceEntryID"].ToString());
        //        }
        //        if (row["FCommitQty"] != null && row["FCommitQty"].ToString() != "")
        //        {
        //            model.FCommitQty = decimal.Parse(row["FCommitQty"].ToString());
        //        }
        //        if (row["FAuxCommitQty"] != null && row["FAuxCommitQty"].ToString() != "")
        //        {
        //            model.FAuxCommitQty = decimal.Parse(row["FAuxCommitQty"].ToString());
        //        }
        //        if (row["FKFDate"] != null && row["FKFDate"].ToString() != "")
        //        {
        //            model.FKFDate = DateTime.Parse(row["FKFDate"].ToString());
        //        }
        //        if (row["FKFPeriod"] != null && row["FKFPeriod"].ToString() != "")
        //        {
        //            model.FKFPeriod = int.Parse(row["FKFPeriod"].ToString());
        //        }
        //        if (row["FDCSPID"] != null && row["FDCSPID"].ToString() != "")
        //        {
        //            model.FDCSPID = int.Parse(row["FDCSPID"].ToString());
        //        }
        //        if (row["FSCSPID"] != null && row["FSCSPID"].ToString() != "")
        //        {
        //            model.FSCSPID = int.Parse(row["FSCSPID"].ToString());
        //        }
        //        if (row["FConsignPrice"] != null && row["FConsignPrice"].ToString() != "")
        //        {
        //            model.FConsignPrice = decimal.Parse(row["FConsignPrice"].ToString());
        //        }
        //        if (row["FConsignAmount"] != null && row["FConsignAmount"].ToString() != "")
        //        {
        //            model.FConsignAmount = decimal.Parse(row["FConsignAmount"].ToString());
        //        }
        //        if (row["FProcessCost"] != null && row["FProcessCost"].ToString() != "")
        //        {
        //            model.FProcessCost = decimal.Parse(row["FProcessCost"].ToString());
        //        }
        //        if (row["FMaterialCost"] != null && row["FMaterialCost"].ToString() != "")
        //        {
        //            model.FMaterialCost = decimal.Parse(row["FMaterialCost"].ToString());
        //        }
        //        if (row["FTaxAmount"] != null && row["FTaxAmount"].ToString() != "")
        //        {
        //            model.FTaxAmount = decimal.Parse(row["FTaxAmount"].ToString());
        //        }
        //        if (row["FMapNumber"] != null)
        //        {
        //            model.FMapNumber = row["FMapNumber"].ToString();
        //        }
        //        if (row["FMapName"] != null)
        //        {
        //            model.FMapName = row["FMapName"].ToString();
        //        }
        //        if (row["FOrgBillEntryID"] != null && row["FOrgBillEntryID"].ToString() != "")
        //        {
        //            model.FOrgBillEntryID = int.Parse(row["FOrgBillEntryID"].ToString());
        //        }
        //        if (row["FOperID"] != null && row["FOperID"].ToString() != "")
        //        {
        //            model.FOperID = int.Parse(row["FOperID"].ToString());
        //        }
        //        if (row["FPlanAmount"] != null && row["FPlanAmount"].ToString() != "")
        //        {
        //            model.FPlanAmount = decimal.Parse(row["FPlanAmount"].ToString());
        //        }
        //        if (row["FProcessPrice"] != null && row["FProcessPrice"].ToString() != "")
        //        {
        //            model.FProcessPrice = decimal.Parse(row["FProcessPrice"].ToString());
        //        }
        //        if (row["FTaxRate"] != null && row["FTaxRate"].ToString() != "")
        //        {
        //            model.FTaxRate = decimal.Parse(row["FTaxRate"].ToString());
        //        }
        //        if (row["FSnListID"] != null && row["FSnListID"].ToString() != "")
        //        {
        //            model.FSnListID = int.Parse(row["FSnListID"].ToString());
        //        }
        //        if (row["FAmtRef"] != null && row["FAmtRef"].ToString() != "")
        //        {
        //            model.FAmtRef = decimal.Parse(row["FAmtRef"].ToString());
        //        }
        //        if (row["FAuxPropID"] != null && row["FAuxPropID"].ToString() != "")
        //        {
        //            model.FAuxPropID = int.Parse(row["FAuxPropID"].ToString());
        //        }
        //        if (row["FCost"] != null && row["FCost"].ToString() != "")
        //        {
        //            model.FCost = decimal.Parse(row["FCost"].ToString());
        //        }
        //        if (row["FPriceRef"] != null && row["FPriceRef"].ToString() != "")
        //        {
        //            model.FPriceRef = decimal.Parse(row["FPriceRef"].ToString());
        //        }
        //        if (row["FAuxPriceRef"] != null && row["FAuxPriceRef"].ToString() != "")
        //        {
        //            model.FAuxPriceRef = decimal.Parse(row["FAuxPriceRef"].ToString());
        //        }
        //        if (row["FFetchDate"] != null && row["FFetchDate"].ToString() != "")
        //        {
        //            model.FFetchDate = DateTime.Parse(row["FFetchDate"].ToString());
        //        }
        //        if (row["FQtyInvoice"] != null && row["FQtyInvoice"].ToString() != "")
        //        {
        //            model.FQtyInvoice = decimal.Parse(row["FQtyInvoice"].ToString());
        //        }
        //        if (row["FQtyInvoiceBase"] != null && row["FQtyInvoiceBase"].ToString() != "")
        //        {
        //            model.FQtyInvoiceBase = decimal.Parse(row["FQtyInvoiceBase"].ToString());
        //        }
        //        if (row["FUnitCost"] != null && row["FUnitCost"].ToString() != "")
        //        {
        //            model.FUnitCost = decimal.Parse(row["FUnitCost"].ToString());
        //        }
        //        if (row["FSecCoefficient"] != null && row["FSecCoefficient"].ToString() != "")
        //        {
        //            model.FSecCoefficient = decimal.Parse(row["FSecCoefficient"].ToString());
        //        }
        //        if (row["FSecQty"] != null && row["FSecQty"].ToString() != "")
        //        {
        //            model.FSecQty = decimal.Parse(row["FSecQty"].ToString());
        //        }
        //        if (row["FSecCommitQty"] != null && row["FSecCommitQty"].ToString() != "")
        //        {
        //            model.FSecCommitQty = decimal.Parse(row["FSecCommitQty"].ToString());
        //        }
        //        if (row["FSourceTranType"] != null && row["FSourceTranType"].ToString() != "")
        //        {
        //            model.FSourceTranType = int.Parse(row["FSourceTranType"].ToString());
        //        }
        //        if (row["FSourceInterId"] != null && row["FSourceInterId"].ToString() != "")
        //        {
        //            model.FSourceInterId = int.Parse(row["FSourceInterId"].ToString());
        //        }
        //        if (row["FSourceBillNo"] != null)
        //        {
        //            model.FSourceBillNo = row["FSourceBillNo"].ToString();
        //        }
        //        if (row["FContractInterID"] != null && row["FContractInterID"].ToString() != "")
        //        {
        //            model.FContractInterID = int.Parse(row["FContractInterID"].ToString());
        //        }
        //        if (row["FContractEntryID"] != null && row["FContractEntryID"].ToString() != "")
        //        {
        //            model.FContractEntryID = int.Parse(row["FContractEntryID"].ToString());
        //        }
        //        if (row["FContractBillNo"] != null)
        //        {
        //            model.FContractBillNo = row["FContractBillNo"].ToString();
        //        }
        //        if (row["FICMOBillNo"] != null)
        //        {
        //            model.FICMOBillNo = row["FICMOBillNo"].ToString();
        //        }
        //        if (row["FICMOInterID"] != null && row["FICMOInterID"].ToString() != "")
        //        {
        //            model.FICMOInterID = int.Parse(row["FICMOInterID"].ToString());
        //        }
        //        if (row["FPPBomEntryID"] != null && row["FPPBomEntryID"].ToString() != "")
        //        {
        //            model.FPPBomEntryID = int.Parse(row["FPPBomEntryID"].ToString());
        //        }
        //        if (row["FOrderInterID"] != null && row["FOrderInterID"].ToString() != "")
        //        {
        //            model.FOrderInterID = int.Parse(row["FOrderInterID"].ToString());
        //        }
        //        if (row["FOrderEntryID"] != null && row["FOrderEntryID"].ToString() != "")
        //        {
        //            model.FOrderEntryID = int.Parse(row["FOrderEntryID"].ToString());
        //        }
        //        if (row["FOrderBillNo"] != null)
        //        {
        //            model.FOrderBillNo = row["FOrderBillNo"].ToString();
        //        }
        //        if (row["FAllHookQTY"] != null && row["FAllHookQTY"].ToString() != "")
        //        {
        //            model.FAllHookQTY = decimal.Parse(row["FAllHookQTY"].ToString());
        //        }
        //        if (row["FAllHookAmount"] != null && row["FAllHookAmount"].ToString() != "")
        //        {
        //            model.FAllHookAmount = decimal.Parse(row["FAllHookAmount"].ToString());
        //        }
        //        if (row["FCurrentHookQTY"] != null && row["FCurrentHookQTY"].ToString() != "")
        //        {
        //            model.FCurrentHookQTY = decimal.Parse(row["FCurrentHookQTY"].ToString());
        //        }
        //        if (row["FCurrentHookAmount"] != null && row["FCurrentHookAmount"].ToString() != "")
        //        {
        //            model.FCurrentHookAmount = decimal.Parse(row["FCurrentHookAmount"].ToString());
        //        }
        //        if (row["FStdAllHookAmount"] != null && row["FStdAllHookAmount"].ToString() != "")
        //        {
        //            model.FStdAllHookAmount = decimal.Parse(row["FStdAllHookAmount"].ToString());
        //        }
        //        if (row["FStdCurrentHookAmount"] != null && row["FStdCurrentHookAmount"].ToString() != "")
        //        {
        //            model.FStdCurrentHookAmount = decimal.Parse(row["FStdCurrentHookAmount"].ToString());
        //        }
        //        if (row["FSCStockID"] != null && row["FSCStockID"].ToString() != "")
        //        {
        //            model.FSCStockID = int.Parse(row["FSCStockID"].ToString());
        //        }
        //        if (row["FDCStockID"] != null && row["FDCStockID"].ToString() != "")
        //        {
        //            model.FDCStockID = int.Parse(row["FDCStockID"].ToString());
        //        }
        //        if (row["FPeriodDate"] != null && row["FPeriodDate"].ToString() != "")
        //        {
        //            model.FPeriodDate = DateTime.Parse(row["FPeriodDate"].ToString());
        //        }
        //        if (row["FCostObjGroupID"] != null && row["FCostObjGroupID"].ToString() != "")
        //        {
        //            model.FCostObjGroupID = int.Parse(row["FCostObjGroupID"].ToString());
        //        }
        //        if (row["FCostOBJID"] != null && row["FCostOBJID"].ToString() != "")
        //        {
        //            model.FCostOBJID = int.Parse(row["FCostOBJID"].ToString());
        //        }
        //        if (row["FDetailID"] != null && row["FDetailID"].ToString() != "")
        //        {
        //            model.FDetailID = int.Parse(row["FDetailID"].ToString());
        //        }
        //        if (row["FMaterialCostPrice"] != null && row["FMaterialCostPrice"].ToString() != "")
        //        {
        //            model.FMaterialCostPrice = decimal.Parse(row["FMaterialCostPrice"].ToString());
        //        }
        //        if (row["FReProduceType"] != null && row["FReProduceType"].ToString() != "")
        //        {
        //            model.FReProduceType = int.Parse(row["FReProduceType"].ToString());
        //        }
        //        if (row["FBomInterID"] != null && row["FBomInterID"].ToString() != "")
        //        {
        //            model.FBomInterID = int.Parse(row["FBomInterID"].ToString());
        //        }
        //        if (row["FDiscountRate"] != null && row["FDiscountRate"].ToString() != "")
        //        {
        //            model.FDiscountRate = decimal.Parse(row["FDiscountRate"].ToString());
        //        }
        //        if (row["FDiscountAmount"] != null && row["FDiscountAmount"].ToString() != "")
        //        {
        //            model.FDiscountAmount = decimal.Parse(row["FDiscountAmount"].ToString());
        //        }
        //        if (row["FSepcialSaleId"] != null && row["FSepcialSaleId"].ToString() != "")
        //        {
        //            model.FSepcialSaleId = int.Parse(row["FSepcialSaleId"].ToString());
        //        }
        //        if (row["FOutCommitQty"] != null && row["FOutCommitQty"].ToString() != "")
        //        {
        //            model.FOutCommitQty = decimal.Parse(row["FOutCommitQty"].ToString());
        //        }
        //        if (row["FOutSecCommitQty"] != null && row["FOutSecCommitQty"].ToString() != "")
        //        {
        //            model.FOutSecCommitQty = decimal.Parse(row["FOutSecCommitQty"].ToString());
        //        }
        //        if (row["FDBCommitQty"] != null && row["FDBCommitQty"].ToString() != "")
        //        {
        //            model.FDBCommitQty = decimal.Parse(row["FDBCommitQty"].ToString());
        //        }
        //        if (row["FDBSecCommitQty"] != null && row["FDBSecCommitQty"].ToString() != "")
        //        {
        //            model.FDBSecCommitQty = decimal.Parse(row["FDBSecCommitQty"].ToString());
        //        }
        //        if (row["FAuxQtyInvoice"] != null && row["FAuxQtyInvoice"].ToString() != "")
        //        {
        //            model.FAuxQtyInvoice = decimal.Parse(row["FAuxQtyInvoice"].ToString());
        //        }
        //        if (row["FOperSN"] != null && row["FOperSN"].ToString() != "")
        //        {
        //            model.FOperSN = int.Parse(row["FOperSN"].ToString());
        //        }
        //        if (row["FCheckStatus"] != null && row["FCheckStatus"].ToString() != "")
        //        {
        //            model.FCheckStatus = int.Parse(row["FCheckStatus"].ToString());
        //        }
        //        if (row["FSplitSecQty"] != null && row["FSplitSecQty"].ToString() != "")
        //        {
        //            model.FSplitSecQty = decimal.Parse(row["FSplitSecQty"].ToString());
        //        }
        //        if (row["FInStockID"] != null && row["FInStockID"].ToString() != "")
        //        {
        //            model.FInStockID = int.Parse(row["FInStockID"].ToString());
        //        }
        //        if (row["FSaleCommitQty"] != null && row["FSaleCommitQty"].ToString() != "")
        //        {
        //            model.FSaleCommitQty = decimal.Parse(row["FSaleCommitQty"].ToString());
        //        }
        //        if (row["FSaleSecCommitQty"] != null && row["FSaleSecCommitQty"].ToString() != "")
        //        {
        //            model.FSaleSecCommitQty = decimal.Parse(row["FSaleSecCommitQty"].ToString());
        //        }
        //        if (row["FSaleAuxCommitQty"] != null && row["FSaleAuxCommitQty"].ToString() != "")
        //        {
        //            model.FSaleAuxCommitQty = decimal.Parse(row["FSaleAuxCommitQty"].ToString());
        //        }
        //        if (row["FSelectedProcID"] != null && row["FSelectedProcID"].ToString() != "")
        //        {
        //            model.FSelectedProcID = int.Parse(row["FSelectedProcID"].ToString());
        //        }
        //        if (row["FVWInStockQty"] != null && row["FVWInStockQty"].ToString() != "")
        //        {
        //            model.FVWInStockQty = decimal.Parse(row["FVWInStockQty"].ToString());
        //        }
        //        if (row["FAuxVWInStockQty"] != null && row["FAuxVWInStockQty"].ToString() != "")
        //        {
        //            model.FAuxVWInStockQty = decimal.Parse(row["FAuxVWInStockQty"].ToString());
        //        }
        //        if (row["FSecVWInStockQty"] != null && row["FSecVWInStockQty"].ToString() != "")
        //        {
        //            model.FSecVWInStockQty = decimal.Parse(row["FSecVWInStockQty"].ToString());
        //        }
        //        if (row["FSecInvoiceQty"] != null && row["FSecInvoiceQty"].ToString() != "")
        //        {
        //            model.FSecInvoiceQty = decimal.Parse(row["FSecInvoiceQty"].ToString());
        //        }
        //        if (row["FCostCenterID"] != null && row["FCostCenterID"].ToString() != "")
        //        {
        //            model.FCostCenterID = int.Parse(row["FCostCenterID"].ToString());
        //        }
        //        if (row["FEntrySelfB0158"] != null)
        //        {
        //            model.FEntrySelfB0158 = row["FEntrySelfB0158"].ToString();
        //        }
        //        if (row["FEntrySelfB0160"] != null)
        //        {
        //            model.FEntrySelfB0160 = row["FEntrySelfB0160"].ToString();
        //        }
        //        if (row["FEntrySelfB0161"] != null)
        //        {
        //            model.FEntrySelfB0161 = row["FEntrySelfB0161"].ToString();
        //        }
        //        if (row["FEntrySelfB0164"] != null)
        //        {
        //            model.FEntrySelfB0164 = row["FEntrySelfB0164"].ToString();
        //        }
        //        if (row["FEntrySelfB0165"] != null)
        //        {
        //            model.FEntrySelfB0165 = row["FEntrySelfB0165"].ToString();
        //        }
        //        if (row["FEntrySelfB0166"] != null)
        //        {
        //            model.FEntrySelfB0166 = row["FEntrySelfB0166"].ToString();
        //        }
        //        if (row["FEntrySelfB0167"] != null)
        //        {
        //            model.FEntrySelfB0167 = row["FEntrySelfB0167"].ToString();
        //        }
        //        if (row["FEntrySelfB0168"] != null)
        //        {
        //            model.FEntrySelfB0168 = row["FEntrySelfB0168"].ToString();
        //        }
        //        if (row["FEntrySelfB0169"] != null)
        //        {
        //            model.FEntrySelfB0169 = row["FEntrySelfB0169"].ToString();
        //        }
        //        if (row["FEntrySelfB0171"] != null)
        //        {
        //            model.FEntrySelfB0171 = row["FEntrySelfB0171"].ToString();
        //        }
        //        if (row["FEntrySelfB0170"] != null && row["FEntrySelfB0170"].ToString() != "")
        //        {
        //            model.FEntrySelfB0170 = decimal.Parse(row["FEntrySelfB0170"].ToString());
        //        }
        //        if (row["FEntrySelfB0162"] != null && row["FEntrySelfB0162"].ToString() != "")
        //        {
        //            model.FEntrySelfB0162 = decimal.Parse(row["FEntrySelfB0162"].ToString());
        //        }
        //        if (row["FEntrySelfB0172"] != null && row["FEntrySelfB0172"].ToString() != "")
        //        {
        //            model.FEntrySelfB0172 = decimal.Parse(row["FEntrySelfB0172"].ToString());
        //        }
        //        if (row["FEntrySelfB0173"] != null && row["FEntrySelfB0173"].ToString() != "")
        //        {
        //            model.FEntrySelfB0173 = decimal.Parse(row["FEntrySelfB0173"].ToString());
        //        }
        //        if (row["FEntrySelfB0157"] != null && row["FEntrySelfB0157"].ToString() != "")
        //        {
        //            model.FEntrySelfB0157 = decimal.Parse(row["FEntrySelfB0157"].ToString());
        //        }
        //        if (row["FEntrySelfB0163"] != null && row["FEntrySelfB0163"].ToString() != "")
        //        {
        //            model.FEntrySelfB0163 = decimal.Parse(row["FEntrySelfB0163"].ToString());
        //        }
        //        if (row["FEntrySelfB0159"] != null && row["FEntrySelfB0159"].ToString() != "")
        //        {
        //            model.FEntrySelfB0159 = decimal.Parse(row["FEntrySelfB0159"].ToString());
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
        //    strSql.Append("select FBrNo,FInterID,FEntryID,FItemID,FQtyMust,FQty,FPrice,FBatchNo,FAmount,FNote,FSCBillInterID,FSCBillNo,FUnitID,FAuxPrice,FAuxQty,FAuxQtyMust,FQtyActual,FAuxQtyActual,FPlanPrice,FAuxPlanPrice,FSourceEntryID,FCommitQty,FAuxCommitQty,FKFDate,FKFPeriod,FDCSPID,FSCSPID,FConsignPrice,FConsignAmount,FProcessCost,FMaterialCost,FTaxAmount,FMapNumber,FMapName,FOrgBillEntryID,FOperID,FPlanAmount,FProcessPrice,FTaxRate,FSnListID,FAmtRef,FAuxPropID,FCost,FPriceRef,FAuxPriceRef,FFetchDate,FQtyInvoice,FQtyInvoiceBase,FUnitCost,FSecCoefficient,FSecQty,FSecCommitQty,FSourceTranType,FSourceInterId,FSourceBillNo,FContractInterID,FContractEntryID,FContractBillNo,FICMOBillNo,FICMOInterID,FPPBomEntryID,FOrderInterID,FOrderEntryID,FOrderBillNo,FAllHookQTY,FAllHookAmount,FCurrentHookQTY,FCurrentHookAmount,FStdAllHookAmount,FStdCurrentHookAmount,FSCStockID,FDCStockID,FPeriodDate,FCostObjGroupID,FCostOBJID,FDetailID,FMaterialCostPrice,FReProduceType,FBomInterID,FDiscountRate,FDiscountAmount,FSepcialSaleId,FOutCommitQty,FOutSecCommitQty,FDBCommitQty,FDBSecCommitQty,FAuxQtyInvoice,FOperSN,FCheckStatus,FSplitSecQty,FInStockID,FSaleCommitQty,FSaleSecCommitQty,FSaleAuxCommitQty,FSelectedProcID,FVWInStockQty,FAuxVWInStockQty,FSecVWInStockQty,FSecInvoiceQty,FCostCenterID,FEntrySelfB0158,FEntrySelfB0160,FEntrySelfB0161,FEntrySelfB0164,FEntrySelfB0165,FEntrySelfB0166,FEntrySelfB0167,FEntrySelfB0168,FEntrySelfB0169,FEntrySelfB0171,FEntrySelfB0170,FEntrySelfB0162,FEntrySelfB0172,FEntrySelfB0173,FEntrySelfB0157,FEntrySelfB0163,FEntrySelfB0159 ");
        //    strSql.Append(" FROM ICStockBillEntry ");
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
        //    strSql.Append(" FBrNo,FInterID,FEntryID,FItemID,FQtyMust,FQty,FPrice,FBatchNo,FAmount,FNote,FSCBillInterID,FSCBillNo,FUnitID,FAuxPrice,FAuxQty,FAuxQtyMust,FQtyActual,FAuxQtyActual,FPlanPrice,FAuxPlanPrice,FSourceEntryID,FCommitQty,FAuxCommitQty,FKFDate,FKFPeriod,FDCSPID,FSCSPID,FConsignPrice,FConsignAmount,FProcessCost,FMaterialCost,FTaxAmount,FMapNumber,FMapName,FOrgBillEntryID,FOperID,FPlanAmount,FProcessPrice,FTaxRate,FSnListID,FAmtRef,FAuxPropID,FCost,FPriceRef,FAuxPriceRef,FFetchDate,FQtyInvoice,FQtyInvoiceBase,FUnitCost,FSecCoefficient,FSecQty,FSecCommitQty,FSourceTranType,FSourceInterId,FSourceBillNo,FContractInterID,FContractEntryID,FContractBillNo,FICMOBillNo,FICMOInterID,FPPBomEntryID,FOrderInterID,FOrderEntryID,FOrderBillNo,FAllHookQTY,FAllHookAmount,FCurrentHookQTY,FCurrentHookAmount,FStdAllHookAmount,FStdCurrentHookAmount,FSCStockID,FDCStockID,FPeriodDate,FCostObjGroupID,FCostOBJID,FDetailID,FMaterialCostPrice,FReProduceType,FBomInterID,FDiscountRate,FDiscountAmount,FSepcialSaleId,FOutCommitQty,FOutSecCommitQty,FDBCommitQty,FDBSecCommitQty,FAuxQtyInvoice,FOperSN,FCheckStatus,FSplitSecQty,FInStockID,FSaleCommitQty,FSaleSecCommitQty,FSaleAuxCommitQty,FSelectedProcID,FVWInStockQty,FAuxVWInStockQty,FSecVWInStockQty,FSecInvoiceQty,FCostCenterID,FEntrySelfB0158,FEntrySelfB0160,FEntrySelfB0161,FEntrySelfB0164,FEntrySelfB0165,FEntrySelfB0166,FEntrySelfB0167,FEntrySelfB0168,FEntrySelfB0169,FEntrySelfB0171,FEntrySelfB0170,FEntrySelfB0162,FEntrySelfB0172,FEntrySelfB0173,FEntrySelfB0157,FEntrySelfB0163,FEntrySelfB0159 ");
        //    strSql.Append(" FROM ICStockBillEntry ");
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
        //    strSql.Append("select count(1) FROM ICStockBillEntry ");
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
        //        strSql.Append("order by T.FDetailID desc");
        //    }
        //    strSql.Append(")AS Row, T.*  from ICStockBillEntry T ");
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
            parameters[0].Value = "ICStockBillEntry";
            parameters[1].Value = "FDetailID";
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
        /// 通过商品代码得到商品编号
        /// </summary>
        /// <param name="connectionName"></param>
        /// <param name="fNumber"></param>
        /// <returns></returns>
        public int GetfItemId(string connectionName, string fNumber)
        {
            //string connectionString = EncryptHelper.Decrypt("77052300", ConfigurationManager.ConnectionStrings[connectionName].ConnectionString);
            StringBuilder strSql = new StringBuilder();
            strSql.Append(String.Format(" select fItemID from t_ICItemCore Where fNumber = '{0}' ", fNumber));
            object obj = SqlHelper.GetSingle(SqlHelper.GetConnectionString(connectionName), strSql.ToString(), null);
            if (obj == null)
            {
                return 0;
            }
            else
            {
                return Convert.ToInt32(obj);
            }
        }

        /// <summary>
        /// 通过单位名称得到单位编号
        /// </summary>
        /// <param name="connectionName"></param>
        /// <param name="fName"></param>
        /// <returns></returns>
        public int GetUnitID(string connectionName, string fName)
        {
            //string connectionString = EncryptHelper.Decrypt("77052300", ConfigurationManager.ConnectionStrings[connectionName].ConnectionString);
            StringBuilder strSql = new StringBuilder();
            strSql.Append(String.Format(" select fItemID from t_MeasureUnit Where FName = '{0}' ", fName));
            object obj = SqlHelper.GetSingle(SqlHelper.GetConnectionString(connectionName), strSql.ToString(), null);
            if (obj == null)
            {
                return 0;
            }
            else
            {
                return Convert.ToInt32(obj);
            }
        }

        /// <summary>
        /// 通过供应商代码得到供应商编号
        /// </summary>
        /// <param name="connectionName"></param>
        /// <param name="fNumber"></param>
        /// <returns></returns>
        public int GetSupplyID(string connectionName, string fNumber)
        {
            //string connectionString = EncryptHelper.Decrypt("77052300", ConfigurationManager.ConnectionStrings[connectionName].ConnectionString);
            StringBuilder strSql = new StringBuilder();
            strSql.Append(String.Format(" select fItemID from t_item Where FNumber = '{0}' ", fNumber));
            object obj = SqlHelper.GetSingle(SqlHelper.GetConnectionString(connectionName), strSql.ToString(), null);
            if (obj == null)
            {
                return 0;
            }
            else
            {
                return Convert.ToInt32(obj);
            }
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
            //string connectionString = EncryptHelper.Decrypt("77052300", ConfigurationManager.ConnectionStrings[connectionName].ConnectionString);
            StringBuilder strSql = new StringBuilder();
            strSql.Append(String.Format(" select count(1) from t_item Where FNumber = '{0}'  and FName = '{1}' ", fNumber, fName));
            object obj = SqlHelper.GetSingle(SqlHelper.GetConnectionString(connectionName), strSql.ToString(), null);
            if (obj == null)
            {
                return false;
            }
            else if (Convert.ToInt32(obj) == 0)
            {
                return false;
            }
            else
            {
                return true;
            }
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
            //string connectionString = EncryptHelper.Decrypt("77052300", ConfigurationManager.ConnectionStrings[connectionName].ConnectionString);
            StringBuilder strSql = new StringBuilder();
            strSql.Append(String.Format(" select count(1) from t_ICItem Where FNumber = '{0}'  and FName = '{1}' ", fNumber, fName));
            object obj = SqlHelper.GetSingle(SqlHelper.GetConnectionString(connectionName), strSql.ToString(), null);
            if (obj == null)
            {
                return false;
            }
            else if (Convert.ToInt32(obj) == 0)
            {
                return false;
            }
            else
            {
                return true;
            }
        }

        /// <summary>
        /// 读取SKU表中的大客户单价
        /// </summary>
        /// <param name="connectionName">分公司数据库</param>
        /// <param name="sLongCode">产品长代码</param>
        /// <returns></returns>
        public decimal getSpecialUnitPriceByLongCode(string connectionName, string sLongCode)
        {
            StringBuilder strSql = new StringBuilder();
            strSql.Append(String.Format(" Select UnitPrice from sku Where ProductCode ='{0}'", sLongCode));
            object obj = SqlHelper.GetSingle(SqlHelper.GetConnectionString(connectionName), strSql.ToString(), null);
            if (obj == null)
            {
                return 0;
            }
            else
            {
                return Decimal.Parse(obj.ToString());
            }
        }
        #endregion  ExtensionMethod
    }
}


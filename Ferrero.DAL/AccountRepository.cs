using System.Data;
using System.Text;
using System.Data.SqlClient;

using Ray.Framework.DBUtility;

namespace Ferrero.DAL
{
    /// <summary>
    /// 数据访问类:t_User
    /// </summary>
    public partial class AccountRepository
    {
        #region  BasicMethod

        #endregion  BasicMethod

        #region  ExtensionMethod
        /// <summary>
        /// 获得数据列表
        /// </summary>
        public int UserLogin(string sConnectionName, string userName, string password)
        {
            ///PassService pass = new PassService();
            if (!string.IsNullOrEmpty(sConnectionName))
            {
                //string connectionString = EncryptHelper.Decrypt("77052300",ConfigurationManager.ConnectionStrings[connectionName].ConnectionString);
                StringBuilder strSql = new StringBuilder();
                strSql.Append("select Count(1) ");
                strSql.Append(" FROM t_User ");
                strSql.Append(" where t_User.FName = @FName and t_User.FSID = @FSID ");

                SqlParameter[] parameters = {
				new SqlParameter("@FName", SqlDbType.NVarChar , 255),
				new SqlParameter("@FSID", SqlDbType.NVarChar , 255)
				};
                parameters[0].Value = userName;
                parameters[1].Value = password;

                object obj = SqlHelper.ExecuteScalar(SqlHelper.GetConnectionString(sConnectionName), CommandType.Text, strSql.ToString(), parameters);
                return obj == null ? 0 : int.Parse(obj.ToString());
            }
            else
            {
                return 0;
            }

        }
        #endregion  ExtensionMethod
    }
}


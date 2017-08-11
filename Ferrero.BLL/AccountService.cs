using System;
using System.Data;
using System.Collections.Generic;

namespace Ferrero.BLL
{
    /// <summary>
    /// t_User
    /// </summary>
    public partial class AccountService
    {
        private readonly Ferrero.DAL.AccountRepository  dal = new Ferrero.DAL.AccountRepository ();

        #region  BasicMethod

        #endregion  BasicMethod

        #region  ExtensionMethod
        /// <summary>
        /// 获得数据列表
        /// </summary>
        public bool UserLogin(string sConnectionName, string fName, string FPassword)
        {
            return dal.UserLogin(sConnectionName, fName, FPassword) > 0 ? true : false;
        }

        #endregion  ExtensionMethod
    }
}
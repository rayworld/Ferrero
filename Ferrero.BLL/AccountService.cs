﻿using System;
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
        public bool UserLogin(string sConnectionString, string fName, string FPassword)
        {
            return dal.UserLogin(sConnectionString, fName, FPassword) > 0 ? true : false;
            //return true;
        }

        #endregion  ExtensionMethod
    }
}
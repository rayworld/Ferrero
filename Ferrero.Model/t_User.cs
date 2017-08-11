using System;

namespace Ferrero.Model
{
    /// <summary>
    /// t_User:实体类(属性说明自动提取数据库字段的描述信息)
    /// </summary>
    [Serializable]
    public class t_User
    {
        #region Model
        private int _fuserid;
        private string _fusername;
        private string _fpassword;
        private int _froleid;
        private int _fdeptid;
        private int _ftitleid;
        /// <summary>
        /// 
        /// </summary>
        public int FUserId
        {
            set { _fuserid = value; }
            get { return _fuserid; }
        }
        /// <summary>
        /// 
        /// </summary>
        public string FUserName
        {
            set { _fusername = value; }
            get { return _fusername; }
        }
        /// <summary>
        /// 
        /// </summary>
        public string FPassword
        {
            set { _fpassword = value; }
            get { return _fpassword; }
        }
        /// <summary>
        /// 
        /// </summary>
        public int FRoleId
        {
            set { _froleid = value; }
            get { return _froleid; }
        }
        /// <summary>
        /// 
        /// </summary>
        public int FDeptId
        {
            set { _fdeptid = value; }
            get { return _fdeptid; }
        }
        /// <summary>
        /// 
        /// </summary>
        public int FTitleId
        {
            set { _ftitleid = value; }
            get { return _ftitleid; }
        }
        #endregion Model
    }
}

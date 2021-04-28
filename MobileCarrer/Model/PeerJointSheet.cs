using System;
using System.Collections.Generic;
using System.Text;

namespace MobileCarrer.Model
{
    public class PeerJointSheet
    {
        public List<PeerJointSheetOne> SheetOne { get; set; }
        public List<PeerJointSheetTwo> SheetTwo { get; set; }
    }


    /// <summary>
    /// 共評表一格式
    /// </summary>
    public class PeerJointSheetOne
    {
        public string DateTime { get; set; }
        /// <summary>
        /// 學號
        /// </summary>
        public string StudentID { get; set; }
        /// <summary>
        /// 姓名
        /// </summary>
        public string StudentName { get; set; }
        /// <summary>
        /// 組別一
        /// </summary>
        public string FoodPanda { get; set; }
        /// <summary>
        /// 組別二
        /// </summary>
        public string FamilyMart { get; set; }
        /// <summary>
        /// 組別三
        /// </summary>
        public string Dcard { get; set; }
        /// <summary>
        /// 組別四
        /// </summary>
        public string Uniqlo { get; set; }
        /// <summary>
        /// 提問
        /// </summary>
        public string Question { get; set; }
        ///// <summary>
        ///// 是否有提問
        ///// </summary>
        //public bool HasQuestion { get; set; }
    }

    /// <summary>
    /// 共評表二格式
    /// </summary>
    public class PeerJointSheetTwo
    {
        public string DateTime { get; set; }
        /// <summary>
        /// 學號
        /// </summary>
        public string StudentID { get; set; }
        /// <summary>
        /// 姓名
        /// </summary>
        public string StudentName { get; set; }
        /// <summary>
        /// 組別一
        /// </summary>
        public string Shopee { get; set; }
        /// <summary>
        /// 組別二
        /// </summary>
        public string PxPay { get; set; }
        /// <summary>
        /// 組別三
        /// </summary>
        public string Tinder { get; set; }
        /// <summary>
        /// 組別四
        /// </summary>
        public string LanguageTrails { get; set; }

        /// <summary>
        /// 組別五
        /// </summary>
        public string PostOffice { get; set; }
        /// <summary>
        /// 組別六
        /// </summary>
        public string OpenPoint { get; set; }
        /// <summary>
        /// 組別七
        /// </summary>
        public string Spotify { get; set; }
        /// <summary>
        /// 提問
        /// </summary>
        public string Question { get; set; }
        ///// <summary>
        ///// 是否有提問
        ///// </summary>
        //public bool HasQuestion { get; set; }
    }
}

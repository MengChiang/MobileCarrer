using System;
using System.Collections.Generic;
using System.Text;

namespace MobileCarrer.Model
{
    public class GroupInformation
    {
        /// <summary>
        /// 組別
        /// </summary>
        public string GroupName { get; set; }



        /// <summary>
        /// 組員1
        /// </summary>
        public string Member1 { get; set; }
        /// <summary>
        /// 組員2
        /// </summary>
        public string Member2 { get; set; }
        /// <summary>
        /// 組員3
        /// </summary>
        public string Member3 { get; set; }
        /// <summary>
        /// 組員4
        /// </summary>
        public string Member4 { get; set; }
        /// <summary>
        /// 組員5
        /// </summary>
        public string Member5 { get; set; }


    }

    public class GroupWithGrade : GroupInformation
    {
        /// <summary>
        /// 組別分數
        /// </summary>
        public string Score { get; set; }
        /// <summary>
        /// 教師評分
        /// </summary>
        public string TeachScore { get; set; }
        /// <summary>
        /// 所有評分(除教師以外)
        /// </summary>
        public Dictionary<string, string> Records { get; set; }
    }
}

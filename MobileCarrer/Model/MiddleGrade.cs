using System;
using System.Collections.Generic;
using System.Text;

namespace MobileCarrer.Model
{
    public class MiddleGrade
    {
        /// <summary>
        /// 總分
        /// </summary>
        public int Total { get; set; }
        /// <summary>
        /// 同儕評分
        /// </summary>
        public int PeerScore { get; set; }
        /// <summary>
        /// 自我評分
        /// </summary>
        public int SelfScore { get; set; }
        /// <summary>
        /// 教師評分
        /// </summary>
        public int TeacherScore { get; set; }
    }
}

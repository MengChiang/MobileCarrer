using System;
using System.Collections.Generic;
using System.Text;

namespace MobileCarrer.Model
{
    public class StudentScore
    {
        /// <summary>
        /// 學生姓名
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// 學號
        /// </summary>
        public string ID { get; set; }
        /// <summary>
        /// 期中成績
        /// </summary>
        public MiddleGrade MidGrade { get; set; }
        //public int MidGrade { get; set; }
        /// <summary>
        /// 期末成績
        /// </summary>
        public int FinalGrade { get; set; }
        /// <summary>
        /// 平時成績
        /// </summary>
        public int UsualGrade { get; set; }
        /// <summary>
        /// 個人加減分
        /// </summary>
        public int Bonus { get; set; }
    }
}

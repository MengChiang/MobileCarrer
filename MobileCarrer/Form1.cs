using CsvHelper;
using MobileCarrer.Model;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Newtonsoft.Json;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace MobileCarrer
{
    public partial class Form1 : Form
    {
        string pathOfFinalStudentList = string.Empty;
        string pathOfJointSheetOne = string.Empty;
        string pathOfJointSheetTwo = string.Empty;
        string pathOfGroupInformation = string.Empty;
        string pathOfExcelOutput = string.Empty;
        List<string> noSelfScoreGuys = new List<string>() { "游芳柔", "劉譯多", "張巧儒" };

        public Form1()
        {
            InitializeComponent();

            var projectPath = GetProjectParentsPath();
            pathOfJointSheetOne = Path.Combine(projectPath, "Data", "sheet1.csv");
            pathOfJointSheetTwo = Path.Combine(projectPath, "Data", "sheet2.csv");
            pathOfGroupInformation = Path.Combine(projectPath, "Data", "goupInfo.csv");
            pathOfFinalStudentList = Path.Combine(projectPath, "Data", "studentlist2.csv");
            pathOfExcelOutput = Path.Combine(projectPath, "Data", "output.xlsx");

            sheet1TextBox.Text = pathOfJointSheetOne;
            sheet2TextBox.Text = pathOfJointSheetTwo;
            groupInfoTextBox.Text = pathOfGroupInformation;
            outputSheetTextBox.Text = pathOfExcelOutput;

            this.openFileDialog1.InitialDirectory = projectPath;
        }

        /// <summary>
        /// 取得上兩層層資料夾路徑
        /// </summary>
        private string GetProjectParentsPath()
        {
            var directoryInfo = new DirectoryInfo(System.Windows.Forms.Application.StartupPath);
            return directoryInfo.Parent.Parent.Parent.FullName;
        }

        /// <summary>
        /// 共評表一的載入按鈕
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void sheet1Button_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.InitialDirectory = sheet2TextBox.Text;
                openFileDialog.Filter = "csv files (*.csv)|*.txt|All files (*.*)|*.*";
                openFileDialog.FilterIndex = 2;
                openFileDialog.RestoreDirectory = true;

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    sheet2TextBox.Text = openFileDialog.FileName;
                }
            }
        }

        /// <summary>
        /// 共評表二的載入按鈕
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void sheet2Button_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.InitialDirectory = sheet2TextBox.Text;
                openFileDialog.Filter = "csv files (*.csv)|*.txt|All files (*.*)|*.*";
                openFileDialog.FilterIndex = 2;
                openFileDialog.RestoreDirectory = true;

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    sheet2TextBox.Text = openFileDialog.FileName;
                }
            }
        }

        /// <summary>
        /// 組員資訊表的載入按鈕
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void sheet3Button_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.InitialDirectory = groupInfoTextBox.Text;
                openFileDialog.Filter = "csv files (*.csv)|*.txt|All files (*.*)|*.*";
                openFileDialog.FilterIndex = 2;
                openFileDialog.RestoreDirectory = true;

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    groupInfoTextBox.Text = openFileDialog.FileName;
                }
            }
        }


        /// <summary>
        /// 輸出路徑的載入按鈕
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void outputButton_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.InitialDirectory = outputSheetTextBox.Text;
                openFileDialog.Filter = "xlsx files (*.xlsx)|*.txt|All files (*.*)|*.*";
                openFileDialog.FilterIndex = 2;
                openFileDialog.RestoreDirectory = true;

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    outputSheetTextBox.Text = openFileDialog.FileName;
                }
            }
        }

        /// <summary>
        /// 按下計算期中成績
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void middleGradeButton_Click(object sender, EventArgs e)
        {
            #region 檢查檔案是否存在
            //檢查共評表一是否存在
            DoesSheet1Exist();

            //檢查共評表二是否存在
            DoesSheet2Exist();

            //檢查組員資訊表是否存在
            DoesGroupInfoExist();

            //檢查輸出檔案是否存在
            DoesOutputSheetExist();
            #endregion

            //取得最終選課清單
            var studentList = GetStudentList();

            //取得組別資料
            var groupInformation = GetGroupInformation();

            //取得共評表相關資料
            var jointSheetData = CalculateJointSheets();

            //取得成績、組別及組員資料
            var groupWithGrades = GetGroupWithGradeInformation(groupInformation, jointSheetData);


            //讀取每個學生依序計算成績
            foreach (var student in studentList)
            {
                student.MidGrade = new MiddleGrade();

                //取得屬於學生位於哪一組的資料
                var groupWithGrade = groupWithGrades.Find(x => x.Member1 == student.Name || x.Member2 == student.Name ||
                x.Member3 == student.Name || x.Member4 == student.Name || x.Member5 == student.Name);

                if (groupWithGrade != null && groupWithGrade.GroupName != "Shopee")
                {
                    student.MidGrade.TeacherScore = ConvertGradeToScore(groupWithGrade.TeachScore);

                    if (groupWithGrade.Records.ContainsKey(student.Name))
                    {
                        student.MidGrade.SelfScore = ConvertGradeToScore(groupWithGrade.Records[student.Name]);
                    }
                    else if (noSelfScoreGuys.Contains(student.Name))
                    {
                        student.MidGrade.SelfScore = ConvertGradeToScore("A");
                    }

                    groupWithGrade.Records.Remove(student.Name);

                    student.MidGrade.PeerScore = (int)Math.Round(groupWithGrade.Records.Values.Where(w => w != null && w.Length > 0).Select(r => ConvertGradeToScore(r)).Average());
                }
                else
                {
                    if (student.Name.IndexOf("羅") >= 0 || student.Name.IndexOf("瑋") >= 0)
                    {
                        groupWithGrade = groupWithGrades.Find(f => f.GroupName == "FoodPanda");

                        student.MidGrade.TeacherScore = ConvertGradeToScore(groupWithGrade.TeachScore);
                        student.MidGrade.SelfScore = ConvertGradeToScore("A");
                        groupWithGrade.Records.Remove(student.Name);
                        student.MidGrade.PeerScore = (int)Math.Round(groupWithGrade.Records.Values.Where(w => w != null && w.Length > 0).Select(r => ConvertGradeToScore(r)).Average());
                    }
                }

                student.MidGrade.Total = (student.MidGrade.SelfScore * 10 / 30) + (student.MidGrade.PeerScore * 15 / 30) + (student.MidGrade.TeacherScore * 5 / 30);
            }

            //寫入Excel相關資料
            WriteToExcel(studentList);

            MessageBox.Show("計算完畢，執行成功!!");
        }

        /// <summary>
        /// 寫入Excel相關資料
        /// </summary>
        /// <param name="studentList"></param>
        private void WriteToExcel(List<StudentScore> studentList)
        {
            IWorkbook workbook;

            using (FileStream fileStream = new FileStream(pathOfExcelOutput, FileMode.Open, FileAccess.Read))
            {
                workbook = new XSSFWorkbook(fileStream);
            }

            var sheet = workbook.GetSheetAt(0);
            var rows = sheet.GetRowEnumerator();

            for (var i = (sheet.FirstRowNum + 3); i <= sheet.LastRowNum; i++)
            {
                var row = sheet.GetRow(i);
                var cell0 = row.GetCell(0);
                var studentNumber = cell0.StringCellValue;
                var studentObject = studentList.Find(x => x.ID == studentNumber);
                if (studentObject != null)
                {
                    var cellScores = new List<int> { 3, 8, 9, 10 };
                    foreach (var cellNum in cellScores)
                    {
                        var cell = row.GetCell(cellNum);
                        if (cell == null)
                        {
                            cell = row.CreateCell(cellNum);
                        }
                        if (cell != null)
                        {
                            switch (cellNum)
                            {
                                //期中成績
                                case 3:
                                    cell.SetCellValue(studentObject.MidGrade.Total.ToString());
                                    break;
                                //同儕自評
                                case 8:
                                    cell.SetCellValue(studentObject.MidGrade.PeerScore.ToString());
                                    break;
                                //自我評分
                                case 9:
                                    cell.SetCellValue(studentObject.MidGrade.SelfScore.ToString());
                                    break;
                                //教師評分
                                case 10:
                                    cell.SetCellValue(studentObject.MidGrade.TeacherScore.ToString());
                                    break;
                            }

                            //if (studentObject.MidGrade.Total < 80)
                            //{
                            //    if(studentObject.MidGrade.Total == 0)
                            //    {
                            //    }
                            //    else
                            //    {
                            //    }
                            //}
                        }
                    }
                }
            }

            using (FileStream fs = new FileStream(pathOfExcelOutput, FileMode.Create, FileAccess.Write))
            {
                workbook.Write(fs);
            }
        }

        /// <summary>
        /// 將等第轉換成數字
        /// </summary>
        /// <returns></returns>
        private int ConvertGradeToScore(string grade)
        {
            if (grade.Length < 5)
            {
                switch (grade)
                {
                    case "A+":
                        return 95;
                    case "A":
                        return 87;
                    case "A-":
                        return 82;
                    case "B+":
                        return 78;
                    case "B":
                        return 75;
                    case "B-":
                        return 70;
                    default:
                        return 87;
                }
            }
            else
            {
                switch (grade)
                {
                    case "A+ (90~100)":
                        return 95;
                    case "A (85~89)":
                        return 87;
                    case "A- (80~84)":
                        return 82;
                    case "B+ (77~79)":
                        return 78;
                    case "B (73~76)":
                        return 75;
                    case "B- (70~72)":
                        return 70;
                    default:
                        return 87;
                }
            }
        }

        /// <summary>
        /// 取得成績、組別及組員資料
        /// </summary>
        /// <param name="groupInformation"></param>
        /// <param name="jointSheetData"></param>
        /// <returns></returns>
        private List<GroupWithGrade> GetGroupWithGradeInformation(List<GroupInformation> groupInformation, PeerJointSheet jointSheetData)
        {
            var groupInfoWithGrade = JsonConvert.DeserializeObject<List<GroupWithGrade>>(JsonConvert.SerializeObject(groupInformation));

            foreach (var group in groupInfoWithGrade)
            {
                //共評表一裡是否有該組別
                Type type = typeof(PeerJointSheetOne);
                var hasThisGroup = type.GetProperties().Any(x => x.Name == group.GroupName);
                if (hasThisGroup)
                {
                    group.Records = new Dictionary<string, string>();
                    foreach (var record in jointSheetData.SheetOne)
                    {
                        var propertyInfo = type.GetProperty(group.GroupName);
                        var score = propertyInfo.GetValue(record).ToString();

                        //教師評分
                        if (record.StudentID.IndexOf("S") > 0)
                        {
                            group.TeachScore = score;
                        }
                        else
                        {
                            if (score != null && score.Length > 0)
                            {
                                group.Records.Add(record.StudentName, score);
                            }
                            //else
                            //{
                            //}
                        }
                    }

                    //if (group.TeachScore == null)
                    //{
                    //}
                }

                //共評表二裡是否有該組別
                type = typeof(PeerJointSheetTwo);
                hasThisGroup = type.GetProperties().Any(x => x.Name == group.GroupName);
                if (hasThisGroup)
                {
                    group.Records = new Dictionary<string, string>();
                    foreach (var record in jointSheetData.SheetTwo)
                    {
                        var propertyInfo = type.GetProperty(group.GroupName);
                        var score = propertyInfo.GetValue(record).ToString();

                        //教師評分
                        if (record.StudentID.IndexOf("S") > 0)
                        {
                            group.TeachScore = score;
                        }
                        else
                        {
                            if (score != null && score.Length > 0)
                            {
                                group.Records.Add(record.StudentName, score);
                            }
                            //else
                            //{
                            //}
                        }
                    }

                    //if (group.TeachScore == null)
                    //{
                    //}
                }
            }

            return groupInfoWithGrade;
        }

        /// <summary>
        /// 取得組別資料
        /// </summary>
        /// <returns></returns>
        private List<GroupInformation> GetGroupInformation()
        {
            var groupInfoList = new List<GroupInformation>();

            using (var reader = new StreamReader(pathOfGroupInformation))
            using (var csv = new CsvReader(reader, CultureInfo.InvariantCulture))
            {
                groupInfoList = csv.GetRecords<GroupInformation>().ToList();
            }

            return groupInfoList;
        }

        /// <summary>
        /// 依最終選課表取得學生清單
        /// </summary>
        /// <returns></returns>
        private List<StudentScore> GetStudentList()
        {
            var studentScoreList = new List<StudentScore>();

            using (var reader = new StreamReader(pathOfFinalStudentList))
            using (var csv = new CsvReader(reader, CultureInfo.InvariantCulture))
            {
                var records = csv.GetRecords<Student>();

                foreach (var record in records)
                {
                    var studentScore = new StudentScore
                    {
                        Name = record.Name,
                        ID = record.ID,
                    };
                    studentScoreList.Add(studentScore);
                }
            }

            return studentScoreList;
        }

        /// <summary>
        /// 取得共評表相關資料
        /// </summary>
        /// <returns></returns>
        private PeerJointSheet CalculateJointSheets()
        {
            var peerJointSheet = new PeerJointSheet()
            {
                SheetOne = new List<PeerJointSheetOne>(),
                SheetTwo = new List<PeerJointSheetTwo>(),
            };

            //讀取共評表一
            using (var reader = new StreamReader(pathOfJointSheetOne))
            using (var csv = new CsvReader(reader, CultureInfo.InvariantCulture))
            {
                var records = csv.GetRecords<PeerJointSheetOne>();
                peerJointSheet.SheetOne = records.ToList();
            }

            //讀取共評表二
            using (var reader = new StreamReader(pathOfJointSheetTwo))
            using (var csv = new CsvReader(reader, CultureInfo.InvariantCulture))
            {
                var records = csv.GetRecords<PeerJointSheetTwo>();
                peerJointSheet.SheetTwo = records.ToList();
            }

            return peerJointSheet;
        }

        /// <summary>
        /// 檢查共評表一是否存在
        /// </summary>
        private void DoesSheet1Exist()
        {
            string sheet1Path = sheet1TextBox.Text;
            DoesFileExistAndShowMessage(sheet1Path, "無法取得共評表一");

        }

        /// <summary>
        /// 檢查共評表二是否存在
        /// </summary>
        private void DoesSheet2Exist()
        {
            var sheet2Path = sheet2TextBox.Text;
            DoesFileExistAndShowMessage(sheet2Path, "無法取得共評表二");
        }

        /// <summary>
        /// 檢查組員資訊表是否存在
        /// </summary>
        private void DoesGroupInfoExist()
        {
            var groupInfoPath = groupInfoTextBox.Text;
            DoesFileExistAndShowMessage(groupInfoPath, "無法取得組員資訊");
        }

        /// <summary>
        /// 檢查輸出表是否存在
        /// </summary>
        private void DoesOutputSheetExist()
        {
            var outputSheetPath = outputSheetTextBox.Text;
            DoesFileExistAndShowMessage(outputSheetPath, "無法取得輸出資料表");
        }

        /// <summary>
        /// 顯示輸出訊息
        /// </summary>
        /// <param name="path"></param>
        /// <param name="message"></param>
        private void DoesFileExistAndShowMessage(string path, string message)
        {
            if (!File.Exists(path))
            {
                MessageBox.Show(message);
            }
        }
    }
}

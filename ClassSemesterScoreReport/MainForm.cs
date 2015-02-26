using Aspose.Cells;
using FISCA.Data;
using FISCA.Presentation.Controls;
using K12.Data;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace ClassSemesterScoreReport
{
    public partial class MainForm : BaseForm
    {
        int _schoolYear, _semester;

        List<string> _classIDs;

        BackgroundWorker _BW;

        public MainForm()
        {
            InitializeComponent();

            _classIDs = K12.Presentation.NLDPanels.Class.SelectedSource;

            _BW = new BackgroundWorker();
            _BW.DoWork += new DoWorkEventHandler(BW_DoWork);
            _BW.RunWorkerCompleted += new RunWorkerCompletedEventHandler(BW_Completed);

            int schoolYear, semester;

            _schoolYear = int.TryParse(K12.Data.School.DefaultSchoolYear, out schoolYear) ? schoolYear : 0;
            _semester = int.TryParse(K12.Data.School.DefaultSemester, out semester) ? semester : 0;

            for (int i = -2; i <= 2; i++)
                cboSchoolYear.Items.Add(_schoolYear + i);

            cboSemester.Items.Add(1);
            cboSemester.Items.Add(2);

            cboSchoolYear.Text = _schoolYear + "";
            cboSemester.Text = _semester + "";
        }

        private void BW_Completed(object sender, RunWorkerCompletedEventArgs e)
        {
            SetFormEnable(true);

            Workbook wb = e.Result as Workbook;

            SaveFileDialog save = new SaveFileDialog();
            save.Title = "另存新檔";
            save.FileName = _schoolYear + "." + _semester + "班級學期成績單.xls";
            save.Filter = "Excel檔案 (*.xls)|*.xls|所有檔案 (*.*)|*.*";
            if (save.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                try
                {
                    wb.Save(save.FileName, SaveFormat.Excel97To2003);
                    System.Diagnostics.Process.Start(save.FileName);
                }
                catch
                {
                    MessageBox.Show("檔案儲存失敗");
                }
            }
        }

        private void BW_DoWork(object sender, DoWorkEventArgs e)
        {
            //取得資料
            Dictionary<string, ClassData> Data = GetData();

            //開始列印
            Workbook wb = new Workbook(new MemoryStream(Properties.Resources.Template));

            #region Style

            //設定Style樣板：四邊框線 水平垂直字中 自動換行
            Style s = wb.Styles[wb.Styles.Add()];
            s.Borders[BorderType.TopBorder].LineStyle = CellBorderType.Thin;
            s.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Thin;
            s.Borders[BorderType.LeftBorder].LineStyle = CellBorderType.Thin;
            s.Borders[BorderType.RightBorder].LineStyle = CellBorderType.Thin;
            s.HorizontalAlignment = TextAlignmentType.Center;
            s.VerticalAlignment = TextAlignmentType.Center;
            s.IsTextWrapped = true;

            //設定Style2樣板：三邊細線 底線粗線 水平垂直字中 自動換行
            Style s2 = wb.Styles[wb.Styles.Add()];
            s2.Borders[BorderType.TopBorder].LineStyle = CellBorderType.Thin;
            s2.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Thick;
            s2.Borders[BorderType.LeftBorder].LineStyle = CellBorderType.Thin;
            s2.Borders[BorderType.RightBorder].LineStyle = CellBorderType.Thin;
            s2.HorizontalAlignment = TextAlignmentType.Center;
            s2.VerticalAlignment = TextAlignmentType.Center;
            s2.IsTextWrapped = true;

            //設定Style3樣板：三邊細線 右線粗線 水平字左 垂直字中
            Style s3 = wb.Styles[wb.Styles.Add()];
            //s3.Borders[BorderType.TopBorder].LineStyle = CellBorderType.Thin;
            s3.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Thin;
            s3.Borders[BorderType.LeftBorder].LineStyle = CellBorderType.Thin;
            s3.Borders[BorderType.RightBorder].LineStyle = CellBorderType.Thick;
            s3.HorizontalAlignment = TextAlignmentType.Left;
            s3.VerticalAlignment = TextAlignmentType.Center;

            //設定Style4樣板：兩邊細線 右邊底線粗線 水平字左 垂直字中
            Style s4 = wb.Styles[wb.Styles.Add()];
            s4.Borders[BorderType.TopBorder].LineStyle = CellBorderType.Thin;
            s4.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Thick;
            s4.Borders[BorderType.LeftBorder].LineStyle = CellBorderType.Thin;
            s4.Borders[BorderType.RightBorder].LineStyle = CellBorderType.Thick;
            s4.HorizontalAlignment = TextAlignmentType.Left;
            s4.VerticalAlignment = TextAlignmentType.Center;

            //設定Style5樣板：兩邊細線 右邊底線粗線 水平字中 垂直字中
            Style s5 = wb.Styles[wb.Styles.Add()];
            s5.Borders[BorderType.TopBorder].LineStyle = CellBorderType.Thin;
            s5.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Thick;
            s5.Borders[BorderType.LeftBorder].LineStyle = CellBorderType.Thin;
            s5.Borders[BorderType.RightBorder].LineStyle = CellBorderType.Thick;
            s5.HorizontalAlignment = TextAlignmentType.Center;
            s5.VerticalAlignment = TextAlignmentType.Center;

            //設定Style6樣板：水平字右 垂直字中
            Style s6 = wb.Styles[wb.Styles.Add()];
            s6.HorizontalAlignment = TextAlignmentType.Right;
            s6.VerticalAlignment = TextAlignmentType.Center;

            #endregion

            //每一班
            foreach (string id in Data.Keys)
            {
                int newSheet = wb.Worksheets.AddCopy(0);
                wb.Worksheets[newSheet].Name = Data[id].Class.Name;

                wb.Worksheets[newSheet].Cells[0, 0].PutValue(string.Format("雙語部  {0} ~ {1}年度第{2}學期  學期成績",_schoolYear + 1911,_schoolYear + 1912,_semester));
                wb.Worksheets[newSheet].Cells[1, 0].PutValue(string.Format("Class: {0}", Data[id].Class.Name));
                wb.Worksheets[newSheet].Cells[2, 0].PutValue("SeatNo");
                wb.Worksheets[newSheet].Cells[2, 1].PutValue("Name");

                //該班的選修最大值
                int maxElective = 0;
                //該班的Domain聯集
                List<string> domainList = new List<string>();
                foreach (StudentObj stu in Data[id].Students.Values)
                {
                    int electiveCount = 0;
                    foreach (string subj in stu.SemesterScoreRecord.Subjects.Keys)
                    {
                        string domain = stu.SemesterScoreRecord.Subjects[subj].Domain;

                        if (domain == "Elective")
                        {
                            electiveCount++;
                            continue;
                        }

                        if (!domainList.Contains(domain))
                            domainList.Add(domain);
                    }

                    //記住最多的選修數量
                    if (electiveCount > maxElective)
                        maxElective = electiveCount;
                }

                //科目名稱排序
                domainList.Sort();

                //加上選修的header
                for (int i = 1; i <= maxElective; i++)
                    domainList.Add("Elective " + i);

                //成績排名
                Data[id].Rank();
                
                //列印columns header
                int colIndex = 2;
                Dictionary<string, int> columnsMapping = new Dictionary<string, int>();
                foreach (string domain in domainList)
                {
                    //列印科目標題
                    wb.Worksheets[newSheet].Cells[2, colIndex].PutValue(domain);

                    //記憶科目索引
                    columnsMapping.Add(domain, colIndex);

                    colIndex++;
                }
                
                //其他column header index
                columnsMapping.Add("Avg", colIndex);
                columnsMapping.Add("AvgGPA", colIndex + 1);
                columnsMapping.Add("Rank", colIndex + 2);
                columnsMapping.Add("Level", colIndex + 3);

                //列印其他標題
                wb.Worksheets[newSheet].Cells[2, columnsMapping["Avg"]].PutValue("Avg");
                wb.Worksheets[newSheet].Cells[2, columnsMapping["AvgGPA"]].PutValue("AvgGPA");
                wb.Worksheets[newSheet].Cells[2, columnsMapping["Rank"]].PutValue("Rank");
                wb.Worksheets[newSheet].Cells[2, columnsMapping["Level"]].PutValue("Level");

                //合併儲存格：First Row合併 ；Second Row Column 前三後三合併
                wb.Worksheets[newSheet].Cells.Merge(0, 0, 1, columnsMapping["Level"] + 1);
                wb.Worksheets[newSheet].Cells.Merge(1, 0, 1, 2);
                wb.Worksheets[newSheet].Cells.Merge(1, 2, 1, columnsMapping["Level"] - 1);

                wb.Worksheets[newSheet].Cells[1, 2].PutValue(string.Format("列印日期: {0}", SelectTime()));
                
                //列印該班學生資料
                int indexRow = 3;
                foreach (StudentObj stu in Data[id].Students.Values)
                {
                    wb.Worksheets[newSheet].Cells[indexRow, 0].PutValue(stu.StudentRecord.SeatNo);
                    wb.Worksheets[newSheet].Cells[indexRow, 1].PutValue(stu.StudentRecord.Name + " " + stu.StudentRecord.EnglishName);

                    wb.Worksheets[newSheet].Cells[indexRow, columnsMapping["Avg"]].PutValue(stu.SemesterScoreRecord.AvgScore);
                    wb.Worksheets[newSheet].Cells[indexRow, columnsMapping["AvgGPA"]].PutValue(stu.SemesterScoreRecord.AvgGPA);
                    wb.Worksheets[newSheet].Cells[indexRow, columnsMapping["Rank"]].PutValue(stu.Rank);

                    //選修課從1開始
                    int electiveCount = 1;
                    foreach (SubjectScore ss in stu.SemesterScoreRecord.Subjects.Values)
                    {
                        string domain = ss.Domain;

                        //若是選修就加上序號
                        if (domain == "Elective")
                        {
                            domain = domain + " " + electiveCount;
                            electiveCount++;
                        }
                            
                        int columnIndex = columnsMapping[domain];
                        wb.Worksheets[newSheet].Cells[indexRow, columnIndex].PutValue(ss.Score);

                        //只列印Chinese的Level
                        if (ss.Domain.ToLower() == "chinese" && !string.IsNullOrWhiteSpace(ss.Level + ""))
                            wb.Worksheets[newSheet].Cells[indexRow, columnsMapping["Level"]].PutValue("Level " + ss.Level);
                    }

                    indexRow++;
                }

                #region 表格style設定
                wb.Worksheets[newSheet].Cells[1, 2].SetStyle(s6);

                Range all = wb.Worksheets[newSheet].Cells.CreateRange(2, 0, indexRow - 2, columnsMapping["Level"] + 1);
                all.SetStyle(s);

                //每5格劃分隔線
                for (int i = 2; i < indexRow; i += 5)
                {
                    Range target = wb.Worksheets[newSheet].Cells.CreateRange(i, 0, 1, columnsMapping["Level"] + 1);
                    target.SetStyle(s2);
                }

                //姓名欄位畫線
                for (int i = 2; i < indexRow; i ++)
                {
                    if (i == 2)
                        wb.Worksheets[newSheet].Cells[i, 1].SetStyle(s5);
                    else if(i % 5 == 2)
                        wb.Worksheets[newSheet].Cells[i, 1].SetStyle(s4);
                    else
                        wb.Worksheets[newSheet].Cells[i, 1].SetStyle(s3);
                }

                #endregion

            }

            wb.Worksheets.RemoveAt(0);

            e.Result = wb;
        }

        /// <summary>
        /// 取得資料
        /// </summary>
        /// <returns></returns>
        public Dictionary<string, ClassData> GetData()
        {
            Dictionary<string, ClassData> data = new Dictionary<string, ClassData>();

            //班級資料
            foreach (ClassRecord cr in K12.Data.Class.SelectByIDs(_classIDs))
                data.Add(cr.ID, new ClassData(cr));

            //學生資料
            List<string> studentIDs = new List<string>();
            Dictionary<string, string> studentToClass = new Dictionary<string, string>();

            foreach (StudentRecord sr in K12.Data.Student.SelectByClassIDs(_classIDs))
            {
                //只列印一般及延修生
                if (sr.Status != StudentRecord.StudentStatus.一般 && sr.Status != StudentRecord.StudentStatus.延修)
                    continue;

                data[sr.RefClassID].Students.Add(sr.ID, new StudentObj(sr));

                studentIDs.Add(sr.ID);

                studentToClass.Add(sr.ID, sr.RefClassID);
            }

            //學期成績資料
            if (studentIDs.Count > 0)
            {
                List<SemesterScoreRecord> ssrs = K12.Data.SemesterScore.SelectBySchoolYearAndSemester(studentIDs, _schoolYear, _semester);
                foreach (SemesterScoreRecord ssr in ssrs)
                {
                    string classID = studentToClass[ssr.RefStudentID];

                    data[classID].Students[ssr.RefStudentID].SemesterScoreRecord = ssr;
                }
            }

            return data;
        }

        private string SelectTime() //取得Server的時間
        {
            QueryHelper Sql = new QueryHelper();
            DataTable dtable = Sql.Select("select now()"); //取得時間
            DateTime dt = DateTime.Now;
            DateTime.TryParse("" + dtable.Rows[0][0], out dt); //Parse資料
            string ComputerSendTime = dt.ToString("yyyy/MM/dd"); //最後時間

            return ComputerSendTime;
        }

        private void btnOK_Click(object sender, EventArgs e)
        {
            if (!int.TryParse(cboSchoolYear.Text, out _schoolYear))
            {
                MessageBox.Show("學年度必須是整數數字");
                return;
            }


            if (!int.TryParse(cboSemester.Text, out _semester))
            {
                MessageBox.Show("學期必須是整數數字");
                return;
            }

            if (_BW.IsBusy)
                MessageBox.Show("系統忙碌中,請稍後再試...");
            else
            {
                SetFormEnable(false);
                _BW.RunWorkerAsync();
            }
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void SetFormEnable(bool b)
        {
            cboSchoolYear.Enabled = b;
            cboSemester.Enabled = b;
            btnOK.Enabled = b;
        }

        public class ClassData
        {
            public ClassRecord Class;
            public Dictionary<string, StudentObj> Students;

            public ClassData(ClassRecord cr)
            {
                Class = cr;
                Students = new Dictionary<string, StudentObj>();
            }

            /// <summary>
            /// 對該班的學生作排名
            /// </summary>
            public void Rank()
            {
                List<StudentObj> rankList = Students.Values.ToList();

                rankList.Sort(delegate(StudentObj x, StudentObj y)
                {
                    decimal xScore = x.SemesterScoreRecord.AvgScore.HasValue ? x.SemesterScoreRecord.AvgScore.Value : 0;
                    decimal yScore = y.SemesterScoreRecord.AvgScore.HasValue ? y.SemesterScoreRecord.AvgScore.Value : 0;
                    decimal xGPA = x.SemesterScoreRecord.AvgGPA.HasValue ? x.SemesterScoreRecord.AvgGPA.Value : 0;
                    decimal yGPA = y.SemesterScoreRecord.AvgGPA.HasValue ? y.SemesterScoreRecord.AvgGPA.Value : 0;

                    //先比GPA,同分就比平均
                    if (xGPA == yGPA)
                        return xScore.CompareTo(yScore);
                    else
                        return xGPA.CompareTo(yGPA);
                });

                rankList.Reverse();

                int rank = 0;
                int count = 0;
                decimal currentScore = decimal.MinValue;
                decimal currentGPA = decimal.MinValue;
                foreach (StudentObj stu in rankList)
                {
                    count++;

                    decimal score = stu.SemesterScoreRecord.AvgScore.HasValue ? stu.SemesterScoreRecord.AvgScore.Value : 0;
                    decimal gpa = stu.SemesterScoreRecord.AvgGPA.HasValue ? stu.SemesterScoreRecord.AvgGPA.Value : 0;

                    if (currentGPA != gpa)
                        rank = count;
                    else if (currentScore != score)
                        rank = count;

                    stu.Rank = rank;
                    currentGPA = gpa;
                    currentScore = score;
                }
            }
        }

        public class StudentObj
        {
            public StudentRecord StudentRecord;
            public SemesterScoreRecord SemesterScoreRecord;
            public int Rank;

            public StudentObj(StudentRecord sr)
            {
                StudentRecord = sr;
                SemesterScoreRecord = new SemesterScoreRecord();
            }
        }
    }
}

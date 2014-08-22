using Aspose.Words;
using Aspose.Words.Tables;
using CourseGradeB;
using CourseGradeB.EduAdminExtendControls;
using CourseGradeB.StuAdminExtendControls;
using FISCA.Data;
using FISCA.Presentation.Controls;
using FISCA.UDT;
using K12.Data;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;

namespace ConductReportForGrade3to6
{
    public partial class Reporter : BaseForm
    {
        private int _schoolYear, _semester;
        AccessHelper _A;
        QueryHelper _Q;
        List<string> _ids;
        Dictionary<string, List<string>> _hrt_template;
        Dictionary<string, List<string>> _common_template;
        static Dictionary<string, string> _SubjToDomain;
        List<Tool.Domain> _domains;
        BackgroundWorker _BW;
        string _校長, _主任;

        public Reporter(List<string> ids)
        {
            InitializeComponent();
            _A = new AccessHelper();
            _Q = new QueryHelper();
            _ids = ids;
            _hrt_template = new Dictionary<string, List<string>>();
            _common_template = new Dictionary<string, List<string>>();
            _SubjToDomain = new Dictionary<string, string>();
            _校長 = K12.Data.School.Configuration["學校資訊"].PreviousData.SelectSingleNode("ChancellorChineseName").InnerText;
            _主任 = K12.Data.School.Configuration["學校資訊"].PreviousData.SelectSingleNode("EduDirectorName").InnerText;

            _BW = new BackgroundWorker();
            _BW.DoWork += new DoWorkEventHandler(BW_DoWork);
            _BW.RunWorkerCompleted += new RunWorkerCompletedEventHandler(BW_Completed);

            _schoolYear = int.Parse(K12.Data.School.DefaultSchoolYear);
            _semester = int.Parse(K12.Data.School.DefaultSemester);

            for (int i = -2; i <= 2; i++)
                cboSchoolYear.Items.Add(_schoolYear + i);

            cboSemester.Items.Add(1);
            cboSemester.Items.Add(2);

            cboSchoolYear.Text = _schoolYear + "";
            cboSemester.Text = _semester + "";

            //讀取上次主任姓名
            //txtDean.Text = Properties.Settings.Default.Dean;

            LoadTemplate();
            SetSubjectMapping();
        }

        public void SetSubjectMapping()
        {
            //取得Grade3-6的Domains並排序
            _domains = Tool.DomainDic[6];

            Tool.Domain hrt = new Tool.Domain();
            hrt.ShortName = "H.R.T";
            hrt.DisplayOrder = -1;
            hrt.Name = "Homeroom";
            _domains.Add(hrt);

            _domains.Sort(delegate(Tool.Domain x, Tool.Domain y)
            {
                int xx = x.DisplayOrder;
                int yy = y.DisplayOrder;
                return xx.CompareTo(yy);
            });

            //建立subject to domain對照(subj.Name不應該重複)
            _SubjToDomain.Add("Homeroom", "Homeroom");
            foreach (SubjectRecord subj in _A.Select<SubjectRecord>())
                _SubjToDomain.Add(subj.Name,  subj.Group);
                
        }

        private void BW_DoWork(object sender, DoWorkEventArgs e)
        {
            string id = string.Join(",", _ids);

            //取得指定學生的班級導師
            Dictionary<string, string> student_class_teacher = new Dictionary<string, string>();
            foreach (SemesterHistoryRecord r in K12.Data.SemesterHistory.SelectByStudentIDs(_ids))
            {
                foreach (SemesterHistoryItem item in r.SemesterHistoryItems)
                {
                    if (item.SchoolYear == _schoolYear && item.Semester == _semester)
                    {
                        if (!student_class_teacher.ContainsKey(item.RefStudentID))
                            student_class_teacher.Add(item.RefStudentID, item.Teacher);
                    }
                }
            }

            //取得指定學生conduct record
            List<ConductRecord> records = _A.Select<ConductRecord>("ref_student_id in (" + id + ") and school_year=" + _schoolYear + " and semester=" + _semester + " and term is null");

            Dictionary<string, ConductObj> student_conduct = new Dictionary<string, ConductObj>();
            foreach (ConductRecord record in records)
            {
                string student_id = record.RefStudentId + "";
                if (!student_conduct.ContainsKey(student_id))
                    student_conduct.Add(student_id, new ConductObj(record));

                student_conduct[student_id].LoadRecord(record);
            }

            //排序tempalte的group及item
            foreach (string group in _hrt_template.Keys)
                _hrt_template[group].Sort(Sorting);

            foreach (string group in _common_template.Keys)
                _common_template[group].Sort(Sorting);


            //開始列印
            Document doc = new Document();

            foreach (ConductObj obj in student_conduct.Values)
            {
                Dictionary<string, string> mergeDic = new Dictionary<string, string>();
                mergeDic.Add("姓名", obj.Student.Name);
                mergeDic.Add("班級", obj.Student.SeatNo + " Gr. " + obj.Class.Name);
                mergeDic.Add("學年度", (_schoolYear + 1911) + "-" + (_schoolYear + 1912));
                mergeDic.Add("學期", _semester == 1 ? _semester + "st" : _semester+"nd");
                mergeDic.Add("班導師", student_class_teacher.ContainsKey(obj.StudentID) ? student_class_teacher[obj.StudentID] : "");
                mergeDic.Add("校長", _校長);
                mergeDic.Add("主任", _主任);

                Document temp = new Aspose.Words.Document(new MemoryStream(Properties.Resources.template));
                DocumentBuilder bu = new DocumentBuilder(temp);

                //Table1
                bu.MoveToMergeField("Comment");
                bu.StartTable();
                bu.ParagraphFormat.Alignment = ParagraphAlignment.Left;

                bu.InsertCell();
                bu.CellFormat.VerticalAlignment = CellVerticalAlignment.Center;
                bu.CellFormat.Width = 30;
                bu.Writeln("");
                bu.Writeln("M = Meets expectations");
                bu.Writeln("S = Meets needs with Support");
                bu.Writeln("N = Not yet within expectations");
                bu.Writeln("N/A = Not applicable");
                
                bu.InsertCell();
                bu.CellFormat.Width = 120;
                bu.CellFormat.VerticalAlignment = CellVerticalAlignment.Top;
                bu.Font.Bold = true;
                bu.Writeln("HOMEROOM TEACHER'S COMMENT");
                bu.Font.Bold = false;
                bu.Writeln(obj.Comment + "");
                bu.EndRow();
                bu.EndTable();

                //Table2
                bu.MoveToMergeField("Conduct");
                bu.StartTable();

                //列印領域名稱
                bu.InsertCell(); //空白欄
                bu.CellFormat.VerticalAlignment = CellVerticalAlignment.Center;
                bu.CellFormat.Width = 50;
                foreach (Tool.Domain domain in _domains)
                {
                    if (domain.ShortName == "Humanities" || domain.ShortName == "Humanity")
                        continue;

                    bu.InsertCell();
                    bu.CellFormat.Width = 10;
                    bu.Write(domain.ShortName);
                    bu.ParagraphFormat.Alignment = ParagraphAlignment.Center;
                }

                bu.EndRow();
                
                //列印group
                foreach (string group in _hrt_template.Keys)
                {
                    bu.InsertCell();
                    bu.CellFormat.Width = 150;
                    bu.Font.Bold = true;
                    bu.Font.Size = 12;
                    bu.Write(group);
                    bu.ParagraphFormat.Alignment = ParagraphAlignment.Left;
                    bu.EndRow();

                    int index = 1;
                    foreach(string item in _hrt_template[group])
                    {
                        bu.Font.Bold = false;
                        bu.Font.Size = 10;
                        bu.InsertCell();
                        bu.CellFormat.Width = 50;
                        bu.Write(index + "." + item);
                        index++;
                        bu.ParagraphFormat.Alignment = ParagraphAlignment.Left;

                        foreach (Tool.Domain domain in _domains)
                        {
                            if (domain.ShortName == "Humanities" || domain.ShortName == "Humanity")
                                continue;

                            string key = domain.Name + "_" + group + "_" + item;

                            bu.InsertCell();
                            bu.CellFormat.Width = 10;
                            string grade = obj.ConductGrade.ContainsKey(key) ? obj.ConductGrade[key] : "";
                            bu.Write(grade);
                            bu.ParagraphFormat.Alignment = ParagraphAlignment.Center;
                        }

                        bu.EndRow();
                    }
                }

                foreach(string group in _common_template.Keys)
                {
                    bu.InsertCell();
                    bu.CellFormat.Width = 150;
                    bu.Font.Bold = true;
                    bu.Font.Size = 12;
                    bu.Write(group);
                    bu.ParagraphFormat.Alignment = ParagraphAlignment.Left;
                    bu.EndRow();

                    int index = 1;
                    foreach(string item in _common_template[group])
                    {
                        bu.Font.Bold = false;
                        bu.Font.Size = 10;
                        bu.InsertCell();
                        bu.CellFormat.Width = 50;
                        bu.Write(index + "." + item);
                        index++;
                        bu.ParagraphFormat.Alignment = ParagraphAlignment.Left;

                        foreach (Tool.Domain domain in _domains)
                        {
                            if (domain.ShortName == "Humanities" || domain.ShortName == "Humanity")
                                continue;

                            string key = domain.Name + "_" + group + "_" + item;

                            bu.InsertCell();
                            bu.CellFormat.Width = 10;
                            string grade = obj.ConductGrade.ContainsKey(key) ? obj.ConductGrade[key] : "";
                            bu.Write(grade);
                            bu.ParagraphFormat.Alignment = ParagraphAlignment.Center;
                        }

                        bu.EndRow();
                    }
                }
                //bu.EndTable();

                //Table3
                //bu.MoveToMergeField("Attend");
                //bu.StartTable();

                bu.Font.Size = 12;
                bu.Font.Bold = true;
                bu.InsertCell();
                bu.CellFormat.Width = 75;
                bu.Write("TOTAL Days of School:");

                bu.InsertCell();
                bu.CellFormat.Width = 75;
                bu.Writeln("TOTAL Days of Absence:");
                bu.Write("(Recorded by one week before the final exam.)");

                bu.EndRow();
                bu.EndTable();

                temp.MailMerge.Execute(mergeDic.Keys.ToArray(), mergeDic.Values.ToArray());
                doc.Sections.Add(doc.ImportNode(temp.FirstSection, true));
            }

            doc.Sections.RemoveAt(0);

            e.Result = doc;
        }

        private int Sorting(string x, string y)
        {
            return x.Length.CompareTo(y.Length);
        }

        private void BW_Completed(object sender, RunWorkerCompletedEventArgs e)
        {
            Document doc = e.Result as Document;
            SaveFileDialog save = new SaveFileDialog();
            save.Title = "另存新檔";
            save.FileName = "ConductGradeReport(for Grade 3-6).doc";
            save.Filter = "Word檔案 (*.doc)|*.doc|所有檔案 (*.*)|*.*";

            if (save.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                try
                {
                    doc.Save(save.FileName, Aspose.Words.SaveFormat.Doc);
                    System.Diagnostics.Process.Start(save.FileName);
                }
                catch
                {
                    MessageBox.Show("檔案儲存失敗");
                }
            }

            //儲存主任姓名
            //Properties.Settings.Default.Save();
        }

        private void LoadTemplate()
        {
            List<ConductSetting> list = _A.Select<ConductSetting>("grade=6");
            if (list.Count > 0)
            {
                ConductSetting setting = list[0];

                XmlDocument xdoc = new XmlDocument();
                if (!string.IsNullOrWhiteSpace(setting.Conduct))
                    xdoc.LoadXml(setting.Conduct);

                foreach (XmlElement elem in xdoc.SelectNodes("//Conduct[@Common]"))
                {
                    string group = elem.GetAttribute("Group");
                    bool common = elem.GetAttribute("Common") == "True" ? true : false;

                    foreach (XmlElement item in elem.SelectNodes("Item"))
                    {
                        string title = item.GetAttribute("Title");

                        if (common)
                        {
                            if (!_common_template.ContainsKey(group))
                                _common_template.Add(group, new List<string>());

                            _common_template[group].Add(title);
                        }
                        else
                        {
                            if (!_hrt_template.ContainsKey(group))
                                _hrt_template.Add(group, new List<string>());

                            _hrt_template[group].Add(title);
                        }
                    }
                }
            }
        }

        private void btnExit_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btnOk_Click(object sender, EventArgs e)
        {
            //_主任 = txtDean.Text;
            _schoolYear = int.Parse(cboSchoolYear.Text);
            _semester = int.Parse(cboSemester.Text);

            if (_BW.IsBusy)
                MessageBox.Show("系統忙碌中,請稍後再試...");
            else
                _BW.RunWorkerAsync();
        }

        public class ConductObj
        {
            public static XmlDocument _xdoc;
            public Dictionary<string, string> ConductGrade = new Dictionary<string, string>();
            public string Comment;
            public string StudentID;
            public StudentRecord Student;
            public ClassRecord Class;

            public ConductObj(ConductRecord record)
            {
                StudentID = record.RefStudentId + "";

                Student = K12.Data.Student.SelectByID(StudentID);
                Class = Student.Class;

                if (Student == null)
                    Student = new StudentRecord();

                if (Class == null)
                    Class = new ClassRecord();
            }

            public void LoadRecord(ConductRecord record)
            {
                string subj = record.Subject;
                if (string.IsNullOrWhiteSpace(subj))
                    subj = "Homeroom";

                string domain = _SubjToDomain.ContainsKey(subj) ? _SubjToDomain[subj] : "";

                if (subj == "Homeroom")
                    Comment = record.Comment;

                //XML
                if (_xdoc == null)
                    _xdoc = new XmlDocument();

                _xdoc.RemoveAll();
                if (!string.IsNullOrWhiteSpace(record.Conduct))
                    _xdoc.LoadXml(record.Conduct);

                foreach (XmlElement elem in _xdoc.SelectNodes("//Conduct"))
                {
                    string group = elem.GetAttribute("Group");

                    foreach (XmlElement item in elem.SelectNodes("Item"))
                    {
                        string title = item.GetAttribute("Title");
                        string grade = item.GetAttribute("Grade");

                        //if (!ConductGrade.ContainsKey(subj + "_" + group + "_" + title))
                        //    ConductGrade.Add(subj + "_" + group + "_" + title, grade);

                        if (!ConductGrade.ContainsKey(domain + "_" + group + "_" + title))
                            ConductGrade.Add(domain + "_" + group + "_" + title, grade);
                    }
                }
            }
        }
    }
}

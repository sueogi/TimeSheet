using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Drawing;
using System.Data;
using System.Data.SqlClient;
using System.Reflection;
using Timesheet.FrmMainFolder;
using Microsoft.Reporting.WinForms;

namespace Timesheet.FrmQuery
{
    partial class FilterPanel : Panel
    {
        #region 參數
        // 系統參數
        private string OnTime = "10:00";
        private string OffTime = "19:00";
        private string Account;
        private string UserName;
        private int criteriaLimit = 3;

        private bool isAdmin = false;
        private bool isPM = false;
        private Dictionary<string, int> dicRole = new Dictionary<string, int>();
        private DataSet ds = new DataSet();
        private string strChartCmd = string.Empty;
        private string strChartRdlc = string.Empty;

        private int extendHeight;
        private int blockHeight = 26;
        private bool flagExtend = false;

        // 點擊增加條件的次數
        private int addProj = 0;
        private int addStep = 0;
        private int addTask = 0;
        private int addAccn = 0;

        // 選擇的條件
        private List<string> Projs = new List<string>();
        private List<string> Steps = new List<string>();
        private List<string> Tasks = new List<string>();
        private List<string> Accns = new List<string>();
        #endregion

        public FilterPanel(Panel middle)
        {
            InitializeComponent();

            // 整個搜尋 Panel 的高度 = (時間條件搜尋 Panel + 其他條件搜尋 Panel)
            // 當點擊展開搜尋條件時記住上次增加條件後的 Panel2 高度
            extendHeight = this.panelMainSearch.Height + this.panelExpandSearch.Height;

            // 取出 FrmMain 拿到的系統參數及使用者帳號            
            Account = (middle as MiddleControlMenuGroup).info.getValue("AccountName");
            string Admin = (middle as MiddleControlMenuGroup).info.getValue("Admin");            

            string strCmd = @"SELECT Name FROM Users WHERE Account = '" + Account + "'";
            UserName = SQLToolsFolder.SQLToolsConnect.selectFunc(strCmd);

            // 比對現在使用者是否為 Admin
            isAdmin = Admin.Equals(Account) ? true : false;

            ExpandClick(null, null);            

            // 填滿預設的搜尋條件資料
            // getComboBoxData 方法裡面會根據使用者是否為 Admin 來決定是否用所有專案來填滿
            ComboBox[] cbs = new ComboBox[] { cbProj0, cbStep0, cbTask0, cbSortBy, cbGroupBy };
            QueryModel.getComboBoxData(cbs, ds, isAdmin, Account);

            // 如果非 Admin 的使用者沒有參與任何專案就 disable 專案階段            
            if (cbs[0].Items.Count == 1)
            {
                panelStep.Enabled = false;
            }

            // 如果使用者不是 Admin，該使用者只能看到自己參與的專案
            // 如果使用者是 PM，該使用者才能查詢所屬專案以下的成員 TimeSheet
            // 所以做出字典檔來得知使用者是否是此專案的PM
            if (!isAdmin)
            {
                List<string> resData = new List<string>();
                SqlCommand sqlCmd = new SqlCommand();
                sqlCmd.CommandText = "SELECT m.RoleID, p.ProjectName FROM Members m join Projects p ON m.ProjectID = p.ProjectID WHERE m.Account = @Account";
                sqlCmd.Parameters.Add("Account", SqlDbType.VarChar, 20).Value = Account;

                if (SQLToolsFolder.SQLToolsConnect.selectFunc(sqlCmd, ref resData) > 0)
                {
                    for (int i = 1; i < resData.Count + 1; i += 2)
                    {
                        dicRole.Add(resData[i], Convert.ToInt32(resData[i - 1]));
                    }
                }
            }
            chartProj.Enabled = (isAdmin && cbProj0.Items.Count > 0 ? true : false);
        }

        private void ExpandClick(object sender, EventArgs e)
        {
            flagExtend = !flagExtend;

            if (flagExtend)
            {
                pbExpand.Image = global::Timesheet.Properties.Resources.ExpandSearch;
                pbExpand.MouseLeave -= pictureBox_MouseLeave;
                pbExpand.MouseEnter -= pictureBox_MouseEnter;

                this.Size = new Size(this.Width, extendHeight);
                this.lineShape1.Visible = false;
                this.lineShape2.Visible = true;

            }
            else
            {
                pbExpand.Image = global::Timesheet.Properties.Resources.ExpandSearchNormal;
                pbExpand.MouseLeave += pictureBox_MouseLeave;
                pbExpand.MouseEnter += pictureBox_MouseEnter;

                this.Size = new Size(this.Width, blockHeight);
                this.lineShape1.Visible = true;
                this.lineShape2.Visible = false;
            }
        }

        private void SearchClick(object sender, EventArgs e)
        {
            // TODO? Cursor
            QueryPanel query = this.Parent as QueryPanel;
            ((query.Parent as MiddleControlMenuGroup).Parent as Form).Cursor = Cursors.WaitCursor;

            switch (cbGroupBy.SelectedItem.ToString())
            {
                case "日期":
                    query.reportViewer1.LocalReport.ReportEmbeddedResource = "Timesheet.ReportGroupByDate.rdlc"; break;
                case "員工":
                    query.reportViewer1.LocalReport.ReportEmbeddedResource = "Timesheet.ReportGroupByAccount.rdlc"; break;
                case "工作類別":
                    query.reportViewer1.LocalReport.ReportEmbeddedResource = "Timesheet.ReportGroupByTaskName.rdlc"; break;
                case "專案":
                    query.reportViewer1.LocalReport.ReportEmbeddedResource = "Timesheet.ReportGroupByProjectName.rdlc";
                    if (chartProj.Checked) query.reportViewer1.LocalReport.SetParameters(new ReportParameter("ChartShow", "False"));
                    else query.reportViewer1.LocalReport.SetParameters(new ReportParameter("ChartShow", "True"));
                    break;
                case "專案階段":
                    query.reportViewer1.LocalReport.ReportEmbeddedResource = "Timesheet.ReportGroupByStepName.rdlc";
                    if (chartStep.Checked) query.reportViewer1.LocalReport.SetParameters(new ReportParameter("ChartShow", "False"));
                    else query.reportViewer1.LocalReport.SetParameters(new ReportParameter("ChartShow", "True"));
                    break;
                default:
                    query.reportViewer1.LocalReport.ReportEmbeddedResource = "Timesheet.ReportGroupByDate.rdlc"; break;
            }

            DataSet report = getReportDataSource();
            if (report.Tables.Count > 0)
            {
                query.SheetViewBindingSource.DataSource = report;
                query.SheetViewBindingSource.DataMember = "Report";
                query.reportViewer1.RefreshReport();
            }
            else
            {
                // 告知使用者錯誤而不是一片空白畫面
                MessageBox.Show("資料庫異常，可能是查詢不到資料或資料庫連線錯誤。");
            }

            ((query.Parent as MiddleControlMenuGroup).Parent as Form).Cursor = Cursors.Default;
        }

        private void PlusClick(object sender, EventArgs e)
        {
            // 沒有參與專案的使用者按了 + 也沒用
            if (cbProj0.Items.Count < 2 || cbAccn0.Items.Count < 2) return;

            ButtonDeluxe btn = sender as ButtonDeluxe;
            int serial = -1;

            switch (btn.TableName)
            {
                case ButtonDeluxe.TableNameEnum.Proj:
                    serial = addProj;
                    // 如果使用者沒有選任何選項就按 + ，第一個 ComboBox 就自動選第二個
                    cbProj0.SelectedIndex = (cbProj0.SelectedIndex == 0 ? 1 : cbProj0.SelectedIndex);
                    break;
                case ButtonDeluxe.TableNameEnum.Step:
                    serial = addStep;
                    cbStep0.SelectedIndex = (cbStep0.SelectedIndex == 0 ? 1 : cbStep0.SelectedIndex);
                    break;
                case ButtonDeluxe.TableNameEnum.Task:
                    serial = addTask;
                    cbTask0.SelectedIndex = (cbTask0.SelectedIndex == 0 ? 1 : cbTask0.SelectedIndex);
                    break;
                case ButtonDeluxe.TableNameEnum.Accn: serial = addAccn; break;
            }
            serial++;

            // 搜尋條件數量限制
            if (serial < criteriaLimit && serial > 0)
            {
                ComboBoxDeluxe cb = getComboBoxItem(btn);
                cb.Size = new System.Drawing.Size(140, 20);
                cb.Name = "cb" + btn.Name.Substring(6, 4) + serial;
                cb.DropDownStyle = ComboBoxStyle.DropDownList;
                cb.TableName = (ComboBoxDeluxe.TableNameEnum)btn.TableName;
                btn.Parent.Controls.Add(cb);

                // TODO? cb.SelectedIndexChanged += .......... 

                ButtonDeluxe btnMinus = new ButtonDeluxe();
                btnMinus.Size = new Size(25, 20);
                btnMinus.Name = "btnMinus" + btn.TableName.ToString() + serial;
                btnMinus.UseVisualStyleBackColor = true;
                btnMinus.TableName = btn.TableName;
                btnMinus.Text = " -";
                btnMinus.Click += MinusClick;
                btn.Parent.Controls.Add(btnMinus);

                // 以下都是 UI 的位移
                this.Height += blockHeight;
                extendHeight = this.Height;
                this.panelExpandSearch.Height += blockHeight;
                this.lineShape2.Y1 += blockHeight;
                this.lineShape2.Y2 = this.lineShape2.Y1;

                if (btn.TableName == ButtonDeluxe.TableNameEnum.Proj)
                {
                    cb.SelectedIndexChanged += cbProj_SelectedIndexChanged;
                    cb.Location = new System.Drawing.Point(cbProj0.Location.X, (52 + (serial * blockHeight)));
                    btnMinus.Location = new Point(226, (52 + (serial * blockHeight)));

                    if (dicRole.Keys.Contains(cb.SelectedItem.ToString()) && dicRole[cb.SelectedItem.ToString()] != 1)
                    {
                        cbAccn0.Items.Clear();
                        panelAccn.Height -= blockHeight * addAccn;
                        panelExpandSearch.Height -= blockHeight * addAccn;
                        lineShape2.Y1 -= blockHeight * addAccn;
                        lineShape2.Y2 = lineShape2.Y1;
                        this.Height -= blockHeight * addAccn;
                        addAccn -= addAccn;

                        foreach (Control item in panelAccn.Controls)
                        {
                            if (!item.Name.Contains("0") && !item.Name.Contains("btnAdd") && !(item is Label))
                            {
                                item.Dispose();
                            }
                            item.Enabled = false;
                        }
                    }
                }
                else
                {
                    cb.Location = new System.Drawing.Point(70, (3 + (serial * blockHeight)));
                    btnMinus.Location = new Point(215, (3 + (serial * blockHeight)));
                    btn.Parent.Height += blockHeight;
                }

                switch (btn.TableName)
                {
                    case ButtonDeluxe.TableNameEnum.Proj:
                        foreach (Control item in btn.Parent.Controls)
                        {
                            if (item is PanelDeluxe)
                            {
                                ((PanelDeluxe)item).Location = new Point(((PanelDeluxe)item).Location.X, ((PanelDeluxe)item).Location.Y + blockHeight);
                            }
                        }
                        addProj = serial;
                        break;
                    case ButtonDeluxe.TableNameEnum.Step:
                        addStep = serial;

                        foreach (Control item in btn.Parent.Parent.Controls)
                        {
                            if (item is PanelDeluxe)
                            {
                                if (((PanelDeluxe)item).TableName != PanelDeluxe.TableNameEnum.Step)
                                {
                                    ((PanelDeluxe)item).Location = new Point(((PanelDeluxe)item).Location.X, ((PanelDeluxe)item).Location.Y + blockHeight);
                                }
                            }
                        }
                        break;
                    case ButtonDeluxe.TableNameEnum.Task:
                        addTask = serial;
                        panelAccn.Location = new Point(panelAccn.Location.X, panelAccn.Location.Y + blockHeight);
                        break;
                    case ButtonDeluxe.TableNameEnum.Accn:
                        addAccn = serial;
                        break;
                }
            }
        }

        private ComboBoxDeluxe getComboBoxItem(ButtonDeluxe btn)
        {
            ComboBoxDeluxe cb = new ComboBoxDeluxe();

            List<string> selected = new List<string>();
            int count;
            switch (btn.TableName)
            {
                case ButtonDeluxe.TableNameEnum.Step: count = addStep; break;
                case ButtonDeluxe.TableNameEnum.Task: count = addTask; break;
                case ButtonDeluxe.TableNameEnum.Accn: count = addAccn; break;
                default: count = addProj; break;
            }

            foreach (Control item in btn.Parent.Controls)
            {
                if (item is ComboBoxDeluxe)
                {
                    if (((ComboBoxDeluxe)item).TableName != ComboBoxDeluxe.TableNameEnum.Sort &&
                        ((ComboBoxDeluxe)item).TableName != ComboBoxDeluxe.TableNameEnum.Group)
                    {
                        selected.Add(((ComboBoxDeluxe)item).SelectedItem.ToString());
                    }
                }
            }

            if (btn.TableName != ButtonDeluxe.TableNameEnum.Accn)
            {

                var query = from x in ds.Tables[btn.TableName.ToString()].AsEnumerable()
                            select new
                            {
                                Name = (btn.TableName == ButtonDeluxe.TableNameEnum.Proj ? x.Field<string>("ProjectName") : x.Field<string>("TaskName"))
                            };

                foreach (var row in query)
                {
                    if (!selected.Contains(row.Name))
                    {
                        cb.Items.Add(row.Name);
                    }
                }
            }
            else
            {
                if (addAccn > 0)
                {
                    List<string> selectedCB = new List<string>();
                    foreach (Control x in btn.Parent.Controls)
                    {
                        if (x is ComboBoxDeluxe)
                        {
                            foreach (var item in ((ComboBoxDeluxe)x).Items)
                            {
                                if (!"任何員工".Equals(item.ToString()))
                                {
                                    selectedCB.Add(item.ToString());
                                }
                            }
                        }
                    }
                    // 不同專案可能有相同員工，所以要挑出不重複的員工
                    var query = (from x in selectedCB.AsEnumerable()
                                 select x).Distinct();
                    foreach (var item in query)
                    {
                        if (!selected.Contains(item))
                        {
                            cb.Items.Add(item);
                        }
                    }
                }
                else
                {
                    cb.Items.Clear();
                    foreach (Control item in btn.Parent.Controls)
                    {
                        if (item is ComboBoxDeluxe)
                        {
                            ComboBoxDeluxe selectedCB = item as ComboBoxDeluxe;
                            foreach (var x in selectedCB.Items)
                            {
                                if (!selected.Contains(x))
                                {
                                    if (!"任何員工".Equals(x))
                                    {
                                        cb.Items.Add(x);
                                    }
                                }
                            }
                        }
                    }
                }
            }
            cb.SelectedIndex = 0;
            return cb;
        }

        private void MinusClick(object sender, EventArgs e)
        {
            ButtonDeluxe btn = sender as ButtonDeluxe;
            // 依點選的 Button 名字尾碼找出是要刪除第幾個 ComboBox
            int btnLocation = Convert.ToInt32(btn.Name.Substring(12, 1));
            int serial = -1;

            serial = getButtonSerial(btn.TableName, serial, true);

            // 從點擊的 Button 選出是哪一類的 ComboBox 要被刪除
            switch (btn.TableName)
            {
                case ButtonDeluxe.TableNameEnum.Step:
                    panelTask.Location = new Point(panelTask.Location.X, panelTask.Location.Y - blockHeight);
                    panelAccn.Location = new Point(panelAccn.Location.X, panelAccn.Location.Y - blockHeight);
                    break;
                case ButtonDeluxe.TableNameEnum.Task:
                    panelAccn.Location = new Point(panelAccn.Location.X, panelAccn.Location.Y - blockHeight);
                    break;
                case ButtonDeluxe.TableNameEnum.Accn:
                    break;
                case ButtonDeluxe.TableNameEnum.Proj:
                    foreach (Control item in btn.Parent.Controls)
                    {
                        if (item is Panel)
                        {
                            ((Panel)item).Location = new Point(((Panel)item).Location.X, ((Panel)item).Location.Y - blockHeight);
                        }
                    }
                    break;
            }

            this.Height -= blockHeight;
            extendHeight = this.Height;
            this.lineShape2.Y1 -= blockHeight;
            this.lineShape2.Y2 = this.lineShape2.Y1;
            btn.Parent.Height -= blockHeight;

            // 先將要被刪除的 ComboBox 的 meta 存起來，再讓下一個 ComboBox 繼承這個 meta
            ComboBoxDeluxe cbOri = btn.Parent.Controls.Find("cb" + btn.TableName.ToString() + btnLocation, false)[0] as ComboBoxDeluxe;
            ButtonDeluxe btnMinusOri = btn.Parent.Controls.Find("btnMinus" + btn.TableName.ToString() + btnLocation, false)[0] as ButtonDeluxe;

            if (serial > 1 && btnLocation == 1)
            {
                ComboBoxDeluxe cbReborn = btn.Parent.Controls.Find("cb" + btn.TableName.ToString() + "2", false)[0] as ComboBoxDeluxe;
                ButtonDeluxe btnMinusReborn = btn.Parent.Controls.Find("btnMinus" + btn.TableName.ToString() + "2", false)[0] as ButtonDeluxe;

                cbReborn.Name = cbOri.Name;
                btnMinusReborn.Name = btnMinusOri.Name;


                cbReborn.Location = new Point(cbOri.Location.X, cbOri.Location.Y);
                btnMinusReborn.Location = new Point(btnMinusOri.Location.X, btnMinusOri.Location.Y);
            }

            serial--;
            getButtonSerial(btn.TableName, serial, false);

            // 避免找不到控制項，最後再消滅
            cbOri.Dispose();
            btnMinusOri.Dispose();

            checkPMProj();
        }

        private void checkPMProj()
        {
            foreach (Control item in this.Controls)
            {
                if (item is ComboBoxDeluxe && ((ComboBoxDeluxe)item).TableName == ComboBoxDeluxe.TableNameEnum.Proj)
                {
                    ComboBoxDeluxe cb = item as ComboBoxDeluxe;
                    if (dicRole.Keys.Contains(cb.SelectedItem.ToString()) && dicRole[cb.SelectedItem.ToString()] != 1)
                    {
                        cbAccn0.Items.Clear();
                        panelAccn.Height -= blockHeight * addAccn;
                        panelExpandSearch.Height -= blockHeight * addAccn;
                        lineShape2.Y1 -= blockHeight * addAccn;
                        lineShape2.Y2 = lineShape2.Y1;
                        this.Height -= blockHeight * addAccn;
                        addAccn -= addAccn;

                        foreach (Control acc in panelAccn.Controls)
                        {
                            if (acc.Location.Y > cbAccn0.Location.Y)
                            {
                                acc.Dispose();
                            }
                            acc.Enabled = false;
                        }
                    }
                }
                else
                {
                    foreach (Control acc in panelAccn.Controls)
                    {
                        acc.Enabled = true;
                    }

                    List<string> selected = new List<string>();
                    foreach (Control x in panelExpandSearch.Controls)
                    {
                        if (x is ComboBoxDeluxe && x.Location.X >= cbProj0.Location.X && x.Location.Y >= cbProj0.Location.Y)
                        {
                            selected.Add(((ComboBoxDeluxe)x).SelectedItem.ToString());
                        }
                    }

                    var q = (from a in ds.Tables["Accn"].AsEnumerable()
                             where selected.Contains(a.Field<string>("ProjectName"))
                             select new { Account = a.Field<string>("Account") }).Distinct();
                    cbAccn0.Items.Clear();
                    foreach (var x in q)
                    {
                        cbAccn0.Items.Add(x.Account);
                    }
                    if (cbAccn0.Items.Count > 0) cbAccn0.SelectedIndex = 0;
                }
            }
        }

        private int getButtonSerial(ButtonDeluxe.TableNameEnum tableName, int serial, bool opposite)
        {
            if (opposite)
            {
                switch (tableName)
                {
                    case ButtonDeluxe.TableNameEnum.Proj: serial = addProj; break;
                    case ButtonDeluxe.TableNameEnum.Step: serial = addStep; break;
                    case ButtonDeluxe.TableNameEnum.Task: serial = addTask; break;
                    case ButtonDeluxe.TableNameEnum.Accn: serial = addAccn; break;
                }
            }
            else
            {
                switch (tableName)
                {
                    case ButtonDeluxe.TableNameEnum.Proj: addProj = serial; break;
                    case ButtonDeluxe.TableNameEnum.Step: addStep = serial; break;
                    case ButtonDeluxe.TableNameEnum.Task: addTask = serial; break;
                    case ButtonDeluxe.TableNameEnum.Accn: addAccn = serial; break;
                }
            }
            return serial;
        }

        private void cbProj_SelectedIndexChanged(object sender, EventArgs e)
        {
            ComboBoxDeluxe cb = sender as ComboBoxDeluxe;

            if (cbProj0.SelectedIndex == 0)
            {
                cbAccn0.Items.Clear();
                chartProj.Checked = false;
                chartStep.Checked = false;
                if (addProj > 0)
                {
                    foreach (Control item in panelExpandSearch.Controls)
                    {
                        // 排除 Label 跟增加的 Button
                        if (item.Name.Contains("Proj") && !item.Name.Contains("Add") && !item.Name.Contains("0"))
                        {
                            if (item is ComboBoxDeluxe || item is ButtonDeluxe)
                            {
                                item.Dispose();
                            }
                        }

                        if (item is PanelDeluxe)
                        {
                            ((PanelDeluxe)item).Location = new Point(((PanelDeluxe)item).Location.X, ((PanelDeluxe)item).Location.Y - (addProj * blockHeight));
                        }
                    }
                    this.Height -= blockHeight * addProj;
                    this.lineShape2.Y1 -= blockHeight * addProj;
                    this.lineShape2.Y2 = this.lineShape2.Y1;
                    addProj = 0;
                }
                // 沒有選專案就不能選員工
                foreach (Control item in panelAccn.Controls)
                {
                    item.Enabled = false;
                }

                if (addAccn > 0)
                {
                    panelAccn.Controls.Find("cbAccn1", false)[0].Dispose();
                    panelAccn.Controls.Find("btnMinusAccn1", false)[0].Dispose();

                    if (addAccn > 1)
                    {
                        panelAccn.Controls.Find("cbAccn2", false)[0].Dispose();
                        panelAccn.Controls.Find("btnMinusAccn2", false)[0].Dispose();
                        addAccn--;
                        panelExpandSearch.Height -= blockHeight;
                        panelAccn.Height -= blockHeight;
                        panelExpandSearch.Parent.Height -= blockHeight;
                        extendHeight -= blockHeight;
                        this.lineShape2.Y1 -= blockHeight;
                        this.lineShape2.Y2 = this.lineShape2.Y1;
                    }

                    addAccn--;
                    panelExpandSearch.Height -= blockHeight;
                    panelAccn.Height -= blockHeight;
                    panelExpandSearch.Parent.Height -= blockHeight;
                    extendHeight -= blockHeight;
                    this.lineShape2.Y1 -= blockHeight;
                    this.lineShape2.Y2 = lineShape2.Y1;
                }
            }
            else
            {
                if (isAdmin || dicRole[cb.SelectedItem.ToString()] == 1)
                {
                    foreach (Control item in panelAccn.Controls)
                    {
                        item.Enabled = true;
                    }

                    List<string> selected = new List<string>();

                    foreach (Control item in ((Control)sender).Parent.Controls)
                    {
                        if (item is ComboBoxDeluxe)
                        {
                            if (((ComboBoxDeluxe)item).TableName != ComboBoxDeluxe.TableNameEnum.Sort && ((ComboBoxDeluxe)item).TableName != ComboBoxDeluxe.TableNameEnum.Group)
                            {
                                selected.Add(((ComboBoxDeluxe)item).SelectedItem.ToString());
                            }
                        }
                    }

                    cbAccn0.Items.Clear();

                    if (isAdmin)
                    {
                        if (selected.Contains("任何專案"))
                        {
                            var query = (from x in ds.Tables["Accn"].AsEnumerable()
                                         select new { Account = x.Field<string>("Account") }).Distinct();

                            cbAccn0.Items.Add("任何員工");
                            foreach (var row in query)
                            {
                                cbAccn0.Items.Add(row.Account);
                            }
                        }
                        else
                        {
                            var query = (from x in ds.Tables["Accn"].AsEnumerable()
                                         where selected.Contains(x.Field<string>("ProjectName"))
                                         select new { Account = x.Field<string>("Account") }).Distinct();

                            foreach (var row in query)
                            {
                                cbAccn0.Items.Add(row.Account);
                            }
                        }
                    }
                    else
                    {
                        if (selected.Contains("任何專案"))
                        {
                            var query = (from x in ds.Tables["Accn"].AsEnumerable()
                                         select new { Account = x.Field<string>("Account") }).Distinct();

                            cbAccn0.Items.Add("任何員工");
                            foreach (var row in query)
                            {
                                cbAccn0.Items.Add(row.Account);
                            }
                        }
                        else
                        {
                            var query = (from x in ds.Tables["Accn"].AsEnumerable()
                                         //where x.Field<string>("ProjectName") == cbProj0.SelectedItem.ToString()
                                         where selected.Contains(x.Field<string>("ProjectName"))
                                         select new { Account = x.Field<string>("Account") }).Distinct();

                            foreach (var row in query)
                            {
                                cbAccn0.Items.Add(row.Account);
                            }
                        }
                    }
                    if (cbAccn0.Items.Count != 0)
                    {
                        cbAccn0.SelectedIndex = 0;
                    }
                }
                else
                {
                    cbAccn0.Items.Clear();
                    panelAccn.Height -= blockHeight * addAccn;
                    panelExpandSearch.Height -= blockHeight * addAccn;
                    lineShape2.Y1 -= blockHeight * addAccn;
                    lineShape2.Y2 = lineShape2.Y1;
                    this.Height -= blockHeight * addAccn;
                    addAccn -= addAccn;

                    foreach (Control item in panelAccn.Controls)
                    {
                        if (!item.Name.Contains("0") && !item.Name.Contains("btnAdd") && !(item is Label))
                        {
                            item.Dispose();
                        }
                        item.Enabled = false;
                    }
                }
            }
        }

        private void chartCheckBox_CheckedChanged(object sender, System.EventArgs e)
        {
            CheckBoxDeluxe checkbox = sender as CheckBoxDeluxe;
            switch (checkbox.TableName)
            {
                case CheckBoxDeluxe.TableNameEnum.Proj:
                    strChartCmd = "SELECT * FROM Sheets s JOIN Projects p ON s.ProjectID = p.ProjectID JOIN Users u ON s.Account = u.Account WHERE p.ProjectName = @ProjectName ";
                    strChartRdlc = "Timesheet.Report2.rdlc";
                    if (checkbox.Checked)
                    {
                        chartStep.Enabled = false;
                        cbAccn0.Items.Clear();
                        foreach (Control item in panelAccn.Controls) item.Enabled = false;
                        btnAddProj.Enabled = false;
                        foreach (Control item in panelExpandSearch.Controls)
                        {
                            if (item is Panel)
                            {
                                foreach (Control subItem in ((Panel)item).Controls) subItem.Enabled = false;
                            }
                        }
                        cbProj0.SelectedIndexChanged -= cbProj_SelectedIndexChanged;
                        cbProj0.SelectedIndex = 1;
                        cbGroupBy.SelectedIndex = 3;
                    }
                    else
                    {
                        cbProj0.SelectedIndexChanged += cbProj_SelectedIndexChanged;
                        btnAddProj.Enabled = true;
                        foreach (Control item in panelExpandSearch.Controls)
                        {
                            if (item is Panel)
                            {
                                foreach (Control subItem in ((Panel)item).Controls)
                                {
                                    subItem.Enabled = true;
                                    if (item.Name.Contains("Accn")) subItem.Enabled = false;
                                }
                            }
                        }
                        //foreach (Control item in panelAccn.Controls) item.Enabled = false;

                        cbProj0.SelectedIndex = 0;
                    }
                    break;
                case CheckBoxDeluxe.TableNameEnum.Step:
                    strChartCmd = @"SELECT Name, SpendTime, ProjectName, t.TaskName StepName FROM Sheets s JOIN Projects p ON s.ProjectID = p.ProjectID JOIN Tasks t ON t.TaskID = s.StageID JOIN Users u ON s.Account = u.Account WHERE p.ProjectName = @ProjectName AND s.Account = @Account";
                    strChartRdlc = "Timesheet.Report3.rdlc";
                    cbGroupBy.SelectedIndex = 4;
                    if (checkbox.Checked)
                    {
                        chartProj.Enabled = false;
                        foreach (Control item in panelExpandSearch.Controls) 
                        {
                            if (item is PanelDeluxe)
                            {
                                foreach (var subItem in ((PanelDeluxe)item).Controls)
                                {
                                    if (subItem is Button)
                                    {
                                        ((Button)subItem).Enabled = false;
                                    }
                                }
                            }
                        }
                        cbTask0.Enabled = !cbTask0.Enabled;
                        cbStep0.Enabled = !cbStep0.Enabled;
                        foreach (Control item in panelAccn.Controls) item.Enabled = true;
                        if (cbProj0.Items.Count > 1) cbProj0.SelectedIndex = 1;
                        if (cbAccn0.Items.Count > 1) cbAccn0.SelectedIndex = 1;
                    }
                    else
                    {
                        chartProj.Enabled = true;
                        cbTask0.Enabled = !cbTask0.Enabled;
                        cbStep0.Enabled = !cbStep0.Enabled;
                        foreach (Control item in panelAccn.Controls) item.Enabled = false;
                        if (cbProj0.Items.Count > 0) cbProj0.SelectedIndex = 0;
                    }
                    break;
            }
        }

        #region 取得 ReportViwer 所需的 DataSource
        private DataSet getReportDataSource()
        {
            string error = string.Empty;

            DataSet viewer = new DataSet();
            try
            {
                using (SqlConnection conn = new SqlConnection(global::Timesheet.Properties.Settings.Default.TimeSheetConnectionString))
                {
                    using (SqlCommand sqlCmd = getSqlCommand())
                    {
                        error = sqlCmd.CommandText;

                        sqlCmd.Connection = conn;
                        using (SqlDataAdapter adapter = new SqlDataAdapter(sqlCmd))
                        {
                            adapter.Fill(viewer, "Report");
                        }
                    }
                }
            }
            catch (SqlException ex)
            {
                MessageBox.Show(ex.Message);
            }
            return viewer;
        }

        private SqlCommand getSqlCommand()
        {
            // 收集搜尋條件
            getSearchCriteria();

            StringBuilder sb = new StringBuilder();
            sb.Append("SELECT Account, Name, CONVERT(CHAR(10), BeginTime, 111) BeginDate, BeginTime, EndTime, SpendTime, Summary, ISNULL(ProjectName, '非專案事務') ProjectName, TaskName, ISNULL(StepName, '無') StepName, OverTimeMinute ");
            sb.Append("FROM SheetView(@OnTime, @OffTime) ");
            sb.Append("WHERE BeginTime >= @BeginTime AND BeginTime <= @EndTime ");

            // ==============================================
            // 選擇專案
            if (cbProj0.SelectedIndex != 0)
            {
                sb.Append("AND ProjectName IN (");

                string str = string.Empty;
                for (int i = 0; i < Projs.Count; i++)
                {
                    str += "'" + Projs[i] + "',";

                    if (!isAdmin)
                    {
                        if (dicRole[Projs[i]] == 1)
                        {
                            isPM = true;
                        }
                    }
                }
                sb.Append(str.TrimEnd(',') + ") ");
            }
            // ==============================================
            // 選擇階段
            if (cbStep0.SelectedIndex != 0)
            {
                sb.Append("AND StepName IN (");

                string str = string.Empty;
                for (int i = 0; i < Steps.Count; i++)
                {
                    str += "'" + Steps[i] + "',";
                }
                sb.Append(str.TrimEnd(',') + ") ");
            }
            // ==============================================
            // 選擇類別
            if (cbTask0.SelectedIndex != 0)
            {
                sb.Append("AND TaskName IN (");

                string str = string.Empty;
                for (int i = 0; i < Tasks.Count; i++)
                {
                    str += "'" + Tasks[i] + "',";
                }
                sb.Append(str.TrimEnd(',') + ") ");
            }
            // ==============================================
            // 選擇員工
            if (cbAccn0.Enabled && (isPM || isAdmin))
            {
                sb.Append("AND Account IN (");

                string str = string.Empty;
                for (int i = 0; i < Accns.Count; i++)
                {
                    str += "'" + Accns[i] + "',";
                }
                sb.Append(str.TrimEnd(',') + ") ");
            }
            // ==============================================
            // 如果使用者不是 Admin 或 PM，或者是使用者只要查詢自己，選擇員工的 ComboBox 不會啟用
            if (!cbAccn0.Enabled && !chartProj.Checked && !chartStep.Checked)
            {
                sb.Append("AND Account = @Account ");
            }
            // ==============================================
            // 排序依據
            sb.Append("ORDER BY ");
            switch (cbSortBy.SelectedItem.ToString())
            {
                case "日期":
                    sb.Append("BeginTime");
                    break;
                case "員工":
                    sb.Append("Account");
                    break;
                case "工作類別":
                    sb.Append("TaskName");
                    break;
                case "專案":
                    sb.Append("ProjectName");
                    break;
                case "專案階段":
                    sb.Append("StepName");
                    break;
                default:
                    sb.Append("BeginTime");
                    break;
            }
            // ==============================================                        

            SqlCommand strCmd = new SqlCommand();
            strCmd.CommandText = sb.ToString();
            strCmd.Parameters.Add("OnTime", SqlDbType.Char, 5).Value = OnTime;
            strCmd.Parameters.Add("OffTime", SqlDbType.Char, 5).Value = OffTime;
            strCmd.Parameters.Add("BeginTime", SqlDbType.DateTime).Value = dtpBegin.Value.ToString("yyyy-MM-dd 00:00:00");
            strCmd.Parameters.Add("EndTime", SqlDbType.DateTime).Value = dtpEnd.Value.ToString("yyyy-MM-dd 23:59:59");
            strCmd.Parameters.Add("Account", SqlDbType.VarChar, 20).Value = Account;
            return strCmd;
        }

        private void getSearchCriteria()
        {
            // 清除上次搜尋的條件
            Projs.Clear();
            Steps.Clear();
            Tasks.Clear();
            Accns.Clear();

            foreach (Control item in panelExpandSearch.Controls)
            {
                if (item is ComboBoxDeluxe)
                {
                    ComboBoxDeluxe cb = item as ComboBoxDeluxe;
                    if (cb.TableName == ComboBoxDeluxe.TableNameEnum.Proj)
                    {
                        (getList(cb.TableName)).Add(cb.SelectedItem.ToString());
                    }
                }
                // 專案以外的條件包在個別的 Panel 裡
                else if (item is Panel)
                {
                    foreach (Control subitem in ((Panel)item).Controls)
                    {
                        if (subitem is ComboBoxDeluxe)
                        {
                            ComboBoxDeluxe subComboBox = subitem as ComboBoxDeluxe;
                            if (subComboBox.SelectedIndex > -1)
                            {
                                (getList(subComboBox.TableName)).Add(subComboBox.SelectedItem.ToString());
                            }
                        }
                    }
                }
            }
        }

        private List<string> getList(ComboBoxDeluxe.TableNameEnum cbEnum)
        {
            // 根據傳進來的列舉值回傳對應的 List
            switch (cbEnum)
            {
                case ComboBoxDeluxe.TableNameEnum.Step: return Steps;
                case ComboBoxDeluxe.TableNameEnum.Task: return Tasks;
                case ComboBoxDeluxe.TableNameEnum.Accn: return Accns;
                default: return Projs;
            }
        }
        #endregion

        #region 圖案事件與提示
        private void control_MouseOver(object sender, EventArgs e)
        {
            ToolTip tip = new ToolTip();
            tip.SetToolTip((PictureBox)sender, ((PictureBox)sender).Name);
        }

        private void pictureBox_MouseEnter(object sender, EventArgs e)
        {
            switch (((PictureBox)sender).Name)
            {
                case "Search": ((PictureBox)sender).Image = global::Timesheet.Properties.Resources.SearchEnter; break;
                case "Exit": ((PictureBox)sender).Image = global::Timesheet.Properties.Resources.ExitSearchEnter; break;
                case "Expand": ((PictureBox)sender).Image = global::Timesheet.Properties.Resources.ExpandSearchEnter; break;
            }
        }

        private void pictureBox_MouseLeave(object sender, EventArgs e)
        {
            switch (((PictureBox)sender).Name)
            {
                case "Search": ((PictureBox)sender).Image = global::Timesheet.Properties.Resources.SearchNormal; break;
                case "Exit": ((PictureBox)sender).Image = global::Timesheet.Properties.Resources.ExitSearchNormal; break;
                case "Expand": ((PictureBox)sender).Image = global::Timesheet.Properties.Resources.ExpandSearchNormal; break;
            }
        }
        #endregion

    }
}

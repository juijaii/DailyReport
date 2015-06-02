using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using DevExpress.Skins;
using DevExpress.LookAndFeel;
using DevExpress.UserSkins;
using DevExpress.XtraEditors;
using DevExpress.XtraBars.Helpers;
using Meteorite.Files;
using System.IO;
using Meteorite;
using Meteorite.Application;
using Meteorite.Developer;
using Meteorite.Controls;
using Meteorite.Database;
using DevExpress.XtraGrid;
using DevExpress.XtraGrid.Views.Grid;
using DevExpress.XtraGrid.Views.BandedGrid;
using DevExpress.XtraPrinting;
using DevExpress.XtraEditors.Controls;
using DevExpress.XtraCharts;
using System.Threading;
using DevExpress.XtraTab;



namespace DailyReport
{
    public partial class Form1 : XtraForm
    {
        string DB_SERVER = string.Empty;
        string DB_NAME = string.Empty;
        string DB_USER = string.Empty;
        string DB_PASS = string.Empty;
        public Framework MF { get; private set; }
        public bool sqlOk { get; private set; }
        public int checkPosi { get; set; }

        private IniReader ini { get; set; }
        public DatabaseConfiguration dc;
        public Form1()
        {
            InitializeComponent();
            ini = new IniReader(Application.StartupPath + @"\conf.ini");
            LoadConfig();
            SetDateTimeFormat();
            //  InitSkinGallery();


        }

        private void LoadConfig()
        {
            #region Read Configurations File
            if (!File.Exists(Application.StartupPath + @"\conf.ini"))
                throw new Exception("conf.ini is missing!");

            DB_SERVER = ini.ReadString("DBServer", "IP");
            DB_NAME = ini.ReadString("DBServer", "DBName", String.Empty);
            DB_USER = ini.ReadString("DBServer", "DBUser", String.Empty);
            DB_PASS = ini.ReadString("DBServer", "DBPassword", String.Empty);
            //PRINTER = ini.ReadString("Settings", "Printer", String.Empty);



            if (String.IsNullOrEmpty(DB_SERVER))
                throw new Exception("No IP Address defined in conf.ini");
            if (String.IsNullOrEmpty(DB_NAME))
                throw new Exception("No DBName defined in conf.ini");
            if (String.IsNullOrEmpty(DB_USER))
                throw new Exception("No DBUser defined in conf.ini");
            if (String.IsNullOrEmpty(DB_PASS))
                throw new Exception("No DBPassword defined in conf.ini");

            #endregion
        }


        void InitSkinGallery()
        {
            SkinHelper.InitSkinGallery(rgbiSkins, true);
        }

        void setDateEdit()
        {
            dateEdit1.Text = Getdate();
            dateEdit2.Text = Getdate();
            dateDaily1.Text = Getdate();
            dateDaily2.Text = Getdate();
        }
        private void Form1_Load(object sender, EventArgs e)
        {
            try
            {
                dc = new DatabaseConfiguration()
                {
                    Server = DB_SERVER,
                    Database = DB_NAME,
                    Username = DB_USER,
                    Password = DB_PASS,
                    StoredProcedure = "UTAX_EPZ_CHECK_IN"
                };
                MF = new Framework(defaultLookAndFeel1, this, new Settings("DAILY REPORT", "DAILY REPORT"
                , Development.Stable
                , RunningMode.PublicRelease), dc);
            }

            catch (Exception x)
            {
                MessageBox.Show(x.Message);
            }
            uscMoveGPZ uc;
            XtraTabPage xtrMoveGPZ = xtrTabMoveGPZ;
            if (xtrMoveGPZ != null)
            {
                uc = (uscMoveGPZ)xtrMoveGPZ.Controls[0];
                if (uc != null)
                    uc.MF = MF;
                //uscMoveGPZ ccc = new uscMoveGPZ(MF);
                //xtraTabControl1.TabPages[1].Controls.Clear();
                //xtraTabControl1.TabPages[1].Controls.Add(ccc);
            }

            xtraTabControl1.ShowTabHeader = DevExpress.Utils.DefaultBoolean.False;
            setDateEdit();
            birdLookup();
            DailyReport_ItemClick(null, null);
            groupControl3.Visible = false;
            groupControl4.Visible = false;
            addIncombo(comboBoxEdit1);
            //forReadOnly();
        }
        void forReadOnly()
        {
            barButtonItem1.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
            iOffice.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
            iOp.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
            iNight.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
            iTimeF.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
            addHoliday.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
            ribbonPageGroup4.Visible = false;
            trainBP.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
            trainList.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
            trainUpload.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
            groupControl13.Visible = false;
            simpleButton22.Visible = false;
            gridColumn45.Visible = false;
            gridColumn94.Visible = false;
            groupControl10.Visible = false;
            groupControl19.Visible = false;
            groupControl21.Visible = false;
        }
        public void birdLookup()
        {
            string sql = new SQLCmd(MF);
            sql = sql + string.Format(" @TYPE = 'STA'");
            DataSet ds = searchTimeOut(sql);
            ControlsUtils.BindLookUpEdit(lookListBu, ds.Tables[0], "AB_NAME", "AB_CODE", false);
            ControlsUtils.BindLookUpEdit(leLookType, ds.Tables[0], "AB_NAME", "AB_CODE", false);

            ControlsUtils.BindLookUpEdit(lookDiviOp, ds.Tables[1], "DIVI_NAME", "DIVI_ID", false);
            ControlsUtils.BindLookUpEdit(lookPosiOp, ds.Tables[2], "POSI_NAME", "POSI_ID", false);
            addColumLlReport(bandedGridView3, ds.Tables[0]);
        }
        public bool checkInt(string input)
        {
            try
            {
                Convert.ToInt32(input);
                return true;
            }
            catch
            {
                return false;
            }
        }
        public bool checkDouble(string input)
        {
            try
            {
                Convert.ToDouble(input);
                return true;
            }
            catch
            {
                return false;
            }
        }
        


        private static void SetDateTimeFormat()
        {
            System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-US");
            System.Threading.Thread.CurrentThread.CurrentCulture.DateTimeFormat.ShortDatePattern = "yyyy-MM-dd";
            System.Globalization.CultureInfo culture = new System.Globalization.CultureInfo("en-US");
            culture.DateTimeFormat.ShortDatePattern = "yyyy-MM-dd";
        }
        public DataSet searchTimeOut(string sql)
        {
            System.Windows.Forms.Cursor.Current = Cursors.WaitCursor;
            DataSet ds = new DataSet();
            try
            {
                siInfo.Caption = "Processing....";
                using (System.Data.SqlClient.SqlConnection DC = new System.Data.SqlClient.SqlConnection(dc.ConnectionString))
                {
                    DC.Open();

                    using (System.Data.SqlClient.SqlDataAdapter DA = new System.Data.SqlClient.SqlDataAdapter(sql, DC))
                    {
                        DA.SelectCommand.CommandTimeout = 500;

                        DA.Fill(ds, "re");
                        sqlOk = true;
                        return ds;
                    }
                }
            }
            catch (Exception x)
            {
                MessageBox.Show(x.Message);
                sqlOk = false;
                return ds;
            }
            finally
            {
                System.Windows.Forms.Cursor.Current = Cursors.Default;
                siInfo.Caption = "Completed";
            }
        }
        #region menuSelect
        private void DailyReport_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            xtraTabControl1.SelectedTabPageIndex = 0;
            siStatus.Caption = "Daily Report : ";
        }

        private void barButtonItem1_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            xtraTabControl1.SelectedTabPageIndex = 1;
            siStatus.Caption = "Calculate Time : ";
        }
        private void iOffice_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            xtraTabControl1.SelectedTabPageIndex = 2;

            siStatus.Caption = "Office Staff : ";
        }
        private void iOp_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            xtraTabControl1.SelectedTabPageIndex = 3;
            siStatus.Caption = "Add Operator : ";
        }
        private void iNight_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            xtraTabControl1.SelectedTabPageIndex = 4;
            siStatus.Caption = "Night Shift Staff: ";
        }
        private void iLeave_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            xtraTabControl1.SelectedTabPageIndex = 5;
            siStatus.Caption = "Leave List: ";
        }
        private void iTime_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            xtraTabControl1.SelectedTabPageIndex = 6;
            siStatus.Caption = "Special Time: ";
        }
        #endregion
        public string Getdate()
        {
            return System.DateTime.Now.ToShortDateString();
        }
        #region CalTime
        private void simpleButton1_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Do not calculate at 11.30am - 13.30 pm And 23.30pm - 01.30am\nห้ามคำนวณในเวลา 11.30 - 13.30 และเวลา 23.30 - 01.30", "Inform", MessageBoxButtons.OK, MessageBoxIcon.Information);
            if (DialogResult.Yes == MessageBox.Show("Do you went calculate? \nYes = calculate \nNo = Do not think", "Please confirm.", MessageBoxButtons.YesNo, MessageBoxIcon.Question))
            {
                string sql = new SQLCmd(MF);
                sql = sql + string.Format(" @TYPE = 'DAILY_CAL'");
                if (textEdit1.Text.Trim() != string.Empty)
                    sql = sql + string.Format(",@STAFF = '{0}' ", textEdit1.Text.Trim());

                sql = sql + string.Format(",@DATE1 = '{0}' ", dateEdit1.Text.Trim());
                sql = sql + string.Format(",@DATE2 = '{0}' ", dateEdit2.Text.Trim());

                DataTable dt = searchTimeOut(sql).Tables[0];
                if (dt.Rows.Count > 0)
                {
                    MessageBox.Show("Calculate Success.", "Inform", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
        }

        #endregion

        #region DailyReport

        private void simpleButton2_Click(object sender, EventArgs e)
        {
            string sql = new SQLCmd(MF);
            sql = sql + string.Format(" @TYPE = 'DAILY_REPORT'");
            if (staffDaily.Text.Trim() != string.Empty)
                sql = sql + string.Format(",@STAFF = '{0}' ", staffDaily.Text.Trim());

            sql = sql + string.Format(",@DATE1 = '{0}' ", dateDaily1.Text.Trim());
            sql = sql + string.Format(",@DATE2 = '{0}' ", dateDaily2.Text.Trim());

            if (sexDailyCombo.Text == "Male")
            {
                sql = sql + string.Format(",@SEX = '{0}' ", "4%");
            }
            else
                if (sexDailyCombo.Text == "Female")
                {
                    sql = sql + string.Format(",@SEX = '{0}' ", "3%");
                }
                else
                    sql = sql + string.Format(",@SEX = '{0}' ", "%");


            if (shiftDailyCombo.Text == "Day")
            {
                sql = sql + string.Format(",@SHIFT = '{0}' ", "A");
            }
            else
                if (shiftDailyCombo.Text == "Night")
                {
                    sql = sql + string.Format(",@SHIFT = '{0}' ", "B");
                }
                else
                    sql = sql + string.Format(",@SHIFT = '{0}' ", "%");


            DataSet ds = searchTimeOut(sql);
            gridControl1.DataSource = ds.Tables[0];

            createTimeLineTable(gridView2, ds.Tables[1]);
            gridControl2.DataSource = ds.Tables[1];

            createTimeLineTable(gridView3, ds.Tables[2]);
            gridControl3.DataSource = ds.Tables[2];


            gridControl4.DataSource = ds.Tables[3];
            xtraTabControl2.SelectedTabPageIndex = 4;
        }
        private void gridView1_RowCellStyle(object sender, DevExpress.XtraGrid.Views.Grid.RowCellStyleEventArgs e)
        {
            if (e.RowHandle >= 0)
            {
                string shift = gridView1.GetRowCellDisplayText(e.RowHandle, gridView1.Columns["SHIFT"]);
                string shiftDate = gridView1.GetRowCellDisplayText(e.RowHandle, gridView1.Columns["SHIFT_DATE"]);
                string stat = gridView1.GetRowCellDisplayText(e.RowHandle, gridView1.Columns["STAT"]);
                string in1 = "";
                string out1 = "";
                string in2 = "";
                string out2 = "";
                if (gridView1.GetRowCellDisplayText(e.RowHandle, gridView1.Columns["IN1"]).Trim() != "")
                {
                    in1 = gridView1.GetRowCellDisplayText(e.RowHandle, gridView1.Columns["IN1"]).Split(' ')[1];
                    if ((Convert.ToDateTime(in1) >= Convert.ToDateTime("08:00:00") && shift == "A") || (Convert.ToDateTime(in1) >= Convert.ToDateTime("20:00:00") && shift == "B"))
                    {
                        if (e.Column.FieldName == "IN1")
                        {
                            e.Appearance.ForeColor = Color.Red;
                        }
                    }
                }
                if (gridView1.GetRowCellDisplayText(e.RowHandle, gridView1.Columns["OUT1"]).Trim() != "")
                {
                    out1 = gridView1.GetRowCellDisplayText(e.RowHandle, gridView1.Columns["OUT1"]).Split(' ')[1];
                    if (((Convert.ToDateTime(out1) < Convert.ToDateTime("11:30:00") && gridView1.GetRowCellDisplayText(e.RowHandle, gridView1.Columns["IN2"]).Trim() != "") || (Convert.ToDateTime(out1) < Convert.ToDateTime("17:00:00") && gridView1.GetRowCellDisplayText(e.RowHandle, gridView1.Columns["IN2"]).Trim() == "")) && shift == "A")
                    {
                        if (e.Column.FieldName == "OUT1")
                        {
                            e.Appearance.ForeColor = Color.Red;
                        }
                    }
                    else
                        if ((shift == "B" && Convert.ToDateTime(out1) < Convert.ToDateTime("05:00:00")) || (shift == "B" && Convert.ToDateTime(out1) > Convert.ToDateTime("21:00:00")))
                        {
                            if (e.Column.FieldName == "OUT1")
                            {
                                e.Appearance.ForeColor = Color.Red;
                            }
                        }
                }
                if (gridView1.GetRowCellDisplayText(e.RowHandle, gridView1.Columns["IN2"]).Trim() != "")
                {
                    in2 = gridView1.GetRowCellDisplayText(e.RowHandle, gridView1.Columns["IN2"]).Split(' ')[1];
                    if (Convert.ToDateTime(in2) > Convert.ToDateTime("13:00:00") && shift == "A")
                    {
                        if (e.Column.FieldName == "IN2")
                        {
                            e.Appearance.ForeColor = Color.Red;
                        }
                    }
                }
                if (gridView1.GetRowCellDisplayText(e.RowHandle, gridView1.Columns["OUT2"]).Trim() != "")
                {
                    out2 = gridView1.GetRowCellDisplayText(e.RowHandle, gridView1.Columns["OUT2"]).Split(' ')[1];
                    if (Convert.ToDateTime(out2) < Convert.ToDateTime("17:00:00") && shift == "A")
                    {
                        if (e.Column.FieldName == "OUT2")
                        {
                            e.Appearance.ForeColor = Color.Red;
                        }
                    }
                }
                if (shift.Trim() == "" && stat.Trim() == "A")
                {
                    e.Appearance.BackColor = Color.Pink;
                }
                else
                    if (stat.Trim() == "AA")
                    {
                        e.Appearance.BackColor = Color.Firebrick;
                    }
                    else
                        if (stat.Trim() == "N")
                        {
                            e.Appearance.BackColor = Color.LightSkyBlue;
                        }
                        else
                            if (stat.Trim() == "H")
                            {
                                e.Appearance.BackColor = Color.PaleGreen;
                            }
                            else
                                if (stat.Trim() == "Y")
                                {
                                    e.Appearance.BackColor = Color.FromArgb(0xFF, 0xFF, 0x99);
                                }
                                else if (stat.Trim() == "A" || stat.Trim() == "AN" || stat.Trim() == "M"
                                     || stat.Trim() == "MA" || stat.Trim() == "S" || stat.Trim() == "SU"
                                     || stat.Trim() == "A/2" || stat.Trim() == "AN/2" || stat.Trim() == "M/2"
                                     || stat.Trim() == "MA/2" || stat.Trim() == "S/2" || stat.Trim() == "SU/2")
                                {
                                    e.Appearance.BackColor = Color.PapayaWhip;
                                }
                                else if (stat.Trim() == "A#" || stat.Trim() == "AN#" || stat.Trim() == "M#"
                                    || stat.Trim() == "MA#" || stat.Trim() == "S#" || stat.Trim() == "SU#"
                                    || stat.Trim() == "A/2#" || stat.Trim() == "AN/2#" || stat.Trim() == "M/2#"
                                    || stat.Trim() == "MA/2#" || stat.Trim() == "S/2#" || stat.Trim() == "SU/2#")
                                {
                                    e.Appearance.BackColor = Color.Thistle;
                                }
                                else
                                    if (shift.Trim() == "" && stat.Trim() != "W" && stat.Trim() != "A" && stat.Trim() != "*R" && stat.Trim() != "+R" && stat.Trim() != "-R" && stat.Trim() != "W/*R" && stat.Trim() != "W/+R" && stat.Trim() != "W/-R")
                                        e.Appearance.BackColor = Color.SandyBrown;
                if (Convert.ToDateTime(shiftDate).DayOfWeek == DayOfWeek.Sunday)
                {

                    //e.Appearance.DrawString(e.Cache, e.DisplayText, e.Bounds, new Font(e.Appearance.Font, FontStyle.Bold), e.Appearance.TextOptions.GetStringFormat());
                    //e.Handled = true;
                }
            }
        }
        public void createTimeLineTable(GridView gv, DataTable dt)
        {
            gv.Columns.Clear();
            for (int i = 0; i < dt.Columns.Count; i++)
            {
                string columName = dt.Columns[i].ColumnName;

                DevExpress.XtraGrid.Columns.GridColumn gc = gv.Columns.Add();
                gc.Name = columName;
                gc.FieldName = columName;
                gc.Caption = columName;
                if (i == 0)
                {
                    gc.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
                    gc.Fixed = DevExpress.XtraGrid.Columns.FixedStyle.Left;
                    gc.AppearanceHeader.BackColor = Color.FromArgb(0xFF, 0xFF, 0x99);
                }
                else
                {
                    gc.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Far;
                    gc.Width = 90;
                    if (i % 2 == 0)
                    {
                        gc.AppearanceHeader.BackColor = Color.MistyRose;
                    }
                    else
                    {
                        gc.AppearanceHeader.BackColor = Color.LightCyan;
                    }
                }
                gc.Visible = true;
            }
        }
        public void setPrintEmployeeRe()
        {
            compositeLink1.Links.Clear();
            compositeLink1.PaperKind = System.Drawing.Printing.PaperKind.A3;
            compositeLink1.Landscape = true;
            PrintableComponentLink pn1 = new PrintableComponentLink();
            pn1.Component = gridControl13;
            PrintableComponentLink pn2 = new PrintableComponentLink();
            pn2.Component = chartControl1;
            PrintableComponentLink pn3 = new PrintableComponentLink();
            pn3.Component = gridControl14;

            Link linkGrid1Report = new Link();
            linkGrid1Report.CreateDetailArea += new CreateAreaEventHandler(linkGrid1Report_CreateDetailArea);
            compositeLink1.Links.Add(linkGrid1Report);
            compositeLink1.Links.Add(pn1);
            compositeLink1.Links.Add(pn2);
            compositeLink1.Links.Add(pn3);
        }
        public void setPrintEmployeeStatus()
        {
            compositeLink1.Links.Clear();
            compositeLink1.PaperKind = System.Drawing.Printing.PaperKind.A4;
            compositeLink1.Landscape = false;
            PrintableComponentLink pn1 = new PrintableComponentLink();
            pn1.Component = gridControl17;
            PrintableComponentLink pn2 = new PrintableComponentLink();
            pn2.Component = gridControl18;
            PrintableComponentLink pn3 = new PrintableComponentLink();
            pn3.Component = gridControl23;

            Link linkGrid1Report = new Link();
            linkGrid1Report.CreateDetailArea += new CreateAreaEventHandler(linkGrid1Report_CreateDetailArea23);
            compositeLink1.Links.Add(linkGrid1Report);
            compositeLink1.Links.Add(pn1);
            compositeLink1.Links.Add(pn2);
            compositeLink1.Links.Add(pn3);
        }
        void linkGrid1Report_CreateDetailArea(object sender, CreateAreaEventArgs e)
        {
            string text = string.Format("RESIGN REPORT {0}", prReYearTxt.Text.Trim());
            TextBrick h1 = e.Graph.DrawString(text, Color.Black, new RectangleF(0, 0, 500, 20), DevExpress.XtraPrinting.BorderSide.None);
            h1.Font = new System.Drawing.Font("Tahoma", 14, FontStyle.Underline);
            h1.StringFormat = new BrickStringFormat(StringAlignment.Near);
        }
        void linkGrid1Report_CreateDetailArea23(object sender, CreateAreaEventArgs e)
        {
            string text = string.Format("{0} {1}", "Daily report   show the status of employee", esDate1.DateTime.ToString("dd MMM yyyy"));
            TextBrick h1 = e.Graph.DrawString(text, Color.Green, new RectangleF(0, 0, 500, 20), DevExpress.XtraPrinting.BorderSide.None);
            h1.Font = new System.Drawing.Font("Tahoma", 14, FontStyle.Underline);
            h1.StringFormat = new BrickStringFormat(StringAlignment.Near);
        }
        private void cardView1_CustomDrawCardCaption(object sender, DevExpress.XtraGrid.Views.Card.CardCaptionCustomDrawEventArgs e)
        {
            e.CardCaption = cardView1.GetRowCellDisplayText(e.RowHandle, cardView1.Columns["STAFF"]);
        }

        private void iPrint_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            printOrSave(true);
        }
        public void printOrSave(bool print)
        {
            if (xtraTabControl1.SelectedTabPage == xtraTabPage17)
            {
                setPrintEmployeeRe();
            }
            if (xtraTabControl1.SelectedTabPage == xtraTabPage23)
            {
                setPrintEmployeeStatus();
            }
            gridColumn24.Visible = false;
            setPagePrint();
            if (!print)
            {
                Cursor = Cursors.WaitCursor;
                string filename = String.Format("{2} on {0:yyyyMMdd} to {1:yyyyMMdd}.xlsx",
                dateDaily1.DateTime,
                dateDaily2.DateTime,
                xtraTabControl2.SelectedTabPage.Text);
                using (SaveFileDialog sf = new SaveFileDialog())
                {
                    sf.DefaultExt = ".xlsx";
                    sf.Filter = "Microsoft Excel 2007 (*.xlsx)|*.xlsx";
                    sf.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
                    if (System.IO.File.Exists(sf.FileName))
                        filename = filename.Replace(".xlsx", "1.xlsx");


                    sf.FileName = filename;
                    if (sf.ShowDialog(this) == DialogResult.OK)
                    {
                        if (xtraTabControl1.SelectedTabPage == xtraTabPage17 || xtraTabControl1.SelectedTabPage == xtraTabPage23)
                        {
                            compositeLink1.CreateDocument();
                            compositeLink1.PrintingSystem.ExportToXlsx(sf.FileName);
                        }
                        else
                        {
                            printableComponentLink1.CreateDocument();
                            printableComponentLink1.PrintingSystem.ExportToXlsx(sf.FileName);
                        }
                        Cursor = Cursors.Default;
                        if (MessageBox.Show("Do you want open file now!", "save Completed", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                        {
                            System.Diagnostics.Process.Start(sf.FileName);
                        }
                    }
                    else
                    {
                        Cursor = Cursors.Default;
                    }

                    Cursor = Cursors.Default;
                }
            }
            else
            {
                if (xtraTabControl1.SelectedTabPage == xtraTabPage17 || xtraTabControl1.SelectedTabPage == xtraTabPage23)
                {
                    compositeLink1.CreateDocument();
                    compositeLink1.ShowPreviewDialog();
                }
                else
                {
                    printableComponentLink1.CreateDocument();
                    printableComponentLink1.ShowPreviewDialog();
                }
            }
            gridColumn24.Visible = true;
            gridView9.OptionsView.ShowViewCaption = true;
        }

        private void iExcel_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            printOrSave(false);
        }

        public string textHead()
        {
            string text = string.Format("DATE : {0} - {1}", dateDaily1.DateTime.ToString("dd/MM/yyyy"), dateDaily2.DateTime.ToString("dd/MM/yyyy"));

            if (xtraTabControl2.SelectedTabPage == xtraTabPage5 && staffDaily.Text.Trim() != "" && cardView1.RowCount == 1)
            {
                string staffName = cardView1.GetRowCellDisplayText(0, cardView1.Columns["NAME"]);
                string late = cardView1.GetRowCellDisplayText(0, cardView1.Columns["LATE"]);
                string ot = cardView1.GetRowCellDisplayText(0, cardView1.Columns["OT"]);
                string oth = cardView1.GetRowCellDisplayText(0, cardView1.Columns["OT_H"]);
                string night = cardView1.GetRowCellDisplayText(0, cardView1.Columns["SHIFT_B"]);
                string day = cardView1.GetRowCellDisplayText(0, cardView1.Columns["SHIFT_A"]);
                string workAll = cardView1.GetRowCellDisplayText(0, cardView1.Columns["DAY_COME"]);
                string breakOut = cardView1.GetRowCellDisplayText(0, cardView1.Columns["OUT_NOON"]);
                string ot3Hr = cardView1.GetRowCellDisplayText(0, cardView1.Columns["HR3"]);
                string workWi = cardView1.GetRowCellDisplayText(0, cardView1.Columns["WORK_WAI"]);


                text = string.Format("{0}   [Work Day(s) = {4}]  [Work full = {9}]  [OT = {3}]  [OT_H = {10}]  [Day Shift = {5}]  [Night Shift = {6}]  [OT 3 Hr = {8}]  [Break Out = {7}]  [Late = {2}]", text, staffName, late, ot, workAll, day, night, breakOut, ot3Hr, workWi, oth);
            }
            return text;
        }
        private void printableComponentLink1_CreateReportHeaderArea(object sender, CreateAreaEventArgs e)
        {
            float y = 15;
            if (xtraTabControl1.SelectedTabPageIndex == 0)
            {
                string text = xtraTabControl2.SelectedTabPage.Text;
                float ttd = 8.5F;
                TextBrick h1 = e.Graph.DrawString(text, Color.Navy, new RectangleF(0, 0, 500, 20), DevExpress.XtraPrinting.BorderSide.None);
                h1.Font = new System.Drawing.Font("Tahoma", 10, FontStyle.Bold);
                h1.StringFormat = new BrickStringFormat(StringAlignment.Near);
                text = textHead();

                TextBrick h2 = e.Graph.DrawString(text, Color.FromArgb(0x00, 0x66, 0x00), new RectangleF(0, y, 1000, 20), DevExpress.XtraPrinting.BorderSide.None);
                h2.Font = new System.Drawing.Font("Tahoma", ttd);
                h2.StringFormat = new BrickStringFormat(StringAlignment.Near);
                y += 10;

            }
            else
                if (xtraTabControl1.SelectedTabPage == xtraTabPage15)
                {
                    string text = "UTAX F.M. CO.,LTD. (EPZ)";
                    TextBrick h1 = e.Graph.DrawString(text, Color.ForestGreen, new RectangleF(0, 0, 500, 20), DevExpress.XtraPrinting.BorderSide.None);
                    h1.Font = new System.Drawing.Font("Tahoma", 14, FontStyle.Bold);
                    h1.StringFormat = new BrickStringFormat(StringAlignment.Near);
                    y += 11;

                    text = string.Format("REPORT EMPLOYEE AS {0}", reEmDate2.DateTime.ToString("MMMM dd, yyyy"));
                    TextBrick h2 = e.Graph.DrawString(text, Color.ForestGreen, new RectangleF(0, y, 500, 20), DevExpress.XtraPrinting.BorderSide.None);
                    h2.Font = new System.Drawing.Font("Tahoma", 14, FontStyle.Bold);
                    h2.StringFormat = new BrickStringFormat(StringAlignment.Near);
                    y += 30;
                }
                else
                    if (xtraTabControl1.SelectedTabPage == xtraTabPage16)
                    {
                        string text = string.Format("SUMMARIZE RESIGNATION {0}", srYearTxt.Text.Trim());
                        TextBrick h1 = e.Graph.DrawString(text, Color.Black, new RectangleF(0, 0, 500, 20), DevExpress.XtraPrinting.BorderSide.None);
                        h1.Font = new System.Drawing.Font("Tahoma", 14, FontStyle.Underline);
                        h1.StringFormat = new BrickStringFormat(StringAlignment.Near);
                        y += 11;
                    }
                    else
                        if (xtraTabControl1.SelectedTabPage == xtraTabPage17)
                        {
                            string text = string.Format("{0} {1}", xtraTabControl3.SelectedTabPage.Text, prReYearTxt.Text.Trim());
                            TextBrick h1 = e.Graph.DrawString(text, Color.Black, new RectangleF(0, 0, 500, 20), DevExpress.XtraPrinting.BorderSide.None);
                            h1.Font = new System.Drawing.Font("Tahoma", 14, FontStyle.Underline);
                            h1.StringFormat = new BrickStringFormat(StringAlignment.Near);
                            y += 11;
                        }
                        else
                            if (xtraTabControl1.SelectedTabPage == xtraTabPage23)
                            {
                                string text = string.Format("{0} {1}", "Daily report   show the status of employee", esDate1.Text.Trim());
                                TextBrick h1 = e.Graph.DrawString(text, Color.Green, new RectangleF(0, 0, 500, 20), DevExpress.XtraPrinting.BorderSide.None);
                                h1.Font = new System.Drawing.Font("Tahoma", 14, FontStyle.Underline);
                                h1.StringFormat = new BrickStringFormat(StringAlignment.Near);
                                y += 11;
                            }
                            else //if (xtraTabControl1.SelectedTabPage != xtraTabPage15)
                            {
                                string text = xtraTabControl1.SelectedTabPage.Text;

                                TextBrick h1 = e.Graph.DrawString(text, Color.Firebrick, new RectangleF(0, 0, 500, 20), DevExpress.XtraPrinting.BorderSide.None);
                                h1.Font = new System.Drawing.Font("Tahoma", 12, FontStyle.Bold);
                                h1.StringFormat = new BrickStringFormat(StringAlignment.Near);

                                if (xtraTabControl1.SelectedTabPage == xtraTabPage14)
                                {
                                    text = string.Format("DATE  : {0} - {1}", statusDate1.DateTime.ToString("dd/MM/yyyy"), statusDate2.DateTime.ToString("dd/MM/yyyy"));
                                    TextBrick h2 = e.Graph.DrawString(text, Color.Firebrick, new RectangleF(0, y, 500, 20), DevExpress.XtraPrinting.BorderSide.None);
                                    h2.Font = new System.Drawing.Font("Tahoma", 10);
                                    h2.StringFormat = new BrickStringFormat(StringAlignment.Near);
                                    y += 20;
                                }
                                else
                                    if (xtraTabControl1.SelectedTabPage == xtraTabPage28)
                                    {
                                        text = string.Format("DATE  : {0} - {1}", trainTotDate1.Text.Trim(), trainTotDate2.Text.Trim());
                                        TextBrick h2 = e.Graph.DrawString(text, Color.Firebrick, new RectangleF(0, y, 500, 20), DevExpress.XtraPrinting.BorderSide.None);
                                        h2.Font = new System.Drawing.Font("Tahoma", 10);
                                        h2.StringFormat = new BrickStringFormat(StringAlignment.Near);
                                        y += 20;
                                    }
                            }
        }

        #endregion

        public void setPagePrint()
        {
            if (xtraTabControl1.SelectedTabPageIndex == 0)
            {
                if (xtraTabControl2.SelectedTabPageIndex == 1)
                {
                    printableComponentLink1.Component = gridControl4;
                    printableComponentLink1.Landscape = true;
                    printableComponentLink1.PaperKind = System.Drawing.Printing.PaperKind.A4;
                }
                else
                    if (xtraTabControl2.SelectedTabPageIndex == 2)
                    {
                        printableComponentLink1.Component = gridControl3;
                        printableComponentLink1.Landscape = true;
                        printableComponentLink1.PaperKind = System.Drawing.Printing.PaperKind.A4;
                    }
                    else
                        if (xtraTabControl2.SelectedTabPageIndex == 3)
                        {
                            printableComponentLink1.Component = gridControl2;
                            printableComponentLink1.Landscape = true;
                            printableComponentLink1.PaperKind = System.Drawing.Printing.PaperKind.A4;
                        }
                        else
                            if (xtraTabControl2.SelectedTabPageIndex == 4)
                            {
                                printableComponentLink1.Component = gridControl1;
                                printableComponentLink1.Landscape = true;
                                printableComponentLink1.PaperKind = System.Drawing.Printing.PaperKind.A4;
                            }
            }
            else
                if (xtraTabControl1.SelectedTabPageIndex == 3)
                {
                    printableComponentLink1.Component = gridControl6;
                    printableComponentLink1.Landscape = false;
                    printableComponentLink1.PaperKind = System.Drawing.Printing.PaperKind.A4;
                }
                else
                    if (xtraTabControl1.SelectedTabPageIndex == 4)
                    {
                        printableComponentLink1.Component = gridControl7;
                        printableComponentLink1.Landscape = false;
                        printableComponentLink1.PaperKind = System.Drawing.Printing.PaperKind.A4;
                    }
                    else
                        if (xtraTabControl1.SelectedTabPageIndex == 5)
                        {
                            printableComponentLink1.Component = gridControl8;
                            printableComponentLink1.Landscape = false;
                            printableComponentLink1.PaperKind = System.Drawing.Printing.PaperKind.A4;
                        }
                        else
                            if (xtraTabControl1.SelectedTabPageIndex == 6)
                            {
                                printableComponentLink1.Component = gridControl9;
                                printableComponentLink1.Landscape = false;
                                printableComponentLink1.PaperKind = System.Drawing.Printing.PaperKind.A4;
                            }
                            else
                                if (xtraTabControl1.SelectedTabPage == xtraTabPage14)
                                {
                                    printableComponentLink1.Component = gridControl10;
                                    printableComponentLink1.Landscape = false;
                                    printableComponentLink1.PaperKind = System.Drawing.Printing.PaperKind.A4;
                                }
                                else
                                    if (xtraTabControl1.SelectedTabPage == xtraTabPage15)
                                    {
                                        printableComponentLink1.Component = gridControl11;
                                        printableComponentLink1.Landscape = false;
                                        gridView9.OptionsView.ShowViewCaption = false;
                                        printableComponentLink1.PaperKind = System.Drawing.Printing.PaperKind.A4;
                                    }
                                    else
                                        if (xtraTabControl1.SelectedTabPage == xtraTabPage16)
                                        {
                                            printableComponentLink1.Component = gridControl12;
                                            printableComponentLink1.Landscape = false;
                                            gridView10.OptionsView.ShowViewCaption = false;
                                            printableComponentLink1.PaperKind = System.Drawing.Printing.PaperKind.A4;
                                        }
                                        else
                                            if (xtraTabControl1.SelectedTabPage == xtraTabPage17)
                                            {
                                                if (xtraTabControl3.SelectedTabPageIndex == 0)
                                                {
                                                    printableComponentLink1.Component = gridControl13;
                                                    printableComponentLink1.Landscape = true;
                                                    printableComponentLink1.PaperKind = System.Drawing.Printing.PaperKind.A4;
                                                }
                                                if (xtraTabControl3.SelectedTabPageIndex == 1)
                                                {
                                                    printableComponentLink1.Component = chartControl1;
                                                    printableComponentLink1.Landscape = true;
                                                    printableComponentLink1.PaperKind = System.Drawing.Printing.PaperKind.A4;
                                                }
                                                if (xtraTabControl3.SelectedTabPageIndex == 2)
                                                {
                                                    printableComponentLink1.Component = gridControl14;
                                                    printableComponentLink1.Landscape = true;
                                                    printableComponentLink1.PaperKind = System.Drawing.Printing.PaperKind.A4;
                                                }
                                            }
                                            else
                                                if (xtraTabControl1.SelectedTabPage == xtraTabPage21)
                                                {
                                                    printableComponentLink1.Component = gridControl15;
                                                    printableComponentLink1.Landscape = true;
                                                    printableComponentLink1.PaperKind = System.Drawing.Printing.PaperKind.A4;
                                                }
                                                else
                                                    if (xtraTabControl1.SelectedTabPage == xtraTabPage22)
                                                    {
                                                        printableComponentLink1.Component = gridControl16;
                                                        printableComponentLink1.Landscape = true;
                                                        printableComponentLink1.PaperKind = System.Drawing.Printing.PaperKind.A4;
                                                    }
                                                    else
                                                        if (xtraTabControl1.SelectedTabPage == xtraTabPage24)
                                                        {
                                                            printableComponentLink1.Component = gridControl19;
                                                            printableComponentLink1.Landscape = true;
                                                            printableComponentLink1.PaperKind = System.Drawing.Printing.PaperKind.A4;
                                                        }
                                                        else
                                                            if (xtraTabControl1.SelectedTabPage == xtraTabPage25)
                                                            {
                                                                printableComponentLink1.Component = gridControl20;
                                                                printableComponentLink1.Landscape = true;
                                                                printableComponentLink1.PaperKind = System.Drawing.Printing.PaperKind.A3;
                                                            }
                                                            else
                                                                if (xtraTabControl1.SelectedTabPage == xtraTabPage26)
                                                                {
                                                                    printableComponentLink1.Component = gridControl21;
                                                                    printableComponentLink1.Landscape = false;
                                                                    printableComponentLink1.PaperKind = System.Drawing.Printing.PaperKind.A4;
                                                                }
                                                                else
                                                                    if (xtraTabControl1.SelectedTabPage == xtraTabPage27)
                                                                    {
                                                                        printableComponentLink1.Component = gridControl22;
                                                                        printableComponentLink1.Landscape = false;
                                                                        printableComponentLink1.PaperKind = System.Drawing.Printing.PaperKind.A4;
                                                                    }
                                                                    else
                                                                        if (xtraTabControl1.SelectedTabPage == xtraTabPage28)
                                                                        {
                                                                            printableComponentLink1.Component = gridControl24;
                                                                            printableComponentLink1.Landscape = false;
                                                                            printableComponentLink1.PaperKind = System.Drawing.Printing.PaperKind.A4;
                                                                        }
        }

        private void repositoryItemCheckEdit4_Click(object sender, EventArgs e)
        {
            string S1 = gridView5.GetRowCellValue(gridView5.FocusedRowHandle, "STAFF_ID").ToString();
            string S2 = gridView5.GetRowCellValue(gridView5.FocusedRowHandle, "NAME").ToString();
            string S3 = gridView5.GetRowCellValue(gridView5.FocusedRowHandle, "NAME2").ToString();
            setStateOffice(2);
            offIdtxt.Text = S1;
            offNametxt.Text = S2;
            offNametxt2.Text = S3;
        }

        private void simpleButton3_Click(object sender, EventArgs e)
        {
            string sql = new SQLCmd(MF);
            sql = sql + string.Format(" @TYPE = 'OFFICE_L'");

            DataTable dt = searchTimeOut(sql).Tables[0];
            gridControl5.DataSource = dt;
        }

        private void simpleButton5_Click(object sender, EventArgs e)
        {
            if (offNametxt.Text.Trim() == "" || !checkInt(offIdtxt.Text.Trim()))
            {
                MessageBox.Show("รูปแบบผิด กรุณาตรจสอบ", "error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            string sql = new SQLCmd(MF);
            sql = sql + string.Format(" @TYPE = 'OFFICE_A'");
            sql = sql + string.Format(",@STAFF = '{0}' ", offIdtxt.Text.Trim());
            sql = sql + string.Format(",@NAME = '{0}' ", offNametxt.Text.Trim());
            if (offNametxt2.Text.Trim() != "")
                sql = sql + string.Format(",@NAME2 = '{0}' ", offNametxt2.Text.Trim());

            DataTable dt = searchTimeOut(sql).Tables[0];
            gridControl5.DataSource = dt;
            setStateOffice(3);
            MessageBox.Show("Add or Edit Success.", "OK", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void simpleButton6_Click(object sender, EventArgs e)
        {
            if (DialogResult.Yes == MessageBox.Show(string.Format("Do you went remove {0} from Office staff ?", offIdtxt.Text.Trim()), "be sure?", MessageBoxButtons.YesNo, MessageBoxIcon.Question))
            {
                string sql = new SQLCmd(MF);
                sql = sql + string.Format(" @TYPE = 'OFFICE_D'");
                sql = sql + string.Format(",@STAFF = '{0}' ", offIdtxt.Text.Trim());
                DataTable dt = searchTimeOut(sql).Tables[0];
                gridControl5.DataSource = dt;
                setStateOffice(3);
                MessageBox.Show("Removed.", "OK", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
        public void setStateOffice(int state)
        {
            if (state == 1)//add b
            {
                groupControl3.Visible = true;
                offIdtxt.Text = string.Empty;
                offNametxt.Text = string.Empty;
                offNametxt2.Text = string.Empty;

                offIdtxt.Enabled = true;
                offNametxt.Enabled = true;
                offNametxt2.Enabled = true;
                simpleButton6.Enabled = false;
            }
            else
                if (state == 2)//edit
                {
                    groupControl3.Visible = true;
                    offIdtxt.Text = string.Empty;
                    offNametxt.Text = string.Empty;
                    offNametxt2.Text = string.Empty;

                    offIdtxt.Enabled = false;
                    offNametxt.Enabled = true;
                    offNametxt2.Enabled = true;
                    simpleButton6.Enabled = true;
                }
                else
                    if (state == 3)//back and fin
                    {
                        groupControl3.Visible = false;
                        offIdtxt.Text = string.Empty;
                        offNametxt.Text = string.Empty;
                        offNametxt2.Text = string.Empty;

                        offIdtxt.Enabled = false;
                        offNametxt.Enabled = true;
                        offNametxt2.Enabled = true;
                        simpleButton6.Enabled = true;
                    }
        }

        private void simpleButton4_Click(object sender, EventArgs e)
        {
            setStateOffice(1);
        }

        private void simpleButton7_Click(object sender, EventArgs e)
        {
            setStateOffice(3);
        }




        #region Add Op page

        public void setStateOp(int state)
        {
            if (state == 1)//add b
            {
                groupControl4.Visible = true;

                txtOpStaff.Text = string.Empty;
                txtOpFN.Text = string.Empty;
                txtOpLN.Text = string.Empty;
                comOpDepart.SelectedIndex = 0;
                comOpStatus.SelectedIndex = 0;
                dateEdit3.Text = string.Empty;

                dateEdit4.Enabled = false;
                dateEdit4.Text = string.Empty;

                comOpTeam.SelectedIndex = 0;
                unselectLook(lookDiviOp);
                unselectLook(lookPosiOp);

                posiChangeDate.Text = string.Empty;
                posiChangeDate.Enabled = true;
                checkPosi = 99;

                comOpReason.SelectedIndex = 0;

                comOpPreName.SelectedIndex = 0;
                txtOpFNt.Text = string.Empty;
                txtOpLNt.Text = string.Empty;

                txtOpStaff.Enabled = true;
                simpleButton9.Enabled = false;
                xtraTabControl1.SelectedTabPageIndex = 8;
            }
            else
                if (state == 2)//edit b
                {
                    groupControl4.Visible = true;

                    txtOpStaff.Text = string.Empty;
                    txtOpFN.Text = string.Empty;
                    txtOpLN.Text = string.Empty;
                    comOpDepart.SelectedIndex = 0;
                    comOpStatus.SelectedIndex = 0;
                    dateEdit3.Text = string.Empty;

                    comOpPreName.SelectedIndex = 0;
                    txtOpFNt.Text = string.Empty;
                    txtOpLNt.Text = string.Empty;

                    dateEdit4.Enabled = false;
                    dateEdit4.Text = string.Empty;

                    posiChangeDate.Text = string.Empty;
                    posiChangeDate.Enabled = false;
                    checkPosi = 99;

                    comOpTeam.SelectedIndex = 0;
                    unselectLook(lookDiviOp);
                    unselectLook(lookPosiOp);
                    comOpReason.SelectedIndex = 0;

                    txtOpStaff.Enabled = false;
                    simpleButton9.Enabled = true;
                    xtraTabControl1.SelectedTabPageIndex = 8;
                }
                else
                    if (state == 3)//back b
                    {
                        groupControl4.Visible = false;

                        txtOpStaff.Text = string.Empty;
                        txtOpFN.Text = string.Empty;
                        txtOpLN.Text = string.Empty;
                        comOpDepart.SelectedIndex = 0;
                        comOpStatus.SelectedIndex = 0;
                        dateEdit3.Text = string.Empty;

                        comOpTeam.SelectedIndex = 0;
                        unselectLook(lookDiviOp);
                        unselectLook(lookPosiOp);
                        comOpReason.SelectedIndex = 0;

                        comOpPreName.SelectedIndex = 0;
                        txtOpFNt.Text = string.Empty;
                        txtOpLNt.Text = string.Empty;

                        dateEdit4.Enabled = true;
                        dateEdit4.Text = string.Empty;

                        posiChangeDate.Text = string.Empty;
                        posiChangeDate.Enabled = false;
                        checkPosi = 99;

                        simpleButton9.Enabled = true;
                        xtraTabControl1.SelectedTabPageIndex = 3;
                    }
        }


        private void comOpStatus_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comOpStatus.SelectedIndex == 2)
            {
                dateEdit4.Enabled = true;
                comOpReason.Enabled = true;
            }
            else
            {
                dateEdit4.Enabled = false;
                comOpReason.Enabled = false;
                comOpReason.SelectedIndex = 0;
                dateEdit4.Text = string.Empty;
            }
        }

        private void simpleButton8_Click(object sender, EventArgs e)
        {
            setStateOp(3);
        }
        private void simpleButton11_Click(object sender, EventArgs e)
        {
            setStateOp(1);
        }

        private void simpleButton12_Click(object sender, EventArgs e)
        {
            string sql = new SQLCmd(MF);
            sql = sql + string.Format(" @TYPE = 'OP_L'");
            sql = sql + string.Format(",@STAFF_ADD = '{0}'", addOpStaffTxt.Text.Trim());


            DataTable dt = searchTimeOut(sql).Tables[0];
            gridControl6.DataSource = dt;
        }

        private void simpleButton9_Click(object sender, EventArgs e)
        {
            if (DialogResult.Yes == MessageBox.Show(string.Format("Do you went remove {0} from Employee staff ?", txtOpStaff.Text.Trim()), "be sure?", MessageBoxButtons.YesNo, MessageBoxIcon.Question))
            {
                string sql = new SQLCmd(MF);
                sql = sql + string.Format(" @TYPE = 'OP_D'");
                sql = sql + string.Format(",@STAFF = '{0}' ", txtOpStaff.Text.Trim());
                DataTable dt = searchTimeOut(sql).Tables[0];
                gridControl6.DataSource = dt;
                setStateOp(3);
                MessageBox.Show("Removed.", "OK", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
        public string gridDateToString(GridView gv, string f_name, string fomat)
        {
            try
            {
                string a = ((DateTime)gv.GetRowCellValue(gv.FocusedRowHandle, f_name)).ToString(fomat);
                return a;
            }
            catch
            {
                return string.Empty;
            }
        }
        private void repositoryItemCheckEdit1_Click(object sender, EventArgs e)
        {
            string staff = gridView4.GetRowCellValue(gridView4.FocusedRowHandle, "STAFF_ID").ToString();
            string fName = gridView4.GetRowCellValue(gridView4.FocusedRowHandle, "F_NAME").ToString();
            string lName = gridView4.GetRowCellValue(gridView4.FocusedRowHandle, "L_NAME").ToString();
            string ftName = gridView4.GetRowCellValue(gridView4.FocusedRowHandle, "T_F_NAME").ToString();
            string ltName = gridView4.GetRowCellValue(gridView4.FocusedRowHandle, "T_L_NAME").ToString();
            string ptName = gridView4.GetRowCellValue(gridView4.FocusedRowHandle, "PRE_NAME").ToString();

            string depasID = gridView4.GetRowCellValue(gridView4.FocusedRowHandle, "DEPARTMENT_ID").ToString();
            string sDate = gridDateToString(gridView4, "DATE_IN", "yyyy-MM-dd"); //((DateTime)gridView4.GetRowCellValue(gridView4.FocusedRowHandle, "DATE_IN")).ToString("yyyy-MM-dd");
            bool stayID = (bool)gridView4.GetRowCellValue(gridView4.FocusedRowHandle, "STAUS");
            string eDate = gridDateToString(gridView4, "DATE_OUT", "yyyy-MM-dd"); //((DateTime)gridView4.GetRowCellValue(gridView4.FocusedRowHandle, "DATE_OUT")).ToString("yyyy-MM-dd");
            string divi = gridView4.GetRowCellValue(gridView4.FocusedRowHandle, "DIVI_NAME").ToString();
            string posi = gridView4.GetRowCellValue(gridView4.FocusedRowHandle, "POSI_NAME").ToString();
            string reason = gridView4.GetRowCellValue(gridView4.FocusedRowHandle, "R_REASON").ToString();
            string team = gridView4.GetRowCellValue(gridView4.FocusedRowHandle, "TEAM").ToString();

            string diviID = gridView4.GetRowCellValue(gridView4.FocusedRowHandle, "DIVISION_ID").ToString();
            string posiID = gridView4.GetRowCellValue(gridView4.FocusedRowHandle, "POSITION_ID").ToString();

            //  barcode = ControlsUtils.ReadLookUpEdit(lookWorkK)[ComboMemberType.Value].ToString();
            setStateOp(2);

            txtOpStaff.Text = staff.Trim();
            txtOpFN.Text = fName.Trim();
            txtOpLN.Text = lName.Trim();

            comOpDepart.SelectedIndex = Convert.ToInt32(depasID.Trim());
            dateEdit3.Text = sDate;
            dateEdit4.Text = eDate;
            checkPosi = Convert.ToInt32(posiID);

            txtOpFNt.Text = ftName.Trim();
            txtOpLNt.Text = ltName.Trim();
            comOpPreName.SelectedIndex = comOpPreName.Properties.Items.IndexOf(ptName);
            lookDiviOp.EditValue = lookDiviOp.Properties.GetKeyValueByDisplayText(divi);
            lookPosiOp.EditValue = lookPosiOp.Properties.GetKeyValueByDisplayText(posi);
            comOpTeam.SelectedIndex = comOpTeam.Properties.Items.IndexOf(team);
            comOpReason.SelectedIndex = comOpReason.Properties.Items.IndexOf(reason);

            if (stayID)
                comOpStatus.SelectedIndex = 1;
            else
                comOpStatus.SelectedIndex = 2;
        }

        private void simpleButton10_Click(object sender, EventArgs e)
        {
            if (!checkInt(txtOpStaff.Text.Trim()) || txtOpFN.Text.Trim() == "" || txtOpLN.Text.Trim() == "" || comOpDepart.SelectedIndex == 0 || comOpStatus.SelectedIndex == 0 || dateEdit3.Text.Trim() == "" || comOpTeam.SelectedIndex == 0 || lookDiviOp.Text.Trim() == "" || lookPosiOp.Text.Trim() == "")
            {
                MessageBox.Show("รูปแบบผิด หรือลงข้อมูลไม่ครบ กรุณาตรจสอบ", "error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            if (txtOpFNt.Text.Trim() == "" || txtOpLNt.Text.Trim() == "" || comOpPreName.SelectedIndex == 0)
            {
                MessageBox.Show("ลงข้อมูลภาษาไทยไม่ครบ กรุณาตรจสอบ", "error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            if (comOpStatus.SelectedIndex == 2 && dateEdit4.Text.Trim() == "")
            {
                MessageBox.Show("พนักงานลาออกแล้วให้ลงวันที่ลาออกด้วย", "error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            if (posiChangeDate.Enabled && posiChangeDate.Text.Trim() == "")
            {
                MessageBox.Show("มีการย้ยตำแหน่งหรือลงข้อมูลพนักงานใหม่\nกรุณาใส่วันที่เปลี่ยนตำแหน่ง", "error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            string sql = new SQLCmd(MF);
            sql = sql + string.Format(" @TYPE = 'OP_A'");
            sql = sql + string.Format(",@STAFF = '{0}' ", txtOpStaff.Text.Trim());

            sql = sql + string.Format(",@NAME = '{0}' ", txtOpFN.Text.Trim());
            sql = sql + string.Format(",@NAME2 = '{0}' ", txtOpLN.Text.Trim());

            sql = sql + string.Format(",@NAME_T = '{0}' ", txtOpFNt.Text.Trim());
            sql = sql + string.Format(",@NAME_T_2 = '{0}' ", txtOpLNt.Text.Trim());
            sql = sql + string.Format(",@P_NAME_ID = {0}", comOpPreName.SelectedIndex);

            sql = sql + string.Format(",@DEPART = {0} ", comOpDepart.SelectedIndex);
            sql = sql + string.Format(",@DATE1 = '{0}' ", dateEdit3.Text.Trim());
            sql = sql + string.Format(",@TEAM = '{0}' ", comOpTeam.Text.Trim());
            sql = sql + string.Format(",@DIVI_ID = {0} ", ControlsUtils.ReadLookUpEdit(lookDiviOp)[ComboMemberType.Value].ToString());
            sql = sql + string.Format(",@POSI_ID = {0} ", ControlsUtils.ReadLookUpEdit(lookPosiOp)[ComboMemberType.Value].ToString());
            sql = sql + string.Format(",@REASON = '{0}' ", comOpReason.Text.Trim());

            if (posiChangeDate.Enabled)
                sql = sql + string.Format(",@DATE3 = '{0}' ", posiChangeDate.Text.Trim());
            if (comOpStatus.SelectedIndex != 0)
            {
                sql = sql + string.Format(",@STATUS = {0} ", comOpStatus.SelectedIndex % 2);
            }
            else
            {
                MessageBox.Show("กรุณาเลือก สถานะ ของพนักงาน", "error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            if (dateEdit4.Text.Trim() != "")
                sql = sql + string.Format(",@DATE2 = '{0}' ", dateEdit4.Text.Trim());
            DataTable dt = searchTimeOut(sql).Tables[0];
            gridControl6.DataSource = dt;
            setStateOp(3);
            MessageBox.Show("Add or Edit Success.", "OK", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }


        #endregion

        private void gridView4_RowCellStyle(object sender, RowCellStyleEventArgs e)
        {
            if (gridView4.GetRowCellDisplayText(e.RowHandle, gridView4.Columns["STATUS"]).Trim() == "RESIGNED")
            {
                //if (e.Column.FieldName == "OUT1")
                //{
                e.Appearance.ForeColor = Color.Red;
                // }

            }
        }

        private void buttonEdit1_Properties_ButtonClick(object sender, DevExpress.XtraEditors.Controls.ButtonPressedEventArgs e)
        {
            OpenFileDialog dlg = new OpenFileDialog();
            ButtonEdit butt = (ButtonEdit)(sender);
            dlg.Filter = "Excel Files|*.xlsx;*.xls;";
            if (dlg.ShowDialog() == DialogResult.OK)
            {
                //StreamReader sr = new StreamReader(dlg.FileName,System.Text.Encoding.UTF8  );
                //buttonEdit1.Text = sr.ReadToEnd;
                butt.Text = dlg.FileName;
            }
        }

        private void simpleButton14_Click(object sender, EventArgs e)
        {
            if (buttonEdit1.Text.Trim() == "")
            {
                MessageBox.Show("ยังไม่ได้เลือก ไฟล์ excel!", "Error buttonEdit1.text is null", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            if (txtUpSta.Text.Trim() == "")
            {
                MessageBox.Show("Staff ID is Emty.", "error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                txtUpSta.Focus();
                return;
            }

            if (!checkInt(txtUpSta.Text.Trim()) || txtUpSta.Text.Trim().Length != 5)
            {
                MessageBox.Show("Staff ID Fomat error.", "error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                txtUpSta.Focus();
                return;
            }
            readexcel();
        }
        public DataTable createTableNight()
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("staff", typeof(string));
            dt.Columns.Add("dateS", typeof(string));
            dt.Columns.Add("dateE", typeof(string));
            dt.Columns.Add("comment", typeof(string));

            return dt;
        }
        public DataTable createTableSpTime()
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("staff", typeof(string));
            dt.Columns.Add("dateS", typeof(string));
            dt.Columns.Add("dateE", typeof(string));
            dt.Columns.Add("type", typeof(string));
            dt.Columns.Add("time", typeof(string));


            return dt;
        }
        public void readexcel()
        {
            Microsoft.Office.Interop.Excel.Application excel = null;
            Microsoft.Office.Interop.Excel.Workbook wb = null;
            Microsoft.Office.Interop.Excel.Worksheet ws = null;

            try
            {
                Cursor.Current = Cursors.WaitCursor;
                object misValue = System.Reflection.Missing.Value;

                System.Data.DataTable dt = createTableNight();

                excel = new Microsoft.Office.Interop.Excel.ApplicationClass();
                wb = excel.Workbooks.Open(buttonEdit1.Text, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                //Workbook wb1 = excel.Workbooks.Open(buttonEdit1.Text,misValue,false,)
                ws = (Microsoft.Office.Interop.Excel.Worksheet)wb.Worksheets.get_Item(1);
                for (int i = 6; i < 307; i++)
                {
                    if ((((ws.UsedRange.Cells[i, 2] as Microsoft.Office.Interop.Excel.Range).Value2) ?? "").ToString().Trim() != "")
                    {
                        string staff = ((ws.UsedRange.Cells[i, 2] as Microsoft.Office.Interop.Excel.Range).Value2).ToString().Trim();
                        string dateStart = ((ws.UsedRange.Cells[i, 3] as Microsoft.Office.Interop.Excel.Range).Text).ToString().Trim();
                        string dateEnd = ((ws.UsedRange.Cells[i, 4] as Microsoft.Office.Interop.Excel.Range).Text).ToString().Trim();
                        string comment = (((ws.UsedRange.Cells[i, 5] as Microsoft.Office.Interop.Excel.Range).Value2) ?? "").ToString().Trim();


                        if (!checkDate(dateStart, "dd/MM/yyyy") || !checkDate(dateEnd, "dd/MM/yyyy"))
                        {
                            MessageBox.Show(string.Format("Date format Error at row {0}", i), "read Excel error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            return;
                        }
                        if ((DateTime.ParseExact(dateEnd, "dd/MM/yyyy", System.Globalization.CultureInfo.InvariantCulture)) < (DateTime.ParseExact(dateStart, "dd/MM/yyyy", System.Globalization.CultureInfo.InvariantCulture)))
                        {
                            MessageBox.Show(string.Format("End Date < Start Date at row {0}", i), "read Excel error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            return;
                        }

                        dt.Rows.Add(staff, dateStart, dateEnd, comment);
                    }
                }
                if (dt.Rows.Count < 1)
                {
                    MessageBox.Show("No data in file", "read Excel error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                if (DialogResult.Yes == MessageBox.Show("ตอนนี้อ่านไฟล์เสร็จแล้ว\nถ้สต้องการ Upload กด  Yes\nถ้าต้องการยกเลิกกด No", "ยืนยัน", MessageBoxButtons.YesNo, MessageBoxIcon.Question))
                {
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        string staff = dt.Rows[i]["staff"].ToString().Trim();
                        string date1 = dt.Rows[i]["dateS"].ToString().Trim();
                        string date2 = dt.Rows[i]["dateE"].ToString().Trim();
                        string comment = dt.Rows[i]["comment"].ToString().Trim();

                        updateStore(staff, date1, date2, comment);
                    }
                }
            }
            catch (Exception x)
            {
                MessageBox.Show(x.Message, "read Excel error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                Cursor.Current = Cursors.Default;

                wb.Close(false, null, null);
                excel.Quit();
                releaseObject(ws);
                releaseObject(wb);
                releaseObject(excel);
            }
        }
        public bool checkDate(string date, string fomat)
        {
            try
            {
                DateTime a = DateTime.ParseExact(date, fomat, System.Globalization.CultureInfo.InvariantCulture);
                //  Convert.ToDateTime(date);
                return true;
            }
            catch
            {
                return false;
            }
        }
        public void updateStore(string staff, string date1, string date2, string comment)
        {
            string sql = new SQLCmd(MF);
            sql = sql + string.Format(" @TYPE = 'SAVE_NIGHT'");
            sql = sql + string.Format(",@STAFF = '{0}' ", staff.Trim());
            sql = sql + string.Format(",@DATE1 = '{0}' ", date1.Trim());
            sql = sql + string.Format(",@DATE2 = '{0}' ", date2.Trim());
            sql = sql + string.Format(",@STAFF_ADD = '{0}' ", txtUpSta.Text.Trim());

            if (comment.Trim() != "")
                sql = sql + string.Format(",@COMMENT = '{0}' ", comment);
            searchTimeOut(sql);
        }
        public void updateStoreSpTime(string staff, string date1, string date2, string type, string time)
        {
            string[] a = time.Split(':');

            string sql = new SQLCmd(MF);
            sql = sql + string.Format(" @TYPE = 'SAVE_ST'");
            sql = sql + string.Format(",@STAFF = '{0}' ", staff.Trim());
            sql = sql + string.Format(",@DATE1 = '{0}' ", date1.Trim());
            sql = sql + string.Format(",@DATE2 = '{0}' ", date2.Trim());
            sql = sql + string.Format(",@STAFF_ADD = '{0}' ", txtSime.Text.Trim());
            sql = sql + string.Format(",@ST_TYPE = '{0}' ", type.Trim());
            sql = sql + string.Format(",@ST_TIME = '{0}' ", time.Trim());
            sql = sql + string.Format(",@ST_HR = {0} ", a[0]);
            sql = sql + string.Format(",@ST_MIN = {0} ", a[1]);

            searchTimeOut(sql);
        }
        public void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                MessageBox.Show("Unable to release the Object " + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }

        private void simpleButton13_Click(object sender, EventArgs e)
        {
            string sql = new SQLCmd(MF);
            sql = sql + string.Format(" @TYPE = 'SHOW_NIGHT'");
            sql = sql + string.Format(",@STAFF_ADD = '{0}' ", txtNightStaff.Text.Trim());
            sql = sql + string.Format(",@DATE1 = '{0}' ", dateEdit5.Text.Trim());
            sql = sql + string.Format(",@DATE2 = '{0}' ", dateEdit6.Text.Trim());

            DataSet ds = searchTimeOut(sql);
            gridControl7.DataSource = ds.Tables[0];
        }

        private void repositoryItemCheckEdit2_Click(object sender, EventArgs e)
        {
            string staff = gridView6.GetRowCellDisplayText(gridView6.FocusedRowHandle, gridView6.Columns["STAFF"]);
            string da = gridView6.GetRowCellDisplayText(gridView6.FocusedRowHandle, gridView6.Columns["SHIFT_DATE"]);
            string com = gridView6.GetRowCellDisplayText(gridView6.FocusedRowHandle, gridView6.Columns["COMMENT"]);

            textEdit2.Text = staff.Trim();
            textEdit3.Text = da.Trim();
            textEdit4.Text = com.Trim();

            groupControl7.Visible = true;
        }

        private void simpleButton15_Click(object sender, EventArgs e)
        {
            textEdit2.Text = string.Empty;
            textEdit3.Text = string.Empty;
            textEdit4.Text = string.Empty;

            groupControl7.Visible = false;
        }

        private void simpleButton17_Click(object sender, EventArgs e)
        {
            string sql = new SQLCmd(MF);
            sql = sql + string.Format(" @TYPE = 'EDIT_NIGHT'");
            sql = sql + string.Format(",@STAFF_ADD = '{0}' ", textEdit2.Text.Trim());
            sql = sql + string.Format(",@DATE1 = '{0}' ", textEdit3.Text.Trim());
            sql = sql + string.Format(",@COMMENT = '{0}' ", textEdit4.Text.Trim());

            searchTimeOut(sql);
            simpleButton15_Click(null, null);
            simpleButton13_Click(null, null);
        }

        private void simpleButton16_Click(object sender, EventArgs e)
        {
            if (DialogResult.Yes == MessageBox.Show("ต้องการจะลบหรือไม่", "ยืนยัน", MessageBoxButtons.YesNo, MessageBoxIcon.Question))
            {
                string sql = new SQLCmd(MF);
                sql = sql + string.Format(" @TYPE = 'DEL_NIGHT'");
                sql = sql + string.Format(",@STAFF_ADD = '{0}' ", textEdit2.Text.Trim());
                sql = sql + string.Format(",@DATE1 = '{0}' ", textEdit3.Text.Trim());

                searchTimeOut(sql);
                simpleButton15_Click(null, null);
                simpleButton13_Click(null, null);
            }
        }
        public void unselectLook(LookUpEdit COUN)
        {
            COUN.EditValue = null;
            COUN.Text = string.Empty;
            COUN.ClosePopup();
        }
        private void COUN_Properties_ButtonClick(object sender, DevExpress.XtraEditors.Controls.ButtonPressedEventArgs e)
        {
            if (e.Button.Kind == ButtonPredefines.Delete)
            {
                LookUpEdit COUN = (LookUpEdit)sender;
                unselectLook(COUN);
            }
        }

        private void lookUpEdit2_Properties_EditValueChanged_1(object sender, EventArgs e)
        {
            if (leLookType.Text.Trim() == "")
                return;
            string type = ControlsUtils.ReadLookUpEdit(leLookType)[ComboMemberType.Value].ToString();
            if (type == "R")
            {
                leReReason.Enabled = true;
            }
            else
            {
                leReReason.Enabled = false;
            }
            if (type == "AN")
            {
                leComEmer.Enabled = true;
            }
            else
            {
                leComEmer.Enabled = false;
                leComEmer.SelectedIndex = 1;
            }
        }

        private void simpleButton18_Click(object sender, EventArgs e)
        {
            string sql = new SQLCmd(MF);
            sql = sql + string.Format(" @TYPE = 'LOAD_LEA'");
            sql = sql + string.Format(",@STAFF_ADD = '{0}' ", textEdit5.Text.Trim());
            sql = sql + string.Format(",@DATE1 = '{0}' ", dateEdit7.Text.Trim());
            sql = sql + string.Format(",@DATE2 = '{0}' ", dateEdit8.Text.Trim());
            if (lookListBu.Text.Trim() != "")
            {
                sql = sql + string.Format(",@AB_CODE = '{0}' ", ControlsUtils.ReadLookUpEdit(lookListBu)[ComboMemberType.Value].ToString());
            }

            gridControl8.DataSource = searchTimeOut(sql).Tables[0];
        }
        public string revertComBoxtoS(int type, int val) //type 0 = emergency 1 = pay 2 = full half    3 send Note
        {
            if (type == 0)
            {
                if (val == 2)
                    return "E";
                else
                    if (val == 1)
                        return "N";
            }
            else if (type == 1)
            {
                if (val == 1)
                    return "P";
                else
                    if (val == 2)
                        return "N";
            }
            else if (type == 2)
            {
                if (val == 1)
                    return "F";
                else
                    if (val == 2)
                        return "H";
            }
            else if (type == 3)
            {
                if (val == 1)
                    return "N";
                else
                    if (val == 2)
                        return "Y";
            }
            return string.Empty;
        }
        public int convertComStoBox(int type, string str)
        {
            if (type == 0)
            {
                if (str == "E")
                    return 2;
                else
                    if (str == "N")
                        return 1;
            }
            else
                if (type == 1)
                {
                    if (str == "P")
                        return 1;
                    else
                        if (str == "N")
                            return 2;
                }
                else
                    if (type == 2)
                    {
                        if (str == "F")
                            return 1;
                        else
                            if (str == "H")
                                return 2;
                    }
                    else if (type == 3)
                    {
                        if (str == "Y")
                            return 2;
                        else
                            if (str == "N")
                                return 1;
                    }
            return 0;
        }
        private void repositoryItemCheckEdit3_CheckedChanged(object sender, EventArgs e)
        {
            string staff = gridView7.GetRowCellDisplayText(gridView7.FocusedRowHandle, gridView7.Columns["STAFF"]);
            string da = gridView7.GetRowCellDisplayText(gridView7.FocusedRowHandle, gridView7.Columns["DATE_SHIFT"]);
            string shife = gridView7.GetRowCellDisplayText(gridView7.FocusedRowHandle, gridView7.Columns["SHIFT"]);
            string u = gridView7.GetRowCellDisplayText(gridView7.FocusedRowHandle, gridView7.Columns["AB_CODE"]);

            string name = gridView7.GetRowCellDisplayText(gridView7.FocusedRowHandle, gridView7.Columns["F_NAME"]);
            string l_name = gridView7.GetRowCellDisplayText(gridView7.FocusedRowHandle, gridView7.Columns["L_NAME"]);
            string half = gridView7.GetRowCellDisplayText(gridView7.FocusedRowHandle, gridView7.Columns["HAFT"]);
            string emer = gridView7.GetRowCellDisplayText(gridView7.FocusedRowHandle, gridView7.Columns["EMERGENCY"]);
            string pay = gridView7.GetRowCellDisplayText(gridView7.FocusedRowHandle, gridView7.Columns["PAY"]);
            string rea = gridView7.GetRowCellDisplayText(gridView7.FocusedRowHandle, gridView7.Columns["R_REASON"]);
            string note = gridView7.GetRowCellDisplayText(gridView7.FocusedRowHandle, gridView7.Columns["NOTE_S"]);


            if (u == "R")
            {
                MessageBox.Show("ไม่สามารถ แก้ไขรายการ Resign ในหน้านี้ได้\nต้องไปแก้ไขในหน้า Add Employee", "Resign Edit", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            else
                if (u == "L")
                {
                    MessageBox.Show("ไม่สามารถ แก้ไขรายการ มาสายได้ \nต้องไปแก้ไขให้เแจ้ง IT", "Resign Edit", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

            setStateLe(0);

            leStaff.Text = staff.Trim();
            leStaff.Enabled = false;
            leDate.Text = da.Trim();
            leName.Text = string.Format("{0}  {1}", name, l_name);
            leReReason.Text = rea;
            leLookType.EditValue = u;
            // leLookType.EditValue = leLookType.Properties.GetKeyValueByDisplayText(u.Trim());
            int em = convertComStoBox(0, emer);
            int py = convertComStoBox(1, pay);
            int ha = convertComStoBox(2, half);
            int noteJa = convertComStoBox(3, note);
            leComEmer.SelectedIndex = em;
            leComPay.SelectedIndex = py;
            leComFull.SelectedIndex = ha;
            leComNote.SelectedIndex = noteJa;

            simpleButton20.Enabled = true;
            xtraTabControl1.SelectedTabPage = xtraTabPage13;
        }

        public void setStateLe(int stateCon)
        {
            if (stateCon == 0) //clear
            {
                leStaff.Text = string.Empty;
                leName.Text = string.Empty;
                leDate.Text = string.Empty;
                unselectLook(leLookType);
                leComFull.SelectedIndex = 0;
                leComEmer.SelectedIndex = 0;
                leComPay.SelectedIndex = 0;
                leDateRe.Text = string.Empty;
                leDateRe.Text = string.Empty;
                leComNote.SelectedIndex = 0;


                leStaff.Enabled = true;
                leName.Enabled = false;
                leReReason.Enabled = true;
                leDateRe.Enabled = true;
                leComEmer.Enabled = true;
            }
        }
        private void simpleButton22_Click(object sender, EventArgs e)
        {
            setStateLe(0);
            leReReason.Enabled = false;
            leDateRe.Enabled = false;
            leComEmer.Enabled = false;
            simpleButton20.Enabled = false;
            xtraTabControl1.SelectedTabPage = xtraTabPage13;
        }

        private void simpleButton19_Click(object sender, EventArgs e)
        {
            setStateLe(0);
            xtraTabControl1.SelectedTabPageIndex = 5;
            simpleButton20.Enabled = true;
        }

        public bool checkQutaLa(string abCode, string dateL, string staff, string em, string forh, string pay)
        {
            string sql = new SQLCmd(MF);
            sql = sql + string.Format(" @TYPE = 'CHECK_QUTA_LA'");
            sql = sql + string.Format(",@STAFF = '{0}' ", staff.Trim());
            sql = sql + string.Format(",@DATE1 = '{0}' ", dateL.Trim());
            sql = sql + string.Format(",@AB_CODE = '{0}' ", abCode.Trim());
            sql = sql + string.Format(",@EMERGENCY = '{0}' ", em.Trim());
            sql = sql + string.Format(",@FULL = '{0}' ", forh.Trim());

            sql = sql + string.Format(",@PAY = '{0}' ", pay.Trim());
            string ss = searchTimeOut(sql).Tables[0].Rows[0]["RES"].ToString();
            if (ss == "OK")
                return true;
            else
            {
                MessageBox.Show(ss, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
        }
        private void simpleButton21_Click(object sender, EventArgs e)
        {
            if (leStaff.Text.Trim() == "" || leDate.Text.Trim() == "" || leLookType.Text.Trim() == "" || leComFull.SelectedIndex == 0 || leComEmer.SelectedIndex == 0 || leComPay.SelectedIndex == 0 || leComNote.SelectedIndex == 0)
            {
                MessageBox.Show("ข้อมูลไม่ครบ", "Add error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            if (!checkQutaLa(ControlsUtils.ReadLookUpEdit(leLookType)[ComboMemberType.Value].ToString(), leDate.Text.Trim(), leStaff.Text.Trim(), revertComBoxtoS(0, leComEmer.SelectedIndex), revertComBoxtoS(2, leComFull.SelectedIndex), revertComBoxtoS(1, leComPay.SelectedIndex)))
            {
                return;
            }

            else
            {
                string sql = new SQLCmd(MF);
                sql = sql + string.Format(" @TYPE = 'ADD_LEA'");
                sql = sql + string.Format(",@STAFF_ADD = '{0}' ", leStaff.Text.Trim());
                // sql = sql + string.Format(",@SHIFT = '{0}' ", comboBoxEdit1.Text.Trim());
                sql = sql + string.Format(",@DATE1 = '{0}' ", leDate.Text.Trim());
                sql = sql + string.Format(",@AB_CODE = '{0}' ", ControlsUtils.ReadLookUpEdit(leLookType)[ComboMemberType.Value].ToString());
                sql = sql + string.Format(",@FULL= '{0}' ", revertComBoxtoS(2, leComFull.SelectedIndex));
                sql = sql + string.Format(",@EMERGENCY = '{0}' ", revertComBoxtoS(0, leComEmer.SelectedIndex));
                sql = sql + string.Format(",@PAY = '{0}' ", revertComBoxtoS(1, leComPay.SelectedIndex));
                sql = sql + string.Format(",@REASON = '{0}' ", leDate.Text.Trim());
                sql = sql + string.Format(",@NOTE = '{0}' ", revertComBoxtoS(3, leComNote.SelectedIndex));

                searchTimeOut(sql);
                simpleButton19_Click(null, null);
                simpleButton18_Click(null, null);
            }
        }

        private void simpleButton20_Click(object sender, EventArgs e)
        {
            if (DialogResult.Yes == MessageBox.Show("ต้องการจะลบหรือไม่", "ยืนยัน", MessageBoxButtons.YesNo, MessageBoxIcon.Question))
            {
                string sql = new SQLCmd(MF);
                sql = sql + string.Format(" @TYPE = 'DEL_LEA'");
                sql = sql + string.Format(",@STAFF_ADD = '{0}' ", leStaff.Text.Trim());
                sql = sql + string.Format(",@DATE1 = '{0}' ", leDate.Text.Trim());

                searchTimeOut(sql);
                simpleButton19_Click(null, null);
                simpleButton18_Click(null, null);
            }
        }

        private void simpleButton23_Click(object sender, EventArgs e)
        {
            if (buttonEdit2.Text.Trim() == "")
            {
                MessageBox.Show("ยังไม่ได้เลือก ไฟล์ excel!", "Error buttonEdit1.text is null", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            if (txtSime.Text.Trim() == "")
            {
                MessageBox.Show("Staff ID is Emty.", "error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                txtSime.Focus();
                return;
            }

            if (!checkInt(txtSime.Text.Trim()) || txtSime.Text.Trim().Length != 5)
            {
                MessageBox.Show("Staff ID Fomat error.", "error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                txtSime.Focus();
                return;
            }
            readexcelSpTime();
        }

        public void readexcelSpTime()
        {
            Microsoft.Office.Interop.Excel.Application excel = null;
            Microsoft.Office.Interop.Excel.Workbook wb = null;
            Microsoft.Office.Interop.Excel.Worksheet ws = null;

            try
            {
                Cursor.Current = Cursors.WaitCursor;
                int count = 0;
                object misValue = System.Reflection.Missing.Value;

                //  System.Data.DataTable dt = createTableSpTime();

                excel = new Microsoft.Office.Interop.Excel.ApplicationClass();
                wb = excel.Workbooks.Open(buttonEdit2.Text, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                //Workbook wb1 = excel.Workbooks.Open(buttonEdit1.Text,misValue,false,)
                ws = (Microsoft.Office.Interop.Excel.Worksheet)wb.Worksheets.get_Item(1);
                for (int i = 4; i < 307; i++)
                {
                    if ((((ws.UsedRange.Cells[i, 2] as Microsoft.Office.Interop.Excel.Range).Value2) ?? "").ToString().Trim() != "")
                    {
                        string date = ((ws.UsedRange.Cells[i, 2] as Microsoft.Office.Interop.Excel.Range).Text).ToString().Trim();
                        string staff = ((ws.UsedRange.Cells[i, 3] as Microsoft.Office.Interop.Excel.Range).Value2).ToString().Trim();
                        string type = ((ws.UsedRange.Cells[i, 4] as Microsoft.Office.Interop.Excel.Range).Value2).ToString().Trim();
                        string hr = (((ws.UsedRange.Cells[i, 5] as Microsoft.Office.Interop.Excel.Range).Value2) ?? "").ToString().Trim();

                        if (!checkDate(date, "yyyy-MM-dd"))
                        {
                            MessageBox.Show(string.Format("Date format Error at row {0}", i), "read Excel error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            return;
                        }
                        if (type.Trim() != "BE" && type.Trim() != "AF")
                        {
                            MessageBox.Show(string.Format("Type Input Time error in Row {0}", i), "read Excel error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            return;
                        }
                        if (!checkDouble(hr))
                        {
                            MessageBox.Show(string.Format("Hr input error in Row {0}", i), "read Excel error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            return;
                        }

                        string sql = new SQLCmd(MF);
                        sql = sql + string.Format(" @TYPE = 'ADD_SP'");
                        sql = sql + string.Format(",@STAFF= '{0}' ", staff);
                        sql = sql + string.Format(",@DATE1 = '{0}' ", date);
                        sql = sql + string.Format(",@AB_CODE = '{0}' ", type);
                        sql = sql + string.Format(",@HR = '{0}' ", hr);

                        DataSet ds = searchTimeOut(sql);
                        if (!sqlOk)
                        {
                            MessageBox.Show(string.Format("error at row {0} \n on Date = {1}", i, date), "read Excel error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            return;
                        }
                        count++;
                    }
                }
                MessageBox.Show(string.Format("OK {0} Rows ", count), "OK", MessageBoxButtons.OK, MessageBoxIcon.Information);
                simpleButton27_Click(null, null);
            }
            catch (Exception x)
            {
                MessageBox.Show(x.Message, "read Excel error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                Cursor.Current = Cursors.Default;

                wb.Close(false, null, null);
                excel.Quit();
                releaseObject(ws);
                releaseObject(wb);
                releaseObject(excel);
            }
        }

        private void simpleButton27_Click(object sender, EventArgs e)
        {
            string sql = new SQLCmd(MF);
            sql = sql + string.Format(" @TYPE = 'LOAD_SP'");
            sql = sql + string.Format(",@STAFF = '{0}' ", spStaffTxt.Text.Trim());
            sql = sql + string.Format(",@DATE1 = '{0}' ", spDate1.Text.Trim());
            sql = sql + string.Format(",@DATE2 = '{0}' ", spDate2.Text.Trim());

            DataSet ds = searchTimeOut(sql);
            gridControl9.DataSource = ds.Tables[0];
        }

        private void iTimeF_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            xtraTabControl1.SelectedTabPageIndex = 7;
        }

        private void atUpb_Click(object sender, EventArgs e)
        {
            if (atPass.Text.Trim() != "40075")
            {
                MessageBox.Show(string.Format("รหัสไม่ถูกต้อง"), "read Excel error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            if (buttonEdit3.Text.Trim() == "")
            {
                MessageBox.Show(string.Format("ไฟล์ยังว่างปล่าว เลือกไฟล์ด้วยฮร๊าฟฟฟ"), "read Excel error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            readexcelAddTime();
        }
        public void readexcelAddTime()
        {
            Microsoft.Office.Interop.Excel.Application excel = null;
            Microsoft.Office.Interop.Excel.Workbook wb = null;
            Microsoft.Office.Interop.Excel.Worksheet ws = null;

            try
            {
                Cursor.Current = Cursors.WaitCursor;

                object misValue = System.Reflection.Missing.Value;

                System.Data.DataTable dt = createTableSpTime();

                excel = new Microsoft.Office.Interop.Excel.ApplicationClass();
                wb = excel.Workbooks.Open(buttonEdit3.Text, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                //Workbook wb1 = excel.Workbooks.Open(buttonEdit1.Text,misValue,false,)
                ws = (Microsoft.Office.Interop.Excel.Worksheet)wb.Worksheets.get_Item(1);
                int count = 0;
                for (int i = 2; i < 307; i++)
                {
                    if ((((ws.UsedRange.Cells[i, 2] as Microsoft.Office.Interop.Excel.Range).Value2) ?? "").ToString().Trim() != "")
                    {
                        string staff = ((ws.UsedRange.Cells[i, 1] as Microsoft.Office.Interop.Excel.Range).Value2).ToString().Trim();
                        string dateShip = ((ws.UsedRange.Cells[i, 2] as Microsoft.Office.Interop.Excel.Range).Text).ToString().Trim();
                        string Ship = ((ws.UsedRange.Cells[i, 3] as Microsoft.Office.Interop.Excel.Range).Text).ToString().Trim();
                        string timeIn = ((ws.UsedRange.Cells[i, 4] as Microsoft.Office.Interop.Excel.Range).Text).ToString().Trim();

                        //string dateStart = ((ws.UsedRange.Cells[i, 3] as Microsoft.Office.Interop.Excel.Range).Text).ToString().Trim();
                        //string dateEnd = ((ws.UsedRange.Cells[i, 4] as Microsoft.Office.Interop.Excel.Range).Text).ToString().Trim();
                        //string type = (((ws.UsedRange.Cells[i, 5] as Microsoft.Office.Interop.Excel.Range).Value2) ?? "").ToString().Trim();
                        //string time = (((ws.UsedRange.Cells[i, 6] as Microsoft.Office.Interop.Excel.Range).Value2) ?? "").ToString().Trim();

                        if (!checkDate(dateShip, "yyyy-MM-dd"))
                        {
                            MessageBox.Show(string.Format("ShipDate format Error at row {0} \n staff = {1}", i, staff), "read Excel error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            return;
                        }
                        if (!checkCerSatff(staff))
                        {
                            MessageBox.Show(string.Format("ไม่มีรหัสพนักงานในระบบ หรือ พนักงานออกไปแล้ว at row {0} \n staff = {1}", i, staff), "read Excel error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            return;
                        }
                        string sql = new SQLCmd(MF);
                        sql = sql + string.Format(" @TYPE = 'ADD_ST'");
                        sql = sql + string.Format(",@STAFF_ADD = '{0}' ", staff);
                        sql = sql + string.Format(",@DATE1 = '{0}' ", dateShip);
                        sql = sql + string.Format(",@SHIFT = '{0}' ", Ship);
                        sql = sql + string.Format(",@DATETIME1 = '{0}' ", timeIn);

                        DataSet ds = searchTimeOut(sql);
                        if (!sqlOk)
                        {
                            MessageBox.Show(string.Format("error at row {0} \n staff = {1}", i, staff), "read Excel error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            return;
                        }
                        count++;
                    }
                }
                MessageBox.Show(string.Format("OK {0} Rows ", count), "OK", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception x)
            {
                MessageBox.Show(x.Message, "read Excel error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                Cursor.Current = Cursors.Default;

                wb.Close(false, null, null);
                excel.Quit();
                releaseObject(ws);
                releaseObject(wb);
                releaseObject(excel);
            }
        }

        public bool checkCerSatff(string staff)
        {
            try
            {
                string sql = new SQLCmd(MF);
                sql = sql + string.Format(" @TYPE = 'CHECK_CER'");
                sql = sql + string.Format(",@STAFF = '{0}' ", staff);

                DataTable dt = searchTimeOut(sql).Tables[0];
                if (dt.Rows.Count > 0)
                    return true;
                else
                    return false;
            }
            catch
            {
                return false;
            }
        }
        public bool checkAllSatff(string staff)
        {
            try
            {
                string sql = new SQLCmd(MF);
                sql = sql + string.Format(" @TYPE = 'CHECK_STAFF'");
                sql = sql + string.Format(",@STAFF = '{0}' ", staff);

                DataTable dt = searchTimeOut(sql).Tables[0];
                if (dt.Rows.Count > 0)
                    return true;
                else
                    return false;
            }
            catch
            {
                return false;
            }
        }

        private void leStaff_KeyDown(object sender, KeyEventArgs e)
        {
            try
            {
                if (leStaff.Text.Trim().Length == 5)
                {
                    string sql = new SQLCmd(MF);
                    sql = sql + string.Format(" @TYPE = 'GET_NAME'");
                    sql = sql + string.Format(",@STAFF = '{0}' ", leStaff.Text.Trim());

                    DataTable dt = searchTimeOut(sql).Tables[0];
                    if (dt.Rows.Count <= 0)
                        return;
                    else
                    {
                        leName.Text = dt.Rows[0]["NAME"].ToString();
                    }
                }
                else
                    leName.Text = string.Empty;
            }
            catch (Exception x)
            {
                MessageBox.Show(x.Message);

                leName.Text = string.Empty;
            }
        }
        public void readexcelAddLeave(string path)
        {
            Microsoft.Office.Interop.Excel.Application excel = null;
            Microsoft.Office.Interop.Excel.Workbook wb = null;
            Microsoft.Office.Interop.Excel.Worksheet ws = null;

            try
            {
                Cursor.Current = Cursors.WaitCursor;

                object misValue = System.Reflection.Missing.Value;

                System.Data.DataTable dt = createTableSpTime();

                excel = new Microsoft.Office.Interop.Excel.ApplicationClass();
                wb = excel.Workbooks.Open(path, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                //Workbook wb1 = excel.Workbooks.Open(buttonEdit1.Text,misValue,false,)
                ws = (Microsoft.Office.Interop.Excel.Worksheet)wb.Worksheets.get_Item(1);
                int count = 0;
                for (int i = 5; i < 1600; i++)
                {
                    if ((((ws.UsedRange.Cells[i, 1] as Microsoft.Office.Interop.Excel.Range).Value2) ?? "").ToString().Trim() != "")
                    {
                        string staff = ((ws.UsedRange.Cells[i, 1] as Microsoft.Office.Interop.Excel.Range).Value2).ToString().Trim();
                        string type = ((ws.UsedRange.Cells[i, 2] as Microsoft.Office.Interop.Excel.Range).Text).ToString().Trim();
                        string date = ((ws.UsedRange.Cells[i, 3] as Microsoft.Office.Interop.Excel.Range).Text).ToString().Trim();
                        string hf = ((ws.UsedRange.Cells[i, 5] as Microsoft.Office.Interop.Excel.Range).Text).ToString().Trim();
                        string en = ((ws.UsedRange.Cells[i, 6] as Microsoft.Office.Interop.Excel.Range).Text).ToString().Trim();
                        string pn = ((ws.UsedRange.Cells[i, 7] as Microsoft.Office.Interop.Excel.Range).Text).ToString().Trim();


                        if (!checkDate(date, "yyyy-MM-dd"))
                        {
                            MessageBox.Show(string.Format("Date format Error at row {0} \n staff = {1}  date = {2}", i, staff, date), "read Excel error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            return;
                        }
                        if (!checkAllSatff(staff))
                        {
                            MessageBox.Show(string.Format("ไม่มีรหัสพนักงานในระบบ  at row {0} \n staff = {1}   date = {2}", i, staff, date), "read Excel error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            return;
                        }

                        string sql = new SQLCmd(MF);
                        sql = sql + string.Format(" @TYPE = 'ADD_LEA'");
                        sql = sql + string.Format(",@STAFF_ADD = '{0}' ", staff.Trim());
                        sql = sql + string.Format(",@DATE1 = '{0}' ", date.Trim());
                        sql = sql + string.Format(",@AB_CODE = '{0}' ", type.Trim());
                        sql = sql + string.Format(",@FULL= '{0}' ", hf.Trim());
                        sql = sql + string.Format(",@EMERGENCY = '{0}' ", en.Trim());
                        sql = sql + string.Format(",@PAY = '{0}' ", pn.Trim());


                        DataSet ds = searchTimeOut(sql);
                        if (!sqlOk)
                        {
                            MessageBox.Show(string.Format("error at row {0} \n staff = {1}", i, staff), "read Excel error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            return;
                        }
                        count++;
                    }
                }
                MessageBox.Show(string.Format("OK {0} Rows ", count), "OK", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception x)
            {
                MessageBox.Show(x.Message, "read Excel error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                Cursor.Current = Cursors.Default;

                wb.Close(false, null, null);
                excel.Quit();
                releaseObject(ws);
                releaseObject(wb);
                releaseObject(excel);
            }
        }
        private void upLeUpbutton_Click(object sender, EventArgs e)
        {
            if (buttonEdit4.Text.Trim() == "")
            {
                MessageBox.Show("เลื่อกไฟล์ด้วย", "read Excel error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            if (upLestaffTxt.Text.Trim() == "")
            {
                MessageBox.Show("รหัสพนักงานว่างอยู่", "read Excel error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            readexcelAddLeave(buttonEdit4.Text.Trim());
        }

        private void gridView7_RowCellStyle(object sender, RowCellStyleEventArgs e)
        {
            if (e.RowHandle >= 0)
            {
                //  out1 = gridView1.GetRowCellDisplayText(e.RowHandle, gridView1.Columns["OUT1"]).Split(' ')[1];
                string u = gridView7.GetRowCellDisplayText(e.RowHandle, gridView7.Columns["AB_CODE"]);
                if (u == "R")
                    e.Appearance.ForeColor = Color.Red;
            }
        }

        private void rStatus_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            xtraTabControl1.SelectedTabPage = xtraTabPage14;
        }

        private void statusLoad_Click(object sender, EventArgs e)
        {
            string sql = new SQLCmd(MF);
            sql = sql + string.Format(" @TYPE = 'STATUS_REPORT'");
            sql = sql + string.Format(",@DATE1 = '{0}' ", statusDate1.Text.Trim());
            sql = sql + string.Format(",@DATE2 = '{0}' ", statusDate2.Text.Trim());

            DataTable dt = searchTimeOut(sql).Tables[0];
            gridControl10.DataSource = dt;
        }

        private void reReEmploy_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            xtraTabControl1.SelectedTabPage = xtraTabPage15;
        }

        private void reResignN_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            xtraTabControl1.SelectedTabPage = xtraTabPage16;
        }

        private void reResignPer_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            xtraTabControl1.SelectedTabPage = xtraTabPage17;
        }

        private void reSuehiro_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            xtraTabControl1.SelectedTabPage = xtraTabPage18;
        }

        private void reEmButt_Click(object sender, EventArgs e)
        {
            if (reEmDate2.Text.Trim() == "")
            {
                MessageBox.Show(string.Format("ไม่ได้กรอกข้อมูลวันที่\nกรุณากรอกวันที่"), "REPORT EMPLOYEE error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            gridView9.ViewCaption = string.Format("UTAX F.M. CO.,LTD. (EPZ)\nREPORT EMPLOYEE AS {0}", reEmDate2.DateTime.ToString("MMMM dd, yyyy"));

            string sql = new SQLCmd(MF);
            sql = sql + string.Format(" @TYPE = 'EMPLOYEE_REPORT'");
            //  sql = sql + string.Format(",@DATE1 = '{0}' ", reEmDate1.Text.Trim());
            sql = sql + string.Format(",@DATE2 = '{0}' ", reEmDate2.Text.Trim());

            DataTable dt = searchTimeOut(sql).Tables[0];
            gridControl11.DataSource = dt;
        }

        private void gridView9_RowCellStyle(object sender, RowCellStyleEventArgs e)
        {
            if (e.RowHandle >= 0)
            {
                //  out1 = gridView1.GetRowCellDisplayText(e.RowHandle, gridView1.Columns["OUT1"]).Split(' ')[1];
                string u = gridView9.GetRowCellDisplayText(e.RowHandle, gridView9.Columns["SECTION"]);
                string team = gridView9.GetRowCellDisplayText(e.RowHandle, gridView9.Columns["TEAM"]);
                if (u == "TOTAL")
                    e.Appearance.BackColor = Color.MistyRose;
                else
                {
                    if (e.Column == gridView9.Columns["TEAM"])
                    {
                        if (team == "A")
                            e.Appearance.BackColor = Color.LightSkyBlue;
                        else
                            if (team == "B")
                                e.Appearance.BackColor = Color.PaleGreen;
                    }
                }
            }
        }
        public void setColumn(DevExpress.XtraGrid.Columns.GridColumn gc, string name, string caption, string fielName, bool visible, int width, DevExpress.Utils.HorzAlignment st)
        {
            gc.Name = name;
            gc.FieldName = fielName;
            gc.Caption = caption;
            gc.Visible = visible;
            gc.Width = width;
            gc.AppearanceCell.TextOptions.HAlignment = st;
        }
        public void createGridColumnSR(GridView gv, DataTable monthList)
        {
            gv.Columns.Clear();
            DevExpress.XtraGrid.Columns.GridColumn sec = gv.Columns.Add();
            setColumn(sec, "srRsection", "SECTION", "SECTION", true, 200, DevExpress.Utils.HorzAlignment.Near);
            DevExpress.XtraGrid.Columns.GridColumn team = gv.Columns.Add();
            setColumn(team, "srRteam", "TEAM", "TEAM", true, 80, DevExpress.Utils.HorzAlignment.Center);

            foreach (DataRow dr in monthList.Rows)
            {
                string cap = dr["CAP"].ToString();
                string fName = dr["FNAME"].ToString();
                DevExpress.XtraGrid.Columns.GridColumn monC = gv.Columns.Add();
                setColumn(monC, string.Format("cg{0}", fName), cap, fName, true, 77, DevExpress.Utils.HorzAlignment.Center);
            }
            DevExpress.XtraGrid.Columns.GridColumn total = gv.Columns.Add();
            setColumn(total, "srRtotal", "TOTAL RESIGNED(UNIT:PERSON)", "TOTAL", true, 150, DevExpress.Utils.HorzAlignment.Center);
        }

        private void srSButton_Click(object sender, EventArgs e)
        {
            if (srYearTxt.Text.Trim() == "")
            {
                MessageBox.Show("กรุณาเลือกปีที่ต้องการค้นหา", "Summary Report error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                srYearTxt.Focus();
                return;
            }
            if (!checkInt(srYearTxt.Text.Trim()) || srYearTxt.Text.Trim().Length != 4)
            {
                MessageBox.Show("รูปแบบปีผิด", "Summary Report error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                srYearTxt.Focus();
                return;
            }
            string sql = new SQLCmd(MF);
            sql = sql + string.Format(" @TYPE = 'SUM_RESIGN'");

            sql = sql + string.Format(",@DATE1 = '{0}' ", srYearTxt.Text.Trim());
            DataSet ds = searchTimeOut(sql);
            DataTable dt = ds.Tables[1];
            createGridColumnSR(gridView10, ds.Tables[0]);
            gridControl12.DataSource = dt;
        }

        private void gridView10_RowCellStyle(object sender, RowCellStyleEventArgs e)
        {
            if (e.RowHandle >= 0)
            {
                //  out1 = gridView1.GetRowCellDisplayText(e.RowHandle, gridView1.Columns["OUT1"]).Split(' ')[1];
                string u = gridView10.GetRowCellDisplayText(e.RowHandle, gridView10.Columns["SECTION"]);
                string team = gridView10.GetRowCellDisplayText(e.RowHandle, gridView10.Columns["TEAM"]);
                if (u == "TOTAL")
                    e.Appearance.BackColor = Color.MistyRose;
                else
                {
                    if (e.Column == gridView10.Columns["TEAM"])
                    {
                        if (team == "A")
                            e.Appearance.BackColor = Color.LightSkyBlue;
                        else
                            if (team == "B")
                                e.Appearance.BackColor = Color.PaleGreen;
                    }
                }
            }
        }

        private void prReLoadButton_Click(object sender, EventArgs e)
        {
            if (prReYearTxt.Text.Trim() == "")
            {
                MessageBox.Show("กรุณาเลือกปีที่ต้องการค้นหา", "Summary Report error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                prReYearTxt.Focus();
                return;
            }
            if (!checkInt(prReYearTxt.Text.Trim()) || prReYearTxt.Text.Trim().Length != 4)
            {
                MessageBox.Show("รูปแบบปีผิด", "Summary Report error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                prReYearTxt.Focus();
                return;
            }
            string sql = new SQLCmd(MF);
            sql = sql + string.Format(" @TYPE = 'PER_RESIGN'");

            sql = sql + string.Format(",@DATE1 = '{0}' ", prReYearTxt.Text.Trim());
            DataSet ds = searchTimeOut(sql);
            DataTable dt = ds.Tables[0];
            gridControl13.DataSource = dt;
            gridControl14.DataSource = ds.Tables[1];
            chartControl1.DataSource = dt;
        }

        private void gridView11_RowCellStyle(object sender, RowCellStyleEventArgs e)
        {
            if (e.RowHandle + 1 == gridView11.RowCount || e.RowHandle + 2 == gridView11.RowCount)
            {
                if (e.Column != gridView11.Columns["HUM"])
                {
                    e.Appearance.BackColor = Color.MistyRose;
                }
            }
        }
        public void chChartType()
        {
        }

        private void checkEdit1_Properties_CheckedChanged(object sender, EventArgs e)
        {
            chartControl1.Series[0].Visible = checkEdit1.Checked;
        }

        private void checkEdit2_Properties_CheckedChanged(object sender, EventArgs e)
        {
            chartControl1.Series[1].Visible = checkEdit2.Checked;
        }

        private void checkEdit3_Properties_CheckedChanged(object sender, EventArgs e)
        {
            chartControl1.Series[2].Visible = checkEdit3.Checked;
        }
        public void addIncombo(ComboBoxEdit cb)
        {
            List<ViewType> myList = new List<ViewType>();
            myList.Add(ViewType.Bar);
            myList.Add(ViewType.StackedBar);
            myList.Add(ViewType.Bar3D);
            myList.Add(ViewType.Line);
            myList.Add(ViewType.StackedLine);
            myList.Add(ViewType.Line3D);
            myList.Add(ViewType.Pie);
            myList.Add(ViewType.Pie3D);
            myList.Add(ViewType.RadarArea);
            cb.Properties.Items.AddRange(myList);
            cb.SelectedIndex = 0;
        }

        private void comboBoxEdit1_SelectedIndexChanged(object sender, EventArgs e)
        {
            for (int i = 0; i < chartControl1.Series.Count; i++)
            {
                chartControl1.Series[i].ChangeView((ViewType)(comboBoxEdit1.EditValue));
            }
        }

        private void upHolidayBut_Click(object sender, EventArgs e)
        {
            if (buttonEdit5.Text.Trim() == "")
            {
                MessageBox.Show(string.Format("ไฟล์ยังว่างปล่าว เลือกไฟล์ด้วยฮร๊าฟฟฟ"), "read Excel error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            readexcelAddHoliday();
        }
        public void readexcelAddHoliday()
        {
            Microsoft.Office.Interop.Excel.Application excel = null;
            Microsoft.Office.Interop.Excel.Workbook wb = null;
            Microsoft.Office.Interop.Excel.Worksheet ws = null;

            try
            {
                Cursor.Current = Cursors.WaitCursor;

                object misValue = System.Reflection.Missing.Value;

                excel = new Microsoft.Office.Interop.Excel.ApplicationClass();
                wb = excel.Workbooks.Open(buttonEdit5.Text, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                //Workbook wb1 = excel.Workbooks.Open(buttonEdit1.Text,misValue,false,)
                ws = (Microsoft.Office.Interop.Excel.Worksheet)wb.Worksheets.get_Item(1);
                int count = 0;
                for (int i = 4; i < 150; i++)
                {
                    if ((((ws.UsedRange.Cells[i, 2] as Microsoft.Office.Interop.Excel.Range).Value2) ?? "").ToString().Trim() != "")
                    {
                        string status = ((ws.UsedRange.Cells[i, 3] as Microsoft.Office.Interop.Excel.Range).Value2).ToString().Trim();
                        string date = ((ws.UsedRange.Cells[i, 2] as Microsoft.Office.Interop.Excel.Range).Text).ToString().Trim();



                        if (!checkDate(date, "yyyy-MM-dd"))
                        {
                            MessageBox.Show(string.Format("HolidayDate format Error at row {0} \n on date = {1}", i, date), "read Excel error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            return;
                        }
                        string sql = new SQLCmd(MF);
                        sql = sql + string.Format(" @TYPE = 'ADD_HOLIDAY'");
                        sql = sql + string.Format(",@DATE1 = '{0}' ", date);
                        sql = sql + string.Format(",@AB_CODE = '{0}' ", status);

                        DataSet ds = searchTimeOut(sql);
                        if (!sqlOk)
                        {
                            MessageBox.Show(string.Format("error at row {0} \n on Date = {1}", i, date), "read Excel error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            return;
                        }
                        count++;
                    }
                }
                MessageBox.Show(string.Format("OK {0} Rows ", count), "OK", MessageBoxButtons.OK, MessageBoxIcon.Information);
                upHolidayLoadBut_Click(null, null);
            }
            catch (Exception x)
            {
                MessageBox.Show(x.Message, "read Excel error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                Cursor.Current = Cursors.Default;

                wb.Close(false, null, null);
                excel.Quit();
                releaseObject(ws);
                releaseObject(wb);
                releaseObject(excel);
            }
        }

        private void upHolidayLoadBut_Click(object sender, EventArgs e)
        {
            string sql = new SQLCmd(MF);
            sql = sql + string.Format(" @TYPE = 'LOAD_HOLIDAY'");
            sql = sql + string.Format(",@DATE1 = '{0}' ", upHolidayTxt.Text.Trim());
            DataSet ds = searchTimeOut(sql);
            gridControl15.DataSource = ds.Tables[0];
        }

        private void repositoryItemCheckEdit5_CheckedChanged(object sender, EventArgs e)
        {
            clcSpTime(true);

            string staff = gridView8.GetRowCellDisplayText(gridView8.FocusedRowHandle, gridView8.Columns["STAFF"]);
            string date = gridView8.GetRowCellDisplayText(gridView8.FocusedRowHandle, gridView8.Columns["DATE_SP"]);
            string type = gridView8.GetRowCellDisplayText(gridView8.FocusedRowHandle, gridView8.Columns["TYPE_SP"]);
            string hr = gridView8.GetRowCellDisplayText(gridView8.FocusedRowHandle, gridView8.Columns["HR_SP"]);

            spEStaffTxt.Text = staff;
            spEDateTxt.Text = date;
            spEHrTxt.Text = hr;
            spETypeCom.SelectedText = type;
        }
        public void clcSpTime(bool show)
        {
            spEStaffTxt.Text = string.Empty;
            spEDateTxt.Text = string.Empty;
            spEHrTxt.Text = string.Empty;
            spETypeCom.SelectedText = string.Empty;
            groupControl11.Visible = show;
        }
        private void spBackBut_Click(object sender, EventArgs e)
        {
            clcSpTime(false);
        }

        private void spBackSave_Click(object sender, EventArgs e)
        {
            string sql = new SQLCmd(MF);
            sql = sql + string.Format(" @TYPE = 'ADD_SP'");
            sql = sql + string.Format(",@STAFF= '{0}' ", spEStaffTxt.Text);
            sql = sql + string.Format(",@DATE1 = '{0}' ", spEDateTxt.Text);
            sql = sql + string.Format(",@AB_CODE = '{0}' ", spETypeCom.Text);
            sql = sql + string.Format(",@HR = '{0}' ", spEHrTxt.Text);

            DataSet ds = searchTimeOut(sql);
            if (!sqlOk)
            {
                MessageBox.Show("edit error", "Edit error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            else
            {
                simpleButton27_Click(null, null);
                clcSpTime(false);
            }
        }

        private void spDelBut_Click(object sender, EventArgs e)
        {
            if (DialogResult.Yes == MessageBox.Show("ต้องการจะลบรายการนี้หรือไม่", "ยืนยัน", MessageBoxButtons.YesNo, MessageBoxIcon.Question))
            {
                string sql = new SQLCmd(MF);
                sql = sql + string.Format(" @TYPE = 'DELETE_SP'");
                sql = sql + string.Format(",@STAFF= '{0}' ", spEStaffTxt.Text);
                sql = sql + string.Format(",@DATE1 = '{0}' ", spEDateTxt.Text);
                DataSet ds = searchTimeOut(sql);
                if (!sqlOk)
                {
                    MessageBox.Show("Delete error", "Delete error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                else
                {
                    simpleButton27_Click(null, null);
                    clcSpTime(false);
                }
            }
        }

        private void llReport_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            xtraTabControl1.SelectedTabPage = xtraTabPage22;
        }
        public void addColumLlReport(BandedGridView gv, DataTable dt)
        {
            if (dt.Rows.Count < 1)
                return;
            GridBand bTn = gv.Bands.AddBand("Leave_Data");
            bTn.Caption = "Leave Data";
            foreach (DataRow dr in dt.Rows)
            {
                string code = dr["AB_CODE"].ToString();
                string name = dr["AB_NAME"].ToString();

                BandedGridColumn C_id = gv.Columns.Add();
                C_id.Name = code;
                C_id.FieldName = code;
                C_id.Caption = name;
                C_id.Visible = true;
                C_id.SummaryItem.FieldName = code;
                C_id.Width = 20;

                C_id.SummaryItem.SummaryType = DevExpress.Data.SummaryItemType.Sum;
                C_id.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Far;
                //C_id.AppearanceHeader.BackColor = Color.FromArgb(0xFF, 0xCC, 0xCC);

                C_id.OwnerBand = bTn;
            }
        }

        private void llrLoadBut_Click(object sender, EventArgs e)
        {
            string sql = new SQLCmd(MF);
            sql = sql + string.Format(" @TYPE = 'LOAD_LEAVE'");
            sql = sql + string.Format(",@STAFF = '{0}' ", llrStaffTxt.Text.Trim());
            sql = sql + string.Format(",@DATE1 = '{0}' ", llrDate1.Text.Trim());
            sql = sql + string.Format(",@DATE2 = '{0}' ", llrDate2.Text.Trim());

            DataSet ds = searchTimeOut(sql);
            gridControl16.DataSource = ds.Tables[0];
        }

        private void seStatus_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            xtraTabControl1.SelectedTabPage = xtraTabPage23;
        }

        private void esLoadBut_Click(object sender, EventArgs e)
        {
            if (esDate1.Text.Trim() == "")
            {
                MessageBox.Show("กรุณาเลือกวันที่", "error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                esDate1.Focus();
                return;
            }
            string sql = new SQLCmd(MF);
            sql = sql + string.Format(" @TYPE = 'LOAD_STATUS_EM'");
            sql = sql + string.Format(",@DATE1 = '{0}' ", esDate1.Text.Trim());
            sql = sql + string.Format(",@AB_CODE = '{0}' ", radioGroup1.SelectedIndex);

            DataSet ds = searchTimeOut(sql);
            gridControl17.DataSource = ds.Tables[0];
            gridControl18.DataSource = ds.Tables[2];
            gridControl23.DataSource = ds.Tables[4];
            if (ds.Tables[1].Rows.Count > 0)
            {
                gridView14.ViewCaption = ds.Tables[1].Rows[0]["HEAD"].ToString();
            }
            else
                gridView14.ViewCaption = string.Empty;

            if (ds.Tables[3].Rows.Count > 0)
            {
                gridView13.ViewCaption = ds.Tables[3].Rows[0]["HEAD_MAIN"].ToString();
            }
            else
                gridView13.ViewCaption = string.Empty;
        }

        private void gridView13_RowCellStyle(object sender, RowCellStyleEventArgs e)
        {
            if (e.RowHandle >= 0)
            {
                string u = gridView13.GetRowCellDisplayText(e.RowHandle, gridView13.Columns["QQ"]);

                if (e.Column == gridView13.Columns["MAIN"])
                {
                    if ((u == "2" || u == "4" || u == "6" || u == "20" || u == "25" || u == "30"))
                    {
                        e.Appearance.BackColor = Color.Khaki;
                        e.Appearance.ForeColor = Color.DarkBlue;
                    }
                }
            }
        }

        private void anTotal_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            xtraTabControl1.SelectedTabPage = xtraTabPage24;
        }

        private void anQuota_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            xtraTabControl1.SelectedTabPage = xtraTabPage25;
        }

        public void addColumATR(BandedGridView gv, DataTable dt)
        {
            if (dt.Rows.Count < 1)
                return;
            if (gv.Bands.Count > 2)
                gv.Bands.RemoveAt(gv.Bands.Count - 1);

            GridBand bTn = gv.Bands.AddBand("MonthYear");
            bTn.Caption = "Year/Month";
            foreach (DataRow dr in dt.Rows)
            {
                string code = dr["DATE_SHIFT"].ToString();
                //string name = dr["AB_NAME"].ToString();

                BandedGridColumn C_id = gv.Columns.Add();
                C_id.Name = code;
                C_id.FieldName = code;
                C_id.Caption = code;
                C_id.Visible = true;
                C_id.SummaryItem.FieldName = code;
                C_id.Width = 50;

                C_id.SummaryItem.SummaryType = DevExpress.Data.SummaryItemType.Sum;
                C_id.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Far;
                //C_id.AppearanceHeader.BackColor = Color.FromArgb(0xFF, 0xCC, 0xCC);

                C_id.OwnerBand = bTn;
            }
        }
        private void atrLoadBut_Click(object sender, EventArgs e)
        {
            string sql = new SQLCmd(MF);
            sql = sql + string.Format(" @TYPE = 'LOAD_ATR'");
            sql = sql + string.Format(",@STAFF = '{0}' ", atrStaff.Text.Trim());
            sql = sql + string.Format(",@DATE1 = '{0}' ", atrDate1.Text.Trim());
            sql = sql + string.Format(",@DATE2 = '{0}' ", atrDate2.Text.Trim());

            DataSet ds = searchTimeOut(sql);
            if (ds.Tables[0].Rows.Count > 0)
            {
                addColumATR(bandedGridView4, ds.Tables[0]);
            }
            gridControl19.DataSource = ds.Tables[1];
        }

        private void simpleButton24_Click(object sender, EventArgs e)
        {
            if (DialogResult.Yes == MessageBox.Show("การ Calculate พักร้อน จะไปเขียนทับข้อมูลอันเดิมของปีที่เลือกมา\nกด Yes เพื่อยืนยันการ Calculate\nกด No เพื่อยกเลิก", "กรุณายืนยัน", MessageBoxButtons.YesNo, MessageBoxIcon.Question))
            {
                if (calAnYearTxt.Text.Trim() == "")
                {
                    MessageBox.Show("กรุณาเลือกปีที่จะทำการ Calculate", "Calculate Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    calAnYearTxt.Focus();
                    return;
                }
                if (!checkInt(calAnYearTxt.Text.Trim()))
                {
                    MessageBox.Show("รูปแบบปีไม่ถูกต้อง", "Calculate Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    calAnYearTxt.Focus();
                    return;
                }
                if (Convert.ToInt32(calAnYearTxt.Text.Trim()) < 2015 && Convert.ToInt32(calAnYearTxt.Text.Trim()) > 2099)
                {
                    MessageBox.Show("ช่วงปีอยู่ในระหว่าง 2015-2099", "Calculate Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    calAnYearTxt.Focus();
                    return;
                }
                string sql = new SQLCmd(MF);
                sql = sql + string.Format(" @TYPE = 'NEW_AN_QUOTA'");
                sql = sql + string.Format(",@DATE1 = '{0}' ", calAnYearTxt.Text.Trim());

                DataSet ds = searchTimeOut(sql);
                if (ds.Tables[0].Rows.Count > 0)
                {
                    MessageBox.Show("Calculate Complete", "Calculate Complete", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
        }
        public void addColumAn(BandedGridView gv, DataTable dt)
        {
            if (dt.Rows.Count < 1)
                return;
            if (gv.Bands.Count > 4)
            {
                gv.Bands.RemoveAt(gv.Bands.Count - 1);
                gv.Bands.RemoveAt(gv.Bands.Count - 1);
            }

            GridBand bTn = gv.Bands.AddBand("MonthYear");
            bTn.Caption = "Year/Month";
            foreach (DataRow dr in dt.Rows)
            {
                string codeQ = dr["MO"].ToString();
                string codeU = dr["MP"].ToString();
                string dis = dr["DIS"].ToString();

                BandedGridColumn C_id = gv.Columns.Add();
                C_id.Name = codeQ;
                C_id.FieldName = codeQ;
                C_id.Caption = dis;
                C_id.Visible = true;
                C_id.SummaryItem.FieldName = codeQ;
                C_id.Width = 50;
                C_id.AppearanceCell.BackColor = Color.LightCyan;
                C_id.SummaryItem.SummaryType = DevExpress.Data.SummaryItemType.Sum;
                C_id.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Far;
                C_id.OwnerBand = bTn;

                BandedGridColumn C_id2 = gv.Columns.Add();
                C_id2.Name = codeU;
                C_id2.FieldName = codeU;
                C_id2.Caption = dis;
                C_id2.Visible = true;
                C_id2.SummaryItem.FieldName = codeU;
                C_id2.Width = 50;
                C_id2.AppearanceCell.BackColor = Color.MistyRose;
                C_id2.SummaryItem.SummaryType = DevExpress.Data.SummaryItemType.Sum;
                C_id2.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Far;
                C_id2.OwnerBand = bTn;
            }
            GridBand tot = gv.Bands.AddBand("balance");
            bTn.Caption = "balance";

            BandedGridColumn ttl = gv.Columns.Add();
            ttl.Name = "TotalHave";
            ttl.FieldName = "SUM_Q";
            ttl.Caption = "Total Quota";
            ttl.Visible = true;
            ttl.SummaryItem.FieldName = "SUM_Q";
            ttl.Width = 50;
            ttl.SummaryItem.SummaryType = DevExpress.Data.SummaryItemType.Sum;
            ttl.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Far;
            ttl.OwnerBand = tot;

            BandedGridColumn total = gv.Columns.Add();
            total.Name = "TotalAn";
            total.FieldName = "SUM_U_AN";
            total.Caption = "Sum An Day";
            total.Visible = true;
            total.SummaryItem.FieldName = "SUM_U_AN";
            total.Width = 50;
            total.SummaryItem.SummaryType = DevExpress.Data.SummaryItemType.Sum;
            total.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Far;
            total.OwnerBand = tot;

            BandedGridColumn bal = gv.Columns.Add();
            bal.Name = "balance";
            bal.FieldName = "REMAIN_AN";
            bal.Caption = "balance";
            bal.Visible = true;
            bal.SummaryItem.FieldName = "REMAIN_AN";
            bal.Width = 50;
            bal.SummaryItem.SummaryType = DevExpress.Data.SummaryItemType.Sum;
            bal.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Far;
            bal.OwnerBand = tot;

            BandedGridColumn EX1 = gv.Columns.Add();
            EX1.Name = "EX1";
            EX1.FieldName = "EM1";
            EX1.Caption = "EM1";
            EX1.Visible = true;
            EX1.Width = 75;
            EX1.OwnerBand = tot;

            BandedGridColumn EX2 = gv.Columns.Add();
            EX2.Name = "EX2";
            EX2.FieldName = "EM2";
            EX2.Caption = "EM2";
            EX2.Visible = true;
            EX2.Width = 75;
            EX2.OwnerBand = tot;
        }
        private void anLoad_Click(object sender, EventArgs e)
        {
            if (!checkInt(loadAnYearTxt.Text.Trim()))
            {
                MessageBox.Show("รูปแบบปีไม่ถูกต้อง", "Calculate Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                loadAnYearTxt.Focus();
                return;
            }
            string sql = new SQLCmd(MF);
            sql = sql + string.Format(" @TYPE = 'LOAD_AN_QUOTA'");
            sql = sql + string.Format(",@DATE1 = '{0}' ", loadAnYearTxt.Text.Trim());
            sql = sql + string.Format(",@STAFF = '{0}' ", loadAnStaffTxt.Text.Trim());


            DataSet ds = searchTimeOut(sql);
            addColumAn(bandedGridView5, ds.Tables[1]);
            gridControl20.DataSource = ds.Tables[0];
        }

        private void bandedGridView5_RowCellStyle(object sender, RowCellStyleEventArgs e)
        {
            if (e.RowHandle >= 0)
            {
                string u = bandedGridView5.GetRowCellDisplayText(e.RowHandle, bandedGridView5.Columns["EX_DATE"]);

                if (u == "R")
                    e.Appearance.ForeColor = Color.Red;
            }
        }

        private void lookPosiOp_EditValueChanged(object sender, EventArgs e)
        {
            //posiChangeDate.Enabled = true;
            try
            {
                if (checkPosi != (Int32)(ControlsUtils.ReadLookUpEdit(lookPosiOp)[ComboMemberType.Value]))
                {
                    posiChangeDate.Enabled = true;
                }
                else
                    posiChangeDate.Enabled = false;
            }
            catch
            {
            }
        }

        private void rPohi_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            xtraTabControl1.SelectedTabPage = xtraTabPage26;
        }

        private void poLoadBut_Click(object sender, EventArgs e)
        {
            string sql = new SQLCmd(MF);
            sql = sql + string.Format(" @TYPE = 'POSI_HI_LOAD'");
            sql = sql + string.Format(",@STAFF = '{0}' ", poStaffTxt.Text.Trim());


            DataSet ds = searchTimeOut(sql);
            gridControl21.DataSource = ds.Tables[0];
        }

        private void rRasin_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            xtraTabControl1.SelectedTabPage = xtraTabPage27;
        }

        private void calRasinBut_Click(object sender, EventArgs e)
        {
            if (calRasinMonTxt.Text.Trim() == "")
            {
                MessageBox.Show("ใส่เดือนที่ต้องการจะ cal ด้วย", "Calculate Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                calRasinMonTxt.Focus();
                return;
            }
            string[] month = calRasinMonTxt.Text.Trim().Split('-');
            if (month[0].Length != 4 || Convert.ToInt32(month[1]) < 1 || Convert.ToInt32(month[1]) > 12 || !checkInt(month[0]))
            {
                MessageBox.Show("รูปแบบเดือนผิด", "Calculate Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                calRasinMonTxt.Focus();
                return;
            }
            string sql = new SQLCmd(MF);
            sql = sql + string.Format(" @TYPE = 'CAL_RASIN'");
            sql = sql + string.Format(",@DATE1 = '{0}-01' ", calRasinMonTxt.Text.Trim());


            DataSet ds = searchTimeOut(sql);
            if (ds.Tables[0].Rows.Count > 0)
            {
                MessageBox.Show("Calculate Complete", "Calculate Complete", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void loadRasinBut_Click(object sender, EventArgs e)
        {
            string sql = new SQLCmd(MF);
            sql = sql + string.Format(" @TYPE = 'LOAD_RASIN'");
            if (loadRasinMonTxt.Text.Trim() != "")
                sql = sql + string.Format(",@DATE1 = '{0}-01' ", loadRasinMonTxt.Text.Trim());
            sql = sql + string.Format(",@STAFF = '{0}' ", loadRasinStaffTxt.Text.Trim());

            DataSet ds = searchTimeOut(sql);
            gridControl22.DataSource = ds.Tables[0];
        }

        private void bandedGridView3_RowCellStyle(object sender, RowCellStyleEventArgs e)
        {
            if (bandedGridView3.GetRowCellDisplayText(e.RowHandle, bandedGridView3.Columns["STATUS_W"]).Trim() == "RESIGNED")
            {
                //if (e.Column.FieldName == "OUT1")
                //{
                e.Appearance.ForeColor = Color.Red;
                // }

            }
        }

        private void trainTotal_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            xtraTabControl1.SelectedTabPage = xtraTabPage28;
        }

        private void trainBP_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            xtraTabControl1.SelectedTabPage = xtraTabPage29;
        }

        private void trainList_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            xtraTabControl1.SelectedTabPage = xtraTabPage30;
        }

        private void trainUpload_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            xtraTabControl1.SelectedTabPage = xtraTabPage31;
        }

        private void trainUploadBut_Click(object sender, EventArgs e)
        {
            if (buttonEdit6.Text.Trim() == "")
            {
                MessageBox.Show(string.Format("ไฟล์ยังว่าง"), "read Excel error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                buttonEdit6.Focus();
                return;
            }
            uploadFileTraining();
        }
        public bool checkStaff(bool statusWork, string staff)
        {
            string sql = new SQLCmd(MF);
            sql = sql + string.Format(" @TYPE = 'CHECK_STAFF_AD'");
            sql = sql + string.Format(",@STAFF = '{0}' ", staff.Trim());
            if (statusWork)
                sql = sql + string.Format(",@STATUS = {0} ", 1);
            DataTable dt = searchTimeOut(sql).Tables[0];
            if (dt.Rows.Count > 0)
                return true;
            else
                return false;
        }
        public void uploadFileTraining()
        {
            Microsoft.Office.Interop.Excel.Application excel = null;
            Microsoft.Office.Interop.Excel.Workbook wb = null;
            Microsoft.Office.Interop.Excel.Worksheet ws = null;

            try
            {
                Cursor.Current = Cursors.WaitCursor;

                object misValue = System.Reflection.Missing.Value;



                excel = new Microsoft.Office.Interop.Excel.ApplicationClass();
                wb = excel.Workbooks.Open(buttonEdit6.Text, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                //Workbook wb1 = excel.Workbooks.Open(buttonEdit1.Text,misValue,false,)
                ws = (Microsoft.Office.Interop.Excel.Worksheet)wb.Worksheets.get_Item(1);

                string date = ((ws.UsedRange.Cells[3, 2] as Microsoft.Office.Interop.Excel.Range).Text).ToString().Trim();
                string place = ((ws.UsedRange.Cells[4, 2] as Microsoft.Office.Interop.Excel.Range).Text).ToString().Trim();
                string trainer = ((ws.UsedRange.Cells[6, 2] as Microsoft.Office.Interop.Excel.Range).Text).ToString().Trim();
                string divi = ((ws.UsedRange.Cells[6, 6] as Microsoft.Office.Interop.Excel.Range).Text).ToString().Trim();

                if (!checkDate(date, "yyyy-MM-dd"))
                {
                    MessageBox.Show(string.Format("Date format Error {0} ", date), "read Excel error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                if (place.Trim() == "" || trainer.Trim() == "" || divi.Trim() == "")
                {
                    MessageBox.Show(string.Format("ลงข้อมูล หัวเรื่องอบรมไม่ครบ"), "read Excel error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                List<string> textDet = new List<string>();
                int sumTime = 0;
                for (int i = 9; i < 15; i++)
                {
                    if ((((ws.UsedRange.Cells[i, 2] as Microsoft.Office.Interop.Excel.Range).Value2) ?? "").ToString().Trim() != "")
                    {
                        string num = ((ws.UsedRange.Cells[i, 1] as Microsoft.Office.Interop.Excel.Range).Value2).ToString().Trim();
                        string detail = ((ws.UsedRange.Cells[i, 2] as Microsoft.Office.Interop.Excel.Range).Text).ToString().Trim();
                        string time = ((ws.UsedRange.Cells[i, 6] as Microsoft.Office.Interop.Excel.Range).Text).ToString().Trim();

                        textDet.Add(detail.Trim());
                        if (checkInt(time))
                            sumTime = sumTime + Convert.ToInt32(time);
                    }
                }
                if (textDet.Count == 0)
                {
                    MessageBox.Show(string.Format("ไม่มีหัวข้อการอบรม"), "read Excel error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                if (sumTime == 0)
                {
                    MessageBox.Show(string.Format("เวลารวมในการอบรมเป็น 0"), "read Excel error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                List<string> listStaff = new List<string>();
                for (int i = 17; i < 150; i++)
                {
                    string staff = ((ws.UsedRange.Cells[i, 1] as Microsoft.Office.Interop.Excel.Range).Text).ToString().Trim();
                    if (staff.Trim() != "")
                    {
                        if (!checkAllSatff(staff))
                        {
                            MessageBox.Show(string.Format("ไม่มีรหัสพนักงาน {0} ในระบบ \nin Row {1} ", staff, i), "read Excel error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            return;
                        }
                        else
                        {
                            int reDouble = checkStackDouble(listStaff, staff);
                            if (reDouble != 999)
                            {
                                MessageBox.Show(string.Format("มีรหัสพนักงาน {0} ซ้ำ\nในแถวที่{1}\nี่ซ้ำกับแถวก่อนหน้าที่ {2} ", staff, i, reDouble), "read Excel error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                return;
                            }
                            listStaff.Add(staff);
                        }
                    }
                }
                if (listStaff.Count == 0)
                {
                    MessageBox.Show(string.Format("ไม่มีพนังงานที่อบรม"), "read Excel error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                //upload zone

                uplosdTrain(date, place, trainer, divi, textDet, sumTime, listStaff);
                MessageBox.Show("Update complete", "Update complete", MessageBoxButtons.OK, MessageBoxIcon.None);
            }
            catch (Exception x)
            {
                MessageBox.Show(x.Message, "read Excel error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                Cursor.Current = Cursors.Default;

                wb.Close(false, null, null);
                excel.Quit();
                releaseObject(ws);
                releaseObject(wb);
                releaseObject(excel);
            }
        }
        public void uplosdTrain(string date, string place, string trainer, string divi, List<string> det, int time, List<string> staff)
        {
            try
            {
                string sql = new SQLCmd(MF);
                sql = sql + string.Format(" @TYPE = 'TRAIN_UP_MAIN'");
                sql = sql + string.Format(",@DATE1 = '{0}' ", date.Trim());
                sql = sql + string.Format(",@TRAIN_BY = '{0}' ", trainer.Trim());
                sql = sql + string.Format(",@TRAIN_PLACE = '{0}' ", place.Trim());
                sql = sql + string.Format(",@TRAIN_DIVI = '{0}' ", divi.Trim());
                sql = sql + string.Format(",@COUNT_DET = {0} ", det.Count);
                sql = sql + string.Format(",@COUNT_STAFF = {0} ", staff.Count);
                sql = sql + string.Format(",@TOT_MIN = {0} ", time);
                DataTable dt = searchTimeOut(sql).Tables[0];
                if (dt.Rows.Count < 1)
                {
                    MessageBox.Show("upload main training error", "Uplaod error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                int kid = (int)(dt.Rows[0]["KID"]);
                for (int i = 0; i < det.Count; i++)
                {
                    string sqlDe = new SQLCmd(MF);
                    sqlDe = sqlDe + string.Format(" @TYPE = 'TRAIN_UP_DET'");
                    sqlDe = sqlDe + string.Format(",@TRAIN_ID = {0} ", kid);
                    sqlDe = sqlDe + string.Format(",@DETAIL = '{0}' ", det[i].Trim());
                    sqlDe = sqlDe + string.Format(",@DATE1 = '{0}' ", date.Trim());
                    sqlDe = sqlDe + string.Format(",@COUNT_DET = {0} ", i + 1);
                    searchTimeOut(sqlDe);//.Tables[0];
                }
                for (int i = 0; i < staff.Count; i++)
                {
                    string sqlDe = new SQLCmd(MF);
                    sqlDe = sqlDe + string.Format(" @TYPE = 'TRAIN_UP_STAFF'");
                    sqlDe = sqlDe + string.Format(",@TRAIN_ID = {0} ", kid);
                    sqlDe = sqlDe + string.Format(",@STAFF = '{0}' ", staff[i].Trim());
                    sqlDe = sqlDe + string.Format(",@DATE1 = '{0}' ", date.Trim());
                    sqlDe = sqlDe + string.Format(",@TOT_MIN = {0} ", time);
                    searchTimeOut(sqlDe);//.Tables[0];
                }
            }
            catch (Exception x)
            {
                MessageBox.Show(x.Message, "Uplaod error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        public int checkStackDouble(List<string> sta, string word)
        {
            for (int i = 0; i < sta.Count; i++)
            {
                if (sta[i] == word)
                {
                    return i;
                }
            }
            return 999;
        }

        private void trainTotBut_Click(object sender, EventArgs e)
        {
            trainSumMin.Text = string.Empty;
            trainCountM.Text = string.Empty;
            trainEqr.Text = string.Empty;

            string sql = new SQLCmd(MF);
            sql = sql + string.Format(" @TYPE = 'TRAIN_TOT_RE'");
            sql = sql + string.Format(",@STAFF_ADD = '{0}'", traTotStaffTxt.Text.Trim());
            sql = sql + string.Format(",@DATE1 = '{0}'", trainTotDate1.Text.Trim());
            sql = sql + string.Format(",@DATE2 = '{0}'", trainTotDate2.Text.Trim());
            DataSet ds = searchTimeOut(sql);
            DataTable dt = ds.Tables[0];
            gridControl24.DataSource = dt;
            DataTable dt2 = ds.Tables[1];
            string s1 = dt2.Rows[0]["TOTIME"].ToString();
            string s2 = dt2.Rows[0]["TOTCOUNT"].ToString();
            string s3 = dt2.Rows[0]["AVG_HO"].ToString();
            trainSumMin.Text = s1;
            trainCountM.Text = s2;
            trainEqr.Text = s3;
        }

        private void gridView18_RowCellStyle(object sender, RowCellStyleEventArgs e)
        {
            if (gridView18.GetRowCellDisplayText(e.RowHandle, gridView18.Columns["STATUS"]).Trim() == "RESIGNED")
            {
                //if (e.Column.FieldName == "OUT1")
                //{
                e.Appearance.ForeColor = Color.Red;
                // }

            }
        }

        private void barButtonItem2_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            xtraTabControl1.SelectedTabPage = xtrTabMoveGPZ;
        }

    

      

     

       

        private DataSet SeachGPZ(string staffid, DateTime date1, DateTime date2)
        {

            try
            {
                string sql = new SQLCmd(MF);
                sql = sql + string.Format(" @TYPE='SEARCH_GPZ'");
                sql = sql + string.Format(",@STAFF='{0}' ", staffid.Trim());
                sql = sql + string.Format(",@DATE1='{0}'", date1.ToString("yyyy-MM-dd"));
                sql = sql + string.Format(",@DATE2='{0}'", date2.ToString("yyyy-MM-dd"));

                return searchTimeOut(sql);

            }
            catch (Exception ex)
            {
                throw;
            }
        }
        //--------------------------------------
       
        //------------------------------
        public void updateGpz(string staff, string date1, string date2)
        {

            // if (CheckStafftoGpz(staff, date1, date2) == false)
            //  {
            string sql = new SQLCmd(MF);
            sql = sql + string.Format(" @TYPE = 'ADD_GPZ'");
            sql = sql + string.Format(",@STAFF = '{0}' ", staff.Trim());
            sql = sql + string.Format(",@DATE1 = '{0}' ", date1.Trim());
            sql = sql + string.Format(",@DATE2 = '{0}' ", date2.Trim());

            DataSet ds = searchTimeOut(sql);
            if (ds.Tables[0].Rows[0]["GG"].ToString() != "OK")
            {
                MessageBox.Show(ds.Tables[0].Rows[0]["GG"].ToString(), "read Excel error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            else
            {
                string res = ds.Tables[0].Rows[0]["RES"].ToString();
                string[] hiBoy = res.Split(',');
                if (hiBoy.Length <= 1 && hiBoy[0] == "")
                    if (hiBoy.Length <= 1 && hiBoy[0] == "")
                        if (hiBoy.Length <= 1 && hiBoy[0] == "")
                            if (hiBoy.Length <= 1 && hiBoy[0] == "")
                            {

                                MessageBox.Show("OK Ja", "update Complete", MessageBoxButtons.OK, MessageBoxIcon.None);

                            }
                            else
                            {
                                string showMeTheMoney = "";
                                for (int i = 0; i < hiBoy.Length; i++)
                                {
                                    showMeTheMoney = string.Format("{0}\n{1}", showMeTheMoney, hiBoy[i]);
                                }
                                MessageBox.Show(string.Format("มีรหัสซำดังนี้{0}", showMeTheMoney), "update Complete", MessageBoxButtons.OK, MessageBoxIcon.None);
                                //MessageBox.Show(string.Format("มีรหัสซำดังนี้{0}", res), "update Complete", MessageBoxButtons.OK, MessageBoxIcon.None);

                            }
            }
            //  }
            // else
            // {
            //       MessageBox.Show("แอดพนักงาน  "+staff+ " ไปเรียบร้อยแล้ว ","ok น่อ",MessageBoxButtons.OK,MessageBoxIcon.Information);
            //  }
        }

        private void btnGPZBrowseFile_Properties_ButtonClick(object sender, ButtonPressedEventArgs e)
        {
            OpenFileDialog dlg = new OpenFileDialog();
            ButtonEdit butt = (ButtonEdit)(sender);
            dlg.Filter = "Excel Files|*.xlsx;*.xls;";
            if (dlg.ShowDialog() == DialogResult.OK)
            {
                //StreamReader sr = new StreamReader(dlg.FileName,System.Text.Encoding.UTF8  );
                //buttonEdit1.Text = sr.ReadToEnd;
                butt.Text = dlg.FileName;
            }
        }
        public DataTable createTableGo2GPZ()
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("staff", typeof(string));
            dt.Columns.Add("dateS", typeof(string));
            dt.Columns.Add("dateE", typeof(string));


            return dt;
        }
        public Stack<string> stackError = new Stack<string>();
        public Boolean CheckStafftoGpz(string staffid)
        {

            string sql = new SQLCmd(MF);
            sql = sql + string.Format(" @TYPE='CHECK_STAFF_JA'");
            sql = sql + string.Format(",@STAFF='{0}' ", staffid);

            DataSet ds = searchTimeOut(sql);

            if (ds.Tables[0].Rows.Count > 0)
            {

                return true;
            }
            else
                return false;
        }

        



        //Function delete selected row
        private void DeleteSelectedRows(DevExpress.XtraGrid.Views.Grid.GridView view)
        {

            if (view == null || view.SelectedRowsCount == 0) return;



            DataRow[] rows = new DataRow[view.SelectedRowsCount];

            for (int i = 0; i < view.SelectedRowsCount; i++)

                rows[i] = view.GetDataRow(view.GetSelectedRows()[i]);



            view.BeginSort();

            try
            {

                foreach (DataRow row in rows)
                    
                    row.Delete();

            }

            finally
            {

                view.EndSort();

            }

        }

       

       

    }
}

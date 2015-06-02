using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Meteorite;
using Meteorite.Database;
using Meteorite.Files;
using DevExpress.XtraEditors;

namespace DailyReport
{
    public partial class uscMoveGPZ : UserControl
    {
        private Meteorite.Framework _mf;
        string DB_SERVER = string.Empty;
        string DB_NAME = string.Empty;
        string DB_USER = string.Empty;
        string DB_PASS = string.Empty;
        //public Framework MF { get; private set; }
        public bool sqlOk { get; private set; }
        public int checkPosi { get; set; }

        private IniReader ini { get; set; }
        public DatabaseConfiguration dc;

        public Meteorite.Framework MF
        {
            get { return _mf; }
            set { _mf = value; }
        }

        public uscMoveGPZ()
        {
            InitializeComponent();
            
        }

        public uscMoveGPZ(Meteorite.Framework mf): this()
        { this._mf = mf; }

        #region TEST
        private void test()
        {
            DBConnection conn = null;
            try
            {
                conn = MF.DbHelper.GetNewConnectionObject();
                conn.SelectQuery("select * from sdfadf");
                conn.Commit();

            }
            catch (Exception x)
            {
                DBHelper.Rollback(conn);
            }
            finally
            {
                DBHelper.CloseConn(conn);
            }
        }
        #endregion

        private void btnDeleteGPZ_Click(object sender, EventArgs e)
        {


            if (DialogResult.Yes == MessageBox.Show("ต้องการจะลบรายการนี้หรือไม่", "ยืนยัน", MessageBoxButtons.YesNo, MessageBoxIcon.Question))
            {
                try
                {
                   
                    if (gvItems == null || gvItems.SelectedRowsCount == 0) return;

                    //DataRow[] rows = new DataRow[gvItems.SelectedRowsCount];

                     DataRow[] rows = new DataRow[gvItems.SelectedRowsCount];
                    for (int i = 0; i < gvItems.SelectedRowsCount; i++)

                        rows[i] = gvItems.GetDataRow(gvItems.GetSelectedRows()[i]);



                    gvItems.BeginSort();

                    try
                    {

                        foreach (DataRow row in rows)
                        {
                            //gvItems.GetDataRow(e.FocusedRowHandle)["Name"].ToString();
                            string staff = gvItems.GetRowCellValue(gvItems.FocusedRowHandle, gvItems.Columns["STAFF_ID"]).ToString();
                            string date = gvItems.GetRowCellValue(gvItems.FocusedRowHandle, gvItems.Columns["DATE_GO"]).ToString();

                            string sql = new SQLCmd(MF);
                            sql = sql + string.Format(" @TYPE='DELETE_GPZ'");
                            sql = sql + string.Format(",@STAFF='{0} '", staff);
                            sql = sql + string.Format(",@DATE1='{0}'", Convert.ToDateTime(date).ToString("yyyy-MM-dd"));
                            searchTimeOut(sql);

                            row.Delete();
                        }
                    }

                    finally
                    {

                        gvItems.EndSort();

                    }



                  
                }
                catch (Exception ex)
                {
                    throw ex;
                }
                finally
                {
                    //update grid
                    //txtSearchStaffGpz.Text = "";
                    //dateSearchGpz1.Text = "";
                    //dateSearchGpz2.Text = "";

                }
            }
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
        private void btnSave2GPZ_Click(object sender, EventArgs e)
        {
            System.Data.DataTable dt = createTableGo2GPZ();


            if (txtSearchStaffGpz.Text.Trim() == "" || !checkInt(txtSearchStaffGpz.Text.Trim()))
            {
                MessageBox.Show("รูปแบบผิด กรุณาตรจสอบ", "error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            try
            {
                string staff = txtSearchStaffGpz.Text.Trim();
                string dateStart = dateSearchGpz1.Text.Trim();
                string dateEnd = dateSearchGpz2.Text.Trim();
                if (!CheckStafftoGpz(staff))
                {
                    MessageBox.Show(string.Format("ไม่มีรหัส พนักงาน {0} ระบบนี้\n ", staff), "read Excel error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;

                }
                else
                {
                    try
                    {// check has record or not ?

                        string sql1 = new SQLCmd(MF);
                        sql1 = sql1 + string.Format(" @TYPE='SEARCH_GPZ'");
                        sql1 = sql1 + string.Format(",@STAFF='{0} '", txtSearchStaffGpz.Text.Trim());
                        sql1 = sql1 + string.Format(",@DATE1='{0}'", dateSearchGpz1.Text.Trim());
                        sql1 = sql1 + string.Format(",@DATE2='{0}'", dateSearchGpz2.Text.Trim());
                         DataSet dss=    searchTimeOut(sql1);

                        //string fuu= dss.Tables[0].Rows[0].ToString();
                        if (dss != null && dss.Tables[0].Rows .Count > 0)
                            MF.ShowInfo("มีข้อมูลนี้ก่อนหน้าอยู่แล้ว ");
                        else
                        {


                            dt.Rows.Add(staff, dateStart, dateEnd);


                            // ds.Tables.Add(dt);
                            gridControl26.DataSource = dt;

                            string sql = new SQLCmd(MF);
                            sql = sql + string.Format(" @TYPE = 'ADD_GPZ'");
                            sql = sql + string.Format(", @STAFF = '{0}' ", txtSearchStaffGpz.Text.Trim());
                            sql = sql + string.Format(",@DATE1 = '{0}' ", dateSearchGpz1.Text.Trim());
                            sql = sql + string.Format(",@DATE2 = '{0}' ", dateSearchGpz2.Text.Trim());

                            DataSet ds = searchTimeOut(sql);
                            if (ds.Tables[0].Rows.Count > 0)
                            {
                                if (ds.Tables[0].Rows[0]["GG"].ToString() != "OK")
                                {
                                    MessageBox.Show(ds.Tables[0].Rows[0]["GG"].ToString(), "error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                    return;
                                }
                                else
                                {
                                    gridControl26.DataSource = ds.Tables[1];
                                    txtSearchStaffGpz.Text = "";
                                    dateSearchGpz1.Text = "";
                                    dateSearchGpz2.Text = "";
                                    MF.ShowInfo("บันทึกข้อมูลเรียบร้อยแล้ว");
                                }
                            }
                       }
                    }
                    catch(Exception x)
                    {
                        MF.ShowError(x);
                    
                    }
                }
                
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        private void btnGPZSearch_Click(object sender, EventArgs e)
        {
            //try
            //{

            //    throw new Exception("test");
            //}
            //catch (Exception x)
            //{
            //    MF.ShowError(x);
            //}
            try
            {
                DataSet ds = SeachGPZ(txtSearchStaffGpz.Text.Trim(), dateSearchGpz1.DateTime, dateSearchGpz2.DateTime);
                 if (ds != null && ds.Tables[0].Rows .Count > 0)
                    gridControl26.DataSource = ds.Tables[0];
                else
                    MessageBox.Show("ไม่มีข้อมูล");
            }
            catch (Exception x)
            {
                MF.ShowError(x);
            }
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

        private void btnGPZBrowseFile_Properties_ButtonClick(object sender, DevExpress.XtraEditors.Controls.ButtonPressedEventArgs e)
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


        public void uploadFileGo2GPZ()
        {
            Microsoft.Office.Interop.Excel.Application excel = null;
            Microsoft.Office.Interop.Excel.Workbook wb = null;
            Microsoft.Office.Interop.Excel.Worksheet ws = null;

            try
            {
                string staff;
                string dateStart;
                string dateEnd;
                Cursor.Current = Cursors.WaitCursor;
                object misValue = System.Reflection.Missing.Value;

                System.Data.DataTable dt = createTableGo2GPZ();
                DataSet ds = new DataSet();
                excel = new Microsoft.Office.Interop.Excel.ApplicationClass();
                wb = excel.Workbooks.Open(btnGPZBrowseFile.Text, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                //Workbook wb1 = excel.Workbooks.Open(buttonEdit1.Text,misValue,false,)
                ws = (Microsoft.Office.Interop.Excel.Worksheet)wb.Worksheets.get_Item(1);
                for (int i = 4; i < 100; i++)
                {
                    if ((((ws.UsedRange.Cells[i, 2] as Microsoft.Office.Interop.Excel.Range).Value2) ?? "").ToString().Trim() != "")
                    {
                        staff = ((ws.UsedRange.Cells[i, 2] as Microsoft.Office.Interop.Excel.Range).Value2).ToString().Trim();
                        dateStart = ((ws.UsedRange.Cells[i, 3] as Microsoft.Office.Interop.Excel.Range).Text).ToString().Trim();
                        dateEnd = ((ws.UsedRange.Cells[i, 4] as Microsoft.Office.Interop.Excel.Range).Text).ToString().Trim();
                        if (!CheckStafftoGpz(staff))
                        {
                            MessageBox.Show(string.Format("ไม่มีรหัส พนักงาน {0} ระบบนี้\nในแถวที่ {1}", staff, i), "read Excel error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            return;

                        }

                        dt.Rows.Add(staff, dateStart, dateEnd);

                    }
                    // ds.Tables.Add(dt);
                    gridControl27.DataSource = dt;
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
                        string staffs = dt.Rows[i]["staff"].ToString().Trim();
                        string date1 = dt.Rows[i]["dateS"].ToString().Trim();
                        string date2 = dt.Rows[i]["dateE"].ToString().Trim();


                        // check has record or not ?
                        DataSet dss = SeachGPZ(staffs.Trim(),Convert.ToDateTime( date1), Convert.ToDateTime(date2));
                        if (dss != null && dss.Tables[0].Rows.Count > 0)
                            MF.ShowInfo("มีข้อมูลนี้ก่อนหน้าอยู่แล้ว  ");
                        else
                        {
                            //no record start insert
                             updateGpz(staffs, date1, date2);
                            //  gridControl27.DataSource = dt.Rows.Add(staff, date1, date2); 
                        }
                    }

                    MessageBox.Show("บันทึกข้อมูลเรียบร้อย", "Upload Complete", MessageBoxButtons.OK, MessageBoxIcon.Error);
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


        public DataSet searchTimeOut(string sql)
        {//สำคัญ
            if (sql == string.Empty)
                return null;
            DBConnection conn = null;
            try
            {
                conn = MF.DbHelper.GetNewConnectionObject();
               DBResult rs= conn.SelectQuery(sql);
                
                conn.Commit();
                return rs.getDataSet();
            }
            catch (Exception x)
            {
                DBHelper.Rollback(conn);
                throw;
               
            }
            finally
            {
                DBHelper.CloseConn(conn);
            }
        }

        private void btnGPZUploadExcel_Click(object sender, EventArgs e)
        {
            

            if (btnGPZBrowseFile.Text.Trim() == "")
            {
                MessageBox.Show(string.Format("ไฟล์ยังว่าง"), "read Excel error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                btnGPZBrowseFile.Focus();
                return;
            }
            uploadFileGo2GPZ();


        }
    }
}

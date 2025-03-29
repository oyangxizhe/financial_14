using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;
using System.Reflection;
using XizheC;
using System.Net;
using System.Web;
using System.Xml;
using System.Collections;
using System.Data.OleDb;
using System.Web.UI;
using System.Web.UI.Adapters;
using System.Web.UI.HtmlControls;
using System.Web.Util;



namespace CSPSS.VOUCHER_MANAGE
{
    public partial class VOUCHERT : Form
    {
        DataTable dt = new DataTable();
        DataTable dt2 = new DataTable();
        DataTable dt3 = new DataTable();
        StringBuilder sqb = new StringBuilder();
        private string _ACID;
        public string ACID
        {
            set { _ACID = value; }
            get { return _ACID; }

        }
        private string _WATER_MARK_CONTENT;
        public string WATER_MARK_CONTENT
        {
            set { _WATER_MARK_CONTENT = value; }
            get { return _WATER_MARK_CONTENT; }
        }
        private string _old_file_name;
        public string old_file_name
        {
            set { _old_file_name = value; }
            get { return _old_file_name; }

        }
        private string _NEW_FILE_NAME;
        public string NEW_FILE_NAME
        {
            set { _NEW_FILE_NAME = value; }
            get { return _NEW_FILE_NAME; }

        }
        private string _ACCOUNTING_PERIOD_START_DATE;
        public string ACCOUNTING_PERIOD_START_DATE
        {
            set { _ACCOUNTING_PERIOD_START_DATE = value; }
            get { return _ACCOUNTING_PERIOD_START_DATE; }

        }
        private string _ACCOUNTING_PERIOD_EXPIRATION_DATE;
        public string ACCOUNTING_PERIOD_EXPIRATION_DATE
        {
            set { _ACCOUNTING_PERIOD_EXPIRATION_DATE = value; }
            get { return _ACCOUNTING_PERIOD_EXPIRATION_DATE; }

        }
        private string _IDO;
        public string IDO
        {
            set { _IDO = value; }
            get { return _IDO; }

        }
        private string _INITIAL_OR_OTHER;
        public string INITIAL_OR_OTHER
        {
            set { _INITIAL_OR_OTHER = value; }
            get { return _INITIAL_OR_OTHER; }
        }
        private string _ADD_OR_UPDATE;
        public string ADD_OR_UPDATE
        {
            set { _ADD_OR_UPDATE = value; }
            get { return _ADD_OR_UPDATE; }
        }
        private bool _IFExecutionSUCCESS;
        public bool IFExecution_SUCCESS
        {
            set { _IFExecutionSUCCESS = value; }
            get { return _IFExecutionSUCCESS; }

        }
        private static bool _IF_DOUBLE_CLICK;
        public static bool IF_DOUBLE_CLICK
        {
            set { _IF_DOUBLE_CLICK = value; }
            get { return _IF_DOUBLE_CLICK; }

        }
        protected int i, j;
        protected int M_int_judge, t;
        basec bc = new basec();
        CVOUCHER vou = new CVOUCHER();
        ExcelToCSHARP etc = new ExcelToCSHARP();
        CFileInfo cfileinfo = new CFileInfo();
        //BaseInfo.FrmCurrency cur = new CSPSS.BASE_INFO.FrmCurrency();
        VOUCHER F1 = new VOUCHER();
        Color c2 = System.Drawing.ColorTranslator.FromHtml("#990033");
        public VOUCHERT()
        {
            InitializeComponent();
        }
        public VOUCHERT(VOUCHER Frm)
        {
            InitializeComponent();
            F1 = Frm;
        }
        private void VOUCHERT_Load(object sender, EventArgs e)
        {

          
            try
            {
                bind();
               
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "提示", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);
            }
        }

        #region bind
        private void bind()
        {

            label52.Text = "";
            label53.Visible = false;
            label55.Visible = false;
            label56.Visible = false;
            label57.Visible = false;
            progressBar1.Visible = false;
            string v1 = bc.getOnlyString("SELECT ADD_NEW FROM RIGHTLIST WHERE  USID='" + LOGIN.USID + "' AND NODE_NAME='录入凭证作业'");
            string v2 = bc.getOnlyString("SELECT EDIT FROM RIGHTLIST WHERE  USID='" + LOGIN.USID + "' AND NODE_NAME='录入凭证作业'");
            string v3 = bc.getOnlyString("SELECT DEL FROM RIGHTLIST WHERE  USID='" + LOGIN.USID + "' AND NODE_NAME='录入凭证作业'");

            string v4 = bc.getOnlyString("SELECT MANAGE FROM RIGHTLIST WHERE  USID='" + LOGIN.USID + "' AND NODE_NAME='录入凭证作业'");
            string v5 = bc.getOnlyString("SELECT FINANCIAL FROM RIGHTLIST WHERE  USID='" + LOGIN.USID + "' AND NODE_NAME='录入凭证作业'");
            string v6 = bc.getOnlyString("SELECT GENERAL_MANAGE FROM RIGHTLIST WHERE  USID='" + LOGIN.USID + "' AND NODE_NAME='录入凭证作业'");
           
     
            if (v1 == "Y")
            {
                btnAdd.Visible = true;
                label9.Visible = true;
                btnSave.Visible = true;
                label17.Visible = true;
                btnupload.Visible = true;
                label28.Visible = true;
                btndelfile.Visible = true;
                label29.Visible = true;
                
            }
            else
            {
                btnAdd.Visible = false;
                label9.Visible = false;
                btnSave.Visible = false;
                label17.Visible = false;
                btnupload.Visible = false;
                label28.Visible = false;
                btndelfile.Visible = false;
                label29.Visible = false;
            }
            if (v2== "Y" || v1=="Y")
            {
               
                btnSave.Visible = true;
                label7.Visible = true;
                btnupload.Visible = true;
                label28.Visible = true;
                btndelfile.Visible = true;
                label29.Visible = true;


            }
            else
            {
               
                btnSave.Visible = false;
                label7.Visible = false;
                btnupload.Visible = false;
                label28.Visible = false;
                btndelfile.Visible = false;
                label29.Visible = false;
            }
            if (v3 =="Y")
            {
                btnDel.Visible = true;
                label5.Visible = true;
            }
            else
            {
                btnDel.Visible = false;
                label5.Visible = false;

            }
            if (v4 == "Y")
            {
                lkmange_audit.Visible = true;
            }
            else
            {
                lkmange_audit.Visible = false;

            }
            if (v5 == "Y")
            {
                lkfinancial_audit.Visible = true;
            }
            else
            {
                lkfinancial_audit.Visible = false;

            }
            if (v6 == "Y")
            {
                lkgeneral_manage.Visible = true;
            }
            else
            {
                lkgeneral_manage.Visible = false;

            }

            if (v5=="Y" || v6 == "Y")
            {
                linkLabel1.Visible = true;
            }
            else
            {
                linkLabel1.Visible = false;

            }
        
            hint.Location = new Point(400, 100);
            hint.ForeColor = Color.Red;
            if (bc.GET_IFExecutionSUCCESS_HINT_INFO(IFExecution_SUCCESS) != "")
            {

                hint.Text = bc.GET_IFExecutionSUCCESS_HINT_INFO(IFExecution_SUCCESS);
            }
            else
            {
                hint.Text = "";
            }
             textBox1.Text = IDO;
        
            DataTable dtx = basec.getdts(vou.getsql +" where A.VOID='" + textBox1.Text + "' ORDER BY  A.VOKEY ASC ");
                if (dtx.Rows.Count > 0)
                {
                   
                   
                    dt = vou.GET_TABLEINFO(dtx,1);
                    dateTimePicker1.Text = dtx.Rows[0]["凭证日期"].ToString();
                 
                    if (dtx.Rows[0]["是否打款"].ToString() == "已打款")
                    {
                    
                        linkLabel1.Text = "已打款";
                    }
                    else
                    {
                        linkLabel1.Text = "未打款";
                    }
                    if (dt.Rows.Count > 0 && dt.Rows.Count < 6)
                    {
                        int n = 6 - dt.Rows.Count;
                        for (int i = 0; i <n; i++)
                        {
                           
                            DataRow dr = dt.NewRow();
                            int b1 = Convert.ToInt32(dt.Rows[dt.Rows.Count - 1]["项次"].ToString());
                            dr["项次"] = Convert.ToString(b1 + 1);
                            //dr["币别"] = dt.Rows[dt.Rows.Count - 1]["币别"].ToString();
                            //dr["汇率"] = decimal.Parse(dt.Rows[dt.Rows.Count - 1]["汇率"].ToString());
                            dt.Rows.Add(dr);
                        }
                    }
                }
                else
                {
                    linkLabel1.Text = "未打款";
                    dt = total1();
                   
                }
         dataGridView1.DataSource = dt;
         bind2();
        }
        #endregion
        #region bind2
        private void bind2()
        {
           
            dt3 = bc.getdt(@"
SELECT cast(0   as   bit)   as   复选框,
old_file_name AS 文件名,FLKEY AS 索引,New_File_Name as 新文件名 FROM WAREFILE WHERE WAREID='" + textBox1.Text + "' AND INITIAL_OR_OTHER='INITIAL'");
            dataGridView2.DataSource = dt3;
            dgvStateControl();
          
            //dateTimePicker1.Text = DateTime.Now.ToString("yyyy/MM/dd").Replace("-", "/");


            this.WindowState = FormWindowState.Maximized;
            Color c = System.Drawing.ColorTranslator.FromHtml("#efdaec");
            t1.BackColor = c;
            t2.BackColor = c;

       
            if (vou.RETURN_MANAGE_AUDIT_STATUS (textBox1 .Text ) == "Y")
            {

                lkmange_audit.Text = "经理已审核";

            }
            else
            {

                lkmange_audit.Text = "经理未审核";
            }
            if (vou.RETURN_FINANCIAL_AUDIT_STATUS (textBox1 .Text )=="Y")
            {

                lkfinancial_audit.Text = "财务已审核";
            }
            else 
            {
                lkfinancial_audit.Text = "财务未审核";
              
            }
            if (vou.RETURN_GENERAL_AUDIT_STATUS (textBox1 .Text )=="Y")
            {
                lkgeneral_manage.Text = "总经理已审核";

            }
            else
            {

                lkgeneral_manage.Text = "总经理未审核";
            }
            IF_DOUBLE_CLICK = false;


            label4.Text = "(1.科目性质为零用金的 签核流程只走财务签核即结束 2.审核与撤审都单击同一个按扭即可)";
            label4.ForeColor = c2;

            label7.Text = "(3.点击文件名另存文件 选中复选框后单击删除按钮删除 )";
            label7.ForeColor = c2;

        }
        #endregion
        #region dgvStateControl
        private void dgvStateControl()
        {
            int i;
            dataGridView1.RowHeadersDefaultCellStyle.BackColor = Color.Lavender;
            dataGridView2.RowHeadersDefaultCellStyle.BackColor = Color.Lavender;
            int numCols1 = dataGridView1.Columns.Count;
            int numCols2 = dataGridView2.Columns.Count;
            //dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;/*自动调整DATAGRIDVIEW的列宽*/
              dataGridView1.Columns["项次"].Width =40;
              dataGridView1.Columns["摘要"].Width =200;
              dataGridView1.Columns["科目"].Width =200;
              //dataGridView1.Columns["币别"].Width =40;
              //dataGridView1.Columns["汇率"].Width =60;
              dataGridView1.Columns["单价"].Width =60;
              dataGridView1.Columns["数量"].Width =60;
              dataGridView1.Columns["支出金额"].Width =80;
              //dataGridView1.Columns["支出本币"].Width =80;
              dataGridView1.Columns["收入金额"].Width =80;
              //dataGridView1.Columns["收入本币"].Width =80;

              dataGridView2.Columns["复选框"].Width = 50;
              dataGridView2.Columns["文件名"].Width = 130;
              dataGridView2.Columns["索引"].Width = 130;
              dataGridView2.Columns["新文件名"].Visible = false;
              dataGridView2.Columns["索引"].Visible = false;
            for (i = 0; i < numCols1; i++)
            {

                dataGridView1.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                this.dataGridView1.Columns[i].SortMode = DataGridViewColumnSortMode.NotSortable;
                //this.dataGridView1.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
                dataGridView1.EnableHeadersVisualStyles = false;
                dataGridView1.Columns[i].HeaderCell.Style.BackColor = Color.Lavender;

            }
            for (i = 0; i < numCols2; i++)
            {

                dataGridView2.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                this.dataGridView2.Columns[i].SortMode = DataGridViewColumnSortMode.NotSortable;
                //this.dataGridView1.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
                dataGridView2.EnableHeadersVisualStyles = false;
                dataGridView2.Columns[i].HeaderCell.Style.BackColor = Color.Lavender;

            }
            for (i = 0; i < dataGridView1.Columns.Count; i++)
            {
                dataGridView1.Columns[i].SortMode = DataGridViewColumnSortMode.NotSortable;
                dataGridView1.Columns[i].DefaultCellStyle.BackColor = Color.OldLace;
                i = i + 1;
            }
            for (i = 0; i < dataGridView2.Columns.Count; i++)
            {
                dataGridView2.Columns[i].SortMode = DataGridViewColumnSortMode.NotSortable;
                dataGridView2.Columns[i].DefaultCellStyle.BackColor = Color.OldLace;
                i = i + 1;
            }
    
            dataGridView1.Columns["摘要"].DefaultCellStyle.BackColor = Color.Yellow;
            dataGridView1.Columns["科目"].DefaultCellStyle.BackColor = Color.Yellow;
            dataGridView1.Columns["支出金额"].DefaultCellStyle.BackColor = Color.Yellow;
            dataGridView1.Columns["收入金额"].DefaultCellStyle.BackColor = Color.Yellow;
            dataGridView1.Columns["单价"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
            dataGridView1.Columns["支出金额"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
            dataGridView1.Columns["收入金额"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;

            dataGridView1.Columns["项次"].ReadOnly = true;
            dataGridView1.Columns["摘要"].ReadOnly = false;
            dataGridView1.Columns["科目"].ReadOnly = false;
            //dataGridView1.Columns["币别"].ReadOnly = true;
            //dataGridView1.Columns["汇率"].ReadOnly = true;
            dataGridView1.Columns["单价"].ReadOnly = false;
            dataGridView1.Columns["数量"].ReadOnly = false;
            dataGridView1.Columns["支出金额"].ReadOnly = false;
            //dataGridView1.Columns["支出本币"].ReadOnly = true;
            dataGridView1.Columns["收入金额"].ReadOnly = false;
            //dataGridView1.Columns["收入本币"].ReadOnly = true;
            
            dataGridView2.Columns["文件名"].ReadOnly = true;
            dataGridView2.Columns["索引"].ReadOnly = true;
          

        }
        #endregion
     
        #region total1
        private DataTable total1()
        {
            DataTable dtt2 = vou.GetTableInfo();
            for (i = 1; i <= 6; i++)
            {
                DataRow dr = dtt2.NewRow();
                dr["项次"] = i;
                //dr["币别"] ="RMB";
                //dr["汇率"] = "1";
                //dr["支出金额"] = "0";
                dtt2.Rows.Add(dr);
            }
            return dtt2;
        }
        #endregion
        #region override enter
        protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
        {
            if (keyData == Keys.Enter &&(( !(ActiveControl is System.Windows.Forms.TextBox) ||
                !((System.Windows.Forms.TextBox)ActiveControl).AcceptsReturn) ))
            {
               
                if (dataGridView1.CurrentCell.ColumnIndex == 7 && 
                    dataGridView1["支出金额",dataGridView1.CurrentCell.RowIndex].Value .ToString ()!=null )
                {
                    
                    SendKeys.SendWait("{Tab}");
                    SendKeys.SendWait("{Tab}");
                }
                else if (dataGridView1.CurrentCell.ColumnIndex == 9 )
                {
                    SendKeys.SendWait("{Tab}");
                    SendKeys.SendWait("{Tab}");
                    SendKeys.SendWait("{Tab}");
                }
                else
                {

                    SendKeys.SendWait("{Tab}");
                }
                return true;
            }
            if (keyData == (Keys.Enter | Keys.Shift))
            {
                SendKeys.SendWait("+{Tab}");
             
                return true;
            }
            if (keyData == (Keys.F7))
            {

                double_info();
              
                return true;
            }
            return base.ProcessCmdKey(ref msg, keyData);
        }
        #endregion
      
        #region juage()
        private bool juage()
        {
            bool b = false;
            for (int k = 0; k <dt.Rows .Count ; k++)
            {
                if (juage(k))
                {
                    b = true;
                    break;
                }
            }
            return b;
        }
        #endregion

        
        #region juage()
        private bool juage(int k)
        {
            bool b = false;
           
                string v1 = dt.Rows[k]["摘要"].ToString();
                string v2 =bc.REMOVE_NAME(dt.Rows[k]["科目"].ToString());
                //string v3 = dt.Rows[k]["币别"].ToString();
                //string v4 = dt.Rows[k]["汇率"].ToString();
                string v5 = dt.Rows[k]["单价"].ToString();
                string v6 = dt.Rows[k]["数量"].ToString();
                string v7 = dt.Rows[k]["支出金额"].ToString();
                string v8 = dt.Rows[k]["收入金额"].ToString();
                if (v2=="" && v7=="" && v8=="")
                {
                
                }
                else  if (bc.CheckKeyInValueIfNoExistsOrEmpty("ACCOUNTANT_COURSE", "ACCODE", v2, "科目"))
                {
                  
                    b = true;
                }
                else if (v2 != "" && v7 == "" && v8 == "")
                {
                    b = true;
                    //MessageBox.Show("科目不为空时需输入相关金额！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    hint.Text = "科目不为空时需输入相关金额！";
                }
                else if (etc.CheckKeyInValueIfExistsDetailCourse("ACCOUNTANT_COURSE", "ACCODE", v2, "科目","存在明细科目，需使用明细科目记帐！")==1)
                {
                    b = true;
                }
                /*else if (bc.CheckKeyInValueIfNoExistsOrEmpty("CURRENCY_MST", "CYCODE", v3, "币别"))
                {
                    b = true;
                }*/
          
                else if (v7 != "" && v8 != "")
                {
                    b = true;
                    //MessageBox.Show("支出金额与收入金额同行只能输入一方！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    hint.Text = "支出金额与收入金额同行只能输入一方！";
                }

             
               
            return b;
        }
        #endregion
        #region dgvDataSourceChanged
        private void dataGridView1_DataSourceChanged(object sender, EventArgs e)
        {
           /* int i;
            for (i = 0; i < dataGridView1.Columns.Count; i++)
            {
                if (dataGridView1.Columns[i].ValueType.ToString() == "System.Decimal")
                {
                    
                    dataGridView1.Columns[i].DefaultCellStyle.Format = "#0.00";
                    dataGridView1.Columns[i].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                }
              
            }
            if (dataGridView1.Columns["汇率"].ValueType.ToString() == "System.Decimal")
            {
                dataGridView1.Columns["汇率"].DefaultCellStyle.Format = "#0.0000";
                dataGridView1.Columns["汇率"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
            }*/
        }
        #endregion
        private void dataGridView1_DataError(object sender, DataGridViewDataErrorEventArgs e)
        {
            //MessageBox.Show("只能输入数字！", "", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            try
            {
                hint.Text = "只能输入数字！";
            }
            catch (Exception)
            {


            }
        }

        private void btnPrint_Click(object sender, EventArgs e)
        {
            try
            {

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "提示", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);

            }
        }
        #region btnExcelPrint
        private void btnExcelPrint_Click(object sender, EventArgs e)
        {
           /* try
            {
                DataTable dtn = boperate.PrintOrder(" WHERE ORID='" + textBox1.Text + "'");
                if (dtn.Rows.Count > 0)
                {
                    string v1 = @"D:\PrintModelForOrder.xls";
                    if (File.Exists(v1))
                    {
                        boperate.ExcelPrint(dtn, "订单", v1);
                    }
                    else
                    {
                        MessageBox.Show("指定路径不存在打印模版！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);

                    }
                }
                else
                {
                    MessageBox.Show("无数据可打印！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);

                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "提示", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);

            }*/
        }
        #endregion
        private void ClearText()
        {
          
            dateTimePicker1.Text = "";
            t1.Text = "";
            t2.Text = "";

        }
        #region save
        private void btnSave_Click(object sender, EventArgs e)
        {

            save();
        }
        #endregion
        private void save()
        {
     

            try
            {
                btnSave.Focus();
                dgvfoucs();
                DataTable dtx = bc.GET_NOEMPTY_ROW_COURSE_DT(dt);
                if (juage2())
                {


                }
                else if (dtx.Rows.Count > 0)
                {
                    vou.VOUCHER_DATE = dateTimePicker1.Text;
                    vou.EMID = LOGIN.EMID;
         
                    vou.ACCOUNTING_PERIOD_EXPIRATION_DATE = DateTime.Now.ToString("yyyy/MM/dd");
                    vou.MANAGE_AUDIT_STATUS = "N";
                    vou.FINANCIAL_AUDIT_STATUS = "N";
                    vou.GENERAL_MANAGE_AUDIT_STATUS = "N";
                    if (linkLabel1 .Text  == "未打款")
                    {
                        vou.IF_PAYFOR = "N";
                    }
                    else
                    {
                        vou.IF_PAYFOR = "Y";
                    }
              
                    vou.save("VOUCHER_MST", "VOUCHER_DET", "VOID", textBox1.Text, dtx);
                    IFExecution_SUCCESS = true;
                    bind();
                    F1.Bind();
                    F1.search();
                }
                else
                {
                    hint.Text = "至少有一项科目才能保存！";

                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "提示", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);

            }


        }

        #region juage2()
        private bool juage2()
        {
            bool b = false;
            string v5 = dt.Compute("sum(支出金额)","").ToString();
            string v6 = dt.Compute("sum(收入金额)","").ToString();
            //string v7 = dt.Compute("sum(支出本币)","").ToString();
            //string v8 = dt.Compute("sum(收入本币)","").ToString();
            DataTable dtx = bc.GET_NOEMPTY_ROW_COURSE_DT(dt);
            string v9 = bc.getOnlyString("SELECT GENERAL_MANAGE FROM RIGHTLIST WHERE  USID='" + LOGIN.USID + "' AND NODE_NAME='录入凭证作业'");
        
            decimal d1 = 0, d2 = 0;
          
            if (!string.IsNullOrEmpty(v5))
            {
                d1 = decimal.Parse(v5);
            }
            if (!string.IsNullOrEmpty(v6))
            {
                d2 = decimal.Parse(v6);
            }
            /*if (!string.IsNullOrEmpty(v7))
            {
                d3 = decimal.Parse(v7);
            }
            /*if (!string.IsNullOrEmpty(v8))
            {
                d4= decimal.Parse(v8);
            }*/
            if (juage())
            {
                b = true;
              
            }
            /*else  if (juage_ABSTRACT_NOEMPTY() >= 0)
            {
                if (dt.Rows[juage_ABSTRACT_NOEMPTY ()]["摘要"].ToString() == "")
                {
                    b = true;
               
                    hint.Text = "项次" + dt.Rows[juage_ABSTRACT_NOEMPTY()]["项次"].ToString() + "摘要不能为空！";
                  
                }

            }*/
           if (ADD_OR_UPDATE =="UPDATE" &&  vou.CheckIfALLOW_SAVEOR_DELETE (textBox1 .Text,LOGIN .USID  ))
            {
               
                b = true;
                hint.Text = vou.ErrowInfo;
           
            }
            else if (ADD_OR_UPDATE == "UPDATE" && bc.getOnlyString ("SELECT EDIT FROM RIGHTLIST WHERE USID='"+LOGIN .USID +"' AND NODE_NAME='录入凭证作业'")!="Y")
            {

                b = true;
                hint.Text = "您没有修改作业的权限";

            }
            return b;
        }
        #endregion
        #region juage3()
        private bool juage3()
        {
            bool b = false;
            string v9 = bc.getOnlyString("SELECT GENERAL_MANAGE FROM RIGHTLIST WHERE  USID='" + LOGIN.USID + "' AND NODE_NAME='录入凭证作业'");
            if (v9 == "N")
            {
                if (linkLabel1.Text == "已打款")
                {
                    b = true;
                    hint.Text = "您只能操作已打款";
                }

            }
            return b;
        }
        #endregion
        #region juage_ABSTRACT_NOEMPTY()
        private int juage_ABSTRACT_NOEMPTY()
        {
           
            int n = 0;
            for (int k = dt.Rows.Count - 1; k >= 0; k--)
            {

                if (dt.Rows[k]["支出金额"].ToString() != "" && dt.Rows[k]["收入金额"].ToString() == ""
                    || dt.Rows[k]["支出金额"].ToString() == "" && dt.Rows[k]["收入金额"].ToString() != "")
                {
                    n = k;
                    break;

                }
            }
            return n;

        }
        #endregion
        private void btnExit_Click(object sender, EventArgs e)
        {
            this.Close();
        }
      
        private void btnDel_Click(object sender, EventArgs e)
        {
          
         
            try
            {
                if (vou.CheckIfALLOW_SAVEOR_DELETE(textBox1.Text,LOGIN .USID ))
                {
                    hint.Text = vou.ErrowInfo;
                }
                else if (MessageBox.Show("确定要删除该条凭证吗？", "提示", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) == DialogResult.OK)
                {
                    basec.getcoms("DELETE VOUCHER_MST WHERE VOID='" + textBox1.Text + "'");
                    basec.getcoms("DELETE VOUCHER_DET WHERE VOID='" + textBox1.Text + "'");
                    bind();
                    ClearText();
                    textBox1.Text = "";
                    F1.Bind();
                    F1.search();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "提示", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);

            }
        }
        #region dgvCellEndEdit
        private void dataGridView1_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
       

            try
            {

                int a = dataGridView1.CurrentCell.ColumnIndex;
                int b = dataGridView1.CurrentCell.RowIndex;
                int c = dataGridView1.Columns.Count - 1;
                int d = dataGridView1.Rows.Count - 1;
                if (a == 2)
                {
                    if (!string.IsNullOrEmpty(dt.Rows[b]["科目"].ToString()))
                    {
                        dt2 = bc.getdt(etc.getsql + " WHERE A.ACCODE='" + dt.Rows[b]["科目"].ToString() + "'");
                        if (dt2.Rows.Count > 0)
                        {
                            string v1 = bc.getOnlyString("SELECT COURSE_NATURE FROM ACCOUNTANT_COURSE WHERE ACCODE='" + dt.Rows[b]["科目"].ToString() + "'");
                            dt.Rows[b]["科目"] = dt.Rows[b]["科目"].ToString() +
                                " " + etc.GetLastCourseAnd_CurrentCourseName(dt.Rows[b]["科目"].ToString()) + " " + v1;

                            if (b != 0)
                            {
                                if (dt.Rows[b]["摘要"].ToString() == "" && dt.Rows[b - 1]["摘要"].ToString() != "")
                                {

                                    dt.Rows[b]["摘要"] = dt.Rows[b - 1]["摘要"].ToString();
                                }
                                if (dt.Rows[b]["支出金额"].ToString() == "" && dt.Rows[b]["收入金额"].ToString() == "" && dt.Rows[b - 1]["支出金额"].ToString() != "")
                                {

                                    dt.Rows[b]["支出金额"] = dt.Rows[b - 1]["支出金额"].ToString();
                                }
                                else if (dt.Rows[b]["支出金额"].ToString() == "" && dt.Rows[b]["收入金额"].ToString() == "" && dt.Rows[b - 1]["收入金额"].ToString() != "")
                                {

                                    dt.Rows[b]["收入金额"] = dt.Rows[b - 1]["收入金额"].ToString();
                                }
                            }
                        }
                    }

                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "提示", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);

            }
        }
        #endregion
        #region dgvDoubleClick
        private void dataGridView1_DoubleClick(object sender, EventArgs e)
        {
        
            try
            {
                int currentrowsindex = dataGridView1.CurrentCell.RowIndex;
                int currentcolumnindex = dataGridView1.CurrentCell.ColumnIndex;
                if (currentcolumnindex == 1)
                {
                    CSPSS.BASE_INFO.ABSTRACT frm = new CSPSS.BASE_INFO.ABSTRACT();
                    frm.a5();
                    frm.ShowDialog();
                    if (IF_DOUBLE_CLICK)
                    {
                        dataGridView1["摘要", currentrowsindex].Value = frm.ABCODE;
                        dataGridView1.CurrentCell = dataGridView1["科目", dataGridView1.CurrentCell.RowIndex];
                        IF_DOUBLE_CLICK = false;
                    }
                }
                if (currentcolumnindex == 2)
                {

                    CSPSS.BASE_INFO.ACCOUNTANT_COURSE frm = new CSPSS.BASE_INFO.ACCOUNTANT_COURSE();
                    frm.a5();
                    frm.ShowDialog();
                    if (IF_DOUBLE_CLICK)
                    {
                        dataGridView1["科目", currentrowsindex].Value = frm.ACCODE;
                        dataGridView1.CurrentCell = dataGridView1["单价", dataGridView1.CurrentCell.RowIndex];
                        IF_DOUBLE_CLICK = false;
                    }

                }
            
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "提示", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);
            }

        }
        #endregion
        private void double_info()
        {

            CSPSS.BASE_INFO.ACCOUNTANT_COURSE frm = new CSPSS.BASE_INFO.ACCOUNTANT_COURSE();
            frm.a5();
            frm.ShowDialog();
            DataGridViewRow dgvr = dataGridView1.CurrentRow;
            int j = dataGridView1.CurrentCell.ColumnIndex;
            if (dataGridView1.Columns[j].Name == "科目")
            {
                dgvr.Cells["科目"].Value = frm.ACCODE;
                //dataGridView1.CurrentCell = dataGridView1["币别", dataGridView1.CurrentCell.RowIndex];
            } 
        }

        #region dgvCellEnter
        private void dataGridView1_CellEnter(object sender, DataGridViewCellEventArgs e)
        {
          
            try
            {
                int a = dataGridView1.CurrentCell.ColumnIndex;
                int b = dataGridView1.CurrentCell.RowIndex;
                int c = dataGridView1.Columns.Count - 1;
                int d = dataGridView1.Rows.Count - 1;


                if (a == c && b == d)
                {
                    if (dt.Rows.Count >= 6)
                    {

                        DataRow dr = dt.NewRow();
                        int b1 = Convert.ToInt32(dt.Rows[dt.Rows.Count - 1]["项次"].ToString());
                        dr["项次"] = Convert.ToString(b1 + 1);
                        //dr["币别"] = dt.Rows[dt.Rows.Count - 1]["币别"].ToString();
                        //dr["汇率"] = decimal.Parse(dt.Rows[dt.Rows.Count - 1]["汇率"].ToString());
                        dt.Rows.Add(dr);
                    }

                }
                dgvfoucs();
           
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "提示", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);

            }

        }
        #endregion
        #region ask
        private void ask(int k)
        {
            int n = k;
            //decimal v1 = decimal.Parse(dt.Rows[k]["汇率"].ToString());
            decimal v2=0, v3=0;
            if (!string.IsNullOrEmpty(dt.Rows[k]["支出金额"].ToString()))
            {
                v2 = decimal.Parse(dt.Rows[k]["支出金额"].ToString());
            }
            if (!string.IsNullOrEmpty(dt.Rows[k]["收入金额"].ToString()))
            {
                v3 = decimal.Parse(dt.Rows[k]["收入金额"].ToString());
            }
         
      
            ask1();
        }
        #endregion
        #region ask1
        private void ask1()
        {
            t1.Text = "";
            t2.Text = "";
         
            string v5 = dt.Compute("sum(支出金额)", "").ToString();
            string v6 = dt.Compute("sum(收入金额)", "").ToString();
            //string v7 = dt.Compute("sum(支出本币)", "").ToString();
            //string v8 = dt.Compute("sum(收入本币)", "").ToString();
            if (!string.IsNullOrEmpty(v5))
            {
                t1.Text = string.Format("{0:F2}", Convert.ToDouble(v5));
            
            }
            /*if (!string.IsNullOrEmpty(v7))
            {
                
                t3.Text = string.Format("{0:F2}", Convert.ToDouble(v7));
            }*/
            if (!string.IsNullOrEmpty(v6))
            {
                t2.Text = string.Format("{0:F2}", Convert.ToDouble(v6));
             
            }
            /*if (!string.IsNullOrEmpty(v8))
            {
                t4.Text = string.Format("{0:F2}", Convert.ToDouble(v8));
            }*/
        }
        #endregion
        #region dgvCellValidating
        private void dataGridView1_CellValidating(object sender, DataGridViewCellValidatingEventArgs e)
        {
      
            try
            {
                if (e.ColumnIndex == 2 && bc.CheckKeyInValueIfNoExists("ACCOUNTANT_COURSE", "ACCODE",
                 bc.REMOVE_NAME(e.FormattedValue.ToString()), "科目"))
                {

                    e.Cancel = true;
                }
                else if (e.ColumnIndex == 2 && e.FormattedValue.ToString() != "" &&
                 etc.CheckKeyInValueIfExistsDetailCourse("ACCOUNTANT_COURSE", "ACCODE", bc.REMOVE_NAME(e.FormattedValue.ToString()),
                 "科目", "存在明细科目，需使用明细科目记帐！") == 1)
                {

                    e.Cancel = true;
                }
                /*else if (e.ColumnIndex == 3 && bc.CheckKeyInValueIfNoExistsOrEmpty("CURRENCY_MST", "CYCODE", e.FormattedValue.ToString(), "币别"))
                {

                    e.Cancel = true;
                }*/
                /*else if (e.ColumnIndex == 4 && bc.CheckKeyInValueIfNoDigitOrEmpty(e.FormattedValue.ToString(), "汇率"))
                {

                    e.Cancel = true;
                }*/
                else if (e.ColumnIndex == 5 && bc.yesno(e.FormattedValue.ToString()) == 0)
                {
                    e.Cancel = true;
                    //MessageBox.Show("单价只能输入数字！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    hint.Text = "单价只能输入数字！";


                }
                else if (e.ColumnIndex == 6 && bc.yesno(e.FormattedValue.ToString()) == 0)
                {
                    e.Cancel = true;

                    hint.Text = "数量只能输入数字！";


                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "提示", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);

            }

        }
        #endregion
        private void dgvfoucs()
        {
            
            for (i = 0; i < dt.Rows .Count ; i++)
            {
                ask(i);
            }
        }
        private void TSMI_Click(object sender, EventArgs e)
        {
            dgvclear(dataGridView1.CurrentCell.RowIndex);
            
        }
        private void dgvclear(int r)
        {
            
            dt.Rows[r]["摘要"] = "";
            dt.Rows[r]["科目"] = null;
            //dt.Rows[r]["币别"] = "";

            //dt.Rows[r]["汇率"] = DBNull.Value;
            dt.Rows[r]["单价"] = "";
            dt.Rows[r]["数量"] = "";
            dt.Rows[r]["支出金额"] = DBNull.Value;
            //dt.Rows[r]["支出本币"] = DBNull.Value;
            dt.Rows[r]["收入金额"] = DBNull.Value;
            //dt.Rows[r]["收入本币"] = DBNull.Value;
            btnSave.Focus();
        }
        private void btnSelect_Click(object sender, EventArgs e)
        {
          
            if (vou.CheckIfALLOW_SAVEOR_DELETE (textBox1 .Text,LOGIN .USID  ))
            {
                hint.Text = vou.ErrowInfo;
            }
            else
            {
                dgvclear(dataGridView1.CurrentCell.RowIndex);
            }
        }

        private void btnAllSelect_Click(object sender, EventArgs e)
        {
            if (vou.CheckIfALLOW_SAVEOR_DELETE (textBox1 .Text,LOGIN .USID  ))
            {
                hint.Text = vou.ErrowInfo;
            }
            else 
            {
                for (i = 0; i < dt.Rows.Count; i++)
                {
                    dgvclear(i);
                }
            }
        }

        private void dataGridView1_RowValidating(object sender, DataGridViewCellCancelEventArgs e)
        {

            try
            {
                int r = dataGridView1.CurrentCell.RowIndex;
                if (dataGridView1["支出金额", r].Value.ToString() != "" && dataGridView1["收入金额", r].Value.ToString() != "")
                {
                    e.Cancel = true;
                    hint.Text = "支出金额与收入金额同行只能输入一方！";

                }
            }
            catch (Exception)
            {

            }
        }

        private void 提取科目F7ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            double_info();
            
        }
        #region lkmange_audit
        private void lkmange_audit_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
         
            try
            {

                if (vou.RETURN_GENERAL_AUDIT_STATUS(textBox1.Text) == "Y")
                {
                    if (vou.RETURN_MANAGE_AUDIT_STATUS(textBox1.Text) == "N")
                    {

                        basec.getcoms("UPDATE VOUCHER_MST SET MANAGE_AUDIT_STATUS='Y',MANAGE_AUDIT_DATE='" + DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss") + "' WHERE VOID='" + textBox1.Text + "'");
                        bind();
                        F1.Bind();
                        F1.search();
                    }
                    else
                    {
                        //hint.Text = "状态为开立或经理已审核才能操作审核与撤审核";
                        hint.Text = "状态为总经理已审核不能操作撤审核";
                    }

                }
                else if (vou.RETURN_FINANCIAL_AUDIT_STATUS(textBox1.Text) == "Y")
                {
                    if (vou.RETURN_MANAGE_AUDIT_STATUS(textBox1.Text) == "N")
                    {

                        basec.getcoms("UPDATE VOUCHER_MST SET MANAGE_AUDIT_STATUS='Y',MANAGE_AUDIT_DATE='" + DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss") + "' WHERE VOID='" + textBox1.Text + "'");
                        bind();
                        F1.Bind();
                        F1.search();
                    }
                    else
                    {
                        //hint.Text = "状态为开立或经理已审核才能操作审核与撤审核";
                        hint.Text = "状态为财务已审核不能操作撤审核";
                    }

                }
                else if (vou.RETURN_MANAGE_AUDIT_STATUS(textBox1.Text) == "N")
                {
                    basec.getcoms("UPDATE VOUCHER_MST SET MANAGE_AUDIT_STATUS='Y',MANAGE_AUDIT_DATE='" + DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss") + "' WHERE VOID='" + textBox1.Text + "'");
                    bind();
                    F1.Bind();
                    F1.search();

                }
                else
                {

                    basec.getcoms("UPDATE VOUCHER_MST SET MANAGE_AUDIT_STATUS='N',MANAGE_AUDIT_DATE='" + DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss") + "' WHERE VOID='" + textBox1.Text + "'");
                    bind();
                    F1.Bind();
                    F1.search();

                }
         

            }
            catch (Exception)
            {

            }
        }
        #endregion
        #region lkfinancial_audit
        private void lkfinancial_audit_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            try
            {
                string s2 = bc.getOnlyString("SELECT FINANCIAL_AUDIT_STATUS FROM VOUCHER_MST WHERE VOID='" + textBox1.Text + "'");
                if (vou.RETURN_GENERAL_AUDIT_STATUS(textBox1.Text) == "Y")
                {
                    if (vou.RETURN_FINANCIAL_AUDIT_STATUS(textBox1.Text) == "N")
                    {
                        basec.getcoms("UPDATE VOUCHER_MST SET FINANCIAL_AUDIT_STATUS='Y',FINANCIAL_AUDIT_DATE='" + DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss") + "' WHERE VOID='" + textBox1.Text + "'");
                        bind();
                        F1.Bind();
                        F1.search();
                    }
                    else
                    {

                        hint.Text = "状态为总经理已审核不能操作撤审核";
                    }
                }
                else if (vou.RETURN_FINANCIAL_AUDIT_STATUS(textBox1.Text) == "N")
                {

                    basec.getcoms("UPDATE VOUCHER_MST SET FINANCIAL_AUDIT_STATUS='Y',FINANCIAL_AUDIT_DATE='" + DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss") + "' WHERE VOID='" + textBox1.Text + "'");
                    bind();
                    F1.Bind();
                    F1.search();
                }
                else
                {
                    basec.getcoms("UPDATE VOUCHER_MST SET FINANCIAL_AUDIT_STATUS='N',FINANCIAL_AUDIT_DATE='" + DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss") + "' WHERE VOID='" + textBox1.Text + "'");
                    bind();
                    F1.Bind();
                    F1.search();

                }
            }
            catch (Exception)
            {


            }
        }
        #endregion
        #region lkgeneral_manage
        private void lkgeneral_manage_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            try
            {
                if (vou.RETURN_GENERAL_AUDIT_STATUS(textBox1.Text) == "N")
                {
                    basec.getcoms("UPDATE VOUCHER_MST SET GENERAL_MANAGE_AUDIT_STATUS='Y',GENERAL_MANAGE_AUDIT_DATE='" + DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss") + "' WHERE VOID='" + textBox1.Text + "'");
                    bind();
                    F1.Bind();
                    F1.search();
                }
                else
                {

                    basec.getcoms("UPDATE VOUCHER_MST SET GENERAL_MANAGE_AUDIT_STATUS='N',GENERAL_MANAGE_AUDIT_DATE='" + DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss") + "' WHERE VOID='" + textBox1.Text + "'");
                    bind();
                    F1.Bind();
                    F1.search();

                }
            }
            catch (Exception)
            {


            }
          
        }
        #endregion 
     
      
    
    
        #region delfile
        public void delfile()
        {

            try
            {
                string v21 = bc.getOnlyString("SELECT EDIT FROM RIGHTLIST WHERE USID='" + LOGIN.USID + "' AND NODE_NAME='录入凭证作业'");
                if (v21 != "Y" && ADD_OR_UPDATE == "UPDATE")
                {
                    hint.Text = "您没有修改权限不能删除文件";
                }
                else
                {
                    if (MessageBox.Show("确定要删除该文件吗？", "提示", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) == DialogResult.OK)
                    {
                        if (dt3.Rows.Count > 0)
                        {

                            for (int i = 0; i < dt3.Rows.Count; i++)
                            {
                                if (dataGridView2.Rows[i].Cells[0].EditedFormattedValue.ToString() == "True")
                                {

                                    string v2 = dt3.Rows[i]["索引"].ToString();
                                    string v3 = bc.getOnlyString("SELECT PATH FROM WAREFILE WHERE FLKEY='" + v2 + "'");
                                    string v4 = bc.FROM_RIGHT_UNTIL_CHAR(v3, 47);
                                    bc.getcom(@"INSERT INTO SERVER_DELETE_FILE(FLKEY,NEW_FILE_NAME) VALUES ('" + v2 + "','" + v4 + "')");

                                    bc.getcom("DELETE WAREFILE WHERE FLKEY='" + v2 + "'");
                                }


                            }
                            bind2();

                        }

                    }

                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "提示", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);

            }


        }
        #endregion
        private void btnAdd_Click(object sender, EventArgs e)
        {
            ClearText();
            IFExecution_SUCCESS = false;
            IDO = vou.GETID();
          
            bind();
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {

            try
            {
                if (juage3())
                {
                }
                else
                {
                    if (linkLabel1.Text == "未打款")
                    {
                        basec.getcoms("UPDATE VOUCHER_MST SET IF_PAYFOR='Y' WHERE VOID='" + textBox1.Text + "'");
                    }
                    else
                    {
                        basec.getcoms("UPDATE VOUCHER_MST SET IF_PAYFOR='N' WHERE VOID='" + textBox1.Text + "'");
                    }
                    bind();
                    F1.Bind();
                    F1.search();

                }
            }
            catch (Exception)
            {


            }
        }

        private void btnupload_Click(object sender, EventArgs e)
        {
            DataTable dty = bc.getdt("SELECT * FROM WAREFILE WHERE WAREID='" + IDO  + "'");
            if (juage())
            {

            }
            /*else if (dty.Rows.Count.ToString() == "2")
            {

                hint.Text = "最多只能上传一张图片";
            }*/
            else
            {

                uploadfile();
            }
            try
            {


            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "提示", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);

            }
        }
        #region uploadfile
        private void uploadfile()
        {
            CFileInfo cfileinfo = new CFileInfo();
            int i = 0;
            label53.Visible = false;
            label55.Visible = false;
            label56.Visible = false;
            label57.Visible = false;
            progressBar1.Visible = false;
            /*  string v2 = bc.getOnlyString("SELECT EDIT FROM RIGHTLIST WHERE USID='" + LOGIN.USID + "' AND NODE_NAME='传单作业'");
              if (v2 != "Y" && ADD_OR_UPDATE == "UPDATE")
              {
                  hint.Text = "您没有修改权限不能修改上传";
              }
              else*/
            label52.Text = "";
            if (bc.RETURN_SERVER_IP_OR_DOMAIN() == "")
            {
                hint.Text = "未设置服务器IP或域名";
               
            }

            else
            {
                OpenFileDialog openf = new OpenFileDialog();
                if (openf.ShowDialog() == DialogResult.OK)
                {

                    Random ro = new Random();
                    string stro = ro.Next(80, 10000000).ToString() + "-";
                    string NeWAREID = DateTime.Now.Minute.ToString() + DateTime.Now.Second.ToString() + DateTime.Now.Millisecond.ToString() + stro;

                    cfileinfo.SERVER_IP_OR_DOMAIN = bc.RETURN_SERVER_IP_OR_DOMAIN();
                    WATER_MARK_CONTENT = "";//水印内容
                    //cfileinfo.UploadImage(openf.FileName, Path.GetFileName(openf.FileName), textBox1 .Text );
                    //this.UploadFile(openf.FileName, System.IO.Path.GetFileName(openf.FileName), "File/", textBox1.Text);

                    string v21 = bc.FROM_RIGHT_UNTIL_CHAR(Path.GetFileName(openf.FileName), 46);
                    old_file_name = Path.GetFileName(openf.FileName);
                    NEW_FILE_NAME = NeWAREID + Path.GetFileName(openf.FileName);
                    //如果上传的是图片文件
                    if (v21 == "jpeg" || v21 == "jpg" || v21 == "JPG" || v21 == "png" || v21 == "bmp" || v21 == "gif")
                    {


                        //裁切小图
                        cfileinfo.MakeThumbnail(openf.FileName, "d:\\80X80" + Path.GetFileName(openf.FileName), 80, 80, "Cut");
                        //裁切700*700
                        cfileinfo.MakeThumbnail(openf.FileName, "d:\\700X700" + Path.GetFileName(openf.FileName), 700, 700, "Cut");

                        //小图加水印
                        cfileinfo.ADD_WATER_MARK("d:\\80X80" + Path.GetFileName(openf.FileName), "d:\\80X80" + NeWAREID + Path.GetFileName(openf.FileName), WATER_MARK_CONTENT);
                        //700*700图加水印
                        cfileinfo.ADD_WATER_MARK("d:\\700X700" + Path.GetFileName(openf.FileName), "d:\\700X700" + NeWAREID + Path.GetFileName(openf.FileName), WATER_MARK_CONTENT);
                        //原图加水印
                        cfileinfo.ADD_WATER_MARK(openf.FileName, "d:\\INITIAL" + NeWAREID + Path.GetFileName(openf.FileName), WATER_MARK_CONTENT);
                        INITIAL_OR_OTHER = "INITIAL";

                        //上传原图
                        i = Upload_Request("http://" + bc.RETURN_SERVER_IP_OR_DOMAIN() + "/webuploadfile/default.aspx", "D:\\INITIAL" + NeWAREID + System.IO.Path.GetFileName(openf.FileName),
                                "INITIAL" + NeWAREID + System.IO.Path.GetFileName(openf.FileName), progressBar1, textBox1.Text);

                        //上传80X80的缩略图
                        INITIAL_OR_OTHER = "80X80";
                        i = Upload_Request("http://" + bc.RETURN_SERVER_IP_OR_DOMAIN() + "/webuploadfile/default.aspx", "D:\\80X80" + NeWAREID + System.IO.Path.GetFileName(openf.FileName),
                                "80X80" + NeWAREID + System.IO.Path.GetFileName(openf.FileName), progressBar1, textBox1.Text);

                        //上传700X700的缩略图
                        INITIAL_OR_OTHER = "700X700";
                        i = Upload_Request("http://" + bc.RETURN_SERVER_IP_OR_DOMAIN() + "/webuploadfile/default.aspx", "D:\\700X700" + NeWAREID + System.IO.Path.GetFileName(openf.FileName),
                                "700X700" + NeWAREID + System.IO.Path.GetFileName(openf.FileName), progressBar1, textBox1.Text);

                        //删除本地临时水印图及剪切图
                        if (File.Exists("d:\\80X80" + NeWAREID + Path.GetFileName(openf.FileName)))
                        {
                            File.Delete("d:\\80X80" + NeWAREID + Path.GetFileName(openf.FileName));
                            File.Delete("d:\\700X700" + NeWAREID + Path.GetFileName(openf.FileName));
                            File.Delete("d:\\80X80" + Path.GetFileName(openf.FileName));
                            File.Delete("d:\\700X700" + Path.GetFileName(openf.FileName));
                            File.Delete("d:\\" + Path.GetFileName(openf.FileName));
                            File.Delete("d:\\INITIAL" + NeWAREID + Path.GetFileName(openf.FileName));
                        }
                        if (i == 1)
                        {
                            label52.Text = "成功上传";
                        }
                        else
                        {
                            label52.Text = "上传失败";
                        }

                        bind2();
                    }
                    else
                    {
                        //MessageBox.Show("只能上传图片格式为jpeg/jpg/png/bmp/gif", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        /*label53.Visible = true;
                        label55.Visible = true;
                        label56.Visible = true;
                        label57.Visible = true;
                        progressBar1.Visible = true;*/
                        //上传的是非图片文件
                        INITIAL_OR_OTHER = "INITIAL";
                        i = Upload_Request("http://" + bc.RETURN_SERVER_IP_OR_DOMAIN() + "/webuploadfile/default.aspx", openf.FileName,
                                                      "INITIAL" + NeWAREID + System.IO.Path.GetFileName(openf.FileName), progressBar1, textBox1.Text);
                        bind2();
                    }

                }
            }

        }
        #endregion
        #region Upload_Request
        public int Upload_Request(string address, string fileNamePath, string saveName, ProgressBar progressBar, string WAREID)
        {
            int returnValue = 0;
            // 要上传的文件

            FileStream fs = new FileStream(fileNamePath, FileMode.Open, FileAccess.Read);
            BinaryReader r = new BinaryReader(fs);
            //时间戳
            string strBoundary = "----------" + DateTime.Now.Ticks.ToString("x");
            byte[] boundaryBytes = Encoding.ASCII.GetBytes("\r\n--" + strBoundary + "\r\n");
            //请求头部信息
            StringBuilder sb = new StringBuilder();
            sb.Append("--");
            sb.Append(strBoundary);
            sb.Append("\r\n");
            sb.Append("Content-Disposition: form-data; name=\"");
            sb.Append("file");
            sb.Append("\"; filename=\"");
            sb.Append(saveName);
            sb.Append("\"");
            sb.Append("\r\n");
            sb.Append("Content-Type: ");
            sb.Append("application/octet-stream");
            sb.Append("\r\n");
            sb.Append("\r\n");
            string strPostHeader = sb.ToString();


            byte[] postHeaderBytes = Encoding.UTF8.GetBytes(strPostHeader);
            // 根据uri创建HttpWebRequest对象
            HttpWebRequest httpReq = (HttpWebRequest)WebRequest.Create(new Uri(address));
            httpReq.Method = "POST";
            //对发送的数据不使用缓存
            httpReq.AllowWriteStreamBuffering = false;
            //设置获得响应的超时时间（300秒）
            httpReq.Timeout = 300000;
            httpReq.ContentType = "multipart/form-data; boundary=" + strBoundary;
            long length = fs.Length + postHeaderBytes.Length + boundaryBytes.Length;
            long fileLength = fs.Length;
            httpReq.ContentLength = length;
            if (fileLength / 1048576.0 > 2.5)
            {

                label52.Visible = false;
                label53.Visible = false;
                label55.Visible = false;
                label56.Visible = false;
                label57.Visible = false;
                progressBar1.Visible = false;
                MessageBox.Show("上传的图片长度为:" + (fileLength / 1048576.0).ToString("F2") + "M" + " 已经大于允许上传的2.5M");
            }
            else
            {
                try
                {
                    progressBar.Maximum = int.MaxValue;
                    progressBar.Minimum = 0;
                    progressBar.Value = 0;
                    //每次上传4k
                    int bufferLength = 4096;
                    byte[] buffer = new byte[bufferLength];
                    //已上传的字节数
                    long offset = 0;
                    //开始上传时间
                    DateTime startTime = DateTime.Now;
                    int size = r.Read(buffer, 0, bufferLength);

                    Stream postStream = httpReq.GetRequestStream();
                    //发送请求头部消息
                    postStream.Write(postHeaderBytes, 0, postHeaderBytes.Length);
                    while (size > 0)
                    {
                        postStream.Write(buffer, 0, size);
                        offset += size;
                        progressBar.Value = (int)(offset * (int.MaxValue / length));
                        TimeSpan span = DateTime.Now - startTime;
                        double second = span.TotalSeconds;
                        label53.Text = "已用时：" + second.ToString("F2") + "秒";

                        if (second > 0.001)
                        {
                            label55.Text = "平均速度：" + (offset / 1024 / second).ToString("0.00") + "KB/秒";
                        }
                        else
                        {
                            label55.Text = "正在连接…";
                        }
                        label56.Text = "已上传：" + (offset * 100.0 / length).ToString("F2") + "%";
                        label57.Text = (offset / 1048576.0).ToString("F2") + "M/" + (fileLength / 1048576.0).ToString("F2") + "M";
                        Application.DoEvents();
                        size = r.Read(buffer, 0, bufferLength);
                    }
                    //添加尾部的时间戳
                    postStream.Write(boundaryBytes, 0, boundaryBytes.Length);
                    postStream.Close();

                    string year = DateTime.Now.ToString("yy");
                    string month = DateTime.Now.ToString("MM");
                    string day = DateTime.Now.ToString("dd");
                    string varDate = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss");
                    string v1 = bc.numYMD(20, 12, "000000000001", "SELECT * FROM WAREFILE", "FLKEY", "FL");
                    string newFileName, uriString;
                    newFileName = System.IO.Path.GetFileName(saveName);
                    uriString = "http://" + bc.RETURN_SERVER_IP_OR_DOMAIN() + "/uploadfile/" + newFileName;


                    String sql = @"
INSERT INTO  WAREFILE 
(
FLKEY,
WAREID,
old_file_name,
NEW_FILE_NAME,
PATH,
INITIAL_OR_OTHER,
DATE,
YEAR,
MONTH,
DAY
) 
VALUES
(
@FLKEY,
@WAREID,
@old_file_name,
@NEW_FILE_NAME,
@PATH,
@INITIAL_OR_OTHER,
@DATE,
@YEAR,
@MONTH,
@DAY

)";
                    SqlConnection sqlcon = bc.getcon();
                    SqlCommand sqlcom = new SqlCommand(sql, sqlcon);
                    sqlcom.Parameters.Add("@FLKEY", SqlDbType.VarChar, 20).Value = v1;
                    sqlcom.Parameters.Add("@WAREID", SqlDbType.VarChar, 20).Value = IDO;
                    sqlcom.Parameters.Add("@old_file_name", SqlDbType.VarChar, 100).Value = old_file_name;
                    sqlcom.Parameters.Add("@NEW_FILE_NAME", SqlDbType.VarChar, 100).Value = NEW_FILE_NAME;
                    sqlcom.Parameters.Add("@PATH", SqlDbType.VarChar, 100).Value = uriString;
                    sqlcom.Parameters.Add("@INITIAL_OR_OTHER", SqlDbType.VarChar, 100).Value = INITIAL_OR_OTHER;
                    sqlcom.Parameters.Add("@DATE", SqlDbType.VarChar, 20).Value = varDate;
                    sqlcom.Parameters.Add("@YEAR", SqlDbType.VarChar, 20).Value = year;
                    sqlcom.Parameters.Add("@MONTH", SqlDbType.VarChar, 20).Value = month;
                    sqlcom.Parameters.Add("@DAY", SqlDbType.VarChar, 20).Value = day;
                    sqlcon.Open();
                    sqlcom.ExecuteNonQuery();
                    sqlcon.Close();


                    //获取服务器端的响应
                    WebResponse webRespon = httpReq.GetResponse();
                    Stream s = webRespon.GetResponseStream();
                    StreamReader sr = new StreamReader(s);
                    //读取服务器端返回的消息
                    String sReturnString = sr.ReadLine();
                    s.Close();
                    sr.Close();
                    if (sReturnString == "Success")
                    {
                        returnValue = 1;
                    }
                    else if (sReturnString == "Error")
                    {
                        returnValue = 0;
                    }
                }
                catch
                {
                    returnValue = 0;
                }
                finally
                {
                    fs.Close();
                    r.Close();
                }
            }
            return returnValue;
        }
        #endregion
        private void btndelfile_Click(object sender, EventArgs e)
        {
            try
            {
                /*string v21 = bc.getOnlyString("SELECT EDIT FROM RIGHTLIST WHERE USID='" + LOGIN.USID + "' AND NODE_NAME='传单作业'");
                if (v21 != "Y" && ADD_OR_UPDATE == "UPDATE")
                {
                    hint.Text = "您没有修改权限不能删除文件";
                }
                else if (vou.CheckIfALLOW_SAVEOR_DELETE(textBox1.Text, LOGIN.USID))
                {
                    hint.Text = vou.ErrowInfo;
                }
                else
                {
                

                }*/
                if (MessageBox.Show("确定要删除该文件吗？", "提示", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) == DialogResult.OK)
                {
                    if (dt3.Rows.Count > 0)
                    {

                        for (int i = 0; i < dt3.Rows.Count; i++)
                        {
                            if (dataGridView2.Rows[i].Cells[0].EditedFormattedValue.ToString() == "True")
                            {

                                string v2 = dt3.Rows[i]["索引"].ToString();
                                string v4 = dt3.Rows[i]["新文件名"].ToString();
                                bc.getcom(@"INSERT INTO SERVER_DELETE_FILE(FLKEY,NEW_FILE_NAME) VALUES ('" + v2 + "','" + v4 + "')");
                                bc.getcom("DELETE WAREFILE WHERE NEW_FILE_NAME='" + v4 + "'");

                            }
                        }
                        bind2();

                    }

                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "提示", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);

            }
        }

        private void dataGridView2_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                int i = dataGridView2.CurrentCell.RowIndex;

                if (dataGridView2.CurrentCell.ColumnIndex == 1)
                {
                    SaveFileDialog sfl = new SaveFileDialog();
                    sfl.FileName = dt3.Rows[dataGridView2.CurrentCell.RowIndex]["文件名"].ToString();
                    sqb = new StringBuilder();
                    sqb.AppendFormat("SELECT PATH FROM WAREFILE WHERE ");
                    sqb.AppendFormat(" FLKEY='{0}'", dataGridView2["索引",i].Value .ToString());
                    sqb.AppendFormat(" AND INITIAL_OR_OTHER='INITIAL'");
                    WebClient wclient = new WebClient();
                    string v1 = bc.getOnlyString(sqb.ToString());
                    wclient.DownloadFile(v1, AppDomain.CurrentDomain.BaseDirectory+ "temp\\"+dt3.Rows[dataGridView2.CurrentCell.RowIndex]["文件名"].ToString());
                    string v2 = AppDomain.CurrentDomain.BaseDirectory  +"temp\\"+ dt3.Rows[dataGridView2.CurrentCell.RowIndex]["文件名"].ToString();
                    /*DataTable dt3x = bc.getdt("SELECT * FROM WAREFILE WHERE FLKEY='" + dt3.Rows[dataGridView1.CurrentCell.RowIndex]["索引"].ToString() + "'");
                    Byte[] byte2 = (byte[])dt3x.Rows[0]["IMAGE_DATA"];
                    System.IO.File.WriteAllBytes(sfl.FileName, byte2);*/
                    if (File.Exists(v2))
                    {
                        System.Diagnostics.Process.Start(v2);

                    }
                    //hint.Text = "已下载";

                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "提示", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);

            }
        }

    }
}

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


namespace CSPSS.BASE_INFO
{
    public partial class ACCOUNTANT_COURSE : Form
    {
        DataTable dt = new DataTable();
        DataTable dt2 = new DataTable();
        DataTable dt3 = new DataTable();
        protected int M_int_judge, t;
        basec bc = new basec();
        ExcelToCSHARP etc = new ExcelToCSHARP();
        Color c = System.Drawing.ColorTranslator.FromHtml("#efdaec");
        CACCOUNTANT_COURSE caccountant_course = new CACCOUNTANT_COURSE();
        PERIOD pe = new PERIOD();
        private string _IDO;
        protected int select;
        public string IDO
        {
            set { _IDO = value; }
            get { return _IDO; }

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
        private string _ACCODE;
        public string ACCODE
        {
            set { _ACCODE = value; }
            get { return _ACCODE; }
        }
        Color c2 = System.Drawing.ColorTranslator.FromHtml("#990033");
        public ACCOUNTANT_COURSE()
        {
            InitializeComponent();
          
           
        }
        private void ACCOUNTANT_COURSE_Load(object sender, EventArgs e)
        {
            
            textBox3.BackColor = Color.Yellow;
            bind();
            label2.Text = "(1.为提高凭证录入的效率，建议设置科目代码 2.科目性质为(1)正常(2)零用金两种 区别在于后续做凭证时签核流程不一样  押金走零用金流程)";
            label2.ForeColor = c2;
            comboBox1.Text = "正常";
   
         
        }
        private void currency()
        {
           
            DataTable dtx = bc.getdt("SELECT * FROM CURRENCY_MST WHERE CYCODE='RMB'");
         

        }

        #region override enter
        protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
        {
            if (keyData == Keys.Enter && ((!(ActiveControl is System.Windows.Forms.TextBox) ||
                !((System.Windows.Forms.TextBox)ActiveControl).AcceptsReturn)))
            {

             
                    SendKeys.SendWait("{Tab}");
                
                return true;
            }
            if (keyData == (Keys.Enter | Keys.Shift))
            {
                SendKeys.SendWait("+{Tab}");

                return true;
            }
        
            return base.ProcessCmdKey(ref msg, keyData);
        }
        #endregion
        #region bind
        private void bind()
        {
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
            if (ADD_OR_UPDATE == "UPDATE")
            {
               
            }
            else
            {
                textBox1.Text = caccountant_course.GETID();
            }
            treeView1.Nodes.Clear();
            dt = bc.getdt("SELECT * FROM ACCOUNTANT_COURSE");
            SHOW_TREEVIEW(dt);
            textBox3.Focus();

            //this.WindowState = FormWindowState.Maximized;
            think();
            

        }
        #endregion
        #region think
        private void think()
        {

            dt2 = bc.getdt("SELECT * FROM ACCOUNTANT_COURSE ");
            AutoCompleteStringCollection inputInfoSource = new AutoCompleteStringCollection();
            AutoCompleteStringCollection inputInfoSource4 = new AutoCompleteStringCollection();
            comboBox2.Items.Clear();
                foreach (DataRow dr in dt2.Rows)
                {

                    comboBox2.Items.Add(dr["ACCODE"].ToString() + " " + dr["ACNAME"].ToString());
                    inputInfoSource.Add(dr["ACCODE"].ToString() + " " + dr["ACNAME"].ToString());


                }
            this.comboBox2.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend;
            this.comboBox2.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.CustomSource;
            this.comboBox2.AutoCompleteCustomSource = inputInfoSource;


        }
        #endregion
        #region show_treeview
        private void SHOW_TREEVIEW(DataTable dt)
        {

           
                dt = bc.GET_DT_TO_DV_TO_DT(dt, "", "PARENT_NODEID='NULL'");
           
                if (dt.Rows.Count > 0)
                {
                    foreach (DataRow dr in dt.Rows)
                    {
                        TreeNode trd = treeView1.Nodes.Add(dr["ACCODE"].ToString()+" "+dr["ACNAME"].ToString());
                   
                        if (trd.Text ==textBox3.Text)
                        {

                            trd.BackColor = c;

                        }
                   

                    }

                }
              
            
        }
        #endregion

        #region show_treeview_O
        private void SHOW_TREEVIEW_O(string ACID,TreeNode trd)
        {

                    dt2 = bc.getdt("SELECT * FROM ACCOUNTANT_COURSE WHERE PARENT_NODEID='" + ACID + "'");
                    if (dt2.Rows.Count > 0)
                    {
                        foreach (DataRow dr1 in dt2.Rows)
                        {

                            TreeNode TRC = new TreeNode();
                            TRC.Text =dr1["ACCODE"].ToString()+" "+dr1["ACNAME"].ToString ();
                            trd.Nodes.Add(TRC);
                            if (TRC.Text == textBox2.Text+" "+textBox3 .Text )
                            {

                                TRC.BackColor = c;
                              
                            }
                            SHOW_TREEVIEW_O(dr1["ACID"].ToString(),TRC);
                          
                        }
                   }
        }
        #endregion
        #region bind1
        private void bind(DataTable dt)
        {

            try
            {
                if (dt.Rows.Count > 0)
                {
                    textBox1.Text = dt.Rows[0]["ACID"].ToString();
                    textBox2.Text = dt.Rows[0]["ACCODE"].ToString();
                    textBox3.Text = dt.Rows[0]["ACNAME"].ToString();
                    bind2(dt.Rows[0]["ACCODE"].ToString());
                  
                }
                think();
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message, "提示", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);
            }
        }
        #endregion
        #region bind2
        private void bind2(string ACCODE)
        {
       
        }
        #endregion
       
        #region save
        protected void save()
        {
            etc.EMID = LOGIN.EMID;
            etc.save(textBox1.Text, textBox2.Text, textBox3.Text, "","","",comboBox1 .Text );
            
            ADD_OR_UPDATE = etc.ADD_OR_UPDATE;
          
           
        }
        private void COURSE_TYPE_LOAD()
        {
            if (textBox2.Text.Length > 0)
            {
                int k = Convert.ToInt32(textBox2.Text.Substring(0, 1));
                dt = etc.GetCOURSE_TypeData(k);
            }

            if (dt.Rows.Count > 0)
            {

                //bind(dt);
            }
            else
            {
                textBox1.Text = "";
                ClearText();
               
            }
            think();
            textBox2.Focus();
       
          
            treeView1.Nodes.Clear();// no allow once again onload
            
            SHOW_TREEVIEW(dt);
            if (textBox2.Text.Length >= 4)
            {
                foreach (TreeNode trd in treeView1.Nodes)
                {
                    if (trd.Text.Substring(0, 4) == textBox2.Text.Substring(0, 4))
                    {

                        trd.ExpandAll();
                    

                    }
                    //MessageBox.Show(trd.Text);
                }
            }
            LoadAgain();
            //
        }
        #endregion
        


        private void dataGridView1_DataSourceChanged(object sender, EventArgs e)
        {
        
        }

        private void dataGridView1_DataError(object sender, DataGridViewDataErrorEventArgs e)
        {
            MessageBox.Show("只能输入数字！", "", MessageBoxButtons.OK, MessageBoxIcon.Warning);

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
        #region excelprint
        private void btnExcelPrint_Click(object sender, EventArgs e)
        {

        }
        #endregion
        #region btnadd

        #endregion
        #region loadagain
        private void LoadAgain()
        {
            ClearText();
            string a1 = bc.numYM(10, 4, "0001", "select * from Accountant_Course", "ACID", "AC");
            if (a1 == "Exceed Limited")
            {
                MessageBox.Show("编码超出限制！", "", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            else
            {
                textBox1.Text = a1;
            }
            //dataGridView1.DataSource = total1();
        }
        #endregion
        private void ClearText()
        {
            textBox2.Text = "";
            textBox3.Text = "";
       
          
        }
        private void btnSave_Click(object sender, EventArgs e)
        {
            save1();

        }
        private void save1()
        {


            try
            {
                if (textBox3.Text == "")
                {
                    
                    hint.Text = "科目名称不能为空！";
                }
                else
                {
                    save();
                    hint.Text = etc.hint;
                    IFExecution_SUCCESS = etc.IFExecution_SUCCESS;
                    if (etc.IFExecution_SUCCESS && etc.ADD_OR_UPDATE =="ADD")
                    {
                        ClearText();
                        bind();
                    }
                    else if(etc.IFExecution_SUCCESS && etc.ADD_OR_UPDATE =="UPDATE")
                    {
                        bind();

                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "提示", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);

            }
        }
        private void btnExit_Click(object sender, EventArgs e)
        {
            this.Close();
        }
        #region btndel
        private void btnDel_Click(object sender, EventArgs e)
        {


            try
            {
                if (MessageBox.Show("确定要删除该条信息吗？", "提示", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) == DialogResult.OK)
                {
                    string v = textBox1.Text;
                    string v1 = bc.getOnlyStringO("ACCOUNTANT_COURSE", "ACNAME", "ACID", v);
                    string v2 = bc.getOnlyStringO("ACCOUNTANT_COURSE", "ACCODE", "ACID", v);

                   if (bc.exists("VOUCHER_DET", "ACID", v, "科目 " + v1 + " " + "已经有做帐记录不允许删除！"))
                    {

                    }

                    else
                    {
                        basec.getcoms("DELETE Accountant_Course WHERE ACID='" + v + "'");
                        bind();
                    }
                    //ClearText();
                    //textBox1.Text = "";
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "提示", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);

            }
            
        }
        #endregion
        #region dgvcellclick
        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {


        }
        #endregion

      


        private void button1_Click(object sender, EventArgs e)
        {
            dt = etc.GetCOURSE_TypeData(1);
            bind(dt);
            treeView1.Nodes.Clear();
            SHOW_TREEVIEW(dt);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            dt = etc.GetCOURSE_TypeData(2);
            bind(dt);
            treeView1.Nodes.Clear();
            SHOW_TREEVIEW(dt);
        }
        private void button3_Click(object sender, EventArgs e)
        {
            dt = etc.GetCOURSE_TypeData(3);
            bind(dt);
            treeView1.Nodes.Clear();
            SHOW_TREEVIEW(dt);
        }
        private void button4_Click(object sender, EventArgs e)
        {
            dt = etc.GetCOURSE_TypeData(4);
            bind(dt);
            treeView1.Nodes.Clear();
            SHOW_TREEVIEW(dt);
        }

        private void button5_Click(object sender, EventArgs e)
        {
            dt = etc.GetCOURSE_TypeData(5);
            bind(dt);
            treeView1.Nodes.Clear();
            SHOW_TREEVIEW(dt);
        }

        private void button6_Click(object sender, EventArgs e)
        {
            dt = etc.GetCOURSE_TypeData(6);
            bind(dt);
            treeView1.Nodes.Clear();
            SHOW_TREEVIEW(dt);
        }

        private void btnSearch_Click(object sender, EventArgs e)
        {
          
            dt = bc.getdt("SELECT * FROM ACCOUNTANT_COURSE ");
            dt = bc.GET_DT_TO_DV_TO_DT(dt, "", "ACCODE LIKE '%" +bc.REMOVE_NAME (comboBox2 .Text ) + "%' AND ACNAME LIKE '%" + textBox5 .Text + "%'");
            if (dt.Rows.Count > 0)
            {
                foreach (DataRow dr in dt.Rows)
                {

                    treeView1.Nodes.Clear();
                    SHOW_TREEVIEW(dt);
                }
            }
            else
            {
                treeView1.Nodes.Clear();
                MessageBox.Show("找不到所要搜索项！");

            }

            foreach (TreeNode trd in treeView1.Nodes)
            {
              

                    trd.ExpandAll();


                
                //MessageBox.Show(trd.Text);
            }

        }

        private void btnToExcel_Click(object sender, EventArgs e)
        {
           
        }
     private void aws(TreeNode trd )
     {
         //MessageBox.Show(trd.Text);
         if (trd.Text ==comboBox2 .Text )
         {
             trd.BackColor = c;
             trd.Checked = true;

          
         }
         foreach (TreeNode trd1 in trd.Nodes)
         {
          
             //MessageBox.Show(trd1.Text);
             aws(trd1);

         }



      }
        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            if (textBox2.Text == "")
            {

            }
            else if (textBox2.Text.Length > 4)
            {

                bind2(textBox2.Text.Substring(0, 4));


            }

        }

     
        public void a5()
        {
            select = 1;
        }

        private void btnAdd_Click(object sender, EventArgs e)
        {
            LoadAgain();
            textBox3.Focus();
            currency();
        
        }


        private void treeView1_AfterSelect(object sender, TreeViewEventArgs e)
        {
            TreeNode trd = treeView1.SelectedNode;
            //MessageBox.Show(trd.Index.ToString() + "-" + trd.Text);
            textBox2.Text = bc.REMOVE_NAME(trd.Text);
            textBox3.Text = bc.getOnlyString("SELECT ACNAME FROM ACCOUNTANT_COURSE WHERE ACCODE='" + bc.REMOVE_NAME(trd.Text) + "'");
            textBox1.Text = bc.getOnlyString("SELECT ACID FROM ACCOUNTANT_COURSE WHERE ACCODE='" + bc.REMOVE_NAME(trd.Text) + "'");
            comboBox1.Text = bc.getOnlyString("SELECT COURSE_NATURE FROM ACCOUNTANT_COURSE WHERE ACCODE='" + bc.REMOVE_NAME(trd.Text) + "'");
            bind2(bc.REMOVE_NAME(trd.Text));
            if (trd.IsExpanded)
            {

                trd.Collapse();

            }
            else
            {
                trd.Expand();

            }
        }

        private void treeView1_DoubleClick(object sender, EventArgs e)
        {
            if (treeView1.Enabled ==true )
            {
                if (select == 1)
                {
                    ACCODE = textBox2.Text;
                    VOUCHER_MANAGE.VOUCHERT.IF_DOUBLE_CLICK = true;
                    this.Close();
                }
            }
        }

        private void button7_Click(object sender, EventArgs e)
        {
            ACCODE = textBox2.Text;
            DataTable dtx = new DataTable();
            dtx.Columns.Add("科目代码", typeof(string));
            VOUCHER_MANAGE.VOUCHERT.IF_DOUBLE_CLICK = true;
            foreach (TreeNode   dr in treeView1.Nodes)
            {
               
                if (dr.Checked)
                {
                    
                    DataRow dr1 = dtx.NewRow();
                    dr1["科目代码"] = bc.REMOVE_NAME(dr.Text.ToString());
                    dtx.Rows.Add(dr1);
                }
            }
            VOUCHER_MANAGE.VOUCHER.GETDT_INFO = dtx;
            VOUCHER_MANAGE.VOUCHER.IF_DOUBLE_CLICK = true;
            this.Close();
           
        }

        private void textBox3_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar ==13)
            {
                save1();
            }
        }

      
    }
}

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
    public partial class ABSTRACT : Form
    {
        DataTable dt = new DataTable();
        DataTable dt2 = new DataTable();
        DataTable dt3 = new DataTable();

     
        protected int M_int_judge, t;
        basec bc = new basec();
        Color c = System.Drawing.ColorTranslator.FromHtml("#efdaec");
        CABSTRACT cABSTRACT = new CABSTRACT();
        PERIOD pe = new PERIOD();
        private string _PARENT_NODEID;
        public string PARENT_NODEID
        {
            set { _PARENT_NODEID = value; }
            get { return _PARENT_NODEID; }
        }
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
        private string _ABCODE;
        public string ABCODE
        {
            set { _ABCODE = value; }
            get { return _ABCODE; }
        }
        DataTable node_dt = new DataTable();
        Color c2 = System.Drawing.ColorTranslator.FromHtml("#990033");
        public ABSTRACT()
        {
            InitializeComponent();
          
           
        }
        private void ABSTRACT_Load(object sender, EventArgs e)
        {
            textBox3.BackColor = Color.Yellow;
            bind();
            label2.Text = "";
            label2.ForeColor = c2;
            treeView1.ContextMenuStrip = contextMenuStrip1;
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
                textBox1.Text = cABSTRACT.GETID();
            }
            treeView1.Nodes.Clear();
            dt = bc.getdt("SELECT * FROM ABSTRACT");
            SHOW_TREEVIEW(dt);
            textBox3.Focus();

            //this.WindowState = FormWindowState.Maximized;
            think();
            PARENT_NODEID = null;

        }
        #endregion
        #region think
        private void think()
        {

            dt2 = bc.getdt("SELECT * FROM ABSTRACT ");
            AutoCompleteStringCollection inputInfoSource = new AutoCompleteStringCollection();
            AutoCompleteStringCollection inputInfoSource4 = new AutoCompleteStringCollection();
            comboBox2.Items.Clear();
                foreach (DataRow dr in dt2.Rows)
                {

                    comboBox2.Items.Add(dr["ABCODE"].ToString() + " " + dr["ABSTRACT"].ToString());
                    inputInfoSource.Add(dr["ABCODE"].ToString() + " " + dr["ABSTRACT"].ToString());


                }
            this.comboBox2.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend;
            this.comboBox2.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.CustomSource;
            this.comboBox2.AutoCompleteCustomSource = inputInfoSource;


        }
        #endregion
        #region show_treeview
        private void SHOW_TREEVIEW(DataTable dt)
        {


            dt = bc.GET_DT_TO_DV_TO_DT(dt, "", "PARENT_NODEID IS NULL");

            if (dt.Rows.Count > 0)
            {
                foreach (DataRow dr in dt.Rows)
                {
                    TreeNode trd = treeView1.Nodes.Add(dr["ABCODE"].ToString()+" "+dr["ABSTRACT"].ToString());

                    if (bc.FROM_RIGHT_UNTIL_CHAR (trd.Text ,' ')==textBox3.Text)
                    {

                        trd.BackColor = c;

                    }
                    if (trd.Checked)
                    {
                        hint.Text = trd.Text.ToString();
                    }
                    SHOW_TREEVIEW_O(dr["ABID"].ToString(), trd);

                }

            }
        

        }
        #endregion

        #region show_treeview_O
        private void SHOW_TREEVIEW_O(string ABID, TreeNode trd)
        {

            dt2 = bc.getdt("SELECT * FROM ABSTRACT WHERE PARENT_NODEID='" + ABID + "'");
            if (dt2.Rows.Count > 0)
            {
                foreach (DataRow dr1 in dt2.Rows)
                {

                    TreeNode TRC = new TreeNode();
                    TRC.Text = dr1["ABCODE"].ToString() + " " + dr1["ABSTRACT"].ToString();
                    trd.Nodes.Add(TRC);
                    if (bc.FROM_RIGHT_UNTIL_CHAR (TRC.Text ,' ')==textBox3.Text)
                    {

                        TRC.BackColor = c;
                    }
                  
                    SHOW_TREEVIEW_O(dr1["ABID"].ToString(), TRC);

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
                    textBox1.Text = dt.Rows[0]["ABID"].ToString();
                    textBox2.Text = dt.Rows[0]["ABCODE"].ToString();
                    textBox3.Text = dt.Rows[0]["ABSTRACT"].ToString();
                    bind2(dt.Rows[0]["ABCODE"].ToString());
                  
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
        private void bind2(string ABCODE)
        {
       
        }
        #endregion
       
        #region save
        protected void save()
        {

            cABSTRACT.PARENT_NODEID = PARENT_NODEID;
            cABSTRACT.EMID = LOGIN.EMID;
            cABSTRACT.save(textBox1.Text, textBox2.Text, textBox3.Text, "", "", "","");
            ADD_OR_UPDATE = cABSTRACT.ADD_OR_UPDATE;
          
           
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
            textBox1.Text = cABSTRACT.GETID();
            hint.ForeColor = Color.Red;
            if (bc.GET_IFExecutionSUCCESS_HINT_INFO(IFExecution_SUCCESS) != "")
            {
                hint.Text = bc.GET_IFExecutionSUCCESS_HINT_INFO(IFExecution_SUCCESS);
            }
            else
            {
                hint.Text = "";
            }
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

            if (textBox3.Text == "")
            {

                hint.Text = "摘要名称不能为空！";
            }
            else
            {
                save();
                hint.Text = cABSTRACT.hint;
                IFExecution_SUCCESS = cABSTRACT.IFExecution_SUCCESS;
                if (cABSTRACT.IFExecution_SUCCESS && cABSTRACT.ADD_OR_UPDATE == "ADD")
                {
                    //ClearText();
                    //bind();
                     COURSE_TYPE_LOAD();
                }
                else if (cABSTRACT.IFExecution_SUCCESS && cABSTRACT.ADD_OR_UPDATE == "UPDATE")
                {
                    //bind();
                    COURSE_TYPE_LOAD();

                }
           

            }
            try
            {
           
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "提示", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);

            }
        }
        public void REQUEST_ABSTRACT()
        {
            select = 1;
        }
        private void COURSE_TYPE_LOAD()
        {
        
            think();
            textBox3.Focus();
            treeView1.Nodes.Clear();// no allow once again onload
            dt = bc.getdt("SELECT * FROM ABSTRACT");
            SHOW_TREEVIEW(dt);
            RETURN_PARENT_NODEID(textBox1.Text, null);
           
               foreach (TreeNode trd in treeView1.Nodes)
                { 
                   string v1 = bc.getOnlyString("SELECT ABID FROM ABSTRACT WHERE ABCODE='"+bc.REMOVE_NAME (trd.Text )+"'");
                    if (PARENT_NODEID ==v1  )
                    {
                        trd.ExpandAll();
                     
                    }
             
                }
            LoadAgain();
            PARENT_NODEID = null;
            //
        }

        public void  RETURN_PARENT_NODEID(string ID,string PARENT_NODEID_O)
        {
        
            string  id3 = bc.getOnlyString("SELECT PARENT_NODEID FROM ABSTRACT WHERE ABID='" + ID + "'");
            if (!string.IsNullOrEmpty(id3))
            {
                PARENT_NODEID = id3;
                RETURN_PARENT_NODEID(id3, id3);

            }
            else
            {
                PARENT_NODEID = ID;
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
                if (MessageBox.Show("提醒：如果删除的结点为根结点，那么删除根结点的同时将一同删除该根结点下的子结点，确定要删除该条信息吗？", "提示", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) == DialogResult.OK)
                {
                    string v = textBox1.Text;
                    string v1 = bc.getOnlyStringO("ABSTRACT", "ABSTRACT", "ABID", v);
                    string v2 = bc.getOnlyStringO("ABSTRACT", "ABCODE", "ABID", v);
                    basec.getcoms("DELETE ABSTRACT WHERE ABID='" + v + "'");
                    hint.Text = "删除成功";
                    bind();
                    ClearText();
                    textBox1.Text = "";
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

      


   

        private void btnSearch_Click(object sender, EventArgs e)
        {
          
            dt = bc.getdt("SELECT * FROM ABSTRACT ");
            dt = bc.GET_DT_TO_DV_TO_DT(dt, "", "ABCODE LIKE '%" +bc.REMOVE_NAME (comboBox2 .Text ) + "%' AND ABSTRACT LIKE '%" + textBox5 .Text + "%'");
            if (dt.Rows.Count > 0)
            {
                
                dt=RETURN_DATA(dt);
                
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
        private DataTable  RETURN_DATA(DataTable dt)
        {
            DataTable dtt = RETURN_EMPTY_DT();
            foreach (DataRow dr in dt.Rows)
            {
                DataRow dr1 = dtt.NewRow();
                RETURN_PARENT_NODEID(dr["ABID"].ToString(), null);
                dr1["ABID"] = PARENT_NODEID;
                DataTable dtx = bc.getdt("SELECT * FROM ABSTRACT WHERE ABID='"+PARENT_NODEID +"'");
                if (dtx.Rows.Count > 0)
                {
                    dr1["ABCODE"] = dtx.Rows[0]["ABCODE"].ToString();
                    dr1["ABSTRACT"] = dtx.Rows[0]["ABSTRACT"].ToString();

                }
                dtt.Rows.Add(dr1);
            }
            DataTable dtt1 = RETURN_EMPTY_DT();
            foreach (DataRow dr2 in dtt.Rows)
            {
                DataTable dtt12 = bc.GET_DT_TO_DV_TO_DT(dtt1, "", "ABID='"+dr2["ABID"].ToString ()+"'");
                if (dtt12.Rows.Count > 0)
                {

                }
                else
                {
                    DataRow dr3 = dtt1.NewRow();
                    dr3["ABID"] = dr2["ABID"].ToString();
                    dr3["ABCODE"] =dr2["ABCODE"].ToString();
                    dr3["ABSTRACT"] =dr2["ABSTRACT"].ToString();
                    dtt1.Rows.Add(dr3);
                }
            }
            return dtt1;
        }
        private DataTable RETURN_EMPTY_DT()
        {
            DataTable dtt = new DataTable();
            dtt.Columns.Add("ABID", typeof(string));
            dtt.Columns.Add("ABCODE", typeof(string));
            dtt.Columns.Add("ABSTRACT", typeof(string));
            dtt.Columns.Add("PARENT_NODEID", typeof(string));
            return dtt;
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
            textBox3.Text = bc.getOnlyString("SELECT ABSTRACT FROM ABSTRACT WHERE ABCODE='" + bc.REMOVE_NAME(trd.Text) + "'");
            textBox1.Text = bc.getOnlyString("SELECT ABID FROM ABSTRACT WHERE ABCODE='" + bc.REMOVE_NAME(trd.Text) + "'");
         
            bind2(bc.REMOVE_NAME(trd.Text));
            if (trd.IsExpanded)
            {

                trd.Collapse();

            }
            else
            {
                trd.Expand();

            }
           listBox1.Items.Add(bc.FROM_RIGHT_UNTIL_CHAR (trd.Text,' '));
        }
        private DataTable dtxa()
        {
            DataTable dtx = new DataTable();
            dtx.Columns.Add("摘要代码", typeof(string));
            return dtx;

        }
        private void treeView1_DoubleClick(object sender, EventArgs e)
        {
            if (treeView1.Enabled ==true )
            {
                if (select == 1)
                {
                    ABCODE = textBox3.Text;
                    VOUCHER_MANAGE.VOUCHERT.IF_DOUBLE_CLICK = true;
                    this.Close();
                }
            }
        }

        private void button7_Click(object sender, EventArgs e)
        {
            ABCODE = textBox3.Text;
            DataTable dtx = new DataTable();
            dtx.Columns.Add("摘要代码", typeof(string));
            for (int i = 0; i < listBox1.Items.Count; i++)
            {
                DataRow dr1 = dtx.NewRow();
                dr1["摘要代码"] = listBox1.Items[i].ToString();
                dtx.Rows.Add(dr1);
            }
     
            VOUCHER_MANAGE.VOUCHERT.IF_DOUBLE_CLICK = true;
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

        private void 新增二级摘要ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            LoadAgain();
            textBox3.Focus();
            hint.Text = "";
            TreeNode trd = treeView1.SelectedNode;
            //MessageBox.Show(trd.Index.ToString() + "-" + trd.Text);
            if (trd != null)
            {
                if (trd.Text.Length > 0)
                {
                    PARENT_NODEID = bc.getOnlyString("SELECT ABID FROM ABSTRACT WHERE ABCODE='" + bc.REMOVE_NAME(trd.Text) + "'");
                }
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            listBox1.Items.Clear();
        }







    }
}

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
using XizheC;

namespace CSPSS.VOUCHER_MANAGE
{
    public partial class VOUCHER : Form
    {
        DataTable dt = new DataTable();
        DataTable dt1 = new DataTable();
        DataTable dt2 = new DataTable();
        DataTable dtx = new DataTable();
        basec bc = new basec();
      
        

        protected int M_int_judge, i, look;
        protected int getdata;

        Color c2 = System.Drawing.ColorTranslator.FromHtml("#990033");
        CVOUCHER vou = new CVOUCHER();
        private static DataTable _GETDT_INFO;
        public  static DataTable GETDT_INFO
        {
            set { _GETDT_INFO = value; }
            get { return _GETDT_INFO; }

        }
        private static bool _IF_DOUBLE_CLICK;
        public static bool IF_DOUBLE_CLICK
        {
            set { _IF_DOUBLE_CLICK = value; }
            get { return _IF_DOUBLE_CLICK; }

        }
        public VOUCHER()
        {
            InitializeComponent();
        }
        #region init
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(VOUCHER));
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.comboBox5 = new System.Windows.Forms.ComboBox();
            this.label16 = new System.Windows.Forms.Label();
            this.comboBox4 = new System.Windows.Forms.ComboBox();
            this.label10 = new System.Windows.Forms.Label();
            this.label9 = new System.Windows.Forms.Label();
            this.textBox2 = new System.Windows.Forms.TextBox();
            this.label8 = new System.Windows.Forms.Label();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.comboBox3 = new System.Windows.Forms.ComboBox();
            this.label5 = new System.Windows.Forms.Label();
            this.chk2 = new System.Windows.Forms.CheckBox();
            this.comboBox1 = new System.Windows.Forms.ComboBox();
            this.label3 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.dtpEndDate = new System.Windows.Forms.DateTimePicker();
            this.dtpStartDate = new System.Windows.Forms.DateTimePicker();
            this.label1 = new System.Windows.Forms.Label();
            this.comboBox2 = new System.Windows.Forms.ComboBox();
            this.label6 = new System.Windows.Forms.Label();
            this.label7 = new System.Windows.Forms.Label();
            this.btnToExcel = new System.Windows.Forms.Button();
            this.textBox5 = new System.Windows.Forms.TextBox();
            this.label11 = new System.Windows.Forms.Label();
            this.label12 = new System.Windows.Forms.Label();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.label18 = new System.Windows.Forms.Label();
            this.btnDel = new System.Windows.Forms.PictureBox();
            this.lkgeneral_manage = new System.Windows.Forms.LinkLabel();
            this.hint = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.label17 = new System.Windows.Forms.Label();
            this.btnAdd = new System.Windows.Forms.PictureBox();
            this.btnExit = new System.Windows.Forms.PictureBox();
            this.btnSearch = new System.Windows.Forms.PictureBox();
            this.label13 = new System.Windows.Forms.Label();
            this.textBox3 = new System.Windows.Forms.TextBox();
            this.label14 = new System.Windows.Forms.Label();
            this.textBox4 = new System.Windows.Forms.TextBox();
            this.label15 = new System.Windows.Forms.Label();
            this.textBox6 = new System.Windows.Forms.TextBox();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.groupBox1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.btnDel)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.btnAdd)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.btnExit)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.btnSearch)).BeginInit();
            this.SuspendLayout();
            // 
            // dataGridView1
            // 
            this.dataGridView1.AllowUserToAddRows = false;
            this.dataGridView1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.dataGridView1.BackgroundColor = System.Drawing.Color.FromArgb(((int)(((byte)(238)))), ((int)(((byte)(245)))), ((int)(((byte)(255)))));
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.EditMode = System.Windows.Forms.DataGridViewEditMode.EditOnEnter;
            this.dataGridView1.Location = new System.Drawing.Point(0, 266);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.RowTemplate.Height = 23;
            this.dataGridView1.Size = new System.Drawing.Size(943, 312);
            this.dataGridView1.TabIndex = 0;
            this.dataGridView1.DataSourceChanged += new System.EventHandler(this.dataGridView1_DataSourceChanged);
            this.dataGridView1.CellClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridView1_CellClick);
            this.dataGridView1.DoubleClick += new System.EventHandler(this.dataGridView1_DoubleClick);
            // 
            // groupBox1
            // 
            this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.groupBox1.Controls.Add(this.comboBox5);
            this.groupBox1.Controls.Add(this.label16);
            this.groupBox1.Controls.Add(this.comboBox4);
            this.groupBox1.Controls.Add(this.label10);
            this.groupBox1.Controls.Add(this.label9);
            this.groupBox1.Controls.Add(this.textBox2);
            this.groupBox1.Controls.Add(this.label8);
            this.groupBox1.Controls.Add(this.textBox1);
            this.groupBox1.Controls.Add(this.comboBox3);
            this.groupBox1.Controls.Add(this.label5);
            this.groupBox1.Controls.Add(this.chk2);
            this.groupBox1.Controls.Add(this.comboBox1);
            this.groupBox1.Controls.Add(this.label3);
            this.groupBox1.Controls.Add(this.label2);
            this.groupBox1.Controls.Add(this.dtpEndDate);
            this.groupBox1.Controls.Add(this.dtpStartDate);
            this.groupBox1.Controls.Add(this.label1);
            this.groupBox1.Controls.Add(this.comboBox2);
            this.groupBox1.Controls.Add(this.label6);
            this.groupBox1.Controls.Add(this.label7);
            this.groupBox1.Controls.Add(this.btnToExcel);
            this.groupBox1.Controls.Add(this.textBox5);
            this.groupBox1.Location = new System.Drawing.Point(3, 127);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(936, 133);
            this.groupBox1.TabIndex = 4;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "查询条件";
            // 
            // comboBox5
            // 
            this.comboBox5.Cursor = System.Windows.Forms.Cursors.Default;
            this.comboBox5.Location = new System.Drawing.Point(137, 107);
            this.comboBox5.Name = "comboBox5";
            this.comboBox5.Size = new System.Drawing.Size(173, 20);
            this.comboBox5.TabIndex = 60;
            this.comboBox5.DropDown += new System.EventHandler(this.comboBox5_DropDown);
            // 
            // label16
            // 
            this.label16.AutoSize = true;
            this.label16.Location = new System.Drawing.Point(78, 115);
            this.label16.Name = "label16";
            this.label16.Size = new System.Drawing.Size(29, 12);
            this.label16.TabIndex = 61;
            this.label16.Text = "摘要";
            // 
            // comboBox4
            // 
            this.comboBox4.Cursor = System.Windows.Forms.Cursors.Default;
            this.comboBox4.Items.AddRange(new object[] {
            "",
            "未打款",
            "已打款"});
            this.comboBox4.Location = new System.Drawing.Point(662, 80);
            this.comboBox4.Name = "comboBox4";
            this.comboBox4.Size = new System.Drawing.Size(121, 20);
            this.comboBox4.TabIndex = 58;
            // 
            // label10
            // 
            this.label10.AutoSize = true;
            this.label10.Location = new System.Drawing.Point(603, 88);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(53, 12);
            this.label10.TabIndex = 59;
            this.label10.Text = "是否打款";
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Location = new System.Drawing.Point(614, 60);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(41, 12);
            this.label9.TabIndex = 57;
            this.label9.Text = "制单人";
            // 
            // textBox2
            // 
            this.textBox2.Location = new System.Drawing.Point(662, 51);
            this.textBox2.Name = "textBox2";
            this.textBox2.Size = new System.Drawing.Size(121, 21);
            this.textBox2.TabIndex = 56;
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Location = new System.Drawing.Point(78, 88);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(41, 12);
            this.label8.TabIndex = 55;
            this.label8.Text = "凭证号";
            // 
            // textBox1
            // 
            this.textBox1.Location = new System.Drawing.Point(137, 79);
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(173, 21);
            this.textBox1.TabIndex = 54;
            // 
            // comboBox3
            // 
            this.comboBox3.Cursor = System.Windows.Forms.Cursors.Default;
            this.comboBox3.Items.AddRange(new object[] {
            "",
            "正常",
            "零用金"});
            this.comboBox3.Location = new System.Drawing.Point(423, 79);
            this.comboBox3.Name = "comboBox3";
            this.comboBox3.Size = new System.Drawing.Size(121, 20);
            this.comboBox3.TabIndex = 52;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(364, 87);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(53, 12);
            this.label5.TabIndex = 53;
            this.label5.Text = "科目性质";
            // 
            // chk2
            // 
            this.chk2.AutoSize = true;
            this.chk2.Location = new System.Drawing.Point(343, 31);
            this.chk2.Name = "chk2";
            this.chk2.Size = new System.Drawing.Size(15, 14);
            this.chk2.TabIndex = 51;
            this.chk2.UseVisualStyleBackColor = true;
            this.chk2.CheckedChanged += new System.EventHandler(this.chk2_CheckedChanged);
            // 
            // comboBox1
            // 
            this.comboBox1.Cursor = System.Windows.Forms.Cursors.Default;
            this.comboBox1.Items.AddRange(new object[] {
            "",
            "经理未审核",
            "经理已审核",
            "财务未审核",
            "财务已审核",
            "总经理未审核",
            "总经理已审核"});
            this.comboBox1.Location = new System.Drawing.Point(423, 51);
            this.comboBox1.Name = "comboBox1";
            this.comboBox1.Size = new System.Drawing.Size(121, 20);
            this.comboBox1.TabIndex = 35;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(388, 59);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(29, 12);
            this.label3.TabIndex = 50;
            this.label3.Text = "状态";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(598, 31);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(11, 12);
            this.label2.TabIndex = 49;
            this.label2.Text = "~";
            // 
            // dtpEndDate
            // 
            this.dtpEndDate.CustomFormat = "yyyy/MM/dd";
            this.dtpEndDate.Enabled = false;
            this.dtpEndDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
            this.dtpEndDate.Location = new System.Drawing.Point(662, 22);
            this.dtpEndDate.Name = "dtpEndDate";
            this.dtpEndDate.Size = new System.Drawing.Size(121, 21);
            this.dtpEndDate.TabIndex = 48;
            // 
            // dtpStartDate
            // 
            this.dtpStartDate.CustomFormat = "yyyy/MM/dd";
            this.dtpStartDate.Enabled = false;
            this.dtpStartDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
            this.dtpStartDate.Location = new System.Drawing.Point(423, 22);
            this.dtpStartDate.Name = "dtpStartDate";
            this.dtpStartDate.Size = new System.Drawing.Size(121, 21);
            this.dtpStartDate.TabIndex = 47;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(364, 31);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(53, 12);
            this.label1.TabIndex = 46;
            this.label1.Text = "凭证期间";
            // 
            // comboBox2
            // 
            this.comboBox2.Cursor = System.Windows.Forms.Cursors.Default;
            this.comboBox2.Location = new System.Drawing.Point(137, 20);
            this.comboBox2.Name = "comboBox2";
            this.comboBox2.Size = new System.Drawing.Size(173, 20);
            this.comboBox2.TabIndex = 5;
            this.comboBox2.DropDown += new System.EventHandler(this.comboBox2_DropDown);
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(78, 28);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(53, 12);
            this.label6.TabIndex = 44;
            this.label6.Text = "科目代码";
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(78, 58);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(53, 12);
            this.label7.TabIndex = 45;
            this.label7.Text = "科目名称";
            // 
            // btnToExcel
            // 
            this.btnToExcel.FlatAppearance.BorderSize = 0;
            this.btnToExcel.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnToExcel.Font = new System.Drawing.Font("宋体", 9F);
            this.btnToExcel.Image = ((System.Drawing.Image)(resources.GetObject("btnToExcel.Image")));
            this.btnToExcel.ImageAlign = System.Drawing.ContentAlignment.TopCenter;
            this.btnToExcel.Location = new System.Drawing.Point(841, 13);
            this.btnToExcel.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            this.btnToExcel.Name = "btnToExcel";
            this.btnToExcel.Size = new System.Drawing.Size(50, 64);
            this.btnToExcel.TabIndex = 11;
            this.btnToExcel.Text = "导出";
            this.btnToExcel.TextAlign = System.Drawing.ContentAlignment.BottomCenter;
            this.btnToExcel.UseVisualStyleBackColor = false;
            this.btnToExcel.Click += new System.EventHandler(this.btnToExcel_Click);
            // 
            // textBox5
            // 
            this.textBox5.Location = new System.Drawing.Point(137, 49);
            this.textBox5.Name = "textBox5";
            this.textBox5.Size = new System.Drawing.Size(173, 21);
            this.textBox5.TabIndex = 6;
            // 
            // label11
            // 
            this.label11.AutoSize = true;
            this.label11.Location = new System.Drawing.Point(857, 95);
            this.label11.Name = "label11";
            this.label11.Size = new System.Drawing.Size(29, 12);
            this.label11.TabIndex = 29;
            this.label11.Text = "退出";
            // 
            // label12
            // 
            this.label12.AutoSize = true;
            this.label12.Location = new System.Drawing.Point(771, 95);
            this.label12.Name = "label12";
            this.label12.Size = new System.Drawing.Size(29, 12);
            this.label12.TabIndex = 28;
            this.label12.Text = "搜索";
            // 
            // groupBox2
            // 
            this.groupBox2.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.groupBox2.Controls.Add(this.label18);
            this.groupBox2.Controls.Add(this.btnDel);
            this.groupBox2.Controls.Add(this.lkgeneral_manage);
            this.groupBox2.Controls.Add(this.hint);
            this.groupBox2.Controls.Add(this.label4);
            this.groupBox2.Controls.Add(this.label11);
            this.groupBox2.Controls.Add(this.label12);
            this.groupBox2.Controls.Add(this.label17);
            this.groupBox2.Controls.Add(this.btnAdd);
            this.groupBox2.Controls.Add(this.btnExit);
            this.groupBox2.Controls.Add(this.btnSearch);
            this.groupBox2.Location = new System.Drawing.Point(3, 3);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(936, 121);
            this.groupBox2.TabIndex = 34;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "菜单栏";
            // 
            // label18
            // 
            this.label18.AutoSize = true;
            this.label18.Location = new System.Drawing.Point(673, 95);
            this.label18.Name = "label18";
            this.label18.Size = new System.Drawing.Size(53, 12);
            this.label18.TabIndex = 405;
            this.label18.Text = "批量删除";
            // 
            // btnDel
            // 
            this.btnDel.Image = ((System.Drawing.Image)(resources.GetObject("btnDel.Image")));
            this.btnDel.InitialImage = null;
            this.btnDel.Location = new System.Drawing.Point(671, 20);
            this.btnDel.Name = "btnDel";
            this.btnDel.Size = new System.Drawing.Size(60, 60);
            this.btnDel.TabIndex = 404;
            this.btnDel.TabStop = false;
            this.btnDel.Click += new System.EventHandler(this.btnDel_Click);
            // 
            // lkgeneral_manage
            // 
            this.lkgeneral_manage.AutoSize = true;
            this.lkgeneral_manage.Location = new System.Drawing.Point(501, 68);
            this.lkgeneral_manage.Name = "lkgeneral_manage";
            this.lkgeneral_manage.Size = new System.Drawing.Size(89, 12);
            this.lkgeneral_manage.TabIndex = 403;
            this.lkgeneral_manage.TabStop = true;
            this.lkgeneral_manage.Text = "总经理批量审核";
            this.lkgeneral_manage.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.lkgeneral_manage_LinkClicked);
            // 
            // hint
            // 
            this.hint.AutoSize = true;
            this.hint.Location = new System.Drawing.Point(196, 59);
            this.hint.Name = "hint";
            this.hint.Size = new System.Drawing.Size(29, 12);
            this.hint.TabIndex = 402;
            this.hint.Text = "hint";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(184, 95);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(41, 12);
            this.label4.TabIndex = 401;
            this.label4.Text = "label4";
            // 
            // label17
            // 
            this.label17.AutoSize = true;
            this.label17.Location = new System.Drawing.Point(28, 95);
            this.label17.Name = "label17";
            this.label17.Size = new System.Drawing.Size(29, 12);
            this.label17.TabIndex = 24;
            this.label17.Text = "新增";
            // 
            // btnAdd
            // 
            this.btnAdd.Image = ((System.Drawing.Image)(resources.GetObject("btnAdd.Image")));
            this.btnAdd.InitialImage = null;
            this.btnAdd.Location = new System.Drawing.Point(12, 20);
            this.btnAdd.Name = "btnAdd";
            this.btnAdd.Size = new System.Drawing.Size(60, 60);
            this.btnAdd.TabIndex = 16;
            this.btnAdd.TabStop = false;
            this.btnAdd.Click += new System.EventHandler(this.btnAdd_Click);
            // 
            // btnExit
            // 
            this.btnExit.Image = ((System.Drawing.Image)(resources.GetObject("btnExit.Image")));
            this.btnExit.InitialImage = null;
            this.btnExit.Location = new System.Drawing.Point(843, 20);
            this.btnExit.Name = "btnExit";
            this.btnExit.Size = new System.Drawing.Size(60, 60);
            this.btnExit.TabIndex = 19;
            this.btnExit.TabStop = false;
            this.btnExit.Click += new System.EventHandler(this.btnExit_Click);
            // 
            // btnSearch
            // 
            this.btnSearch.Image = ((System.Drawing.Image)(resources.GetObject("btnSearch.Image")));
            this.btnSearch.InitialImage = null;
            this.btnSearch.Location = new System.Drawing.Point(757, 20);
            this.btnSearch.Name = "btnSearch";
            this.btnSearch.Size = new System.Drawing.Size(60, 60);
            this.btnSearch.TabIndex = 18;
            this.btnSearch.TabStop = false;
            this.btnSearch.Click += new System.EventHandler(this.btnSearch_Click);
            // 
            // label13
            // 
            this.label13.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.label13.AutoSize = true;
            this.label13.Location = new System.Drawing.Point(92, 591);
            this.label13.Name = "label13";
            this.label13.Size = new System.Drawing.Size(53, 12);
            this.label13.TabIndex = 57;
            this.label13.Text = "合计支出";
            // 
            // textBox3
            // 
            this.textBox3.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.textBox3.Location = new System.Drawing.Point(151, 583);
            this.textBox3.Name = "textBox3";
            this.textBox3.Size = new System.Drawing.Size(115, 21);
            this.textBox3.TabIndex = 56;
            // 
            // label14
            // 
            this.label14.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.label14.AutoSize = true;
            this.label14.Location = new System.Drawing.Point(362, 590);
            this.label14.Name = "label14";
            this.label14.Size = new System.Drawing.Size(53, 12);
            this.label14.TabIndex = 59;
            this.label14.Text = "合计收入";
            // 
            // textBox4
            // 
            this.textBox4.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.textBox4.Location = new System.Drawing.Point(421, 582);
            this.textBox4.Name = "textBox4";
            this.textBox4.Size = new System.Drawing.Size(112, 21);
            this.textBox4.TabIndex = 58;
            // 
            // label15
            // 
            this.label15.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.label15.AutoSize = true;
            this.label15.Location = new System.Drawing.Point(634, 592);
            this.label15.Name = "label15";
            this.label15.Size = new System.Drawing.Size(53, 12);
            this.label15.TabIndex = 61;
            this.label15.Text = "合计余额";
            // 
            // textBox6
            // 
            this.textBox6.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.textBox6.Location = new System.Drawing.Point(693, 584);
            this.textBox6.Name = "textBox6";
            this.textBox6.Size = new System.Drawing.Size(113, 21);
            this.textBox6.TabIndex = 60;
            // 
            // VOUCHER
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(238)))), ((int)(((byte)(245)))), ((int)(((byte)(255)))));
            this.ClientSize = new System.Drawing.Size(942, 616);
            this.Controls.Add(this.label15);
            this.Controls.Add(this.textBox6);
            this.Controls.Add(this.label14);
            this.Controls.Add(this.textBox4);
            this.Controls.Add(this.label13);
            this.Controls.Add(this.textBox3);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.dataGridView1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "VOUCHER";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "凭证查询作业";
            this.Load += new System.EventHandler(this.VOUCHER_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.btnDel)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.btnAdd)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.btnExit)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.btnSearch)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }
        #endregion
  
    
        private void VOUCHER_Load(object sender, EventArgs e)
        {

            Bind();
        }
        #region Bind
        public void Bind()
        {
            textBox3.BackColor = Color.Yellow;
            textBox4.BackColor = Color.Yellow;
            textBox6.BackColor = Color.Yellow;
            textBox3.TextAlign = HorizontalAlignment.Right;
            textBox4.TextAlign = HorizontalAlignment.Right;
            textBox6.TextAlign = HorizontalAlignment.Right;
            textBox3.ReadOnly = true;
            textBox4.ReadOnly = true;
            textBox6.ReadOnly = true;
            string v1 = bc.getOnlyString("SELECT ADD_NEW FROM RIGHTLIST WHERE USID='" + LOGIN.USID + "' AND NODE_NAME='录入凭证作业'");
            string v3 = bc.getOnlyString("SELECT DEL FROM RIGHTLIST WHERE  USID='" + LOGIN.USID + "' AND NODE_NAME='录入凭证作业'");
            string v6 = bc.getOnlyString("SELECT GENERAL_MANAGE FROM RIGHTLIST WHERE  USID='" + LOGIN.USID + "' AND NODE_NAME='录入凭证作业'");
            if (v1 == "Y")
            {
                btnAdd.Visible = true;

                label17.Visible = true;


            }
            else
            {
                btnAdd.Visible = false;

                label17.Visible = false;
            }
            if (v3 == "Y" )
            {
                btnDel.Visible = true;
                label18.Visible = true;
            }
            else
            {
                btnDel.Visible = false;
                label18.Visible = false;

            }
            if (v6 == "Y")
            {
                lkgeneral_manage.Visible = true;
            }
            else
            {
                lkgeneral_manage.Visible = false;

            }
            this.WindowState = FormWindowState.Maximized;
            //dt = basec.getdts(vou.getsqlX + " WHERE B.STATUS!='INITIAL' ORDER BY A.VOKEY");
            //dataGridView1.DataSource = dt;
            dt1 = bc.getdt("SELECT VOID FROM VOUCHER_MST");
            AutoCompleteStringCollection inputInfoSource = new AutoCompleteStringCollection();
            //dgvStateControl();
            think();
            hint.Text = "";
            hint.ForeColor = Color.Red;
            try
            {
            
              
            }
            catch (Exception)
            {


            }
            label4.Text = "(状态为总经理已审核或科目性质为零用金且财务已审核时表明已入账)";
            label4.ForeColor = c2;
        }
        #endregion
        #region think
        private void think()
        {

            dt2 = bc.getdt("SELECT * FROM ACCOUNTANT_COURSE ");
            AutoCompleteStringCollection inputInfoSource = new AutoCompleteStringCollection();
       
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
        #region dgvStateControl
        private void dgvStateControl()
        {
            dataGridView1.Columns["复选框"].Width = 50;
            dataGridView1.Columns["单价"].Width = 40;
            dataGridView1.Columns["数量"].Width = 40;
            //dataGridView1.Columns["币别"].Width = 40;
            dataGridView1.Columns["项次"].Width = 40;
            //dataGridView1.Columns["汇率"].Width = 40;
            dataGridView1.Columns["科目代码"].Width = 100;
            dataGridView1.Columns["科目名称"].Width = 120;
            //dataGridView1.Columns["日期"].Width = 80;
            dataGridView1.Columns["凭证号"].Width = 80;
            dataGridView1.Columns["凭证日期"].Width = 80;
            dataGridView1.Columns["状态"].Width = 80;
            dataGridView1.Columns["科目性质"].Width = 70;
            dataGridView1.Columns["摘要"].Width = 200;
            dataGridView1.Columns["支出金额"].Width = 80;
            dataGridView1.Columns["收入金额"].Width = 80;
            dataGridView1.Columns["公司余额"].Width = 80;
            dataGridView1.Columns["是否打款"].Width = 80;
            dataGridView1.Columns["制单人"].Width = 80;
            dataGridView1.Columns["制单日期"].Width = 120;
            dataGridView1.Columns["单价"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dataGridView1.Columns["支出金额"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dataGridView1.Columns["收入金额"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dataGridView1.Columns["公司余额"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            int i;
            dataGridView1.RowHeadersDefaultCellStyle.BackColor = Color.Lavender;
            int numCols1 = dataGridView1.Columns.Count;
            for (i = 0; i < numCols1; i++)
            {

                dataGridView1.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                this.dataGridView1.Columns[i].SortMode = DataGridViewColumnSortMode.NotSortable;
    

                dataGridView1.EnableHeadersVisualStyles = false;
                dataGridView1.Columns[i].HeaderCell.Style.BackColor = Color.Lavender;

            }
            for (i = 0; i < dataGridView1.Columns.Count; i++)
            {
                dataGridView1.Columns[i].SortMode = DataGridViewColumnSortMode.NotSortable;
                dataGridView1.Columns[i].DefaultCellStyle.BackColor = Color.OldLace;
                i = i + 1;
            }

            for (i = 0; i < dataGridView1.Columns.Count; i++)
            {
                dataGridView1.Columns[i].SortMode = DataGridViewColumnSortMode.NotSortable;
                if (i !=0)
                {
                    dataGridView1.Columns[i].ReadOnly = true;
                }
             
                if (i == 0)
                {
                    dataGridView1.Columns[i].Visible = true;

                }
            }

            
        }
        #endregion

        #region add

        #endregion
  


        #region override enter
        protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
        {
            if (keyData == Keys.Enter &&
             (
             (
              !(ActiveControl is System.Windows.Forms.TextBox) ||
              !((System.Windows.Forms.TextBox)ActiveControl).AcceptsReturn)
             )
             )
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


        #region doubleclick
        private void dataGridView1_DoubleClick(object sender, EventArgs e)
        {

            if (getdata != 0)
            {
                if (getdata == 1)
                {
                    int intCurrentRowNumber = this.dataGridView1.CurrentCell.RowIndex;
                    string s1 = this.dataGridView1.Rows[intCurrentRowNumber].Cells[0].Value.ToString().Trim();
                    string s2 = this.dataGridView1.Rows[intCurrentRowNumber].Cells[18].Value.ToString().Trim();
                    string s3 = this.dataGridView1.Rows[intCurrentRowNumber].Cells[19].Value.ToString().Trim();
                    string s4 = this.dataGridView1.Rows[intCurrentRowNumber].Cells[20].Value.ToString().Trim();
                    string s5 = this.dataGridView1.Rows[intCurrentRowNumber].Cells[21].Value.ToString().Trim();
                    /*CSPSS.VOUCHER_MANAGE.FrmSellTableT.data4[0] = "doubleclick";
                    CSPSS.VOUCHER_MANAGE.FrmSellTableT.data1[0] = s1;
                    CSPSS.VOUCHER_MANAGE.FrmSellTableT.data1[1] = s2;
                    CSPSS.VOUCHER_MANAGE.FrmSellTableT.data1[2] = s3;
                    CSPSS.VOUCHER_MANAGE.FrmSellTableT.data1[3] = s4;
                    CSPSS.VOUCHER_MANAGE.FrmSellTableT.data1[4] = s5;*/

                    this.Close();
                }

            }
            else
            {
                VOUCHERT frm = new VOUCHERT(this);
                frm.IDO  = dt.Rows[dataGridView1.CurrentCell.RowIndex]["凭证号"].ToString();
                frm.ADD_OR_UPDATE = "UPDATE";
                frm.ShowDialog();
            }
        }
        #endregion
        public void a2()
        {

            getdata = 1;

        }
        public void a3()
        {

            getdata = 2;

        }
   
        private void btnToExcel_Click(object sender, EventArgs e)
        {
            if (dt.Rows.Count > 0)
            {
               
                bc.dgvtoExcel(dataGridView1, "凭证明细");
                
            }
            else
            {
                MessageBox.Show("没有数据可导出！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void dataGridView1_DataSourceChanged(object sender, EventArgs e)
        {
            int i;
            for (i = 0; i < dataGridView1.Columns.Count; i++)
            {
                if (dataGridView1.Columns[i].ValueType.ToString() == "System.Decimal")
                {
                    dataGridView1.Columns[i].DefaultCellStyle.Format = "#0.0000";
                    dataGridView1.Columns[i].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                }

            }
        }

        private void btnAdd_Click(object sender, EventArgs e)
        {
            string a1 = bc.numYMD(12, 4, "0001", "select * from VOUCHER_MST", "VOID", "VO");
            if (a1 == "Exceed Limited")
            {
                MessageBox.Show("编码超出限制！", "", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            else
            {

               
                VOUCHERT frm = new VOUCHERT(this);
                frm.IDO = a1;
                frm.ADD_OR_UPDATE = "ADD";
                frm.ShowDialog();

            }
        }

        private void btnSearch_Click(object sender, EventArgs e)
        {
            
            try
            {
                search();
            }
            catch (Exception)
            {
                MessageBox.Show("不能出现(单引号+变量+单引号)+变量的格式");

            }

        }
        #region search()
        public void search()
        {
            string v1 = dtpStartDate.Value.ToString("yyyy/MM/dd 00:00:00").Replace("-", "/");
            string v2 = dtpEndDate.Value.ToString("yyyy/MM/dd 23:59:59").Replace("-", "/");

            string v3 = dtpStartDate.Value.ToString("yyyy/MM/dd").Replace("-", "/");
            string v4 = dtpEndDate.Value.ToString("yyyy/MM/dd").Replace("-", "/");
            dt = basec.getdts(vou.getsqlX + " ORDER BY C.ACNAME ASC");
            string v5=comboBox1 .Text ;
            string v6 = "";
            string v7="";
            string v8 = "";
        
            hint.Text = "";
            if (bc.JUAGE_IF_EXISTS_ANYONE_CHAR (comboBox2 .Text ,','))
            {
                v7 = " C.ACCODE IN " + "(" + comboBox2.Text + ")";
            }
            else if (bc.JUAGE_IF_EXISTS_ANYONE_CHAR(comboBox2.Text, 39))
            {
                v7 = " C.ACCODE IN " + "(" + comboBox2.Text + ")";

            }
            else
            {
                v7 = " C.ACCODE LIKE '%"+bc.REMOVE_NAME (comboBox2 .Text) +"%'";

            }
            if (bc.JUAGE_IF_EXISTS_ANYONE_CHAR(comboBox5.Text, ','))
            {
                v8 = " A.ABSTRACT IN " + "(" + comboBox5.Text + ")";
            }
            else if (bc.JUAGE_IF_EXISTS_ANYONE_CHAR(comboBox5.Text, 39))
            {
                v8 = " A.ABSTRACT IN " + "(" + comboBox5.Text + ")";
            }
            else
            {
                v8 = " A.ABSTRACT LIKE '%" + bc.REMOVE_NAME(comboBox5.Text) + "%'";
            }
            if (comboBox4.Text == "未打款")
            {
                v6 = "N";
            }
            else if(comboBox4 .Text =="已打款")
            {
                v6 = "Y";
            }
        
            if (chk2.Checked)
            {
             
                if (comboBox1.Text == "经理未审核")
                {


                    search_o(vou.getsqlX + " WHERE  "+v7+" AND "+v8+"  AND C.ACNAME LIKE '%" + textBox5.Text +
                        "%' AND B.MANAGE_AUDIT_STATUS='N'  AND B.VOUCHER_DATE BETWEEN '" + v3 + "' AND '" + v4 +
                        "' AND C.COURSE_NATURE LIKE '%" + comboBox3.Text + "%' AND A.VOID LIKE '%" + textBox1.Text +
                        "%' AND B.IF_PAYFOR LIKE '%"+v6+"%' AND E.ENAME LIKE '%"+textBox2 .Text +"%'");

                }
                else    if (comboBox1.Text == "经理已审核")
                {
                  

                    search_o(vou.getsqlX + " WHERE  "+v7+" AND "+v8+"  AND C.ACNAME LIKE '%" + textBox5.Text +
                        "%' AND B.MANAGE_AUDIT_STATUS='Y' AND B.FINANCIAL_AUDIT_STATUS='N' AND B.GENERAL_MANAGE_AUDIT_STATUS='N' AND B.VOUCHER_DATE BETWEEN '" + v3 + "' AND '" + v4 +
                        "' AND C.COURSE_NATURE LIKE '%" + comboBox3.Text + "%' AND A.VOID LIKE '%" + textBox1.Text +
                        "%' AND B.IF_PAYFOR LIKE '%" + v6 + "%' AND E.ENAME LIKE '%" + textBox2.Text + "%'");
                     
                }
                else if (comboBox1.Text == "财务未审核")
                {

                    search_o(vou.getsqlX + " WHERE  "+v7+" AND "+v8+"  AND C.ACNAME LIKE '%" + textBox5.Text +
                        "%' AND B.FINANCIAL_AUDIT_STATUS='N' AND B.VOUCHER_DATE BETWEEN '" + v3 + "' AND '" + v4 +
                        "' AND C.COURSE_NATURE LIKE '%" + comboBox3.Text + "%' AND A.VOID LIKE '%" + textBox1.Text +
                        "%' AND B.IF_PAYFOR LIKE '%" + v6 + "%' AND E.ENAME LIKE '%" + textBox2.Text + "%'");

                }
                else if (comboBox1.Text == "财务已审核")
                {
                  
                    search_o(vou.getsqlX + " WHERE  "+v7+" AND "+v8+"  AND C.ACNAME LIKE '%" + textBox5.Text +
                        "%' AND B.FINANCIAL_AUDIT_STATUS='Y' AND B.GENERAL_MANAGE_AUDIT_STATUS='N' AND B.VOUCHER_DATE BETWEEN '" + v3 + "' AND '" + v4 +
                        "' AND C.COURSE_NATURE LIKE '%" + comboBox3.Text + "%' AND A.VOID LIKE '%" + textBox1.Text +
                        "%' AND B.IF_PAYFOR LIKE '%" + v6 + "%' AND E.ENAME LIKE '%" + textBox2.Text + "%'");

                }
                else if (comboBox1.Text == "总经理未审核")
                {
                    search_o(vou.getsqlX + " WHERE  "+v7+" AND "+v8+"  AND C.ACNAME LIKE '%" + textBox5.Text +
                      "%' AND B.GENERAL_MANAGE_AUDIT_STATUS='N' AND B.VOUCHER_DATE BETWEEN '" + v3 + "' AND '" + v4 +
                      "' AND C.COURSE_NATURE LIKE '%" + comboBox3.Text + "%' AND A.VOID LIKE '%" + textBox1.Text +
                      "%' AND B.IF_PAYFOR LIKE '%" + v6 + "%' AND E.ENAME LIKE '%" + textBox2.Text + "%'");
                }
                else if (comboBox1.Text == "总经理已审核")
                {
                    search_o(vou.getsqlX + " WHERE  "+v7+" AND "+v8+"  AND C.ACNAME LIKE '%" + textBox5.Text +
                      "%' AND B.GENERAL_MANAGE_AUDIT_STATUS ='Y' AND B.VOUCHER_DATE BETWEEN '" + v3 + "' AND '" + v4 +
                      "' AND C.COURSE_NATURE LIKE '%" + comboBox3.Text + "%' AND A.VOID LIKE '%" + textBox1.Text +
                      "%' AND B.IF_PAYFOR LIKE '%" + v6 + "%' AND E.ENAME LIKE '%" + textBox2.Text + "%'");
                }
             
                else
                {
                   
                    search_o(vou.getsqlX + " WHERE "+v7+" AND "+v8+"  AND C.ACNAME LIKE '%" + textBox5.Text +
                      "%' AND B.VOUCHER_DATE BETWEEN '" + v3 + "' AND '" + v4 +
                      "' AND C.COURSE_NATURE LIKE '%" + comboBox3.Text + "%' AND A.VOID LIKE '%" + textBox1.Text +
                      "%' AND B.IF_PAYFOR LIKE '%" + v6 + "%' AND E.ENAME LIKE '%" + textBox2.Text + "%'");
                }
            }
            else
            {
                 if (comboBox1.Text == "经理未审核")
                {


                    search_o(vou.getsqlX + " WHERE "+v7+" AND "+v8+"  AND C.ACNAME LIKE '%" + textBox5.Text +
                        "%' AND B.MANAGE_AUDIT_STATUS='N'   AND C.COURSE_NATURE LIKE '%" + comboBox3.Text + "%' AND A.VOID LIKE '%" + textBox1.Text +
                        "%' AND B.IF_PAYFOR LIKE '%" + v6 + "%' AND E.ENAME LIKE '%" + textBox2.Text + "%'");

                }
                else if (comboBox1.Text == "经理已审核")
                {


                    search_o(vou.getsqlX + @" WHERE "+v7+" AND "+v8+"  AND C.ACNAME LIKE '%" + textBox5.Text +
                        "%' AND B.MANAGE_AUDIT_STATUS='Y' AND B.FINANCIAL_AUDIT_STATUS='N' AND B.GENERAL_MANAGE_AUDIT_STATUS='N' AND C.COURSE_NATURE LIKE '%" +
                        comboBox3.Text + "%' AND A.VOID LIKE '%" + textBox1.Text + "%' AND B.IF_PAYFOR LIKE '%" + v6 + "%' AND E.ENAME LIKE '%" + textBox2.Text + "%'");

                }
                else if (comboBox1.Text == "财务未审核")
                {

                    search_o(vou.getsqlX + " WHERE "+v7+" AND "+v8+"  AND C.ACNAME LIKE '%" + textBox5.Text +
                        "%' AND B.FINANCIAL_AUDIT_STATUS='N'  AND C.COURSE_NATURE LIKE '%" + comboBox3.Text +
                        "%' AND A.VOID LIKE '%" + textBox1.Text + "%' AND B.IF_PAYFOR LIKE '%" + v6 + "%' AND E.ENAME LIKE '%" + textBox2.Text + "%'");

                }
                else if (comboBox1.Text == "财务已审核")
                {

                    search_o(vou.getsqlX + " WHERE "+v7+" AND "+v8+"  AND C.ACNAME LIKE '%" + textBox5.Text +
                        "%' AND B.FINANCIAL_AUDIT_STATUS='Y' AND B.GENERAL_MANAGE_AUDIT_STATUS='N'  AND C.COURSE_NATURE LIKE '%" + comboBox3.Text +
                        "%' AND A.VOID LIKE '%" + textBox1.Text + "%' AND B.IF_PAYFOR LIKE '%" + v6 + "%' AND E.ENAME LIKE '%" + textBox2.Text + "%'");

                }
                else if (comboBox1.Text == "总经理未审核")
                {
                    search_o(vou.getsqlX + " WHERE "+v7+" AND "+v8+"  AND C.ACNAME LIKE '%" + textBox5.Text +
                      "%' AND B.GENERAL_MANAGE_AUDIT_STATUS='N'  AND C.COURSE_NATURE LIKE '%" + comboBox3.Text +
                      "%' AND A.VOID LIKE '%" + textBox1.Text + "%' AND B.IF_PAYFOR LIKE '%" + v6 + "%' AND E.ENAME LIKE '%" + textBox2.Text + "%'");
                }
                else if (comboBox1.Text == "总经理已审核")
                {
                    search_o(vou.getsqlX + " WHERE "+ v7+" AND C.ACNAME LIKE '%" + textBox5.Text +
                      "%' AND B.GENERAL_MANAGE_AUDIT_STATUS ='Y' AND C.COURSE_NATURE LIKE '%" + comboBox3.Text +
                      "%' AND A.VOID LIKE '%" + textBox1.Text + "%' AND B.IF_PAYFOR LIKE '%" + v6 + "%' AND E.ENAME LIKE '%" + textBox2.Text + "%'");
                }
           
                else
                {
                 
                    search_o(vou.getsqlX + " WHERE "+v7+" AND "+v8+"  AND C.ACNAME LIKE '%" + textBox5.Text +
                      "%'  AND C.COURSE_NATURE LIKE '%" + comboBox3.Text + "%' AND A.VOID LIKE '%" + textBox1.Text +
                      "%' AND B.IF_PAYFOR LIKE '%" + v6 + "%' AND E.ENAME LIKE '%" + textBox2.Text + "%'");
                }


            }

    
        }
        #endregion

        #region search_o()
        public void search_o(string sql)
        {
            string sqlo =" ORDER BY A.VOID ASC";
            string v7 = bc.getOnlyString("SELECT SCOPE FROM SCOPE_OF_AUTHORIZATION WHERE USID='"+LOGIN .USID +"'");
            //string v7 = "Y";
            if (v7 == "Y")
            {
             
                dt = bc.getdt(sql+sqlo);
                
            }
            else if (v7 == "GROUP")
            {

                dt = bc.getdt(sql + @" AND B.MAKERID IN (SELECT EMID FROM USERINFO A WHERE USER_GROUP IN 
 (SELECT USER_GROUP FROM USERINFO WHERE USID='"+LOGIN .USID +"'))"+sqlo );
            }
            else
            {
                dt = bc.getdt(sql + " AND B.MAKERID='"+LOGIN .EMID +"'"+sqlo );

            }
          
            dt = vou.GET_CALCULATE(dt);
         
           
            if (dt.Rows.Count > 0)
            {
                dataGridView1.DataSource = dt;
                decimal d1 = 0;
                decimal d2 = 0;

                string v1 = dt.Compute("SUM(收入金额)","").ToString();
               
                if (!string.IsNullOrEmpty(v1))
                {
                    d1 = decimal.Parse(v1);
                }
                string v2 = dt.Compute("SUM(支出金额)","").ToString();
                if (!string.IsNullOrEmpty(v2))
                {
                    d2 = decimal.Parse(v2);
                }
                textBox3.Text = (d2).ToString();
                textBox4.Text = (d1).ToString();
                textBox6.Text = (d1 - d2).ToString();
               

                dgvStateControl();
            }
            else
            {


                hint.Text = "找不到所要搜索项！";
                dataGridView1.DataSource = dt;
                textBox3.Text = "";
                textBox4.Text = "";
                textBox6.Text = "";
            }
          

        }
        #endregion
        private void btnExit_Click(object sender, EventArgs e)
        {
            this.Close();
        }



        private void button1_Click(object sender, EventArgs e)
        {
            PERIOD period = new PERIOD();
            MessageBox.Show(period.NEXT_FINANCIAL_YEAR + "," + period.NEXT_PERIOD+","+period .NEXT_PERIOD_t );
        }

        private void chk2_CheckedChanged(object sender, EventArgs e)
        {
            if (chk2.Checked)
            {
                dtpStartDate.Enabled = true;
                dtpEndDate.Enabled = true;
            }
            else
            {
                dtpStartDate.Enabled = false;
                dtpEndDate.Enabled = false;

            }
        }

        private void comboBox2_DropDown(object sender, EventArgs e)
        {
            IF_DOUBLE_CLICK = false;
            BASE_INFO.ACCOUNTANT_COURSE FRM = new CSPSS.BASE_INFO.ACCOUNTANT_COURSE();
            FRM.ShowDialog();
            this.comboBox2.IntegralHeight = false;//使组合框不调整大小以显示其所有项
            this.comboBox2.DroppedDown = false;//使组合框不显示其下拉部分
            this.comboBox2.IntegralHeight = true;//恢复默认值
            string v1 = "";
         
            if (IF_DOUBLE_CLICK)
            {
                foreach (DataRow dr in GETDT_INFO.Rows)
                {
                    v1 = v1 + "'" + dr["科目代码"].ToString() + "'" + ",";


                }
                comboBox2.Text = v1.Substring(0, v1.Length - 1);
            }
        }

        private void comboBox5_DropDown(object sender, EventArgs e)
        {
            IF_DOUBLE_CLICK = false;
            BASE_INFO.ABSTRACT FRM = new CSPSS.BASE_INFO.ABSTRACT();
            FRM.ShowDialog();
            this.comboBox5.IntegralHeight = false;//使组合框不调整大小以显示其所有项
            this.comboBox5.DroppedDown = false;//使组合框不显示其下拉部分
            this.comboBox5.IntegralHeight = true;//恢复默认值
            string v1 = "";

            if (IF_DOUBLE_CLICK)
            {
                foreach (DataRow dr in GETDT_INFO.Rows)
                {
                    v1 = v1 + "'" + dr["摘要代码"].ToString() + "'" + ",";


                }
                comboBox5.Text = v1.Substring(0, v1.Length - 1);
            }
        }

        private void btnDel_Click(object sender, EventArgs e)
        {
            try
            {
                if (MessageBox.Show("确定要删除选中凭证吗？", "提示", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) == DialogResult.OK)
                {
                    for (int i = 0; i < dataGridView1.Rows.Count; i++)
                    {
                        if (dataGridView1.Rows[i].Cells[0].EditedFormattedValue.ToString() == "True")
                        {
                            basec.getcoms("DELETE VOUCHER_MST WHERE VOID='" + dt.Rows[i]["凭证号"].ToString() + "'");
                            basec.getcoms("DELETE VOUCHER_DET WHERE VOID='" + dt.Rows[i]["凭证号"].ToString() + "'");
                         
                        }

                    }
                    search();
                  

                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "提示", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);

            }
       
        
      
        
          
        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            //MessageBox.Show(dt.Rows[dataGridView1.CurrentCell.RowIndex][1].ToString() + "," + dt.Rows[dataGridView1.CurrentCell.RowIndex][3].ToString());
        }

        private void lkgeneral_manage_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
           
            try
            {

            }
            catch (Exception)
            {


            }

            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
                if (dataGridView1.Rows[i].Cells[0].EditedFormattedValue.ToString() == "True")
                {

                    if (vou.RETURN_GENERAL_AUDIT_STATUS(dt.Rows[i]["凭证号"].ToString()) == "N")
                    {
                        basec.getcoms(@"
UPDATE
VOUCHER_MST 
SET 
GENERAL_MANAGE_AUDIT_STATUS='Y',
GENERAL_MANAGE_AUDIT_DATE='" + DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss") +
                         "' WHERE VOID='" + dt.Rows[i]["凭证号"].ToString() + "'");
                    
                     
                    }

                }
            }

      
        }

     
    }
}

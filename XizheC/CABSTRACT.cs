using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Globalization;
using System.Data.SqlClient;
using System.IO;
using System.Data.OleDb;
using XizheC;

namespace XizheC
{
    public class CABSTRACT
    {

        private string _getsql;
        public string getsql
        {
            set { _getsql = value; }
            get { return _getsql; ; }

        }
 
        private bool _IFExecutionSUCCESS;
        public bool IFExecution_SUCCESS
        {
            set { _IFExecutionSUCCESS = value; }
            get { return _IFExecutionSUCCESS; }

        }
        private string _ABID;
        public string ABID
        {
            set { _ABID = value; }
            get { return _ABID; }
        }
        private string _ABCODE;
        public string ABCODE
        {
            set { _ABCODE = value; }
            get { return _ABCODE; }


        }
        private string _EMID;
        public  string EMID
        {
            set { _EMID = value; }
            get { return _EMID; }

        }
  
        private bool _IFCONSULENZA;
        public bool IFCONSULENZA
        {
            set { _IFCONSULENZA = value; }
            get { return _IFCONSULENZA; }
        }
        private string _hint;
        public string hint
        {
            set { _hint = value; }
            get { return _hint; }
        }
        private string _ADD_OR_UPDATE;
        public string ADD_OR_UPDATE
        {
            set { _ADD_OR_UPDATE = value; }
            get { return _ADD_OR_UPDATE; }
        }
        private string _PARENT_NODEID;
        public string PARENT_NODEID
        {
            set { _PARENT_NODEID = value; }
            get { return _PARENT_NODEID; }
        }
        string sql = @"
SELECT A.ABID AS ABID,
A.ABCODE AS ABCODE,
A.ABSTRACT AS ABSTRACT,
(SELECT ENAME FROM EMPLOYEEINFO WHERE EMID=A.MAKERID ) AS MAKER,
A.DATE AS  DATE,
A.PARENT_NODEID AS PARENT_NODEID
FROM ABSTRACT A
";
        string sql1 = @"INSERT INTO ABSTRACT(
ABID,
ABCODE,
ABSTRACT,
MAKERID,
DATE,
YEAR,
MONTH,
PARENT_NODEID
) 
VALUES 
(
@ABID,
@ABCODE,
@ABSTRACT,
@MAKERID,
@DATE,
@YEAR,
@MONTH,
@PARENT_NODEID
)

";
        string sql2 = @"UPDATE ABSTRACT SET 
ABID=@ABID,
ABCODE=@ABCODE,
ABSTRACT=@ABSTRACT,
MAKERID=@MAKERID,
DATE=@DATE,
YEAR=@YEAR,
MONTH=@MONTH,
PARENT_NODEID=@PARENT_NODEID
";
        basec bc = new basec();
        DataTable dt = new DataTable();
        DataTable dto = new DataTable();
        public CABSTRACT()
        {
            IFExecution_SUCCESS = true;
            getsql = sql;
          

        }
        public string GETID()
        {
            string v1 = bc.numYM(10, 4, "0001", "SELECT * FROM ABSTRACT", "ABID", "AB");
            string GETID = "";
            if (v1 != "Exceed Limited")
            {
                GETID = v1;
            }
            return GETID;
        }
    
        #region save
        public void save(string ABID, string ABCODE, string ABSTRACT, string COURSE_TYPE, string BALANCE_DIRECTION,string CYCODE,string COURSE_NATURE)
        {
            string year = DateTime.Now.ToString("yy");
            string month = DateTime.Now.ToString("MM");
            string day = DateTime.Now.ToString("dd");
            string varDate = DateTime.Now.ToString("yyy/MM/dd HH:mm:ss").Replace("-", "/");
            string v1 = bc.getOnlyString("SELECT ABCODE FROM ABSTRACT WHERE  ABID='" + ABID + "'");
            string v2 = bc.getOnlyString("SELECT ABSTRACT FROM ABSTRACT WHERE  ABID='" + ABID + "'");
            string v3 = PARENT_NODEID;
            //string varMakerID;
            if (!bc.exists("SELECT ABID FROM ABSTRACT WHERE ABID='" + ABID + "'"))
            {
                if (bc.exists("SELECT * FROM ABSTRACT WHERE ABCODE='" + ABCODE + "'"))
                {
                    IFExecution_SUCCESS = false;
                   
                    hint = "摘要代码已经存在于系统！";

                }
                else if (bc.exists("SELECT * FROM ABSTRACT WHERE  ABSTRACT='" + ABSTRACT + "'"))
                {
                    IFExecution_SUCCESS = false;
                    hint = "摘要名称已经存在于系统！";

                }
                else
                {
                    IFExecution_SUCCESS = true;

                    SQlcommandE(sql1, ABID, ABCODE, ABSTRACT, COURSE_TYPE, BALANCE_DIRECTION, v3, CYCODE,COURSE_NATURE );
                    ADD_OR_UPDATE = "ADD";
                }

            }
        
            else if (v1 != ABCODE && v2 == ABSTRACT)
            {
                if (bc.exists("SELECT * FROM ABSTRACT WHERE ABCODE='" + ABCODE + "'"))
                {
                    IFExecution_SUCCESS = false;
                    hint = "摘要代码已经存在于系统！";

                }
                else
                {
                    IFExecution_SUCCESS = true;
                    SQlcommandE(sql2 + " WHERE ABID='" + ABID + "'", ABID, ABCODE, ABSTRACT, COURSE_TYPE, BALANCE_DIRECTION, v3, CYCODE, COURSE_NATURE);
                    ADD_OR_UPDATE = "UPDATE";

                }
            }
            else if (v1 == ABCODE && v2 != ABSTRACT)
            {
                if (bc.exists("SELECT * FROM ABSTRACT WHERE ABSTRACT='" + ABSTRACT + "'"))
                {
                    IFExecution_SUCCESS = false;
                    hint = "摘要名称已经存在于系统！";

                }
                else
                {
                    IFExecution_SUCCESS = true;
                    SQlcommandE(sql2 + " WHERE ABID='" + ABID + "'", ABID, ABCODE, ABSTRACT, COURSE_TYPE, BALANCE_DIRECTION, v3, CYCODE, COURSE_NATURE);
                    ADD_OR_UPDATE = "UPDATE";

                }
            }
            else if (v1 != ABCODE && v2 != ABSTRACT)
            {
                if (bc.exists("SELECT * FROM ABSTRACT WHERE ABCODE='" + ABCODE + "'"))
                {
                    IFExecution_SUCCESS = false;
                    hint = "摘要代码已经存在于系统！";

                }
                else if (bc.exists("SELECT * FROM ABSTRACT WHERE  ABSTRACT='" + ABSTRACT + "'"))
                {
                    IFExecution_SUCCESS = false;
                    hint = "摘要名称已经存在于系统！";

                }
                else
                {
                    IFExecution_SUCCESS = true;
                    SQlcommandE(sql2 + " WHERE ABID='" + ABID + "'", ABID, ABCODE, ABSTRACT, COURSE_TYPE, BALANCE_DIRECTION, v3, CYCODE, COURSE_NATURE);
                    ADD_OR_UPDATE = "UPDATE";

                }
            }
            else
            {
                IFExecution_SUCCESS = true;
                SQlcommandE(sql2 + " WHERE ABID='" + ABID + "'", ABID, ABCODE, ABSTRACT, COURSE_TYPE, BALANCE_DIRECTION, v3, CYCODE, COURSE_NATURE);
                ADD_OR_UPDATE = "UPDATE";


            }
        }
        #endregion
        #region SQlcommandE
        protected void SQlcommandE(string sql, string v1, string v2, string v3, string v4, string v5, string v6, string v7, string v8)
        {
            string year = DateTime.Now.ToString("yy");
            string month = DateTime.Now.ToString("MM");
            string day = DateTime.Now.ToString("dd");
            string varDate = DateTime.Now.ToString("yyy/MM/dd HH:mm:ss").Replace("-", "/");
            //string varMakerID = bc.getOnlyString("SELECT EMID FROM USERINFO WHERE USID='" + n2 + "'");
            string varMakerID = EMID;
            SqlConnection sqlcon = bc.getcon();
            SqlCommand sqlcom = new SqlCommand(sql, sqlcon);
            sqlcom.Parameters.Add("@ABID", SqlDbType.VarChar, 20).Value = v1;
            if (v2 == "")
            {
                sqlcom.Parameters.Add("@ABCODE", SqlDbType.VarChar, 20).Value = v1;
            }
            else
            {
                sqlcom.Parameters.Add("@ABCODE", SqlDbType.VarChar, 20).Value = v2;
            }
            sqlcom.Parameters.Add("@ABSTRACT", SqlDbType.VarChar, 20).Value = v3;
            sqlcom.Parameters.Add("@MAKERID", SqlDbType.VarChar, 20).Value = varMakerID;
            sqlcom.Parameters.Add("@DATE", SqlDbType.VarChar, 20).Value = varDate;
            sqlcom.Parameters.Add("@YEAR", SqlDbType.VarChar, 20).Value = year;
            sqlcom.Parameters.Add("@MONTH", SqlDbType.VarChar, 20).Value = month;

            if (!string.IsNullOrEmpty(v6))
            {
                sqlcom.Parameters.Add("@PARENT_NODEID", SqlDbType.VarChar, 20).Value = PARENT_NODEID;
            }
            else
            {
                sqlcom.Parameters.Add("@PARENT_NODEID", SqlDbType.VarChar, 20).Value = DBNull.Value;
            }


            sqlcon.Open();
            sqlcom.ExecuteNonQuery();
            sqlcon.Close();
        }
        #endregion
    }
}

using System;
using System.Collections.Generic;
using System.Text;
using System.Data.OleDb;
using System.Data;
using System.Windows.Forms;

namespace myStatisticalAnalysis
{
    class ReadMDB
    {
        public object MDBreader(string sqlstr, string filepath = null, string passWord = null)
        {

            if (filepath == null)
                filepath = Application.StartupPath + "\\SPC_DATA.accdb";
            //   string tablename = (string)argsList[3];
            OleDbDataAdapter myadapter;
            // string strcon = sqlstr;
            if (passWord == "" || passWord == null)
            {
                myadapter = new OleDbDataAdapter(sqlstr, InitialConnection(filepath));
            }
            else
            {
                myadapter = new OleDbDataAdapter(sqlstr, InitialConnection(filepath, passWord));
            }
            
            DataTable DT1 = new DataTable();
            //  DataSet DS1 = new DataSet();
            //   myadapter.Fill(DS1,"nihao");
            myadapter.Fill(DT1);
            //myadapter.Fill(
            return DT1;

        }
        public void InsertmdbNoRepFast(DataTable DT, string TableName)
        {
            //  ReadMDB readmdb1 = new ReadMDB();

            if (IsTableExist("testdata000"))
                MDBreader(" drop table testdata000 ");
            MDBreader(" select *  into testdata000   from  testdata where 1=0");
            SaveDataTable("testdata000", DT); ;
            MDBreader("Insert into testdata  select * from testdata000  where not exists (select * from testdata where testdata.数据=testdata000.数据 and testdata.日期时间=testdata000.日期时间)");
           

        }

        public void InsertmdbNoRep(DataTable DT, string TableName, string[] colName = null)
        {
       //     ReadMDB readmdb1 = new ReadMDB();
            if (colName == null)
            {
                colName = GetColumnsName(TableName);
            }

            foreach (DataRow DR in DT.Rows)
            {
                string s0 = "DELETE * FROM " + TableName + "  WHERE  ";
                string s1 = "insert into " + TableName + "  values(  ";
                for (int i = 0; i < DT.Columns.Count; i++)
                {
                    if (i == 0)
                    {
                        s0 += (colName[i] + " = " + DR[i].ToString());
                        s1 += DR[i].ToString();
                    }
                    else
                    {
                        s0 += (colName[i] + " = '" + DR[i].ToString() + "'");
                        s1 += (" '" + DR[i].ToString() + "'");
                    }
                    if (i != DT.Columns.Count - 1)
                    {
                        s0 += " and ";
                        s1 += ",";
                    }
                    else
                    {
                        s1 += ")";
                    }
                }
                MDBreader(s0);
                MDBreader(s1);
            }
        }


        public bool IsTableExist(string TableName, string filepath = null)
        {
            if (filepath == null)
            {
                filepath = Application.StartupPath + "\\SPC_DATA.accdb"; 
            }
            OleDbConnection conn = InitialConnection(filepath);
         //   conn.GetOleDbSchemaTable(
            conn.Open();
            DataTable schemaTable = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables,
                                                             new object[] { null, null, null, "TABLE" });
            conn.Close();

            if (schemaTable != null)
            {
                for (Int32 row = 0; row < schemaTable.Rows.Count; row++)
                {
                    string col_name = schemaTable.Rows[row]["TABLE_NAME"].ToString();
                    if (col_name == TableName)
                    {
                        return true;
                    }
                }
            }

            return false;
 
        }

        public bool SaveDataTable(string TableName, DataTable DT,string filepath = null)
        {
            if (filepath == null)
            {
                filepath = Application.StartupPath + "\\SPC_DATA.accdb";
            }
       

           
            OleDbDataAdapter myadapter;
            // string strcon = sqlstr;

            myadapter = new OleDbDataAdapter("select * from " + TableName, InitialConnection(filepath));

            OleDbCommandBuilder cb = new OleDbCommandBuilder(myadapter);
            //myadapter.Fill(DT);
            myadapter.Update(DT);
            
            //myadapter.Fill(
            return true;

       

        }

        public string[] GetColumnsName(string TableName, string filepath = null)
        {
           
            if (filepath == null)
            {
                filepath = Application.StartupPath + "\\SPC_DATA.accdb";
            }
            OleDbConnection conn = InitialConnection(filepath);
            
            conn.Open();
            DataTable columnTable = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Columns, new object[] { null, null, TableName, null });
            conn.Close();
            conn.Dispose();
            string[] ColumnsName = new string[columnTable.Rows.Count];

            foreach (DataRow dr2 in columnTable.Rows)
            {
               
                ColumnsName[(long)dr2["ORDINAL_POSITION"] - 1] = dr2["COLUMN_NAME"].ToString();
             
            }


            return ColumnsName;
          

        }

        public List<string> GetColumnsName1(string TableName, string filepath = null)
        {

            if (filepath == null)
            {
                filepath = Application.StartupPath + "\\SPC_DATA.accdb";
            }
            OleDbConnection conn = InitialConnection(filepath);

            conn.Open();
            DataTable columnTable = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Columns, new object[] { null, null, TableName, null });
            conn.Close();

            List<string> ColumnsName = new List<string>();

          //  ColumnsName[]
            foreach (DataRow dr2 in columnTable.Rows)
            {

                ColumnsName[(int)((long)dr2["ORDINAL_POSITION"] - 1)] = dr2["COLUMN_NAME"].ToString();

            }


            return ColumnsName;


        }

        public OleDbConnection InitialConnection(string filepath)
        {
            // string constr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + filepath;
            // string constr = "Provider=Microsoft.Jet.OleDb.4.0;Data Source=" + filepath;
            string password = "31204000498";
            string constr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + filepath + ";Jet OLEDB:Database Password=" + password + ";";
         
            OleDbConnection connection = new OleDbConnection(constr);
            return connection;
        }
        public OleDbConnection InitialConnection(string filepath, string password)
        {
            string constr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + filepath + ";Jet OLEDB:Database Password=" + password + ";";
            OleDbConnection connection = new OleDbConnection(constr);
            return connection;

        }
    }
}

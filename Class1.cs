using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace chungju
{
    internal class Class1
    {
        public static OleDbConnection con = new OleDbConnection();


        public bool ExcelConnection(string strName)
        {
            string fileName = strName;

            try
            {
                bool hasHeaders = true;
                string HDR = hasHeaders ? "Yes" : "No";

                if (fileName.Substring(fileName.LastIndexOf('.')).ToLower() == ".xlsx")
                    con.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + fileName + ";Extended Properties=\"Excel 12.0;HDR=" + HDR + ";\"";
                else
                    con.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + fileName + ";Extended Properties=\"Excel 8.0;HDR=" + HDR + ";\"";
                con.Open();
            }
            catch (Exception ex)
            {
                return false;
            }

            return true;
        }

        public void ExcelClose()
        {
            if(con.State == System.Data.ConnectionState.Open)
            {
                con.Close();
            }
        }

        public System.Data.DataTable XlsDataTable(string query, System.Data.DataTable dataTable, DataSet dataSet, string sheet)
        {
            OleDbCommand cmd = new OleDbCommand(query, con);
            cmd.CommandType = CommandType.Text;

            dataTable = new System.Data.DataTable(sheet);
            dataSet.Tables.Add(dataTable);

            new OleDbDataAdapter(cmd).Fill(dataTable);

            return dataTable;
        }

        public System.Data.DataTable XlsDataTable(string query, System.Data.DataTable dataTable, DataSet dataSet, string sheet, int ConGubun)
        {
            OleDbCommand cmd = new OleDbCommand(query, con);
            cmd.CommandType = CommandType.Text;

            if (dataSet == null)
            {
                dataSet = new DataSet();
            }

            dataTable = new System.Data.DataTable(sheet);
            dataSet.Tables.Add(dataTable);

            new OleDbDataAdapter(cmd).Fill(dataTable);

            return dataTable;
        }
    }
}

using DevExpress.ClipboardSource.SpreadsheetML;
using DevExpress.Spreadsheet;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Database_To_Excel_Schema_Converter
{
    public partial class Form1 : DevExpress.XtraBars.Ribbon.RibbonForm
    {
        private IWorkbook workbook;

        public string ConString { get; set; }
        public Form1()
        {
            InitializeComponent();
            ConString = "Data Source=THUNIS-FLIGHT;Initial Catalog=DataBRI;Integrated Security=True;Asynchronous Processing=true;";
        }

        private void CreateXLSSheetAndPopulateColumnData(string TableName)
        {
            String[] tableRestrictions = new String[4];

            tableRestrictions[2] = TableName;
            using (SqlConnection conn = new SqlConnection(ConString))
            {
                conn.Open();
                DataTable metaDataTable = conn.GetSchema("Columns", tableRestrictions);
                metaDataTable.DefaultView.Sort="ORDINAL_POSITION";

                DataTable dtSort = metaDataTable.DefaultView.ToTable();

                workbook = xls.Document;
                workbook.Worksheets.Add().Name = metaDataTable.Rows[0]["TABLE_NAME"].ToString();
                DevExpress.Spreadsheet.Worksheet ws = workbook.Worksheets[metaDataTable.Rows[0]["TABLE_NAME"].ToString()];
                ws.Cells[0, 0].Value = "Column Name";
                ws.Cells[0,0].ColumnWidth= 800;
                ws.Cells[0, 1].Value = "Data Type";
                ws.Cells[0, 1].ColumnWidth = 400;

                for (int i = 0; i < dtSort.Rows.Count; i++)
                {
                    ws.Cells[i+1, 0].Value = dtSort.Rows[i]["COLUMN_NAME"].ToString();
                    ws.Cells[i + 1, 1].Value = dtSort.Rows[i]["DATA_TYPE"].ToString();

                }


            }
        }

        private DataTable GetSchema()
        {
            using (SqlConnection conn = new SqlConnection(ConString))
            {
                conn.Open();
                DataTable metaDataTable = conn.GetSchema("Tables");
                return metaDataTable;
            }

        }

        private void barButCreateXLS_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            DataTable dt = GetSchema();
            foreach (DataRow dr in dt.Rows)
            {
                if (dr["TABLE_TYPE"].ToString() == "BASE TABLE")
                {
                    CreateXLSSheetAndPopulateColumnData(dr["TABLE_NAME"].ToString());
                }
            }

        }
    }
}

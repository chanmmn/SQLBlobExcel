using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConAppReadBlob
{
    class Program
    {
        static void Main(string[] args)
        {
            ReadBBlob();
            //ExportExcelFromDB();
        }

        public static void ReadBBlob()
        {
            SqlConnection con = new SqlConnection("Data Source =172.18.123.74; Initial Catalog = ispec.ptp.com.my; Integrated Security = True");

            SqlDataAdapter da = new SqlDataAdapter("Select * From ATTACHMENT WHERE [ATTACHFILE] IS NOT NULL AND[ATTACHMENT_ID] = 345", con);
            SqlCommandBuilder MyCB = new SqlCommandBuilder(da);
            DataSet ds = new DataSet("MyImages");

            byte[] MyData = new byte[0];
            da.Fill(ds, "MyImages");
            DataRow myRow;
            myRow = ds.Tables["MyImages"].Rows[0];

            //MyData = (byte[])myRow["imgField"];
            MyData = (byte[])myRow["ATTACHFILE"];
            int ArraySize = new int();
            ArraySize = MyData.GetUpperBound(0);
                       
            FileStream fs = new FileStream(@"C:\temp\Spare Part Data Requirements.xlsx", FileMode.OpenOrCreate, FileAccess.Write);
            fs.Write(MyData, 0, ArraySize);
            fs.Close();
        }

        public static void ExportExcelFromDB()
        {
            //string filepathtostore = @"D:\TPMS\Uploaded_Boq\Raveena_boq_From_Db.xlsx";
            string filepathtostore = @"C:\temp\QIFFARAH_VENTURES_SB.xlsx";
            RetrieveExcelFileFromDatabase(12096, filepathtostore);
        }

        public static void RetrieveExcelFileFromDatabase(int ID, string excelFileName)
        {
            byte[] excelContents;

            //string selectStmt = "SELECT FileContent FROM dbo.Tender_Excel_Source WHERE file_sequence_no = @ID";
            string selectStmt = "Select [ATTACHFILE] From ATTACHMENT WHERE [ATTACHFILE] IS NOT NULL AND[ATTACHMENT_ID] = @ID";

            using (SqlConnection connection = new SqlConnection("Data Source =172.18.123.74; Initial Catalog = ispec.ptp.com.my; Integrated Security = True"))
            using (SqlCommand cmdSelect = new SqlCommand(selectStmt, connection))
            {
                cmdSelect.Parameters.Add("@ID", SqlDbType.Int).Value = ID;

                connection.Open();
                excelContents = (byte[])cmdSelect.ExecuteScalar();
                connection.Close();
            }

            File.WriteAllBytes(excelFileName, excelContents);
        }
    }
}

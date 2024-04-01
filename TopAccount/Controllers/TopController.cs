using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using VBIDE = Microsoft.Vbe.Interop;
using Excel = Microsoft.Office.Interop.Excel;
using ClosedXML.Excel;
using TopAccount.Models;

namespace TopAccount.Controllers
{
    public class TopController : Controller
    {
        public SqlConnection con = new SqlConnection(ConfigurationManager.ConnectionStrings["ConnectionString"].ConnectionString);
        public ActionResult Index()
        {
            Consolidated _model = Bind();
            ViewBag.Display = "none";
            return View(_model);
        }

        private Consolidated Bind()
        {
            Parameter model = new Parameter();
            DataTable dt = new DataTable();//GetSL();
            dt.Columns.Add("Role");

            string[] _SL = new string[] {"EAS","SAP","ORC","ECAS","EAIS" };

            for(int i = 0; i < _SL.Length; i++)
            {
                DataRow dr = dt.NewRow();
                dr["Role"] = _SL[i].ToString();
                    
                dt.Rows.Add(dr);
            }

            model.SLId = "EAS";


            //if (dt.Rows.Count == 1) { model.SLId = dt.Rows[0]["Role"].ToString(); }
            //else { model.SLId = "EAS"; }

            for (int i = 0; i < dt.Rows.Count; i++)
            {
                string Role = dt.Rows[i]["Role"].ToString();
                model.SL.Add(new SelectListItem { Text = Role, Value = Role });
            }

              

            Consolidated _model = new Consolidated();
            _model.item2 = model;
            return _model;
        }

        [HttpPost]
        public ActionResult Index(Consolidated model, string cookieValue)
        {
            ViewBag.Display = "block";
            string SL = model.item2.SLId;
                     
           
                
                ControllerContext.HttpContext.Response.Cookies.Add(new HttpCookie("dlc", cookieValue));
                return getFinal(SL);
            

            //return GreaterTo(SL, Finyear, from, to, From, To);

        }

        private ActionResult getFinal(string SL)
        {



            DataSet ds = GetData(SL);
            DataTable dt_Account_Wise_Revenue = ds.Tables[0];
            DataTable dt_Org_100 = ds.Tables[1];

            DataTable dt = GetData_Top15(SL);

            string templatePath = string.Empty;

            templatePath = Server.MapPath("~/Template/"+SL+"_TopAccounts_OrgClassification.xlsx");

            var wb = new XLWorkbook(templatePath);
            IXLWorksheet Worksheet = wb.Worksheet("AccountWise_Revenue&TopAccounts");
            Worksheet.Row(19).Cell(1).InsertData(dt_Account_Wise_Revenue.Rows);

            IXLWorksheet Worksheet1 = wb.Worksheet("Org 300 Accounts_Charts");
            Worksheet1.Row(3).Cell(2).InsertData(dt_Org_100.Rows);

            IXLWorksheet Worksheet2 = wb.Worksheet("Top 15 Accounts");
            Worksheet2.Row(3).Cell(1).InsertData(dt.Rows);

            string filename = "";

            //string month_diff=""
            //if(From==To)
            //{
            //    string k = From =="04" ? "Apr"
            //     month_diff = From;
            //}


            string snapshot = DateTime.Now.ToString("ddMMMyyyy_HHmm");
            filename = SL + "_" + GETUser() + "_TopAccounts_OrgClassification_" + snapshot + ".xlsx";
            wb.SaveAs(Server.MapPath("~/ExcelOperation/" + filename));

            //GenerateReport(filename);

            string fullPath = Path.Combine(Server.MapPath("~/ExcelOperation"), filename);
            return File(fullPath, "application/vnd.ms-excel", filename);
        }

        private static string GETUser()
        {
            string user = "";

            string[] machineUser = System.Web.HttpContext.Current.User.Identity.Name.Split('\\');
            if (machineUser.Length == 2)
                user = machineUser[1];
            return user;
        }

        private DataSet GetData(string SL)
        {

            string SP = "";
            if (SL == "EAS") { SP = "sp_SL_Data_OrgClassification_EAS_Online"; } else { SP = "sp_SL_Data_OrgClassification_Online"; }
            SqlCommand cmd = new SqlCommand(SP, con);
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.CommandTimeout = int.MaxValue;
            //cmd.Parameters.AddWithValue("@UserId", GETUser());
            if (SL != "EAS") { cmd.Parameters.AddWithValue("@SL", SL); }
              
            SqlDataAdapter sdr = new SqlDataAdapter(cmd);
            DataSet ds = new DataSet();
            sdr.Fill(ds);
            return ds;
        }

        private DataTable GetData_Top15(string SL)
        {

            string SP = "";
            SP = "SP_Dinesh_DP_Top15Accounts"; 
            SqlCommand cmd = new SqlCommand(SP, con);
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.CommandTimeout = int.MaxValue;
            cmd.Parameters.AddWithValue("@ParamSL", SL);
            SqlDataAdapter sdr = new SqlDataAdapter(cmd);
            DataSet ds = new DataSet();
            sdr.Fill(ds);
            return ds.Tables[0];
        }

        private DataTable GetSL()
        {
            SqlCommand cmd = new SqlCommand("spBEGetSU_dummy", con);
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.CommandTimeout = int.MaxValue;
            cmd.Parameters.AddWithValue("@userid", GETUser());

            SqlDataAdapter sdr = new SqlDataAdapter(cmd);
            DataSet ds = new DataSet();
            sdr.Fill(ds);
            return ds.Tables[0];
        }

      



        void GenerateReport(string fname)
        {


            Microsoft.Office.Interop.Excel.Application oExcel;
            Microsoft.Office.Interop.Excel.Workbook oBook = default(Microsoft.Office.Interop.Excel.Workbook);
            VBIDE.VBComponent oModule;
            //try
            {

                string folder = "~/ExcelOperation";
                var myDir = new DirectoryInfo(Server.MapPath(folder));

                string templatefolder = "~/template";
                var templatemyDir = new DirectoryInfo(Server.MapPath(templatefolder));

                String sCode;
                Object oMissing = System.Reflection.Missing.Value;
                //instance of excel
                oExcel = new Microsoft.Office.Interop.Excel.Application();

                oBook = oExcel.Workbooks.
                    Open(myDir.FullName + "\\" + fname + "", 0, false, 5, "", "", true,
                    Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                Microsoft.Office.Interop.Excel.Sheets WRss = oBook.Sheets;

                //oModule = oBook.VBProject.VBComponents.Add(VBIDE.vbext_ComponentType.vbext_ct_StdModule);

                //sCode = "sub Macro()\r\n" +

                //    System.IO.File.ReadAllText(templatemyDir.FullName + "\\PerformanceDeck_macro.txt") +
                //        "\nend sub";
                //oModule.CodeModule.AddFromString(sCode);

                //oExcel.GetType().InvokeMember("Run",
                //                System.Reflection.BindingFlags.Default |
                //                System.Reflection.BindingFlags.InvokeMethod,
                //                null, oExcel, new string[] { "Macro" });



                /////////////////////////////////////
                oBook.RefreshAll();
                oBook.Save();
                oBook.Close(false, myDir.FullName + "\\" + fname + "", null);


                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(oBook);
                oBook = null;

                oExcel.Quit();
                oExcel = null;



                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
                GC.WaitForPendingFinalizers();



            }

        }

        public void ReleaseObject(object o)
        {
            try
            {
                if (o != null)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(o);
            }
            catch (Exception) { }
            finally { o = null; }
        }
        private string GetUserID()
        {
            string[] machineUsers = User.Identity.Name.Split('\\');
            if (machineUsers.Length == 2)
                return machineUsers[1];
            return "";
        }

        public DataSet Get_Data(string UserId, string SL, string Finyear, string duration, string Period, string Type)
        {

            DataSet ds = new DataSet();
            SqlCommand cmd = new SqlCommand("sp_Project-wise_Contract-wise_Summary", con);
            cmd.CommandTimeout = int.MaxValue;
            cmd.Parameters.AddWithValue("@UserId", UserId);
            cmd.Parameters.AddWithValue("@SU", SL);
            cmd.Parameters.AddWithValue("@fyr", Finyear);
            cmd.Parameters.AddWithValue("@duration", duration);
            cmd.Parameters.AddWithValue("@Period", Period);
            cmd.Parameters.AddWithValue("@Type", Type);
            cmd.CommandType = CommandType.StoredProcedure;
            SqlDataAdapter sdr = new SqlDataAdapter(cmd);
            sdr.Fill(ds);
            return ds;
        }

        public ActionResult About()
        {
            ViewBag.Message = "Your application description page.";

            return View();
        }

        public ActionResult Contact()
        {
            ViewBag.Message = "Your contact page.";

            return View();
        }
    }
}
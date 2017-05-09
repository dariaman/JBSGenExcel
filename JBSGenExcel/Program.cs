using MySql.Data.MySqlClient;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace JBSGenExcel
{
    class Program
    {
        static string constring = ConfigurationManager.AppSettings["DefaultDB"];        
        static string DirBilling = ConfigurationManager.AppSettings["FileBilling"];
        static string TempMandiriCCFile = ConfigurationManager.AppSettings["TemplateMandiriCC"];
        static string TempBCAacFile = ConfigurationManager.AppSettings["TemplateBCAac"];
        static string MandiriccFile = "Mandiri_" + DateTime.Now.ToString("ddMMyyyy") + ".xls";
        static string BCAacFile = "BCAac" + DateTime.Now.ToString("yyyyMMdd") + ".xls";
        static string VaRegulerPremi = "VARegulerPremi" + DateTime.Now.ToString("yyyyMMdd") + ".xls";
        static void Main(string[] args)
        {
            if (args[0] == "mandiricc")
            {
                FileInfo FileName = new FileInfo(DirBilling + MandiriccFile);
                if (!FileName.Exists)
                {
                    genMandiriCC();
                }
            }
            else if (args[0] == "bcaac")
            {
                FileInfo FileName = new FileInfo(DirBilling + BCAacFile);
                if (!FileName.Exists)
                {
                    genBCAac();
                }
            }
            else if (args[0] == "va")
            {
                FileInfo FileName = new FileInfo(DirBilling + VaRegulerPremi);
                if (FileName.Exists) { FileName.Delete(); }
                genVARegulerPremi();
            }
        }
        public static void genMandiriCC()
        {
            MySqlConnection con = new MySqlConnection(constring);
            HSSFWorkbook hssfwb;
            MySqlCommand cmd;
            FileInfo FileName;
            try
            {
                FileName = new FileInfo(TempMandiriCCFile);
                if (FileName.Exists)
                {
                    FileName = new FileInfo(DirBilling + MandiriccFile);
                }
                else
                {
                    Exception ex = new Exception("File template tidak ditemukan");
                    throw ex;
                }
                cmd = new MySqlCommand("BillingMandiriCC_sp", con);
                cmd.CommandType = System.Data.CommandType.StoredProcedure;

            }
            catch (Exception ex)
            {
                throw ex;
            }

            using (FileStream file = new FileStream(TempMandiriCCFile, FileMode.Open, FileAccess.Read))
            {
                hssfwb = new HSSFWorkbook(file);
            }

            using (FileStream file = new FileStream(FileName.FullName.ToString(), FileMode.Create, FileAccess.ReadWrite))
            {
                con.Open();
                using (MySqlDataReader reader = cmd.ExecuteReader())
                {
                    try
                    {
                        ISheet sheet = hssfwb.GetSheet("sheet1");
                        int j = 1;
                        int i = 15;
                        while (reader.Read())
                        {
                            IRow row = sheet.GetRow(i);
                            row.GetCell(0).SetCellValue(j);
                            row.GetCell(2).SetCellValue(reader["a"].ToString());
                            row.GetCell(4).SetCellValue(reader["b"].ToString());
                            row.GetCell(6).SetCellValue(reader["c"].ToString());
                            row.GetCell(8).SetCellValue(reader["d"].ToString());
                            row.GetCell(10).SetCellValue(reader["e"].ToString());
                            row.GetCell(12).SetCellValue(reader["f"].ToString());

                            i++;
                            j++;
                        }
                    }
                    catch (Exception ex) { throw ex; }
                    finally
                    {
                        hssfwb.Write(file);
                        file.Close();
                        con.Dispose();
                        con.Close();
                    }
                }
            }

        }

        public static void genBCAac()
        {
            MySqlConnection con = new MySqlConnection(constring);
            HSSFWorkbook hssfwb;
            MySqlCommand cmd;
            FileInfo FileName;
            try
            {
                FileName = new FileInfo(TempBCAacFile);
                if (FileName.Exists)
                {
                    FileName = new FileInfo(DirBilling + BCAacFile);
                }
                else
                {
                    Exception ex = new Exception("File template tidak ditemukan");
                    throw ex;
                }
                cmd = new MySqlCommand("BillingBcaAC_sp", con);
                cmd.CommandType = System.Data.CommandType.StoredProcedure;

            }
            catch (Exception ex)
            {
                throw ex;
            }

            using (FileStream file = new FileStream(TempBCAacFile, FileMode.Open, FileAccess.Read))
            {
                hssfwb = new HSSFWorkbook(file);
            }

            using (FileStream file = new FileStream(FileName.FullName.ToString(), FileMode.Create, FileAccess.ReadWrite))
            {
                con.Open();
                using (MySqlDataReader reader = cmd.ExecuteReader())
                {
                    try
                    {
                        ISheet sheet = hssfwb.GetSheet("sheet1");
                        int i = 1;
                        while (reader.Read())
                        {
                            IRow row;
                            row = sheet.CreateRow(i);
                            row.CreateCell(0).SetCellValue(reader["a"].ToString());
                            row.CreateCell(1).SetCellValue(reader["b"].ToString());
                            row.CreateCell(2).SetCellValue(reader["c"].ToString());
                            row.CreateCell(3).SetCellValue(reader["d"].ToString());
                            row.CreateCell(4).SetCellValue(reader["e"].ToString());
                            row.CreateCell(5).SetCellValue(reader["f"].ToString());
                            row.CreateCell(6).SetCellValue(reader["g"].ToString());
                            row.CreateCell(7).SetCellValue(reader["h"].ToString());
                            row.CreateCell(8).SetCellValue(reader["i"].ToString());
                            row.CreateCell(9).SetCellValue(reader["j"].ToString());


                            i++;
                        }

                    }
                    catch (Exception ex)
                    {
                        throw ex;
                    }
                    finally
                    {
                        hssfwb.Write(file);
                        file.Close();
                        con.Dispose();
                        con.Close();
                    }
                }
            }
        }

        public static void genVARegulerPremi()
        {
            MySqlConnection con = new MySqlConnection(constring);
            MySqlCommand cmd;
            FileInfo FileName;
            HSSFWorkbook hssfwb = new HSSFWorkbook();

            FileName = new FileInfo(DirBilling + VaRegulerPremi);
            cmd = new MySqlCommand("GenVARegulerPremi_sp", con);
            cmd.CommandType = System.Data.CommandType.StoredProcedure;

            using (FileStream file = new FileStream(FileName.FullName.ToString(), FileMode.Create, FileAccess.ReadWrite))
            {
                con.Open();
                using (MySqlDataReader reader = cmd.ExecuteReader())
                {
                    try
                    {
                        ISheet sheet = hssfwb.CreateSheet("sheet1");
                        IRow row;
                        row = sheet.CreateRow(0);
                        row.CreateCell(0).SetCellValue("No Polis");
                        row.CreateCell(1).SetCellValue("Pemegang Polis");
                        row.CreateCell(2).SetCellValue("Premi");

                        int i = 1;
                        while (reader.Read())
                        {                            
                            row = sheet.CreateRow(i);
                            row.CreateCell(0).SetCellValue(reader["a"].ToString());
                            row.CreateCell(1).SetCellValue(reader["b"].ToString());
                            row.CreateCell(2).SetCellValue(reader["c"].ToString());
                            i++;
                        }

                    }
                    catch (Exception ex)
                    {
                        throw ex;
                    }
                    finally
                    {
                        hssfwb.Write(file);
                        file.Close();
                        con.Dispose();
                        con.Close();
                    }
                }
            }
        }
    }
}

using MySql.Data.MySqlClient;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace JBSGenExcel
{
    class Program
    {
        static string constring = ConfigurationManager.AppSettings["DefaultDB"];
        static string con21 = ConfigurationManager.AppSettings["Life21DB"];
        static string con21p = ConfigurationManager.AppSettings["Life21P"];

        static string DirBilling = ConfigurationManager.AppSettings["FileBilling"];
        static string DirResult = ConfigurationManager.AppSettings["DirResult"];

        static string TempMandiriCCFile = ConfigurationManager.AppSettings["TemplateMandiriCC"];
        static string TempBCAacFile = ConfigurationManager.AppSettings["TemplateBCAac"];
        static string MandiriccFile = "Mandiri_" + DateTime.Now.ToString("ddMMyyyy") + ".xls";
        static string BCAacFile = "BCAac" + DateTime.Now.ToString("yyyyMMdd") + ".xls";
        static string VaRegulerPremi = "VARegulerPremi" + DateTime.Now.ToString("yyyyMMdd") + ".xls";

        static void Main(string[] args)
        {
            if (args.Length > 0)
            {
                if (args[0] == "mandiricc")
                {
                    FileInfo FileName = new FileInfo(DirBilling + MandiriccFile);
                    if (!FileName.Exists) genMandiriCC();
                }
                else if (args[0] == "bcaac")
                {
                    FileInfo FileName = new FileInfo(DirBilling + BCAacFile);
                    if (!FileName.Exists) genBCAac();
                }
                else if (args[0] == "va")
                {
                    FileInfo FileName = new FileInfo(DirBilling + VaRegulerPremi);
                    if (FileName.Exists) { FileName.Delete(); }
                    genVARegulerPremi();
                }
                else if (args[0] == "sync")
                {
                    SyncAmountBilling();
                }
            }
            else
            {
                throw new Exception("Parameter kosong");
            }

        }
        public static void genMandiriCC()
        {
            MySqlConnection con = new MySqlConnection(constring);
            HSSFWorkbook hssfwb;
            MySqlCommand cmd;
            MySqlCommand cmd2;
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
            }
            catch (Exception ex)
            {
                throw ex;
            }

            using (FileStream file = new FileStream(TempMandiriCCFile, FileMode.Open, FileAccess.Read))
            {
                hssfwb = new HSSFWorkbook(file);
            }

            //Untuk data
            cmd = new MySqlCommand("BillingMandiriCC_sp", con);
            cmd.CommandType = System.Data.CommandType.StoredProcedure;


            // Untuk Header File Mandiri
            cmd2 = new MySqlCommand(@"SELECT SUM(jlh) AS jlh,SUM(total) AS total
                                        FROM(
	                                        SELECT COUNT(1) AS jlh, SUM(b.`TotalAmount`) AS total
	                                        FROM `billing` b
	                                        WHERE b.`IsDownload`= 1 AND b.`BankIdDownload`= 2
	                                        UNION ALL
	                                        SELECT COUNT(1) AS jlh, SUM(b.`TotalAmount`) AS total
	                                        FROM `billing_others` b
	                                        WHERE b.`IsDownload`= 1 AND b.`BankIdDownload`= 2
	                                        UNION ALL
	                                        SELECT COUNT(1) AS jlh, SUM(b.`TotalAmount`) AS total
	                                        FROM `quote_billing` b
	                                        WHERE b.`IsDownload`= 1 AND b.`BankIdDownload`= 2
                                        )a;	", con);
            cmd2.CommandType = System.Data.CommandType.Text;
            using (FileStream file = new FileStream(FileName.FullName.ToString(), FileMode.Create, FileAccess.ReadWrite))
            {
                con.Open();

                ISheet sheet = hssfwb.GetSheet("sheet1");
                IRow row;

                using (MySqlDataReader reader = cmd.ExecuteReader())
                {
                    try
                    {
                        int j = 1;
                        int i = 15;
                        while (reader.Read())
                        {
                            row = sheet.GetRow(i);
                            if (row == null) row = sheet.CreateRow(i);
                            if (row.GetCell(0) == null) row.CreateCell(0);
                            row.GetCell(0).SetCellValue(j);
                            if (row.GetCell(2) == null) row.CreateCell(2);
                            row.GetCell(2).SetCellValue(reader["a"].ToString());
                            if (row.GetCell(4) == null) row.CreateCell(4);
                            row.GetCell(4).SetCellValue(reader["b"].ToString());
                            if (row.GetCell(6) == null) row.CreateCell(6);
                            row.GetCell(6).SetCellValue(reader["c"].ToString());
                            if (row.GetCell(8) == null) row.CreateCell(8);
                            row.GetCell(8).SetCellValue(reader["d"].ToString());
                            if (row.GetCell(10) == null) row.CreateCell(10);
                            row.GetCell(10).SetCellValue(reader["e"].ToString());
                            if (row.GetCell(12) == null) row.CreateCell(12);
                            row.GetCell(12).SetCellValue(reader["f"].ToString());

                            i++;
                            j++;
                        }
                    }
                    catch (Exception ex) { throw ex; }
                    finally
                    {
                        reader.Close();
                    }
                }

                try
                {
                    using (MySqlDataReader reader2 = cmd2.ExecuteReader())
                    {
                        while (reader2.Read())
                        {
                            row = sheet.GetRow(3);
                            row.GetCell(4).SetCellValue("01010452216");
                            row = sheet.GetRow(5);
                            row.GetCell(4).SetCellValue("ASURANSI JAGA DIRI RECURRING");
                            row = sheet.GetRow(7);
                            row.GetCell(4).SetCellValue(DateTime.Now.ToString("ddMMyyyy"));
                            row = sheet.GetRow(8);
                            row.GetCell(4).SetCellValue("AFI0910121");
                            row = sheet.GetRow(9);
                            row.GetCell(4).SetCellValue("607");
                            row = sheet.GetRow(10);
                            row.GetCell(4).SetCellValue("C");
                            row = sheet.GetRow(11);
                            row.GetCell(4).SetCellValue(reader2["jlh"].ToString());
                            row = sheet.GetRow(12);
                            row.GetCell(4).SetCellValue(reader2["total"].ToString());
                        }
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

        static void SyncAmountBilling()
        {
            MySqlConnection con = new MySqlConnection(constring);
            MySqlConnection cons21 = new MySqlConnection(con21);
            MySqlCommand cmd;
            DataTable dt = new DataTable();

            cmd = new MySqlCommand("GetBillAmount", cons21);
            cmd.CommandType = System.Data.CommandType.StoredProcedure;
            cons21.Open();
            int jlh;
            using (MySqlDataReader dr = cmd.ExecuteReader())
            {
                dt.Load(dr);
                jlh = dt.Rows.Count;
            }
            cons21.Close();

            cmd = new MySqlCommand("DELETE FROM `dumpbilling`", con);
            cmd.CommandType = System.Data.CommandType.Text;
            con.Open();
            cmd.ExecuteNonQuery();
            con.Close();
            int baris = 0;

            try
            {
                while (baris < jlh)
                {
                    StringBuilder sCommand = new StringBuilder("INSERT INTO dumpbilling VALUES ");
                    using (MySqlConnection mConnection = new MySqlConnection(constring))
                    {
                        List<string> Rows = new List<string>();
                        Rows.Clear();
                        for (int i = 0; i < 1000; i++)
                        {
                            Rows.Add(string.Format("('{0}','{1}','{2}','{3}')", dt.Rows[baris][0], dt.Rows[baris][1], dt.Rows[baris][2], dt.Rows[baris][3]));
                            baris++;
                            if (baris == jlh) break;
                        }
                        sCommand.Append(string.Join(",", Rows));
                        sCommand.Append(";");
                        mConnection.Open();
                        using (MySqlCommand myCmd = new MySqlCommand(sCommand.ToString(), mConnection))
                        {
                            myCmd.CommandType = CommandType.Text;
                            myCmd.ExecuteNonQuery();
                        }
                        mConnection.Close();
                    }
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }

            try {
                cmd = new MySqlCommand("DELETE FROM `dumpbilling`", con);
                cmd.CommandType = System.Data.CommandType.Text;
                con.Open();
                cmd.ExecuteNonQuery();
                con.Close();
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
    }
}

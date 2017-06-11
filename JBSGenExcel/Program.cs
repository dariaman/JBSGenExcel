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
                else if (args[0] == "resultmandiricc")
                {
                    FileInfo FileName = new FileInfo(DirResult + args[1].ToString());
                    if (FileName.Exists) resultMandiricc(FileName);
                    else throw new Exception(@"File tidak ditemukan => " + FileName.FullName + "xxxx");
                }
                else if (args[0] == "resultmegaonuscc")
                {
                    FileInfo FileName = new FileInfo(DirResult + args[1].ToString());
                    if (FileName.Exists) resultMegaOnUscc(FileName);
                    else throw new Exception(@"File tidak ditemukan => " + FileName.FullName + "xxxx");
                }
                else if (args[0] == "resultmegaoffuscc")
                {
                    FileInfo FileName = new FileInfo(DirResult + args[1].ToString());
                    if (FileName.Exists) resultMegaOffUscc(FileName);
                    else throw new Exception(@"File tidak ditemukan => " + FileName.FullName + "xxxx");
                }
                else if (args[0] == "resultbnicc")
                {
                    FileInfo FileName = new FileInfo(DirResult + args[1].ToString());
                    if (FileName.Exists) resultBNIcc(FileName);
                    else throw new Exception(@"File tidak ditemukan => " + FileName.FullName + "xxxx");
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
                            if (row == null)  row = sheet.CreateRow(i);
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
        public static void resultMandiricc(FileInfo FileName)
        {

            HSSFWorkbook hssfwb = new HSSFWorkbook();
            string NoPolis = "", AppCode = "", Description = "";
            bool isApprove = false;

            int PolicyID = -1, BillingID = -1, recurring_seq = -1, CycleDate = 0;
            DateTime DueDatePre = new DateTime(2000, 1, 1), BillDate = new DateTime(2000, 1, 1);
            decimal BillAmount = 0;
            string Period = "", CCno = "", CCexp = "", ccName = "", addr = "", telp = "";

            string xFileName = Path.GetFileNameWithoutExtension(FileName.FullName).ToLower() +
                Path.GetRandomFileName().Replace(".", "").Substring(0, 8).ToLower() + ".xls";
            using (FileStream file = new FileStream(FileName.FullName, FileMode.Open, FileAccess.Read))
            {
                MySqlConnection con = new MySqlConnection(constring);
                MySqlConnection conLife21 = new MySqlConnection(con21);
                MySqlCommand cmdjbs;
                MySqlCommand cmd21;
                MySqlTransaction tranjbs;
                MySqlTransaction tran21;

                hssfwb = new HSSFWorkbook(file);
                ISheet sheet = hssfwb.GetSheetAt(0); // Utk Mandiri CC sheet 1 adalah transaksi yg sukses
                int row = 0;
                for (row = 1; row <= sheet.LastRowNum; row++) // mulai dari baris 2
                {
                    con.Open();
                    conLife21.Open();
                    // transaction for JBS
                    tranjbs = con.BeginTransaction();
                    cmdjbs = con.CreateCommand();
                    cmdjbs.Transaction = tranjbs;

                    // transaction for Life21
                    tran21 = conLife21.BeginTransaction();
                    cmd21 = conLife21.CreateCommand();
                    cmd21.Transaction = tran21;
                    try
                    {
                        if (sheet.GetRow(row) != null)
                        {
                            // Jika cell null karena beda cell null dengan cell empty
                            if (sheet.GetRow(row).GetCell(5) == null)  continue; // untuk kolom polisNo

                            NoPolis = Convert.ToString(sheet.GetRow(row).GetCell(5));
                            if (NoPolis == "") continue; // menghindari nopolis kosong
                            AppCode = Convert.ToString(sheet.GetRow(row).GetCell(3)); // pada result MandiriCC kolom AuthCode
                            Description = Convert.ToString(sheet.GetRow(row).GetCell(4)); // Pada Kolom TC
                            isApprove = true;

                            // Ambil data polis billing yang akan di update
                            cmdjbs.Parameters.Clear();
                            cmdjbs.CommandType = CommandType.StoredProcedure;
                            cmdjbs.CommandText = @"FindPolisCCGetBillSeq";
                            cmdjbs.Parameters.Add(new MySqlParameter("@NoPolis", MySqlDbType.VarChar) { Value = NoPolis });
                            using (var rd = cmdjbs.ExecuteReader())
                            {
                                while (rd.Read())
                                {
                                    PolicyID = Convert.ToInt32(rd["policy_id"]);
                                    BillingID = Convert.ToInt32(rd["BillingID"]);
                                    recurring_seq = Convert.ToInt32(rd["recurring_seq"]);
                                    BillDate = Convert.ToDateTime(rd["BillingDate"]);
                                    DueDatePre = Convert.ToDateTime(rd["due_dt_pre"]);
                                    BillAmount = Convert.ToDecimal(rd["TotalAmount"]);

                                    Period = rd["PeriodeBilling"].ToString();
                                    CycleDate = Convert.ToInt32(rd["cycleDate"]);
                                    CCno = rd["cc_no"].ToString();
                                    CCexp = rd["cc_expiry"].ToString();
                                    ccName = rd["cc_name"].ToString();
                                    addr = rd["cc_address"].ToString();
                                    telp = rd["cc_telephone"].ToString();
                                }

                                if (PolicyID < 1 || BillingID < 1 || recurring_seq < 1)
                                {
                                    throw new Exception("Polis tidak ditemukan,mungkin billingnya tidak dalam status download atau terdapat kesalahan pada data textfile...");
                                }
                            }

                            // insert transaction bank JBS approve atw reject
                            cmdjbs.Parameters.Clear();
                            cmdjbs.CommandType = CommandType.StoredProcedure;
                            cmdjbs.CommandText = @"InsertTransactionBank;";
                            cmdjbs.Parameters.Add(new MySqlParameter("@FileName", MySqlDbType.VarChar) { Value = xFileName });
                            cmdjbs.Parameters.Add(new MySqlParameter("@Trancode", MySqlDbType.VarChar) { Value = "mandiricc" }); // hardCode BNI CC
                            cmdjbs.Parameters.Add(new MySqlParameter("@IsApprove", MySqlDbType.Bit) { Value = isApprove });
                            cmdjbs.Parameters.Add(new MySqlParameter("@policyID", MySqlDbType.VarChar) { Value = PolicyID });
                            cmdjbs.Parameters.Add(new MySqlParameter("@Seq", MySqlDbType.VarChar) { Value = recurring_seq });
                            cmdjbs.Parameters.Add(new MySqlParameter("@IDBill", MySqlDbType.VarChar) { Value = BillingID });
                            cmdjbs.Parameters.Add(new MySqlParameter("@approvalCode", MySqlDbType.VarChar) { Value = AppCode });
                            cmdjbs.Parameters.Add(new MySqlParameter("@BankID", MySqlDbType.Int32) { Value = 0 }); // bukan BCA (boleh bankCode asli asal jgn 1)
                            cmdjbs.Parameters.Add(new MySqlParameter("@ErrCode", MySqlDbType.VarChar) { Value = Description });
                            var uid = cmdjbs.ExecuteScalar().ToString();

                            //{// ============================ Proses Insert Received ===========================
                                cmd21.Parameters.Clear();
                                cmd21.CommandType = CommandType.StoredProcedure;
                                cmd21.CommandText = @"ReceiptInsert";
                                cmd21.Parameters.Add(new MySqlParameter("@BillingDate", MySqlDbType.Date) { Value = BillDate });
                                cmd21.Parameters.Add(new MySqlParameter("@policy_id", MySqlDbType.Int32) { Value = PolicyID });
                                cmd21.Parameters.Add(new MySqlParameter("@receipt_amount", MySqlDbType.Decimal) { Value = BillAmount });
                                cmd21.Parameters.Add(new MySqlParameter("@Source_download", MySqlDbType.VarChar) { Value = "CC" });
                                cmd21.Parameters.Add(new MySqlParameter("@recurring_seq", MySqlDbType.Int32) { Value = recurring_seq });
                                cmd21.Parameters.Add(new MySqlParameter("@bank_acc_id", MySqlDbType.Int32) { Value = 2 }); // Mandiri
                                cmd21.Parameters.Add(new MySqlParameter("@due_dt_pre", MySqlDbType.Date) { Value = DueDatePre });
                                var receiptID = cmd21.ExecuteScalar().ToString();

                                // ============================ Proses Insert Pilis CC Transaction Life21 ===========================
                                cmd21.Parameters.Clear();
                                cmd21.CommandType = CommandType.StoredProcedure;
                                cmd21.CommandText = @"InsertPolistransCC";
                                cmd21.Parameters.Add(new MySqlParameter("@PolisID", MySqlDbType.Int32) { Value = PolicyID });
                                cmd21.Parameters.Add(new MySqlParameter("@Transdate", MySqlDbType.Date) { Value = BillDate });
                                cmd21.Parameters.Add(new MySqlParameter("@Seq", MySqlDbType.Int32) { Value = recurring_seq });
                                cmd21.Parameters.Add(new MySqlParameter("@Amount", MySqlDbType.Decimal) { Value = BillAmount });
                                cmd21.Parameters.Add(new MySqlParameter("@DueDatePre", MySqlDbType.Date) { Value = DueDatePre });
                                cmd21.Parameters.Add(new MySqlParameter("@Period", MySqlDbType.VarChar) { Value = Period });
                                cmd21.Parameters.Add(new MySqlParameter("@CycleDate", MySqlDbType.Int32) { Value = CycleDate });
                                cmd21.Parameters.Add(new MySqlParameter("@BankID", MySqlDbType.Int32) { Value = 2 }); // Mandiri
                                cmd21.Parameters.Add(new MySqlParameter("@CCno", MySqlDbType.VarChar) { Value = CCno });
                                cmd21.Parameters.Add(new MySqlParameter("@CCExpiry", MySqlDbType.VarChar) { Value = CCexp });
                                cmd21.Parameters.Add(new MySqlParameter("@CCName", MySqlDbType.VarChar) { Value = ccName });
                                cmd21.Parameters.Add(new MySqlParameter("@CCAddrs", MySqlDbType.VarChar) { Value = addr });
                                cmd21.Parameters.Add(new MySqlParameter("@CCtelp", MySqlDbType.VarChar) { Value = telp });
                                cmd21.Parameters.Add(new MySqlParameter("@receiptID", MySqlDbType.Int32) { Value = receiptID });
                                var CCTransID = cmd21.ExecuteScalar().ToString();

                                // Update table billing
                                cmdjbs.Parameters.Clear();
                                cmdjbs.CommandType = CommandType.Text;
                                cmdjbs.CommandText = @"UPDATE `billing` SET `IsDownload`=0,
			                                                `IsClosed`=1,
			                                                `status_billing`='P',
			                                                `status_billing_dateUpdate`=@tgl,
			                                                `paid_date`=@billDate,
                                                            Life21TranID=@TransactionID,
			                                                `ReceiptID`=@receiptID,
			                                                `PaymentTransactionID`=@uid
		                                                WHERE `BillingID`=@idBill;";
                                cmdjbs.Parameters.Add(new MySqlParameter("@tgl", MySqlDbType.DateTime) { Value = DateTime.Now });
                                cmdjbs.Parameters.Add(new MySqlParameter("@billDate", MySqlDbType.DateTime) { Value = BillDate });
                                cmdjbs.Parameters.Add(new MySqlParameter("@TransactionID", MySqlDbType.Int32) { Value = CCTransID });
                                cmdjbs.Parameters.Add(new MySqlParameter("@receiptID", MySqlDbType.Int32) { Value = receiptID });
                                cmdjbs.Parameters.Add(new MySqlParameter("@uid", MySqlDbType.VarChar) { Value = uid });
                                cmdjbs.Parameters.Add(new MySqlParameter("@idBill", MySqlDbType.Int32) { Value = BillingID });
                                cmdjbs.ExecuteNonQuery();

                                // Update Polis Last Transaction
                                cmdjbs.Parameters.Clear();
                                cmdjbs.CommandType = CommandType.Text;
                                cmdjbs.CommandText = @"UPDATE `policy_last_trans` AS pt
		                                                INNER JOIN `billing` AS bx ON bx.policy_id=pt.policy_Id
			                                                SET pt.BillingID=bx.BillingID,
			                                                pt.recurring_seq=bx.recurring_seq,
			                                                pt.due_dt_pre=bx.due_dt_pre,
			                                                pt.source=bx.Source_download,
			                                                pt.receipt_id=bx.`ReceiptID`,
			                                                pt.receipt_date=bx.BillingDate,
			                                                pt.bank_id=bx.BankIdDownload
		                                                WHERE pt.policy_Id=@policyID AND bx.BillingID=@idBill;";
                                cmdjbs.Parameters.Add(new MySqlParameter("@policyID", MySqlDbType.Int32) { Value = PolicyID });
                                cmdjbs.Parameters.Add(new MySqlParameter("@idBill", MySqlDbType.Int32) { Value = BillingID });
                                cmdjbs.ExecuteNonQuery();
                        }
                        tranjbs.Commit();
                        tran21.Commit();

                        PolicyID = -1;
                        BillingID = -1;
                        recurring_seq = -1;
                        AppCode = "";
                        Description = "";
                    }
                    catch (Exception ex)
                    {
                        tranjbs.Rollback();
                        tran21.Rollback();
                        cmdjbs.CommandType = CommandType.Text;
                        cmdjbs.Parameters.Clear();
                        cmdjbs.CommandText = @"INSERT INTO `log_error_upload_result`(TranCode,line,FileName,exceptionApp)
                                            SELECT @TranCode,@line,@FileName,@exceptionApp";
                        cmdjbs.Parameters.Add(new MySqlParameter("@TranCode", MySqlDbType.VarChar) { Value = "bnicc" });
                        cmdjbs.Parameters.Add(new MySqlParameter("@line", MySqlDbType.Int32) { Value = row });
                        cmdjbs.Parameters.Add(new MySqlParameter("@FileName", MySqlDbType.VarChar) { Value = xFileName });
                        cmdjbs.Parameters.Add(new MySqlParameter("@exceptionApp", MySqlDbType.VarChar) { Value ="S1, " +  ex.Message.Substring(0, ex.Message.Length < 255 ? ex.Message.Length : 250) });
                        cmdjbs.ExecuteNonQuery();
                    }
                    finally
                    {
                        con.Dispose();
                        con.Close();
                        conLife21.Close();
                    }
                }

                //// Untuk Transaksi yang gagal
                sheet = hssfwb.GetSheetAt(1); // Utk Mandiri CC sheet 2 adalah transaksi yg gagal (Reject)
                isApprove = false;
                for (row = 1; row <= sheet.LastRowNum; row++) // mulai dari baris 2
                {
                    con.Open();
                    // transaction for JBS
                    tranjbs = con.BeginTransaction();
                    cmdjbs = con.CreateCommand();
                    cmdjbs.Transaction = tranjbs;
                    try
                    {
                        if (sheet.GetRow(row) != null)
                        {
                            // Jika cell null karena beda cell null dengan cell empty
                            if (sheet.GetRow(row).GetCell(3) == null) continue; // untuk kolom polisNo

                            NoPolis = Convert.ToString(sheet.GetRow(row).GetCell(3));
                            if (NoPolis == "") continue; // menghindari nopolis kosong
                            AppCode = Convert.ToString(sheet.GetRow(row).GetCell(4)); // pada result MandiriCC kolom AuthCode
                            Description = Convert.ToString(sheet.GetRow(row).GetCell(5)); // Pada Kolom TC

                            // Ambil data polis billing yang akan di update
                            cmdjbs.Parameters.Clear();
                            cmdjbs.CommandType = CommandType.StoredProcedure;
                            cmdjbs.CommandText = @"FindPolisCCGetBillSeq";
                            cmdjbs.Parameters.Add(new MySqlParameter("@NoPolis", MySqlDbType.VarChar) { Value = NoPolis });
                            using (var rd = cmdjbs.ExecuteReader())
                            {
                                while (rd.Read())
                                {
                                    PolicyID = Convert.ToInt32(rd["policy_id"]);
                                    BillingID = Convert.ToInt32(rd["BillingID"]);
                                    recurring_seq = Convert.ToInt32(rd["recurring_seq"]);
                                }

                                if (PolicyID < 1 || BillingID < 1 || recurring_seq < 1)
                                {
                                    throw new Exception("Polis tidak ditemukan,mungkin billingnya tidak dalam status download atau terdapat kesalahan pada data textfile...");
                                }
                            }
                            // insert transaction bank JBS approve atw reject
                            cmdjbs.Parameters.Clear();
                            cmdjbs.CommandType = CommandType.StoredProcedure;
                            cmdjbs.CommandText = @"InsertTransactionBank;";
                            cmdjbs.Parameters.Add(new MySqlParameter("@FileName", MySqlDbType.VarChar) { Value = xFileName });
                            cmdjbs.Parameters.Add(new MySqlParameter("@Trancode", MySqlDbType.VarChar) { Value = "mandiricc" }); // hardCode BNI CC
                            cmdjbs.Parameters.Add(new MySqlParameter("@IsApprove", MySqlDbType.Bit) { Value = false });
                            cmdjbs.Parameters.Add(new MySqlParameter("@policyID", MySqlDbType.VarChar) { Value = PolicyID });
                            cmdjbs.Parameters.Add(new MySqlParameter("@Seq", MySqlDbType.VarChar) { Value = recurring_seq });
                            cmdjbs.Parameters.Add(new MySqlParameter("@IDBill", MySqlDbType.VarChar) { Value = BillingID });
                            cmdjbs.Parameters.Add(new MySqlParameter("@approvalCode", MySqlDbType.VarChar) { Value = AppCode });
                            cmdjbs.Parameters.Add(new MySqlParameter("@BankID", MySqlDbType.Int32) { Value = 0 }); // bukan BCA (boleh bankCode asli asal jgn 1)
                            cmdjbs.Parameters.Add(new MySqlParameter("@ErrCode", MySqlDbType.VarChar) { Value = Description });
                            var uid = cmdjbs.ExecuteScalar().ToString();

                            cmdjbs.CommandType = CommandType.Text;
                            cmdjbs.Parameters.Clear();
                            cmdjbs.CommandText = @"UPDATE `billing` SET IsDownload=0 WHERE `BillingID`=@billid";
                            cmdjbs.Parameters.Add(new MySqlParameter("@billid", MySqlDbType.Int32) { Value = BillingID });
                            cmdjbs.ExecuteNonQuery();
                            tranjbs.Commit();
                        }
                    }
                    catch (Exception ex)
                    {
                        tranjbs.Rollback();
                        cmdjbs.CommandType = CommandType.Text;
                        cmdjbs.Parameters.Clear();
                        cmdjbs.CommandText = @"INSERT INTO `log_error_upload_result`(TranCode,line,FileName,exceptionApp)
                                            SELECT @TranCode,@line,@FileName,@exceptionApp";
                        cmdjbs.Parameters.Add(new MySqlParameter("@TranCode", MySqlDbType.VarChar) { Value = "bnicc" });
                        cmdjbs.Parameters.Add(new MySqlParameter("@line", MySqlDbType.Int32) { Value = row });
                        cmdjbs.Parameters.Add(new MySqlParameter("@FileName", MySqlDbType.VarChar) { Value = xFileName });
                        cmdjbs.Parameters.Add(new MySqlParameter("@exceptionApp", MySqlDbType.VarChar) { Value ="S2, " +  ex.Message.Substring(0, ex.Message.Length < 255 ? ex.Message.Length : 253) });
                        cmdjbs.ExecuteNonQuery();
                    }
                    finally
                    {
                        con.Dispose();
                        con.Close();
                    }
                    PolicyID = -1;
                    BillingID = -1;
                    recurring_seq = -1;
                    AppCode = "";
                    Description = "";
                }

                    //hssfwb.Dispose();
                    file.Close();
            }
        }
        public static void resultMegaOnUscc(FileInfo FileName)
        {

            HSSFWorkbook hssfwb = new HSSFWorkbook();
            string NoPolis = "", AppCode = "", Description = "";
            bool isApprove = false;

            int PolicyID = -1, BillingID = -1, recurring_seq = -1, CycleDate = 0;
            DateTime DueDatePre = new DateTime(2000, 1, 1), BillDate = new DateTime(2000, 1, 1);
            decimal BillAmount = 0;
            string Period = "", CCno = "", CCexp = "", ccName = "", addr = "", telp = "";

            string xFileName = Path.GetFileNameWithoutExtension(FileName.FullName).ToLower() +
                Path.GetRandomFileName().Replace(".", "").Substring(0, 8).ToLower() + ".xls";
            using (FileStream file = new FileStream(FileName.FullName, FileMode.Open, FileAccess.Read))
            {
                MySqlConnection con = new MySqlConnection(constring);
                MySqlConnection conLife21 = new MySqlConnection(con21);
                MySqlCommand cmdjbs;
                MySqlCommand cmd21;
                MySqlTransaction tranjbs;
                MySqlTransaction tran21;

                hssfwb = new HSSFWorkbook(file);
                ISheet sheet = hssfwb.GetSheetAt(0); // Utk sheet 1 adalah transaksi yg sukses
                int row = 0;
                for (row = 1; row <= sheet.LastRowNum; row++) // mulai dari baris 2
                {
                    con.Open();
                    conLife21.Open();
                    // transaction for JBS
                    tranjbs = con.BeginTransaction();
                    cmdjbs = con.CreateCommand();
                    cmdjbs.Transaction = tranjbs;

                    // transaction for Life21
                    tran21 = conLife21.BeginTransaction();
                    cmd21 = conLife21.CreateCommand();
                    cmd21.Transaction = tran21;
                    try
                    {
                        if (sheet.GetRow(row) != null)
                        {
                            // Jika cell null karena beda cell null dengan cell empty
                            if (sheet.GetRow(row).GetCell(1) == null) continue; // untuk kolom polisNo

                            NoPolis = Convert.ToString(sheet.GetRow(row).GetCell(1));
                            if (NoPolis == "") continue; // menghindari nopolis kosong
                            NoPolis = NoPolis.Substring(NoPolis.Length-11); // ambil 11 karakter di kanan
                            AppCode = Convert.ToString(sheet.GetRow(row).GetCell(4)); // pada result MandiriCC kolom AuthCode
                            Description = Convert.ToString(sheet.GetRow(row).GetCell(5)); // Pada Kolom TC
                            isApprove = true;

                            // Ambil data polis billing yang akan di update
                            cmdjbs.Parameters.Clear();
                            cmdjbs.CommandType = CommandType.StoredProcedure;
                            cmdjbs.CommandText = @"FindPolisCCGetBillSeq";
                            cmdjbs.Parameters.Add(new MySqlParameter("@NoPolis", MySqlDbType.VarChar) { Value = NoPolis });
                            using (var rd = cmdjbs.ExecuteReader())
                            {
                                while (rd.Read())
                                {
                                    PolicyID = Convert.ToInt32(rd["policy_id"]);
                                    BillingID = Convert.ToInt32(rd["BillingID"]);
                                    recurring_seq = Convert.ToInt32(rd["recurring_seq"]);
                                    BillDate = Convert.ToDateTime(rd["BillingDate"]);
                                    DueDatePre = Convert.ToDateTime(rd["due_dt_pre"]);
                                    BillAmount = Convert.ToDecimal(rd["TotalAmount"]);

                                    Period = rd["PeriodeBilling"].ToString();
                                    CycleDate = Convert.ToInt32(rd["cycleDate"]);
                                    CCno = rd["cc_no"].ToString();
                                    CCexp = rd["cc_expiry"].ToString();
                                    ccName = rd["cc_name"].ToString();
                                    addr = rd["cc_address"].ToString();
                                    telp = rd["cc_telephone"].ToString();
                                }

                                if (PolicyID < 1 || BillingID < 1 || recurring_seq < 1)
                                {
                                    throw new Exception("Polis tidak ditemukan,mungkin billingnya tidak dalam status download atau terdapat kesalahan pada data textfile...");
                                }
                            }

                            // insert transaction bank JBS approve atw reject
                            cmdjbs.Parameters.Clear();
                            cmdjbs.CommandType = CommandType.StoredProcedure;
                            cmdjbs.CommandText = @"InsertTransactionBank;";
                            cmdjbs.Parameters.Add(new MySqlParameter("@FileName", MySqlDbType.VarChar) { Value = xFileName });
                            cmdjbs.Parameters.Add(new MySqlParameter("@Trancode", MySqlDbType.VarChar) { Value = "megaonus" }); // hardCode MegaOnUs CC
                            cmdjbs.Parameters.Add(new MySqlParameter("@IsApprove", MySqlDbType.Bit) { Value = isApprove });
                            cmdjbs.Parameters.Add(new MySqlParameter("@policyID", MySqlDbType.VarChar) { Value = PolicyID });
                            cmdjbs.Parameters.Add(new MySqlParameter("@Seq", MySqlDbType.VarChar) { Value = recurring_seq });
                            cmdjbs.Parameters.Add(new MySqlParameter("@IDBill", MySqlDbType.VarChar) { Value = BillingID });
                            cmdjbs.Parameters.Add(new MySqlParameter("@approvalCode", MySqlDbType.VarChar) { Value = AppCode });
                            cmdjbs.Parameters.Add(new MySqlParameter("@BankID", MySqlDbType.Int32) { Value = 0 }); // bukan BCA (boleh bankCode asli asal jgn 1)
                            cmdjbs.Parameters.Add(new MySqlParameter("@ErrCode", MySqlDbType.VarChar) { Value = Description });
                            var uid = cmdjbs.ExecuteScalar().ToString();

                            //{// ============================ Proses Insert Received ===========================
                            cmd21.Parameters.Clear();
                            cmd21.CommandType = CommandType.StoredProcedure;
                            cmd21.CommandText = @"ReceiptInsert";
                            cmd21.Parameters.Add(new MySqlParameter("@BillingDate", MySqlDbType.Date) { Value = BillDate });
                            cmd21.Parameters.Add(new MySqlParameter("@policy_id", MySqlDbType.Int32) { Value = PolicyID });
                            cmd21.Parameters.Add(new MySqlParameter("@receipt_amount", MySqlDbType.Decimal) { Value = BillAmount });
                            cmd21.Parameters.Add(new MySqlParameter("@Source_download", MySqlDbType.VarChar) { Value = "CC" });
                            cmd21.Parameters.Add(new MySqlParameter("@recurring_seq", MySqlDbType.Int32) { Value = recurring_seq });
                            cmd21.Parameters.Add(new MySqlParameter("@bank_acc_id", MySqlDbType.Int32) { Value = 12 }); // Mega
                            cmd21.Parameters.Add(new MySqlParameter("@due_dt_pre", MySqlDbType.Date) { Value = DueDatePre });
                            var receiptID = cmd21.ExecuteScalar().ToString();

                            // ============================ Proses Insert Pilis CC Transaction Life21 ===========================
                            cmd21.Parameters.Clear();
                            cmd21.CommandType = CommandType.StoredProcedure;
                            cmd21.CommandText = @"InsertPolistransCC";
                            cmd21.Parameters.Add(new MySqlParameter("@PolisID", MySqlDbType.Int32) { Value = PolicyID });
                            cmd21.Parameters.Add(new MySqlParameter("@Transdate", MySqlDbType.Date) { Value = BillDate });
                            cmd21.Parameters.Add(new MySqlParameter("@Seq", MySqlDbType.Int32) { Value = recurring_seq });
                            cmd21.Parameters.Add(new MySqlParameter("@Amount", MySqlDbType.Decimal) { Value = BillAmount });
                            cmd21.Parameters.Add(new MySqlParameter("@DueDatePre", MySqlDbType.Date) { Value = DueDatePre });
                            cmd21.Parameters.Add(new MySqlParameter("@Period", MySqlDbType.VarChar) { Value = Period });
                            cmd21.Parameters.Add(new MySqlParameter("@CycleDate", MySqlDbType.Int32) { Value = CycleDate });
                            cmd21.Parameters.Add(new MySqlParameter("@BankID", MySqlDbType.Int32) { Value = 12 }); // Mega
                            cmd21.Parameters.Add(new MySqlParameter("@CCno", MySqlDbType.VarChar) { Value = CCno });
                            cmd21.Parameters.Add(new MySqlParameter("@CCExpiry", MySqlDbType.VarChar) { Value = CCexp });
                            cmd21.Parameters.Add(new MySqlParameter("@CCName", MySqlDbType.VarChar) { Value = ccName });
                            cmd21.Parameters.Add(new MySqlParameter("@CCAddrs", MySqlDbType.VarChar) { Value = addr });
                            cmd21.Parameters.Add(new MySqlParameter("@CCtelp", MySqlDbType.VarChar) { Value = telp });
                            cmd21.Parameters.Add(new MySqlParameter("@receiptID", MySqlDbType.Int32) { Value = receiptID });
                            var CCTransID = cmd21.ExecuteScalar().ToString();

                            // Update table billing
                            cmdjbs.Parameters.Clear();
                            cmdjbs.CommandType = CommandType.Text;
                            cmdjbs.CommandText = @"UPDATE `billing` SET `IsDownload`=0,
			                                                `IsClosed`=1,
			                                                `status_billing`='P',
			                                                `status_billing_dateUpdate`=@tgl,
			                                                `paid_date`=@billDate,
                                                            Life21TranID=@TransactionID,
			                                                `ReceiptID`=@receiptID,
			                                                `PaymentTransactionID`=@uid
		                                                WHERE `BillingID`=@idBill;";
                            cmdjbs.Parameters.Add(new MySqlParameter("@tgl", MySqlDbType.DateTime) { Value = DateTime.Now });
                            cmdjbs.Parameters.Add(new MySqlParameter("@billDate", MySqlDbType.DateTime) { Value = BillDate });
                            cmdjbs.Parameters.Add(new MySqlParameter("@TransactionID", MySqlDbType.Int32) { Value = CCTransID });
                            cmdjbs.Parameters.Add(new MySqlParameter("@receiptID", MySqlDbType.Int32) { Value = receiptID });
                            cmdjbs.Parameters.Add(new MySqlParameter("@uid", MySqlDbType.VarChar) { Value = uid });
                            cmdjbs.Parameters.Add(new MySqlParameter("@idBill", MySqlDbType.Int32) { Value = BillingID });
                            cmdjbs.ExecuteNonQuery();

                            // Update Polis Last Transaction
                            cmdjbs.Parameters.Clear();
                            cmdjbs.CommandType = CommandType.Text;
                            cmdjbs.CommandText = @"UPDATE `policy_last_trans` AS pt
		                                                INNER JOIN `billing` AS bx ON bx.policy_id=pt.policy_Id
			                                                SET pt.BillingID=bx.BillingID,
			                                                pt.recurring_seq=bx.recurring_seq,
			                                                pt.due_dt_pre=bx.due_dt_pre,
			                                                pt.source=bx.Source_download,
			                                                pt.receipt_id=bx.`ReceiptID`,
			                                                pt.receipt_date=bx.BillingDate,
			                                                pt.bank_id=bx.BankIdDownload
		                                                WHERE pt.policy_Id=@policyID AND bx.BillingID=@idBill;";
                            cmdjbs.Parameters.Add(new MySqlParameter("@policyID", MySqlDbType.Int32) { Value = PolicyID });
                            cmdjbs.Parameters.Add(new MySqlParameter("@idBill", MySqlDbType.Int32) { Value = BillingID });
                            cmdjbs.ExecuteNonQuery();
                        }
                        tranjbs.Commit();
                        tran21.Commit();
                    }
                    catch (Exception ex)
                    {
                        tranjbs.Rollback();
                        tran21.Rollback();
                        cmdjbs.CommandType = CommandType.Text;
                        cmdjbs.Parameters.Clear();
                        cmdjbs.CommandText = @"INSERT INTO `log_error_upload_result`(TranCode,line,FileName,exceptionApp)
                                            SELECT @TranCode,@line,@FileName,@exceptionApp";
                        cmdjbs.Parameters.Add(new MySqlParameter("@TranCode", MySqlDbType.VarChar) { Value = "megaonus" });
                        cmdjbs.Parameters.Add(new MySqlParameter("@line", MySqlDbType.Int32) { Value = row });
                        cmdjbs.Parameters.Add(new MySqlParameter("@FileName", MySqlDbType.VarChar) { Value = xFileName });
                        cmdjbs.Parameters.Add(new MySqlParameter("@exceptionApp", MySqlDbType.VarChar) { Value = "S1, " + ex.Message.Substring(0, ex.Message.Length < 255 ? ex.Message.Length : 250) });
                        cmdjbs.ExecuteNonQuery();
                    }
                    finally
                    {
                        con.Dispose();
                        con.Close();
                        conLife21.Close();
                    }

                    PolicyID = -1;
                    BillingID = -1;
                    recurring_seq = -1;
                    AppCode = "";
                    Description = "";
                }

                //// Untuk Transaksi yang gagal
                sheet = hssfwb.GetSheetAt(1); // Utk Mandiri CC sheet 2 adalah transaksi yg gagal (Reject)
                isApprove = false;
                for (row = 1; row <= sheet.LastRowNum; row++) // mulai dari baris 2
                {
                    con.Open();
                    // transaction for JBS
                    tranjbs = con.BeginTransaction();
                    cmdjbs = con.CreateCommand();
                    cmdjbs.Transaction = tranjbs;
                    try
                    {
                        if (sheet.GetRow(row) != null)
                        {
                            // Jika cell null karena beda cell null dengan cell empty
                            if (sheet.GetRow(row).GetCell(1) == null) continue; // untuk kolom polisNo

                            NoPolis = Convert.ToString(sheet.GetRow(row).GetCell(1));
                            if (NoPolis == "") continue; // menghindari nopolis kosong
                            NoPolis = NoPolis.Substring(NoPolis.Length - 11); // ambil 11 karakter di kanan
                            AppCode = Convert.ToString(sheet.GetRow(row).GetCell(4));
                            Description = Convert.ToString(sheet.GetRow(row).GetCell(5));

                            // Ambil data polis billing yang akan di update
                            cmdjbs.Parameters.Clear();
                            cmdjbs.CommandType = CommandType.StoredProcedure;
                            cmdjbs.CommandText = @"FindPolisCCGetBillSeq";
                            cmdjbs.Parameters.Add(new MySqlParameter("@NoPolis", MySqlDbType.VarChar) { Value = NoPolis });
                            using (var rd = cmdjbs.ExecuteReader())
                            {
                                while (rd.Read())
                                {
                                    PolicyID = Convert.ToInt32(rd["policy_id"]);
                                    BillingID = Convert.ToInt32(rd["BillingID"]);
                                    recurring_seq = Convert.ToInt32(rd["recurring_seq"]);
                                }

                                if (PolicyID < 1 || BillingID < 1 || recurring_seq < 1)
                                {
                                    throw new Exception("Polis tidak ditemukan,mungkin billingnya tidak dalam status download atau terdapat kesalahan pada data textfile...");
                                }
                            }
                            // insert transaction bank JBS approve atw reject
                            cmdjbs.Parameters.Clear();
                            cmdjbs.CommandType = CommandType.StoredProcedure;
                            cmdjbs.CommandText = @"InsertTransactionBank;";
                            cmdjbs.Parameters.Add(new MySqlParameter("@FileName", MySqlDbType.VarChar) { Value = xFileName });
                            cmdjbs.Parameters.Add(new MySqlParameter("@Trancode", MySqlDbType.VarChar) { Value = "megaonus" }); // hardCode MegaOnUs CC
                            cmdjbs.Parameters.Add(new MySqlParameter("@IsApprove", MySqlDbType.Bit) { Value = false });
                            cmdjbs.Parameters.Add(new MySqlParameter("@policyID", MySqlDbType.VarChar) { Value = PolicyID });
                            cmdjbs.Parameters.Add(new MySqlParameter("@Seq", MySqlDbType.VarChar) { Value = recurring_seq });
                            cmdjbs.Parameters.Add(new MySqlParameter("@IDBill", MySqlDbType.VarChar) { Value = BillingID });
                            cmdjbs.Parameters.Add(new MySqlParameter("@approvalCode", MySqlDbType.VarChar) { Value = AppCode });
                            cmdjbs.Parameters.Add(new MySqlParameter("@BankID", MySqlDbType.Int32) { Value = 0 }); // bukan BCA (boleh bankCode asli asal jgn 1)
                            cmdjbs.Parameters.Add(new MySqlParameter("@ErrCode", MySqlDbType.VarChar) { Value = Description });
                            var uid = cmdjbs.ExecuteScalar().ToString();

                            cmdjbs.CommandType = CommandType.Text;
                            cmdjbs.Parameters.Clear();
                            cmdjbs.CommandText = @"UPDATE `billing` SET IsDownload=0 WHERE `BillingID`=@billid";
                            cmdjbs.Parameters.Add(new MySqlParameter("@billid", MySqlDbType.Int32) { Value = BillingID });
                            cmdjbs.ExecuteNonQuery();
                            tranjbs.Commit();
                        }
                    }
                    catch (Exception ex)
                    {
                        tranjbs.Rollback();
                        cmdjbs.CommandType = CommandType.Text;
                        cmdjbs.Parameters.Clear();
                        cmdjbs.CommandText = @"INSERT INTO `log_error_upload_result`(TranCode,line,FileName,exceptionApp)
                                            SELECT @TranCode,@line,@FileName,@exceptionApp";
                        cmdjbs.Parameters.Add(new MySqlParameter("@TranCode", MySqlDbType.VarChar) { Value = "megaonus" });
                        cmdjbs.Parameters.Add(new MySqlParameter("@line", MySqlDbType.Int32) { Value = row });
                        cmdjbs.Parameters.Add(new MySqlParameter("@FileName", MySqlDbType.VarChar) { Value = xFileName });
                        cmdjbs.Parameters.Add(new MySqlParameter("@exceptionApp", MySqlDbType.VarChar) { Value = "S2, " + ex.Message.Substring(0, ex.Message.Length < 255 ? ex.Message.Length : 253) });
                        cmdjbs.ExecuteNonQuery();
                    }
                    finally
                    {
                        con.Dispose();
                        con.Close();
                    }
                    PolicyID = -1;
                    BillingID = -1;
                    recurring_seq = -1;
                    AppCode = "";
                    Description = "";
                }

                //hssfwb.Dispose();
                file.Close();
            }
        }
        public static void resultMegaOffUscc(FileInfo FileName)
        {

            HSSFWorkbook hssfwb = new HSSFWorkbook();
            string NoPolis = "", AppCode = "", Description = "";
            bool isApprove = false;

            int PolicyID = -1, BillingID = -1, recurring_seq = -1, CycleDate = 0;
            DateTime DueDatePre = new DateTime(2000, 1, 1), BillDate = new DateTime(2000, 1, 1);
            decimal BillAmount = 0;
            string Period = "", CCno = "", CCexp = "", ccName = "", addr = "", telp = "";

            string xFileName = Path.GetFileNameWithoutExtension(FileName.FullName).ToLower() +
                Path.GetRandomFileName().Replace(".", "").Substring(0, 8).ToLower() + ".xls";
            using (FileStream file = new FileStream(FileName.FullName, FileMode.Open, FileAccess.Read))
            {
                MySqlConnection con = new MySqlConnection(constring);
                MySqlConnection conLife21 = new MySqlConnection(con21);
                MySqlCommand cmdjbs;
                MySqlCommand cmd21;
                MySqlTransaction tranjbs;
                MySqlTransaction tran21;

                hssfwb = new HSSFWorkbook(file);
                ISheet sheet = hssfwb.GetSheetAt(0); // Utk sheet 1 adalah transaksi yg sukses
                int row = 0;
                for (row = 1; row <= sheet.LastRowNum; row++) // mulai dari baris 2
                {
                    con.Open();
                    conLife21.Open();
                    // transaction for JBS
                    tranjbs = con.BeginTransaction();
                    cmdjbs = con.CreateCommand();
                    cmdjbs.Transaction = tranjbs;

                    // transaction for Life21
                    tran21 = conLife21.BeginTransaction();
                    cmd21 = conLife21.CreateCommand();
                    cmd21.Transaction = tran21;
                    try
                    {
                        if (sheet.GetRow(row) != null)
                        {
                            // Jika cell null karena beda cell null dengan cell empty
                            if (sheet.GetRow(row).GetCell(1) == null) continue; // untuk kolom polisNo

                            NoPolis = Convert.ToString(sheet.GetRow(row).GetCell(1));
                            if (NoPolis == "") continue; // menghindari nopolis kosong
                            NoPolis = NoPolis.Substring(NoPolis.Length - 11); // ambil 11 karakter di kanan
                            AppCode = Convert.ToString(sheet.GetRow(row).GetCell(4)); // pada result MandiriCC kolom AuthCode
                            Description = Convert.ToString(sheet.GetRow(row).GetCell(5)); // Pada Kolom TC
                            isApprove = true;

                            // Ambil data polis billing yang akan di update
                            cmdjbs.Parameters.Clear();
                            cmdjbs.CommandType = CommandType.StoredProcedure;
                            cmdjbs.CommandText = @"FindPolisCCGetBillSeq";
                            cmdjbs.Parameters.Add(new MySqlParameter("@NoPolis", MySqlDbType.VarChar) { Value = NoPolis });
                            using (var rd = cmdjbs.ExecuteReader())
                            {
                                while (rd.Read())
                                {
                                    PolicyID = Convert.ToInt32(rd["policy_id"]);
                                    BillingID = Convert.ToInt32(rd["BillingID"]);
                                    recurring_seq = Convert.ToInt32(rd["recurring_seq"]);
                                    BillDate = Convert.ToDateTime(rd["BillingDate"]);
                                    DueDatePre = Convert.ToDateTime(rd["due_dt_pre"]);
                                    BillAmount = Convert.ToDecimal(rd["TotalAmount"]);

                                    Period = rd["PeriodeBilling"].ToString();
                                    CycleDate = Convert.ToInt32(rd["cycleDate"]);
                                    CCno = rd["cc_no"].ToString();
                                    CCexp = rd["cc_expiry"].ToString();
                                    ccName = rd["cc_name"].ToString();
                                    addr = rd["cc_address"].ToString();
                                    telp = rd["cc_telephone"].ToString();
                                }

                                if (PolicyID < 1 || BillingID < 1 || recurring_seq < 1)
                                {
                                    throw new Exception("Polis tidak ditemukan,mungkin billingnya tidak dalam status download atau terdapat kesalahan pada data textfile...");
                                }
                            }

                            // insert transaction bank JBS approve atw reject
                            cmdjbs.Parameters.Clear();
                            cmdjbs.CommandType = CommandType.StoredProcedure;
                            cmdjbs.CommandText = @"InsertTransactionBank;";
                            cmdjbs.Parameters.Add(new MySqlParameter("@FileName", MySqlDbType.VarChar) { Value = xFileName });
                            cmdjbs.Parameters.Add(new MySqlParameter("@Trancode", MySqlDbType.VarChar) { Value = "megaoffus" }); // hardCode MegaOnUs CC
                            cmdjbs.Parameters.Add(new MySqlParameter("@IsApprove", MySqlDbType.Bit) { Value = isApprove });
                            cmdjbs.Parameters.Add(new MySqlParameter("@policyID", MySqlDbType.VarChar) { Value = PolicyID });
                            cmdjbs.Parameters.Add(new MySqlParameter("@Seq", MySqlDbType.VarChar) { Value = recurring_seq });
                            cmdjbs.Parameters.Add(new MySqlParameter("@IDBill", MySqlDbType.VarChar) { Value = BillingID });
                            cmdjbs.Parameters.Add(new MySqlParameter("@approvalCode", MySqlDbType.VarChar) { Value = AppCode });
                            cmdjbs.Parameters.Add(new MySqlParameter("@BankID", MySqlDbType.Int32) { Value = 0 }); // bukan BCA (boleh bankCode asli asal jgn 1)
                            cmdjbs.Parameters.Add(new MySqlParameter("@ErrCode", MySqlDbType.VarChar) { Value = Description });
                            var uid = cmdjbs.ExecuteScalar().ToString();

                            //{// ============================ Proses Insert Received ===========================
                            cmd21.Parameters.Clear();
                            cmd21.CommandType = CommandType.StoredProcedure;
                            cmd21.CommandText = @"ReceiptInsert";
                            cmd21.Parameters.Add(new MySqlParameter("@BillingDate", MySqlDbType.Date) { Value = BillDate });
                            cmd21.Parameters.Add(new MySqlParameter("@policy_id", MySqlDbType.Int32) { Value = PolicyID });
                            cmd21.Parameters.Add(new MySqlParameter("@receipt_amount", MySqlDbType.Decimal) { Value = BillAmount });
                            cmd21.Parameters.Add(new MySqlParameter("@Source_download", MySqlDbType.VarChar) { Value = "CC" });
                            cmd21.Parameters.Add(new MySqlParameter("@recurring_seq", MySqlDbType.Int32) { Value = recurring_seq });
                            cmd21.Parameters.Add(new MySqlParameter("@bank_acc_id", MySqlDbType.Int32) { Value = 12 }); // Mega
                            cmd21.Parameters.Add(new MySqlParameter("@due_dt_pre", MySqlDbType.Date) { Value = DueDatePre });
                            var receiptID = cmd21.ExecuteScalar().ToString();

                            // ============================ Proses Insert Pilis CC Transaction Life21 ===========================
                            cmd21.Parameters.Clear();
                            cmd21.CommandType = CommandType.StoredProcedure;
                            cmd21.CommandText = @"InsertPolistransCC";
                            cmd21.Parameters.Add(new MySqlParameter("@PolisID", MySqlDbType.Int32) { Value = PolicyID });
                            cmd21.Parameters.Add(new MySqlParameter("@Transdate", MySqlDbType.Date) { Value = BillDate });
                            cmd21.Parameters.Add(new MySqlParameter("@Seq", MySqlDbType.Int32) { Value = recurring_seq });
                            cmd21.Parameters.Add(new MySqlParameter("@Amount", MySqlDbType.Decimal) { Value = BillAmount });
                            cmd21.Parameters.Add(new MySqlParameter("@DueDatePre", MySqlDbType.Date) { Value = DueDatePre });
                            cmd21.Parameters.Add(new MySqlParameter("@Period", MySqlDbType.VarChar) { Value = Period });
                            cmd21.Parameters.Add(new MySqlParameter("@CycleDate", MySqlDbType.Int32) { Value = CycleDate });
                            cmd21.Parameters.Add(new MySqlParameter("@BankID", MySqlDbType.Int32) { Value = 12 }); // Mega
                            cmd21.Parameters.Add(new MySqlParameter("@CCno", MySqlDbType.VarChar) { Value = CCno });
                            cmd21.Parameters.Add(new MySqlParameter("@CCExpiry", MySqlDbType.VarChar) { Value = CCexp });
                            cmd21.Parameters.Add(new MySqlParameter("@CCName", MySqlDbType.VarChar) { Value = ccName });
                            cmd21.Parameters.Add(new MySqlParameter("@CCAddrs", MySqlDbType.VarChar) { Value = addr });
                            cmd21.Parameters.Add(new MySqlParameter("@CCtelp", MySqlDbType.VarChar) { Value = telp });
                            cmd21.Parameters.Add(new MySqlParameter("@receiptID", MySqlDbType.Int32) { Value = receiptID });
                            var CCTransID = cmd21.ExecuteScalar().ToString();

                            // Update table billing
                            cmdjbs.Parameters.Clear();
                            cmdjbs.CommandType = CommandType.Text;
                            cmdjbs.CommandText = @"UPDATE `billing` SET `IsDownload`=0,
			                                                `IsClosed`=1,
			                                                `status_billing`='P',
			                                                `status_billing_dateUpdate`=@tgl,
			                                                `paid_date`=@billDate,
                                                            Life21TranID=@TransactionID,
			                                                `ReceiptID`=@receiptID,
			                                                `PaymentTransactionID`=@uid
		                                                WHERE `BillingID`=@idBill;";
                            cmdjbs.Parameters.Add(new MySqlParameter("@tgl", MySqlDbType.DateTime) { Value = DateTime.Now });
                            cmdjbs.Parameters.Add(new MySqlParameter("@billDate", MySqlDbType.DateTime) { Value = BillDate });
                            cmdjbs.Parameters.Add(new MySqlParameter("@TransactionID", MySqlDbType.Int32) { Value = CCTransID });
                            cmdjbs.Parameters.Add(new MySqlParameter("@receiptID", MySqlDbType.Int32) { Value = receiptID });
                            cmdjbs.Parameters.Add(new MySqlParameter("@uid", MySqlDbType.VarChar) { Value = uid });
                            cmdjbs.Parameters.Add(new MySqlParameter("@idBill", MySqlDbType.Int32) { Value = BillingID });
                            cmdjbs.ExecuteNonQuery();

                            // Update Polis Last Transaction
                            cmdjbs.Parameters.Clear();
                            cmdjbs.CommandType = CommandType.Text;
                            cmdjbs.CommandText = @"UPDATE `policy_last_trans` AS pt
		                                                INNER JOIN `billing` AS bx ON bx.policy_id=pt.policy_Id
			                                                SET pt.BillingID=bx.BillingID,
			                                                pt.recurring_seq=bx.recurring_seq,
			                                                pt.due_dt_pre=bx.due_dt_pre,
			                                                pt.source=bx.Source_download,
			                                                pt.receipt_id=bx.`ReceiptID`,
			                                                pt.receipt_date=bx.BillingDate,
			                                                pt.bank_id=bx.BankIdDownload
		                                                WHERE pt.policy_Id=@policyID AND bx.BillingID=@idBill;";
                            cmdjbs.Parameters.Add(new MySqlParameter("@policyID", MySqlDbType.Int32) { Value = PolicyID });
                            cmdjbs.Parameters.Add(new MySqlParameter("@idBill", MySqlDbType.Int32) { Value = BillingID });
                            cmdjbs.ExecuteNonQuery();
                        }
                        tranjbs.Commit();
                        tran21.Commit();
                    }
                    catch (Exception ex)
                    {
                        tranjbs.Rollback();
                        tran21.Rollback();
                        cmdjbs.CommandType = CommandType.Text;
                        cmdjbs.Parameters.Clear();
                        cmdjbs.CommandText = @"INSERT INTO `log_error_upload_result`(TranCode,line,FileName,exceptionApp)
                                            SELECT @TranCode,@line,@FileName,@exceptionApp";
                        cmdjbs.Parameters.Add(new MySqlParameter("@TranCode", MySqlDbType.VarChar) { Value = "megaoffus" });
                        cmdjbs.Parameters.Add(new MySqlParameter("@line", MySqlDbType.Int32) { Value = row });
                        cmdjbs.Parameters.Add(new MySqlParameter("@FileName", MySqlDbType.VarChar) { Value = xFileName });
                        cmdjbs.Parameters.Add(new MySqlParameter("@exceptionApp", MySqlDbType.VarChar) { Value = "S1, " + ex.Message.Substring(0, ex.Message.Length < 255 ? ex.Message.Length : 250) });
                        cmdjbs.ExecuteNonQuery();
                    }
                    finally
                    {
                        con.Dispose();
                        con.Close();
                        conLife21.Close();
                    }

                    PolicyID = -1;
                    BillingID = -1;
                    recurring_seq = -1;
                    AppCode = "";
                    Description = "";
                }

                //// Untuk Transaksi yang gagal
                sheet = hssfwb.GetSheetAt(1); // Utk Mandiri CC sheet 2 adalah transaksi yg gagal (Reject)
                isApprove = false;
                for (row = 1; row <= sheet.LastRowNum; row++) // mulai dari baris 2
                {
                    con.Open();
                    // transaction for JBS
                    tranjbs = con.BeginTransaction();
                    cmdjbs = con.CreateCommand();
                    cmdjbs.Transaction = tranjbs;
                    try
                    {
                        if (sheet.GetRow(row) != null)
                        {
                            // Jika cell null karena beda cell null dengan cell empty
                            if (sheet.GetRow(row).GetCell(1) == null) continue; // untuk kolom polisNo

                            NoPolis = Convert.ToString(sheet.GetRow(row).GetCell(1));
                            if (NoPolis == "") continue; // menghindari nopolis kosong
                            NoPolis = NoPolis.Substring(NoPolis.Length - 11); // ambil 11 karakter di kanan
                            AppCode = Convert.ToString(sheet.GetRow(row).GetCell(4));
                            Description = Convert.ToString(sheet.GetRow(row).GetCell(5));

                            // Ambil data polis billing yang akan di update
                            cmdjbs.Parameters.Clear();
                            cmdjbs.CommandType = CommandType.StoredProcedure;
                            cmdjbs.CommandText = @"FindPolisCCGetBillSeq";
                            cmdjbs.Parameters.Add(new MySqlParameter("@NoPolis", MySqlDbType.VarChar) { Value = NoPolis });
                            using (var rd = cmdjbs.ExecuteReader())
                            {
                                while (rd.Read())
                                {
                                    PolicyID = Convert.ToInt32(rd["policy_id"]);
                                    BillingID = Convert.ToInt32(rd["BillingID"]);
                                    recurring_seq = Convert.ToInt32(rd["recurring_seq"]);
                                }

                                if (PolicyID < 1 || BillingID < 1 || recurring_seq < 1)
                                {
                                    throw new Exception("Polis tidak ditemukan,mungkin billingnya tidak dalam status download atau terdapat kesalahan pada data textfile...");
                                }
                            }
                            // insert transaction bank JBS approve atw reject
                            cmdjbs.Parameters.Clear();
                            cmdjbs.CommandType = CommandType.StoredProcedure;
                            cmdjbs.CommandText = @"InsertTransactionBank;";
                            cmdjbs.Parameters.Add(new MySqlParameter("@FileName", MySqlDbType.VarChar) { Value = xFileName });
                            cmdjbs.Parameters.Add(new MySqlParameter("@Trancode", MySqlDbType.VarChar) { Value = "megaoffus" }); // hardCode MegaOnUs CC
                            cmdjbs.Parameters.Add(new MySqlParameter("@IsApprove", MySqlDbType.Bit) { Value = false });
                            cmdjbs.Parameters.Add(new MySqlParameter("@policyID", MySqlDbType.VarChar) { Value = PolicyID });
                            cmdjbs.Parameters.Add(new MySqlParameter("@Seq", MySqlDbType.VarChar) { Value = recurring_seq });
                            cmdjbs.Parameters.Add(new MySqlParameter("@IDBill", MySqlDbType.VarChar) { Value = BillingID });
                            cmdjbs.Parameters.Add(new MySqlParameter("@approvalCode", MySqlDbType.VarChar) { Value = AppCode });
                            cmdjbs.Parameters.Add(new MySqlParameter("@BankID", MySqlDbType.Int32) { Value = 0 }); // bukan BCA (boleh bankCode asli asal jgn 1)
                            cmdjbs.Parameters.Add(new MySqlParameter("@ErrCode", MySqlDbType.VarChar) { Value = Description });
                            var uid = cmdjbs.ExecuteScalar().ToString();

                            cmdjbs.CommandType = CommandType.Text;
                            cmdjbs.Parameters.Clear();
                            cmdjbs.CommandText = @"UPDATE `billing` SET IsDownload=0 WHERE `BillingID`=@billid";
                            cmdjbs.Parameters.Add(new MySqlParameter("@billid", MySqlDbType.Int32) { Value = BillingID });
                            cmdjbs.ExecuteNonQuery();
                            tranjbs.Commit();
                        }
                    }
                    catch (Exception ex)
                    {
                        tranjbs.Rollback();
                        cmdjbs.CommandType = CommandType.Text;
                        cmdjbs.Parameters.Clear();
                        cmdjbs.CommandText = @"INSERT INTO `log_error_upload_result`(TranCode,line,FileName,exceptionApp)
                                            SELECT @TranCode,@line,@FileName,@exceptionApp";
                        cmdjbs.Parameters.Add(new MySqlParameter("@TranCode", MySqlDbType.VarChar) { Value = "megaoffus" });
                        cmdjbs.Parameters.Add(new MySqlParameter("@line", MySqlDbType.Int32) { Value = row });
                        cmdjbs.Parameters.Add(new MySqlParameter("@FileName", MySqlDbType.VarChar) { Value = xFileName });
                        cmdjbs.Parameters.Add(new MySqlParameter("@exceptionApp", MySqlDbType.VarChar) { Value = "S2, " + ex.Message.Substring(0, ex.Message.Length < 255 ? ex.Message.Length : 253) });
                        cmdjbs.ExecuteNonQuery();
                    }
                    finally
                    {
                        con.Dispose();
                        con.Close();
                    }
                    PolicyID = -1;
                    BillingID = -1;
                    recurring_seq = -1;
                    AppCode = "";
                    Description = "";
                }

                //hssfwb.Dispose();
                file.Close();
            }
        }
        public static void resultBNIcc(FileInfo FileName)
        {
            
            HSSFWorkbook hssfwb = new HSSFWorkbook();
            string NoPolis="", AppCode = "",Description="";
            bool isApprove = false;

            int PolicyID=-1,BillingID=-1,recurring_seq=-1, CycleDate=0;
            DateTime DueDatePre = new DateTime(2000, 1, 1), BillDate = new DateTime(2000, 1, 1);
            decimal BillAmount=0;
            string Period="", CCno="", CCexp="", ccName="", addr="", telp="";

            string xFileName = Path.GetFileNameWithoutExtension(FileName.FullName).ToLower() +
                Path.GetRandomFileName().Replace(".", "").Substring(0, 8).ToLower() + ".xls";
            using (FileStream file = new FileStream(FileName.FullName, FileMode.Open, FileAccess.Read))
            {
                hssfwb = new HSSFWorkbook(file);
                ISheet sheet = hssfwb.GetSheetAt(0);

                MySqlConnection con = new MySqlConnection(constring);
                MySqlConnection conLife21 = new MySqlConnection(con21);
                MySqlCommand cmdjbs;
                MySqlCommand cmd21;
                MySqlTransaction tranjbs;
                MySqlTransaction tran21;

                
                
                int row=0 ;
                for (row=1; row <= sheet.LastRowNum; row++) // mulai dari baris 2
                {
                    con.Open();
                    conLife21.Open();
                    // transaction for JBS
                    tranjbs = con.BeginTransaction();
                    cmdjbs = con.CreateCommand();
                    cmdjbs.Transaction = tranjbs;

                    // transaction for Life21
                    tran21 = conLife21.BeginTransaction();
                    cmd21 = conLife21.CreateCommand();
                    cmd21.Transaction = tran21;
                    try
                    {
                        if (sheet.GetRow(row) != null)
                        {
                            // Jika cell null karena beda cell null dengan cell empty
                            if ((sheet.GetRow(row).GetCell(6) == null) && (sheet.GetRow(row).GetCell(8) == null) && (sheet.GetRow(row).GetCell(9) == null)) continue;

                            NoPolis = Convert.ToString(sheet.GetRow(row).GetCell(6));
                            if (NoPolis == "") continue; // menghindari nopolis kosong
                            AppCode =Convert.ToString( sheet.GetRow(row).GetCell(8));
                            Description = Convert.ToString(sheet.GetRow(row).GetCell(9));
                            isApprove = (AppCode == "" ? false : true);

                            // Ambil data polis billing yang akan di update
                            cmdjbs.Parameters.Clear();
                            cmdjbs.CommandType = CommandType.StoredProcedure;
                            cmdjbs.CommandText = @"FindPolisCCGetBillSeq";
                            cmdjbs.Parameters.Add(new MySqlParameter("@NoPolis", MySqlDbType.VarChar) { Value = NoPolis });
                            using (var rd = cmdjbs.ExecuteReader())
                            {
                                while (rd.Read())
                                {
                                    PolicyID = Convert.ToInt32(rd["policy_id"]);
                                    BillingID = Convert.ToInt32(rd["BillingID"]);
                                    recurring_seq = Convert.ToInt32(rd["recurring_seq"]);
                                    BillDate = Convert.ToDateTime(rd["BillingDate"]);
                                    DueDatePre = Convert.ToDateTime(rd["due_dt_pre"]);
                                    BillAmount = Convert.ToDecimal(rd["TotalAmount"]);

                                    Period = rd["PeriodeBilling"].ToString();
                                    CycleDate = Convert.ToInt32(rd["cycleDate"]);
                                    CCno = rd["cc_no"].ToString();
                                    CCexp = rd["cc_expiry"].ToString();
                                    ccName = rd["cc_name"].ToString();
                                    addr = rd["cc_address"].ToString();
                                    telp = rd["cc_telephone"].ToString();
                                }

                                if (PolicyID < 1 || BillingID < 1 || recurring_seq < 1)
                                {
                                    throw new Exception("Polis tidak ditemukan,mungkin billingnya tidak dalam status download atau terdapat kesalahan pada data textfile...");
                                }
                            }

                            // insert transaction bank JBS approve atw reject
                            cmdjbs.Parameters.Clear();
                            cmdjbs.CommandType = CommandType.StoredProcedure;
                            cmdjbs.CommandText = @"InsertTransactionBank;";
                            cmdjbs.Parameters.Add(new MySqlParameter("@FileName", MySqlDbType.VarChar) { Value = xFileName });
                            cmdjbs.Parameters.Add(new MySqlParameter("@Trancode", MySqlDbType.VarChar) { Value = "bnicc" }); // hardCode BNI CC
                            cmdjbs.Parameters.Add(new MySqlParameter("@IsApprove", MySqlDbType.Bit) { Value = isApprove });
                            cmdjbs.Parameters.Add(new MySqlParameter("@policyID", MySqlDbType.VarChar) { Value = PolicyID });
                            cmdjbs.Parameters.Add(new MySqlParameter("@Seq", MySqlDbType.VarChar) { Value = recurring_seq });
                            cmdjbs.Parameters.Add(new MySqlParameter("@IDBill", MySqlDbType.VarChar) { Value = BillingID });
                            cmdjbs.Parameters.Add(new MySqlParameter("@approvalCode", MySqlDbType.VarChar) { Value = AppCode });
                            cmdjbs.Parameters.Add(new MySqlParameter("@BankID", MySqlDbType.Int32) { Value = 0 }); // bukan BCA (boleh bankCode asli asal jgn 1)
                            cmdjbs.Parameters.Add(new MySqlParameter("@ErrCode", MySqlDbType.VarChar) { Value = Description });
                            var uid = cmdjbs.ExecuteScalar().ToString();

                            if (isApprove) // jika transaksi d approve bank, ada flag approve di file
                            {// ============================ Proses Insert Received ===========================
                                cmd21.Parameters.Clear();
                                cmd21.CommandType = CommandType.StoredProcedure;
                                cmd21.CommandText = @"ReceiptInsert";
                                cmd21.Parameters.Add(new MySqlParameter("@BillingDate", MySqlDbType.Date) { Value = BillDate });
                                cmd21.Parameters.Add(new MySqlParameter("@policy_id", MySqlDbType.Int32) { Value = PolicyID });
                                cmd21.Parameters.Add(new MySqlParameter("@receipt_amount", MySqlDbType.Decimal) { Value = BillAmount });
                                cmd21.Parameters.Add(new MySqlParameter("@Source_download", MySqlDbType.VarChar) { Value = "CC" });
                                cmd21.Parameters.Add(new MySqlParameter("@recurring_seq", MySqlDbType.Int32) { Value = recurring_seq });
                                cmd21.Parameters.Add(new MySqlParameter("@bank_acc_id", MySqlDbType.Int32) { Value = 3 }); // BankCode BNI
                                cmd21.Parameters.Add(new MySqlParameter("@due_dt_pre", MySqlDbType.Date) { Value = DueDatePre });
                                var receiptID = cmd21.ExecuteScalar().ToString();

                                // ============================ Proses Insert Pilis CC Transaction Life21 ===========================
                                cmd21.Parameters.Clear();
                                cmd21.CommandType = CommandType.StoredProcedure;
                                cmd21.CommandText = @"InsertPolistransCC";
                                cmd21.Parameters.Add(new MySqlParameter("@PolisID", MySqlDbType.Int32) { Value = PolicyID });
                                cmd21.Parameters.Add(new MySqlParameter("@Transdate", MySqlDbType.Date) { Value = BillDate });
                                cmd21.Parameters.Add(new MySqlParameter("@Seq", MySqlDbType.Int32) { Value = recurring_seq });
                                cmd21.Parameters.Add(new MySqlParameter("@Amount", MySqlDbType.Decimal) { Value = BillAmount });
                                cmd21.Parameters.Add(new MySqlParameter("@DueDatePre", MySqlDbType.Date) { Value = DueDatePre });
                                cmd21.Parameters.Add(new MySqlParameter("@Period", MySqlDbType.VarChar) { Value = Period });
                                cmd21.Parameters.Add(new MySqlParameter("@CycleDate", MySqlDbType.Int32) { Value = CycleDate });
                                cmd21.Parameters.Add(new MySqlParameter("@BankID", MySqlDbType.Int32) { Value = 3 }); // BankCode BNI
                                cmd21.Parameters.Add(new MySqlParameter("@CCno", MySqlDbType.VarChar) { Value = CCno });
                                cmd21.Parameters.Add(new MySqlParameter("@CCExpiry", MySqlDbType.VarChar) { Value = CCexp });
                                cmd21.Parameters.Add(new MySqlParameter("@CCName", MySqlDbType.VarChar) { Value = ccName });
                                cmd21.Parameters.Add(new MySqlParameter("@CCAddrs", MySqlDbType.VarChar) { Value = addr });
                                cmd21.Parameters.Add(new MySqlParameter("@CCtelp", MySqlDbType.VarChar) { Value = telp });
                                cmd21.Parameters.Add(new MySqlParameter("@receiptID", MySqlDbType.Int32) { Value = receiptID });
                                var CCTransID = cmd21.ExecuteScalar().ToString();

                                // Update table billing
                                cmdjbs.Parameters.Clear();
                                cmdjbs.CommandType = CommandType.Text;
                                cmdjbs.CommandText = @"UPDATE `billing` SET `IsDownload`=0,
			                                                `IsClosed`=1,
			                                                `status_billing`='P',
			                                                `status_billing_dateUpdate`=@tgl,
			                                                `paid_date`=@billDate,
                                                            Life21TranID=@TransactionID,
			                                                `ReceiptID`=@receiptID,
			                                                `PaymentTransactionID`=@uid
		                                                WHERE `BillingID`=@idBill;";
                                cmdjbs.Parameters.Add(new MySqlParameter("@tgl", MySqlDbType.DateTime) { Value = DateTime.Now });
                                cmdjbs.Parameters.Add(new MySqlParameter("@billDate", MySqlDbType.DateTime) { Value = BillDate });
                                cmdjbs.Parameters.Add(new MySqlParameter("@TransactionID", MySqlDbType.Int32) { Value = CCTransID });
                                cmdjbs.Parameters.Add(new MySqlParameter("@receiptID", MySqlDbType.Int32) { Value = receiptID });
                                cmdjbs.Parameters.Add(new MySqlParameter("@uid", MySqlDbType.VarChar) { Value = uid });
                                cmdjbs.Parameters.Add(new MySqlParameter("@idBill", MySqlDbType.Int32) { Value = BillingID });
                                cmdjbs.ExecuteNonQuery();

                                // Update Polis Last Transaction
                                cmdjbs.Parameters.Clear();
                                cmdjbs.CommandType = CommandType.Text;
                                cmdjbs.CommandText = @"UPDATE `policy_last_trans` AS pt
		                                                INNER JOIN `billing` AS bx ON bx.policy_id=pt.policy_Id
			                                                SET pt.BillingID=bx.BillingID,
			                                                pt.recurring_seq=bx.recurring_seq,
			                                                pt.due_dt_pre=bx.due_dt_pre,
			                                                pt.source=bx.Source_download,
			                                                pt.receipt_id=bx.`ReceiptID`,
			                                                pt.receipt_date=bx.BillingDate,
			                                                pt.bank_id=bx.BankIdDownload
		                                                WHERE pt.policy_Id=@policyID AND bx.BillingID=@idBill;";
                                cmdjbs.Parameters.Add(new MySqlParameter("@policyID", MySqlDbType.Int32) { Value = PolicyID });
                                cmdjbs.Parameters.Add(new MySqlParameter("@idBill", MySqlDbType.Int32) { Value = BillingID });
                                cmdjbs.ExecuteNonQuery();

                            }
                            else // jika transaksi d reject bank
                            {//billing hanya ganti flag download, kolom lain tetap sbg status terakhir
                                cmdjbs.CommandType = CommandType.Text;
                                cmdjbs.Parameters.Clear();
                                cmdjbs.CommandText = @"UPDATE `billing` SET IsDownload=0 WHERE `BillingID`=@billid";
                                cmdjbs.Parameters.Add(new MySqlParameter("@billid", MySqlDbType.Int32) { Value = BillingID });
                                cmdjbs.ExecuteNonQuery();
                            }

                            AppCode = "";
                            Description = "";
                            isApprove = false;
                        }
                        tranjbs.Commit();
                        tran21.Commit();
                    }
                    catch (Exception ex)
                    {
                        tranjbs.Rollback();
                        tran21.Rollback();
                        cmdjbs.CommandType = CommandType.Text;
                        cmdjbs.Parameters.Clear();
                        cmdjbs.CommandText = @"INSERT INTO `log_error_upload_result`(TranCode,line,FileName,exceptionApp)
                                            SELECT @TranCode,@line,@FileName,@exceptionApp";
                        cmdjbs.Parameters.Add(new MySqlParameter("@TranCode", MySqlDbType.VarChar) { Value = "bnicc" });
                        cmdjbs.Parameters.Add(new MySqlParameter("@line", MySqlDbType.Int32) { Value = row });
                        cmdjbs.Parameters.Add(new MySqlParameter("@FileName", MySqlDbType.VarChar) { Value = xFileName });
                        cmdjbs.Parameters.Add(new MySqlParameter("@exceptionApp", MySqlDbType.VarChar) { Value = ex.Message.Substring(0, ex.Message.Length < 255 ? ex.Message.Length : 253) });
                        cmdjbs.ExecuteNonQuery();
                    }
                    finally
                    {
                        con.Dispose();
                        con.Close();
                        conLife21.Close();
                    }
                }
                //hssfwb.Dispose();
                file.Close();
            }
        }
    }
}

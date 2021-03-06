﻿using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Data;
using System.IO;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Globalization;
using Gizmox.Controls;
using JDataEngine;
using JurisAuthenticator;
using JurisUtilityBase.Properties;
using System.Data.OleDb;
using System.Diagnostics;
using Microsoft.Office.Interop.Excel;
using System.Runtime.ExceptionServices;

namespace JurisUtilityBase
{
    public partial class UtilityBaseMain : Form
    {
        #region Private  members

        private JurisUtility _jurisUtility;

        #endregion

        #region Public properties

        public string CompanyCode { get; set; }

        public string JurisDbName { get; set; }

        public string JBillsDbName { get; set; }

        string clientFile = "";

        string matterFile = "";

        string origFile = "";

        string addyFile = "";

        #endregion

        #region Constructor

        public UtilityBaseMain()
        {
            InitializeComponent();
            _jurisUtility = new JurisUtility();
        }

        #endregion

        #region Public methods

        public void LoadCompanies()
        {
            var companies = _jurisUtility.Companies.Cast<object>().Cast<Instance>().ToList();
//            listBoxCompanies.SelectedIndexChanged -= listBoxCompanies_SelectedIndexChanged;
            listBoxCompanies.ValueMember = "Code";
            listBoxCompanies.DisplayMember = "Key";
            listBoxCompanies.DataSource = companies;
//            listBoxCompanies.SelectedIndexChanged += listBoxCompanies_SelectedIndexChanged;
            var defaultCompany = companies.FirstOrDefault(c => c.Default == Instance.JurisDefaultCompany.jdcJuris);
            if (companies.Count > 0)
            {
                listBoxCompanies.SelectedItem = defaultCompany ?? companies[0];
            }
        }

        #endregion

        #region MainForm events

        private void Form1_Load(object sender, EventArgs e)
        {
        }

        private void listBoxCompanies_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (_jurisUtility.DbOpen)
            {
                _jurisUtility.CloseDatabase();
            }
            CompanyCode = "Company" + listBoxCompanies.SelectedValue;
            _jurisUtility.SetInstance(CompanyCode);
            JurisDbName = _jurisUtility.Company.DatabaseName;
            JBillsDbName = "JBills" + _jurisUtility.Company.Code;
            _jurisUtility.OpenDatabase();
            if (_jurisUtility.DbOpen)
            {
                ///GetFieldLengths();
            }

        }



        #endregion



        private void DoDaFix()
        {






            UpdateStatus("LedgerHistory Updated.", 1, 1);

            MessageBox.Show("The process is complete", "Confirmation", MessageBoxButtons.OK, MessageBoxIcon.None);
            
            clientFile = "";
            addyFile = "";
            matterFile = "";
            origFile = "";
        }


        private void processClients()
        {
            string sysparam = "  select SpTxtValue from sysparam where SpName = 'FldClient'";
            DataSet dds = _jurisUtility.RecordsetFromSQL(sysparam);
            int clientLength = 5;
            string cell = "";
            if (dds != null && dds.Tables.Count > 0)
            {
                cell = dds.Tables[0].Rows[0].ToString();
                
                foreach (DataRow dr in dds.Tables[0].Rows)
                {
                    cell = dr[0].ToString();
                }
                
            }
            string[] test = cell.Split(',');
            clientLength = Convert.ToInt32(test[2]);

            string line;
            int counter = 0;
            List<string> clientItems = new List<string>();
            System.IO.StreamReader file =  new System.IO.StreamReader(clientFile);
            string errorSQL = "";
            while ((line = file.ReadLine()) != null)
            {
                try
                {
                    if (counter > 0)
                    {
                    errorSQL = line;
                        clientItems = line.Split('\t').ToList();
                            string sql = "Insert into Client(CliSysNbr,CliCode,CliNickName,CliReportingName,CliSourceOfBusiness,CliPhoneNbr,CliFaxNbr,CliContactName,CliDateOpened,CliOfficeCode,CliBillingAtty,CliPracticeClass, "
                               + " CliFeeSch,CliTaskCodeXref,CliExpSch,CliExpCodeXref,CliBillFormat,CliBillAgreeCode,CliFlatFeeIncExp,CliRetainerType,CliExpFreqCode,CliFeeFreqCode,CliBillMonth,CliBillCycle, "
                               + " CliExpThreshold,CliFeeThreshold,CliInterestPcnt,CliInterestDays,CliDiscountOption,CliDiscountPcnt,CliSurchargeOption,CliSurchargePcnt,CliTax1Exempt,CliTax2Exempt,CliTax3Exempt,CliBudgetOption,CliReqPhaseOnTrans, "
                               + " CliReqTaskCdOnTime,CliReqActyCdOnTime,CliReqTaskCdOnExp,CliPrimaryAddr,CliType,CliEditFormat,CliThresholdOption,CliRespAtty,CliBillingField01,CliBillingField02,CliBillingField03,CliBillingField04,CliBillingField05, "
                               + " CliBillingField06,CliBillingField07,CliBillingField08,CliBillingField09,CliBillingField10,CliBillingField11,CliBillingField12,CliBillingField13,CliBillingField14,CliBillingField15,CliBillingField16,CliBillingField17,CliBillingField18,CliBillingField19, "
                               + " CliBillingField20,CliCTerms,CliCStatus,CliCStatus2)  "
                               + " values( case when (select max(clisysnbr) from client) is null then 1 else ((select max(clisysnbr) from client) + 1) end, right('000000000000' + '" + clientItems[0].Trim() + "', " + clientLength + "), left('" + clientItems[1].Replace("\"", "").Replace("'", "''") + "', 30), left('" + clientItems[1].Replace("\"", "").Replace("'", "''") + "', 30), "
                               + " '', left('" + clientItems[11] + "', 20), left('" + clientItems[12] + "', 20), " +
                               "left('" + clientItems[13].Replace("\"", "").Replace("'", "''") + "', 30), '" + clientItems[6] + "', '" + clientItems[2] + "', "
                             + " (select empsysnbr from employee where empid = '" + clientItems[5] + "'), "
                             + "'" + clientItems[3] + "', 'STDR',null,'STDR',null, "
                               + " 'BF01','"
                               + clientItems[8] + "','N','', '" + clientItems[10] + "', '" + clientItems[9] + "' ,1,1,0.00,0.00,0.0000,30,0,0.0000, "
                               + " 0,0.0000,'N','N','N',0,'N','N','N','N',null,0,'BF01', "
                               + " 0,null,'" + clientItems[14].Replace("\"", "").Replace("'", "''").Replace("|", "\r\n") + "','','','','','','','','','','','', "
                               + " '','','','','','','','',0,0,'')";

                            errorSQL = sql;


                            _jurisUtility.ExecuteNonQuery(0, sql);
                        
                    }
                    counter++;
                    clientItems.Clear();
                    errorSQL = "";
                }
                catch (Exception ex1)
                {
                    MessageBox.Show(ex1.Message + "\r\n" + errorSQL);


                }
            }

            file.Close();

            string SQL = " delete from documenttree where dtdocclass = 4200 and dtparentid <> 2";
            _jurisUtility.ExecuteNonQuery(0, SQL);

            SQL = "Insert into DocumentTree(dtdocid, dtsystemcreated, dtdocclass,dtdoctype,  dtparentid, dttitle, dtkeyl) "
                       + " select(select max(dtdocid)  from documenttree) + rank() Over(order by clisysnbr) as DTID, 'Y',4200,'R', 22, Clireportingname, Clisysnbr "
                       + " from Client ";
            _jurisUtility.ExecuteNonQuery(0, SQL);

            SQL = " Update sysparam set spnbrvalue=(select max(dtdocid) from documenttree) where spname='LastSysNbrDocTree'";
            _jurisUtility.ExecuteNonQuery(0, SQL);

            SQL = " update sysparam set spnbrvalue = (select max(CliSysNbr) from client) where spname = 'LastSysNbrClient'";
            _jurisUtility.ExecuteNonQuery(0, SQL);

        }

        private void processAddresses()
        {
            string sysparam = "  select SpTxtValue from sysparam where SpName = 'FldClient'";
            DataSet dds = _jurisUtility.RecordsetFromSQL(sysparam);
            int clientLength = 5;
            string cell = "";
            if (dds != null && dds.Tables.Count > 0)
            {
                cell = dds.Tables[0].Rows[0].ToString();

                foreach (DataRow dr in dds.Tables[0].Rows)
                {
                    cell = dr[0].ToString();
                }

            }
            string[] test = cell.Split(',');
            clientLength = Convert.ToInt32(test[2]);
            string line;
            int counter = 0;
            List<string> clientItems = new List<string>();
            System.IO.StreamReader file = new System.IO.StreamReader(addyFile);
            string errorSQL = "";
            while ((line = file.ReadLine()) != null)
            {
                try
                {
                    Random rr = new Random();
                    if (counter > 0)
                    {
                        clientItems = line.Split('\t').ToList();
                        
                        string sql = "Insert into BillingAddress(BilAdrSysNbr, BilAdrCliNbr, BilAdrUsageFlg, BilAdrNickName, BilAdrPhone, BilAdrFax, BilAdrContact, BilAdrName, BilAdrAddress, BilAdrCity, BilAdrState, BilAdrZip, BilAdrCountry, BilAdrType, BilAdrEmail) " +
                            " values (case when(select max(biladrsysnbr) from billingaddress) is null then 1 else ((select max(biladrsysnbr) from billingaddress) +1) end, case when (select clisysnbr from client where dbo.jfn_FormatClientCode(clicode) = right('000000000000' + '" + clientItems[0].Trim() + "', " + clientLength + ")) is null then 99999 else (select clisysnbr from client where dbo.jfn_FormatClientCode(clicode) = right('000000000000' + '" + clientItems[0] + "', " + clientLength + ")) end, " +
                            " 'C', left('" + clientItems[1].Trim() + "' + '-' + '" + clientItems[2].Replace("\"", "").Replace("'", "''") + "', 25), left('" + clientItems[3].Replace("\"", "").Replace("'", "''") + "', 20), "
                             + "         left('" + clientItems[4].Replace("\"", "").Replace("'", "''") + "', 20), left('" + clientItems[5].Replace("\"", "").Replace("'", "''") + "', 30), " +
                            " left('" + clientItems[6].Replace("\"", "").Replace("'", "''") + "', 50), " +
                            "case when '" + clientItems[7].Replace("\"", "").Replace("'", "''") + "' = '' then 'N/A' when '" + clientItems[8].Replace("\"", "").Replace("'", "''") + "' = '' then  left('" + clientItems[7].Replace("\"", "").Replace("'", "''") + "', 250) else left('" + clientItems[7].Replace("\"", "").Replace("'", "''") + "' + char(13) + char(10) + '" + clientItems[8].Replace("\"", "").Replace("'", "''") + "',250) end, " 
                            + " left('" + clientItems[9] + "', 20), left('" + clientItems[10] + "', 2), left('" + clientItems[11].Replace("-", "") + "', 9),left('" + clientItems[12] + "', 20), 0, left('" + clientItems[13] + "', 255) )";

                        errorSQL = sql;


                        _jurisUtility.ExecuteNonQuery(0, sql);

                        if (!_jurisUtility.error)
                         {
                        sql = "Insert into BillCopy(BilCpyBillTo,BilCpyBilAdr,BilCpyComment,BilCpyNbrOfCopies,BilCpyPrintFormat,BilCpyEmailFormat,BilCpyExportFormat,BilCpyARFormat) "
                            + " values (  (select matbillto from matter inner join client on matclinbr = clisysnbr where matcode = right('000000000000' + '" + clientItems[1].Trim() + "', 12) and dbo.jfn_formatClientCode(clicode) = right('000000000000' + '" + clientItems[0] + "', " + clientLength + "))," 
                                + "  (select max(BilAdrSysNbr) from billingaddress) ,'',1,1,0,0,0 )";


                            errorSQL = sql;
                            _jurisUtility.ExecuteNonQuery(0, sql);
                        }

                    }
                    counter++;
                    clientItems.Clear();
                    errorSQL = "";
                }
                catch (Exception ex1)
                {
                    System.Windows.Forms.Clipboard.SetText(errorSQL);
                    MessageBox.Show(ex1.Message + "\r\n" + errorSQL);


                }
            }


            file.Close();
            string SQL = "update sysparam set spnbrvalue = (select max(biladrsysnbr) from billingaddress) where spname = 'LastSysNbrBillAddress'";
            _jurisUtility.ExecuteNonQuery(0, SQL);



        }


        private void processMatters()
        {
            string sysparam = "  select SpTxtValue from sysparam where SpName = 'FldClient'";
            DataSet dds = _jurisUtility.RecordsetFromSQL(sysparam);
            int clientLength = 5;
            string cell = "";
            if (dds != null && dds.Tables.Count > 0)
            {
                cell = dds.Tables[0].Rows[0].ToString();

                foreach (DataRow dr in dds.Tables[0].Rows)
                {
                    cell = dr[0].ToString();
                }

            }
            string[] test = cell.Split(',');
            clientLength = Convert.ToInt32(test[2]);
            try
            {
                //adds temp billto so we can add the matter first then assign it the right billto
                string ss = "  insert into billto ([BillToSysNbr],[BillToCliNbr],[BillToUsageFlg] ,[BillToNickName],[BillToBillingAtty],[BillToBillFormat] ,[BillToEditFormat] ,[BillToRespAtty]) " +
                            "values(1, (select top 1 clisysnbr from client), 'M', 'TestForConv', (select empsysnbr from employee where empid = 'SMGR'), 'BF01', 'BF01', 0)";
                _jurisUtility.ExecuteNonQuery(0, ss);
                string line;
                int counter = 0;
                List<string> clientItems = new List<string>();
                System.IO.StreamReader file = new System.IO.StreamReader(matterFile);
                string errorSQL = "";
                while ((line = file.ReadLine()) != null)
                {
                    try
                    {
                        if (counter > 0)
                        {
                            

                                clientItems = line.Split('\t').ToList();
                                string sql = "Insert into Matter(MatSysNbr,MatCliNbr,MatBillTo,MatCode,MatNickName,MatReportingName,MatDescription,MatRemarks,MatPhoneNbr,MatFaxNbr,MatContactName,MatDateOpened,MatStatusFlag,MatLockFlag, "
                                   + "  MatDateClosed,MatOfficeCode,MatPracticeClass,MatFeeSch,MatTaskCodeXref,MatExpSch,MatExpCodeXref,MatQuickAction,MatBillAgreeCode,MatFlatFeeIncExp,MatRetainerType,MatFltFeeOrRetainer,MatExpFreqCode,"
                                  + "   MatFeeFreqCode,MatBillMonth,MatBillCycle,MatExpThreshold,MatFeeThreshold,MatInterestPcnt,MatInterestDays,MatDiscountOption,MatDiscountPcnt,MatSurchargeOption,MatSurchargePcnt,MatSplitMethod,MatSplitThreshold,"
                                   + "  MatSplitPriorAmtBld,MatBudgetOption,MatBudgetPhase,MatReqPhaseOnTrans,MatReqTaskCdOnTime,MatReqActyCdOnTime,MatReqTaskCdOnExp,MatTax1Exempt,MatTax2Exempt,MatTax3Exempt,MatDateLastWork,MatDateLastExp"
                                   + "  , MatDateLastBill,MatDateLastStmt,MatDateLastPaymt,MatLastPaymtAmt,MatARLastBill,MatPaySinceLastBill,MatAdjSinceLastBill,MatPPDBalance,MatVisionAddr,MatThresholdOption,MatType,MatBillingField01,MatBillingField02,"
                                  + "   MatBillingField03,MatBillingField04,MatBillingField05,MatBillingField06,MatBillingField07,MatBillingField08,MatBillingField09,MatBillingField10,MatBillingField11,MatBillingField12,MatBillingField13,MatBillingField14,MatBillingField15,MatBillingField16,"
                                   + "  MatBillingField17,MatBillingField18,MatBillingField19,MatBillingField20,MatCTerms,MatCStatus,MatCStatus2) "
                                   + "     values( case when (select max(MatSysNbr) from matter) is null then 1 else ((select max(MatSysNbr) from matter) + 1) end, case when (select clisysnbr from client where dbo.jfn_FormatClientCode(clicode) = right('000000000000' + '" + clientItems[0].Trim() + "', " + clientLength + ")) is null then 99999 else (select clisysnbr from client where dbo.jfn_FormatClientCode(clicode) = right('000000000000' + '" + clientItems[0] + "', " + clientLength + ")) end, 1,  "
                                   + "       right('000000000000'+ '"  + clientItems[1].Trim() + "', 12), left('" + clientItems[2].Replace("\"", "").Replace("'", "''") + "', 30), left('" + clientItems[2].Replace("\"", "").Replace("'", "''") + "', 30), left('" + clientItems[3].Replace("\"", "").Replace("'", "''") + "', 254), '', left('" + clientItems[14].Replace("\"", "").Replace("'", "''") + "', 20), "
                                 + "         left('" + clientItems[15].Replace("\"", "").Replace("'", "''") + "', 20), left('" + clientItems[16].Replace("\"", "").Replace("'", "''") + "', 30), '" + clientItems[10] + "',case when '" + clientItems[11] + "' = '1/1/1900' then 'O' else 'C' end ,'0', "
                                 + " '" + clientItems[11] + "','" + clientItems[4] + "','" + clientItems[5] + "','STDR',null,'STDR',null,0,'" + clientItems[8] + "', "
                                  + "        'N', '', case when '" + clientItems[9] + "' = '' then 0 else " + clientItems[9] + " end ,'" + clientItems[13] + "','" + clientItems[12] + "',1,1,0.00,0.00,0.0000,0,0,0.0000, "
                                   + "       0,0.0000,0,0.000,0.00,0,0,'N', 'N', 'N', 'N', 'N',"
                                + "          'N','N','01/01/1900','01/01/1900','01/01/1900','01/01/1900','01/01/1900',0,0,0,0,0,0, "
                                 + "         0, 0,'" + clientItems[17].Replace("\"", "").Replace("'", "''").Replace("|", "\r\n") + "','" + clientItems[18].Replace("\"", "").Replace("'", "''").Replace("|", "\r\n") + "','" + clientItems[19].Replace("\"", "").Replace("'", "''").Replace("|", "\r\n") + "','','" + clientItems[20].Replace("\"", "").Replace("'", "''").Replace("|", "\r\n") + "','" + clientItems[21].Replace("\"", "").Replace("'", "''").Replace("|", "\r\n") + "','','','','','','', "
                                + "          '','','','','','','','',null, null, null)";

                                _jurisUtility.ExecuteNonQuery(0, sql);

                                if (!_jurisUtility.error)
                                {
                                    sql = "Insert into BillTo (BillToSysNbr,BillToCliNbr,BillToUsageFlg,BillToNickName,BillToBillingAtty,BillToBillFormat,BillToEditFormat,BillToRespAtty) " +
                                        "values (((select max(BillToSysNbr) from billto) + 1), " +
                                        "case when (select clisysnbr from client where dbo.jfn_FormatClientCode(clicode) = right('000000000000' + '" + clientItems[0] + "', " + clientLength + ")) is null then 2 else (select clisysnbr from client where dbo.jfn_FormatClientCode(clicode) = right('000000000000' + '" + clientItems[0] + "', " + clientLength + ")) end, 'M', left('" + clientItems[0] + clientItems[1] + "', 30), (select empsysnbr from employee where empid = '" + clientItems[6].Trim() + "'), 'BF01', 'BF01', null)";
                                    errorSQL = sql;

                                    _jurisUtility.ExecuteNonQuery(0, sql);
                                }

                                if (!_jurisUtility.error)
                                {
                                    sql = "update matter set matbillto = (select max(billtosysnbr) from billto) where matsysnbr = (select max(matsysnbr) from  matter)";
                                    errorSQL = sql;
                                    _jurisUtility.ExecuteNonQuery(0, sql);
                                }
                            

                        }
                        counter++;
                        clientItems.Clear();
                        errorSQL = "";
                    }
                    catch (Exception ex1)
                    {
                        
                        MessageBox.Show(ex1.Message + "\r\n" + errorSQL);


                    }
                }

                file.Close();

                string SQL = "delete from billto where billtosysnbr = 1";
                //_jurisUtility.ExecuteNonQuery(0, SQL);

                SQL = "update sysparam set spnbrvalue = (select max(billtosysnbr) from billto) where spname = 'LastSysNbrBillTo'";
                _jurisUtility.ExecuteNonQuery(0, SQL);

                SQL = "update sysparam set spnbrvalue = (select max(matsysnbr) from matter) where spname = 'LastSysNbrMatter'";
                _jurisUtility.ExecuteNonQuery(0, SQL);
            }
            catch (Exception ex2)

            {
                MessageBox.Show(ex2.Message);

            }
        }

        private void processOrig(int flag)
        {
            if (flag == 1) //clients
            {
                string line;
                int counter = 0;
                List<string> clientItems = new List<string>();
                System.IO.StreamReader file = new System.IO.StreamReader(origFile);
                string errorSQL = "";
                while ((line = file.ReadLine()) != null)
                {
                    try
                    {
                        if (counter > 0)
                        {
                            clientItems = line.Split('\t').ToList();

                        string sql = "update CliOrigAtty set COrigAtty = (select empsysnbr from employee where empid = '" + clientItems[1].Trim() + "') where COrigCli = (select clisysnbr from client where dbo.jfn_FormatClientCode(clicode) = right('000000000000' + '" + clientItems[0] + "', 4))";
                            errorSQL = sql;


                            _jurisUtility.ExecuteNonQuery(0, sql);



                        }
                        counter++;
                        clientItems.Clear();
                        errorSQL = "";
                    }
                    catch (Exception ex1)
                    {
                        System.Windows.Forms.Clipboard.SetText(errorSQL);
                        MessageBox.Show(ex1.Message + "\r\n" + errorSQL);


                    }
                }


                file.Close();

            }
            else
            {
                string line;
                int counter = 0;
                List<string> clientItems = new List<string>();
                System.IO.StreamReader file = new System.IO.StreamReader(origFile);
                string errorSQL = "";
                while ((line = file.ReadLine()) != null)
                {
                    try
                    {
                        Random rr = new Random();
                        if (counter > 0)
                        {
                            clientItems = line.Split('\t').ToList();

                            string sql = "update MatOrigAtty set MOrigAtty = (select empsysnbr from employee where empid = '" + clientItems[2].Trim() + "') where MOrigMat = (select matsysnbr from matter inner join client on matclinbr = clisysnbr where matcode = right('000000000000' + '" + clientItems[1].Trim() + "', 12) and dbo.jfn_formatClientCode(clicode) = right('000000000000' + '" + clientItems[0].Trim() + "', 4))";
                            errorSQL = sql;


                            _jurisUtility.ExecuteNonQuery(0, sql);



                        }
                        counter++;
                        clientItems.Clear();
                        errorSQL = "";
                    }
                    catch (Exception ex1)
                    {
                        System.Windows.Forms.Clipboard.SetText(errorSQL);
                        MessageBox.Show(ex1.Message + "\r\n" + errorSQL);


                    }
                }


                file.Close();
            }

        }


        private bool VerifyFirmName()
        {
            //    Dim SQL     As String
            //    Dim rsDB    As ADODB.Recordset
            //
            //    SQL = "SELECT CASE WHEN SpTxtValue LIKE '%firm name%' THEN 'Y' ELSE 'N' END AS Firm FROM SysParam WHERE SpName = 'FirmName'"
            //    Cmd.CommandText = SQL
            //    Set rsDB = Cmd.Execute
            //
            //    If rsDB!Firm = "Y" Then
            return true;
            //    Else
            //        VerifyFirmName = False
            //    End If

        }

        private bool FieldExistsInRS(DataSet ds, string fieldName)
        {

            foreach (DataColumn column in ds.Tables[0].Columns)
            {
                if (column.ColumnName.Equals(fieldName, StringComparison.OrdinalIgnoreCase))
                    return true;
            }
            return false;
        }


        private static bool IsDate(String date)
        {
            try
            {
                DateTime dt = DateTime.Parse(date);
                return true;
            }
            catch
            {
                return false;
            }
        }

        private static bool IsNumeric(object Expression)
        {
            double retNum;

            bool isNum = Double.TryParse(Convert.ToString(Expression), System.Globalization.NumberStyles.Any, System.Globalization.NumberFormatInfo.InvariantInfo, out retNum);
            return isNum; 
        }

        private void WriteLog(string comment)
        {
            var sql =
                string.Format("Insert Into UtilityLog(ULTimeStamp,ULWkStaUser,ULComment) Values('{0}','{1}', '{2}')",
                    DateTime.Now, GetComputerAndUser(), comment);
            _jurisUtility.ExecuteNonQueryCommand(0, sql);
        }

        private string GetComputerAndUser()
        {
            var computerName = Environment.MachineName;
            var windowsIdentity = System.Security.Principal.WindowsIdentity.GetCurrent();
            var userName = (windowsIdentity != null) ? windowsIdentity.Name : "Unknown";
            return computerName + "/" + userName;
        }


        private void UpdateStatus(string status, long step, long steps)
        {
            labelCurrentStatus.Text = status;

            if (steps == 0)
            {
                progressBar.Value = 0;
                labelPercentComplete.Text = string.Empty;
            }
            else
            {
                double pctLong = Math.Round(((double)step/steps)*100.0);
                int percentage = (int)Math.Round(pctLong, 0);
                if ((percentage < 0) || (percentage > 100))
                {
                    progressBar.Value = 0;
                    labelPercentComplete.Text = string.Empty;
                }
                else
                {
                    progressBar.Value = percentage;
                    labelPercentComplete.Text = string.Format("{0} percent complete", percentage);
                }
            }
        }



        private void button1_Click(object sender, EventArgs e)
        {
            DoDaFix();
        }

        private void buttonReport_Click(object sender, EventArgs e)
        {

            System.Environment.Exit(0);
          
        }


        private void buttonClients_Click(object sender, EventArgs e)
        {
            OpenFileDialogOpen.Title = "Select Client Text File (tab delimited)";
            OpenFileDialogOpen.Multiselect = false;
            OpenFileDialogOpen.DefaultExt = "txt";
            OpenFileDialogOpen.Filter = "txt files (*.txt)|*.txt|All files (*.*)|*.*";
            OpenFileDialogOpen.FilterIndex = 1;
            DialogResult dr = OpenFileDialogOpen.ShowDialog();

            if (dr == DialogResult.OK)
            {
                clientFile = OpenFileDialogOpen.FileName;
                processClients();
                MessageBox.Show("The process is complete", "Confirmation", MessageBoxButtons.OK, MessageBoxIcon.None);
            }
        }

        private void buttonMatters_Click(object sender, EventArgs e)
        {
            OpenFileDialogOpen.Title = "Select Matter Text File (tab delimited)";
            OpenFileDialogOpen.Multiselect = false;
            OpenFileDialogOpen.DefaultExt = "txt";
            OpenFileDialogOpen.Filter = "txt files (*.txt)|*.txt|All files (*.*)|*.*";
            OpenFileDialogOpen.FilterIndex = 1;
            DialogResult dr = OpenFileDialogOpen.ShowDialog();

            if (dr == DialogResult.OK)
            {
                matterFile = OpenFileDialogOpen.FileName;
                //processMatters();
                processMatters();
                MessageBox.Show("The process is complete", "Confirmation", MessageBoxButtons.OK, MessageBoxIcon.None);
            }
        }

        private void buttonAddy_Click(object sender, EventArgs e)
        {
            OpenFileDialogOpen.Title = "Select Matter Address Text File (tab delimited)";
            OpenFileDialogOpen.Multiselect = false;
            OpenFileDialogOpen.DefaultExt = "txt";
            OpenFileDialogOpen.Filter = "txt files (*.txt)|*.txt|All files (*.*)|*.*";
            OpenFileDialogOpen.FilterIndex = 1;
            DialogResult dr = OpenFileDialogOpen.ShowDialog();

            if (dr == DialogResult.OK)
            {
                addyFile = OpenFileDialogOpen.FileName;
                processAddresses();
                MessageBox.Show("The process is complete", "Confirmation", MessageBoxButtons.OK, MessageBoxIcon.None);
            }
        }

        private void buttonOrig_Click(object sender, EventArgs e)
        {
            OpenFileDialogOpen.Title = "Select Client Originator Text File (tab delimited)";
            OpenFileDialogOpen.Multiselect = false;
            OpenFileDialogOpen.DefaultExt = "txt";
            OpenFileDialogOpen.Filter = "txt files (*.txt)|*.txt|All files (*.*)|*.*";
            OpenFileDialogOpen.FilterIndex = 1;
            DialogResult dr = OpenFileDialogOpen.ShowDialog();

            if (dr == DialogResult.OK)
            {
                origFile = OpenFileDialogOpen.FileName;
                processOrig(1);
                MessageBox.Show("The process is complete", "Confirmation", MessageBoxButtons.OK, MessageBoxIcon.None);
                origFile = "";
            }
        }

        private void buttonMatOrig_Click(object sender, EventArgs e)
        {
            OpenFileDialogOpen.Title = "Select Matter Originator Text File (tab delimited)";
            OpenFileDialogOpen.Multiselect = false;
            OpenFileDialogOpen.DefaultExt = "txt";
            OpenFileDialogOpen.Filter = "txt files (*.txt)|*.txt|All files (*.*)|*.*";
            OpenFileDialogOpen.FilterIndex = 1;
            DialogResult dr = OpenFileDialogOpen.ShowDialog();

            if (dr == DialogResult.OK)
            {
                origFile = OpenFileDialogOpen.FileName;
                processOrig(2);
                MessageBox.Show("The process is complete", "Confirmation", MessageBoxButtons.OK, MessageBoxIcon.None);
                origFile = "";
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            //get rid of dupes


            // " ) t on s.BilAdrCliNbr = t.BilAdrCliNbr and s.[BilAdrContact] = t.BilAdrContact and s.[BilAdrName] = t.BilAdrName and s.[BilAdrAddress] = t.BilAdrAddress and  " +
            //" s.[BilAdrCity] = t.BilAdrCity and s.[BilAdrState] = t.BilAdrState and s.[BilAdrZip] = t.BilAdrZip order by t.BilAdrCliNbr, s.BilAdrSysNbr";
           string ss = "  select s.BilAdrSysNbr, t.BilAdrCliNbr, t.BilAdrContact " + //t.*
" from[BillingAddress] s  " +
" inner join(" +
" select BilAdrCliNbr, [BilAdrContact], count(*) as qty  " +
" from[BillingAddress]  " +
" group by BilAdrCliNbr,[BilAdrContact] " +
" having count(*) > 1  " +
" ) t on s.BilAdrCliNbr = t.BilAdrCliNbr and s.[BilAdrContact] = t.BilAdrContact    " +
"  order by t.BilAdrCliNbr, s.BilAdrSysNbr, t.BilAdrContact";
            DataSet dds = _jurisUtility.RecordsetFromSQL(ss);

            //store addresses from table to list
            List<Addy> addys = new List<Addy>();
            Addy address = null;

            if (dds != null && dds.Tables.Count > 0)
            {
                foreach (DataRow dr in dds.Tables[0].Rows)
                {
                    address = new Addy();
                   // address.adr = dr[4].ToString();
                   // address.city = dr[5].ToString();
                    address.cli = Convert.ToInt32(dr[1].ToString());
                   address.contact = dr[2].ToString();
                  //  address.name = dr[3].ToString();
                  //  address.state = dr[6].ToString();
                    address.sys = Convert.ToInt32(dr[0].ToString());
                  // address.zip = dr[7].ToString();
                    addys.Add(address);
                }
            }

            //separate them by client and store first billadrsys (we will be removing the rest
            addys = addys.OrderBy(c => c.cli).ThenBy(c => c.contact).ThenBy(c => c.sys).ToList();

            int currentClient = 0;
            billAdr bb = null;
            int currentGoodSys = 0;
            string currentContact = "";
            List<billAdr> adrList = new List<billAdr>();
            foreach (Addy aa in addys)
            {
                bb = new billAdr();
                if (aa.cli == currentClient && aa.contact.Equals(currentContact))
                {
                    bb.goodSys = currentGoodSys;
                    bb.cli = currentClient;
                    bb.isBad = true;
                    bb.badSys = aa.sys;
                }
                else
                {
                    currentClient = aa.cli;
                    currentContact = aa.contact;
                    bb.goodSys = aa.sys;
                    bb.cli = aa.cli;
                    bb.isBad = false;
                    currentGoodSys = aa.sys;
                    bb.badSys = 0;
                }


                adrList.Add(bb);
            }

            adrList = adrList.OrderBy(c => c.cli).ThenBy(c => c.badSys).ToList();

            //remove the dupes and repoint billcopy
            foreach (billAdr ba in adrList)
            {
                if (ba.isBad)
                {
                    //update billcopy
                    string sss = "update billcopy set BilCpyBilAdr = " + ba.goodSys + " from billcopy " +
                    " inner join billto on billtosysnbr = BilCpyBillTo " +
                    " inner join matter on billtosysnbr = matbillto " +
                    " inner join client on clisysnbr = matclinbr " +
                    " where BilCpyBilAdr = " + ba.badSys;
                    _jurisUtility.ExecuteNonQuery(0, sss);

                    //remove from billingaddress
                    sss = "delete from billingaddress where BilAdrSysNbr = " + ba.badSys;
                     _jurisUtility.ExecuteNonQuery(0, sss);
                }

            }
            



            string SQL = "update sysparam set spnbrvalue = (select max(biladrsysnbr) from billingaddress) where spname = 'LastSysNbrBillAddress'";
            _jurisUtility.ExecuteNonQuery(0, SQL);

            MessageBox.Show("Done");
        }

        private void processDefaultAddys()
        {
            string sysparam = "  select SpNbrValue from sysparam where SpName = 'LastSysNbrBillAddress'";
            DataSet dds = _jurisUtility.RecordsetFromSQL(sysparam);
            int LastAddyID = 5;
            string cell = "";
            if (dds != null && dds.Tables.Count > 0)
            {
                cell = dds.Tables[0].Rows[0].ToString();

                foreach (DataRow dr in dds.Tables[0].Rows)
                {
                    cell = dr[0].ToString();
                }

            }
            LastAddyID = Convert.ToInt32(cell) + 1;
            //clients
            dds.Clear();
            sysparam = "SELECT clisysnbr, clicode FROM Client where clisysnbr not in (select biladrclinbr from billingaddress)";
            dds = _jurisUtility.RecordsetFromSQL(sysparam);
            if (dds != null && dds.Tables.Count > 0)
            {
                foreach (DataRow dr in dds.Tables[0].Rows)
                {
                    LastAddyID = LastAddyID + 1;
                    string sql = "Insert into BillingAddress(BilAdrSysNbr, BilAdrCliNbr, BilAdrUsageFlg, BilAdrNickName, BilAdrPhone, BilAdrFax, BilAdrContact, BilAdrName, BilAdrAddress, BilAdrCity, BilAdrState, BilAdrZip, BilAdrCountry, BilAdrType, BilAdrEmail) " +
                            " values (" + LastAddyID + ", " + dr[0].ToString() + ", " +
                            " 'C', 'Default - ' + '" + dr[1].ToString() + "', '', "
                            + "  '', 'FIRM', " +
                            " 'FIRM', " +
                            "'', "
                            + " '', '', '','', 0, '')";

                    _jurisUtility.ExecuteNonQuery(0, sql);

                }
                string SQL = "update sysparam set spnbrvalue = (select max(biladrsysnbr) from billingaddress) where spname = 'LastSysNbrBillAddress'";
                _jurisUtility.ExecuteNonQuery(0, SQL);

            }






          //  sql = "Insert into BillCopy(BilCpyBillTo,BilCpyBilAdr,BilCpyComment,BilCpyNbrOfCopies,BilCpyPrintFormat,BilCpyEmailFormat,BilCpyExportFormat,BilCpyARFormat) "
//+ " values (  (select matbillto from matter inner join client on matclinbr = clisysnbr where matcode = right('000000000000' + '" + clientItems[1].Trim() + "', 12) and dbo.jfn_formatClientCode(clicode) = right('000000000000' + '" + clientItems[0] + "', " + clientLength + ")),"
//+ "  (select max(BilAdrSysNbr) from billingaddress) ,'',1,1,0,0,0 )";


           // _jurisUtility.ExecuteNonQuery(0, sql);



        }

        private void processDefaultAddysMat()
        {
            string sysparam = "select matsysnbr, matbillto, matclinbr from matter " +
                            " inner join  BillTo on BillTo.BillToSysNbr = Matter.MatBillTo " +
                            " left outer JOIN BillCopy ON BillTo.BillToSysNbr = BillCopy.BilCpyBillTo " +
                            " left outer JOIN BillingAddress ON BillingAddress.BilAdrSysNbr = BillCopy.BilCpyBilAdr " +
                            " where BilAdrSysNbr is null";


            DataSet dds = _jurisUtility.RecordsetFromSQL(sysparam);
            if (dds != null && dds.Tables.Count > 0)
            {
                foreach (DataRow dr in dds.Tables[0].Rows)
                {
                    string sql = "Insert into BillCopy(BilCpyBillTo, BilCpyBilAdr, BilCpyComment, BilCpyNbrOfCopies, BilCpyPrintFormat, BilCpyEmailFormat, BilCpyExportFormat, BilCpyARFormat) "
                                   + " values ( " + dr[1].ToString() + ","
                                       + "  (select top 1 BilAdrSysNbr from billingaddress where biladrclinbr = " + dr[2].ToString() + ") ,'',1,1,0,0,0 )";

                    _jurisUtility.ExecuteNonQuery(0, sql);

                }


            }

        }


        private void button3_Click(object sender, EventArgs e)
        {
            string sql = "insert into CliOrigAtty ([COrigCli],[COrigAtty],[COrigPcnt]) select clisysnbr, 1, 100.00 from client";

            _jurisUtility.ExecuteNonQuery(0, sql);

            sql = "  insert into [MatOrigAtty] ([MOrigMat],[MOrigAtty],[MOrigPcnt]) select matsysnbr, 1, 100.00 from matter";

            _jurisUtility.ExecuteNonQuery(0, sql);
            MessageBox.Show("The process is complete", "Confirmation", MessageBoxButtons.OK, MessageBoxIcon.None);
        }

        private void button4_Click(object sender, EventArgs e)
        {
            processDefaultAddys();
            processDefaultAddysMat();
                MessageBox.Show("The process is complete", "Confirmation", MessageBoxButtons.OK, MessageBoxIcon.None);
            
        }

        private void button5_Click(object sender, EventArgs e)
        {

            string line;
            List<string> clientItems = new List<string>();
            System.IO.StreamReader file = new System.IO.StreamReader(@"c:\intel\attys.txt");
            while ((line = file.ReadLine()) != null)
            {
                        clientItems = line.Split('\t').ToList();
                foreach (string item in clientItems)
                    File.AppendAllText(@"c:\intel\wtf.txt", "sum(case when empinitials = '" + item + "' then FSPBilHrsEntered else 0 end) as " + item + "," + Environment.NewLine);
            }

            file.Close();
            MessageBox.Show("The process is complete", "Confirmation", MessageBoxButtons.OK, MessageBoxIcon.None);
        }
    }
}

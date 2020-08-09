using System;
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
                        clientItems = line.Split('\t').ToList();
                        string sql = "Insert into Client(CliSysNbr,CliCode,CliNickName,CliReportingName,CliSourceOfBusiness,CliPhoneNbr,CliFaxNbr,CliContactName,CliDateOpened,CliOfficeCode,CliBillingAtty,CliPracticeClass, "
                           + " CliFeeSch,CliTaskCodeXref,CliExpSch,CliExpCodeXref,CliBillFormat,CliBillAgreeCode,CliFlatFeeIncExp,CliRetainerType,CliExpFreqCode,CliFeeFreqCode,CliBillMonth,CliBillCycle, "
                           + " CliExpThreshold,CliFeeThreshold,CliInterestPcnt,CliInterestDays,CliDiscountOption,CliDiscountPcnt,CliSurchargeOption,CliSurchargePcnt,CliTax1Exempt,CliTax2Exempt,CliTax3Exempt,CliBudgetOption,CliReqPhaseOnTrans, "
                           + " CliReqTaskCdOnTime,CliReqActyCdOnTime,CliReqTaskCdOnExp,CliPrimaryAddr,CliType,CliEditFormat,CliThresholdOption,CliRespAtty,CliBillingField01,CliBillingField02,CliBillingField03,CliBillingField04,CliBillingField05, "
                           + " CliBillingField06,CliBillingField07,CliBillingField08,CliBillingField09,CliBillingField10,CliBillingField11,CliBillingField12,CliBillingField13,CliBillingField14,CliBillingField15,CliBillingField16,CliBillingField17,CliBillingField18,CliBillingField19, "
                           + " CliBillingField20,CliCTerms,CliCStatus,CliCStatus2)  "
                           + " Select (select max(clisysnbr) from client) + 1, right('000000000000' + '" + clientItems[0] + "', 12), left('" + clientItems[1].Replace("\"", "").Replace("'", "''") + "', 30), left('" + clientItems[1].Replace("\"", "").Replace("'", "''") + "', 30), "
                           + " CliSourceOfBusiness, left('" + clientItems[9] + "', 20), left('" + clientItems[10] + "', 20), left('" + clientItems[11].Replace("\"", "").Replace("'", "''") + "', 30), '" + clientItems[6] + "', '" + clientItems[2] + "', "
                         + " (select empsysnbr from employee where empinitials = '" + clientItems[4] + "') as CliBillingAtty, "
                         + "'" + clientItems[3] + "', 'STDR',null,'STDR',null, "
                           + " 'BF01','" + clientItems[5] + "',CliFlatFeeIncExp,CliRetainerType, '" + clientItems[7] + "', '" + clientItems[8] + "' ,CliBillMonth,CliBillCycle,CliExpThreshold,CliFeeThreshold,CliInterestPcnt,CliInterestDays,CliDiscountOption,CliDiscountPcnt, "
                           + " CliSurchargeOption,CliSurchargePcnt,CliTax1Exempt,CliTax2Exempt,CliTax3Exempt,CliBudgetOption,CliReqPhaseOnTrans,CliReqTaskCdOnTime,CliReqActyCdOnTime,CliReqTaskCdOnExp,null,CliType,'P100', "
                           + " CliThresholdOption,null,'','','','','','','','','','','','', "
                           + " '','','','','','','','',CliCTerms,CliCStatus,CliCStatus2 "
                           + " from Client where clisysnbr = (select min(clisysnbr) from client)";

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
                            " values (((select max(BilAdrSysNbr) from billingaddress) + 1), (select clisysnbr from client where dbo.jfn_FormatClientCode(clicode) = '" + clientItems[0] + "'), " +
                            " 'M', left('" + clientItems[2].Replace("\"", "").Replace("'", "''") + "', 25) + '" + rr.Next(10000,99999).ToString() + "', left('" + clientItems[3].Replace("\"", "").Replace("'", "''") + "', 20), "
                             + "         left('" + clientItems[4].Replace("\"", "").Replace("'", "''") + "', 20), left('" + clientItems[5].Replace("\"", "").Replace("'", "''") + "', 30), " +
                            " left('" + clientItems[2].Replace("\"", "").Replace("'", "''") + "', 50), left(left('" + clientItems[6].Replace("\"", "").Replace("'", "''") + "', 250) + char(10) + char(13) + left('" + clientItems[7].Replace("\"", "").Replace("'", "''") + "', 250), 250), " 
                            + " left('" + clientItems[8] + "', 20), left('" + clientItems[9] + "', 2), left('" + clientItems[10] + "', 9),left('" + clientItems[11] + "', 20), 0, left('" + clientItems[12] + "', 255) )";

                        errorSQL = sql;


                        _jurisUtility.ExecuteNonQuery(0, sql);

                        sql = "Insert into BillCopy(BilCpyBillTo,BilCpyBilAdr,BilCpyComment,BilCpyNbrOfCopies,BilCpyPrintFormat,BilCpyEmailFormat,BilCpyExportFormat,BilCpyARFormat) "
                                + " values (  (select matbillto from matter inner join client on matclinbr = clisysnbr where dbo.jfn_FormatMatterCode(matcode) = '" + clientItems[1] + "' and dbo.jfn_formatClientCode(clicode) = '" + clientItems[0] + "')"
                                + " , (select max(BilAdrSysNbr) from billingaddress) ,'',1,1,0,0,0 )";


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
            string SQL = "update sysparam set spnbrvalue = (select max(biladrsysnbr) from billingaddress) where spname = 'LastSysNbrBillAddress'";
            _jurisUtility.ExecuteNonQuery(0, SQL);

          

        }

        private void processMatters()
        {
            try
            {
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
                               + "       Select ((select max(MatSysNbr) from matter) + 1), case when (select clisysnbr from client where dbo.jfn_FormatClientCode(clicode) = '" + clientItems[0] + "') is null then 2 else (select clisysnbr from client where dbo.jfn_FormatClientCode(clicode) = '" + clientItems[0] + "') end, 1,  "
                               + "       right('0000000000' + '" + clientItems[1] + "', 12), left('" + clientItems[2].Replace("\"", "").Replace("'", "''") + "', 30), left('" + clientItems[2].Replace("\"", "").Replace("'", "''") + "', 30), '" + clientItems[3].Replace("\"", "").Replace("'", "''") + "', MatRemarks, left('" + clientItems[11].Replace("\"", "").Replace("'", "''") + "', 20), "
                             + "         left('" + clientItems[12].Replace("\"", "").Replace("'", "''") + "', 20), left('" + clientItems[13].Replace("\"", "").Replace("'", "''") + "', 30), '" + clientItems[8] + "', 'O','0', '01/01/1900','" + clientItems[4] + "','" + clientItems[5] + "','STDR',null,'STDR',null,MatQuickAction,'" + clientItems[7] + "', "
                              + "        MatFlatFeeIncExp,MatRetainerType,MatFltFeeOrRetainer,'" + clientItems[10] + "','" + clientItems[9] + "',MatBillMonth,MatBillCycle,MatExpThreshold,MatFeeThreshold,MatInterestPcnt,MatInterestDays,MatDiscountOption,MatDiscountPcnt, "
                               + "       MatSurchargeOption,MatSurchargePcnt,MatSplitMethod,MatSplitThreshold,MatSplitPriorAmtBld,MatBudgetOption,MatBudgetPhase,MatReqPhaseOnTrans,MatReqTaskCdOnTime,MatReqActyCdOnTime,MatReqTaskCdOnExp,MatTax1Exempt, "
                            + "          MatTax2Exempt,MatTax3Exempt,'01/01/1900','01/01/1900','01/01/1900','01/01/1900','01/01/1900',0,0,0,0,0,MatVisionAddr, "
                             + "         MatThresholdOption,MatType,'','','','','','','','','','','','', "
                            + "          '','','','','','','','',MatCTerms,MatCStatus,MatCStatus2 "
                              + "        from matter where matsysnbr = (select min(matsysnbr) from matter where matstatusflag = 'O')";



                            _jurisUtility.ExecuteNonQuery(0, sql);

                            Random rr = new Random();
                            sql = "Insert into BillTo (BillToSysNbr,BillToCliNbr,BillToUsageFlg,BillToNickName,BillToBillingAtty,BillToBillFormat,BillToEditFormat,BillToRespAtty) " +
                                "values (((select max(BillToSysNbr) from billto) + 1), " +
                                "case when (select clisysnbr from client where dbo.jfn_FormatClientCode(clicode) = '" + clientItems[0] + "') is null then 2 else (select clisysnbr from client where dbo.jfn_FormatClientCode(clicode) = '" + clientItems[0] + "') end, 'M', left('" + clientItems[2].Replace("\"", "").Replace("'", "''") + "', 26) + '" + rr.Next(1000, 9999).ToString() + "', (select empsysnbr from employee where empinitials = '" + clientItems[6] + "'), 'CR40', 'P100', null)";
                            errorSQL = sql;

                            _jurisUtility.ExecuteNonQuery(0, sql);


                            sql = "update matter set matbillto = (select max(billtosysnbr) from billto) where matsysnbr = (select max(matsysnbr) from  matter)";
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

                string SQL = "delete from billto where billtosysnbr = 1";
                _jurisUtility.ExecuteNonQuery(0, SQL);

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

            string line;
            int counter = 0;
            List<string> clientItems = new List<string>();
            System.IO.StreamReader file = new System.IO.StreamReader(origFile);
            string errorSQL = "";
            List<string> AllLines = new List<string>();
            while ((line = file.ReadLine()) != null)
            {
                AllLines.Add(line);
            }
            file.Close();

            var distinctList = AllLines.Distinct().ToList();
            foreach (string oneLine in distinctList)
            {
                try
                {
                    if (counter > 0)
                    {
                        clientItems = oneLine.Split('\t').ToList();
                        string sql = "";
                        if (flag == 1)
                        {
                            sql = " 	Insert into CliOrigAtty(COrigCli,COrigAtty,COrigPcnt) " +
                                " values ((select clisysnbr from client where dbo.jfn_FormatClientCode(clicode) = '" + clientItems[1] + "'), " +
                                " (select empsysnbr from employee where empinitials = '" + clientItems[0] + "'), cast('" + clientItems[2] + "' as decimal(5,2)))";
                        }
                        else
                        {
                            sql = "Insert into MatOrigAtty(MOrigMat, MOrigAtty, MOrigPcnt) " +
                                " values ((select matsysnbr from matter inner join client on matclinbr = clisysnbr where dbo.jfn_FormatMatterCode(matcode) = '" + clientItems[2] + "' and dbo.jfn_formatClientCode(clicode) = '" + clientItems[1] + "'), " +
                                " (select empsysnbr from employee where empinitials = '" + clientItems[0] + "'), cast('" + clientItems[3] + "' as decimal(5,2)))";

                        }
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
            //add any leftovers that have no originator
            string SQL = "  insert into MatOrigAtty (MOrigMat,MOrigAtty,MOrigPcnt) " +
                        " select a.matsysnbr, (select billtobillingatty from billto inner join matter on matbillto = billtosysnbr where matsysnbr = a.matsysnbr), 100.00 " +
                        " from matter a where a.matsysnbr in ((select matsysnbr from matter where matsysnbr not in (SELECT distinct MOrigMat FROM MatOrigAtty)))";
            _jurisUtility.ExecuteNonQuery(0, SQL);

            SQL = "  insert into CliOrigAtty (cOrigCli,cOrigAtty,COrigPcnt) " +
                    " select c.clisysnbr, clibillingatty, 100.00 " +
                    " from client c where c.clisysnbr not in (select distinct COrigCli from CliOrigAtty)";
            _jurisUtility.ExecuteNonQuery(0, SQL);

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
    }
}

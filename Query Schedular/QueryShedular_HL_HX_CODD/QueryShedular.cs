/*'Copyright notice: ã 2004 by Bird Information Systems Pvt. Ltd. All rights reserved.
'********************************************************************************************
' This file contains trade secrets of Bird Information Systems. No part
' may be reproduced or transmitted in any form by any means or for any purpose
' without the express written permission of Bird Information Systems.
'********************************************************************************************
'$Author: Neeraj $Logfile: /AAMS/Queryschedular/QueryShedular.cs $
'$Workfile: QueryShedular.cs $
'$Revision: 1 $
'$Archive: /AAMS/Queryschedular/QueryShedular.cs $
'$Modtime: 6/15/10 2:57p $
*/

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Configuration;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using QueryShedular_HL_HX_CODD;
using System.Data.SqlClient;
using System.Diagnostics;
using System.IO;
using System.Windows;
using System.Data.OleDb;



namespace QueryShedular_HL_HX_CODD
{     
    public partial class frmQueryShedular : Form
    {
        #region Variable Declaration
        private String strFilePath = System.Configuration.ConfigurationSettings.AppSettings["OutputFolder"];
        private bool isExecutingCODD;
        private bool isExecutingHL;
        private bool isExecutingHX;
        private String strMonthyearCODD = String.Empty;
        private String strMonthyearHX = String.Empty;
        private String strMonthyearHL = String.Empty;

        private Stopwatch myCallBackWatchCODD = new Stopwatch();
        private Stopwatch myCallBackWatchHX = new Stopwatch();
        private Stopwatch myCallBackWatchHL = new Stopwatch();
        private Stopwatch myCallBackWatchGrp = new Stopwatch();


        SqlConnection objConCODD;
        SqlConnection objConHL;
        SqlConnection objConHX;
        SqlConnection objConLivedatabase;

        String strDisplay = String.Empty;
        String strStoredProcName = String.Empty;

        public DataTable myTable1;
        public DataTable myTable2;
        public DataTable myTable3;
        public DataTable myTable4;
        public DataTable myTable4_1;

        public DataSet gblGrpds;        

        public Boolean boolExport;
        public int intYear;
        public int intMonth;
        public String StrCountry;
        public String StrCountryCode;
        public Boolean blndogrpAdjustment;

        public SqlTransaction objSqlTransaction;
        public SqlBulkCopy objSqlbulkCopy;

        public SqlCommand cmd;
        public SqlCommand cmd1;
        public SqlCommand cmd2;

        public Boolean boolBulkInsert;

        // For connection string Password Decription 
        public IPOSS.bizBarcode.bzBarcode objBarcode = new IPOSS.bizBarcode.bzBarcode();
        public String strKey = "1793";
        public String strIpaddress;
        public String strHostname;
        public List<string> list = new List<string>();
        public String printString;        
        public String strGrpdata_Process_Status = "";
        public String strGrpdata_Found_Status = "";
        public string[] strgrpColumnarray = new string[7];
        public String strMissingOfficeid;
      

        #endregion

        #region Declaration of delegate for Timeinfo and DataResult
        private delegate void displayDataCODD(DataTable exportDTCODD);
        private delegate void displayTimeInfoDelegateCODD(String Text);
        //private delegate void ExportExcel(DataTable objDT , String strMonth  , String strYear, String QueryType);

        public delegate void displayDataHX(DataTable exportDTHX);
        private delegate void displayTimeInfoDelegateHX(String Text);

        private delegate void displayDataHL(DataTable exportDTHL);
        private delegate void displayTimeInfoDelegateHL(String Text);

        private delegate void displayDataGrp(DataTable exportDTGrp);
        private delegate void displayTimeInfoDelegateGrp(String Text);


        #endregion

        #region frmQueryShedular
        public frmQueryShedular()
        {
            InitializeComponent();
        }
        #endregion

        #region btnExit_Click
        private void btnExit_Click(object sender, EventArgs e)
        {
            this.Close();
            //Application.Exit();
        }
        #endregion

        #region Form Load frmQueryShedular_Load
        private void frmQueryShedular_Load(object sender, EventArgs e)
        {
            try
            {
                ConfigureOutPutFolder();
                ConnectToDB();
                disableEnablepanel();
                EnableDisableMenuItems();
                fillControls();
                clearControls();
                grpQuery.BackgroundImage = Myresource.graybackground;
                grpButton.BackgroundImage = Myresource.graybackground;
                chkCODD.CheckedChanged += new EventHandler(chkCODD_CheckedChanged);
                chkHX.CheckedChanged += new EventHandler(chkHX_CheckedChanged);
                chkHL.CheckedChanged += new EventHandler(chkHL_CheckedChanged);
                chkGrpData.CheckedChanged += new EventHandler(chkGrpData_CheckedChanged);

                lblCbar.Visible = false;
                lblHLbar.Visible = false;
                lblHXbar.Visible = false;

                pbarCODD.Visible = false;
                pbarHL.Visible = false;
                pbarHX.Visible = false;
                grpStatusbar.Visible = false;

                IPNetworking objIp = new IPNetworking();
                objIp.GetIP4Address(out strIpaddress, out strHostname);
                this.Text = "QUERY SHEDULAR NIDT  SYSTEM IP/HOST [" + strIpaddress + "]" + "[" + strHostname + "]";


                foreach (ToolStripMenuItem item in contextMenuStrip1.Items)
                {
                    if (item.HasDropDownItems)
                    {
                        DoSubItems(item);
                    }
                }
                this.ContextMenuStrip = this.contextMenuStrip1;


                foreach (Control ctrl in this.Controls)
                {
                    ctrl.MouseDown += new MouseEventHandler(ctrl_MouseDown);

                    if (ctrl.GetType() == typeof(GroupBox))
                    {
                        foreach (Control ctrl1 in ((GroupBox)ctrl).Controls)
                        {
                            ctrl1.MouseDown += new MouseEventHandler(ctrl_MouseDown);
                        }
                        //((System.Windows.Forms.Button)ctrl).Enabled = false;                    
                    }
                }
            }
            catch (Exception exep)
            {
                MessageBox.Show(exep.Message);
            }
        }


        #endregion

        #region ctrl_MouseDown
        void ctrl_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                //contextMenuStrip1.Enabled=false;
            }
        }
        #endregion

        #region ContextMenu Click Code
        private void DoSubItems(ToolStripMenuItem item)
        {
            foreach (ToolStripMenuItem subitem in item.DropDownItems)
            {
                subitem.Click += new EventHandler(item_Click);
            }
        }

        void item_Click(object sender, EventArgs e)
        {
            String strFormat;
            try
            {
                ToolStripMenuItem clickedMenu = sender as ToolStripMenuItem;
                if (clickedMenu.OwnerItem.Text.ToUpper().Trim() == "CODD")
                {
                    strFormat = clickedMenu.Text;
                    //MessageBox.Show("CODD");
                    ExportToExcel_CSV(myTable1, "CODD", strFormat);
                }
                if (clickedMenu.OwnerItem.Text.ToUpper().Trim() == "HX")
                {
                    //MessageBox.Show("HX");
                    strFormat = clickedMenu.Text;
                    //MessageBox.Show("CODD");
                    ExportToExcel_CSV(myTable2, "HX", strFormat);
                }
                if (clickedMenu.OwnerItem.Text.ToUpper().Trim() == "HL")
                {
                    //MessageBox.Show("HL");
                    strFormat = clickedMenu.Text;
                    //MessageBox.Show("CODD");
                    ExportToExcel_CSV(myTable3, "HL", strFormat);
                }
            }
            catch (Exception exep)
            {
                MessageBox.Show(exep.Message);
            }
        }
        #endregion

        #region ExportToEXCEL_CSV
        private void ExportToExcel_CSV(DataTable ObjDT, String strType, String strFormat)
        {
            ExportToExcel_WIN_App objExport = new ExportToExcel_WIN_App();


            if (strType == "CODD")
            {
                //this.dataGridView1.DataSource = ObjDT;                
                try
                {
                    if (strFormat.ToUpper().Trim() == "XLS")
                    {
                        this.toolStripProgressBar1.Minimum = 1;
                        this.toolStripProgressBar1.Maximum = 5;
                        this.toolStripProgressBar1.Value = 1;
                        this.toolStripProgressBar1.Value = this.toolStripProgressBar1.Value + 1;
                        DisplayStatus("Exporting CODD in xls..");

                        boolExport = objExport.ExportToExcel(ObjDT, strMonthyearCODD, "CODD", drpCcountry.SelectedValue.ToString(), strFilePath, toolStripStatusLabel1);
                        this.toolStripProgressBar1.Value = 5;
                        if (boolExport == true)
                        {
                            DisplayStatus("Successfully exported CODD.." + strFilePath);
                        }
                        else
                        {
                            toolStripProgressBar1.Value = 1;
                            pbarCODD.Value = 1;
                        }
                    }
                    else if (strFormat.ToUpper().Trim() == "CSV")
                    {
                        this.toolStripProgressBar1.Minimum = 1;
                        this.toolStripProgressBar1.Maximum = 5;
                        this.toolStripProgressBar1.Value = 1;
                        this.toolStripProgressBar1.Value = this.toolStripProgressBar1.Value + 1;
                        DisplayStatus("Exporting CODD in csv..");

                        boolExport = objExport.Exportcsv(ObjDT, "CODD", drpCcountry.SelectedValue.ToString(), strMonthyearCODD, strFilePath, toolStripStatusLabel1);
                        this.toolStripProgressBar1.Value = 5;
                        if (boolExport == true)
                        {
                            DisplayStatus("Successfully exported CODD.." + strFilePath);
                        }
                        else
                        {
                            toolStripProgressBar1.Value = 1;
                            pbarCODD.Value = 1;
                        }
                    }
                }
                catch (Exception exe)
                {
                    DisplayStatus("Error while export CODD.." + exe.Message);
                }
            }

            if (strType == "HX")
            {
                //Boolean boolExport;
                //this.dataGridView2.DataSource = ObjDT;                
                try
                {
                    if (strFormat.ToUpper().Trim() == "XLS")
                    {
                        this.toolStripProgressBar1.Minimum = 1;
                        this.toolStripProgressBar1.Maximum = 5;
                        this.toolStripProgressBar1.Value = 1;
                        this.toolStripProgressBar1.Value = this.toolStripProgressBar1.Value + 1;
                        DisplayStatus("Exporting HX in xls..");

                        //ExportToExcel_WIN_App objExport = new ExportToExcel_WIN_App();
                        boolExport = objExport.ExportToExcel(ObjDT, strMonthyearHX, "HX", drpHXcountry.SelectedValue.ToString(), strFilePath, toolStripStatusLabel1);
                        this.toolStripProgressBar1.Value = 5;
                        if (boolExport == true)
                        {
                            DisplayStatus("Successfully exported HX in xls format.." + strFilePath);
                        }
                        else
                        {
                            toolStripProgressBar1.Value = 1;
                            pbarHX.Value = 1;
                        }
                    }
                    else if (strFormat.ToUpper().Trim() == "CSV")
                    {
                        this.toolStripProgressBar1.Minimum = 1;
                        this.toolStripProgressBar1.Maximum = 5;
                        this.toolStripProgressBar1.Value = 1;
                        this.toolStripProgressBar1.Value = this.toolStripProgressBar1.Value + 1;
                        DisplayStatus("Exporting HX in csv..");

                        objExport.Exportcsv(ObjDT, "HX", drpHXcountry.SelectedValue.ToString(), strMonthyearHX, strFilePath, toolStripStatusLabel1);
                        this.toolStripProgressBar1.Value = 5;
                        if (boolExport == true)
                        {
                            DisplayStatus("Successfully exported HX in csv foramt.." + strFilePath);
                        }
                        else
                        {
                            toolStripProgressBar1.Value = 1;
                            pbarCODD.Value = 1;
                        }
                    }
                }
                catch (Exception exe)
                {
                    DisplayStatus("Error while export HX.." + exe.Message);
                }
            }

            if (strType == "HL")
            {
                //Boolean boolExport;                
                //this.dataGridView3.DataSource = ObjDT;                
                try
                {
                    if (strFormat.ToUpper().Trim() == "XLS")
                    {
                        this.toolStripProgressBar1.Minimum = 1;
                        this.toolStripProgressBar1.Maximum = 5;
                        this.toolStripProgressBar1.Value = 1;
                        this.toolStripProgressBar1.Value = this.toolStripProgressBar1.Value + 1;
                        DisplayStatus("Exporting HL in XLS..");

                        //ExportToExcel_WIN_App objExport = new ExportToExcel_WIN_App();
                        boolExport = objExport.ExportToExcel(ObjDT, strMonthyearHL, "HL", drpHLcountry.SelectedValue.ToString(), strFilePath, toolStripStatusLabel1);
                        this.toolStripProgressBar1.Value = 5;
                        if (boolExport == true)
                        {
                            DisplayStatus("Successfully exported HL in xls format.." + strFilePath);
                        }
                        else
                        {
                            toolStripProgressBar1.Value = 1;
                            pbarHL.Value = 1;
                        }
                    }
                    else if (strFormat.ToUpper().Trim() == "CSV")
                    {
                        this.toolStripProgressBar1.Minimum = 1;
                        this.toolStripProgressBar1.Maximum = 5;
                        this.toolStripProgressBar1.Value = 1;
                        this.toolStripProgressBar1.Value = this.toolStripProgressBar1.Value + 1;
                        DisplayStatus("Exporting HL in csv..");

                        boolExport = objExport.Exportcsv(ObjDT, "HL", drpHLcountry.SelectedValue.ToString(), strMonthyearHL, strFilePath, toolStripStatusLabel1);
                        this.toolStripProgressBar1.Value = 5;
                        if (boolExport == true)
                        {
                            DisplayStatus("Successfully exported HL in csv foramt.." + strFilePath);
                        }
                        else
                        {
                            toolStripProgressBar1.Value = 1;
                            pbarCODD.Value = 1;
                        }
                    }
                }
                catch (Exception exe)
                {
                    DisplayStatus("Error while export HL.." + exe.Message);
                }
            }
        }
        #endregion

        #region  Upload NIDT to AAMS
        private void UploadNIDTtoAAMS()
        {
            SqlCommand SqlCommandHL = null;

            try
            {
                lblHLbar.Visible = true;
                pbarHL.Visible = true;
                lblHL.Visible = false;

                lblTimeSHL.Visible = true;
                hlstime.Visible = true;
                lblTimeEHL.Visible = true;
                hletime.Visible = true;

                hlstime.Text = String.Format("{0:T}", DateTime.Now);

                pbarHL.Maximum = 100;
                pbarHL.Minimum = 1;
                pbarHL.Value = 1;
                pbarHL.Value = pbarHL.Value + 10;
                DisplayStatus("Extracting NIDT data...");

                isExecutingHL = true;
                btnExport.Enabled = true;

                //SqlCommandHL = new SqlCommand("select top 1000  location_code,name,address from location_master", objConHL);
                //call to make stored Proc name

                strStoredProcName = GetProcdureName(drpHLMonth.SelectedItem.ToString());

                //strStoredProcName = "UP_NIDT_PRODUCTIVITY_FEB_NEW";

                int param_month = System.Convert.ToInt16(drpHLMonth.SelectedIndex);
                int param_year = System.Convert.ToInt16(drpHLYear.SelectedItem);
                String param_country = System.Convert.ToString(drpHLcountry.Text);
                intYear = param_year;
                intMonth = param_month;
                StrCountry = param_country;
                StrCountryCode = System.Convert.ToString(drpHLcountry.SelectedValue);

                //String param_country = System.Convert.ToString(drpHLcountry.SelectedValue);
                strMonthyearHL = drpHLMonth.SelectedItem.ToString() + "" + param_year.ToString();

                SqlCommandHL = new SqlCommand();
                SqlCommandHL.CommandType = CommandType.StoredProcedure;
                SqlCommandHL.CommandText = strStoredProcName;
                SqlCommandHL.Connection = objConHL;
                SqlCommandHL.Parameters.Add(new SqlParameter("@MONTH", SqlDbType.Int));
                SqlCommandHL.Parameters["@MONTH"].Value = param_month;
                SqlCommandHL.Parameters.Add(new SqlParameter("@YEAR", SqlDbType.Int));
                SqlCommandHL.Parameters["@YEAR"].Value = param_year;

                SqlCommandHL.Connection.Open();

                //start clock for HL 
                myCallBackWatchHL.Start();
                AsyncCallback myCallBackHL = new AsyncCallback(HandleCallbackHL);
                SqlCommandHL.BeginExecuteReader(myCallBackHL, SqlCommandHL);
            }
            catch (Exception exe)
            {                
                DisplayStatus("Error while export HL.." + exe.Message);
                if (((System.Data.SqlClient.SqlException)(exe)).Number == 53)
                {
                    MessageBox.Show("NIDT server line may be down.please contact to Admin!", "AAMS Admin", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                }
                else
                {
                    MessageBox.Show(exe.Message, "AAMS Admin", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                }
                isExecutingHL = false;
                pbarHL.Maximum = 100;
                pbarHL.Minimum = 1;
                pbarHL.Value = 1;
                pbarHL.Visible = false;
            }
            finally
            {
                
            }
        }
        #endregion

        #region  Return Missing Offidceid from NIDT server for th ecurrent Upload
        private Boolean retMissingOfficeid()
        {
            DisplayStatus("Calculating missing officeid's...");

            SqlConnection sqlcon = new SqlConnection();            
            sqlcon = objConHL;
            if (sqlcon.State == ConnectionState.Open)
            {
                sqlcon.Close();
            }
            sqlcon.Open();

            strMissingOfficeid = string.Empty;
            try
            {
                //CASE STATUS 1 , FILE WAS LOADED FINALLY AND USE WANTS TO RELOAD AGAIN
                using (SqlCommand objcmd = new SqlCommand("UP_GET_MISSING_OFFICEID", sqlcon))
                {
                    int param_month = System.Convert.ToInt16(drpHLMonth.SelectedIndex);
                    int param_year = System.Convert.ToInt16(drpHLYear.SelectedItem);

                    objcmd.CommandType = CommandType.StoredProcedure;
                    objcmd.CommandTimeout = 60;
                    objcmd.Parameters.AddWithValue("@MONTH", param_month);
                    objcmd.Parameters.AddWithValue("@YEAR", param_year);

                    SqlParameter OprmRESULT = new SqlParameter();
                    OprmRESULT.ParameterName = "@RESULT";
                    OprmRESULT.SqlDbType = SqlDbType.VarChar;
                    OprmRESULT.Size = 5000;
                    OprmRESULT.Direction = ParameterDirection.Output;
                    objcmd.Parameters.Add(OprmRESULT);
                    objcmd.ExecuteNonQuery();
                    strMissingOfficeid = (string)OprmRESULT.Value.ToString();
                    if (strMissingOfficeid.Replace(",", "") != "")
                    {
                        //MessageBox.Show("Please note the following officeids are missing\n" + strMissingOfficeid + "\n\ndo you want to exit or continue transfer HL data into Live server?", "AAMS Admin", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                        DisplayStatus("Missing Officeid's found...");
                        return true;
                    }
                    else
                    {
                        DisplayStatus("Missing Officeid's not found...");    
                        return false;
                    }
                }
            }
            catch (Exception exe)
            {
                this.toolStripStatusLabel1.Text = exe.Message;
            }
            finally
            { 
              if (sqlcon != null)
              {
                  sqlcon.Close();
              }
            }
        return false;
        }
        #endregion


        #region  Configure Excel Export Folder
        private void ConfigureOutPutFolder()
        {
            try
            {
                // Code to Configure Folder in Local System
                DirectoryInfo objDir;
                //objDir = new DirectoryInfo(@"C:\ExportQuery");
                objDir = new DirectoryInfo(strFilePath);
                if (objDir.Exists == false)
                {
                    //MessageBox.Show(strFilePath);
                    objDir.Create();
                }
            }
            catch (Exception exe)
            {
                this.toolStripStatusLabel1.Text = exe.Message;
            }
        }
        #endregion

        #region GetProcdureName
        private String GetProcdureName(String strMonth)
        {
            String StrShortMonth = String.Empty;
            switch (strMonth)
            {
                case "January":
                    StrShortMonth = "Jan";
                    break;
                case "February":
                    StrShortMonth = "Feb";
                    break;
                case "March":
                    StrShortMonth = "Mar";
                    break;
                case "April":
                    StrShortMonth = "Apr";
                    break;
                case "May":
                    StrShortMonth = "May";
                    break;
                case "June":
                    StrShortMonth = "Jun";
                    break;
                case "July":
                    StrShortMonth = "Jul";
                    break;
                case "August":
                    StrShortMonth = "Aug";
                    break;
                case "September":
                    StrShortMonth = "Sep";
                    break;
                case "October":
                    StrShortMonth = "Oct";
                    break;
                case "November":
                    StrShortMonth = "Nov";
                    break;
                case "December":
                    StrShortMonth = "Dec";
                    break;
            }
            return "UP_NIDT_PRODUCTIVITY_" + StrShortMonth;
            //return "UP_NIDT_PRODUCTIVITY_MAR_test";
        }
        #endregion

        #region ValidateControls
        private Boolean ValidateControls()
        {
            Boolean objCntl;
            objCntl = false;
            if (chkCODD.Checked == true || chkHL.Checked == true || chkHX.Checked == true || chkGrpData.Checked == true)
            {
                if (chkCODD.Checked == true)
                {
                    if (drpCMonth.SelectedIndex == 0)
                    {
                        lblC.Visible = true;
                        objCntl = true;
                    }
                    else
                    {
                        if (drpCyear.SelectedIndex == 0)
                        {
                            lblC.Visible = true;
                        }
                        else
                        {
                            lblC.Visible = false;
                        }

                    }
                    if (drpCyear.SelectedIndex == 0)
                    {
                        lblC.Visible = true;
                        objCntl = true;
                    }
                    else
                    {
                        if (drpCMonth.SelectedIndex == 0)
                        {
                            lblC.Visible = true;
                        }
                        else
                        {
                            lblC.Visible = false;
                        }
                    }
                }
                if (chkHL.Checked == true)
                {
                    if (drpHLMonth.SelectedIndex == 0)
                    {
                        lblHL.Visible = true;
                        objCntl = true;
                    }
                    else
                    {
                        if (drpHLMonth.SelectedIndex == 0)
                        {
                            lblHL.Visible = true;
                        }
                        else
                        {
                            lblHL.Visible = false;
                        }
                    }

                    if (drpHLYear.SelectedIndex == 0)
                    {
                        lblHL.Visible = true;
                        objCntl = true;
                    }
                    else
                    {
                        if (drpHLMonth.SelectedIndex == 0)
                        {
                            lblHL.Visible = true;
                        }
                        else
                        {
                            lblHL.Visible = false;
                        }
                    }
                }
                if (chkHX.Checked == true)
                {
                    if (drpHXMonth.SelectedIndex == 0)
                    {
                        lblHX.Visible = true;
                        objCntl = true;
                    }
                    else
                    {
                        if (drpHXYear.SelectedIndex == 0)
                        {
                            lblHX.Visible = true;
                        }
                        else
                        {
                            lblHX.Visible = false;
                        }
                    }
                    if (drpHXYear.SelectedIndex == 0)
                    {
                        lblHX.Visible = true;
                        objCntl = true;
                    }
                    else
                    {
                        if (drpHXMonth.SelectedIndex == 0)
                        {
                            lblHX.Visible = true;
                        }
                        else
                        {
                            lblHX.Visible = false;
                        }
                    }
                }
                if (chkGrpData.Checked == true)
                {
                    if (drpGrpMonth.SelectedIndex == 0)
                    {
                        lblGrp.Visible = true;
                        objCntl = true;
                    }
                    else
                    {
                        if (drpGrpYear.SelectedIndex == 0)
                        {
                            lblGrp.Visible = true;
                        }
                        else
                        {
                            lblGrp.Visible = false;
                        }
                        if (txtFilename.Text.Trim()=="")
                        {
                            lblGrp.Visible = true;
                            objCntl = true;
                        }
                    }
                    if (drpGrpYear.SelectedIndex == 0)
                    {
                        lblGrp.Visible = true;
                        objCntl = true;
                    }
                    else
                    {
                        if (drpGrpMonth.SelectedIndex == 0)
                        {
                            lblGrp.Visible = true;
                        }
                        else
                        {
                            lblGrp.Visible = false;
                        }
                        if (txtFilename.Text.Trim() == "")
                        {
                            lblGrp.Visible = true;
                            objCntl = true;
                        }
                    }
                }
            }
            return objCntl;
        }
        #endregion

        #region Execute Button Click
        void btnExecute_Click(object sender, EventArgs e)
        {
            timer1.Enabled = true;
            Boolean objvalidate;
            EnableDisableMenuItems();
            objvalidate = ValidateControls();

            if (objvalidate == false)
            {
                if (isExecutingCODD || isExecutingHL || isExecutingHX)
                {
                    MessageBox.Show(this, "Already executing. Please wait until " +
                    "the current query has completed.");
                }
                else
                {
                    SqlCommand SqlCommandCODD = null;
                    SqlCommand SqlCommandHX = null;                   
                    SqlCommand SqlCommandGrp = null;

                    dataGridView1.DataSource = null;
                    dataGridView2.DataSource = null;
                    dataGridView3.DataSource = null;
                    this.toolStripStatusLabel1.Text = "";
                    grpStatusbar.Visible = true;
                    clearControls();
                    try
                    {
                        if (chkCODD.Checked == true && chkHL.Checked == true && chkHX.Checked == true)
                        {
                            DisplayStatus("Connecting...");
                            DisplayStatus("Executing...");
                        }

                        if (chkCODD.Checked == false && chkHL.Checked == false && chkHX.Checked == false)
                        {
                            isExecutingCODD = false;
                        }

                        if (chkCODD.Checked == true)
                        {
                            lblCbar.Visible = true;
                            pbarCODD.Visible = true;
                            lblC.Visible = false;

                            lblTimeSCODD.Visible = true;
                            coddstime.Visible = true;
                            lblTimeECODD.Visible = true;
                            coddetime.Visible = true;

                            coddstime.Text = String.Format("{0:T}", DateTime.Now);

                            pbarCODD.Maximum = 100;
                            pbarCODD.Minimum = 1;
                            pbarCODD.Value = 1;
                            pbarCODD.Value = pbarCODD.Value + 10;

                            //objConCODD.Open();

                            //start the watch for 
                            myCallBackWatchCODD.Start();

                            isExecutingCODD = true;
                            btnExport.Enabled = true;

                            int param_month = System.Convert.ToInt16(drpCMonth.SelectedIndex);
                            int param_year = System.Convert.ToInt16(drpCyear.SelectedItem);
                            String param_country = System.Convert.ToString(drpCcountry.Text);
                            intYear = param_year;
                            intMonth = param_month;
                            StrCountry = param_country;
                            StrCountryCode = System.Convert.ToString(drpCcountry.SelectedValue);

                            //SqlCommandCODD = new SqlCommand("select top 100 * from location_master", objConCODD);

                            SqlCommandCODD = new SqlCommand();
                            SqlCommandCODD.CommandType = CommandType.StoredProcedure;
                            SqlCommandCODD.CommandText = "UP_NIDT_PRODUCTIVITY_CODD";
                            SqlCommandCODD.Connection = objConCODD;

                            SqlCommandCODD.Parameters.Add(new SqlParameter("@MONTH", SqlDbType.Int));
                            SqlCommandCODD.Parameters["@MONTH"].Value = param_month;

                            SqlCommandCODD.Parameters.Add(new SqlParameter("@YEAR", SqlDbType.Int));
                            SqlCommandCODD.Parameters["@YEAR"].Value = param_year;
                            strMonthyearCODD = drpCMonth.SelectedItem.ToString() + "" + param_year.ToString();

                            SqlCommandCODD.Parameters.Add(new SqlParameter("@COUNTRY", SqlDbType.VarChar, 3));
                            if (StrCountryCode == "0")
                                SqlCommandCODD.Parameters["@COUNTRY"].Value = DBNull.Value;
                            else
                                SqlCommandCODD.Parameters["@COUNTRY"].Value = StrCountryCode;


                            SqlCommandCODD.Connection.Open();

                            AsyncCallback myCallBackCODD = new AsyncCallback(HandleCallbackCODD);
                            SqlCommandCODD.BeginExecuteReader(myCallBackCODD, SqlCommandCODD);

                        }
                        else
                        {
                            lblCbar.Visible = false;
                            pbarCODD.Visible = false;

                            lblTimeSCODD.Visible = false;
                            coddstime.Visible = false;
                            lblTimeECODD.Visible = false;
                            coddetime.Visible = false;

                        }
                        if (chkHL.Checked == true)
                        {
                            /*New code implemented for the Group data date 30/04/2015
                             * 1. check group data availability in system
                             * 2. Alert to User i.e data found/not found
                             * 3. keep found data in global datatable
                             * 4. Alert Payments exists for the selected Month and Year
                             */
                            strGrpdata_Process_Status = "";

                            strStoredProcName = "[UP_GET_INC_AIRLINE_GROUP_DATA]";
                            int param_monthgrp = System.Convert.ToInt16(drpHLMonth.SelectedIndex);
                            int param_yeargrp = System.Convert.ToInt16(drpHLYear.SelectedItem);

                            SqlCommandGrp = new SqlCommand();
                            SqlCommandGrp.CommandType = CommandType.StoredProcedure;
                            SqlCommandGrp.CommandText = strStoredProcName;
                            SqlCommandGrp.Connection = objConLivedatabase;
                            SqlCommandGrp.Parameters.Add(new SqlParameter("@MONTH", SqlDbType.Int));
                            SqlCommandGrp.Parameters["@MONTH"].Value = param_monthgrp;
                            SqlCommandGrp.Parameters.Add(new SqlParameter("@YEAR", SqlDbType.Int));
                            SqlCommandGrp.Parameters["@YEAR"].Value = param_yeargrp;

                            if (SqlCommandGrp.Connection.State == ConnectionState.Open)
                            {
                                SqlCommandGrp.Connection.Close();
                            }

                            SqlCommandGrp.Connection.Open();

                            AsyncCallback myCallBackGrp = new AsyncCallback(HandleCallbackGrp);
                            SqlCommandGrp.BeginExecuteReader(myCallBackGrp, SqlCommandGrp);
                            /*end New code implemented for the Group data date 30/04/2015*/
                        }
                        else
                        {

                            lblHLbar.Visible = false;
                            pbarHL.Visible = false;

                            lblTimeSHL.Visible = false;
                            hlstime.Visible = false;
                            lblTimeEHL.Visible = false;
                            hletime.Visible = false;

                        }
                        if (chkHX.Checked == true)
                        {
                            lblHXbar.Visible = true;
                            pbarHX.Visible = true;
                            lblHX.Visible = false;

                            lblTimeSHX.Visible = true;
                            hxstime.Visible = true;
                            lblTimeEHX.Visible = true;
                            hxetime.Visible = true;

                            hxstime.Text = String.Format("{0:T}", DateTime.Now);

                            pbarHX.Maximum = 100;
                            pbarHX.Minimum = 1;
                            pbarHX.Value = 1;
                            pbarHX.Value = pbarHX.Value + 10;

                            //objConHX.Open();
                            //start the watch for 
                            myCallBackWatchHX.Start();
                            isExecutingHX = true;
                            btnExport.Enabled = true;
                            //SqlCommandHX = new SqlCommand("select top 1000  location_code,name,address from location_master", objConHX);

                            int param_month = System.Convert.ToInt16(drpHXMonth.SelectedIndex);
                            int param_year = System.Convert.ToInt16(drpHXYear.SelectedItem);
                            String param_country = System.Convert.ToString(drpHXcountry.Text);
                            intYear = param_year;
                            intMonth = param_month;
                            StrCountry = param_country;
                            StrCountryCode = System.Convert.ToString(drpHXcountry.SelectedValue);
                            //SqlCommandCODD = new SqlCommand("select top 100 * from location_master", objConCODD);

                            SqlCommandHX = new SqlCommand();
                            SqlCommandHX.CommandType = CommandType.StoredProcedure;
                            SqlCommandHX.CommandText = "UP_NIDT_PRODUCTIVITY_HX";
                            SqlCommandHX.Connection = objConHX;

                            SqlCommandHX.Parameters.Add(new SqlParameter("@MONTH", SqlDbType.Int));
                            SqlCommandHX.Parameters["@MONTH"].Value = param_month;

                            SqlCommandHX.Parameters.Add(new SqlParameter("@YEAR", SqlDbType.Int));
                            SqlCommandHX.Parameters["@YEAR"].Value = param_year;

                            SqlCommandHX.Parameters.Add(new SqlParameter("@COUNTRY", SqlDbType.VarChar, 3));
                            if (StrCountryCode == "0")
                                SqlCommandHX.Parameters["@COUNTRY"].Value = DBNull.Value;
                            else
                                SqlCommandHX.Parameters["@COUNTRY"].Value = StrCountryCode;

                            strMonthyearHX = drpHXMonth.SelectedItem.ToString() + "" + param_year.ToString();

                            SqlCommandHX.Connection.Open();

                            AsyncCallback myCallBackHX = new AsyncCallback(HandleCallbackHX);
                            SqlCommandHX.BeginExecuteReader(myCallBackHX, SqlCommandHX);
                        }
                        else
                        {
                            lblHXbar.Visible = false;
                            pbarHX.Visible = false;

                            lblTimeSHX.Visible = false;
                            hxstime.Visible = false;
                            lblTimeEHX.Visible = false;
                            hxetime.Visible = false;
                        }

                    }

                    catch (Exception ex)
                    {
                        isExecutingCODD = false;

                        DisplayStatus(string.Format("Ready (last error: {0})", ex.Message));
                        if (objConCODD != null)
                        {
                            objConCODD.Close();
                        }
                        if (objConHL != null)
                        {
                            objConHL.Close();
                        }
                        if (objConHX != null)
                        {
                            objConHX.Close();
                        }
                        //MessageBox.Show(ex.Message);
                    }
                }
            }
        }
        #endregion

        #region DisplayStatus
        private void DisplayStatus(string Text)
        {
            //strRepalced = Text;
            strDisplay = Text;

            //strDisplay= strRepalced.Replace("Connecting...", " ");
            //strDisplay = strDisplay + strRepalced.Replace("Executing...", " ");
            this.toolStripStatusLabel1.Text = strDisplay;
        }
        #endregion

        #region Delegate Haldlers for CODD , HX , HL
        private void HandleCallbackCODD(IAsyncResult myResult)
        {
            try
            {
                SqlCommand myCmd1 = (SqlCommand)myResult.AsyncState;
                SqlDataReader myReader1 = myCmd1.EndExecuteReader(myResult);

                //myTable1.Clear();
                myTable1 = new DataTable();
                myTable1.Load(myReader1);

                // Stop the watch so we can see how long it took to process
                myCallBackWatchCODD.Stop();
                String myCallBackTime = myCallBackWatchCODD.ElapsedMilliseconds.ToString();
                displayTimeInfoDelegateHX myWatchdisplayCODD = new displayTimeInfoDelegateHX(displayCODDTime);
                this.Invoke(myWatchdisplayCODD, myCallBackTime);

                displayDataCODD myDataDelegate = new displayDataCODD(DisplayDataResultsCODD);
                this.Invoke(myDataDelegate, myTable1);

            }
            catch (Exception exep)
            {
                this.Invoke(new displayDataCODD(DisplayDataResultsCODD), String.Format("Ready(last error: {0}", exep.Message));
                //MessageBox.Show(exep.Message);
            }
            finally
            {
                isExecutingCODD = false;
                if (objConCODD != null)
                {
                    objConCODD.Close();
                }
                //myTable1.Clear();
                //myTable1=null;
            }
        }

        private void HandleCallbackHX(IAsyncResult myResult)
        {
            try
            {
                SqlCommand myCmd2 = (SqlCommand)myResult.AsyncState;
                SqlDataReader myReader2 = myCmd2.EndExecuteReader(myResult);

                myTable2 = new DataTable();
                myTable2.Load(myReader2);

                // Stop the watch so we can see how long it took to process
                myCallBackWatchHX.Stop();

                String myCallBackTime = myCallBackWatchHX.ElapsedMilliseconds.ToString();
                displayTimeInfoDelegateHX myWatchdisplayHX = new displayTimeInfoDelegateHX(displayHXTime);
                this.Invoke(myWatchdisplayHX, myCallBackTime);

                displayDataHX myDataDelegate = new displayDataHX(DisplayDataResultsHX);
                this.Invoke(myDataDelegate, myTable2);
            }
            catch (Exception exep)
            {
                this.Invoke(new displayDataHX(DisplayDataResultsHX), String.Format("Ready(last error: {0}", exep.Message));
                //MessageBox.Show(exep.Message);
            }
            finally
            {
                isExecutingHX = false;
                if (objConHX != null)
                {
                    objConHX.Close();
                }
                //myTable2.Clear();
                //myTable2=null;
            }
        }

        private void HandleCallbackHL(IAsyncResult myResult)
        {
            try
            {
                SqlCommand myCmd3 = (SqlCommand)myResult.AsyncState;
                SqlDataReader myReader3 = myCmd3.EndExecuteReader(myResult);

                myTable3 = new DataTable();
                myTable3.Load(myReader3);
                DisplayStatus("NIDT data Extracted...");
                // Stop the watch so we can see how long it took to process
                myCallBackWatchHL.Stop();

                String myCallBackTime = myCallBackWatchHL.ElapsedMilliseconds.ToString();
                displayTimeInfoDelegateHL myWatchdisplayHL = new displayTimeInfoDelegateHL(displayHLTime);
                this.Invoke(myWatchdisplayHL, myCallBackTime);

                displayDataHL myDataDelegate = new displayDataHL(DisplayDataResultsHL);
                this.Invoke(myDataDelegate, myTable3);
            }
            catch (Exception exep)
            {
                this.Invoke(new displayDataHL(DisplayDataResultsHL), String.Format("Ready(last error: {0}", exep.Message));
                //MessageBox.Show(exep.Message);
            }
            finally
            {
                isExecutingHL = false;
                if (objConHL != null)
                {
                    objConHL.Close();
                }
                //myTable3.Clear();
                //myTable3=null;
            }
        }

        private void HandleCallbackGrp(IAsyncResult myResult)
        {
            try
            {
                SqlCommand myCmd4 = (SqlCommand)myResult.AsyncState;
                SqlDataReader myReader4 = myCmd4.EndExecuteReader(myResult);


                //myReader4.Read();
                myTable4 = new DataTable();
                myTable4.Load(myReader4); // Note this will automatically call the command NextResult() on the reader

                // Result set 2
                myTable4_1 = new DataTable();
                myTable4_1.Load(myReader4); // 


                // Stop the watch so we can see how long it took to process
                myCallBackWatchGrp.Stop();
                String myCallBackTime = myCallBackWatchGrp.ElapsedMilliseconds.ToString();

                displayTimeInfoDelegateGrp myWatchdisplayGrp = new displayTimeInfoDelegateGrp(displayGrpTime);
                this.Invoke(myWatchdisplayGrp, myCallBackTime);

                displayDataGrp myDataDelegate = new displayDataGrp(DisplayDataResultsGrp);
                this.Invoke(myDataDelegate, myTable4_1);
            }
            catch (Exception exep)
            {
                MessageBox.Show(exep.Message,"AAMS Admin", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                //this.Invoke(new displayDataGrp(DisplayDataResultsGrp),myTable4_1,  exep.Message);                                
            }
            finally
            {
                //isExecutingHL = false;
                if (objConLivedatabase != null)
                {
                    objConLivedatabase.Close();
                }                
            }
        }

        #endregion

        #region private void displayTime
        private void displayHLTime(String Text)
        {
            try
            {
                hletime.Text = Math.Round(TimeSpan.FromMilliseconds(double.Parse(Text)).TotalMinutes, 5).ToString() + " Minutes";
            }
            catch (Exception exe)
            {
                this.toolStripStatusLabel1.Text = "Error occured while display time : " + exe.Message;
            }
        }
        private void displayGrpTime(String Text)
        {
            try
            {
                hletime.Text = Math.Round(TimeSpan.FromMilliseconds(double.Parse(Text)).TotalMinutes, 5).ToString() + " Minutes";
            }
            catch (Exception exe)
            {
                this.toolStripStatusLabel1.Text = "Error occured while display time : " + exe.Message;
            }
        }
        private void displayHXTime(String Text)
        {
            try
            {
                hxetime.Text = Math.Round(TimeSpan.FromMilliseconds(double.Parse(Text)).TotalMinutes, 5).ToString() + " Minutes";
            }
            catch (Exception exe)
            {
                this.toolStripStatusLabel1.Text = "Error occured while display time : " + exe.Message;
            }
        }
        private void displayCODDTime(String Text)
        {
            try
            {
                coddetime.Text = Math.Round(TimeSpan.FromMilliseconds(double.Parse(Text)).TotalMinutes, 5).ToString() + " Minutes";
            }
            catch (Exception exe)
            {
                this.toolStripStatusLabel1.Text = "Error occured while display time : " + exe.Message;
            }
        }
        #endregion

        //private void ExportToExcel(DataTable objDT , String strMonth  , String strYear, String QueryType)
        //{
        //    try
        //    {
        //        this.toolStripProgressBar1.Minimum=1;
        //        this.toolStripProgressBar1.Maximum=5;
        //        this.toolStripProgressBar1.Value=1;
        //        this.toolStripProgressBar1.Value=this.toolStripProgressBar1.Value+1;
        //        this.toolStripStatusLabel1.Text = "Exporting CODD..";
        //        ExportToExcel_WIN_App objExport = new ExportToExcel_WIN_App();
        //        this.toolStripProgressBar1.Value=5;
        //        objExport.ExportToExcel(objDT,String.Empty,String.Empty,"CODD");
        //    }
        //    catch (Exception exe)
        //    {
        //        this.toolStripStatusLabel1.Text=exe.Message;
        //    }
        //}

        #region Display Result in grid & Export date to Excel
        private void DisplayDataResultsCODD(DataTable ObjDT)
        {
            foreach (ToolStripMenuItem item in contextMenuStrip1.Items)
            {
                if (item.HasDropDownItems)
                {
                    if (item.Text.ToUpper().Trim() == "CODD")
                    {
                        item.Enabled = true;
                        DoDisableSubItems(item, true);
                    }
                    else
                    {
                        if (item.Enabled == false)
                        {
                            item.Enabled = false;
                            DoDisableSubItems(item, false);
                        }
                    }
                }
            }
            pbarCODD.Value = pbarCODD.Value + 5;
            pbarCODD.Value = 100;
            lblRowcountCODD.Text = "Row Count : " + ObjDT.Rows.Count;
            DisplayStatus("CODD Ready...");
            pbarCODD.Cursor = Cursors.Default;

            /*Export data into database server Newly code Implemented as on date 17-12-2010 [Neeraj Goswami]*/
            DialogResult dlgResult = MessageBox.Show("Do you want to continue to Transfer CODD data into Live server database?", "AAMS Admin", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (dlgResult == DialogResult.Yes)
            {
                boolBulkInsert = ExportBulkInsert(ObjDT, "CODD","1");
                if (boolBulkInsert == true)
                {
                    MessageBox.Show("CODD data for the period of " + intMonth.ToString() + "/" + intYear.ToString() + "  successfully transfered to Live Server database ", "AAMS Admin", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }

        }

        private void DisplayDataResultsHX(DataTable ObjDT)
        {
            foreach (ToolStripMenuItem item in contextMenuStrip1.Items)
            {
                if (item.HasDropDownItems)
                {
                    if (item.Text.ToUpper().Trim() == "HX")
                    {
                        item.Enabled = true;
                        DoDisableSubItems(item, true);
                    }
                    else
                    {
                        if (item.Enabled == false)
                        {
                            item.Enabled = false;
                            DoDisableSubItems(item, false);
                        }
                    }
                }
            }
            pbarHX.Value = pbarHX.Value + 5;
            pbarHX.Value = 100;
            lblRowcountHX.Text = "Row Count : " + ObjDT.Rows.Count;
            DisplayStatus("HX Ready...");
            pbarHX.Cursor = Cursors.Default;

            /*Export data into database server Newly code Implemented as on date 17-12-2010 [Neeraj Goswami]*/
            DialogResult dlgResult = MessageBox.Show("Do you want to continue Transfer HX data into Live server database?", "AAMS Admin", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (dlgResult == DialogResult.Yes)
            {
                boolBulkInsert = ExportBulkInsert(ObjDT, "HX","1");
                if (boolBulkInsert == true)
                {
                    MessageBox.Show("HX data for the period of " + intMonth.ToString() + "/" + intYear.ToString() + "  successfully transfered to Live Server database ", "AAMS Admin", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
        }

        private void DisplayDataResultsHL(DataTable ObjDT)
        {
            try
            {
                foreach (ToolStripMenuItem item in contextMenuStrip1.Items)
                {
                    if (item.HasDropDownItems)
                    {
                        if (item.Text.ToUpper().Trim() == "HL")
                        {
                            item.Enabled = true;
                            DoDisableSubItems(item, true);
                        }
                        else
                        {
                            if (item.Enabled == false)
                            {
                                item.Enabled = false;
                                DoDisableSubItems(item, false);
                            }
                        }
                    }
                }
                pbarHL.Value = pbarHL.Value + 5;
                pbarHL.Value = 100;
                DisplayStatus("HL Ready...");
                lblRowcounthl.Text = "Row Count : " + ObjDT.Rows.Count;
                pbarHL.Cursor = Cursors.Default;

                

                /*Export data into database server Newly code Implemented as on date 17-12-2010 [Neeraj Goswami]*/
                DialogResult dlgResult = MessageBox.Show("do you want to continue transfer HL data into live server database?", "AAMS Admin", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                
                if (dlgResult == DialogResult.Yes)
                {
                    DisplayStatus("Adjusting group data into productivity ...");
                    /*do Group data productivity adjustment*/
                    if (blndogrpAdjustment)
                    {
                        //MessageBox.Show(myTable3.Rows + Environment.NewLine , "AAMS Admin", MessageBoxButtons.OK, MessageBoxIcon.Question);
                        //MessageBox.Show(myTable4.Rows + Environment.NewLine , "AAMS Admin", MessageBoxButtons.OK, MessageBoxIcon.Question);                        

                        DataRow[] myTable4row;
                        String strAirlineCode = "";
                        int prod;
                        int dom;
                        int intl;

                        prod = 0;
                        dom = 0;
                        intl = 0;

                        foreach (DataColumn col in myTable3.Columns)
                            col.ReadOnly = false;

                        foreach (DataRow r in myTable3.Rows)
                        {
                            //for 9W 
                            myTable4row = myTable4.Select("AirLineCode ='9W' and Officeid ='" + r[4].ToString() + "'"); //AirLineCode ='9W' and Officeid='AMDUG3219'

                            foreach (DataRow temp in myTable4row)
                            {
                                strAirlineCode = Convert.ToString(temp["AirLineCode"]);
                                if (strAirlineCode == "9W")
                                {
                                    prod = Convert.ToInt32(r["PRODUCTIVITY_CODD_PK_HX"]);
                                    dom = Convert.ToInt32(r["9W_DOM_CODD_PK_HX"]);
                                    intl = Convert.ToInt32(r["INTL"]);

                                    prod = prod - Convert.ToInt32(temp["Productivity"]);
                                    dom = dom - Convert.ToInt32(temp["Dom"]);
                                    intl = intl - Convert.ToInt32(temp["Intl"]);
                                    
                                    r["PRODUCTIVITY_CODD_PK_HX"] = prod;
                                    r["9W_DOM_CODD_PK_HX"] = dom;
                                    r["INTL"] = intl;
                                }
                            }
                            myTable3.AcceptChanges();

                            //for AI 
                            myTable4row = myTable4.Select("AirLineCode ='AI' and Officeid ='" + r[4].ToString() + "'"); //AirLineCode ='9W' and Officeid='AMDUG3219'
                            foreach (DataRow temp in myTable4row)
                            {
                                prod = 0;
                                dom = 0;
                                intl = 0;
                                strAirlineCode = Convert.ToString(temp["AirLineCode"]);

                                prod = Convert.ToInt32(r["PRODUCTIVITY_CODD_PK_HX"]);
                                dom = Convert.ToInt32(r["AI_DOM_CODD_PK_HX"]);
                                intl = Convert.ToInt32(r["INTL"]);

                                prod = prod - Convert.ToInt32(temp["Productivity"]);
                                dom = dom - Convert.ToInt32(temp["Dom"]);
                                intl = intl - Convert.ToInt32(temp["Intl"]);

                                r["PRODUCTIVITY_CODD_PK_HX"] = prod;
                                r["AI_DOM_CODD_PK_HX"] = dom;
                                r["INTL"] = intl;
                            }
                            myTable3.AcceptChanges();
                        }
                        ObjDT = myTable3;
                    }
                    /*end do Group data productivity adjustment*/


                    /*Code Missing Officeid's newly implemented as on dated 10/03/2015 [Neeraj Goswami]
                    * One more change is required in Scheduler that when scheduler upload the data to the server and at that time 
                    * if any officeid is mismatched between AAMS data and NIDT data then system should show the list of missing id list 
                    * and if we want to upload the data without adding that id in AAMS then it should be loaded successfully. 
                    * Right now it rejects whole data to upload if any id not exist.*/

                    DisplayStatus("checking missing officeid's...");

                    DataView dv = new DataView(ObjDT);
                    dv.RowFilter = "[LCODE] is null";

                    string[] TobeDistinct = { "OFFICEID" };
                    string strMissingOffIds = "";
                    DataTable dtDistinct = GetDistinctRecords(dv, TobeDistinct);

                    if (dtDistinct != null)
                    {
                        if (dtDistinct.Rows.Count > 0)
                        {
                            int i = 0;
                            foreach (DataRow r in dtDistinct.Rows)
                            {
                                strMissingOffIds += r[0].ToString() + " , ";
                                i++;
                            }
                            DialogResult dlgMissingofficeid = MessageBox.Show("Missing Officeid's exists , do you want to by pass these officeId's and continue transfer HL data into live server ?" + Environment.NewLine + Environment.NewLine + strMissingOffIds, "AAMS Admin", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                            if (dlgMissingofficeid == DialogResult.Yes)
                            {
                                DataTable dtTarget = new DataTable();
                                //try
                                //{    
                                dtTarget = ObjDT.Clone();
                                DataRow[] rowsToCopy;
                                rowsToCopy = ObjDT.Select("[LCODE] is not null");
                                foreach (DataRow temp in rowsToCopy)
                                {
                                    dtTarget.ImportRow(temp);
                                }

                                ObjDT = dtTarget;
                                dtTarget = null;
                                DisplayStatus("Adding final NIDT data to AAMS server...");
                                boolBulkInsert = ExportBulkInsert(ObjDT, "HL", "1");

                                if (boolBulkInsert == true)
                                {
                                    if (blndogrpAdjustment)
                                    {
                                        if (objConLivedatabase.State == ConnectionState.Open)
                                        {
                                            objConLivedatabase.Close();
                                        }
                                        objConLivedatabase.Open();

                                        cmd2 = new SqlCommand("UPDATE T_INC_NIDT_PRODUCTIVITY_GROUPDATA_MAIN SET STATUS = 1 WHERE MONTH = " + intMonth + " AND YEAR = " + intYear, objConLivedatabase);
                                        // 2. Call Execute query 
                                        int intRowaffected = cmd2.ExecuteNonQuery();
                                    }
                                    MessageBox.Show("HL data for the period of " + intMonth.ToString() + "/" + intYear.ToString() + "  successfully transfered to Live Server database ", "AAMS Admin", MessageBoxButtons.OK, MessageBoxIcon.Information);
                                }
                            }
                        }
                        else
                        {
                            DisplayStatus("Adding final NIDT data to AAMS server...");
                            boolBulkInsert = ExportBulkInsert(ObjDT, "HL", "1");
                            if (boolBulkInsert == true)
                            {
                                if (blndogrpAdjustment)
                                {
                                    if (objConLivedatabase.State == ConnectionState.Open)
                                    {
                                        objConLivedatabase.Close();
                                    }
                                    objConLivedatabase.Open();

                                    cmd2 = new SqlCommand("UPDATE T_INC_NIDT_PRODUCTIVITY_GROUPDATA_MAIN SET STATUS = 1 WHERE MONTH = " + intMonth + " AND YEAR = " + intYear, objConLivedatabase);
                                    // 2. Call Execute query 
                                    int intRowaffected = cmd2.ExecuteNonQuery();
                                }
                                MessageBox.Show("HL data for the period of " + intMonth.ToString() + "/" + intYear.ToString() + "  successfully transfered to Live Server database ", "AAMS Admin", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            }
                            else if (boolBulkInsert == false)
                            {
                                DisplayStatus("Operation canceled for uploading nidt to aams server...");
                            }
                        }
                    }
                }
                else if (dlgResult == DialogResult.No)
                {
                    DisplayStatus("Adjusting group data into productivity ...");
                    /*do Group data productivity adjustment*/
                    if (blndogrpAdjustment)
                    {
                        //MessageBox.Show(myTable3.Rows + Environment.NewLine , "AAMS Admin", MessageBoxButtons.OK, MessageBoxIcon.Question);
                        //MessageBox.Show(myTable4.Rows + Environment.NewLine , "AAMS Admin", MessageBoxButtons.OK, MessageBoxIcon.Question);                        

                        DataRow[] myTable4row;
                        String strAirlineCode = "";
                        int prod;
                        int dom;
                        int intl;

                        prod = 0;
                        dom = 0;
                        intl = 0;

                        foreach (DataColumn col in myTable3.Columns)
                            col.ReadOnly = false;

                        foreach (DataRow r in myTable3.Rows)
                        {
                            //for 9W 
                            myTable4row = myTable4.Select("AirLineCode ='9W' and Officeid ='" + r[4].ToString() + "'"); //AirLineCode ='9W' and Officeid='AMDUG3219'
                            foreach (DataRow temp in myTable4row)
                            {
                                strAirlineCode = Convert.ToString(temp["AirLineCode"]);
                                if (strAirlineCode == "9W")
                                {
                                    prod = Convert.ToInt32(r["PRODUCTIVITY_CODD_PK_HX"]);
                                    dom = Convert.ToInt32(r["9W_DOM_CODD_PK_HX"]);
                                    intl = Convert.ToInt32(r["INTL"]);

                                    prod = prod - Convert.ToInt32(temp["Productivity"]);
                                    dom = dom - Convert.ToInt32(temp["Dom"]);
                                    intl = intl - Convert.ToInt32(temp["Intl"]);


                                    r["PRODUCTIVITY_CODD_PK_HX"] = prod;
                                    r["9W_DOM_CODD_PK_HX"] = dom;
                                    r["INTL"] = intl;
                                }
                            }
                            myTable3.AcceptChanges();

                            //for AI 
                            myTable4row = myTable4.Select("AirLineCode ='AI' and Officeid ='" + r[4].ToString() + "'"); //AirLineCode ='9W' and Officeid='AMDUG3219'
                            foreach (DataRow temp in myTable4row)
                            {
                                prod = 0;
                                dom = 0;
                                intl = 0;
                                strAirlineCode = Convert.ToString(temp["AirLineCode"]);

                                prod = Convert.ToInt32(r["PRODUCTIVITY_CODD_PK_HX"]);
                                dom = Convert.ToInt32(r["AI_DOM_CODD_PK_HX"]);
                                intl = Convert.ToInt32(r["INTL"]);

                                prod = prod - Convert.ToInt32(temp["Productivity"]);
                                dom = dom - Convert.ToInt32(temp["Dom"]);
                                intl = intl - Convert.ToInt32(temp["Intl"]);

                                r["PRODUCTIVITY_CODD_PK_HX"] = prod;
                                r["AI_DOM_CODD_PK_HX"] = dom;
                                r["INTL"] = intl;
                            }
                            myTable3.AcceptChanges();
                        }
                        ObjDT = myTable3;
                    }

                   /*Code Missing Officeid's newly implemented as on dated 10/03/2015 [Neeraj Goswami]
                   * One more change is required in Scheduler that when scheduler upload the data to the server and at that time 
                   * if any officeid is mismatched between AAMS data and NIDT data then system should show the list of missing id list 
                   * and if we want to upload the data without adding that id in AAMS then it should be loaded successfully. 
                   * Right now it rejects whole data to upload if any id not exist.*/

                    DisplayStatus("checking missing officeid's...");

                    DataView dv = new DataView(ObjDT);
                    dv.RowFilter = "[LCODE] is null";

                    string[] TobeDistinct = { "OFFICEID" };
                    string strMissingOffIds = "";
                    DataTable dtDistinct = GetDistinctRecords(dv, TobeDistinct);

                    if (dtDistinct != null)
                    {
                        if (dtDistinct.Rows.Count > 0)
                        {
                            int i = 0;
                            foreach (DataRow r in dtDistinct.Rows)
                            {
                                strMissingOffIds += r[0].ToString() + " , ";
                                i++;
                            }
                            DialogResult dlgMissingofficeid = MessageBox.Show("Missing Officeid's exists , do you want to by pass these officeId's and continue transfer HL data into live server ?" + Environment.NewLine + Environment.NewLine + strMissingOffIds, "AAMS Admin", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                            if (dlgMissingofficeid == DialogResult.Yes)
                            {
                                DataTable dtTarget = new DataTable();
                                //try
                                //{    
                                dtTarget = ObjDT.Clone();
                                DataRow[] rowsToCopy;
                                rowsToCopy = ObjDT.Select("[LCODE] is not null");
                                foreach (DataRow temp in rowsToCopy)
                                {
                                    dtTarget.ImportRow(temp);
                                }

                                ObjDT = dtTarget;
                                dtTarget = null;
                                DisplayStatus("Adding final NIDT data to AAMS server...");
                                boolBulkInsert = ExportBulkInsert(ObjDT, "HL", "0");

                                if (boolBulkInsert == true)
                                {
                                    if (blndogrpAdjustment)
                                    {
                                        if (objConLivedatabase.State == ConnectionState.Open)
                                        {
                                            objConLivedatabase.Close();
                                        }
                                        objConLivedatabase.Open();

                                        cmd2 = new SqlCommand("UPDATE T_INC_NIDT_PRODUCTIVITY_GROUPDATA_MAIN SET STATUS = 1 WHERE MONTH = " + intMonth + " AND YEAR = " + intYear, objConLivedatabase);
                                        // 2. Call Execute query 
                                        int intRowaffected = cmd2.ExecuteNonQuery();
                                    }
                                    MessageBox.Show("HL data for the period of " + intMonth.ToString() + "/" + intYear.ToString() + "  successfully transfered to Live Server database ", "AAMS Admin", MessageBoxButtons.OK, MessageBoxIcon.Information);
                                }
                            }
                        }
                        else
                        {
                            DisplayStatus("Adding final NIDT data to AAMS server...");
                            boolBulkInsert = ExportBulkInsert(ObjDT, "HL", "0");
                            if (boolBulkInsert == true)
                            {
                                if (blndogrpAdjustment)
                                {
                                    if (objConLivedatabase.State == ConnectionState.Open)
                                    {
                                        objConLivedatabase.Close();
                                    }
                                    objConLivedatabase.Open();

                                    cmd2 = new SqlCommand("UPDATE T_INC_NIDT_PRODUCTIVITY_GROUPDATA_MAIN SET STATUS = 1 WHERE MONTH = " + intMonth + " AND YEAR = " + intYear, objConLivedatabase);
                                    // 2. Call Execute query 
                                    int intRowaffected = cmd2.ExecuteNonQuery();
                                }
                                MessageBox.Show("HL data for the period of " + intMonth.ToString() + "/" + intYear.ToString() + "  successfully transfered to Live Server database ", "AAMS Admin", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            }
                            else if (boolBulkInsert == false)
                            {
                                DisplayStatus("Operation canceled for uploading nidt to aams server...");
                            }
                        }
                    }   
                }
            }
            catch (Exception exep)
            {
                MessageBox.Show(exep.Message);
            }
            finally
            {
                //dtTarget = null;
                //isExecutingHL = false;
                //if (objConHL != null)
                //{
                //    objConHL.Close();
                //}
                ////myTable3.Clear();
                ////myTable3=null;
            }
        }
        

        private void DisplayDataResultsGrp(DataTable ObjDT)
        {
            try
            {
                foreach (ToolStripMenuItem item in contextMenuStrip1.Items)
                {
                    if (item.HasDropDownItems)
                    {
                        if (item.Text.ToUpper().Trim() == "HL")
                        {
                            item.Enabled = true;
                            DoDisableSubItems(item, true);
                        }
                        else
                        {
                            if (item.Enabled == false)
                            {
                                item.Enabled = false;
                                DoDisableSubItems(item, false);
                            }
                        }
                    }
                }
                pbarHL.Maximum = 100;
                pbarHL.Minimum = 1;
                pbarHL.Value = 1;

                pbarHL.Value = pbarHL.Value + 5;
                pbarHL.Value = 100;
                DisplayStatus("Group data Ready...");
                lblRowcounthl.Text = "Row Count : " + ObjDT.Rows.Count;
                pbarHL.Cursor = Cursors.Default;

                if (ObjDT.Rows.Count >= 1)
                {
                    strGrpdata_Found_Status = "TRUE";

                    foreach (DataRow r in ObjDT.Rows)
                    {
                        strGrpdata_Process_Status = r[2].ToString();
                    }

                    if (strGrpdata_Process_Status.ToUpper() == "FALSE")
                    {
                        strGrpdata_Process_Status = "Unprocessed";
                    }
                    else if (strGrpdata_Process_Status.ToUpper() == "TRUE")
                    {
                        strGrpdata_Process_Status = "Processed";
                    }

                    if (strGrpdata_Process_Status == "Unprocessed")
                    {
                        DisplayStatus("<<Available>>Airline Group data productivity...");
                        DialogResult dlgResult = MessageBox.Show("<<Available>>\nAirline Group data productivity is available without adjusted for the selected Month and Year in Live Server.\n\n\ndo you want to continue to process HL data?", "AAMS Admin", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                        if (dlgResult == DialogResult.Yes)
                        {                            
                            blndogrpAdjustment = true;

                            //Calculate Missing officeids
                            Boolean blnmissOFID =  retMissingOfficeid();
                            if (blnmissOFID == true)
                            {
                                DialogResult dlgmisofidResult = MessageBox.Show("Please note the following officeids are missing\n" + strMissingOfficeid + "\ncontinue transfer HL data into Live server\n\n\ndo you want to exit?", "AAMS Admin", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                                if (dlgmisofidResult == DialogResult.No)
                                {
                                    UploadNIDTtoAAMS();
                                }                                
                            }
                            else
                            {
                                UploadNIDTtoAAMS();
                            }                            
                        }
                    }
                    else if ((strGrpdata_Process_Status == "Processed"))
                    {
                        DisplayStatus("<<Adjusted>>Airline Group data productivity...");
                        DialogResult dlgResult = MessageBox.Show("<<Adjusted>>\nAirline Group productivity already adjusted in Live server.\n\n\ndo you want to continue to process HL data?", "AAMS Admin", MessageBoxButtons.YesNo, MessageBoxIcon.Stop);
                        if (dlgResult == DialogResult.Yes)
                        {
                            blndogrpAdjustment = true;
                            DisplayStatus("<<Adjusted>>Airline Group data productivity...");
                            blndogrpAdjustment = true;

                            //Calculate Missing officeids
                            Boolean blnmissOFID = retMissingOfficeid();
                            if (blnmissOFID == true)
                            {
                                DialogResult dlgmisofidResult = MessageBox.Show("Please note the following officeids are missing\n" + strMissingOfficeid + "\ncontinue transfer HL data into Live server\n\n\ndo you want to exit?", "AAMS Admin", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                                if (dlgmisofidResult == DialogResult.No)
                                {
                                    UploadNIDTtoAAMS();
                                }
                            }
                            else
                            {
                                UploadNIDTtoAAMS();
                            }                            
                        }
                    }
                }
                else if(ObjDT.Rows.Count == 0)
                {
                    strGrpdata_Process_Status = "Unavailable";
                    DisplayStatus("<<Not Available>>Airline Group data productivity...");

                    DialogResult dlgResult = MessageBox.Show("<<Not Available>>\nAirline Group data productivity is not available for the selected Month and Year in Live Server.\n\n\ndo you want to continue to process HL data?", "AAMS Admin", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                    if (dlgResult == DialogResult.Yes)
                    {
                        DisplayStatus("<<Not Available>>Airline Group data productivity...");
                        Boolean blnmissOFid = retMissingOfficeid();
                        if (blnmissOFid == true)
                        {
                            DialogResult dlgmisofidResult = MessageBox.Show("Please note the following officeids are missing\n" + strMissingOfficeid + "\ncontinue transfer HL data into Live server\n\n\ndo you want to exit?", "AAMS Admin", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                            if (dlgmisofidResult == DialogResult.No)
                            {
                                UploadNIDTtoAAMS();
                            }
                        }
                        else
                        {
                            UploadNIDTtoAAMS();
                        }
                    }
                    //else
                    //{
                    //    UploadNIDTtoAAMS();
                    //}
                }                
            }            
            catch (Exception exep)
            {
                MessageBox.Show(exep.Message);
            }
            finally
            {
                //pbarHL.Maximum = 100;
                //pbarHL.Minimum = 1;
                //pbarHL.Value = 1;
                ObjDT = null;                
            }
        }

        //Following function will return Distinct records for LCODE column.
        public static DataTable GetDistinctRecords(DataView dt, string[] Columns)
        {
            DataTable dtUniqRecords = new DataTable();
            dtUniqRecords = dt.ToTable(true, Columns);
            return dtUniqRecords;
        }
        #endregion

        #region Mapping Column before Bulk Insert
        // Mapping of each Column while Exporting to database server
        private Boolean ExportBulkInsert(DataTable objBulkDT, String strExportType,String optionaldoBulkTransfer)
        {
            int intCount;

            if (objBulkDT.Rows.Count == 0)
            {
                MessageBox.Show("data is not available to Transter for the period of " + intMonth.ToString() + "/" + intYear.ToString() + "  to Live Server ", "AAMS Admin", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return false;
            }

            if (optionaldoBulkTransfer == "0")
            {
                DataColumn colMonth = new DataColumn("MONTH", typeof(System.Int16));
                DataColumn colYear = new DataColumn("YEAR", typeof(System.Int16));
                DataColumn colType = new DataColumn("UPLOAD_TYPE", typeof(string));

                colMonth.DefaultValue = intMonth;
                colYear.DefaultValue = intYear;
                colType.DefaultValue = "HL";

                objBulkDT.Columns.Add(colMonth);
                objBulkDT.Columns.Add(colYear);
                objBulkDT.Columns.Add(colType);
                return false;
            }

            try
            {
                if (objConLivedatabase.State == ConnectionState.Open)
                {
                    objConLivedatabase.Close();
                }
                objConLivedatabase.Open();

                if (strExportType == "HL")
                {
                    // 1. Instantiate a new command with a query and connection
                    //cmd = new SqlCommand("SELECT YEAR , MONTH FROM T_INC_NIDT_PRODUCTIVITY_EXPORTED WHERE COUNTRY  = '" + StrCountry +"' AND  MONTH = "  + intMonth + " AND YEAR = " + intYear + " GROUP BY YEAR,MONTH", objConLivedatabase);
                    cmd = new SqlCommand("SELECT YEAR , MONTH FROM T_INC_NIDT_PRODUCTIVITY_EXPORTED WHERE COUNTRY  = '" + StrCountry + "' AND  MONTH = " + intMonth + " AND YEAR = " + intYear + " GROUP BY YEAR,MONTH", objConLivedatabase);
                }
                else if (strExportType == "HX")
                {
                    cmd = new SqlCommand("SELECT HX_BOOKINGS = SUM(HX_BOOKINGS)  FROM T_INC_NIDT_PRODUCTIVITY_EXPORTED WHERE COUNTRY  = '" + StrCountry + "' AND  MONTH = " + intMonth + " AND YEAR = " + intYear + " GROUP BY YEAR,MONTH HAVING ISNULL(SUM(HX_BOOKINGS) ,0) !=0 ", objConLivedatabase);
                }
                else if (strExportType == "CODD")
                {
                    cmd = new SqlCommand("SELECT CODD = SUM(CODD) FROM T_INC_NIDT_PRODUCTIVITY_EXPORTED WHERE COUNTRY  = '" + StrCountry + "' AND  MONTH = " + intMonth + " AND YEAR = " + intYear + " GROUP BY YEAR,MONTH HAVING ISNULL(SUM(CODD),0) !=0 ", objConLivedatabase);
                }


                // 2. Call Execute reader to get query results
                SqlDataReader rdr = cmd.ExecuteReader();
                if (rdr.HasRows == false)
                {
                    if (objConLivedatabase.State == ConnectionState.Open)
                    {
                        objConLivedatabase.Close();
                    }
                    objConLivedatabase.Open();

                    if (strExportType == "HL")
                    {
                        DataColumn colMonth = new DataColumn("MONTH", typeof(System.Int16));
                        DataColumn colYear = new DataColumn("YEAR", typeof(System.Int16));
                        DataColumn colType = new DataColumn("UPLOAD_TYPE", typeof(string));

                        colMonth.DefaultValue = intMonth;
                        colYear.DefaultValue = intYear;
                        colType.DefaultValue = "HL";

                        objBulkDT.Columns.Add(colMonth);
                        objBulkDT.Columns.Add(colYear);
                        objBulkDT.Columns.Add(colType);


                        intCount = objBulkDT.Rows.Count;

                        if (intCount > 1)
                        {
                            this.toolStripProgressBar1.Minimum = 0;
                            this.toolStripProgressBar1.Maximum = intCount;
                            this.toolStripProgressBar1.Value = 1;
                            this.toolStripProgressBar1.Value = this.toolStripProgressBar1.Value + 1;
                        }                       


                        SqlBulkCopyColumnMapping mapping1 = new SqlBulkCopyColumnMapping("YEAR", "YEAR");
                        SqlBulkCopyColumnMapping mapping2 = new SqlBulkCopyColumnMapping("MONTH", "MONTH");
                        SqlBulkCopyColumnMapping mapping3 = new SqlBulkCopyColumnMapping("LCODE", "LCODE");
                        SqlBulkCopyColumnMapping mapping4 = new SqlBulkCopyColumnMapping("CHAIN_CODE", "CHAIN_CODE");
                        SqlBulkCopyColumnMapping mapping5 = new SqlBulkCopyColumnMapping("CHAIN_NAME", "CHAIN_NAME");
                        SqlBulkCopyColumnMapping mapping6 = new SqlBulkCopyColumnMapping("OFFICEID", "OFFICEID");
                        SqlBulkCopyColumnMapping mapping7 = new SqlBulkCopyColumnMapping("COUNTRY", "COUNTRY");
                        SqlBulkCopyColumnMapping mapping8 = new SqlBulkCopyColumnMapping("PRODUCTIVITY_CODD_PK_HX", "PRODUCTIVITY_CODD_PK_HX");
                        SqlBulkCopyColumnMapping mapping9 = new SqlBulkCopyColumnMapping("INTL", "INTL");
                        SqlBulkCopyColumnMapping mapping10 = new SqlBulkCopyColumnMapping("TTL_HL", "TTL_HL");
                        SqlBulkCopyColumnMapping mapping11 = new SqlBulkCopyColumnMapping("HL_INTL", "HL_INTL");
                        SqlBulkCopyColumnMapping mapping12 = new SqlBulkCopyColumnMapping("S2_DOM_CODD_PK_HX", "S2_DOM_CODD_PK_HX");
                        SqlBulkCopyColumnMapping mapping13 = new SqlBulkCopyColumnMapping("S2_HL_NETSTATUS", "S2_HL_NETSTATUS");
                        SqlBulkCopyColumnMapping mapping14 = new SqlBulkCopyColumnMapping("IC_DOM_CODD_PK_HX", "IC_DOM_CODD_PK_HX");
                        SqlBulkCopyColumnMapping mapping15 = new SqlBulkCopyColumnMapping("IC_HL_NETSTATUS", "IC_HL_NETSTATUS");
                        SqlBulkCopyColumnMapping mapping16 = new SqlBulkCopyColumnMapping("9W_DOM_CODD_PK_HX", "9W_DOM_CODD_PK_HX");
                        SqlBulkCopyColumnMapping mapping17 = new SqlBulkCopyColumnMapping("9W_HL_NETSTATUS", "9W_HL_NETSTATUS");
                        SqlBulkCopyColumnMapping mapping18 = new SqlBulkCopyColumnMapping("AI_DOM_CODD_PK_HX", "AI_DOM_CODD_PK_HX");
                        SqlBulkCopyColumnMapping mapping19 = new SqlBulkCopyColumnMapping("AI_HL_NETSTATUS", "AI_HL_NETSTATUS");
                        SqlBulkCopyColumnMapping mapping20 = new SqlBulkCopyColumnMapping("IT_DOM_CODD_PK_HX", "IT_DOM_CODD_PK_HX");
                        SqlBulkCopyColumnMapping mapping21 = new SqlBulkCopyColumnMapping("IT_HL_NETSTATUS", "IT_HL_NETSTATUS");
                        SqlBulkCopyColumnMapping mapping22 = new SqlBulkCopyColumnMapping("ITRED_DOM_CODD_PK_HX", "ITRED_DOM_CODD_PK_HX");
                        SqlBulkCopyColumnMapping mapping23 = new SqlBulkCopyColumnMapping("ITRED_HL_NETSTATUS", "ITRED_HL_NETSTATUS");
                        SqlBulkCopyColumnMapping mapping24 = new SqlBulkCopyColumnMapping("I7_DOM_CODD_PK_HX", "I7_DOM_CODD_PK_HX");
                        SqlBulkCopyColumnMapping mapping25 = new SqlBulkCopyColumnMapping("I7_HL_NETSTATUS", "I7_HL_NETSTATUS");
                        SqlBulkCopyColumnMapping mapping26 = new SqlBulkCopyColumnMapping("TOTALPK", "TOTALPK");
                        SqlBulkCopyColumnMapping mapping27 = new SqlBulkCopyColumnMapping("DOM_PK_IC", "DOM_PK_IC");
                        SqlBulkCopyColumnMapping mapping28 = new SqlBulkCopyColumnMapping("DOM_PK_IT", "DOM_PK_IT");
                        SqlBulkCopyColumnMapping mapping29 = new SqlBulkCopyColumnMapping("DOM_PK_AI", "DOM_PK_AI");
                        SqlBulkCopyColumnMapping mapping30 = new SqlBulkCopyColumnMapping("DOM_PK_9W", "DOM_PK_9W");
                        SqlBulkCopyColumnMapping mapping31 = new SqlBulkCopyColumnMapping("CODD", "CODD");
                        SqlBulkCopyColumnMapping mapping32 = new SqlBulkCopyColumnMapping("ROI", "ROI");
                        SqlBulkCopyColumnMapping mapping33 = new SqlBulkCopyColumnMapping("S2_HX", "S2_HX");
                        SqlBulkCopyColumnMapping mapping34 = new SqlBulkCopyColumnMapping("IC_HX", "IC_HX");
                        SqlBulkCopyColumnMapping mapping35 = new SqlBulkCopyColumnMapping("9W_HX", "9W_HX");
                        SqlBulkCopyColumnMapping mapping36 = new SqlBulkCopyColumnMapping("AI_HX", "AI_HX");
                        SqlBulkCopyColumnMapping mapping37 = new SqlBulkCopyColumnMapping("IT_HX", "IT_HX");
                        SqlBulkCopyColumnMapping mapping38 = new SqlBulkCopyColumnMapping("I7_HX", "I7_HX");
                        SqlBulkCopyColumnMapping mapping39 = new SqlBulkCopyColumnMapping("HX_BOOKINGS", "HX_BOOKINGS");
                        SqlBulkCopyColumnMapping mapping40 = new SqlBulkCopyColumnMapping("DOM_PK_S2", "DOM_PK_S2");

                        // newly added as on dated 09/03/2015
                        SqlBulkCopyColumnMapping mapping41 = new SqlBulkCopyColumnMapping("UK_DOM_CODD_PK_HX", "UK_DOM_CODD_PK_HX");
                        SqlBulkCopyColumnMapping mapping42 = new SqlBulkCopyColumnMapping("UK_HL_NETSTATUS", "UK_HL_NETSTATUS");
                        SqlBulkCopyColumnMapping mapping43 = new SqlBulkCopyColumnMapping("DOM_PK_UK", "DOM_PK_UK");
                        SqlBulkCopyColumnMapping mapping44 = new SqlBulkCopyColumnMapping("UK_HX", "UK_HX");

                        SqlBulkCopyColumnMapping mapping45 = new SqlBulkCopyColumnMapping("UPLOAD_TYPE", "UPLOAD_TYPE");

                        objSqlTransaction = objConLivedatabase.BeginTransaction(IsolationLevel.RepeatableRead);

                        objSqlbulkCopy = new SqlBulkCopy(objConLivedatabase, SqlBulkCopyOptions.CheckConstraints, objSqlTransaction);
                        objSqlbulkCopy.SqlRowsCopied += new SqlRowsCopiedEventHandler(objSqlbulkCopy_SqlRowsCopied);

                        objSqlbulkCopy.ColumnMappings.Add(mapping1);
                        objSqlbulkCopy.ColumnMappings.Add(mapping2);
                        objSqlbulkCopy.ColumnMappings.Add(mapping3);
                        objSqlbulkCopy.ColumnMappings.Add(mapping4);
                        objSqlbulkCopy.ColumnMappings.Add(mapping5);
                        objSqlbulkCopy.ColumnMappings.Add(mapping6);
                        objSqlbulkCopy.ColumnMappings.Add(mapping7);
                        objSqlbulkCopy.ColumnMappings.Add(mapping8);
                        objSqlbulkCopy.ColumnMappings.Add(mapping9);
                        objSqlbulkCopy.ColumnMappings.Add(mapping10);
                        objSqlbulkCopy.ColumnMappings.Add(mapping11);
                        objSqlbulkCopy.ColumnMappings.Add(mapping12);
                        objSqlbulkCopy.ColumnMappings.Add(mapping13);
                        objSqlbulkCopy.ColumnMappings.Add(mapping14);
                        objSqlbulkCopy.ColumnMappings.Add(mapping15);
                        objSqlbulkCopy.ColumnMappings.Add(mapping16);
                        objSqlbulkCopy.ColumnMappings.Add(mapping17);
                        objSqlbulkCopy.ColumnMappings.Add(mapping18);
                        objSqlbulkCopy.ColumnMappings.Add(mapping19);
                        objSqlbulkCopy.ColumnMappings.Add(mapping20);
                        objSqlbulkCopy.ColumnMappings.Add(mapping21);
                        objSqlbulkCopy.ColumnMappings.Add(mapping22);
                        objSqlbulkCopy.ColumnMappings.Add(mapping23);
                        objSqlbulkCopy.ColumnMappings.Add(mapping24);
                        objSqlbulkCopy.ColumnMappings.Add(mapping25);
                        objSqlbulkCopy.ColumnMappings.Add(mapping26);
                        objSqlbulkCopy.ColumnMappings.Add(mapping27);
                        objSqlbulkCopy.ColumnMappings.Add(mapping28);
                        objSqlbulkCopy.ColumnMappings.Add(mapping29);
                        objSqlbulkCopy.ColumnMappings.Add(mapping30);
                        objSqlbulkCopy.ColumnMappings.Add(mapping31);
                        objSqlbulkCopy.ColumnMappings.Add(mapping32);
                        objSqlbulkCopy.ColumnMappings.Add(mapping33);
                        objSqlbulkCopy.ColumnMappings.Add(mapping34);
                        objSqlbulkCopy.ColumnMappings.Add(mapping35);
                        objSqlbulkCopy.ColumnMappings.Add(mapping36);
                        objSqlbulkCopy.ColumnMappings.Add(mapping37);
                        objSqlbulkCopy.ColumnMappings.Add(mapping38);
                        objSqlbulkCopy.ColumnMappings.Add(mapping39);
                        objSqlbulkCopy.ColumnMappings.Add(mapping40);
                        // newly added as on dated 09/03/2015
                        objSqlbulkCopy.ColumnMappings.Add(mapping41);
                        objSqlbulkCopy.ColumnMappings.Add(mapping42);
                        objSqlbulkCopy.ColumnMappings.Add(mapping43);
                        objSqlbulkCopy.ColumnMappings.Add(mapping44);
                        objSqlbulkCopy.ColumnMappings.Add(mapping45);

                        //do enable the below lines for Live Working
                        //objSqlbulkCopy.DestinationTableName = "T_INC_NIDT_PRODUCTIVITY_EXPORTED";
                        objSqlbulkCopy.DestinationTableName = "T_INC_NIDT_PRODUCTIVITY_EXPORTED";

                        objSqlbulkCopy.BatchSize = 1000;
                        objSqlbulkCopy.NotifyAfter = 5;
                        //objSqlbulkCopy.WriteToServer(objBulkDT);

                        DataTableReader reader = objBulkDT.CreateDataReader();

                        using (ValidatingDataReader validator = new ValidatingDataReader(reader, objConLivedatabase, objSqlbulkCopy, objSqlTransaction))
                        {
                            DisplayStatus("Transfering HL data...");
                            objSqlbulkCopy.WriteToServer(validator);
                        }

                        objSqlTransaction.Commit();
                        this.toolStripProgressBar1.Value = intCount;
                        objSqlTransaction = null;
                        DisplayStatus("HL data Transfered...");
                        return true;
                    }
                    else if (strExportType == "HX")
                    {
                        DataColumn colMonth = new DataColumn("MONTH", typeof(System.Int16));
                        DataColumn colYear = new DataColumn("YEAR", typeof(System.Int16));
                        DataColumn colType = new DataColumn("UPLOAD_TYPE", typeof(string));

                        colMonth.DefaultValue = intMonth;
                        colYear.DefaultValue = intYear;
                        colType.DefaultValue = "HX";

                        objBulkDT.Columns.Add(colMonth);
                        objBulkDT.Columns.Add(colYear);
                        objBulkDT.Columns.Add(colType);

                        intCount = objBulkDT.Rows.Count;

                        toolStripProgressBar1.Minimum = 0;
                        toolStripProgressBar1.Maximum = intCount;

                        SqlBulkCopyColumnMapping mapping1 = new SqlBulkCopyColumnMapping("YEAR", "YEAR");
                        SqlBulkCopyColumnMapping mapping2 = new SqlBulkCopyColumnMapping("MONTH", "MONTH");
                        SqlBulkCopyColumnMapping mapping3 = new SqlBulkCopyColumnMapping("LCODE", "LCODE");
                        SqlBulkCopyColumnMapping mapping4 = new SqlBulkCopyColumnMapping("CHAIN_CODE", "CHAIN_CODE");
                        SqlBulkCopyColumnMapping mapping5 = new SqlBulkCopyColumnMapping("CHAIN_NAME", "CHAIN_NAME");
                        SqlBulkCopyColumnMapping mapping6 = new SqlBulkCopyColumnMapping("OFFICEID", "OFFICEID");
                        SqlBulkCopyColumnMapping mapping7 = new SqlBulkCopyColumnMapping("COUNTRY", "COUNTRY");
                        SqlBulkCopyColumnMapping mapping8 = new SqlBulkCopyColumnMapping("HX_BOOKINGS", "HX_BOOKINGS");
                        SqlBulkCopyColumnMapping mapping9 = new SqlBulkCopyColumnMapping("UPLOAD_TYPE", "UPLOAD_TYPE");


                        objSqlTransaction = objConLivedatabase.BeginTransaction(IsolationLevel.RepeatableRead);
                        objSqlbulkCopy = new SqlBulkCopy(objConLivedatabase, SqlBulkCopyOptions.CheckConstraints, objSqlTransaction);
                        objSqlbulkCopy.SqlRowsCopied += new SqlRowsCopiedEventHandler(objSqlbulkCopy_SqlRowsCopied);

                        objSqlbulkCopy.ColumnMappings.Add(mapping1);
                        objSqlbulkCopy.ColumnMappings.Add(mapping2);
                        objSqlbulkCopy.ColumnMappings.Add(mapping3);
                        objSqlbulkCopy.ColumnMappings.Add(mapping4);
                        objSqlbulkCopy.ColumnMappings.Add(mapping5);
                        objSqlbulkCopy.ColumnMappings.Add(mapping6);
                        objSqlbulkCopy.ColumnMappings.Add(mapping7);
                        objSqlbulkCopy.ColumnMappings.Add(mapping8);
                        objSqlbulkCopy.ColumnMappings.Add(mapping9);


                        objSqlbulkCopy.DestinationTableName = "T_INC_NIDT_PRODUCTIVITY_EXPORTED";
                        objSqlbulkCopy.BatchSize = 1000;
                        objSqlbulkCopy.NotifyAfter = 5;
                        //objSqlbulkCopy.WriteToServer(objBulkDT);

                        //objSqlbulkCopy.WriteToServer(objBulkDT);
                        DataTableReader reader = objBulkDT.CreateDataReader();

                        using (ValidatingDataReader validator = new ValidatingDataReader(reader, objConLivedatabase, objSqlbulkCopy, objSqlTransaction))
                        {
                            objSqlbulkCopy.WriteToServer(validator);
                        }

                        objSqlTransaction.Commit();
                        objSqlTransaction = null;
                        return true;
                    }
                    else if (strExportType == "CODD")
                    {
                        DataColumn colMonth = new DataColumn("MONTH", typeof(System.Int16));
                        DataColumn colYear = new DataColumn("YEAR", typeof(System.Int16));
                        DataColumn colType = new DataColumn("UPLOAD_TYPE", typeof(string));

                        colMonth.DefaultValue = intMonth;
                        colYear.DefaultValue = intYear;
                        colType.DefaultValue = "CODD";

                        objBulkDT.Columns.Add(colMonth);
                        objBulkDT.Columns.Add(colYear);
                        objBulkDT.Columns.Add(colType);


                        intCount = objBulkDT.Rows.Count;
                        toolStripProgressBar1.Minimum = 0;
                        toolStripProgressBar1.Maximum = intCount;

                        SqlBulkCopyColumnMapping mapping1 = new SqlBulkCopyColumnMapping("[YEAR]", "[YEAR]");
                        SqlBulkCopyColumnMapping mapping2 = new SqlBulkCopyColumnMapping("[MONTH]", "[MONTH]");
                        SqlBulkCopyColumnMapping mapping3 = new SqlBulkCopyColumnMapping("[LCODE]", "[LCODE]");
                        SqlBulkCopyColumnMapping mapping4 = new SqlBulkCopyColumnMapping("[CHAIN_CODE]", "[CHAIN_CODE]");
                        SqlBulkCopyColumnMapping mapping5 = new SqlBulkCopyColumnMapping("[CHAIN_NAME]", "[CHAIN_NAME]");
                        SqlBulkCopyColumnMapping mapping6 = new SqlBulkCopyColumnMapping("[OFFICEID]", "[OFFICEID]");
                        SqlBulkCopyColumnMapping mapping7 = new SqlBulkCopyColumnMapping("[COUNTRY]", "[COUNTRY]");
                        SqlBulkCopyColumnMapping mapping8 = new SqlBulkCopyColumnMapping("[CODD]", "[CODD]");
                        SqlBulkCopyColumnMapping mapping9 = new SqlBulkCopyColumnMapping("[UPLOAD_TYPE]", "[UPLOAD_TYPE]");


                        objSqlTransaction = objConLivedatabase.BeginTransaction(IsolationLevel.RepeatableRead);

                        objSqlbulkCopy = new SqlBulkCopy(objConLivedatabase, SqlBulkCopyOptions.CheckConstraints, objSqlTransaction);
                        objSqlbulkCopy.SqlRowsCopied += new SqlRowsCopiedEventHandler(objSqlbulkCopy_SqlRowsCopied);

                        objSqlbulkCopy.ColumnMappings.Add(mapping1);
                        objSqlbulkCopy.ColumnMappings.Add(mapping2);
                        objSqlbulkCopy.ColumnMappings.Add(mapping3);
                        objSqlbulkCopy.ColumnMappings.Add(mapping4);
                        objSqlbulkCopy.ColumnMappings.Add(mapping5);
                        objSqlbulkCopy.ColumnMappings.Add(mapping6);
                        objSqlbulkCopy.ColumnMappings.Add(mapping7);
                        objSqlbulkCopy.ColumnMappings.Add(mapping8);
                        objSqlbulkCopy.ColumnMappings.Add(mapping9);


                        objSqlbulkCopy.DestinationTableName = "T_INC_NIDT_PRODUCTIVITY_EXPORTED";
                        objSqlbulkCopy.BatchSize = 1000;
                        objSqlbulkCopy.NotifyAfter = 5;

                        //objSqlbulkCopy.WriteToServer(objBulkDT);
                        DataTableReader reader = objBulkDT.CreateDataReader();

                        using (ValidatingDataReader validator = new ValidatingDataReader(reader, objConLivedatabase, objSqlbulkCopy, objSqlTransaction))
                        {
                            objSqlbulkCopy.WriteToServer(validator);
                        }

                        objSqlTransaction.Commit();
                        objSqlTransaction = null;
                        return true;
                    }
                }
                else
                {
                    DataColumn colMonth = new DataColumn("MONTH", typeof(System.Int16));
                    DataColumn colYear = new DataColumn("YEAR", typeof(System.Int16));

                    colMonth.DefaultValue = intMonth;
                    colYear.DefaultValue = intYear;

                    objBulkDT.Columns.Add(colMonth);
                    objBulkDT.Columns.Add(colYear);

                    if (strExportType == "HL")
                    {
                        DataColumn colType = new DataColumn("UPLOAD_TYPE", typeof(string));
                        colType.DefaultValue = "HL";
                        objBulkDT.Columns.Add(colType);
                    }
                    else if (strExportType == "HX")
                    {
                        DataColumn colType = new DataColumn("UPLOAD_TYPE", typeof(string));
                        colType.DefaultValue = "HX";
                        objBulkDT.Columns.Add(colType);
                    }
                    else if (strExportType == "CODD")
                    {
                        DataColumn colType = new DataColumn("UPLOAD_TYPE", typeof(string));
                        colType.DefaultValue = "CODD";
                        objBulkDT.Columns.Add(colType);
                    }

                    intCount = objBulkDT.Rows.Count;

                    toolStripProgressBar1.Minimum = 0;
                    toolStripProgressBar1.Maximum = intCount;

                    DialogResult dlgResult = MessageBox.Show("NIDT data for the month/year [" + intMonth.ToString() + "/" + intYear.ToString() + "] already exists in Live Server database."+ "\n\n" + "do you want to continue Upload data into Live Server ? ", "AAMS Admin", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);

                    if (dlgResult == DialogResult.Yes)
                    {
                        Cursor.Current = Cursors.WaitCursor;
                        if (objConLivedatabase.State == ConnectionState.Open)
                        {
                            objConLivedatabase.Close();
                        }
                        objConLivedatabase.Open();

                        if (strExportType == "HL")
                        {
                            // INSERT LOG BEFORE DELETE FROM T_INC_NIDT_PRODUCTIVITY_EXPORTED TABLE
                            // 1. Instantiate a new command with a query and connection
                            // HL DATA OF INDIA CONTAINS THREE COUNTRY DATA INCLUDING HL,HX,CODD 'India,Bhutan,Tba'
                            //cmd1 = new SqlCommand("DELETE FROM T_INC_NIDT_PRODUCTIVITY_EXPORTED OUTPUT  DELETED.[YEAR],DELETED.[MONTH],DELETED.[LCODE],DELETED.[CHAIN_CODE],DELETED.[CHAIN_NAME],DELETED.[OFFICEID],DELETED.[COUNTRY],DELETED.[PRODUCTIVITY_CODD_PK_HX],DELETED.[INTL],DELETED.[TTL_HL],DELETED.[HL_INTL],DELETED.[S2_DOM_CODD_PK_HX],DELETED.[S2_HL_NETSTATUS],DELETED.[IC_DOM_CODD_PK_HX],DELETED.[IC_HL_NETSTATUS],DELETED.[9W_DOM_CODD_PK_HX],DELETED.[9W_HL_NETSTATUS],DELETED.[AI_DOM_CODD_PK_HX],DELETED.[AI_HL_NETSTATUS],DELETED.[IT_DOM_CODD_PK_HX],DELETED.[IT_HL_NETSTATUS],DELETED.[ITRED_DOM_CODD_PK_HX],DELETED.[ITRED_HL_NETSTATUS],DELETED.[I7_DOM_CODD_PK_HX],DELETED.[I7_HL_NETSTATUS],DELETED.[TOTALPK],DELETED.[DOM_PK_IC],DELETED.[DOM_PK_IT],DELETED.[DOM_PK_AI],DELETED.[DOM_PK_9W],DELETED.[CODD],DELETED.[ROI],DELETED.[S2_HX],DELETED.[IC_HX],DELETED.[9W_HX],DELETED.[AI_HX],DELETED.[IT_HX],DELETED.[I7_HX],DELETED.[HX_BOOKINGS],DELETED.[DOM_PK_S2] INTO T_INC_NIDT_PRODUCTIVITY_EXPORTED_LOG ([YEAR],[MONTH],[LCODE],[CHAIN_CODE],[CHAIN_NAME],[OFFICEID],[COUNTRY],[PRODUCTIVITY_CODD_PK_HX],[INTL],[TTL_HL],[HL_INTL],[S2_DOM_CODD_PK_HX],[S2_HL_NETSTATUS],[IC_DOM_CODD_PK_HX],[IC_HL_NETSTATUS],[9W_DOM_CODD_PK_HX],[9W_HL_NETSTATUS],[AI_DOM_CODD_PK_HX],[AI_HL_NETSTATUS],[IT_DOM_CODD_PK_HX],[IT_HL_NETSTATUS],[ITRED_DOM_CODD_PK_HX],[ITRED_HL_NETSTATUS],[I7_DOM_CODD_PK_HX],[I7_HL_NETSTATUS],[TOTALPK],[DOM_PK_IC],[DOM_PK_IT],[DOM_PK_AI],[DOM_PK_9W],[CODD],[ROI],[S2_HX],[IC_HX],[9W_HX],[AI_HX],[IT_HX],[I7_HX],[HX_BOOKINGS],[DOM_PK_S2]) WHERE COUNTRY IN ('India','Bhutan','Tba') AND  MONTH = " + intMonth + " AND YEAR = " + intYear, objConLivedatabase);
                            cmd1 = new SqlCommand("DELETE FROM T_INC_NIDT_PRODUCTIVITY_EXPORTED OUTPUT  DELETED.[YEAR],DELETED.[MONTH],DELETED.[LCODE],DELETED.[CHAIN_CODE],DELETED.[CHAIN_NAME],DELETED.[OFFICEID],DELETED.[COUNTRY],DELETED.[PRODUCTIVITY_CODD_PK_HX],DELETED.[INTL],DELETED.[TTL_HL],DELETED.[HL_INTL],DELETED.[S2_DOM_CODD_PK_HX],DELETED.[S2_HL_NETSTATUS],DELETED.[IC_DOM_CODD_PK_HX],DELETED.[IC_HL_NETSTATUS],DELETED.[9W_DOM_CODD_PK_HX],DELETED.[9W_HL_NETSTATUS],DELETED.[AI_DOM_CODD_PK_HX],DELETED.[AI_HL_NETSTATUS],DELETED.[IT_DOM_CODD_PK_HX],DELETED.[IT_HL_NETSTATUS],DELETED.[ITRED_DOM_CODD_PK_HX],DELETED.[ITRED_HL_NETSTATUS],DELETED.[I7_DOM_CODD_PK_HX],DELETED.[I7_HL_NETSTATUS],DELETED.[TOTALPK],DELETED.[DOM_PK_IC],DELETED.[DOM_PK_IT],DELETED.[DOM_PK_AI],DELETED.[DOM_PK_9W],DELETED.[CODD],DELETED.[ROI],DELETED.[S2_HX],DELETED.[IC_HX],DELETED.[9W_HX],DELETED.[AI_HX],DELETED.[IT_HX],DELETED.[I7_HX],DELETED.[HX_BOOKINGS],DELETED.[DOM_PK_S2],DELETED.[UK_DOM_CODD_PK_HX],DELETED.[UK_HL_NETSTATUS],DELETED.[DOM_PK_UK],DELETED.[UK_HX],DELETED.UPLOAD_TYPE ,'" + strIpaddress + "' INTO T_INC_NIDT_PRODUCTIVITY_EXPORTED_LOG ([YEAR],[MONTH],[LCODE],[CHAIN_CODE],[CHAIN_NAME],[OFFICEID],[COUNTRY],[PRODUCTIVITY_CODD_PK_HX],[INTL],[TTL_HL],[HL_INTL],[S2_DOM_CODD_PK_HX],[S2_HL_NETSTATUS],[IC_DOM_CODD_PK_HX],[IC_HL_NETSTATUS],[9W_DOM_CODD_PK_HX],[9W_HL_NETSTATUS],[AI_DOM_CODD_PK_HX],[AI_HL_NETSTATUS],[IT_DOM_CODD_PK_HX],[IT_HL_NETSTATUS],[ITRED_DOM_CODD_PK_HX],[ITRED_HL_NETSTATUS],[I7_DOM_CODD_PK_HX],[I7_HL_NETSTATUS],[TOTALPK],[DOM_PK_IC],[DOM_PK_IT],[DOM_PK_AI],[DOM_PK_9W],[CODD],[ROI],[S2_HX],[IC_HX],[9W_HX],[AI_HX],[IT_HX],[I7_HX],[HX_BOOKINGS],[DOM_PK_S2],[UK_DOM_CODD_PK_HX],[UK_HL_NETSTATUS],[DOM_PK_UK],[UK_HX],UPLOAD_TYPE,IPADDRESS) WHERE UPLOAD_TYPE = 'HL' AND  MONTH = " + intMonth + " AND YEAR = " + intYear, objConLivedatabase);
                            // 2. Call Execute query 
                            int intRowaffected = cmd1.ExecuteNonQuery();

                            //INSERT HL DATA AFTER DELETE QUERY
                            intCount = objBulkDT.Rows.Count;

                            if (intCount > 1)
                            {
                                this.toolStripProgressBar1.Minimum = 0;
                                this.toolStripProgressBar1.Maximum = intCount;
                                this.toolStripProgressBar1.Value = 1;
                                this.toolStripProgressBar1.Value = this.toolStripProgressBar1.Value + 1;
                            }

                            SqlBulkCopyColumnMapping mapping1 = new SqlBulkCopyColumnMapping("YEAR", "YEAR");
                            SqlBulkCopyColumnMapping mapping2 = new SqlBulkCopyColumnMapping("MONTH", "MONTH");
                            SqlBulkCopyColumnMapping mapping3 = new SqlBulkCopyColumnMapping("LCODE", "LCODE");
                            SqlBulkCopyColumnMapping mapping4 = new SqlBulkCopyColumnMapping("CHAIN_CODE", "CHAIN_CODE");
                            SqlBulkCopyColumnMapping mapping5 = new SqlBulkCopyColumnMapping("CHAIN_NAME", "CHAIN_NAME");
                            SqlBulkCopyColumnMapping mapping6 = new SqlBulkCopyColumnMapping("OFFICEID", "OFFICEID");
                            SqlBulkCopyColumnMapping mapping7 = new SqlBulkCopyColumnMapping("COUNTRY", "COUNTRY");
                            SqlBulkCopyColumnMapping mapping8 = new SqlBulkCopyColumnMapping("PRODUCTIVITY_CODD_PK_HX", "PRODUCTIVITY_CODD_PK_HX");
                            SqlBulkCopyColumnMapping mapping9 = new SqlBulkCopyColumnMapping("INTL", "INTL");
                            SqlBulkCopyColumnMapping mapping10 = new SqlBulkCopyColumnMapping("TTL_HL", "TTL_HL");
                            SqlBulkCopyColumnMapping mapping11 = new SqlBulkCopyColumnMapping("HL_INTL", "HL_INTL");
                            SqlBulkCopyColumnMapping mapping12 = new SqlBulkCopyColumnMapping("S2_DOM_CODD_PK_HX", "S2_DOM_CODD_PK_HX");
                            SqlBulkCopyColumnMapping mapping13 = new SqlBulkCopyColumnMapping("S2_HL_NETSTATUS", "S2_HL_NETSTATUS");
                            SqlBulkCopyColumnMapping mapping14 = new SqlBulkCopyColumnMapping("IC_DOM_CODD_PK_HX", "IC_DOM_CODD_PK_HX");
                            SqlBulkCopyColumnMapping mapping15 = new SqlBulkCopyColumnMapping("IC_HL_NETSTATUS", "IC_HL_NETSTATUS");
                            SqlBulkCopyColumnMapping mapping16 = new SqlBulkCopyColumnMapping("9W_DOM_CODD_PK_HX", "9W_DOM_CODD_PK_HX");
                            SqlBulkCopyColumnMapping mapping17 = new SqlBulkCopyColumnMapping("9W_HL_NETSTATUS", "9W_HL_NETSTATUS");
                            SqlBulkCopyColumnMapping mapping18 = new SqlBulkCopyColumnMapping("AI_DOM_CODD_PK_HX", "AI_DOM_CODD_PK_HX");
                            SqlBulkCopyColumnMapping mapping19 = new SqlBulkCopyColumnMapping("AI_HL_NETSTATUS", "AI_HL_NETSTATUS");
                            SqlBulkCopyColumnMapping mapping20 = new SqlBulkCopyColumnMapping("IT_DOM_CODD_PK_HX", "IT_DOM_CODD_PK_HX");
                            SqlBulkCopyColumnMapping mapping21 = new SqlBulkCopyColumnMapping("IT_HL_NETSTATUS", "IT_HL_NETSTATUS");
                            SqlBulkCopyColumnMapping mapping22 = new SqlBulkCopyColumnMapping("ITRED_DOM_CODD_PK_HX", "ITRED_DOM_CODD_PK_HX");
                            SqlBulkCopyColumnMapping mapping23 = new SqlBulkCopyColumnMapping("ITRED_HL_NETSTATUS", "ITRED_HL_NETSTATUS");
                            SqlBulkCopyColumnMapping mapping24 = new SqlBulkCopyColumnMapping("I7_DOM_CODD_PK_HX", "I7_DOM_CODD_PK_HX");
                            SqlBulkCopyColumnMapping mapping25 = new SqlBulkCopyColumnMapping("I7_HL_NETSTATUS", "I7_HL_NETSTATUS");
                            SqlBulkCopyColumnMapping mapping26 = new SqlBulkCopyColumnMapping("TOTALPK", "TOTALPK");
                            SqlBulkCopyColumnMapping mapping27 = new SqlBulkCopyColumnMapping("DOM_PK_IC", "DOM_PK_IC");
                            SqlBulkCopyColumnMapping mapping28 = new SqlBulkCopyColumnMapping("DOM_PK_IT", "DOM_PK_IT");
                            SqlBulkCopyColumnMapping mapping29 = new SqlBulkCopyColumnMapping("DOM_PK_AI", "DOM_PK_AI");
                            SqlBulkCopyColumnMapping mapping30 = new SqlBulkCopyColumnMapping("DOM_PK_9W", "DOM_PK_9W");
                            SqlBulkCopyColumnMapping mapping31 = new SqlBulkCopyColumnMapping("CODD", "CODD");
                            SqlBulkCopyColumnMapping mapping32 = new SqlBulkCopyColumnMapping("ROI", "ROI");
                            SqlBulkCopyColumnMapping mapping33 = new SqlBulkCopyColumnMapping("S2_HX", "S2_HX");
                            SqlBulkCopyColumnMapping mapping34 = new SqlBulkCopyColumnMapping("IC_HX", "IC_HX");
                            SqlBulkCopyColumnMapping mapping35 = new SqlBulkCopyColumnMapping("9W_HX", "9W_HX");
                            SqlBulkCopyColumnMapping mapping36 = new SqlBulkCopyColumnMapping("AI_HX", "AI_HX");
                            SqlBulkCopyColumnMapping mapping37 = new SqlBulkCopyColumnMapping("IT_HX", "IT_HX");
                            SqlBulkCopyColumnMapping mapping38 = new SqlBulkCopyColumnMapping("I7_HX", "I7_HX");
                            SqlBulkCopyColumnMapping mapping39 = new SqlBulkCopyColumnMapping("HX_BOOKINGS", "HX_BOOKINGS");
                            SqlBulkCopyColumnMapping mapping40 = new SqlBulkCopyColumnMapping("DOM_PK_S2", "DOM_PK_S2");
                            // newly added as on dated 09/03/2015
                            SqlBulkCopyColumnMapping mapping41 = new SqlBulkCopyColumnMapping("UK_DOM_CODD_PK_HX", "UK_DOM_CODD_PK_HX");
                            SqlBulkCopyColumnMapping mapping42 = new SqlBulkCopyColumnMapping("UK_HL_NETSTATUS", "UK_HL_NETSTATUS");
                            SqlBulkCopyColumnMapping mapping43 = new SqlBulkCopyColumnMapping("DOM_PK_UK", "DOM_PK_UK");
                            SqlBulkCopyColumnMapping mapping44 = new SqlBulkCopyColumnMapping("UK_HX", "UK_HX");
                            SqlBulkCopyColumnMapping mapping45 = new SqlBulkCopyColumnMapping("UPLOAD_TYPE", "UPLOAD_TYPE");

                            objSqlTransaction = objConLivedatabase.BeginTransaction(IsolationLevel.RepeatableRead);

                            objSqlbulkCopy = new SqlBulkCopy(objConLivedatabase, SqlBulkCopyOptions.CheckConstraints, objSqlTransaction);
                            objSqlbulkCopy.SqlRowsCopied += new SqlRowsCopiedEventHandler(objSqlbulkCopy_SqlRowsCopied);

                            objSqlbulkCopy.ColumnMappings.Add(mapping1);
                            objSqlbulkCopy.ColumnMappings.Add(mapping2);
                            objSqlbulkCopy.ColumnMappings.Add(mapping3);
                            objSqlbulkCopy.ColumnMappings.Add(mapping4);
                            objSqlbulkCopy.ColumnMappings.Add(mapping5);
                            objSqlbulkCopy.ColumnMappings.Add(mapping6);
                            objSqlbulkCopy.ColumnMappings.Add(mapping7);
                            objSqlbulkCopy.ColumnMappings.Add(mapping8);
                            objSqlbulkCopy.ColumnMappings.Add(mapping9);
                            objSqlbulkCopy.ColumnMappings.Add(mapping10);
                            objSqlbulkCopy.ColumnMappings.Add(mapping11);
                            objSqlbulkCopy.ColumnMappings.Add(mapping12);
                            objSqlbulkCopy.ColumnMappings.Add(mapping13);
                            objSqlbulkCopy.ColumnMappings.Add(mapping14);
                            objSqlbulkCopy.ColumnMappings.Add(mapping15);
                            objSqlbulkCopy.ColumnMappings.Add(mapping16);
                            objSqlbulkCopy.ColumnMappings.Add(mapping17);
                            objSqlbulkCopy.ColumnMappings.Add(mapping18);
                            objSqlbulkCopy.ColumnMappings.Add(mapping19);
                            objSqlbulkCopy.ColumnMappings.Add(mapping20);
                            objSqlbulkCopy.ColumnMappings.Add(mapping21);
                            objSqlbulkCopy.ColumnMappings.Add(mapping22);
                            objSqlbulkCopy.ColumnMappings.Add(mapping23);
                            objSqlbulkCopy.ColumnMappings.Add(mapping24);
                            objSqlbulkCopy.ColumnMappings.Add(mapping25);
                            objSqlbulkCopy.ColumnMappings.Add(mapping26);
                            objSqlbulkCopy.ColumnMappings.Add(mapping27);
                            objSqlbulkCopy.ColumnMappings.Add(mapping28);
                            objSqlbulkCopy.ColumnMappings.Add(mapping29);
                            objSqlbulkCopy.ColumnMappings.Add(mapping30);
                            objSqlbulkCopy.ColumnMappings.Add(mapping31);
                            objSqlbulkCopy.ColumnMappings.Add(mapping32);
                            objSqlbulkCopy.ColumnMappings.Add(mapping33);
                            objSqlbulkCopy.ColumnMappings.Add(mapping34);
                            objSqlbulkCopy.ColumnMappings.Add(mapping35);
                            objSqlbulkCopy.ColumnMappings.Add(mapping36);
                            objSqlbulkCopy.ColumnMappings.Add(mapping37);
                            objSqlbulkCopy.ColumnMappings.Add(mapping38);
                            objSqlbulkCopy.ColumnMappings.Add(mapping39);
                            objSqlbulkCopy.ColumnMappings.Add(mapping40);
                            // newly added as on dated 09/03/2015
                            objSqlbulkCopy.ColumnMappings.Add(mapping41);
                            objSqlbulkCopy.ColumnMappings.Add(mapping42);
                            objSqlbulkCopy.ColumnMappings.Add(mapping43);
                            objSqlbulkCopy.ColumnMappings.Add(mapping44);
                            objSqlbulkCopy.ColumnMappings.Add(mapping45);

                            objSqlbulkCopy.DestinationTableName = "T_INC_NIDT_PRODUCTIVITY_EXPORTED";
                            objSqlbulkCopy.BatchSize = 1000;
                            objSqlbulkCopy.NotifyAfter = 5;
                            //objSqlbulkCopy.WriteToServer(objBulkDT);

                            DataTableReader reader = objBulkDT.CreateDataReader();

                            using (ValidatingDataReader validator = new ValidatingDataReader(reader, objConLivedatabase, objSqlbulkCopy, objSqlTransaction))
                            {
                                DisplayStatus("Transfering HL data...");
                                objSqlbulkCopy.WriteToServer(validator);
                            }

                            objSqlTransaction.Commit();
                            this.toolStripProgressBar1.Value = intCount;
                            objSqlTransaction = null;
                            DisplayStatus("HL data Transfered...");
                            MessageBox.Show("HL data for the period of " + intMonth.ToString() + "/" + intYear.ToString() + "  successfully transfered to Live Server database ", "AAMS Admin", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            //END INSERT HL DATA
                        }
                        else if (strExportType == "HX")
                        {
                            cmd1 = new SqlCommand("DELETE FROM T_INC_NIDT_PRODUCTIVITY_EXPORTED OUTPUT  DELETED.[YEAR],DELETED.[MONTH],DELETED.[LCODE],DELETED.[CHAIN_CODE],DELETED.[CHAIN_NAME],DELETED.[OFFICEID],DELETED.[COUNTRY],DELETED.[PRODUCTIVITY_CODD_PK_HX],DELETED.[INTL],DELETED.[TTL_HL],DELETED.[HL_INTL],DELETED.[S2_DOM_CODD_PK_HX],DELETED.[S2_HL_NETSTATUS],DELETED.[IC_DOM_CODD_PK_HX],DELETED.[IC_HL_NETSTATUS],DELETED.[9W_DOM_CODD_PK_HX],DELETED.[9W_HL_NETSTATUS],DELETED.[AI_DOM_CODD_PK_HX],DELETED.[AI_HL_NETSTATUS],DELETED.[IT_DOM_CODD_PK_HX],DELETED.[IT_HL_NETSTATUS],DELETED.[ITRED_DOM_CODD_PK_HX],DELETED.[ITRED_HL_NETSTATUS],DELETED.[I7_DOM_CODD_PK_HX],DELETED.[I7_HL_NETSTATUS],DELETED.[TOTALPK],DELETED.[DOM_PK_IC],DELETED.[DOM_PK_IT],DELETED.[DOM_PK_AI],DELETED.[DOM_PK_9W],DELETED.[CODD],DELETED.[ROI],DELETED.[S2_HX],DELETED.[IC_HX],DELETED.[9W_HX],DELETED.[AI_HX],DELETED.[IT_HX],DELETED.[I7_HX],DELETED.[HX_BOOKINGS],DELETED.[DOM_PK_S2],DELETED.[UK_DOM_CODD_PK_HX],DELETED.[UK_HL_NETSTATUS],DELETED.[DOM_PK_UK],DELETED.[UK_HX],DELETED.UPLOAD_TYPE ,'" + strIpaddress + "' INTO T_INC_NIDT_PRODUCTIVITY_EXPORTED_LOG ([YEAR],[MONTH],[LCODE],[CHAIN_CODE],[CHAIN_NAME],[OFFICEID],[COUNTRY],[PRODUCTIVITY_CODD_PK_HX],[INTL],[TTL_HL],[HL_INTL],[S2_DOM_CODD_PK_HX],[S2_HL_NETSTATUS],[IC_DOM_CODD_PK_HX],[IC_HL_NETSTATUS],[9W_DOM_CODD_PK_HX],[9W_HL_NETSTATUS],[AI_DOM_CODD_PK_HX],[AI_HL_NETSTATUS],[IT_DOM_CODD_PK_HX],[IT_HL_NETSTATUS],[ITRED_DOM_CODD_PK_HX],[ITRED_HL_NETSTATUS],[I7_DOM_CODD_PK_HX],[I7_HL_NETSTATUS],[TOTALPK],[DOM_PK_IC],[DOM_PK_IT],[DOM_PK_AI],[DOM_PK_9W],[CODD],[ROI],[S2_HX],[IC_HX],[9W_HX],[AI_HX],[IT_HX],[I7_HX],[HX_BOOKINGS],[DOM_PK_S2],[UK_DOM_CODD_PK_HX],[UK_HL_NETSTATUS],[DOM_PK_UK],[UK_HX],UPLOAD_TYPE,IPADDRESS) WHERE UPLOAD_TYPE = 'HX' AND COUNTRY  = '" + StrCountry + "' AND  MONTH = " + intMonth + " AND YEAR = " + intYear, objConLivedatabase);
                            // 2. Call Execute query 
                            int intRowaffected = cmd1.ExecuteNonQuery();

                            //INSERT HX DATA AFTER DELETE QUUERY
                            intCount = objBulkDT.Rows.Count;
                            toolStripProgressBar1.Minimum = 0;
                            toolStripProgressBar1.Maximum = intCount;

                            SqlBulkCopyColumnMapping mapping1 = new SqlBulkCopyColumnMapping("YEAR", "YEAR");
                            SqlBulkCopyColumnMapping mapping2 = new SqlBulkCopyColumnMapping("MONTH", "MONTH");
                            SqlBulkCopyColumnMapping mapping3 = new SqlBulkCopyColumnMapping("LCODE", "LCODE");
                            SqlBulkCopyColumnMapping mapping4 = new SqlBulkCopyColumnMapping("CHAIN_CODE", "CHAIN_CODE");
                            SqlBulkCopyColumnMapping mapping5 = new SqlBulkCopyColumnMapping("CHAIN_NAME", "CHAIN_NAME");
                            SqlBulkCopyColumnMapping mapping6 = new SqlBulkCopyColumnMapping("OFFICEID", "OFFICEID");
                            SqlBulkCopyColumnMapping mapping7 = new SqlBulkCopyColumnMapping("COUNTRY", "COUNTRY");
                            SqlBulkCopyColumnMapping mapping8 = new SqlBulkCopyColumnMapping("HX_BOOKINGS", "HX_BOOKINGS");
                            SqlBulkCopyColumnMapping mapping9 = new SqlBulkCopyColumnMapping("UPLOAD_TYPE", "UPLOAD_TYPE");


                            objSqlTransaction = objConLivedatabase.BeginTransaction(IsolationLevel.RepeatableRead);
                            objSqlbulkCopy = new SqlBulkCopy(objConLivedatabase, SqlBulkCopyOptions.CheckConstraints, objSqlTransaction);
                            objSqlbulkCopy.SqlRowsCopied += new SqlRowsCopiedEventHandler(objSqlbulkCopy_SqlRowsCopied);

                            objSqlbulkCopy.ColumnMappings.Add(mapping1);
                            objSqlbulkCopy.ColumnMappings.Add(mapping2);
                            objSqlbulkCopy.ColumnMappings.Add(mapping3);
                            objSqlbulkCopy.ColumnMappings.Add(mapping4);
                            objSqlbulkCopy.ColumnMappings.Add(mapping5);
                            objSqlbulkCopy.ColumnMappings.Add(mapping6);
                            objSqlbulkCopy.ColumnMappings.Add(mapping7);
                            objSqlbulkCopy.ColumnMappings.Add(mapping8);
                            objSqlbulkCopy.ColumnMappings.Add(mapping9);


                            objSqlbulkCopy.DestinationTableName = "T_INC_NIDT_PRODUCTIVITY_EXPORTED";
                            objSqlbulkCopy.BatchSize = 1000;
                            objSqlbulkCopy.NotifyAfter = 5;
                            //objSqlbulkCopy.WriteToServer(objBulkDT);

                            //objSqlbulkCopy.WriteToServer(objBulkDT);
                            DataTableReader reader = objBulkDT.CreateDataReader();

                            using (ValidatingDataReader validator = new ValidatingDataReader(reader, objConLivedatabase, objSqlbulkCopy, objSqlTransaction))
                            {
                                objSqlbulkCopy.WriteToServer(validator);
                            }

                            objSqlTransaction.Commit();
                            objSqlTransaction = null;
                            MessageBox.Show("HX data for the period of " + intMonth.ToString() + "/" + intYear.ToString() + "  successfully transfered to Live Server database ", "AAMS Admin", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            //END INSERT HX DATA
                        }
                        else if (strExportType == "CODD")
                        {
                            cmd1 = new SqlCommand("DELETE FROM T_INC_NIDT_PRODUCTIVITY_EXPORTED OUTPUT  DELETED.[YEAR],DELETED.[MONTH],DELETED.[LCODE],DELETED.[CHAIN_CODE],DELETED.[CHAIN_NAME],DELETED.[OFFICEID],DELETED.[COUNTRY],DELETED.[PRODUCTIVITY_CODD_PK_HX],DELETED.[INTL],DELETED.[TTL_HL],DELETED.[HL_INTL],DELETED.[S2_DOM_CODD_PK_HX],DELETED.[S2_HL_NETSTATUS],DELETED.[IC_DOM_CODD_PK_HX],DELETED.[IC_HL_NETSTATUS],DELETED.[9W_DOM_CODD_PK_HX],DELETED.[9W_HL_NETSTATUS],DELETED.[AI_DOM_CODD_PK_HX],DELETED.[AI_HL_NETSTATUS],DELETED.[IT_DOM_CODD_PK_HX],DELETED.[IT_HL_NETSTATUS],DELETED.[ITRED_DOM_CODD_PK_HX],DELETED.[ITRED_HL_NETSTATUS],DELETED.[I7_DOM_CODD_PK_HX],DELETED.[I7_HL_NETSTATUS],DELETED.[TOTALPK],DELETED.[DOM_PK_IC],DELETED.[DOM_PK_IT],DELETED.[DOM_PK_AI],DELETED.[DOM_PK_9W],DELETED.[CODD],DELETED.[ROI],DELETED.[S2_HX],DELETED.[IC_HX],DELETED.[9W_HX],DELETED.[AI_HX],DELETED.[IT_HX],DELETED.[I7_HX],DELETED.[HX_BOOKINGS],DELETED.[DOM_PK_S2],DELETED.[UK_DOM_CODD_PK_HX],DELETED.[UK_HL_NETSTATUS],DELETED.[DOM_PK_UK],DELETED.[UK_HX],DELETED.UPLOAD_TYPE ,'" + strIpaddress + "' INTO T_INC_NIDT_PRODUCTIVITY_EXPORTED_LOG ([YEAR],[MONTH],[LCODE],[CHAIN_CODE],[CHAIN_NAME],[OFFICEID],[COUNTRY],[PRODUCTIVITY_CODD_PK_HX],[INTL],[TTL_HL],[HL_INTL],[S2_DOM_CODD_PK_HX],[S2_HL_NETSTATUS],[IC_DOM_CODD_PK_HX],[IC_HL_NETSTATUS],[9W_DOM_CODD_PK_HX],[9W_HL_NETSTATUS],[AI_DOM_CODD_PK_HX],[AI_HL_NETSTATUS],[IT_DOM_CODD_PK_HX],[IT_HL_NETSTATUS],[ITRED_DOM_CODD_PK_HX],[ITRED_HL_NETSTATUS],[I7_DOM_CODD_PK_HX],[I7_HL_NETSTATUS],[TOTALPK],[DOM_PK_IC],[DOM_PK_IT],[DOM_PK_AI],[DOM_PK_9W],[CODD],[ROI],[S2_HX],[IC_HX],[9W_HX],[AI_HX],[IT_HX],[I7_HX],[HX_BOOKINGS],[DOM_PK_S2],[UK_DOM_CODD_PK_HX],[UK_HL_NETSTATUS],[DOM_PK_UK],[UK_HX],UPLOAD_TYPE,IPADDRESS) WHERE UPLOAD_TYPE = 'CODD' AND COUNTRY  = '" + StrCountry + "' AND  MONTH = " + intMonth + " AND YEAR = " + intYear, objConLivedatabase);
                            // 2. Call Execute query 
                            int intRowaffected = cmd1.ExecuteNonQuery();

                            //INSERT CODD DATA AFTER DELETE QUERY
                            intCount = objBulkDT.Rows.Count;
                            toolStripProgressBar1.Minimum = 0;
                            toolStripProgressBar1.Maximum = intCount;

                            SqlBulkCopyColumnMapping mapping1 = new SqlBulkCopyColumnMapping("[YEAR]", "[YEAR]");
                            SqlBulkCopyColumnMapping mapping2 = new SqlBulkCopyColumnMapping("[MONTH]", "[MONTH]");
                            SqlBulkCopyColumnMapping mapping3 = new SqlBulkCopyColumnMapping("[LCODE]", "[LCODE]");
                            SqlBulkCopyColumnMapping mapping4 = new SqlBulkCopyColumnMapping("[CHAIN_CODE]", "[CHAIN_CODE]");
                            SqlBulkCopyColumnMapping mapping5 = new SqlBulkCopyColumnMapping("[CHAIN_NAME]", "[CHAIN_NAME]");
                            SqlBulkCopyColumnMapping mapping6 = new SqlBulkCopyColumnMapping("[OFFICEID]", "[OFFICEID]");
                            SqlBulkCopyColumnMapping mapping7 = new SqlBulkCopyColumnMapping("[COUNTRY]", "[COUNTRY]");
                            SqlBulkCopyColumnMapping mapping8 = new SqlBulkCopyColumnMapping("[CODD]", "[CODD]");
                            SqlBulkCopyColumnMapping mapping9 = new SqlBulkCopyColumnMapping("[UPLOAD_TYPE]", "[UPLOAD_TYPE]");


                            objSqlTransaction = objConLivedatabase.BeginTransaction(IsolationLevel.RepeatableRead);

                            objSqlbulkCopy = new SqlBulkCopy(objConLivedatabase, SqlBulkCopyOptions.CheckConstraints, objSqlTransaction);
                            objSqlbulkCopy.SqlRowsCopied += new SqlRowsCopiedEventHandler(objSqlbulkCopy_SqlRowsCopied);

                            objSqlbulkCopy.ColumnMappings.Add(mapping1);
                            objSqlbulkCopy.ColumnMappings.Add(mapping2);
                            objSqlbulkCopy.ColumnMappings.Add(mapping3);
                            objSqlbulkCopy.ColumnMappings.Add(mapping4);
                            objSqlbulkCopy.ColumnMappings.Add(mapping5);
                            objSqlbulkCopy.ColumnMappings.Add(mapping6);
                            objSqlbulkCopy.ColumnMappings.Add(mapping7);
                            objSqlbulkCopy.ColumnMappings.Add(mapping8);
                            objSqlbulkCopy.ColumnMappings.Add(mapping9);


                            objSqlbulkCopy.DestinationTableName = "T_INC_NIDT_PRODUCTIVITY_EXPORTED";
                            objSqlbulkCopy.BatchSize = 1000;
                            objSqlbulkCopy.NotifyAfter = 5;

                            //objSqlbulkCopy.WriteToServer(objBulkDT);
                            DataTableReader reader = objBulkDT.CreateDataReader();

                            using (ValidatingDataReader validator = new ValidatingDataReader(reader, objConLivedatabase, objSqlbulkCopy, objSqlTransaction))
                            {
                                objSqlbulkCopy.WriteToServer(validator);
                            }

                            objSqlTransaction.Commit();
                            objSqlTransaction = null;
                            MessageBox.Show("CODD data for the period of " + intMonth.ToString() + "/" + intYear.ToString() + "  successfully transfered to Live Server database ", "AAMS Admin", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            //END INSERT CODD DATA
                        }
                    }
                    else if (dlgResult == DialogResult.No)
                    {
                        Cursor.Current = Cursors.Default;
                        return false;
                    }
                    Cursor.Current = Cursors.Default;
                    return true;
                }
            }
            catch (Exception e1)
            {
                MessageBox.Show("Problem in BulkInserting." + "\n\n" + "Contact to Admin and send them error message : " + e1.Message + " \n\n" + e1.StackTrace, "AAMS Admin", MessageBoxButtons.OK, MessageBoxIcon.Stop);

                if (objConLivedatabase.State != ConnectionState.Closed)
                {
                    if (objSqlTransaction != null)
                    {
                        objSqlTransaction.Rollback();
                        MessageBox.Show(e1.Message);
                    }
                    objConLivedatabase.Close();
                }
                return false;
            }
            finally
            {
                //myTable1 = null;
                //myTable2 = null;
                //myTable3 = null;
                //myTable4 = null;
                //myTable4_1 = null;
                //gblGrpds = null;

                if (objConLivedatabase.State == ConnectionState.Open)
                {
                    objConLivedatabase.Close();
                }
            }
            return true;
        }

        void objSqlbulkCopy_SqlRowsCopied(object sender, SqlRowsCopiedEventArgs e)
        {
            try
            {
                if (toolStripProgressBar1.Value + (int)e.RowsCopied < toolStripProgressBar1.Maximum)
                {
                    toolStripProgressBar1.Value = toolStripProgressBar1.Value + (int)e.RowsCopied;
                }
                else
                {
                    toolStripProgressBar1.Value = toolStripProgressBar1.Maximum;
                }
            }
            catch (Exception exec1)
            {
                MessageBox.Show("error occured while increasing progressbar : " + exec1.Message);
            }
        }
        #endregion

        #region DoDisableSubItems
        private void DoDisableSubItems(ToolStripMenuItem item, Boolean blnSubItem)
        {
            foreach (ToolStripMenuItem subitem in item.DropDownItems)
            {
                subitem.Enabled = blnSubItem;
            }
        }
        #endregion

        #region GetConecitonString to Connect Database
        private void ConnectToDB()
        {
            try
            {
                String strMyStringNIDT = String.Empty;
                String strMyStringAAMS = String.Empty;

                list.Add("---------- NIDT SERVER CONFIGURATION  -----------");
                strMyStringNIDT = ConfigurationManager.ConnectionStrings["QueryShedularConnectionStringNIDT"].ToString();
                strMyStringNIDT = DecryptConnectionString(strMyStringNIDT);
                list.Add("");
                list.Add("---------- AAMS SERVER CONFIGURATION -----------");
                strMyStringAAMS = ConfigurationManager.ConnectionStrings["QueryShedularConnectionStringAAMS"].ToString();
                strMyStringAAMS = DecryptConnectionString(strMyStringAAMS);
                list.Add("");
                list.Add("Please confirm the above configurations. Would you like to continue !");

                if (strMyStringAAMS == "" || strMyStringNIDT == "")
                {
                    MessageBox.Show("Error while decrypt password from encrypted Connection String , please check password in App.config in <<connectionStrings>> setting.", "AAMS Admin", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    this.Close();
                }

                foreach (object o in list)
                {
                    printString += "\n" + o.ToString();
                }

                DialogResult dlgResult = MessageBox.Show(printString, "AAMS Admin", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (dlgResult == DialogResult.No) { this.Close(); }


                objConCODD = new SqlConnection(strMyStringNIDT);
                objConHL = new SqlConnection(strMyStringNIDT);
                objConHX = new SqlConnection(strMyStringNIDT);
                objConLivedatabase = new SqlConnection(strMyStringAAMS);

                //cmd = new SqlCommand("select top 10 location_code,name,address from location_master", objConCODD);
                //SqlDataReader sqlDr = cmd.ExecuteReader(CommandBehavior.CloseConnection);
                //DataTable objTable = new DataTable();
                //objTable.Load(sqlDr);
            }
            catch (Exception exep)
            {
                MessageBox.Show(exep.Message);
            }
        }
        #endregion

        #region DecryptPassword from Connection string String
        private String DecryptConnectionString(String MyConnectionString)
        {
            try
            {
                String strEncryptPassword = String.Empty;
                String strDecryptPassword = String.Empty;

                string[] words = MyConnectionString.Split(';');
                int splitindexposition = 0;
                foreach (string word in words)
                {
                    if (word.Replace(" ", "").ToString().Contains("DataSource="))
                    {
                        list.Add(word);
                    }
                    if (word.Replace(" ", "").ToString().Contains("InitialCatalog="))
                    {
                        list.Add(word);
                    }

                    if (word.Replace(" ", "").ToString().Contains("password="))
                    {
                        int index = word.Replace(" ", "").IndexOf("=");
                        if (index != -1)
                        {
                            index = index + 1;
                            strEncryptPassword = word.Replace(" ", "").ToString().Replace("password=", "");
                            strDecryptPassword = objBarcode.Decrypt(strEncryptPassword, strKey);
                            words[splitindexposition] = "password=" + strDecryptPassword;
                        }
                    }
                    splitindexposition++;
                }

                MyConnectionString = String.Empty;

                if (words.Length > 1)
                {
                    foreach (string s in words)
                    {
                        MyConnectionString = MyConnectionString + s + ";";
                    }
                }
                else
                {
                    throw new Exception("error while decrypt password from encrypted Connection String ");
                }
            }
            catch (Exception exep)
            {
                MyConnectionString = "";
                MessageBox.Show(exep.Message, "AAMS Admin", MessageBoxButtons.OK, MessageBoxIcon.Error);
                this.Close();
            }

            return MyConnectionString;

        }
        #endregion

        #region CheckBox Event
        void chkHL_CheckedChanged(object sender, EventArgs e)
        {
            try
            {
                if (chkHL.Checked == true) 
                { 
                    panelHL.Enabled = true;
                    chkGrpData.Checked = false;
                    panelGrpData.Enabled = false;
                    lblGrp.Visible = false;
                }
                else if (chkHL.Checked == false) { panelHL.Enabled = false; }
                lblHL.Visible = false;
            }
            catch (Exception exep)
            {
                MessageBox.Show(exep.Message);
            }
        }

        void chkHX_CheckedChanged(object sender, EventArgs e)
        {
            try
            {
                if (chkHX.Checked == true) 
                { 
                    panelHX.Enabled = true;
                    chkGrpData.Checked = false;
                    panelGrpData.Enabled = false;
                    lblGrp.Visible = false;
                }
                else if (chkHX.Checked == false) { panelHX.Enabled = false; }
                lblHX.Visible = false;
            }
            catch (Exception exep)
            {
                MessageBox.Show(exep.Message);
            }
        }

        void chkCODD_CheckedChanged(object sender, EventArgs e)
        {
            try
            {
                if (chkCODD.Checked == true) 
                { 
                    panelCODD.Enabled = true;
                    chkGrpData.Checked = false;
                    panelGrpData.Enabled = false;
                    lblGrp.Visible = false;
                }
                else if (chkCODD.Checked == false) { panelCODD.Enabled = false; }
                lblC.Visible = false;
            }
            catch (Exception exep)
            {
                MessageBox.Show(exep.Message);
            }
        }

        void chkGrpData_CheckedChanged(object sender, EventArgs e)
        {
            try
            {
                if (chkGrpData.Checked == true) 
                { 
                    panelGrpData.Enabled = true;

                    chkHL.Checked = false;
                    panelHL.Enabled = false;
                    lblHL.Visible = false;

                    chkCODD.Checked = false;
                    panelCODD.Enabled = false;
                    lblC.Visible = false;

                    chkHX.Checked = false;
                    panelHX.Enabled = false;
                    lblHX.Visible = false;

                }
                else if (chkGrpData.Checked == false) { panelGrpData.Enabled = false; }
                lblGrp.Visible = false;
            }
            catch (Exception exep)
            {
                MessageBox.Show(exep.Message);
            }
        }



        #endregion

        #region disable Panel & Lable
        private void disableEnablepanel()
        {
            panelCODD.Enabled = false;
            panelHL.Enabled = false;
            panelHX.Enabled = false;
            panelGrpData.Enabled = false;
            lblC.Visible = false;
            lblHL.Visible = false;
            lblHX.Visible = false;
            lblGrp.Visible = false;
            btnExport.Enabled = false;
        }
        #endregion

        #region clearControl
        private void clearControls()
        {
            coddstime.Text = "";
            coddetime.Text = "";

            hlstime.Text = "";
            hletime.Text = "";

            hxstime.Text = "";
            hxetime.Text = "";
        }
        #endregion

        #region fill Month & year into dropdown
        private void fillControls()
        {
            try
            {
                //QueryClass objQueryClass = new QueryClass();
                QueryClass.fillYear(drpCyear);
                QueryClass.fillYear(drpHLYear);
                QueryClass.fillYear(drpHXYear);
                QueryClass.fillYear(drpGrpYear);

                QueryClass.fillMonth(drpCMonth);
                QueryClass.fillMonth(drpHLMonth);
                QueryClass.fillMonth(drpHXMonth);
                QueryClass.fillMonth(drpGrpMonth);

                QueryClass.fillCountry(drpCcountry, "");
                QueryClass.fillCountry(drpHXcountry, "");
                QueryClass.fillCountry(drpHLcountry, "IN");
                QueryClass.fillCountry(drpGrpcountry, "IN");

                QueryClass.fillReportColuimn(listBoxDefault);
            }
            catch (Exception exep)
            {
                MessageBox.Show("error while populating dropdown " + exep.Message);
            }
        }
        #endregion

        #region Form Closing Event
        private void frmQueryShedular_FormClosing(object sender, FormClosingEventArgs e)
        {
            try
            {
                if (isExecutingCODD || isExecutingHL || isExecutingHX)
                {
                    MessageBox.Show(this, "Can't close the form until the pending asynchronous command has completed. Please wait...");
                    e.Cancel = true;
                }
            }
            catch (Exception exep)
            {
                MessageBox.Show(exep.Message);
            }
        }
        #endregion

        #region timer1_Tick
        private void timer1_Tick(object sender, EventArgs e)
        {
            lblTimer.Text = String.Format("{0:T}", DateTime.Now);
            if (isExecutingCODD || isExecutingHL || isExecutingHX)
            { }
            else { timer1.Enabled = false; }
        }
        #endregion

        #region trackBar1_Scroll
        private void trackBar1_Scroll(object sender, EventArgs e)
        {
            //this.toolStripStatusLabel1.Text =((System.Windows.Forms.TrackBar)(sender)).Value.ToString();
            if (((System.Windows.Forms.TrackBar)(sender)).Value > 1)
                this.Opacity = ((float)((System.Windows.Forms.TrackBar)(sender)).Value / 10);
        }
        #endregion

        #region btnExport_Click
        private void btnExport_Click(object sender, EventArgs e)
        {
            try
            {
                if (isExecutingCODD || isExecutingHL || isExecutingHX)
                {
                    MessageBox.Show(this, "Can't export until the pending asynchronous command has completed. Please wait...");
                    return;
                }


                //contextMenuStrip1.Enabled=true;
                btnExport.ContextMenuStrip = contextMenuStrip1;
                Cursor.Position = new Point(Cursor.Position.X, Cursor.Position.Y);
                this.contextMenuStrip1.Show(btnExport, btnExport.PointToClient(Cursor.Position));
            }
            catch (Exception exep)
            {
                MessageBox.Show("error while Exporting " + exep.Message);
            }
        }
        #endregion

        #region EnableDisableMenuItems
        private void EnableDisableMenuItems()
        {
            foreach (ToolStripMenuItem item in contextMenuStrip1.Items)
            {
                if (item.HasDropDownItems)
                {
                    if (item.Text.ToUpper().Trim() == "CODD")
                    {
                        item.Enabled = false;
                        DoDisableSubItems(item, false);
                    }
                }
            }

            foreach (ToolStripMenuItem item in contextMenuStrip1.Items)
            {
                if (item.HasDropDownItems)
                {
                    if (item.Text.ToUpper().Trim() == "HX")
                    {
                        item.Enabled = false;
                        DoDisableSubItems(item, false);
                    }
                }
            }

            foreach (ToolStripMenuItem item in contextMenuStrip1.Items)
            {
                if (item.HasDropDownItems)
                {
                    if (item.Text.ToUpper().Trim() == "HL")
                    {
                        item.Enabled = false;
                        DoDisableSubItems(item, false);
                    }
                }
            }
        }
        #endregion

        #region Connection Information Message
        private void btnconninfo_Click(object sender, EventArgs e)
        {
            try
            {
                MessageBox.Show(printString, "AAMS Admin", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception exep)
            {
                MessageBox.Show(exep.Message, "AAMS Admin", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        #endregion

        #region OpenFile dialog to select Group data
        private void btnBrowse_Click(object sender, EventArgs e)
        {
            try
            {
                // Create an instance of the open file dialog box.
                OpenFileDialog openFileDialog1 = new OpenFileDialog();

                // Set filter options and filter index.
                //openFileDialog1.Filter = "xls files (*.xls)|*.xls|csv files (*.csv)|*.csv";
                //openFileDialog1.FilterIndex = 1;

                openFileDialog1.InitialDirectory = Application.StartupPath;
                openFileDialog1.Title = "Browse group data file to upload";

                openFileDialog1.CheckFileExists = true;
                openFileDialog1.CheckPathExists = true;

                openFileDialog1.DefaultExt = "txt";
                openFileDialog1.Filter = "xls files (*.xls)|*.xls|xlsx files (*.xlsx)|*.xlsx";
                openFileDialog1.FilterIndex = 2;
                openFileDialog1.RestoreDirectory = true;

                openFileDialog1.ReadOnlyChecked = true;
                openFileDialog1.ShowReadOnly = true;

                if (openFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    txtFilename.Text = openFileDialog1.FileName;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "AAMS Admin", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        #endregion

        #region Group data  - Upload        
        private void btnUpload_Click(object sender, EventArgs e)
        {
            Boolean objvalidate;            
            objvalidate = ValidateControls();
            if (objvalidate== false)
            {
                this.toolStripProgressBar1.Minimum = 1;
                this.toolStripProgressBar1.Maximum = 5;
                this.toolStripProgressBar1.Value = 1;
                this.toolStripProgressBar1.Value = this.toolStripProgressBar1.Value + 1;
                DisplayStatus("Checking excel file format...");
                this.toolStripProgressBar1.Value = 5;

                ExcelHandler OExcelHandler = new ExcelHandler();            
                DataSet dsGrpUpload = OExcelHandler.GetDataFromExcel(txtFilename.Text);
                strgrpColumnarray = OExcelHandler.strColumnarray;
                

                if (dsGrpUpload != null)
                {
                    for (int i = 0; i <= dsGrpUpload.Tables[0].Rows.Count - 1; i++)
                    {
                        if (dsGrpUpload.Tables[0].Rows[i][strgrpColumnarray[4]] == DBNull.Value) //"Productivity"
                        {
                            dsGrpUpload.Tables[0].Rows[i][strgrpColumnarray[4]] = 0;
                        }

                        if (dsGrpUpload.Tables[0].Rows[i][strgrpColumnarray[5]] == DBNull.Value) //"Dom"
                        {
                            dsGrpUpload.Tables[0].Rows[i][strgrpColumnarray[5]] = 0;
                        }

                        if (dsGrpUpload.Tables[0].Rows[i][strgrpColumnarray[6]] == DBNull.Value) //"Intl"
                        {
                            dsGrpUpload.Tables[0].Rows[i][strgrpColumnarray[6]] = 0;
                        }
                    }

                    gblGrpds = null;
                    dsGrpUpload.AcceptChanges();
                    gblGrpds = dsGrpUpload;

                    SqlConnection sqlcon = new SqlConnection();
                    sqlcon = objConLivedatabase;
                    string strResult;
                    SqlTransaction objTrans = null;

                    ////////////to be polulate from drop down////////////////////////
                    int intMonth = System.Convert.ToInt16(drpGrpMonth.SelectedIndex);
                    int intYear = System.Convert.ToInt16(drpGrpYear.SelectedItem); ;
                    /////////////////end////////////////////////////////////////////

                    string MessageBoxTitle = "Upload Group Data";
                    string MessageBoxContent = "GroupData is available for the selected month [" + intMonth + "] and year [" + intYear + "]." + "\n\n" + " Are you sure want to upload again ? ";

                    IPNetworking objIp = new IPNetworking();
                    objIp.GetIP4Address(out strIpaddress, out strHostname);

                    if (sqlcon.State == ConnectionState.Open)
                    {
                        sqlcon.Close();
                    }

                    sqlcon.Open();
                    objTrans = sqlcon.BeginTransaction();   //Transaction 
                    try
                    {
                        if (gblGrpds != null)
                        {
                            if (gblGrpds.Tables.Count > 0)
                            {
                                //////////////////////////  check between selection and uploaded file ///////////////////////////
                                if (((int)gblGrpds.Tables[0].Rows[0]["Month"] != intMonth) || ((int)gblGrpds.Tables[0].Rows[0]["Year"] != intYear))
                                {
                                    MessageBox.Show("Please select the correct Month and Year.The selected file containing Month," + gblGrpds.Tables[0].Rows[0]["Month"] + " and Year," + gblGrpds.Tables[0].Rows[0]["Year"], MessageBoxTitle, MessageBoxButtons.OK, MessageBoxIcon.Stop);
                                    return;

                                }
                                //CASE STATUS 1 , FILE WAS LOADED FINALLY AND USE WANTS TO RELOAD AGAIN
                                using (SqlCommand objcmd = new SqlCommand("UP_INC_GROUP_DATA_CHECK_UPDATE", sqlcon))
                                {
                                    objcmd.CommandType = CommandType.StoredProcedure;
                                    objcmd.Parameters.AddWithValue("@MONTH", intMonth);
                                    objcmd.Parameters.AddWithValue("@YEAR", intYear);
                                    objcmd.Parameters.AddWithValue("@STATEMENT", "STATEMENT4");
                                    objcmd.Parameters.AddWithValue("@IPADDRESS", strIpaddress);

                                    SqlParameter OprmRESULT = new SqlParameter();
                                    OprmRESULT.ParameterName = "@RESULT";
                                    OprmRESULT.SqlDbType = SqlDbType.VarChar;
                                    OprmRESULT.Size = 11;
                                    OprmRESULT.Direction = ParameterDirection.Output;
                                    objcmd.Parameters.Add(OprmRESULT);
                                    objcmd.Transaction = objTrans;   //Transaction 
                                    objcmd.ExecuteNonQuery();
                                    strResult = (string)OprmRESULT.Value.ToString();
                                    if (strResult == "NO")
                                    {
                                        MessageBox.Show("Group data available in system with status T(rue),pls contact to Admin for uploading this file.", MessageBoxTitle, MessageBoxButtons.OK, MessageBoxIcon.Stop);
                                        return;
                                    }
                                }
                                //END CASE

                                //////////////////////////////////////////////////////////////////////////
                                using (SqlCommand objcmd = new SqlCommand("UP_INC_GROUP_DATA_CHECK_UPDATE", sqlcon))
                                {
                                    objcmd.CommandType = CommandType.StoredProcedure;
                                    objcmd.Parameters.AddWithValue("@MONTH", intMonth);
                                    objcmd.Parameters.AddWithValue("@YEAR", intYear);
                                    objcmd.Parameters.AddWithValue("@STATEMENT", "STATEMENT1");
                                    objcmd.Parameters.AddWithValue("@IPADDRESS", strIpaddress);
                                    SqlParameter OprmRESULT = new SqlParameter();
                                    OprmRESULT.ParameterName = "@RESULT";
                                    OprmRESULT.SqlDbType = SqlDbType.VarChar;
                                    OprmRESULT.Size = 10;
                                    OprmRESULT.Direction = ParameterDirection.Output;
                                    objcmd.Parameters.Add(OprmRESULT);
                                    objcmd.Transaction = objTrans;
                                    objcmd.ExecuteNonQuery();
                                    strResult = (string)OprmRESULT.Value.ToString();
                                }


                                if (strResult == "EXIST")
                                {
                                    DialogResult dialogResult = MessageBox.Show(MessageBoxContent, MessageBoxTitle, MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                                    if (dialogResult == DialogResult.Yes)
                                    {
                                        using (SqlCommand objcmd = new SqlCommand("UP_INC_GROUP_DATA_CHECK_UPDATE", sqlcon))
                                        {
                                            objcmd.CommandType = CommandType.StoredProcedure;
                                            objcmd.Parameters.AddWithValue("@MONTH", intMonth);
                                            objcmd.Parameters.AddWithValue("@YEAR", intYear);
                                            objcmd.Parameters.AddWithValue("@STATEMENT", "STATEMENT2");
                                            objcmd.Parameters.AddWithValue("@IPADDRESS", strIpaddress);
                                            SqlParameter OprmRESULT = new SqlParameter();
                                            OprmRESULT.ParameterName = "@RESULT";
                                            OprmRESULT.SqlDbType = SqlDbType.VarChar;
                                            OprmRESULT.Size = 10;
                                            OprmRESULT.Direction = ParameterDirection.Output;
                                            objcmd.Parameters.Add(OprmRESULT);
                                            objcmd.Transaction = objTrans;
                                            objcmd.ExecuteNonQuery();
                                            strResult = (string)OprmRESULT.Value.ToString();
                                        }
                                    }
                                    else
                                    {
                                        return;
                                    }
                                }
                                else if (strResult == "NOTEXIST")
                                {
                                    using (SqlCommand objcmd = new SqlCommand("UP_INC_GROUP_DATA_CHECK_UPDATE", sqlcon))
                                    {
                                        objcmd.CommandType = CommandType.StoredProcedure;
                                        objcmd.Parameters.AddWithValue("@MONTH", intMonth);
                                        objcmd.Parameters.AddWithValue("@YEAR", intYear);
                                        objcmd.Parameters.AddWithValue("@STATEMENT", "STATEMENT2");
                                        objcmd.Parameters.AddWithValue("@IPADDRESS", strIpaddress);
                                        SqlParameter OprmRESULT = new SqlParameter();
                                        OprmRESULT.ParameterName = "@RESULT";
                                        OprmRESULT.SqlDbType = SqlDbType.VarChar;
                                        OprmRESULT.Size = 10;
                                        OprmRESULT.Direction = ParameterDirection.Output;
                                        objcmd.Parameters.Add(OprmRESULT);
                                        objcmd.Transaction = objTrans;
                                        objcmd.ExecuteNonQuery();
                                        strResult = (string)OprmRESULT.Value.ToString();
                                    }
                                }
                            }

                            using (SqlCommand objcmdDelete = new SqlCommand("DELETE FROM T_INC_NIDT_PRODUCTIVITY_GROUPDATA_DETAILS OUTPUT DELETED.[Month],DELETED.[Year],DELETED.[AirLineCode],DELETED.[Officeid],DELETED.[Productivity],DELETED.[Dom],DELETED.[Intl],DELETED.[FileName],DELETED.[FileUploadDTTI],DELETED.[IPADDRESS],'" + strIpaddress + "' INTO T_INC_NIDT_PRODUCTIVITY_GROUPDATA_DETAILS_LOG([Month],[Year],[AirLineCode],[Officeid],[Productivity],[Dom],[Intl],[FileName],[FileUploadDTTI],[IPADDRESS],[IPADDRESS_DEL]) where month=" + intMonth + " and year=" + intYear, sqlcon))
                            {
                                objcmdDelete.Transaction = objTrans;
                                int intDeletedRows = objcmdDelete.ExecuteNonQuery();
                                //MessageBox.Show(intDeletedRows + " : Rows Deleted");
                            }

                            using (SqlCommand objcmdDelete = new SqlCommand("delete from T_INC_NIDT_PRODUCTIVITY_GROUPDATA_DETAILS where month=" + intMonth + " and year=" + intYear, sqlcon))
                            {
                                objcmdDelete.Transaction = objTrans;
                                int intDeletedRows = objcmdDelete.ExecuteNonQuery();
                                //MessageBox.Show(intDeletedRows + " : Rows Deleted");
                            }

                            using (SqlCommand objcmd = new SqlCommand("UP_INC_GROUP_DATA_CHECK_UPDATE", sqlcon))
                            {
                                objcmd.CommandType = CommandType.StoredProcedure;
                                objcmd.Parameters.AddWithValue("@MONTH", intMonth);
                                objcmd.Parameters.AddWithValue("@YEAR", intYear);
                                objcmd.Parameters.AddWithValue("@STATEMENT", "STATEMENT3");
                                objcmd.Parameters.AddWithValue("@IPADDRESS", strIpaddress);

                                SqlParameter OprmRESULT = new SqlParameter();

                                OprmRESULT.ParameterName = "@RESULT";
                                OprmRESULT.SqlDbType = SqlDbType.VarChar;
                                OprmRESULT.Size = 10;
                                OprmRESULT.Direction = ParameterDirection.Output;
                                objcmd.Parameters.Add(OprmRESULT);
                                objcmd.Transaction = objTrans;
                                objcmd.ExecuteNonQuery();
                                strResult = (string)OprmRESULT.Value.ToString();
                            }

                            this.toolStripProgressBar1.Minimum = 1;
                            this.toolStripProgressBar1.Maximum = 5;
                            this.toolStripProgressBar1.Value = 1;
                            this.toolStripProgressBar1.Value = this.toolStripProgressBar1.Value + 1;


                            //////////////////////bulkcopy dataset to sql table/////////////////////////////
                            using (SqlBulkCopy sqlBulkCopy = new SqlBulkCopy(sqlcon, SqlBulkCopyOptions.Default, objTrans))
                            {
                                sqlBulkCopy.DestinationTableName = "T_INC_NIDT_PRODUCTIVITY_GROUPDATA_DETAILS";

                                /////////////////////////////adding extra column in datatable//////////////
                                gblGrpds.Tables[0].Columns.Add("FileName", typeof(string));
                                gblGrpds.Tables[0].Columns.Add("FileUploadDTTI", typeof(DateTime));
                                gblGrpds.Tables[0].Columns.Add("IPADDRESS", typeof(string));
                                ////////////////////////////end///////////////////////////////////////////

                                ///////////////////////////////coloumn mapping////////////////////////////
                                sqlBulkCopy.ColumnMappings.Add(strgrpColumnarray[0], "Month");
                                sqlBulkCopy.ColumnMappings.Add(strgrpColumnarray[1], "Year");
                                sqlBulkCopy.ColumnMappings.Add(strgrpColumnarray[2], "AirLineCode");
                                sqlBulkCopy.ColumnMappings.Add(strgrpColumnarray[3], "Officeid");
                                sqlBulkCopy.ColumnMappings.Add(strgrpColumnarray[4], "Productivity");
                                sqlBulkCopy.ColumnMappings.Add(strgrpColumnarray[5], "Dom");
                                sqlBulkCopy.ColumnMappings.Add(strgrpColumnarray[6], "Intl");

                                sqlBulkCopy.ColumnMappings.Add("FileName", "FileName");
                                sqlBulkCopy.ColumnMappings.Add("FileUploadDTTI", "FileUploadDTTI");
                                sqlBulkCopy.ColumnMappings.Add("IPADDRESS", "IPADDRESS");

                                for (int i = 0; i <= gblGrpds.Tables[0].Rows.Count - 1; i++)
                                {
                                    gblGrpds.Tables[0].Rows[i]["FileName"] = txtFilename.Text;
                                    gblGrpds.Tables[0].Rows[i]["FileUploadDTTI"] = DateTime.Now;
                                    gblGrpds.Tables[0].Rows[i]["IPADDRESS"] = strIpaddress;
                                }
                                gblGrpds.AcceptChanges();
                                sqlBulkCopy.WriteToServer(gblGrpds.Tables[0]);
                            }
                            objTrans.Commit();
                            this.toolStripProgressBar1.Value = 5;
                            DisplayStatus(gblGrpds.Tables[0].Rows.Count + ":Rows Added for Group data productivity..");
                            MessageBox.Show(gblGrpds.Tables[0].Rows.Count + " : Rows Added Sucessfully..", MessageBoxTitle);
                            gblGrpds = null;
                        }
                        else
                        {
                            MessageBox.Show("Dataset Cleaned.", MessageBoxTitle);
                            return;
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Problem in uploading.Try again may be the column is not matched with specified format." + "\n\n" + "Contact to Admin and send them error message : " + ex.Message + " \n\n" + ex.StackTrace, MessageBoxTitle, MessageBoxButtons.OK, MessageBoxIcon.Stop);
                        objTrans.Rollback();
                    }
                    finally
                    {
                        pbarHL.Maximum = 100;
                        pbarHL.Minimum = 1;
                        pbarHL.Value = 1;

                        if (sqlcon.State == ConnectionState.Open)
                        {
                            sqlcon.Close();
                        }
                    }
                }
                else 
                {
                    DisplayStatus("Error in excel file format...");
                }
        }
        #endregion
    }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                this.Cursor = Cursors.WaitCursor;
                frmReportSetting objfrmInput = new frmReportSetting();
                objfrmInput.WindowState = FormWindowState.Normal;
                objfrmInput.ShowDialog();                
            }
            catch (Exception ex)
            {
                MessageBox.Show("Problem in loading setting form : " + ex.Message + " \n\n" + ex.StackTrace, "AAMS Admin", MessageBoxButtons.OK, MessageBoxIcon.Stop);                
            }
            finally
            {
                this.Cursor = Cursors.Arrow;
            }
        }
    }

    #region Group data - Read from Excel
    public class ExcelHandler
    {
        String printErrorString;
        System.Data.OleDb.OleDbConnection oledbcn;
        public string[] strColumnarray = new string[7];

        // Return data in dataset from excel file. '   
        public DataSet GetDataFromExcel(string a_sFilepath)
        {         
            
            DataSet ds_excel = new DataSet();
            string extension;
            extension = Path.GetExtension(a_sFilepath);

            if (extension == ".xls")
            {
                //Provider=Microsoft.ACE.OLEDB.4.0 for .xls
                oledbcn = new System.Data.OleDb.OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + a_sFilepath + ";Extended Properties= Excel 8.0");
            }
            else if (extension== ".xlsx")
            {
                //Provider=Microsoft.ACE.OLEDB.12.0 for .xlsx
                oledbcn = new System.Data.OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + a_sFilepath + ";Extended Properties= Excel 8.0");
            }

            try
            {
                oledbcn.Open();
            }
            catch (OleDbException ex)
            {
                //Console.WriteLine(ex.Message);
                oledbcn.Close();
            }
            catch (Exception ex)
            {
                //Console.WriteLine(ex.Message);
                oledbcn.Close();
            }

            // It Represents Excel data table Schema.'
            System.Data.DataTable dt = new System.Data.DataTable();
            dt = oledbcn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
            if (dt != null || dt.Rows.Count > 0)
            {
                for (int sheet_count = 0; sheet_count <= 0; sheet_count++)
                {
                    try
                    {
                        // Create Query to get Data from sheet. '
                        string sheetname = dt.Rows[sheet_count]["table_name"].ToString();

                        DataTable dtColumns = oledbcn.GetSchema("Columns", new string[] { null, null, sheetname, null });
                        List<string> columns = new List<string>();

                        foreach (DataRow dr in dtColumns.Rows)
                        {
                          if (dr[6].ToString()=="1") // month
                          {
                              if (dr[3].ToString().ToUpper().Trim() != "MONTH")
                              {
                                  columns.Add("Month");
                              }
                              if (dr[3].ToString().ToUpper().Trim() == "MONTH")
                              {
                                  strColumnarray[0] = dr[3].ToString();
                              }                              
                          }
                          else if (dr[6].ToString() == "2") // Year
                          {
                              if (dr[3].ToString().ToUpper().Trim() != "YEAR")
                              {
                                  columns.Add("Year");
                              }
                              if (dr[3].ToString().ToUpper().Trim() == "YEAR")
                              {
                                  strColumnarray[1] = dr[3].ToString();
                              }      
                          }
                          else if (dr[6].ToString() == "3") // AirLineCode
                          {
                              if (dr[3].ToString().ToUpper().Trim() != "AIRLINECODE")
                              {
                                  columns.Add("AirLineCode");
                              }
                              if (dr[3].ToString().ToUpper().Trim() == "AIRLINECODE")
                              {
                                  strColumnarray[2] = dr[3].ToString();
                              }      
                          }
                          else if (dr[6].ToString() == "4") // Officeid
                          {
                              if (dr[3].ToString().ToUpper().Trim() != "OFFICEID")
                              {
                                  columns.Add("Officeid");
                              }
                              if (dr[3].ToString().ToUpper().Trim() == "OFFICEID")
                              {
                                  strColumnarray[3] = dr[3].ToString();
                              }      
                          }
                          else if (dr[6].ToString() == "5") // Productivity
                          {
                              if (dr[3].ToString().ToUpper().Trim() != "PRODUCTIVITY")
                              {
                                  columns.Add("Productivity");
                              }
                              if (dr[3].ToString().ToUpper().Trim() == "PRODUCTIVITY")
                              {
                                  strColumnarray[4] = dr[3].ToString();
                              }      
                          }
                          else if (dr[6].ToString() == "6") // Dom
                          {
                              if (dr[3].ToString().ToUpper().Trim() != "DOM")
                              {
                                  columns.Add("Dom");
                              }
                              if (dr[3].ToString().ToUpper().Trim() == "DOM")
                              {
                                  strColumnarray[5] = dr[3].ToString();
                              }      
                          }
                          else if (dr[6].ToString() == "7") // Intl
                          {
                              if (dr[3].ToString().ToUpper().Trim() != "INTL")
                              {
                                  columns.Add("Intl");
                              }
                              if (dr[3].ToString().ToUpper().Trim() == "INTL")
                              {
                                  strColumnarray[6] = dr[3].ToString();
                              }      
                          }   
                        }
                        if (columns.Count > 1)
                        {
                            columns.Add("\n\nThe excel file should be in following format\n\nMonth||Year||AirLineCode||Officeid||Productivity||Dom||Intl");
                            foreach (object o in columns)
                            {
                                printErrorString += "\n" + o.ToString();
                            }

                            DialogResult dlgResult = MessageBox.Show("The following Columns are not in corect position\n" + printErrorString, "AAMS Admin", MessageBoxButtons.OK, MessageBoxIcon.Stop);

                            ds_excel = null;
                        }
                        else 
                        {
                            OleDbDataAdapter da = new OleDbDataAdapter("SELECT * FROM [" + sheetname + "]", oledbcn);
                            System.Data.DataTable dt2 = new System.Data.DataTable();
                            dt2.Columns.Add(strColumnarray[0], System.Type.GetType("System.Int32"));  //Month
                            dt2.Columns.Add(strColumnarray[1], System.Type.GetType("System.Int32"));  //Year
                            dt2.Columns.Add(strColumnarray[2], System.Type.GetType("System.String")); //AirLineCode
                            dt2.Columns.Add(strColumnarray[3], System.Type.GetType("System.String")); //Officeid
                            dt2.Columns.Add(strColumnarray[4], System.Type.GetType("System.Int32"));  //Productivity
                            dt2.Columns.Add(strColumnarray[5], System.Type.GetType("System.Int32"));  //Dom 
                            dt2.Columns.Add(strColumnarray[6], System.Type.GetType("System.Int32"));  //Intl
                            dt2.TableName = sheetname;
                            ds_excel.Tables.Add(dt2);
                            da.Fill(ds_excel.Tables[0]);
                        }                        
                    }
                    catch (DataException ex)
                    {
                        MessageBox.Show(ex.Message, "AAMS Admin", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message, "AAMS Admin", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
            oledbcn.Close();
            return ds_excel;
        }
    }
    #endregion
}
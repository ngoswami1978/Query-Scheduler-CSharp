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

namespace QueryShedular_HL_HX_CODD
{    
    public partial class frmQueryShedular : Form
    {

        #region Variable Declaration        
            private String strFilePath = System.Configuration.ConfigurationSettings.AppSettings["OutputFolder"];
            private bool isExecutingCODD;
            private bool isExecutingHL;
            private bool isExecutingHX;
            private String strMonthyearCODD= String.Empty;
            private String strMonthyearHX= String.Empty;
            private String strMonthyearHL= String.Empty;

            private Stopwatch myCallBackWatchCODD = new Stopwatch();
            private Stopwatch myCallBackWatchHX = new Stopwatch();
            private Stopwatch myCallBackWatchHL = new Stopwatch();

            SqlConnection objConCODD;
            SqlConnection objConHL;
            SqlConnection objConHX;
            SqlConnection objConLivedatabase;

            String strDisplay = String.Empty;            
            String strStoredProcName=String.Empty;

            public DataTable myTable1 ;
            public DataTable myTable2 ;
            public DataTable myTable3 ;
            public Boolean boolExport;
            public int intYear;
            public int intMonth;
            public String  StrCountry;
            public String  StrCountryCode;


            public SqlTransaction objSqlTransaction ;
            public SqlBulkCopy objSqlbulkCopy;

            public SqlCommand cmd;
            public SqlCommand cmd1;

            public Boolean boolBulkInsert ;

        #endregion

        #region Declaration of delegate for Timeinfo and DataResult
            private delegate void displayDataCODD(DataTable exportDTCODD);
            private delegate void displayTimeInfoDelegateCODD(String Text);
            //private delegate void ExportExcel(DataTable objDT , String strMonth  , String strYear, String QueryType);

            public delegate void displayDataHX(DataTable exportDTHX);
            private delegate void displayTimeInfoDelegateHX(String Text);

            private delegate void displayDataHL(DataTable exportDTHL);
            private delegate void displayTimeInfoDelegateHL(String Text);
            

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
                chkCODD.CheckedChanged +=new EventHandler(chkCODD_CheckedChanged);
                chkHX.CheckedChanged +=new EventHandler(chkHX_CheckedChanged);
                chkHL.CheckedChanged +=new EventHandler(chkHL_CheckedChanged);
                lblCbar.Visible = false;
                lblHLbar.Visible = false;
                lblHXbar.Visible = false;
                pbarCODD.Visible = false;
                pbarHL.Visible = false;
                pbarHX.Visible = false;
                grpStatusbar.Visible = false;                

                foreach (ToolStripMenuItem item in contextMenuStrip1.Items)
                {
                    if (item.HasDropDownItems )
                    {
                        DoSubItems(item);
                    }
                }
                this.ContextMenuStrip = this.contextMenuStrip1;


                foreach (Control ctrl in this.Controls)
                {
                    ctrl.MouseDown +=new MouseEventHandler(ctrl_MouseDown);

                    if (ctrl.GetType() == typeof(GroupBox)) 
                    {
                        foreach (Control ctrl1 in ((GroupBox)ctrl).Controls )
                        {
                            ctrl1.MouseDown +=new MouseEventHandler(ctrl_MouseDown);
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
            foreach (ToolStripMenuItem  subitem in item.DropDownItems)
            {
                subitem.Click +=new EventHandler(item_Click);
            }
        }

        void item_Click(object sender, EventArgs e)
        {
            String strFormat;
            try
            {
                ToolStripMenuItem clickedMenu = sender as ToolStripMenuItem;                
                if (clickedMenu.OwnerItem.Text.ToUpper().Trim() =="CODD") 
                {
                    strFormat= clickedMenu.Text;
                    //MessageBox.Show("CODD");
                    ExportToExcel_CSV(myTable1,"CODD",strFormat);                    
                }                
                if (clickedMenu.OwnerItem.Text.ToUpper().Trim() =="HX") 
                {
                    //MessageBox.Show("HX");
                    strFormat= clickedMenu.Text;
                    //MessageBox.Show("CODD");
                    ExportToExcel_CSV(myTable2,"HX",strFormat);
                }                
                if (clickedMenu.OwnerItem.Text.ToUpper().Trim() =="HL") 
                {
                    //MessageBox.Show("HL");
                    strFormat= clickedMenu.Text;
                    //MessageBox.Show("CODD");
                    ExportToExcel_CSV(myTable3,"HL",strFormat);
                }                
            }
            catch (Exception exep)
            {                
                MessageBox.Show(exep.Message);
            }
        }
        #endregion

        #region ExportToEXCEL_CSV
        private void ExportToExcel_CSV(DataTable ObjDT, String strType,String strFormat)
        {            
            ExportToExcel_WIN_App objExport = new ExportToExcel_WIN_App();
            

            if(strType=="CODD")
            {                
                //this.dataGridView1.DataSource = ObjDT;                
                try
                {
                    if (strFormat.ToUpper().Trim() =="XLS")
                    {
                        this.toolStripProgressBar1.Minimum=1;
                        this.toolStripProgressBar1.Maximum=5;
                        this.toolStripProgressBar1.Value=1;
                        this.toolStripProgressBar1.Value=this.toolStripProgressBar1.Value+1;
                        DisplayStatus("Exporting CODD in xls..");
                        
                        boolExport = objExport.ExportToExcel(ObjDT, strMonthyearCODD, "CODD", drpCcountry.SelectedValue.ToString(), strFilePath, toolStripStatusLabel1);
                        this.toolStripProgressBar1.Value=5;
                        if (boolExport == true)
                        {
                            DisplayStatus("Successfully exported CODD.."  + strFilePath);
                        }
                        else
                        {
                            toolStripProgressBar1.Value = 1;
                            pbarCODD.Value = 1;
                        }
                    }
                    else if (strFormat.ToUpper().Trim() =="CSV")
                    {                        
                        this.toolStripProgressBar1.Minimum=1;
                        this.toolStripProgressBar1.Maximum=5;
                        this.toolStripProgressBar1.Value=1;
                        this.toolStripProgressBar1.Value=this.toolStripProgressBar1.Value+1;
                        DisplayStatus("Exporting CODD in csv..");

                        boolExport= objExport.Exportcsv(ObjDT,"CODD",drpCcountry.SelectedValue.ToString(),strMonthyearCODD,strFilePath,toolStripStatusLabel1);
                        this.toolStripProgressBar1.Value=5;
                        if (boolExport == true)
                        {
                            DisplayStatus("Successfully exported CODD.."  + strFilePath);
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

            if (strType=="HX")
            {
                //Boolean boolExport;
                //this.dataGridView2.DataSource = ObjDT;                
                try
                {
                    if (strFormat.ToUpper().Trim() =="XLS")
                    {
                        this.toolStripProgressBar1.Minimum=1;
                        this.toolStripProgressBar1.Maximum=5;
                        this.toolStripProgressBar1.Value=1;
                        this.toolStripProgressBar1.Value=this.toolStripProgressBar1.Value+1;
                        DisplayStatus("Exporting HX in xls..");
                        
                        //ExportToExcel_WIN_App objExport = new ExportToExcel_WIN_App();
                        boolExport = objExport.ExportToExcel(ObjDT, strMonthyearHX, "HX", drpHXcountry.SelectedValue.ToString(), strFilePath, toolStripStatusLabel1);
                        this.toolStripProgressBar1.Value=5;
                        if (boolExport == true)
                        {
                            DisplayStatus("Successfully exported HX.."  + strFilePath);
                        }
                        else
                        {
                            toolStripProgressBar1.Value = 1;
                            pbarHX.Value = 1;
                        }
                    }
                    else if (strFormat.ToUpper().Trim() =="CSV")
                    {
                        this.toolStripProgressBar1.Minimum=1;
                        this.toolStripProgressBar1.Maximum=5;
                        this.toolStripProgressBar1.Value=1;
                        this.toolStripProgressBar1.Value=this.toolStripProgressBar1.Value+1;
                        DisplayStatus("Exporting HX in csv..");
                        objExport.Exportcsv(ObjDT,"HX",drpHXcountry.SelectedValue.ToString(),strMonthyearHX,strFilePath,toolStripStatusLabel1);
                        this.toolStripProgressBar1.Value=5;
                        if (boolExport == true)
                        {
	                        DisplayStatus("Successfully exported HX.."  + strFilePath);
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

            if (strType=="HL")
            {
                //Boolean boolExport;                
                //this.dataGridView3.DataSource = ObjDT;                
                try
                {
                    if (strFormat.ToUpper().Trim() =="XLS")
                    {
                        this.toolStripProgressBar1.Minimum=1;
                        this.toolStripProgressBar1.Maximum=5;
                        this.toolStripProgressBar1.Value=1;
                        this.toolStripProgressBar1.Value=this.toolStripProgressBar1.Value+1;      
                        DisplayStatus("Exporting HL in XLS..");

                        //ExportToExcel_WIN_App objExport = new ExportToExcel_WIN_App();
                        boolExport = objExport.ExportToExcel(ObjDT, strMonthyearHL, "HL", drpHLcountry.SelectedValue.ToString(), strFilePath, toolStripStatusLabel1);
                        this.toolStripProgressBar1.Value=5;
                        if (boolExport == true)
                        {
                            DisplayStatus("Successfully exported HL.."  + strFilePath);
                        }
                        else
                        {
                            toolStripProgressBar1.Value = 1;
                            pbarHL.Value = 1;
                        }
                    }
                    else if (strFormat.ToUpper().Trim() =="CSV")
                    {
                        this.toolStripProgressBar1.Minimum=1;
                        this.toolStripProgressBar1.Maximum=5;
                        this.toolStripProgressBar1.Value=1;
                        this.toolStripProgressBar1.Value=this.toolStripProgressBar1.Value+1;
                        DisplayStatus("Exporting HL in csv..");
                        objExport.Exportcsv(ObjDT,"HL",drpHLcountry.SelectedValue.ToString(),strMonthyearHL,strFilePath,toolStripStatusLabel1);
                        this.toolStripProgressBar1.Value=5;
                        if (boolExport == true)
                        {
	                        DisplayStatus("Successfully exported HL.."  + strFilePath);
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

        #region  Configure Excel Export Folder
        private void ConfigureOutPutFolder()
        {
            try
            {
                // Code to Configure Folder in Local System
                DirectoryInfo objDir;
                //objDir = new DirectoryInfo(@"C:\ExportQuery");
                objDir = new DirectoryInfo(strFilePath);
                if (objDir.Exists==false) 
                {
                    //MessageBox.Show(strFilePath);
                    objDir.Create();
                }                
            }
            catch (Exception exe )
            {                
                this.toolStripStatusLabel1.Text= exe.Message;
            }

        }
        #endregion

        #region GetProcdureName
        private String GetProcdureName(String strMonth)
        {
            String StrShortMonth=String.Empty;
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
                    StrShortMonth ="Jul";
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
                if (chkCODD.Checked == true || chkHL.Checked == true || chkHX.Checked== true)
                {
                    if (chkCODD.Checked== true)
                    {
                        if (drpCMonth.SelectedIndex ==0 )
                        {
                            lblC.Visible=true;
                            objCntl= true;
                        }
                        else
                        {
                            if (drpCyear.SelectedIndex==0)
                            {
                                lblC.Visible=true;
                            }
                            else
                            {
                                lblC.Visible=false;
                            }
                            
                        }
                        if (drpCyear.SelectedIndex==0)
                        {
                            lblC.Visible=true;
                            objCntl= true;
                        }
                        else
                        {
                            if (drpCMonth.SelectedIndex ==0 )
                            {
                                lblC.Visible=true;
                            }
                            else
                            {
                                lblC.Visible=false;
                            }                            
                        }
                    }
                  if (chkHL.Checked== true)
                    {
                        if (drpHLMonth.SelectedIndex ==0 )
                        {
                            lblHL.Visible=true;
                            objCntl= true;
                        }
                        else
                        {
                            if (drpHLMonth.SelectedIndex ==0)
                            {
                                lblHL.Visible=true;
                            }
                            else
                            {
                                lblHL.Visible=false;
                            }                            
                        }

                        if (drpHLYear.SelectedIndex==0)
                        {                            
                            lblHL.Visible=true;
                            objCntl= true;
                        }
                      else
                        {
                            if (drpHLMonth.SelectedIndex ==0)
                            {
                                lblHL.Visible=true;
                            }
                            else
                            {
                                lblHL.Visible=false;
                            }                            
                        }
                    }
                  if (chkHX.Checked== true)
                    {
                        if (drpHXMonth.SelectedIndex ==0 )
                        {
                            lblHX.Visible=true;
                            objCntl= true;
                        }
                        else
                        {
                            if (drpHXYear.SelectedIndex==0)
                            {
                                lblHX.Visible=true;
                            }
                            else
                            {
                                lblHX.Visible=false;
                            }                            
                        }
                        if (drpHXYear.SelectedIndex==0)
                        {
                            lblHX.Visible=true;
                            objCntl= true;
                        }
                        else
                        {
                            if (drpHXMonth.SelectedIndex ==0)
                            {
                                lblHX.Visible=true;
                            }else
                            {
                                lblHX.Visible=false;
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
            timer1.Enabled=true;
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
                    SqlCommand SqlCommandHL = null;
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
                        lblC.Visible=false;

                        lblTimeSCODD.Visible=true;
                        coddstime.Visible=true;
                        lblTimeECODD.Visible=true;
                        coddetime.Visible=true;

                        coddstime.Text= String.Format("{0:T}",DateTime.Now);

                        pbarCODD.Maximum = 100;
                        pbarCODD.Minimum = 1;
                        pbarCODD.Value = 1;
                        pbarCODD.Value = pbarCODD.Value + 10;

                        //objConCODD.Open();

                        //start the watch for 
                        myCallBackWatchCODD.Start();                      

                        isExecutingCODD = true;
                        btnExport.Enabled=true;

                        int param_month = System.Convert.ToInt16(drpCMonth.SelectedIndex);
                        int param_year = System.Convert.ToInt16(drpCyear.SelectedItem);
                        String param_country = System.Convert.ToString(drpCcountry.Text);
                        intYear = param_year;
                        intMonth=param_month;
                        StrCountry = param_country;
                        StrCountryCode = System.Convert.ToString(drpCcountry.SelectedValue);

                        //SqlCommandCODD = new SqlCommand("select top 100 * from location_master", objConCODD);
                     
                        SqlCommandCODD = new SqlCommand();
                        SqlCommandCODD.CommandType=CommandType.StoredProcedure;
                        SqlCommandCODD.CommandText = "UP_NIDT_PRODUCTIVITY_CODD";
                        SqlCommandCODD.Connection = objConCODD;

                        SqlCommandCODD.Parameters.Add(new SqlParameter("@MONTH", SqlDbType.Int));
                        SqlCommandCODD.Parameters["@MONTH"].Value = param_month;

                        SqlCommandCODD.Parameters.Add(new SqlParameter("@YEAR", SqlDbType.Int));
                        SqlCommandCODD.Parameters["@YEAR"].Value = param_year;
                        strMonthyearCODD =  drpCMonth.SelectedItem.ToString() + "" +param_year.ToString();

                        SqlCommandCODD.Parameters.Add(new SqlParameter("@COUNTRY", SqlDbType.VarChar,3));
                        if (StrCountryCode=="0")
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

                        lblTimeSCODD.Visible=false;
                        coddstime.Visible=false;
                        lblTimeECODD.Visible=false;
                        coddetime.Visible=false;

                    }
                    if (chkHL.Checked == true)
                    {
                        
                        lblHLbar.Visible = true;
                        pbarHL.Visible = true;
                        lblHL.Visible=false;

                        lblTimeSHL.Visible=true;
                        hlstime.Visible=true;
                        lblTimeEHL.Visible=true;
                        hletime.Visible=true;

                        hlstime.Text= String.Format("{0:T}",DateTime.Now);

                        pbarHL.Maximum = 100;
                        pbarHL.Minimum = 1;
                        pbarHL.Value = 1;
                        pbarHL.Value = pbarHL.Value + 10;

                        isExecutingHL = true;
                        btnExport.Enabled=true;
                        //SqlCommandHL = new SqlCommand("select top 1000  location_code,name,address from location_master", objConHL);
                        //call to make stored Proc name
                        strStoredProcName = GetProcdureName(drpHLMonth.SelectedItem.ToString());

                        int param_month = System.Convert.ToInt16(drpHLMonth.SelectedIndex);
                        int param_year = System.Convert.ToInt16(drpHLYear.SelectedItem);
                        String param_country = System.Convert.ToString(drpHLcountry.Text);
                        intYear = param_year;
                        intMonth=param_month;
                        StrCountry = param_country;
                        StrCountryCode = System.Convert.ToString(drpHLcountry.SelectedValue);

                        //String param_country = System.Convert.ToString(drpHLcountry.SelectedValue);
                        strMonthyearHL =  drpHLMonth.SelectedItem.ToString() + "" +param_year.ToString();
                     
                        SqlCommandHL = new SqlCommand();
                        SqlCommandHL.CommandType=CommandType.StoredProcedure;
                      
                        SqlCommandHL.CommandText = strStoredProcName;

                        
                        SqlCommandHL.Connection = objConHL;

                        SqlCommandHL.Parameters.Add(new SqlParameter("@MONTH", SqlDbType.Int));
                        SqlCommandHL.Parameters["@MONTH"].Value = param_month;

                        SqlCommandHL.Parameters.Add(new SqlParameter("@YEAR", SqlDbType.Int));
                        SqlCommandHL.Parameters["@YEAR"].Value = param_year;

                        //SqlCommandHL.Parameters.Add(new SqlParameter("@COUNTRY", SqlDbType.VarChar,3));
                        //if (param_country=="0")
                        //{
                        //    SqlCommandHL.Parameters["@COUNTRY"].Value = DBNull.Value;
                        //}
                        //else
                        //{
                        //    SqlCommandHL.Parameters["@COUNTRY"].Value = param_country;
                        //}                       

                        SqlCommandHL.Connection.Open();

                        //start the watch for 
                        myCallBackWatchHL.Start();                      
                        AsyncCallback myCallBackHL = new AsyncCallback(HandleCallbackHL);
                        SqlCommandHL.BeginExecuteReader(myCallBackHL, SqlCommandHL);
                    }
                    else 
                    {
                        
                        lblHLbar.Visible = false;
                        pbarHL.Visible = false;

                        lblTimeSHL.Visible=false;
                        hlstime.Visible=false;
                        lblTimeEHL.Visible=false;
                        hletime.Visible=false;

                    }
                    if (chkHX.Checked == true)
                    {
                        lblHXbar.Visible = true;
                        pbarHX.Visible = true;
                        lblHX.Visible=false;

                        lblTimeSHX.Visible=true;
                        hxstime.Visible=true;
                        lblTimeEHX.Visible=true;
                        hxetime.Visible=true;

                        hxstime.Text= String.Format("{0:T}",DateTime.Now);

                        pbarHX.Maximum = 100;
                        pbarHX.Minimum = 1;
                        pbarHX.Value = 1;
                        pbarHX.Value = pbarHX.Value + 10;

                        //objConHX.Open();
                        //start the watch for 
                        myCallBackWatchHX.Start();
                        isExecutingHX = true;
                        btnExport.Enabled=true;
                        //SqlCommandHX = new SqlCommand("select top 1000  location_code,name,address from location_master", objConHX);

                        int param_month = System.Convert.ToInt16(drpHXMonth.SelectedIndex);
                        int param_year = System.Convert.ToInt16(drpHXYear.SelectedItem);
                        String param_country = System.Convert.ToString(drpHXcountry.Text);
                        intYear = param_year;
                        intMonth=param_month;
                        StrCountry = param_country;
                        StrCountryCode = System.Convert.ToString(drpHXcountry.SelectedValue);
                        //SqlCommandCODD = new SqlCommand("select top 100 * from location_master", objConCODD);
                     
                        SqlCommandHX = new SqlCommand();
                        SqlCommandHX.CommandType=CommandType.StoredProcedure;
                        SqlCommandHX.CommandText = "UP_NIDT_PRODUCTIVITY_HX";
                        SqlCommandHX.Connection = objConHX;

                        SqlCommandHX.Parameters.Add(new SqlParameter("@MONTH", SqlDbType.Int));
                        SqlCommandHX.Parameters["@MONTH"].Value = param_month;

                        SqlCommandHX.Parameters.Add(new SqlParameter("@YEAR", SqlDbType.Int));
                        SqlCommandHX.Parameters["@YEAR"].Value = param_year;

                        SqlCommandHX.Parameters.Add(new SqlParameter("@COUNTRY", SqlDbType.VarChar,3));
                        if (StrCountryCode=="0")
                            SqlCommandHX.Parameters["@COUNTRY"].Value = DBNull.Value;
                        else
                            SqlCommandHX.Parameters["@COUNTRY"].Value = StrCountryCode;

                        strMonthyearHX =  drpHXMonth.SelectedItem.ToString() + "" +param_year.ToString();

                        SqlCommandHX.Connection.Open();

                        AsyncCallback myCallBackHX = new AsyncCallback(HandleCallbackHX);
                        SqlCommandHX.BeginExecuteReader(myCallBackHX, SqlCommandHX);
                    }
                    else
                    {
                        lblHXbar.Visible = false;
                        pbarHX.Visible = false;

                        lblTimeSHX.Visible=false;
                        hxstime.Visible=false;
                        lblTimeEHX.Visible=false;
                        hxetime.Visible=false;
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
                this.Invoke(myWatchdisplayCODD,myCallBackTime);

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
                this.Invoke(myWatchdisplayHX,myCallBackTime);

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

                // Stop the watch so we can see how long it took to process
                myCallBackWatchHL.Stop();

                String myCallBackTime = myCallBackWatchHL.ElapsedMilliseconds.ToString();
                displayTimeInfoDelegateHL myWatchdisplayHL = new displayTimeInfoDelegateHL(displayHLTime);
                this.Invoke(myWatchdisplayHL,myCallBackTime);

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
        #endregion

        #region private void displayTime
        private void displayHLTime(String Text)
        {
            try
            {
                hletime.Text= Math.Round(TimeSpan.FromMilliseconds(double.Parse(Text)).TotalMinutes,5).ToString() + " Minutes";
            }
            catch (Exception exe )
            {                
                this.toolStripStatusLabel1.Text ="Error occured while display time : " +  exe.Message;
            }
        }

        private void displayHXTime(String Text)
        {
            try
            {
                hxetime.Text= Math.Round(TimeSpan.FromMilliseconds(double.Parse(Text)).TotalMinutes,5).ToString()+ " Minutes";
            }
            catch (Exception exe)
            {                
                this.toolStripStatusLabel1.Text ="Error occured while display time : " +  exe.Message;
            } 
        }
        private void displayCODDTime(String Text)
        {
            try
            {
                coddetime.Text = Math.Round(TimeSpan.FromMilliseconds(double.Parse(Text)).TotalMinutes,5).ToString()+ " Minutes";
            }
            catch (Exception exe)
            {                
                this.toolStripStatusLabel1.Text ="Error occured while display time : " +  exe.Message;
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
                    if (item.HasDropDownItems )
                    {
                        if (item.Text.ToUpper().Trim() =="CODD")
                        {
                            item.Enabled=true;
                            DoDisableSubItems(item,true);
                        }
                        else
                        {
                            if (item.Enabled==false)
                            {
                                item.Enabled=false;
                                DoDisableSubItems(item,false);
                            }
                        }
                    }
                }
                pbarCODD.Value = pbarCODD.Value + 5;                
                pbarCODD.Value = 100;
                lblRowcountCODD.Text= "Row Count : " + ObjDT.Rows.Count;
                DisplayStatus("CODD Ready...");                
                pbarCODD.Cursor= Cursors.Default;

                /*Export data into database server Newly code Implemented as on date 17-12-2010 [Neeraj Goswami]*/
                DialogResult dlgResult = MessageBox.Show("Do you want to continue to Transfer CODD data into Live server database?", "AAMS Admin", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (dlgResult == DialogResult.Yes)
                {
                    boolBulkInsert = ExportBulkInsert(ObjDT,"CODD");
                    if (boolBulkInsert == true)
                    {
                        MessageBox.Show("CODD data for the period of " + intMonth.ToString() + "/" + intYear.ToString() + "  successfully transfered to Live Server database ","AAMS Admin",MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }                   
                }           

            }
        
            private void DisplayDataResultsHX(DataTable ObjDT)
            {
                foreach (ToolStripMenuItem item in contextMenuStrip1.Items)
                {
                    if (item.HasDropDownItems )
                    {
                        if (item.Text.ToUpper().Trim() =="HX")
                        {
                            item.Enabled=true;
                            DoDisableSubItems(item,true);
                        }
                        else
                        {
                            if (item.Enabled==false)
                            {
                                item.Enabled=false;
                                DoDisableSubItems(item,false);
                            }
                        }
                    }
                }
                pbarHX.Value = pbarHX.Value + 5;
                pbarHX.Value = 100;
                lblRowcountHX.Text= "Row Count : " + ObjDT.Rows.Count;
                DisplayStatus("HX Ready...");
                pbarHX.Cursor= Cursors.Default;

                /*Export data into database server Newly code Implemented as on date 17-12-2010 [Neeraj Goswami]*/
                DialogResult dlgResult = MessageBox.Show("Do you want to continue to Transfer HX data into Live server database?", "AAMS Admin", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (dlgResult == DialogResult.Yes)
                {
                    boolBulkInsert = ExportBulkInsert(ObjDT,"HX");
                    if (boolBulkInsert == true)
                    {
                        MessageBox.Show("HX data for the period of " + intMonth.ToString() + "/" + intYear.ToString() + "  successfully transfered to Live Server database ","AAMS Admin",MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }                   
                }
            }

            private void DisplayDataResultsHL(DataTable ObjDT)
            {
                foreach (ToolStripMenuItem item in contextMenuStrip1.Items)
                {
                    if (item.HasDropDownItems )
                    {
                        if (item.Text.ToUpper().Trim() =="HL")
                        {
                            item.Enabled=true;
                            DoDisableSubItems(item,true);
                        }
                        else
                        {
                            if (item.Enabled==false)
                            {
                                item.Enabled=false;
                                DoDisableSubItems(item,false);
                            }
                        }
                    }
                }
                pbarHL.Value = pbarHL.Value + 5;
                pbarHL.Value = 100;
                DisplayStatus("HL Ready...");
                lblRowcounthl.Text= "Row Count : " + ObjDT.Rows.Count;
                pbarHL.Cursor= Cursors.Default;                

                /*Export data into database server Newly code Implemented as on date 17-12-2010 [Neeraj Goswami]*/
                DialogResult dlgResult = MessageBox.Show("Do you want to continue to Transfer HL data into Live server database?", "AAMS Admin", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (dlgResult == DialogResult.Yes)
                {
                    boolBulkInsert = ExportBulkInsert(ObjDT,"HL");
                    if (boolBulkInsert == true)
                    {
                        MessageBox.Show("HL data for the period of " + intMonth.ToString() + "/" + intYear.ToString() + "  successfully transfered to Live Server database ","AAMS Admin",MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }                   
                }                  
            }
        #endregion


        #region Mapping Column before Bulk Insert
        // Mapping of each Column while Exporting to database server
        private Boolean ExportBulkInsert(DataTable objBulkDT, String strExportType)
        {   
            int intCount;
            

            if (objBulkDT.Rows.Count==0)
            {
                MessageBox.Show("data is not available to Transter for the period of " + intMonth.ToString() + "/" + intYear.ToString() + "  successfully transfered to Live Server database ","AAMS Admin",MessageBoxButtons.OK, MessageBoxIcon.Information);
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
                    cmd = new SqlCommand("SELECT YEAR , MONTH FROM T_INC_NIDT_PRODUCTIVITY_EXPORTED WHERE COUNTRY  = '" + StrCountry +"' AND  MONTH = "  + intMonth + " AND YEAR = " + intYear + " GROUP BY YEAR,MONTH", objConLivedatabase);
                }
                else if (strExportType == "HX")
                {
                    cmd = new SqlCommand("SELECT HX_BOOKINGS = SUM(HX_BOOKINGS)  FROM T_INC_NIDT_PRODUCTIVITY_EXPORTED WHERE COUNTRY  = '" + StrCountry +"' AND  MONTH = "  + intMonth + " AND YEAR = " + intYear + " GROUP BY YEAR,MONTH HAVING ISNULL(SUM(HX_BOOKINGS) ,0) !=0 ", objConLivedatabase);
                }
                else if (strExportType == "CODD")
                {
                    cmd = new SqlCommand("SELECT CODD = SUM(CODD) FROM T_INC_NIDT_PRODUCTIVITY_EXPORTED WHERE COUNTRY  = '" + StrCountry +"' AND  MONTH = "  + intMonth + " AND YEAR = " + intYear + " GROUP BY YEAR,MONTH HAVING ISNULL(SUM(CODD),0) !=0 ", objConLivedatabase);
                }


                // 2. Call Execute reader to get query results
                SqlDataReader rdr = cmd.ExecuteReader();
                if (rdr.HasRows== false )
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

                        colMonth.DefaultValue=intMonth;
                        colYear.DefaultValue=intYear;
                        colType.DefaultValue = "HL";

                        objBulkDT.Columns.Add(colMonth);
                        objBulkDT.Columns.Add(colYear);
                        objBulkDT.Columns.Add(colType); 
                        

                        intCount = objBulkDT.Rows.Count;
                        toolStripProgressBar1.Minimum=0;
                        toolStripProgressBar1.Maximum= intCount;                

                        SqlBulkCopyColumnMapping mapping1 = new SqlBulkCopyColumnMapping("YEAR","YEAR");
                        SqlBulkCopyColumnMapping mapping2 = new SqlBulkCopyColumnMapping("MONTH","MONTH");
                        SqlBulkCopyColumnMapping mapping3 = new SqlBulkCopyColumnMapping("LCODE","LCODE");
                        SqlBulkCopyColumnMapping mapping4 = new SqlBulkCopyColumnMapping("CHAIN_CODE","CHAIN_CODE");
                        SqlBulkCopyColumnMapping mapping5 = new SqlBulkCopyColumnMapping("CHAIN_NAME","CHAIN_NAME");
                        SqlBulkCopyColumnMapping mapping6 = new SqlBulkCopyColumnMapping("OFFICEID","OFFICEID");
                        SqlBulkCopyColumnMapping mapping7 = new SqlBulkCopyColumnMapping("COUNTRY","COUNTRY");
                        SqlBulkCopyColumnMapping mapping8 = new SqlBulkCopyColumnMapping("PRODUCTIVITY_CODD_PK_HX","PRODUCTIVITY_CODD_PK_HX");
                        SqlBulkCopyColumnMapping mapping9 = new SqlBulkCopyColumnMapping("INTL","INTL");
                        SqlBulkCopyColumnMapping mapping10 = new SqlBulkCopyColumnMapping("TTL_HL","TTL_HL");
                        SqlBulkCopyColumnMapping mapping11 = new SqlBulkCopyColumnMapping("HL_INTL","HL_INTL");
                        SqlBulkCopyColumnMapping mapping12 = new SqlBulkCopyColumnMapping("S2_DOM_CODD_PK_HX","S2_DOM_CODD_PK_HX");
                        SqlBulkCopyColumnMapping mapping13 = new SqlBulkCopyColumnMapping("S2_HL_NETSTATUS","S2_HL_NETSTATUS");
                        SqlBulkCopyColumnMapping mapping14 = new SqlBulkCopyColumnMapping("IC_DOM_CODD_PK_HX","IC_DOM_CODD_PK_HX");
                        SqlBulkCopyColumnMapping mapping15 = new SqlBulkCopyColumnMapping("IC_HL_NETSTATUS","IC_HL_NETSTATUS");
                        SqlBulkCopyColumnMapping mapping16 = new SqlBulkCopyColumnMapping("9W_DOM_CODD_PK_HX","9W_DOM_CODD_PK_HX");
                        SqlBulkCopyColumnMapping mapping17 = new SqlBulkCopyColumnMapping("9W_HL_NETSTATUS","9W_HL_NETSTATUS");
                        SqlBulkCopyColumnMapping mapping18 = new SqlBulkCopyColumnMapping("AI_DOM_CODD_PK_HX","AI_DOM_CODD_PK_HX");
                        SqlBulkCopyColumnMapping mapping19 = new SqlBulkCopyColumnMapping("AI_HL_NETSTATUS","AI_HL_NETSTATUS");
                        SqlBulkCopyColumnMapping mapping20 = new SqlBulkCopyColumnMapping("IT_DOM_CODD_PK_HX","IT_DOM_CODD_PK_HX");
                        SqlBulkCopyColumnMapping mapping21 = new SqlBulkCopyColumnMapping("IT_HL_NETSTATUS","IT_HL_NETSTATUS");
                        SqlBulkCopyColumnMapping mapping22 = new SqlBulkCopyColumnMapping("ITRED_DOM_CODD_PK_HX","ITRED_DOM_CODD_PK_HX");
                        SqlBulkCopyColumnMapping mapping23 = new SqlBulkCopyColumnMapping("ITRED_HL_NETSTATUS","ITRED_HL_NETSTATUS");
                        SqlBulkCopyColumnMapping mapping24 = new SqlBulkCopyColumnMapping("I7_DOM_CODD_PK_HX","I7_DOM_CODD_PK_HX");
                        SqlBulkCopyColumnMapping mapping25 = new SqlBulkCopyColumnMapping("I7_HL_NETSTATUS","I7_HL_NETSTATUS");
                        SqlBulkCopyColumnMapping mapping26 = new SqlBulkCopyColumnMapping("TOTALPK","TOTALPK");
                        SqlBulkCopyColumnMapping mapping27 = new SqlBulkCopyColumnMapping("DOM_PK_IC","DOM_PK_IC");
                        SqlBulkCopyColumnMapping mapping28 = new SqlBulkCopyColumnMapping("DOM_PK_IT","DOM_PK_IT");
                        SqlBulkCopyColumnMapping mapping29 = new SqlBulkCopyColumnMapping("DOM_PK_AI","DOM_PK_AI");
                        SqlBulkCopyColumnMapping mapping30 = new SqlBulkCopyColumnMapping("DOM_PK_9W","DOM_PK_9W");
                        SqlBulkCopyColumnMapping mapping31 = new SqlBulkCopyColumnMapping("CODD","CODD");
                        SqlBulkCopyColumnMapping mapping32 = new SqlBulkCopyColumnMapping("ROI","ROI");
                        SqlBulkCopyColumnMapping mapping33 = new SqlBulkCopyColumnMapping("S2_HX","S2_HX");
                        SqlBulkCopyColumnMapping mapping34 = new SqlBulkCopyColumnMapping("IC_HX","IC_HX");
                        SqlBulkCopyColumnMapping mapping35 = new SqlBulkCopyColumnMapping("9W_HX","9W_HX");
                        SqlBulkCopyColumnMapping mapping36 = new SqlBulkCopyColumnMapping("AI_HX","AI_HX");
                        SqlBulkCopyColumnMapping mapping37 = new SqlBulkCopyColumnMapping("IT_HX","IT_HX");
                        SqlBulkCopyColumnMapping mapping38 = new SqlBulkCopyColumnMapping("I7_HX","I7_HX");
                        SqlBulkCopyColumnMapping mapping39 = new SqlBulkCopyColumnMapping("HX_BOOKINGS","HX_BOOKINGS");
                        SqlBulkCopyColumnMapping mapping40 = new SqlBulkCopyColumnMapping("DOM_PK_S2", "DOM_PK_S2");
                        SqlBulkCopyColumnMapping mapping41 = new SqlBulkCopyColumnMapping("UPLOAD_TYPE", "UPLOAD_TYPE");

                        objSqlTransaction = objConLivedatabase.BeginTransaction(IsolationLevel.RepeatableRead);
                    
                        objSqlbulkCopy = new SqlBulkCopy(objConLivedatabase, SqlBulkCopyOptions.CheckConstraints, objSqlTransaction);
                        objSqlbulkCopy.SqlRowsCopied +=new SqlRowsCopiedEventHandler(objSqlbulkCopy_SqlRowsCopied);

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
                        objSqlbulkCopy.ColumnMappings.Add(mapping41);

                        objSqlbulkCopy.DestinationTableName = "T_INC_NIDT_PRODUCTIVITY_EXPORTED";
                        objSqlbulkCopy.BatchSize = 1000;
                        objSqlbulkCopy.NotifyAfter = 5;
                        //objSqlbulkCopy.WriteToServer(objBulkDT);

                        DataTableReader reader = objBulkDT.CreateDataReader();

                        using (ValidatingDataReader validator = new ValidatingDataReader(reader, objConLivedatabase, objSqlbulkCopy,objSqlTransaction))
                        {
	                        objSqlbulkCopy.WriteToServer(validator);
                        }


                        objSqlTransaction.Commit();
                        objSqlTransaction =null;
                        return true;

                      }
                    else if (strExportType == "HX")
                    {
                        DataColumn colMonth = new DataColumn("MONTH", typeof(System.Int16));
                        DataColumn colYear = new DataColumn("YEAR", typeof(System.Int16));
                        DataColumn colType = new DataColumn("UPLOAD_TYPE", typeof(string));    

                        colMonth.DefaultValue=intMonth;
                        colYear.DefaultValue=intYear;
                        colType.DefaultValue = "HX";

                        objBulkDT.Columns.Add(colMonth);
                        objBulkDT.Columns.Add(colYear);
                        objBulkDT.Columns.Add(colType); 

                        intCount = objBulkDT.Rows.Count;

                        toolStripProgressBar1.Minimum=0;
                        toolStripProgressBar1.Maximum= intCount;                

                        SqlBulkCopyColumnMapping mapping1 = new SqlBulkCopyColumnMapping("YEAR","YEAR");
                        SqlBulkCopyColumnMapping mapping2 = new SqlBulkCopyColumnMapping("MONTH","MONTH");
                        SqlBulkCopyColumnMapping mapping3 = new SqlBulkCopyColumnMapping("LCODE","LCODE");
                        SqlBulkCopyColumnMapping mapping4 = new SqlBulkCopyColumnMapping("CHAIN_CODE","CHAIN_CODE");
                        SqlBulkCopyColumnMapping mapping5 = new SqlBulkCopyColumnMapping("CHAIN_NAME","CHAIN_NAME");
                        SqlBulkCopyColumnMapping mapping6 = new SqlBulkCopyColumnMapping("OFFICEID","OFFICEID");
                        SqlBulkCopyColumnMapping mapping7 = new SqlBulkCopyColumnMapping("COUNTRY","COUNTRY");
                        SqlBulkCopyColumnMapping mapping8 = new SqlBulkCopyColumnMapping("HX_BOOKINGS", "HX_BOOKINGS");
                        SqlBulkCopyColumnMapping mapping9 = new SqlBulkCopyColumnMapping("UPLOAD_TYPE", "UPLOAD_TYPE");


                        objSqlTransaction = objConLivedatabase.BeginTransaction(IsolationLevel.RepeatableRead);                    
                        objSqlbulkCopy = new SqlBulkCopy(objConLivedatabase, SqlBulkCopyOptions.CheckConstraints, objSqlTransaction);
                        objSqlbulkCopy.SqlRowsCopied +=new SqlRowsCopiedEventHandler(objSqlbulkCopy_SqlRowsCopied);

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

                        using (ValidatingDataReader validator = new ValidatingDataReader(reader, objConLivedatabase, objSqlbulkCopy,objSqlTransaction))
                        {
	                        objSqlbulkCopy.WriteToServer(validator);
                        }

                        objSqlTransaction.Commit();
                        objSqlTransaction =null;
                        return true;

                    }
                    else if( strExportType=="CODD")
                    {
                        DataColumn colMonth = new DataColumn("MONTH", typeof(System.Int16));
                        DataColumn colYear = new DataColumn("YEAR", typeof(System.Int16));
                        DataColumn colType = new DataColumn("UPLOAD_TYPE", typeof(string));    

                        colMonth.DefaultValue=intMonth;
                        colYear.DefaultValue=intYear;
                        colType.DefaultValue = "CODD";
                        
                        objBulkDT.Columns.Add(colMonth);
                        objBulkDT.Columns.Add(colYear);
                        objBulkDT.Columns.Add(colType); 


                        intCount = objBulkDT.Rows.Count;
                        toolStripProgressBar1.Minimum=0;
                        toolStripProgressBar1.Maximum= intCount;                

                        SqlBulkCopyColumnMapping mapping1 = new SqlBulkCopyColumnMapping("[YEAR]","[YEAR]");
                        SqlBulkCopyColumnMapping mapping2 = new SqlBulkCopyColumnMapping("[MONTH]","[MONTH]");
                        SqlBulkCopyColumnMapping mapping3 = new SqlBulkCopyColumnMapping("[LCODE]","[LCODE]");
                        SqlBulkCopyColumnMapping mapping4 = new SqlBulkCopyColumnMapping("[CHAIN_CODE]","[CHAIN_CODE]");
                        SqlBulkCopyColumnMapping mapping5 = new SqlBulkCopyColumnMapping("[CHAIN_NAME]","[CHAIN_NAME]");
                        SqlBulkCopyColumnMapping mapping6 = new SqlBulkCopyColumnMapping("[OFFICEID]","[OFFICEID]");
                        SqlBulkCopyColumnMapping mapping7 = new SqlBulkCopyColumnMapping("[COUNTRY]","[COUNTRY]");
                        SqlBulkCopyColumnMapping mapping8 = new SqlBulkCopyColumnMapping("[CODD]", "[CODD]");
                        SqlBulkCopyColumnMapping mapping9 = new SqlBulkCopyColumnMapping("[UPLOAD_TYPE]", "[UPLOAD_TYPE]");


                        objSqlTransaction = objConLivedatabase.BeginTransaction(IsolationLevel.RepeatableRead);
                    
                        objSqlbulkCopy = new SqlBulkCopy(objConLivedatabase, SqlBulkCopyOptions.CheckConstraints, objSqlTransaction);
                        objSqlbulkCopy.SqlRowsCopied +=new SqlRowsCopiedEventHandler(objSqlbulkCopy_SqlRowsCopied);

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

                        using (ValidatingDataReader validator = new ValidatingDataReader(reader, objConLivedatabase, objSqlbulkCopy,objSqlTransaction))
                        {
	                        objSqlbulkCopy.WriteToServer(validator);
                        }

                        objSqlTransaction.Commit();
                        objSqlTransaction =null;
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

                        DialogResult dlgResult = MessageBox.Show("NIDT data for the month/year " + intMonth.ToString() + "/" + intYear.ToString() + " already exists in Live Server database. do you want to continue Upload data into Live Server ? ", "AAMS Admin", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);

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
                                cmd1 = new SqlCommand("DELETE FROM T_INC_NIDT_PRODUCTIVITY_EXPORTED OUTPUT  DELETED.[YEAR],DELETED.[MONTH],DELETED.[LCODE],DELETED.[CHAIN_CODE],DELETED.[CHAIN_NAME],DELETED.[OFFICEID],DELETED.[COUNTRY],DELETED.[PRODUCTIVITY_CODD_PK_HX],DELETED.[INTL],DELETED.[TTL_HL],DELETED.[HL_INTL],DELETED.[S2_DOM_CODD_PK_HX],DELETED.[S2_HL_NETSTATUS],DELETED.[IC_DOM_CODD_PK_HX],DELETED.[IC_HL_NETSTATUS],DELETED.[9W_DOM_CODD_PK_HX],DELETED.[9W_HL_NETSTATUS],DELETED.[AI_DOM_CODD_PK_HX],DELETED.[AI_HL_NETSTATUS],DELETED.[IT_DOM_CODD_PK_HX],DELETED.[IT_HL_NETSTATUS],DELETED.[ITRED_DOM_CODD_PK_HX],DELETED.[ITRED_HL_NETSTATUS],DELETED.[I7_DOM_CODD_PK_HX],DELETED.[I7_HL_NETSTATUS],DELETED.[TOTALPK],DELETED.[DOM_PK_IC],DELETED.[DOM_PK_IT],DELETED.[DOM_PK_AI],DELETED.[DOM_PK_9W],DELETED.[CODD],DELETED.[ROI],DELETED.[S2_HX],DELETED.[IC_HX],DELETED.[9W_HX],DELETED.[AI_HX],DELETED.[IT_HX],DELETED.[I7_HX],DELETED.[HX_BOOKINGS],DELETED.[DOM_PK_S2],DELETED.UPLOAD_TYPE  INTO T_INC_NIDT_PRODUCTIVITY_EXPORTED_LOG ([YEAR],[MONTH],[LCODE],[CHAIN_CODE],[CHAIN_NAME],[OFFICEID],[COUNTRY],[PRODUCTIVITY_CODD_PK_HX],[INTL],[TTL_HL],[HL_INTL],[S2_DOM_CODD_PK_HX],[S2_HL_NETSTATUS],[IC_DOM_CODD_PK_HX],[IC_HL_NETSTATUS],[9W_DOM_CODD_PK_HX],[9W_HL_NETSTATUS],[AI_DOM_CODD_PK_HX],[AI_HL_NETSTATUS],[IT_DOM_CODD_PK_HX],[IT_HL_NETSTATUS],[ITRED_DOM_CODD_PK_HX],[ITRED_HL_NETSTATUS],[I7_DOM_CODD_PK_HX],[I7_HL_NETSTATUS],[TOTALPK],[DOM_PK_IC],[DOM_PK_IT],[DOM_PK_AI],[DOM_PK_9W],[CODD],[ROI],[S2_HX],[IC_HX],[9W_HX],[AI_HX],[IT_HX],[I7_HX],[HX_BOOKINGS],[DOM_PK_S2],UPLOAD_TYPE) WHERE UPLOAD_TYPE = 'HL' AND  MONTH = " + intMonth + " AND YEAR = " + intYear, objConLivedatabase);
                                // 2. Call Execute query 
                                int intRowaffected = cmd1.ExecuteNonQuery();

                                //INSERT HL DATA AFTER DELETE QUERY
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
                                SqlBulkCopyColumnMapping mapping41 = new SqlBulkCopyColumnMapping("UPLOAD_TYPE", "UPLOAD_TYPE");

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
                                objSqlbulkCopy.ColumnMappings.Add(mapping41);

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
                                MessageBox.Show("HL data for the period of " + intMonth.ToString() + "/" + intYear.ToString() + "  successfully transfered to Live Server database ", "AAMS Admin", MessageBoxButtons.OK, MessageBoxIcon.Information);
                                //END INSERT HL DATA
                            }
                            else if (strExportType == "HX")
                            {
                                cmd1 = new SqlCommand("DELETE FROM T_INC_NIDT_PRODUCTIVITY_EXPORTED OUTPUT  DELETED.[YEAR],DELETED.[MONTH],DELETED.[LCODE],DELETED.[CHAIN_CODE],DELETED.[CHAIN_NAME],DELETED.[OFFICEID],DELETED.[COUNTRY],DELETED.[PRODUCTIVITY_CODD_PK_HX],DELETED.[INTL],DELETED.[TTL_HL],DELETED.[HL_INTL],DELETED.[S2_DOM_CODD_PK_HX],DELETED.[S2_HL_NETSTATUS],DELETED.[IC_DOM_CODD_PK_HX],DELETED.[IC_HL_NETSTATUS],DELETED.[9W_DOM_CODD_PK_HX],DELETED.[9W_HL_NETSTATUS],DELETED.[AI_DOM_CODD_PK_HX],DELETED.[AI_HL_NETSTATUS],DELETED.[IT_DOM_CODD_PK_HX],DELETED.[IT_HL_NETSTATUS],DELETED.[ITRED_DOM_CODD_PK_HX],DELETED.[ITRED_HL_NETSTATUS],DELETED.[I7_DOM_CODD_PK_HX],DELETED.[I7_HL_NETSTATUS],DELETED.[TOTALPK],DELETED.[DOM_PK_IC],DELETED.[DOM_PK_IT],DELETED.[DOM_PK_AI],DELETED.[DOM_PK_9W],DELETED.[CODD],DELETED.[ROI],DELETED.[S2_HX],DELETED.[IC_HX],DELETED.[9W_HX],DELETED.[AI_HX],DELETED.[IT_HX],DELETED.[I7_HX],DELETED.[HX_BOOKINGS],DELETED.[DOM_PK_S2],DELETED.UPLOAD_TYPE INTO T_INC_NIDT_PRODUCTIVITY_EXPORTED_LOG ([YEAR],[MONTH],[LCODE],[CHAIN_CODE],[CHAIN_NAME],[OFFICEID],[COUNTRY],[PRODUCTIVITY_CODD_PK_HX],[INTL],[TTL_HL],[HL_INTL],[S2_DOM_CODD_PK_HX],[S2_HL_NETSTATUS],[IC_DOM_CODD_PK_HX],[IC_HL_NETSTATUS],[9W_DOM_CODD_PK_HX],[9W_HL_NETSTATUS],[AI_DOM_CODD_PK_HX],[AI_HL_NETSTATUS],[IT_DOM_CODD_PK_HX],[IT_HL_NETSTATUS],[ITRED_DOM_CODD_PK_HX],[ITRED_HL_NETSTATUS],[I7_DOM_CODD_PK_HX],[I7_HL_NETSTATUS],[TOTALPK],[DOM_PK_IC],[DOM_PK_IT],[DOM_PK_AI],[DOM_PK_9W],[CODD],[ROI],[S2_HX],[IC_HX],[9W_HX],[AI_HX],[IT_HX],[I7_HX],[HX_BOOKINGS],[DOM_PK_S2],UPLOAD_TYPE) WHERE UPLOAD_TYPE = 'HX' AND COUNTRY  = '" + StrCountry + "' AND  MONTH = " + intMonth + " AND YEAR = " + intYear, objConLivedatabase);
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
                                cmd1 = new SqlCommand("DELETE FROM T_INC_NIDT_PRODUCTIVITY_EXPORTED OUTPUT  DELETED.[YEAR],DELETED.[MONTH],DELETED.[LCODE],DELETED.[CHAIN_CODE],DELETED.[CHAIN_NAME],DELETED.[OFFICEID],DELETED.[COUNTRY],DELETED.[PRODUCTIVITY_CODD_PK_HX],DELETED.[INTL],DELETED.[TTL_HL],DELETED.[HL_INTL],DELETED.[S2_DOM_CODD_PK_HX],DELETED.[S2_HL_NETSTATUS],DELETED.[IC_DOM_CODD_PK_HX],DELETED.[IC_HL_NETSTATUS],DELETED.[9W_DOM_CODD_PK_HX],DELETED.[9W_HL_NETSTATUS],DELETED.[AI_DOM_CODD_PK_HX],DELETED.[AI_HL_NETSTATUS],DELETED.[IT_DOM_CODD_PK_HX],DELETED.[IT_HL_NETSTATUS],DELETED.[ITRED_DOM_CODD_PK_HX],DELETED.[ITRED_HL_NETSTATUS],DELETED.[I7_DOM_CODD_PK_HX],DELETED.[I7_HL_NETSTATUS],DELETED.[TOTALPK],DELETED.[DOM_PK_IC],DELETED.[DOM_PK_IT],DELETED.[DOM_PK_AI],DELETED.[DOM_PK_9W],DELETED.[CODD],DELETED.[ROI],DELETED.[S2_HX],DELETED.[IC_HX],DELETED.[9W_HX],DELETED.[AI_HX],DELETED.[IT_HX],DELETED.[I7_HX],DELETED.[HX_BOOKINGS],DELETED.[DOM_PK_S2],DELETED.UPLOAD_TYPE INTO T_INC_NIDT_PRODUCTIVITY_EXPORTED_LOG ([YEAR],[MONTH],[LCODE],[CHAIN_CODE],[CHAIN_NAME],[OFFICEID],[COUNTRY],[PRODUCTIVITY_CODD_PK_HX],[INTL],[TTL_HL],[HL_INTL],[S2_DOM_CODD_PK_HX],[S2_HL_NETSTATUS],[IC_DOM_CODD_PK_HX],[IC_HL_NETSTATUS],[9W_DOM_CODD_PK_HX],[9W_HL_NETSTATUS],[AI_DOM_CODD_PK_HX],[AI_HL_NETSTATUS],[IT_DOM_CODD_PK_HX],[IT_HL_NETSTATUS],[ITRED_DOM_CODD_PK_HX],[ITRED_HL_NETSTATUS],[I7_DOM_CODD_PK_HX],[I7_HL_NETSTATUS],[TOTALPK],[DOM_PK_IC],[DOM_PK_IT],[DOM_PK_AI],[DOM_PK_9W],[CODD],[ROI],[S2_HX],[IC_HX],[9W_HX],[AI_HX],[IT_HX],[I7_HX],[HX_BOOKINGS],[DOM_PK_S2],UPLOAD_TYPE) WHERE UPLOAD_TYPE = 'CODD' AND COUNTRY  = '" + StrCountry + "' AND  MONTH = " + intMonth + " AND YEAR = " + intYear, objConLivedatabase);
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
                        Cursor.Current = Cursors.Default;
                        return false;            
                    }
            }
            catch (Exception e1)
            {
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
                MessageBox.Show("error occured while increasing progressbar : " + exec1.Message );
	        } 
        }   
        #endregion 


        #region DoDisableSubItems
        private void DoDisableSubItems(ToolStripMenuItem item, Boolean blnSubItem)
        {
            foreach (ToolStripMenuItem  subitem in item.DropDownItems)
            {
                subitem.Enabled= blnSubItem;
            }
        }
        #endregion        

        #region GetConecitonString to Connect Database
            private void ConnectToDB()
            {
                try      
                {
                    String strMyString = String.Empty;
                    strMyString = ConfigurationManager.ConnectionStrings["QueryShedularConnectionString"].ToString();
                    objConCODD = new SqlConnection(strMyString);
                    objConHL = new SqlConnection(strMyString);
                    objConHX = new SqlConnection(strMyString);
                    objConLivedatabase = new SqlConnection(ConfigurationManager.ConnectionStrings["QueryShedularConnectionStringLive"].ToString());
          
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

        #region CheckBox Event
            void chkHL_CheckedChanged(object sender, EventArgs e)
            {
                 try
                {
                    if (chkHL.Checked==true){panelHL.Enabled=true;}
                    else if (chkHL.Checked==false){panelHL.Enabled=false;}
                     lblHL.Visible=false;
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
                    if (chkHX.Checked==true){panelHX.Enabled=true;}
                    else if (chkHX.Checked==false){panelHX.Enabled=false;}
                    lblHX.Visible=false;
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
                    if (chkCODD.Checked==true){panelCODD.Enabled=true;}
                    else if (chkCODD.Checked==false){panelCODD.Enabled=false;}
                    lblC.Visible=false;
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
                panelCODD.Enabled=false;
                panelHL.Enabled=false;
                panelHX.Enabled=false;
                lblC.Visible=false;
                lblHL.Visible=false;
                lblHX.Visible=false;
                btnExport.Enabled=false;
            }
        #endregion

        #region clearControl
        private void clearControls()
        {
            coddstime.Text="";
            coddetime.Text="";

            hlstime.Text="";
            hletime.Text="";

            hxstime.Text="";
            hxetime.Text="";
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
                    QueryClass.fillMonth(drpCMonth);
                    QueryClass.fillMonth(drpHLMonth);
                    QueryClass.fillMonth(drpHXMonth);
                    QueryClass.fillCountry(drpCcountry,"");
                    QueryClass.fillCountry(drpHXcountry,"");
                    QueryClass.fillCountry(drpHLcountry,"IN");
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
            lblTimer.Text= String.Format("{0:T}",DateTime.Now);
            if (isExecutingCODD || isExecutingHL || isExecutingHX)
            {}
            else{timer1.Enabled=false;}
        }
        #endregion

        #region trackBar1_Scroll
        private void trackBar1_Scroll(object sender, EventArgs e)
        {
            //this.toolStripStatusLabel1.Text =((System.Windows.Forms.TrackBar)(sender)).Value.ToString();
            if (((System.Windows.Forms.TrackBar)(sender)).Value >1)
            this.Opacity = ((float)((System.Windows.Forms.TrackBar)(sender)).Value/10);
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
                btnExport.ContextMenuStrip=contextMenuStrip1;
                Cursor.Position = new Point(Cursor.Position.X , Cursor.Position.Y );
                this.contextMenuStrip1.Show(btnExport, btnExport.PointToClient(Cursor.Position));                
            }
            catch (Exception exep )
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
                if (item.HasDropDownItems )
                {
                    if (item.Text.ToUpper().Trim() =="CODD")
                    {
                        item.Enabled=false;
                        DoDisableSubItems(item,false);
                    }                            
                }
            }

            foreach (ToolStripMenuItem item in contextMenuStrip1.Items)
            {
                if (item.HasDropDownItems )
                {
                    if (item.Text.ToUpper().Trim() =="HX")
                    {
                        item.Enabled=false;
                        DoDisableSubItems(item,false);
                    }                            
                }
            }
        
            foreach (ToolStripMenuItem item in contextMenuStrip1.Items)
            {
                if (item.HasDropDownItems )
                {
                    if (item.Text.ToUpper().Trim() =="HL")
                    {
                        item.Enabled=false;
                        DoDisableSubItems(item,false);
                    }                            
                }
            }                            
        }
        #endregion        
    }
}





/*'Copyright notice: ã 2004 by Bird Information Systems Pvt. Ltd. All rights reserved.
'********************************************************************************************
' This file contains trade secrets of Bird Information Systems. No part
' may be reproduced or transmitted in any form by any means or for any purpose
' without the express written permission of Bird Information Systems.
'********************************************************************************************
'$Author: Neeraj $Logfile: /AAMS/Queryschedular/ExportToExcel_WIN_App $
'$Workfile: ExportToExcel_WIN_App $
'$Revision: 1 $
'$Archive: /AAMS/Queryschedular/ExportToExcel_WIN_App $
'$Modtime: 6/15/10 2:57p $
*/

using System;
using System.Collections.Generic;
using System.Text;
using System.ComponentModel;
using System.Data;
using System.Configuration;
using System.Drawing;
using System.Diagnostics;
using System.Windows.Forms;
using System.IO;

namespace QueryShedular_HL_HX_CODD
{
    public class ExportToExcel_WIN_App
    {
        Microsoft.Office.Interop.Excel.Application excel;

        public Boolean ExportToExcel(DataTable objDT, String strMonthYear, String QueryType, String Countryname, String FileSavePath, ToolStripStatusLabel toolStripStatusLabel1)
        {
        int columnCount ;
        Int64 rowCount;
        Boolean boolRec;
        boolRec=false;
        
        try
        {
            object[,] stringArray = new object[objDT.Rows.Count + 1,objDT.Columns.Count ];

            // Add Column Header
            for (int i =0 ;i< objDT.Columns.Count ;i++)
            {
                stringArray[0, i] = objDT.Columns[i].ColumnName.ToString();
            }

            // Add Column data
            for(int row = 0; row < objDT.Rows.Count; ++row)
            {
                for(int col = 0; col < objDT.Columns.Count; col++)
                {
                    stringArray[row + 1, col] = objDT.Rows[row][col].ToString();
                }
            }

        excel = new Microsoft.Office.Interop.Excel.Application();
        if (excel==null)
        {
            MessageBox.Show("ERROR: EXCEL couldn't be started! ","Amadeus agent management system",MessageBoxButtons.OK,MessageBoxIcon.Warning);                
        }
        if (Countryname=="0"){Countryname="";}

            //String strFilename = @"c:\";
            String strFilename = FileSavePath + @"\";
            strFilename = strFilename+QueryType+Countryname+strMonthYear+"_"+(String.Format("{0:T}",DateTime.Now).Replace(":",""));
            strFilename = strFilename.Replace(" ","");
            strFilename = strFilename + ".xls";
            excel.Application.Workbooks.Add(true);
                        
            //excel.Visible = true;

            Microsoft.Office.Interop.Excel.Worksheet worksheet = (Microsoft.Office.Interop.Excel.Worksheet)excel.ActiveSheet;            
            worksheet.Activate();

            //worksheet.get_Range("A1:Z" + objDT.Rows.Count  , Type.Missing).Value2= stringArray;
            // To find the last cell index, we do the following thing.
            columnCount = objDT.Columns.Count;
            rowCount = objDT.Rows.Count;
            rowCount = rowCount + 1;
            if (rowCount!=0)
            {
                boolRec = true;
                string lastColumn = GetLastColumnName(columnCount);
                string usedRange = "A1:" + lastColumn + rowCount.ToString();              
                Microsoft.Office.Interop.Excel.Range oRange;                
                oRange = worksheet.get_Range(usedRange , Type.Missing);
                oRange.Value2= stringArray;
                
                //***************Setting Number Format***************************************
                //Selection.TextToColumns Destination:=Range("C1"), DataType:=xlDelimited, _
                //TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
                //Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
                //:=Array(1, 1), TrailingMinusNumbers:=True

                //string usedNumberRange = "C2:" + lastColumn + rowCount.ToString();                
                //oRange = worksheet.get_Range(usedNumberRange, Type.Missing);
                //oRange.Select();
                //oRange.TextToColumns("C2", Microsoft.Office.Interop.Excel.XlTextParsingType.xlDelimited, Microsoft.Office.Interop.Excel.XlTextQualifier.xlTextQualifierDoubleQuote,false,true, false,false, false, false,false,Microsoft.Office.Interop.Excel.XlColumnDataType.xlTextFormat,false,false,true);
                
                //oRange.NumberFormat="$0.00";
                //**************************************************************************

                excel.ActiveCell.Worksheet.SaveAs(strFilename,Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excel);

                //excel = null;
                //MessageBox.Show("Data's are exported to Excel Succesfully in '" + strFilename + "'","Amadeus agent management system",MessageBoxButtons.OK, MessageBoxIcon.Information);
                //Microsoft.Office.Interop.Excel.Workbook xlWorkBook;
                //Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet;
                //Microsoft.Office.Interop.Excel.Application xlApp;
                //xlApp = new Microsoft.Office.Interop.Excel.Application();
                //int[,] squareArray = new int[1,1];
                //xlWorkBook =  xlApp.Workbooks.Open(strFilename,Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,Type.Missing);
                //xlWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)xlWorkBook.ActiveSheet;
                //xlApp.Visible=true;
                //string usedNumberRange = "C2:" + lastColumn + rowCount.ToString();                
                //oRange = xlWorkSheet.get_Range(usedNumberRange, Type.Missing);
                //oRange.Select();
                //oRange.TextToColumns("C2", Microsoft.Office.Interop.Excel.XlTextParsingType.xlDelimited, Microsoft.Office.Interop.Excel.XlTextQualifier.xlTextQualifierDoubleQuote,false,true, false,false, false, false,false,squareArray,Type.Missing,Type.Missing,true);
                //xlWorkSheet As Excel.Worksheet

            }
            else
                toolStripStatusLabel1.Text="";
        }
            
        catch (Exception exec )
        {
            if (excel == null)
            {
                toolStripStatusLabel1.Text= "EXCEL couldn't be started ,please install office !";                
                //MessageBox.Show("ERROR: EXCEL couldn't be started ,please install office !", "Amadeus agent management system", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            else
                //MessageBox.Show(exec.Message, "Amadeus agent management system", MessageBoxButtons.OK, MessageBoxIcon.Warning);                
                toolStripStatusLabel1.Text = exec.Message.ToString();                

        }
        finally
        {
            foreach (System.Diagnostics.Process pr  in System.Diagnostics.Process.GetProcessesByName("EXCEL"))
            {
                pr.Kill();            
            }               
            
        }
            return boolRec;
        }

        private string GetLastColumnName(int lastColumnIndex)
        {
            string lastColumn = "";
            // check whether the column count is > 26
            if (lastColumnIndex > 26)
            {
            // If the column count is > 26, the the last column index will be something 
            // like "AA", "DE", "BC" etc

            // Get the first letter
            // ASCII index 65 represent char. 'A'. So, we use 64 in this calculation as a starting point
            char first = Convert.ToChar(64 + ((lastColumnIndex - 1)/26));

            // Get the second letter
            char second = Convert.ToChar(64 + (lastColumnIndex%26 == 0? 26 : lastColumnIndex%26));

            // Concat. them
            lastColumn = first.ToString() + second.ToString();
            }
            else
            {
            // ASCII index 65 represent char. 'A'. So, we use 64 in this calculation as a starting point
            lastColumn = Convert.ToChar(64 + lastColumnIndex).ToString();
            }
            return lastColumn;
        }

        public Boolean Exportcsv(DataTable objDT, String strType,String Countryname,String strMonthYear, String strFilename,ToolStripStatusLabel toolStripStatusLabel1)
        {
	        String str = "";
            Boolean boolRec=false;

            strFilename = strFilename+"/"+strType+Countryname+strMonthYear+"_"+(String.Format("{0:T}",DateTime.Now).Replace(":",""));
            strFilename = strFilename.Replace(" ","");
            strFilename = strFilename + ".csv";

	        if (File.Exists(strFilename)) {
		        File.Delete(strFilename);
	        }
	        FileStream objfilestream = new FileStream(strFilename, FileMode.Create, FileAccess.Write);
	        StreamWriter objFileWriter = new StreamWriter(objfilestream);

	        try 
            {
		        foreach (DataColumn c in objDT.Columns) 
                {
			        str = str.ToString() + c.ColumnName.ToString() +   ",".ToString() ;
		        }
		        objFileWriter.WriteLine(str);
                foreach (DataRow r in objDT.Rows) {
			        str = "";
			        for (int i = 0; i <= objDT.Columns.Count - 1; i++) 
                    {
				        str = str + r[i].ToString() + ",".ToString();
			        }
			        objFileWriter.WriteLine(str);
		        }
                boolRec=true;                
	        } 
            catch (Exception ex) 
            {
                toolStripStatusLabel1.Text = ex.Message.ToString();  
                boolRec=false;
	        } 
            finally {
		        objFileWriter.Flush();
		        objFileWriter.Close();
		        objFileWriter = null;
		        objfilestream = null;
	        }
            return boolRec;
        }
    }
}
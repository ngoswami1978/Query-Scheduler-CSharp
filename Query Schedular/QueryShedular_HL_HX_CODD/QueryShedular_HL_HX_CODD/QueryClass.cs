/*'Copyright notice: ã 2004 by Bird Information Systems Pvt. Ltd. All rights reserved.
'********************************************************************************************
' This file contains trade secrets of Bird Information Systems. No part
' may be reproduced or transmitted in any form by any means or for any purpose
' without the express written permission of Bird Information Systems.
'********************************************************************************************
'$Author: Neeraj $Logfile: /AAMS/Queryschedular/QueryClass.cs $
'$Workfile: QueryClass.cs $
'$Revision: 1 $
'$Archive: /AAMS/Queryschedular/QueryClass.cs $
'$Modtime: 6/15/10 2:57p $
*/

using System;
using System.Collections.Generic;
using System.Text;
using System.Data;
using System.Windows.Forms;
using System.Xml;

namespace QueryShedular_HL_HX_CODD
{
    public class QueryClass
    {
        public static void fillYear(ComboBox objCmbYear)
        {
            for (int i = DateTime.Now.Year; i > DateTime.Now.Year-4;i--)
            {
                objCmbYear.Items.Add(i.ToString());
            }           
            objCmbYear.Items.Insert(0,"Select one");
            objCmbYear.SelectedIndex=0;
        }
        public static void fillMonth(ComboBox objCmbMnth)
        {
            for (int i =1 ; i < 13;i++)
            {
                DateTime date = new DateTime(1900, i, 1);                  
                objCmbMnth.Items.Add(date.ToString("MMMM").ToString());
            }           
            objCmbMnth.Items.Insert(0,"Select one");
            objCmbMnth.SelectedIndex=0;
        }
        public static void fillCountry(ComboBox objCmbCountry,String CountryCode)
        {
            //DataTable objDT = new DataTable("Country");
            //DataRow objRow;
            //DataColumn objcol =new DataColumn("ContryName",typeof(System.String));
            //objDT.Columns.Add(objcol);

            //objcol =new DataColumn("ContryId",typeof(System.String));
            //objDT.Columns.Add(objcol);

            XmlNodeReader objXmlReader;
            DataSet ds = new DataSet();
            XmlDocument  objOutXml = new XmlDocument();
            String strInput = "<MS_LISTCOUNTRY_OUTPUT><COUNTRY  COUNTRY_CODE='0' COUNTRY_NAME='Select one'/><COUNTRY  COUNTRY_CODE='BD' COUNTRY_NAME='Bangladesh'/><COUNTRY  COUNTRY_CODE='IN' COUNTRY_NAME='India' /><COUNTRY  COUNTRY_CODE='NP' COUNTRY_NAME='Nepal' /> <COUNTRY  COUNTRY_CODE='LK' COUNTRY_NAME='Srilanka' /><COUNTRY  COUNTRY_CODE='BT' COUNTRY_NAME='Bhutan' /><COUNTRY  COUNTRY_CODE='ML' COUNTRY_NAME='Maldives' /><COUNTRY  COUNTRY_CODE='TB' COUNTRY_NAME='TBA' /> <Errors Status='False'><Error Code='' Description='' /></Errors></MS_LISTCOUNTRY_OUTPUT>";
            objOutXml.LoadXml(strInput);

            objXmlReader = new XmlNodeReader(objOutXml);
            ds.ReadXml(objXmlReader);
            objCmbCountry.DataSource = ds.Tables["COUNTRY"];
            objCmbCountry.DisplayMember = "COUNTRY_NAME";
            objCmbCountry.ValueMember = "COUNTRY_CODE";
            objCmbCountry.SelectedIndex = 0;
            if (CountryCode !="")
            {
                objCmbCountry.SelectedIndex = objCmbCountry.FindString("India");
            }
        }
    }
}

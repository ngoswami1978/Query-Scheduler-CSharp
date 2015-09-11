using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace QueryShedular_HL_HX_CODD
{
    public partial class frmReportSetting : Form
    {
        public frmReportSetting()
        {
            InitializeComponent();
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            // An item must be selected
            if (listBox1.SelectedItems.Count > 0)
            {
                object selected = listBox1.SelectedItem;
                int indx = listBox1.Items.IndexOf(selected);
                int totl = listBox1.Items.Count;
                // If the item is right at the top, throw it right down to the bottom
                if (indx == 0)
                {
                    listBox1.Items.Remove(selected);
                    listBox1.Items.Insert(totl - 1, selected);
                    listBox1.SetSelected(totl - 1, true);
                }
                // To move the selected item upwards in the listbox
                else
                {
                    listBox1.Items.Remove(selected);
                    listBox1.Items.Insert(indx - 1, selected);
                    listBox1.SetSelected(indx - 1, true);
                }
            }
            QueryClass.fillReportColuimn(listBox1);            
        }

        private void button2_Click_1(object sender, EventArgs e)
        {
            // An item must be selected
            if (listBox1.SelectedItems.Count > 0)
            {
                object selected = listBox1.SelectedItem;
                int indx = listBox1.Items.IndexOf(selected);
                int totl = listBox1.Items.Count;
                // If the item is last in the listbox, move it all the way to the top
                if (indx == totl - 1)
                {
                    listBox1.Items.Remove(selected);
                    listBox1.Items.Insert(0, selected);
                    listBox1.SetSelected(0, true);
                }
                // To move the selected item downwards in the listbox
                else
                {
                    listBox1.Items.Remove(selected);
                    listBox1.Items.Insert(indx + 1, selected);
                    listBox1.SetSelected(indx + 1, true);
                }
            }
            QueryClass.fillReportColuimn(listBox1);
        }
        private void frmReportSetting_Load(object sender, EventArgs e)
        {            
            listBox1.Items.Clear();
            foreach (string str in QueryClass.lstReportColumn)
            {
                listBox1.Items.Add(str);
            }
        }               
    }
}
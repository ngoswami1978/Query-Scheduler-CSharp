using System;
using System.Collections.Generic;
using System.Windows.Forms;

namespace QueryShedular_HL_HX_CODD
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the appliA tion.
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new frmQueryShedular());
        }
    }
}
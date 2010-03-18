using System;
using System.Windows.Forms;
using Microsoft.VisualStudio.Tools.Applications.Runtime;
using Word = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;
using System.IO;

namespace Word2003Tools4Dominique
{
    public partial class ThisAddIn
    {
        public const string ADDIN_TITLE = "Word Index";
        CustomMenu menu;


        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            try
            {
                // Menu creation
                menu = new CustomMenu(this.Application);
                // Create default filters file
                if (!File.Exists(Util.GetFiltersFilePath())) Util.GenerateDefaultFiltersFile();

            }
            catch (Exception exception)
            {
                MessageBox.Show(exception.Message);
            }
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            try
            {
                // Delete WordIndex menu
                menu.Delete(this.Application);
            }
            catch (Exception exception)
            {
                MessageBox.Show(exception.Message);
            }
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}

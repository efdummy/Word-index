using System;
using System.Collections.Generic;
using System.Text;
using Microsoft.VisualStudio.Tools.Applications.Runtime;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Interop.Word;
using System.ComponentModel;

namespace Word2003Tools4Dominique
{
    public class CustomMenu
    {
        // missing
        private global::System.Object missing = global::System.Type.Missing;

        // Used by index filter
        private int _MinWordLength = 4;

        // Commandes de menu
        Office.CommandBarButton menuCommand00;
        //Office.CommandBarButton menuCommand01;
        //Office.CommandBarButton menuCommand02;
        Office.CommandBarButton menuCommand03;
        Office.CommandBarButton menuCommand04;
        //Office.CommandBarButton menuCommand05;
        //Office.CommandBarButton menuCommand06;
        //Office.CommandBarButton menuCommand07;
        //Office.CommandBarButton menuCommand08;


        Office.CommandBarButton menuCommand20;

        // Nom et éléments du menu
        private string MENU_ID = "Fullenbaums-WordIndex";
        private string MENU_CAPTION = "Word Inde&x";
        private string MENU_ITEM0_ID = "MI00";
        private string MENU_ITEM3_ID = "MI03";
        private string MENU_ITEM4_ID = "MI04";
        //private string MENU_ITEM5_ID = "MI05";
        //private string MENU_ITEM6_ID = "MI06";
        //private string MENU_ITEM7_ID = "MI07";
        //private string MENU_ITEM8_ID = "MI08";

        private string MENU_ITEM20_ID = "MI20";

        private string MENU_ITEM0_CAPTION = "&Extract words according pattern...";
        private string MENU_ITEM3_CAPTION = "Extract all &words for index...";
        private string MENU_ITEM4_CAPTION = "Edit index &filters...";
        //private string MENU_ITEM5_CAPTION = "";
        //private string MENU_ITEM6_CAPTION = "&";
        //private string MENU_ITEM7_CAPTION = "&";
        //private string MENU_ITEM8_CAPTION = "&";


        private string MENU_ITEM20_CAPTION = "About Word &Index...";
        
        private int MENU_ITEM0_FACEID = 0;
        //private int MENU_ITEM1_FACEID = 576;
        //private int MENU_ITEM2_FACEID = 576;
        private int MENU_ITEM3_FACEID = 0;
        private int MENU_ITEM4_FACEID = 0;
        //private int MENU_ITEM5_FACEID = 0;
        //private int MENU_ITEM6_FACEID = 0;
        //private int MENU_ITEM7_FACEID = 0;
        //private int MENU_ITEM8_FACEID = 0;

        private int MENU_ITEM20_FACEID = 0;

        // Menu word
        Application _appli;
        Office.CommandBarPopup _menu;

        public CustomMenu(Application appli)
        {
            // Suppression du menu personnalisé s'il existe
            CheckIfMenuBarExistsAndDelete(appli, MENU_ID);
            // Création du menu personnalisé
            AddMenuBar(appli);
            // Mémorisation de l'appli et du menu
            _appli = appli;
            _menu = (Office.CommandBarPopup)
                    appli.CommandBars.ActiveMenuBar.FindControl(
                    Office.MsoControlType.msoControlPopup, System.Type.Missing, MENU_ID, true, true);
        }
        public void SetMinWordLength(int min)
        {
            _MinWordLength = min;
        }
        public int GetMinWordLength()
        {
            return _MinWordLength;
        }


        // If the menu already exists, remove it.
        private void CheckIfMenuBarExistsAndDelete(Application appli, string menuId)
        {
                Office.CommandBarPopup foundMenu = (Office.CommandBarPopup)
                    appli.CommandBars.ActiveMenuBar.FindControl(
                    Office.MsoControlType.msoControlPopup, System.Type.Missing, menuId, true, true);

                if (foundMenu != null)
                {
                    foundMenu.Delete(true);
                }
        }
        public void EnableMenu(bool Enabled)
        {
            _menu.Enabled = Enabled;
        }
        public void Delete(Application appli)
        {
            CheckIfMenuBarExistsAndDelete(appli, MENU_ID);
        }

        // Create the menu, if it does not exist.
        private void AddMenuBar(Application appli)
        {
                Office.CommandBarPopup cmdBarControl = null;
                Office.CommandBar menubar = (Office.CommandBar)appli.CommandBars.ActiveMenuBar;
                int controlCount = menubar.Controls.Count;
                string menuCaption = MENU_CAPTION;

                // Add the menu.
                cmdBarControl = (Office.CommandBarPopup)menubar.Controls.Add(
                    Office.MsoControlType.msoControlPopup, missing, missing, controlCount, missing);

                if (cmdBarControl != null)
                {
                    cmdBarControl.Caption = menuCaption;
                    cmdBarControl.Tag = MENU_ID;


                    // Add the menu's commands.
                    AddMenuCommand(cmdBarControl, ref menuCommand00, MENU_ITEM0_ID, MENU_ITEM0_CAPTION, new Microsoft.Office.Core._CommandBarButtonEvents_ClickEventHandler(menuCommand0_Click), MENU_ITEM0_FACEID);
                    AddMenuCommand(cmdBarControl, ref menuCommand03, MENU_ITEM3_ID, MENU_ITEM3_CAPTION, new Microsoft.Office.Core._CommandBarButtonEvents_ClickEventHandler(menuCommand3_Click), MENU_ITEM3_FACEID);
                    AddMenuCommand(cmdBarControl, ref menuCommand04, MENU_ITEM4_ID, MENU_ITEM4_CAPTION, new Microsoft.Office.Core._CommandBarButtonEvents_ClickEventHandler(menuCommand4_Click), MENU_ITEM4_FACEID);
                    //AddMenuCommand(cmdBarControl, ref menuCommand05, MENU_ITEM5_ID, MENU_ITEM5_CAPTION, new Microsoft.Office.Core._CommandBarButtonEvents_ClickEventHandler(menuCommand5_Click), MENU_ITEM5_FACEID);
                    //AddMenuCommand(cmdBarControl, ref menuCommand06, MENU_ITEM6_ID, MENU_ITEM6_CAPTION, new Microsoft.Office.Core._CommandBarButtonEvents_ClickEventHandler(menuCommand6_Click), MENU_ITEM6_FACEID);
                    //AddMenuCommand(cmdBarControl, ref menuCommand07, MENU_ITEM7_ID, MENU_ITEM7_CAPTION, new Microsoft.Office.Core._CommandBarButtonEvents_ClickEventHandler(menuCommand7_Click), MENU_ITEM7_FACEID);
                    //AddMenuCommand(cmdBarControl, ref menuCommand08, MENU_ITEM8_ID, MENU_ITEM8_CAPTION, new Microsoft.Office.Core._CommandBarButtonEvents_ClickEventHandler(menuCommand8_Click), MENU_ITEM8_FACEID);

                    AddMenuCommand(cmdBarControl, ref menuCommand20, MENU_ITEM20_ID, MENU_ITEM20_CAPTION, new Microsoft.Office.Core._CommandBarButtonEvents_ClickEventHandler(menuCommand20_Click), MENU_ITEM20_FACEID);
                }
        }
        private void AddMenuCommand(Office.CommandBarPopup cmdBarControl, ref Office.CommandBarButton menuCommand, string cmdID, string cmdCAPTION, Microsoft.Office.Core._CommandBarButtonEvents_ClickEventHandler clickHandler, int faceID)
        {
            menuCommand = (Office.CommandBarButton)cmdBarControl.Controls.Add(
                Office.MsoControlType.msoControlButton, missing, missing, missing, missing);
            //menuCommand = cmdBarControl.Controls.Add(
            //    Office.MsoControlType.msoControlButton, missing, missing, missing, missing);
            menuCommand.Caption = cmdCAPTION;
            menuCommand.Tag = cmdID;
            menuCommand.FaceId = faceID;
            menuCommand.Click += clickHandler;
        }
        
        const string I18N_MINUS = "a-zàáâãäåæçèéêëìíîïðñòóôõöøùúûüýþÿ";
        const string I18N_CAPS  = "A-ZÀÁÂÃÄÅÆÇÈÉÊËÌÍÎÏÐÑÒÓÔÕÖØÙÚÛÜÝÞ";
        //const string SUSPICIOUS_INITIALS_PATTERN = "Mc|[" + I18N_CAPS + "][.][" + I18N_CAPS + "][.][" + I18N_CAPS + "][.]|[" + I18N_CAPS + "][.][" + I18N_CAPS + "][.]";
        //const string AUTHORS_PATTERN_ALL = "[" + I18N_CAPS + "][-" + I18N_CAPS + I18N_MINUS +" ]*[,][ \u00A0](([" + I18N_CAPS + "][.][- \u00A0][" + I18N_CAPS + "][.][- \u00A0][" + I18N_CAPS + "][.])|([" + I18N_CAPS + "][.][- \u00A0][" + I18N_CAPS + "][.])|([" + I18N_CAPS + "][.]))";
        const string AUTHORS_PATTERN_CAPS = "[" + I18N_CAPS + "][-" + I18N_CAPS + " \u00A0]*[,][ \u00A0](([" + I18N_CAPS + "][.][- \u00A0][" + I18N_CAPS + "][.][- \u00A0][" + I18N_CAPS + "][.])|([" + I18N_CAPS + "][.][- \u00A0][" + I18N_CAPS + "][.])|([" + I18N_CAPS + "][.]))";
        //const string WORDS_PATTERN = "[^\\n\\r\\s\\[\\]\\?:,‚;'’‘()\"“”„«»…§&><!*/=+.0-9]+";
        const string WORDS_PATTERN = @"([^0-9\W]|[-])+";

        private void menuCommand0_Click(Microsoft.Office.Core.CommandBarButton Ctrl, ref bool CancelDefault)
        {
            // Extract authors' names (NNNN I. I. pattern)
            Application appli = (Application)Ctrl.Application;
            string strText = "";
            if (appli != null) if (appli.Selection != null) strText = appli.Selection.Text;
            /*
            progressForm.Start(Commands.DoExtraction, ThreadParamsAndProgress.DoNothingToDo, new ThreadParameters(null, strText, AUTHORS_PATTERN_ALL, false, null));
             */
            ExtractForm fm = new ExtractForm(strText);
            fm.ShowDialog();
        }

        
        
        
        
        private void menuCommand3_Click(Microsoft.Office.Core.CommandBarButton Ctrl, ref bool CancelDefault)
        {
            // Open ProgressForm and Background workers manager
            ThreadParamsAndProgress progressForm = new ThreadParamsAndProgress();
            // Get all selected text
            Application appli = (Application)Ctrl.Application;
            string strText = "";
            if (appli != null) if (appli.Selection != null) strText = appli.Selection.Text;
            // Parse all words
            List<string> wordListAll = new List<string>();
            int rc1 = Commands.ExtractPattern(strText, WORDS_PATTERN, true, ref wordListAll);

            // Multi-threading represents a boost only if there are many words to process
            int count = wordListAll.Count;
            int sliceSize;
            int remainder;
            ThreadParamsAndProgress.DoFinalizeThreadsWork DoFinalize = new ThreadParamsAndProgress.DoFinalizeThreadsWork(Commands.IndexFinalize);
            if (count > 3) 
            {
                switch (Environment.ProcessorCount)
                {
                    case 2:
                        // Prepare parameters
                        sliceSize = count / 2;
                        Math.DivRem(count, 2, out remainder);
                        List<string> wordlist21 = wordListAll.GetRange(0, sliceSize);
                        List<string> wordlist22 = wordListAll.GetRange(sliceSize, sliceSize+remainder);
                        ThreadParameters p21 = new ThreadParameters(null, strText, WORDS_PATTERN, GetMinWordLength(), true, false, wordlist21);
                        ThreadParameters p22 = new ThreadParameters(null, strText, WORDS_PATTERN, GetMinWordLength(), true, false, wordlist22);
                        // Start threads
                        progressForm.Start(Commands.DoExtractionForIndexFiltered, DoFinalize, p21, p22);
                        break;
                    case 3:
                        // Prepare parameters
                        sliceSize = count / 3;
                        Math.DivRem(count, 3, out remainder);
                        List<string> wordlist31 = wordListAll.GetRange(0, sliceSize);
                        List<string> wordlist32 = wordListAll.GetRange(sliceSize, sliceSize);
                        List<string> wordlist33 = wordListAll.GetRange(2 * sliceSize, sliceSize + remainder);
                        ThreadParameters p31 = new ThreadParameters(null, strText, WORDS_PATTERN, GetMinWordLength(), true, false, wordlist31);
                        ThreadParameters p32 = new ThreadParameters(null, strText, WORDS_PATTERN, GetMinWordLength(), true, false, wordlist32);
                        ThreadParameters p33 = new ThreadParameters(null, strText, WORDS_PATTERN, GetMinWordLength(), true, false, wordlist33);
                        // Start threads
                        progressForm.Start(Commands.DoExtractionForIndexFiltered, DoFinalize, p31, p32, p33);
                        break;
                    case 4:
                    case 5:
                    case 6:
                    case 7:
                    case 8:
                        // Prepare parameters
                        sliceSize = count / 4;
                        Math.DivRem(count, 4, out remainder);
                        List<string> wordlist41 = wordListAll.GetRange(0, sliceSize);
                        List<string> wordlist42 = wordListAll.GetRange(sliceSize, sliceSize);
                        List<string> wordlist43 = wordListAll.GetRange(2 * sliceSize, sliceSize);
                        List<string> wordlist44 = wordListAll.GetRange(3 * sliceSize, sliceSize + remainder);
                        ThreadParameters p41 = new ThreadParameters(null, strText, WORDS_PATTERN, GetMinWordLength(), true, false, wordlist41);
                        ThreadParameters p42 = new ThreadParameters(null, strText, WORDS_PATTERN, GetMinWordLength(), true, false, wordlist42);
                        ThreadParameters p43 = new ThreadParameters(null, strText, WORDS_PATTERN, GetMinWordLength(), true, false, wordlist43);
                        ThreadParameters p44 = new ThreadParameters(null, strText, WORDS_PATTERN, GetMinWordLength(), true, false, wordlist44);
                        // Start threads
                        progressForm.Start(Commands.DoExtractionForIndexFiltered, DoFinalize, p41, p42, p43, p44);
                        break;
                    default:
                        ThreadParameters p1 = new ThreadParameters(null, strText, WORDS_PATTERN, GetMinWordLength(), true, false, wordListAll);
                        progressForm.Start(Commands.DoExtractionForIndexFiltered, DoFinalize, p1);
                        break;
                }
            }
            else
            {
                ThreadParameters p1 = new ThreadParameters(null, strText, WORDS_PATTERN, GetMinWordLength(), true, false, wordListAll);
                progressForm.Start(Commands.DoExtractionForIndexFiltered, DoFinalize, p1);
            }
        }

        private void menuCommand4_Click(Microsoft.Office.Core.CommandBarButton Ctrl, ref bool CancelDefault)
        {
            FiltersForm fm = new FiltersForm(GetMinWordLength());
            fm.ShowDialog();
            SetMinWordLength(fm.MinValue);
        }


        private void menuCommand20_Click(Microsoft.Office.Core.CommandBarButton Ctrl, ref bool CancelDefault)
        {
            // Display version
            Util.DisplayVersion();
        }
    }
}

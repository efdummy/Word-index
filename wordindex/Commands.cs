using System;
using System.Collections.Generic;
using System.Text;
using Microsoft.Office.Interop.Word;
using System.Windows.Forms;
using System.Text.RegularExpressions;
using System.IO;
using System.Reflection;
using System.ComponentModel;


namespace Word2003Tools4Dominique
{
    static class Commands
    {
        // Tout les combien de fois on informe sur la progression
        const int REPORT_FACTOR = 40;

        const int MAX_AUTHOR_INITIALS_LENGTH = 33;

        //_________________________________________________________________________________________________________

        //                     DELEGATES THAT RUN ON BACKGROUND WORKER THREADS

        //

        // Extraction Delegate that run on the background worker thread
        public static void DoExtraction(object sender, DoWorkEventArgs e)
        {
            BackgroundWorker worker = sender as BackgroundWorker;
            ThreadParameters p = (ThreadParameters)e.Argument;
            string resultFilePath = Util.GetResultFilePath(p.ThreadNum);
            List<string> wordlistresult = new List<string>();
            int rc = Commands.ExtractPattern(worker, p.Text, p.Pattern, p.IgnoreCase, ref wordlistresult);
            Util.WriteToFile(resultFilePath, wordlistresult, Msg.MSG_NOTHING_FOUND_OR_NO_TEXT);
            e.Result = rc;
            if (rc != 0)
            {
                e.Cancel = true;
            }
            else
            {
                Util.SortUniq(resultFilePath, Util.GetSortedUniqResultFilePath());
                Util.DisplayFile(Util.GetSortedUniqResultFilePath());
            }
            e.Cancel = true;
            worker.CancelAsync();
        }

        // Extract Delegate for index that run on the background worker thread
        public static void DoExtractionForIndexFiltered(object sender, DoWorkEventArgs e)
        {
            BackgroundWorker worker = sender as BackgroundWorker;
            ThreadParameters p = (ThreadParameters)e.Argument;
            List<string> wordListAccepted = new List<string>();
            List<string> wordListRejected = new List<string>();
            int rc = Commands.ExtractWordsForIndexFiltered(worker, p.StringList, p.Pattern, ref wordListAccepted, ref wordListRejected, Util.GetFilters(), p.MinWordLength);
            if (rc != 0) e.Result = rc;
            if (rc != 0)
            {
                e.Cancel = true;
            }
            else
            {
                Util.SortedUniqToFile(wordListAccepted, Util.GetResultFilePath(p.ThreadNum), Msg.MSG_NOTHING_FOUND_OR_NO_TEXT);
                Util.SortedUniqToFile(wordListRejected, Util.GetRejectedFilePath(p.ThreadNum), Msg.MSG_NOTHING_REJECTED);
            }
            e.Cancel = true;
            worker.CancelAsync();
        }

        // Edition des fichiers une fois que tous les threads sont terminés
        public static void IndexFinalize(ThreadParameters p)
        {
            File.Delete(Util.GetFinalRejectedFilePath());
            File.Delete(Util.GetFinalResultFilePath());
            for (int i = 0; i < p.NumberOfWorkingThreads; i++)
            {
                Util.AppendFiles(Util.GetRejectedFilePath(i), Util.GetFinalRejectedFilePath());
                Util.AppendFiles(Util.GetResultFilePath(i), Util.GetFinalResultFilePath());
            }
            Util.SortUniq(Util.GetFinalRejectedFilePath(), Util.GetSortedUniqRejectedFilePath());
            Util.SortUniq(Util.GetFinalResultFilePath(), Util.GetSortedUniqResultFilePath());

            Util.DisplayFile(Util.GetSortedUniqRejectedFilePath());
            Util.DisplayFile(Util.GetSortedUniqResultFilePath());
        }

        //_________________________________________________________________________________________________________

        //                                    FUNCTIONS OF DELEGATES

        //

        public static int ExtractPattern(BackgroundWorker worker, string strText, string pattern, bool ignoreCase, ref List<string> wordlist)
        {
            int rc = 0;
            string strPattern = pattern;

            if (!String.IsNullOrEmpty(strText))
            {

                //___ Expression régulière à rechercher
                Regex regExp;
                if (ignoreCase)
                    regExp = new Regex(strPattern, RegexOptions.Compiled|RegexOptions.IgnoreCase);
                else
                    regExp = new Regex(strPattern, RegexOptions.Compiled);

                //___ Exécution de la recherche
                MatchCollection oMatches = regExp.Matches(strText);

                //___ S'il y a des résultats
                if (oMatches.Count != 0)
                {
                    int loopcount=0;
                    double total = oMatches.Count;
                    int res;

                    foreach (Match oMatch in oMatches)
                    {
                        loopcount++;
                        // Attention, on filtre les match de plus de 33 lettres
                        if (oMatch.Length < MAX_AUTHOR_INITIALS_LENGTH)
                        {
                            
                            if (!wordlist.Contains(oMatch.Value)) wordlist.Add(oMatch.Value);
                            // Indication de progression au background worker
                            res=Math.DivRem(loopcount, REPORT_FACTOR, out res);
                            if (res == 0)
                            worker.ReportProgress((int)Math.Truncate((loopcount / total) * 100), (object)oMatch.Value);
                        }
                        if (worker.CancellationPending)
                        {
                            rc = 1;
                            break;
                        }
                    }
                }
            }
            return rc;
        }

        public static int ExtractPattern(string strText, string pattern, bool ignoreCase, ref List<string> wordlist)
        {
            int rc = 0;
            string strPattern = pattern;

            if (!String.IsNullOrEmpty(strText))
            {
                //___ Expression régulière à rechercher
                Regex regExp;
                if (ignoreCase)
                    regExp = new Regex(strPattern, RegexOptions.Compiled|RegexOptions.IgnoreCase);
                else
                    regExp = new Regex(strPattern, RegexOptions.Compiled);
                //___ Exécution de la recherche
                MatchCollection oMatches = regExp.Matches(strText);
                //___ S'il y a des résultats
                if (oMatches.Count != 0)
                {
                    int loopcount = 0;
                    double total = oMatches.Count;
                    foreach (Match oMatch in oMatches)
                    {
                        loopcount++;
                        // Attention, on filtre les match de plus de 33 lettres
                        if (oMatch.Length < MAX_AUTHOR_INITIALS_LENGTH)
                        {

                            if (!wordlist.Contains(oMatch.Value)) wordlist.Add(oMatch.Value);
                        }
                    }
                }
            }
            return rc;
        }

        public static int ExtractWordsForIndexFiltered(BackgroundWorker worker, List<string> wordListInput, string pattern, ref List<string> acceptedWords, ref List<string> rejectedWords , List<string> filters, int minWordLength)
        {
            int rc = 0;

            //___ S'il y a des mots à traiter
            if (wordListInput.Count > 0)
            {
                int loopcount = 0;

                double total = wordListInput.Count;
                int res;
                    
                    foreach (string match in wordListInput)
                    {
                        if (isNotFiltered(match, filters, minWordLength))
                        {
                            if (!acceptedWords.Contains(match)) acceptedWords.Add(match);
                        }
                        else
                        {
                            if (!rejectedWords.Contains(match)) rejectedWords.Add(match);
                        }
                        // Indication de progression au background worker
                        loopcount++;
                        Math.DivRem(loopcount, REPORT_FACTOR, out res);
                        if (res == 0)
                            worker.ReportProgress((int)Math.Truncate((loopcount / total) * 100), (object)match);
                        if (worker.CancellationPending)
                        {
                            rc = 1;
                            break;
                        }
                    }
            }
            return rc;
        }

        private static bool isNotFiltered(string word, List<string> filters, int wordMinLength)
        {
            bool notExcluded = true;

            if (word.Length < wordMinLength)
                notExcluded = false;
            else
            {
                foreach (string filter in filters)
                {
                    Regex regExp = new Regex("^" + filter + "\\z", RegexOptions.Compiled | RegexOptions.IgnoreCase);
                    MatchCollection oMatches = regExp.Matches(word);
                    if (oMatches.Count > 0)
                    {
                        notExcluded = false;
                        break;
                    }
                }
            }
            return notExcluded;
        }

    }
}

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Windows.Forms;
using System.Reflection;

namespace Word2003Tools4Dominique
{
    static class Util
    {
        const string TEMP_DIR = "c:\\Temp";
        const string RESULT_FILE_NAME = "Word2003Tools.result";
        const string PROCESSED_OCCURRENCES_FILE_NAME = "Word2003Tools.ProcessedOccurrences";
        const string RESULT_FILE_NAME_REJECTED = "Word2003Tools.rejected";
        const string RESULT_FILE_EXT = ".txt";
        const string FILTERS_FILE = "filterspatterns-fr.txt";
        const string PATTERNS_FILE = "extractpatterns-fr.txt";

        public const string REGULAR_EXPRESSION_HELP_URL1 = "http://en.wikipedia.org/wiki/Regular_expression";
        public const string REGULAR_EXPRESSION_HELP_URL2 = "http://msdn.microsoft.com/en-us/library/1400241x(VS.85).aspx";

        public const string COMMENT_MARKER = " // ";
        // Path for files
        public static string GetResultFilePath(int ThreadNum)
        {
            if (!Directory.Exists(TEMP_DIR)) Directory.CreateDirectory(TEMP_DIR);
            return Path.Combine(TEMP_DIR, RESULT_FILE_NAME + ThreadNum + RESULT_FILE_EXT);
        }
        public static string GetProcessedOccurrencesFilePath()
        {
            if (!Directory.Exists(TEMP_DIR)) Directory.CreateDirectory(TEMP_DIR);
            return Path.Combine(TEMP_DIR, PROCESSED_OCCURRENCES_FILE_NAME + RESULT_FILE_EXT);
        }
        public static string GetRejectedFilePath(int ThreadNum)
        {
            return Path.Combine(TEMP_DIR, RESULT_FILE_NAME_REJECTED + ThreadNum + RESULT_FILE_EXT);
        }
        public static string GetFinalResultFilePath()
        {
            if (!Directory.Exists(TEMP_DIR)) Directory.CreateDirectory(TEMP_DIR);
            return Path.Combine(TEMP_DIR, RESULT_FILE_NAME + RESULT_FILE_EXT);
        }
        public static string GetFinalRejectedFilePath()
        {
            return Path.Combine(TEMP_DIR, RESULT_FILE_NAME_REJECTED + RESULT_FILE_EXT);
        }
        public static string GetSortedUniqResultFilePath()
        {
            return Path.Combine(TEMP_DIR, RESULT_FILE_NAME+".u.txt");
        }
        public static string GetSortedUniqRejectedFilePath()
        {
            return Path.Combine(TEMP_DIR, RESULT_FILE_NAME_REJECTED + ".u.txt");
        }
        
        // Write a string to a file
        public static void WriteToFile(string tempFilePath, string stringToWrite)
        {
            StreamWriter SW;
            SW = File.CreateText(tempFilePath);
            SW.Write(stringToWrite);
            SW.Close();
        }
        
        // Write a string list to a file
        public static void WriteToFile(string FilePath, List<string> LinesList, string EmptyListMsg)
        {
            StreamWriter SW;
            SW = File.CreateText(FilePath);
            if (LinesList.Count > 0)
                foreach (string line in LinesList) SW.WriteLine(line);
            else
                SW.WriteLine(EmptyListMsg);
            SW.Close();
        }

        // Append two files
        public static void AppendFiles(string SourceFilePath, string TargetFilePath)
        {
            using (TextWriter tw = new StreamWriter(TargetFilePath, true))
            {
                using (StreamReader tr = new StreamReader(SourceFilePath))
                {
                    while (!tr.EndOfStream)
                    {
                        tw.WriteLine(tr.ReadLine());
                    }
                    tr.Close();
                }
                tw.Close();
            }
        }

        // Create file with default filters for index words
        public static void GenerateDefaultFiltersFile()
        {
            StreamWriter SW;
            SW = File.CreateText(GetFiltersFilePath());
            SW.WriteLine(@"\S{1}-\S{1}");
            SW.WriteLine(@"[IVXCL]{2,8}[ème]*");
            SW.WriteLine(@"[IVXCL]{1,8}-[IVXCL]{1,8}");
            SW.WriteLine(@"[IVXCL]{1,8}e-[IVXCL]{1,8}e");
            SW.WriteLine(@"\S{1,12}er");
            SW.WriteLine(@"\S{1,12}ir");
            SW.WriteLine(@"\S{1,12}tion[s]*");
            SW.WriteLine(@"\S{1,12}sion[s]*");
            SW.Write(@"\S{1,12}ement");
            SW.Close();
        }
        // Create file with default filters for index
        public static void GenerateDefaultPatternsFile()
        {
            StreamWriter SW;
            SW = File.CreateText(GetPatternsFilePath());
            SW.WriteLine(@"Chiffres : Toute suite de chiffres arabes // [0-9]+");
            SW.WriteLine(@"Chiffres romains : Toute suite de chiffres romains // [IVXLCD]+");
            SW.WriteLine(@"Mots et chiffres : Toute suite de caractères alphanumériques // [\w]+");
            SW.WriteLine(@"Mots : Toute suite de caractères et de tirets sauf les chiffres // ([^0-9\W]|[-])+");
            SW.WriteLine(@"Mots : Tout mot en minuscule ou en majuscule contenant les lettres indiquées // [A-Za-zÀÁÂÃÄÅÆÇÈÉÊËÌÍÎÏÐÑÒÓÔÕÖØÙÚÛÜÝÞàáâãäåæçèéêëìíîïðñòóôõöøùúûüýþÿ-]+");
            SW.WriteLine(@"Mots grecs : Toute suite de caractères grecs // [Α-Ωα-ωʹ͵ͺ;΄΅Ά·ΈΉΊΌΎΏΐΪΫάέήίΰϊϋόύώϐϑϕϖϚϜϞϠϰἀἁἂἃἄἅἆἇἈἉἊἋἌἍἎἏἐἑἒἓἔἕἘἙἚἛἜἝἠἡἢἣἤἥἦἧἨἩἪἫἬἭἮἯἰἱἲἳἴἵἶἷἸἹἺἻἼἽἾἿὀὁὂὃὄὅὈὉὊὋὌὍὐὑὒὓὔὕὖὗὙὛὝὟὠὡὢὣὤὥὦὧὨὩὪὫὬὭὮὯὰάὲέὴήὶίὸόὺύὼώᾀᾁᾂᾃᾄᾅᾆᾇᾈᾉᾊᾋᾌᾍᾎᾏᾐᾑᾒᾓᾔᾕᾖᾗᾘᾙᾚᾛᾜᾝᾞᾟᾠᾡᾢᾣᾤᾥᾦᾧᾨᾩᾪᾫᾬᾭᾮᾯᾰᾱᾲᾳᾴᾶᾷᾸᾹᾺΆᾼ᾽ι᾿῀῁ῂῃῄῆῇῈΈῊΉῌ῍῎῏ῐῑῒΐῖῗῘῙῚΊ῝῞῟ῠῡῢΰῤῥῦῧῨῩῪΎῬ῭΅`ῲῳῴῶῷῸΌῺΏῼ´῾]+");
            SW.WriteLine(@"Mots en majuscules : Tout mot en majuscule contenant les lettres indiquées // [A-ZÀÁÂÃÄÅÆÇÈÉÊËÌÍÎÏÐÑÒÓÔÕÖØÙÚÛÜÝÞ]+");
            SW.WriteLine(@"Mots en minuscules : Tout mot en minuscule contenant les lettres indiqués // [a-zàáâãäåæçèéêëìíîïðñòóôõöøùúûüýþÿ-]+");
            SW.WriteLine( "Mots en ement : Toute suite de 1 à 12 caractères qui ne sont pas des espaces et qui se termine par \"ement\" // \\S{1,12}ement");
            SW.WriteLine( "Mots en tion : Toute suite de 1 à  12 caractères qui ne sont pas des espaces et qui se termine par \"tion\" et éventuellement qui se termine par \"tions\" // \\S{1,12}tion[s]*");
            SW.WriteLine(@"L-L : Tout caractère qui n'est pas un espace suivi d'un tiret et suivi d'un caractère qui n'est pas un espace // \\S{1}-\S{1}");
            SW.WriteLine( "Siècles : Tout chiffre romain de 2 à 8 caractères suivis éventuellement de 'e' ou de \"ème\" // [IVXLCD]{2,8}[ème]*");
            SW.WriteLine(@"Période : Tout chiffre romain de 1 à 8 caractères suivi d'un 'e' et d'un tiret et suivi d'un autre chiffre romain suivi d'un 'e' // [IVXLCD]{1,8}e-[IVXLCD]{1,8}e");
            SW.WriteLine(@"Initiales suspectes : Initiales de prénom suspectes par manque de caractères espace // Mc|([A-ZÀÁÂÃÄÅÆÇÈÉÊËÌÍÎÏÐÑÒÓÔÕÖØÙÚÛÜÝÞ][.]){2,3}");
            SW.WriteLine(@"Noms d'auteurs : Format Nnnnn I. I. (Majuscules ou minuscules) // [A-ZÀÁÂÃÄÅÆÇÈÉÊËÌÍÎÏÐÑÒÓÔÕÖØÙÚÛÜÝÞ][-A-ZÀÁÂÃÄÅÆÇÈÉÊËÌÍÎÏÐÑÒÓÔÕÖØÙÚÛÜÝÞa-zàáâãäåæçèéêëìíîïðñòóôõöøùúûüýþÿ ]*[,][ \u00A0](([A-ZÀÁÂÃÄÅÆÇÈÉÊËÌÍÎÏÐÑÒÓÔÕÖØÙÚÛÜÝÞ][.][- \u00A0][A-ZÀÁÂÃÄÅÆÇÈÉÊËÌÍÎÏÐÑÒÓÔÕÖØÙÚÛÜÝÞ][.][- \u00A0][A-ZÀÁÂÃÄÅÆÇÈÉÊËÌÍÎÏÐÑÒÓÔÕÖØÙÚÛÜÝÞ][.])|([A-ZÀÁÂÃÄÅÆÇÈÉÊËÌÍÎÏÐÑÒÓÔÕÖØÙÚÛÜÝÞ][.][- \u00A0][A-ZÀÁÂÃÄÅÆÇÈÉÊËÌÍÎÏÐÑÒÓÔÕÖØÙÚÛÜÝÞ][.])|([A-ZÀÁÂÃÄÅÆÇÈÉÊËÌÍÎÏÐÑÒÓÔÕÖØÙÚÛÜÝÞ][.]))");
            SW.WriteLine(@"Noms d'auteurs : Format NNNNN I. I. (Majuscules seulement) // [A-ZÀÁÂÃÄÅÆÇÈÉÊËÌÍÎÏÐÑÒÓÔÕÖØÙÚÛÜÝÞ][-A-ZÀÁÂÃÄÅÆÇÈÉÊËÌÍÎÏÐÑÒÓÔÕÖØÙÚÛÜÝÞ \u00A0]*[,][ \u00A0](([A-ZÀÁÂÃÄÅÆÇÈÉÊËÌÍÎÏÐÑÒÓÔÕÖØÙÚÛÜÝÞ][.][- \u00A0][A-ZÀÁÂÃÄÅÆÇÈÉÊËÌÍÎÏÐÑÒÓÔÕÖØÙÚÛÜÝÞ][.][- \u00A0][A-ZÀÁÂÃÄÅÆÇÈÉÊËÌÍÎÏÐÑÒÓÔÕÖØÙÚÛÜÝÞ][.])|([A-ZÀÁÂÃÄÅÆÇÈÉÊËÌÍÎÏÐÑÒÓÔÕÖØÙÚÛÜÝÞ][.][- \u00A0][A-ZÀÁÂÃÄÅÆÇÈÉÊËÌÍÎÏÐÑÒÓÔÕÖØÙÚÛÜÝÞ][.])|([A-ZÀÁÂÃÄÅÆÇÈÉÊËÌÍÎÏÐÑÒÓÔÕÖØÙÚÛÜÝÞ][.]))");
            SW.Close();
        }

        // Display a file with Notepad
        public static void DisplayFile(string filePath)
        {
            System.Diagnostics.Process.Start("notepad.exe", filePath);
        }

        // Display a web site page
        public static void BrowseUrl(string URL)
        {
            System.Diagnostics.Process.Start("explorer", URL);
        }

        // Create a file from a string list with no duplicated lines
        public static void UniqToFile(List<string> lines, string filePathOut)
        {
            using (StreamWriter w = new StreamWriter(new FileStream(filePathOut, FileMode.Create)))
            {
                foreach (string line in lines.Distinct())
                {
                    w.WriteLine(line);
                }
            }
        }

        // Create a file from a string list with no duplicated lines and sorted
        public static void SortedUniqToFile(List<string> lines, string filePathOut, string EmptyListMsg)
        {
            lines.Sort();
            using (StreamWriter w = new StreamWriter(new FileStream(filePathOut, FileMode.Create)))
            {
                if (lines.Count == 0)
                    w.WriteLine(EmptyListMsg);
                else
                {
                    foreach (string line in lines.Distinct())
                    {
                        w.WriteLine(line);
                    }
                }
            }
        }
        public static void UniqToScreen(List<string> lines)
        {
            foreach (string line in lines.Distinct())
            {
                Console.WriteLine(line);
            }
        }
        public static string GetFiltersFilePath()
        {
            return Path.Combine(TEMP_DIR, FILTERS_FILE);
        }
        public static List<string> GetFilters()
        {
            return GetLines(GetFiltersFilePath());
        }
        public static string GetPatternsFilePath()
        {
            return Path.Combine(TEMP_DIR, PATTERNS_FILE);
        }
        public static List<string> GetPatterns()
        {
            return GetLines(GetPatternsFilePath());
        }

        public static List<string> GetLines(string filePath)
        {
            List<string> lines = new List<string>();
            if (File.Exists(filePath))
            {
                using (StreamReader r = new StreamReader(new FileStream(filePath, FileMode.Open, FileAccess.Read)))
                {
                    string line;
                    while ((line = r.ReadLine()) != null)
                    {
                        lines.Add(line);
                    }
                }
            }
            else lines.Add(Msg.MSG_FILE_NOT_FOUND+ " " +filePath);
            return lines;
        }


        public static void SortUniq(string filePath, string resultFilePath)
        {
            if (File.Exists(filePath))
            {
                string filePathOut = resultFilePath;
                List<string> lines = Util.GetLines(filePath);
                lines.Sort();
                Util.UniqToFile(lines, filePathOut);
            }
        }

        public static void DisplayVersion()
        {
            MessageBox.Show(ThisAddIn.ADDIN_TITLE + " " +
                            Msg.MSG_VERSION + " " +
                            Assembly.GetExecutingAssembly().GetName().Version.ToString() + "\n" +
                            Msg.MSG_COPYRIGHT,
                            ThisAddIn.ADDIN_TITLE);
        }

        public static void NotYetImplemented()
        {
            MessageBox.Show(Msg.MSG_NOT_AVAILABLE, ThisAddIn.ADDIN_TITLE);
        }


    }
}

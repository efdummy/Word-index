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
        const string RESULT_FILE_NAME = "wordindex.Result";
        const string PROCESSED_OCCURRENCES_FILE_NAME = "wordindex.ProcessedOccurrences";
        const string RESULT_FILE_NAME_REJECTED = "wordindex.rejected";
        const string RESULT_FILE_EXT = ".txt";
        const string FILTERS_FILE = "wordindex.filterspatterns-fr.txt";
        const string PATTERNS_FILE = "wordindex.extractpatterns-fr.txt";

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

    /* Most used french verbs (can be used as filters)

    [a][b][a][n][d][o][n][n][e][a-z]*
    [a][c][c][e][p][t][e][a-z]*
    [a][c][c][o][m][p][a][g][n][e][a-z]*
    [a][c][c][o][r][d][e][a-z]*
    [a][c][h][e][t][e][a-z]*
    [a][c][h][e][v][e][a-z]*
    [a][d][r][e][s][s][e][a-z]*
    [a][g][i][t][e][a-z]*
    [a][i][d][e][a-z]*
    [a][i][m][e][a-z]*
    [a][j][o][u][t][e][a-z]*
    [a][l][l][u][m][e][a-z]*
    [a][m][e][n][e][a-z]*
    [a][m][u][s][e][a-z]*
    [a][n][n][o][n][c][e][a-z]*
    [a][p][p][a][r][t][e][n][i][r]
    [a][p][p][e][l][e][a-z]*
    [a][p][p][o][r][t][e][a-z]*
    [a][p][p][r][e][n][d][r][e]
    [a][p][p][r][o][c][h][e][a-z]*
    [a][p][p][u][y][e][a-z]*
    [a][r][r][a][c][h][e][a-z]*
    [a][r][r][ê][t][e][a-z]*
    [a][r][r][i][v][e][a-z]*
    [a][s][s][u][r][e][a-z]*
    [a][t][t][a][c][h][e][a-z]*
    [a][t][t][i][r][e][a-z]*
    [a][v][a][n][c][e][a-z]*
    [a][v][o][u][e][a-z]*
    [b][a][i][s][s][e][a-z]*
    [b][r][i][l][l][e][a-z]*
    [b][r][i][s][e][a-z]*
    [b][r][û][l][e][a-z]*
    [c][a][c][h][e][a-z]*
    [c][a][u][s][e][a-z]*
    [c][e][s][s][e][a-z]*
    [c][h][a][n][g][e][a-z]*
    [c][h][a][n][t][e][a-z]*
    [c][h][a][r][g][e][a-z]*
    [c][h][e][r][c][h][e][a-z]*
    [c][o][m][m][e][n][c][e][a-z]*
    [c][o][m][p][o][s][e][a-z]*
    [c][o][m][p][t][e][a-z]*
    [c][o][n][s][i][d][é][r][e][a-z]*
    [c][o][n][t][i][n][u][e][a-z]*
    [c][o][u][c][h][e][a-z]*
    [c][o][u][l][e][a-z]*
    [c][o][u][p][e][a-z]*
    [c][r][a][i][n][d][r][e]
    [c][r][é][e][a-z]*
    [c][r][i][e][a-z]*
    [d][a][n][s][e][a-z]*
    [d][é][c][i][d][e][a-z]*
    [d][é][c][l][a][r][e][a-z]*
    [d][e][m][a][n][d][e][a-z]*
    [d][e][m][e][u][r][e][a-z]*
    [d][é][s][i][r][e][a-z]*
    [d][e][v][i][n][e][a-z]*
    [d][i][r][i][g][e][a-z]*
    [d][i][s][t][i][n][g][u][e][a-z]*
    [d][o][n][n][e][a-z]*
    [d][o][u][t][e][a-z]*
    [d][r][e][s][s][e][a-z]*
    [d][u][r][e][a-z]*
    [é][c][h][a][p][p][e][a-z]*
    [é][c][l][a][i][r][e][a-z]*
    [é][c][l][a][t][e][a-z]*
    [é][c][o][u][t][e][a-z]*
    [é][l][e][v][e][a-z]*
    [e][m][b][r][a][s][s][e][a-z]*
    [e][m][m][e][n][e][a-z]*
    [e][m][p][ê][c][h][e][a-z]*
    [e][m][p][l][o][y][e][a-z]*
    [e][m][p][o][r][t][e][a-z]*
    [e][n][g][a][g][e][a-z]*
    [e][n][l][e][v][e][a-z]*
    [e][n][t][o][u][r][e][a-z]*
    [e][n][t][r][a][î][n][e][a-z]*
    [e][n][t][r][e][a-z]*
    [e][n][v][o][y][e][a-z]*
    [é][p][r][o][u][v][e][a-z]*
    [e][s][p][é][r][e][a-z]*
    [e][s][s][a][y][e][a-z]*
    [é][t][o][n][n][e][a-z]*
    [é][v][i][t][e][a-z]*
    [e][x][a][m][i][n][e][a-z]*
    [e][x][i][g][e][a-z]*
    [e][x][i][s][t][e][a-z]*
    [e][x][p][l][i][q][u][e][a-z]*
    [e][x][p][r][i][m][e][a-z]*
    [f][e][r][m][e][a-z]*
    [f][i][x][e][a-z]*
    [f][o][r][c][e][a-z]*
    [f][o][r][m][e][a-z]*
    [f][r][a][p][p][e][a-z]*
    [g][a][g][n][e][a-z]*
    [g][a][r][d][e][a-z]*
    [g][l][i][s][s][e][a-z]*
    [h][a][b][i][t][e][a-z]*
    [h][é][s][i][t][e][a-z]*
    [i][g][n][o][r][e][a-z]*
    [i][m][a][g][i][n][e][a-z]*
    [i][m][p][o][r][t][e][a-z]*
    [i][m][p][o][s][e][a-z]*
    [i][n][d][i][q][u][e][a-z]*
    [i][n][s][t][a][l][l][e][a-z]*
    [j][e][t][e][a-z]*
    [j][o][u][e][a-z]*
    [j][u][g][e][a-z]*
    [l][e][v][e][a-z]*
    [l][i][v][r][e][a-z]*
    [m][a][n][g][e][a-z]*
    [m][a][n][q][u][e][a-z]*
    [m][a][r][c][h][e][a-z]*
    [m][a][r][i][e][a-z]*
    [m][a][r][q][u][e][a-z]*
    [m][ê][l][e][a-z]*
    [m][e][n][e][a-z]*
    [m][o][n][t][e][a-z]*
    [m][o][n][t][r][e][a-z]*
    [n][o][m][m][e][a-z]*
    [o][b][l][i][g][e][a-z]*
    [o][b][s][e][a-z]*[v][e][a-z]*
    [o][c][c][u][p][e][a-z]*
    [o][s][e][a-z]*
    [o][u][b][l][i][e][a-z]*
    [p][a][r][l][e][a-z]*
    [p][a][r][v][e][n][i][r]
    [p][a][s][s][e][a-z]*
    [p][a][y][e][a-z]*
    [p][e][n][c][h][e][a-z]*
    [p][é][n][é][t][r][e][a-z]*
    [p][e][n][s][e][a-z]*
    [p][l][a][c][e][a-z]*
    [p][l][e][u][r][e][a-z]*
    [p][o][r][t][e][a-z]*
    [p][o][s][e][a-z]*
    [p][o][s][s][é][d][e][a-z]*
    [p][o][u][s][s][e][a-z]*
    [p][r][é][p][a][r][e][a-z]*
    [p][r][é][s][e][n][t][e][a-z]*
    [p][r][ê][t][e][a-z]*
    [p][r][i][e][a-z]*
    [p][r][o][m][e][n][e][a-z]*
    [p][r][o][n][o][n][c][e][a-z]*
    [p][r][o][p][o][s][e][a-z]*
    [p][r][o][u][v][e][a-z]*
    [q][u][i][t][t][e][a-z]*
    [r][a][c][o][n][t][e][a-z]*
    [r][a][p][p][e][l][e][a-z]*
    [r][e][f][u][s][e][a-z]*
    [r][e][g][a][r][d][e][a-z]*
    [r][e][l][e][v][e][a-z]*
    [r][e][m][a][r][q][u][e][a-z]*
    [r][e][m][o][n][t][e][a-z]*
    [r][e][n][c][o][n][t][r][e][a-z]*
    [r][e][n][t][r][e][a-z]*
    [r][é][p][é][t][e][a-z]*
    [r][e][p][o][s][e][a-z]*
    [r][e][p][r][é][s][e][n][t][e][a-z]*
    [r][e][s][p][i][r][e][a-z]*
    [r][e][s][s][e][m][b][l][e][a-z]*
    [r][e][s][t][e][a-z]*
    [r][e][t][i][r][e][a-z]*
    [r][e][t][o][u][r][n][e][a-z]*
    [r][e][t][r][o][u][v][e][a-z]*
    [r][é][v][e][i][l][l][e][a-z]*
    [r][é][v][e][a-z]*
    [r][o][u][l][e][a-z]*
    [s][a][l][u][e][a-z]*
    [s][a][u][v][e][a-z]*
    [s][e][m][b][l][e][a-z]*
    [s][é][p][a][r][e][a-z]*
    [s][e][r][r][e][a-z]*
    [s][o][n][g][e][a-z]*
    [s][o][n][n][e][a-z]*
    [s][o][u][f][f][l][e][a-z]*
    [s][o][u][l][e][v][e][a-z]*
    [t][e][n][t][e][a-z]*
    [t][e][r][m][i][n][e][a-z]*
    [t][i][r][e][a-z]*
    [t][o][m][b][e][a-z]*
    [t][o][u][c][h][e][a-z]*
    [t][o][u][r][n][e][a-z]*
    [t][r][a][î][n][e][a-z]*
    [t][r][a][i][t][e][a-z]*
    [t][r][a][v][a][i][l][l][e][a-z]*
    [t][r][a][v][e][r][s][e][a-z]*
    [t][r][e][m][b][l][e][a-z]*
    [t][r][o][m][p][e][a-z]*
    [t][r][o][u][v][e][a-z]*
    [t][u][e][a-z]*
    [v][o][l][e][a-z]* 
     
     * */


}

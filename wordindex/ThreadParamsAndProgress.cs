using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;


namespace Word2003Tools4Dominique
{
    //             Fenêtre d'attente créant les background thread et affichant la progression
    //
    //             Pour réutiliser :
    //             - Adapter la classe ThreadParameters ci-dessous
    //             - Ecrire le code du thread le squelette ci-dessous
    //               
    //               public static void DoSomething(object sender, DoWorkEventArgs e)
    //               {
    //                  BackgroundWorker worker = sender as BackgroundWorker;
    //                  ThreadParameters p = (ThreadParameters)e.Argument;
    //                  //... Prepare parameters
    //                  rc=DoSomethingFunction(...);
    //                  if (rc != 0)
    //                  {
    //                      e.Result = rc;
    //                      e.Cancel = true;
    //                  }
    //                  else
    //                  {
    //                  //... Finalize this thread work
    //                  }
    //                  e.Cancel = true;
    //                  worker.CancelAsync();
    //               }
    //
    //             - Ecrire le code de finalisation des travaux de tous les threads
    //
    //               public static void DoFinalizeAllThreadsWork(ThreadParameters p)
    //               {
    //                // Finalize all threads works here
    //               }
    //
    //             - Lancer les threads avec le code suivant
    //                
    //               // Open progress form
    //               ThreadParamsAndProgress progressForm = new ThreadParamsAndProgress();
    //               // Prepare parameters for all threads for your need
    //               ThreadParameters p41 = new ThreadParameters(...);
    //               ThreadParameters p42 = new ThreadParameters(...);
    //               ThreadParameters p43 = new ThreadParameters(...);
    //               ThreadParameters p44 = new ThreadParameters(...);
    //               // Start 4 threads
    //               progressForm.Start(Commands.DoExtractionForIndexFiltered, DoFinalize, p41, p42, p43, p44);
    //


    public partial class ThreadParamsAndProgress : Form
    {
        // Worker's array
        BackgroundWorker [] WorkersArray;
        const int MAX_WORKER = 4;
        string COMMON_MSG1;
        string COMMON_MSG2;
        // Nb threads
        int nbWorker;
        // Nb threads that have'nt complete
        int nbRemainingWorkers;
        // Lock object for UI updates for progress report
        static object locker = new object();

        string msgPrefix;

        public ThreadParamsAndProgress()
        {
            InitializeComponent();
            WorkersArray = new BackgroundWorker[MAX_WORKER];
            label1.Text = "Starting working with " + Environment.ProcessorCount + " CPU... ";
            Show();
            label1.Update();
            buttonCancel.Update();
            COMMON_MSG1 = " thread is working on " + Environment.ProcessorCount + " CPU... ";
            COMMON_MSG2 = " threads are working on " + Environment.ProcessorCount + " CPU... ";
        }

        private void ProgressForm_Load(object sender, EventArgs e)
        {
        }

        void AllocateBackgroundWorkerAndSetDelegate(DoWorkEventHandler DoWork, int ThreadNum)
        {
            WorkersArray[ThreadNum] = new BackgroundWorker();
            WorkersArray[ThreadNum].WorkerSupportsCancellation = true;
            WorkersArray[ThreadNum].WorkerReportsProgress = true;
            WorkersArray[ThreadNum].DoWork += new DoWorkEventHandler(DoWork);
            WorkersArray[ThreadNum].RunWorkerCompleted += new RunWorkerCompletedEventHandler(Completed);
            WorkersArray[ThreadNum].ProgressChanged += new ProgressChangedEventHandler(ProgressChanged);
        }

        // Delegate pour finaliser le travail de tous les threads
        public delegate void DoFinalizeThreadsWork(ThreadParameters p);
        DoFinalizeThreadsWork DoFinalizeCurrent;

        public void Start(DoWorkEventHandler DoWork, DoFinalizeThreadsWork DoFinalize, ThreadParameters p)
        {
            nbWorker = 1;
            nbRemainingWorkers = 1;
            AllocateBackgroundWorkerAndSetDelegate(DoWork, 0);
            DoFinalizeCurrent = DoFinalize;

            msgPrefix = nbWorker + COMMON_MSG1;
            label1.Text = msgPrefix;

            p.ThreadNum = 0;
            WorkersArray[0].RunWorkerAsync(p);
        }

        public void Start(DoWorkEventHandler DoWork, DoFinalizeThreadsWork DoFinalize, ThreadParameters Thread1p, ThreadParameters Thread2p)
        {
            nbWorker = 2;
            nbRemainingWorkers = 2;
            AllocateBackgroundWorkerAndSetDelegate(DoWork, 0);
            AllocateBackgroundWorkerAndSetDelegate(DoWork, 1);
            DoFinalizeCurrent = DoFinalize;

            msgPrefix = nbWorker + COMMON_MSG2;
            label1.Text = msgPrefix;

            Show();

            Thread1p.NumberOfWorkingThreads = nbWorker;
            Thread2p.NumberOfWorkingThreads = nbWorker;
            Thread1p.ThreadNum = 0;
            Thread2p.ThreadNum = 1;
            WorkersArray[0].RunWorkerAsync(Thread1p);
            WorkersArray[1].RunWorkerAsync(Thread2p);
        }
        public void Start(DoWorkEventHandler DoWork, DoFinalizeThreadsWork DoFinalize, ThreadParameters Thread1p, ThreadParameters Thread2p, ThreadParameters Thread3p)
        {
            nbWorker = 3;
            nbRemainingWorkers = 3;
            AllocateBackgroundWorkerAndSetDelegate(DoWork, 0);
            AllocateBackgroundWorkerAndSetDelegate(DoWork, 1);
            AllocateBackgroundWorkerAndSetDelegate(DoWork, 2);
            DoFinalizeCurrent = DoFinalize;

            msgPrefix = nbWorker + COMMON_MSG2;
            label1.Text = msgPrefix;

            Show();

            Thread1p.NumberOfWorkingThreads = nbWorker;
            Thread2p.NumberOfWorkingThreads = nbWorker;
            Thread3p.NumberOfWorkingThreads = nbWorker;
            Thread1p.ThreadNum = 0;
            Thread2p.ThreadNum = 1;
            Thread3p.ThreadNum = 2;
            WorkersArray[0].RunWorkerAsync(Thread1p);
            WorkersArray[1].RunWorkerAsync(Thread2p);
            WorkersArray[2].RunWorkerAsync(Thread3p);
        }
        public void Start(DoWorkEventHandler DoWork, DoFinalizeThreadsWork DoFinalize, ThreadParameters Thread1p, ThreadParameters Thread2p, ThreadParameters Thread3p, ThreadParameters Thread4p)
        {
            nbWorker = 4;
            nbRemainingWorkers = 4;
            AllocateBackgroundWorkerAndSetDelegate(DoWork, 0);
            AllocateBackgroundWorkerAndSetDelegate(DoWork, 1);
            AllocateBackgroundWorkerAndSetDelegate(DoWork, 2);
            AllocateBackgroundWorkerAndSetDelegate(DoWork, 3);
            DoFinalizeCurrent = DoFinalize;

            msgPrefix = nbWorker + COMMON_MSG2;
            label1.Text = msgPrefix;

            Show();

            Thread1p.NumberOfWorkingThreads = nbWorker;
            Thread2p.NumberOfWorkingThreads = nbWorker;
            Thread3p.NumberOfWorkingThreads = nbWorker;
            Thread4p.NumberOfWorkingThreads = nbWorker;
            Thread1p.ThreadNum = 0;
            Thread2p.ThreadNum = 1;
            Thread3p.ThreadNum = 2;
            Thread4p.ThreadNum = 3;
            WorkersArray[0].RunWorkerAsync(Thread1p);
            WorkersArray[1].RunWorkerAsync(Thread2p);
            WorkersArray[2].RunWorkerAsync(Thread3p);
            WorkersArray[3].RunWorkerAsync(Thread4p);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            for (int i=0; i<nbWorker; i++)
                WorkersArray[i].CancelAsync();
            Close();
        }

        //_______ Background worker thread methods
        void ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            lock (locker)
            {
                BackgroundWorker worker = sender as BackgroundWorker;
                if (!worker.CancellationPending)
                {
                    label1.Text = msgPrefix + ((int)(e.ProgressPercentage)).ToString() + "%   " + e.UserState.ToString();
                }
            }
        }
        void Completed(object sender, RunWorkerCompletedEventArgs e)
        {
            lock (locker)
            {
                nbRemainingWorkers--;
                if (nbRemainingWorkers == 0)
                {
                    label1.Text = msgPrefix + ("100" + "%   Finalizing reports...");
                    label1.Update();
                    if (DoFinalizeCurrent != null)
                        DoFinalizeCurrent(new ThreadParameters(nbWorker));
                    label1.Text = "Completed.";
                    Close();
                }
            }
        }
        // Edition des fichiers une fois que tous les threads sont terminés
        public static void DoNothingToDo(ThreadParameters ThreadParam)
        {

        }

        
    }

    // _______ This is the only class to adapt
    // Paramètres à passer au threads lors du lancement (start)
    public class ThreadParameters
    {
        public int ThreadNum;
        public int NumberOfWorkingThreads;
        public Microsoft.Office.Interop.Word.Application Appli;
        public string Text;
        public string Pattern;
        public int MinWordLength;
        public bool IgnoreCase;
        public bool OnOff;
        public List<string> StringList;

        public ThreadParameters(Microsoft.Office.Interop.Word.Application AppliWord, string TextToParse, string PatternToUse, int MinIndexWordLength, bool IgnoreCaseOption, bool boolOption, List<string> WordList)
        {
            Appli = AppliWord;
            Text = TextToParse;
            Pattern = PatternToUse;
            MinWordLength = MinIndexWordLength;
            IgnoreCase = IgnoreCaseOption;
            OnOff = boolOption;
            StringList = WordList;
        }
        public ThreadParameters(int NbWorker)
        {
            NumberOfWorkingThreads = NbWorker;
        }
    }

}

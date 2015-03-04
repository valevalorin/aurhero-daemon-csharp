using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Threading.Tasks;
using System.Timers;
using Npgsql;
using System.IO;
using System.Runtime.InteropServices;

namespace AurHero
{
    public partial class Service1 : ServiceBase
    {
        static Timer QueueTimer;
        static bool Processing;
        static NpgsqlConnection conn;
        static NpgsqlDataAdapter adapter;
        static NpgsqlCommand insertQueueLog;
        static NpgsqlCommand removeQueueTrack;
        const string GetQueue = "SELECT * from queue_tracks";
        static string RecordingCommand = "-c 2 -t waveaudio \"Line 1\" \"{0}.wav\" trim 0 00:{1}";
        static string Folder = "D:\\Music\\";

        public Service1()
        {
            InitializeComponent();
        }

        protected override void OnStart(string[] args)
        {
            Process(); 
        }

        protected override void OnStop()
        {
            adapter.Dispose();
            conn.Close();
        }

        public void Process()
        {
            try
            {
                //Connect To Database
                string connstring = String.Format("Server={0};Port={1};" +
                        "User Id={2};Password={3};Database={4};",
                        "127.0.0.1", 5432, "aurhero",
                        "traindriverchangegravity", "aurhero_dev");
                // Making connection with Npgsql provider
                conn = new NpgsqlConnection(connstring);
                conn.Open();
                adapter = new NpgsqlDataAdapter(GetQueue, conn);
                removeQueueTrack = new NpgsqlCommand("", conn);
                insertQueueLog = new NpgsqlCommand("", conn);

                //Start Queue Timer
                //Processing = false;
                //QueueTimer = new Timer(5000);
                //QueueTimer.Elapsed += ProcessQueue;
                //QueueTimer.Enabled = true;

                DataSet set = new DataSet();
                DataTable table = new DataTable();
                set.Reset();
                adapter.Fill(set);
                table = set.Tables[0];

                DataRow row;

                using (StreamWriter sw = new StreamWriter("C:\\service_log.txt"))
                {
                    row = table.AsEnumerable().First<DataRow>();
                    sw.WriteLine(row.ItemArray[0].ToString());
                }

                //0:ID - 1:Link - 2:Title - 3:Artist - 4:Album - 56 - 7:Duration
                string tempFolder = Folder + row.ItemArray[3];
                int duration = int.Parse(row.ItemArray[7].ToString())+15;

                //Check Dir
                if (!Directory.Exists(tempFolder))
                    Directory.CreateDirectory(tempFolder);


                if (!Directory.Exists(tempFolder + "\\" + row.ItemArray[4]) && !String.IsNullOrEmpty(row.ItemArray[4].ToString()))
                {
                    tempFolder = tempFolder + "\\" + row.ItemArray[4];
                    Directory.CreateDirectory(tempFolder);
                }

                //Start Recording
                //System.Diagnostics.Process recProc = RecordingProcess(tempFolder, row, duration);

                //Start Playing
                PlayingProcess(row);
                
                //Wait
                //System.Threading.Thread.Sleep((duration+2) * 1000);

                //Kill Firefox
                System.Threading.Thread.Sleep(3000);
                KillFirefox();

                //Stop Recording
                
                //recProc.Kill();

                //Clean Up

            }
            catch (Exception ex)
            {
                using (StreamWriter sw = new StreamWriter("C:\\service_log.txt"))
                {
                    sw.WriteLine(ex.Message);
                }

            }
        }

        private static void ProcessQueue(Object source, ElapsedEventArgs e)
        {
            if (!Processing)
            {
                Processing = true;

                DataSet set = new DataSet();
                DataTable table = new DataTable();
                set.Reset();
                adapter.Fill(set);
                table = set.Tables[0];

                DataRow row;

                using (StreamWriter sw = new StreamWriter("C:\\service_log.txt"))
                {
                    row = table.AsEnumerable().First<DataRow>();
                    sw.WriteLine(row.ItemArray[0].ToString());
                }

                //0:ID - 1:Link - 2:Title - 3:Artist - 4:Album
                

                Processing = false;
            }
        }

        private void InsertQueueLog(string link, string title, string artist, string album, int duration)
        {

        }

        private Process RecordingProcess(string tempFolder, DataRow row, int duration)
        {
            System.Diagnostics.ProcessStartInfo startInfo = new System.Diagnostics.ProcessStartInfo();
            //startInfo.WindowStyle = System.Diagnostics.ProcessWindowStyle.Hidden;
            startInfo.UseShellExecute = false;
            startInfo.FileName = @"D:\Program Files (x86)\sox-14-4-2\sox.exe";
            startInfo.Arguments = String.Format(RecordingCommand, tempFolder + "\\" + row.ItemArray[2], duration);

            return System.Diagnostics.Process.Start(startInfo);
        }

        private Process PlayingProcess(DataRow row)
        {
            System.Diagnostics.ProcessStartInfo playInfo = new System.Diagnostics.ProcessStartInfo();
            //startInfo.WindowStyle = System.Diagnostics.ProcessWindowStyle.Hidden;
            playInfo.UseShellExecute = false;
            playInfo.FileName = @"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe";
            playInfo.Arguments = row.ItemArray[1].ToString();

            return System.Diagnostics.Process.Start(playInfo); 
        }

        private void KillFirefox()
        {
            Process[] processlist = System.Diagnostics.Process.GetProcesses();

            foreach(Process proc in processlist){
                if(proc.ProcessName.Equals("firefox"))
                    proc.Close();
            }
        }
    }
}

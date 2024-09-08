using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Threading.Tasks;
using System.Data.SQLite;
using System.IO.Ports;
using System.Collections;
using System.Text.RegularExpressions;
using System.Threading;
using System.Globalization;

namespace InvTransService
{
    public partial class InvTransService : ServiceBase
    {
        private System.Threading.Timer _timer;
        private DateTime _targetTime;
        private Thread ReadThread;
        private Thread RunThread;
        public static Thread monitoringThread;
        private Thread ParseThread;
        public SerialPort _serialPort;
        public string Buffer = null;
        public static int RefreshFlag = 0;

        public InvTransService()
        {

            InitializeComponent();
            if (!EventLog.SourceExists("MySource"))
            {
                EventLog.CreateEventSource("MySource", "MyNewLog");
            }
            eventLog1.Source = "MySource";
            eventLog1.Log = "MyNewLog";



            //    OpeningPort();
           
            //    InitializeThreads();
        }

        protected override void OnStart(string[] args)
        {
          //  System.Diagnostics.Debugger.Launch();
            eventLog1.WriteEntry("MySimpleService started now.");
            _serialPort = new SerialPort("COM3"); // Replace with your COM port
            _serialPort.BaudRate = 9600;
            _serialPort.Parity = Parity.None;
            _serialPort.StopBits = StopBits.One;
            _serialPort.DataBits = 8;
            _serialPort.DataReceived += new SerialDataReceivedEventHandler(SerialPort_DataReceived);
            InitializeThreads();
            _targetTime = new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day, 16,51, 0); // Set target time to 2:00 PM
            _timer = new System.Threading.Timer(CheckTime, null, 0, 60000); // Check every minute
         


            //    RunThread.Start();
        }

        protected override void OnStop()
        {
            eventLog1.WriteEntry("MySimpleService stopped.");
        }
        private void CheckTime(object state)
        {
            if (DateTime.Now >= _targetTime && DateTime.Now < _targetTime.AddMinutes(1))
            {
                OpeningPort();
                eventLog1.WriteEntry("timer Click");
                RunTheProc();
            }

        }
        private void RunTheProc()
        {
            RunThread.Start();
            ReadThread.Start();
            
        }
        private void InitializeThreads()
        {
            // Initialize and start the COM port listening thread
            ReadThread = new Thread(ListenToComPort);
            ReadThread.IsBackground = true;
            

            // Initialize and start the Run thread
            RunThread = new Thread(RunningMethod);
            RunThread.IsBackground = true;
           



            //Initialize and start the parsing thraed
            ParseThread = new Thread(ParsingMethod);
            ParseThread.IsBackground = true;
            
        }
        public static string ConvertGregorianToSolar(DateTime gregorianDate)
        {
            PersianCalendar persianCalendar = new PersianCalendar();
            int year = persianCalendar.GetYear(gregorianDate);
            int month = persianCalendar.GetMonth(gregorianDate);
            int day = persianCalendar.GetDayOfMonth(gregorianDate);

            return $"{year}-{month:D2}-{day:D2}";
        }
        private void OpeningPort()
        {
            
           try
           {
                _serialPort.Open();
           }
           catch (Exception e)
           {
           }
            
        }

        private void SerialPort_DataReceived(object sender, SerialDataReceivedEventArgs e)
        {
            string data = _serialPort.ReadExisting();
            Buffer += data;
        }

        private void SendingMessage()
        {
            ReadSms();
            string QuestionEnergyLowWord = "MB30119=?";
            string QuestionEnergyHighWord = "MB30120=?";
            string connectionString = "Data Source=C:\\Users\\Koosha\\source\\repos\\GetNum\\GetNum\\bin\\Debug\\library.db";
            string GetSimNumQuery = "SELECT SimNum FROM DeviceInfoTable";
            List<string> columnData = new List<string>();
            using (SQLiteConnection connection = new SQLiteConnection(connectionString))
            {
                SQLiteCommand command = new SQLiteCommand(GetSimNumQuery, connection);
                connection.Open();
                using (SQLiteDataReader reader = command.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        columnData.Add(reader["SimNum"].ToString());
                    }
                }
            }
            string[] ListOfNumbers = columnData.ToArray();
            //opening the port for sending message
            
            //initializing the sms procedure

            _serialPort.WriteLine("AT+CMGF=1\r"); // Set SMS text mode
            Thread.Sleep(400);
            _serialPort.WriteLine("AT+CSCS=\"GSM\"" + '\r');
            Thread.Sleep(400);
            //sending batch messages
            Thread.Sleep(2000);
            ParseThread.Start();
            eventLog1.WriteEntry("parsing thread start and sending message begin");
            foreach (var item in ListOfNumbers)
            {
                if (_serialPort.IsOpen)
                {
                    _serialPort.WriteLine("AT+CMGS=" + "\"" + item + "\"" + '\r');
                    Thread.Sleep(400);
                    _serialPort.WriteLine(QuestionEnergyHighWord + (char)26 + '\r');
                    Thread.Sleep(400);
                }
                else
                {
                    try
                    {
                        _serialPort.Open();
                    }
                    catch (Exception e)
                    {
                        eventLog1.WriteEntry("this happened "+ e);
                    }
                }
                Thread.Sleep(5000);
            }
            Thread.Sleep(2000);
            foreach (var item in ListOfNumbers)
            {
                if (_serialPort.IsOpen)
                {
                    Thread.Sleep(400);
                    _serialPort.WriteLine("AT+CMGS=" + "\"" + item + "\"" + '\r');
                    Thread.Sleep(300);
                    _serialPort.WriteLine(QuestionEnergyLowWord + (char)26 + '\r');
                    Thread.Sleep(300);

                }
                else
                {
                    try
                    {
                        _serialPort.Open();
                    }
                    catch (Exception)
                    {
                    }
                }
                Thread.Sleep(5000);

            }
            eventLog1.WriteEntry("sending message ended");
        }

        private void ReadSms()
        {
            try
            {
                _serialPort.WriteLine("AT+CMGD=0,4\r"); // delete pre sms
                Thread.Sleep(400);
                _serialPort.WriteLine("AT+CMGF=1\r"); // Set SMS text mode
                Thread.Sleep(400);

                _serialPort.WriteLine("AT+CPMS=\"SM\"\r"); // Select SIM storage
                Thread.Sleep(400);

                _serialPort.WriteLine("AT+CNMI=2,2,0,0,0\r");
                Thread.Sleep(400);

                // Read a specific message (example: message at index 1)
                // _serialPort.WriteLine("AT+CMGR=1\r");
                Thread.Sleep(500);
            }
            catch (Exception e)
            {
                eventLog1.WriteEntry("this is happendddd : "+ e);
            }

        }

        private void ParsingMethod()
        {
            eventLog1.WriteEntry("parsing method start waiting");
            int WaitTime = 120000;
            string[,] DataInDB;
            Thread.Sleep(WaitTime);
            eventLog1.WriteEntry("parsing method start Working");
            string[] Messages = ExtractCMTMessages(Buffer);
            DataInDB = CreateArray(Messages);
            

            InsertInDataBase(DataInDB);

            while (true)
            {
                // Keep the threading alive
                Thread.Sleep(200);
            }

        }
       
        private void ChekingError(string[,] DataInDB, int[] TodayEnergy)
         {
             string connectionString = "Data Source=C:\\Users\\Koosha\\source\\repos\\GetNum\\GetNum\\bin\\Debug\\library.db";
             string GetSimNumQuery = "SELECT SimNum , SID FROM DeviceInfoTable";
             List<string> AllSimList = new List<string>();
             List<string> SidList = new List<string>();
             List<string> ResSimList = new List<string>();
             using (SQLiteConnection connection = new SQLiteConnection(connectionString))
             {
                 SQLiteCommand command = new SQLiteCommand(GetSimNumQuery, connection);
                 connection.Open();

                 using (SQLiteDataReader reader = command.ExecuteReader())
                 {
                     while (reader.Read())
                     {
                         AllSimList.Add(reader["SimNum"].ToString());
                         SidList.Add(reader["sid"].ToString());
                     }
                 }
             }
             string[,] StatArr = new string[SidList.Count, 3];

             for (int i = 0; i < DataInDB.GetLength(0); i++)
             {
                 ResSimList.Add(DataInDB[i, 0]);
             }
             for (int i = 0; i < SidList.Count; i++)
             {
                 StatArr[i, 0] = SidList[i];
             }

             for (int i = 0; i < AllSimList.Count; i++)
             {
                 StatArr[i, 1] = AllSimList[i];
             }
             for (int i = 0; i < AllSimList.Count; i++)
             {
                 int index = ResSimList.IndexOf(AllSimList[i]);
                 if (index != -1)// contians the number
                 {

                     if (TodayEnergy[index] != -1 && TodayEnergy[index] < 10)
                     {
                         StatArr[i, 2] = "Warning";
                     }
                     else if (TodayEnergy[index] != -1 && TodayEnergy[index] > 10)
                     {
                         StatArr[i, 2] = "OK";
                     }
                     else if (TodayEnergy[index] == -1)
                     {
                         StatArr[i, 2] = "Fail";
                     }
                 }
                 else // this number not responses
                 {
                     StatArr[i, 2] = "Fail";
                 }
             }

             using (SQLiteConnection connection = new SQLiteConnection(connectionString))
             {

                 connection.Open();
                 string InsStatusQuery;
                 DateTime currentDate = DateTime.Today;
                 string today = ConvertGregorianToSolar(currentDate);
                //  string today = currentDate.ToString("yyyy-MM-dd");
                for (int i = 0; i < StatArr.GetLength(0); i++)
                 {
                     if (StatArr[i, 2] == "OK")
                     {
                         InsStatusQuery = " update DeviceTodayFeed set status =@status where (sid=@OKsid And date=@date );";
                         using (SQLiteCommand command = new SQLiteCommand(InsStatusQuery, connection))
                         {
                             command.Parameters.AddWithValue("@OKsid", StatArr[i, 0]);
                             command.Parameters.AddWithValue("@status", StatArr[i, 2]);
                             command.Parameters.AddWithValue("@date", today);
                             command.ExecuteNonQuery();
                         }
                     }
                     if (StatArr[i, 2] == "Warning")
                     {
                         InsStatusQuery = " update DeviceTodayFeed set status =@status where (sid=@Warningsid And date=@date);";
                         using (SQLiteCommand command = new SQLiteCommand(InsStatusQuery, connection))
                         {
                             command.Parameters.AddWithValue("@Warningsid", StatArr[i, 0]);
                             command.Parameters.AddWithValue("@status", StatArr[i, 2]);
                             command.Parameters.AddWithValue("@date", today);
                             command.ExecuteNonQuery();
                         }
                     }
                     if (StatArr[i, 2] == "Fail")
                     {
                         InsStatusQuery = "INSERT INTO DeviceTodayFeed (sid, energy, date, status) VALUES (@sid, 0 , @date, @status);";
                         using (SQLiteCommand command = new SQLiteCommand(InsStatusQuery, connection))
                         {
                             command.Parameters.AddWithValue("@sid", StatArr[i, 0]);
                             command.Parameters.AddWithValue("@date", today);
                             command.Parameters.AddWithValue("@status", StatArr[i, 2]);
                             command.ExecuteNonQuery();
                         }
                     }

                 }


             }

         }

        private void InsertInDataBase(string[,] DataInDB)
         {

             int[] YesterdayEnergy = new int[DataInDB.GetLength(0)];
             int[] TodayEnergy = new int[DataInDB.GetLength(0)];

             string simnum;
             string connectionString = "Data Source=C:\\Users\\Koosha\\source\\repos\\GetNum\\GetNum\\bin\\Debug\\library.db";
             using (SQLiteConnection connection = new SQLiteConnection(connectionString))
             {
                 try
                 {
                     connection.Open();
                     // Create a table if it doesn't exist
                     string createTableQuery = "CREATE TABLE IF NOT EXISTS DeviceFeedLog (SID INTEGER,Energy Integer,Date TEXT)";
                     using (SQLiteCommand createTableCmd = new SQLiteCommand(createTableQuery, connection))
                     {
                         createTableCmd.ExecuteNonQuery();
                     }

                     // Insert data from TextBox
                     DateTime currentDate = DateTime.Today;
                     string today = ConvertGregorianToSolar( currentDate);

                     for (int i = 0; i < DataInDB.GetLength(0); i++)
                     {
                         if (DataInDB[i, 1] != "-1")
                         {
                             simnum = DataInDB[i, 0];

                             string selectSidQuery = $"SELECT sid FROM deviceinfotable WHERE simnum = '{simnum}'";
                             using (var command = new SQLiteCommand(selectSidQuery, connection))
                             {
                                 int sid = Convert.ToInt32(command.ExecuteScalar());

                                 // Insert new row into devicefeedlog
                                 string insertDataQuery = "INSERT INTO devicefeedlog (sid, energy, date) VALUES (@sid, @energy, @date)";
                                 using (var insertCommand = new SQLiteCommand(insertDataQuery, connection))
                                 {
                                     insertCommand.Parameters.AddWithValue("@sid", sid);
                                     insertCommand.Parameters.AddWithValue("@energy", DataInDB[i, 1]);
                                     insertCommand.Parameters.AddWithValue("@date", today);
                                     insertCommand.ExecuteNonQuery();
                                 }
                             }
                         }
                     }




                     // Create a table2 if it doesn't exist
                     string createTable2Query = "CREATE TABLE IF NOT EXISTS DeviceTodayFeed (SID INTEGER,Energy Integer,Date TEXT, Status TEXT)";
                     using (SQLiteCommand createTable2Cmd = new SQLiteCommand(createTable2Query, connection))
                     {
                         createTable2Cmd.ExecuteNonQuery();
                     }

                     
                     currentDate = DateTime.Today;
                     today = ConvertGregorianToSolar(currentDate);
                    string yesterday = ConvertGregorianToSolar(currentDate.AddDays(-1));

                     for (int i = 0; i < DataInDB.GetLength(0); i++)
                     {
                         simnum = DataInDB[i, 0];
                         string selectSidQuery = $"SELECT sid FROM deviceinfotable WHERE simnum = '{simnum}'";
                         using (var command = new SQLiteCommand(selectSidQuery, connection))
                         {
                             int sid = Convert.ToInt32(command.ExecuteScalar());

                             // Insert new row into devicefeedlog
                             string SelectyesterdayEnergy = $"SELECT energy FROM devicefeedlog WHERE(sid={sid} AND date='{yesterday}')";
                             using (var selectyesterdayenergyCMD = new SQLiteCommand(SelectyesterdayEnergy, connection))
                             {
                                 var yesenergy = selectyesterdayenergyCMD.ExecuteScalar();
                                 YesterdayEnergy[i] = yesenergy != null ? Convert.ToInt32(yesenergy) : 0;
                                 selectyesterdayenergyCMD.ExecuteNonQuery();
                             }
                         }
                     }
                     for (int i = 0; i < DataInDB.GetLength(0); i++)
                     {
                         TodayEnergy[i] = Convert.ToInt32(DataInDB[i, 1]) - YesterdayEnergy[i];
                         if (DataInDB[i, 1] == "-1")
                         {
                             TodayEnergy[i] = -1;
                         }
                     }





                     for (int i = 0; i < DataInDB.GetLength(0); i++)
                     {
                         if (DataInDB[i, 1] != "-1")
                         {
                             simnum = DataInDB[i, 0];
                             string selectSidQuery = $"SELECT sid FROM deviceinfotable WHERE simnum = '{simnum}'";
                             using (var command = new SQLiteCommand(selectSidQuery, connection))
                             {
                                 int sid = Convert.ToInt32(command.ExecuteScalar());
                                 // Insert new row into devicefeedlog
                                 string insertDataQuery = "INSERT INTO DeviceTodayFeed (sid, energy, date) VALUES (@sid, @energy, @date)";
                                 using (var insertCommand = new SQLiteCommand(insertDataQuery, connection))
                                 {
                                     insertCommand.Parameters.AddWithValue("@sid", sid);
                                     insertCommand.Parameters.AddWithValue("@energy", TodayEnergy[i]);
                                     insertCommand.Parameters.AddWithValue("@date", today);
                                     insertCommand.ExecuteNonQuery();
                                 }
                             }
                         }


                     }
                 }
                 catch (Exception )
                 {
                 }
             }

             ChekingError(DataInDB, TodayEnergy);
             
        }

        private string[,] CreateArray(string[] Messages)
         {
             string[,] resultArray = new string[Messages.Length, 2];

             for (int t = 0; t < Messages.Length; t++)
             {
                 string input = Messages[t];

                 // Split the input string into lines
                 string[] lines = input.Split(new[] { '\n' }, StringSplitOptions.None);

                 for (int i = 0; i < lines.Length - 1; i++)
                 {
                     // Regex to find 13-digit substrings starting with +98
                     Match match = Regex.Match(lines[i], @"\+98\d{10}");

                     if (match.Success)
                     {
                         resultArray[t, 0] = match.Value;
                         resultArray[t, 1] = lines[i + 1];
                         break; // Stop after finding the first match
                     }
                 }
             }

             HashSet<string> uniqueSims = new HashSet<string>();
             for (int j = 0; j < Messages.Length; j++)
             {
                 uniqueSims.Add(resultArray[j, 0]);
             }
             int sims = uniqueSims.Count;
             string[] arrayOfSims = uniqueSims.ToArray();
             string[,] finalArray = new string[sims, 4]; // [0-simnumbers][1-highword][2-lowword][3-energy]
             string[,] returnArray = new string[sims, 2];
             for (int k = 0; k < sims; k++)
             {
                 finalArray[k, 0] = arrayOfSims[k];
             }
             for (int j = 0; j < Messages.Length; j++)
             {
                 for (int i = 0; i < sims; i++)
                 {
                     if (finalArray[i, 0] == resultArray[j, 0] && resultArray[j, 1].Contains("MB30119="))
                     {
                         finalArray[i, 1] = resultArray[j, 1].Substring(8);
                     }
                     if (finalArray[i, 0] == resultArray[j, 0] && resultArray[j, 1].Contains("MB30120="))
                     {
                         finalArray[i, 2] = resultArray[j, 1].Substring(8);
                     }
                 }
             }

             for (int i = 0; i < sims; i++)
             {
               
                try
                 {
                     int energy = (int.Parse(finalArray[i, 2]) << 16) + int.Parse(finalArray[i, 1]);
                     finalArray[i, 3] = energy.ToString();
                 }
                 catch (ArgumentNullException) // handleing the one packet of energy lost
                 {
                     int energy = -1;
                     finalArray[i, 3] = energy.ToString();
                 }
                 
             }
             for (int i = 0; i < sims; i++)
             {
                 returnArray[i, 0] = finalArray[i, 0];
                 returnArray[i, 1] = finalArray[i, 3];
             }

             return returnArray;
         }

        public static string[] ExtractCMTMessages(string input)
         {
             List<string> messages = new List<string>();
             string[] lines = input.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);

             for (int i = 0; i < lines.Length; i++)
             {
                 if (lines[i].StartsWith("+CMT:"))
                 {
                     string message = lines[i];
                     if (i + 1 < lines.Length) message += "\n" + lines[i + 1];
                     messages.Add(message);
                 }
             }

             return messages.ToArray();
         }

        private void ListenToComPort()
        {
             _serialPort.DataReceived += SerialPort_DataReceived;

             while (true)
             {
                 // Keep the thread alive
                 Thread.Sleep(200);
             }
        }
        
         public void RunningMethod()
         {
             try
             {
                 SendingMessage();
                 while (true)
                 {
                     // Keep the thread alive
                     Thread.Sleep(200);
                 }
             }
             catch (ThreadAbortException)
             {
                 Console.WriteLine("ThreadAbortException caught. Cleaning up...");
             }
             finally
             {
                 RefreshFlag = 1;
             }

         }
         

    }
}

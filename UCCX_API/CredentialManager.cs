using System;
using System.IO;
using Microsoft.Extensions.Configuration;

namespace UCCX_API
{
    class CredentialManager : APIHandler
    {
        public string Env { get; set; }
        public string RootURL { get; set; }
        public string ExcelFile { get; set; }
        public string Username { get; set; }
        public string Password { get; set; }
        public IConfigurationRoot Configuration { get; set; }
        public IConfigurationRoot ConfigurationEnv { get; set; }
        public string LogPath { get; set; }
        public CredentialManager()
        {
            SetEnv();
            SetConfig();
            UpdateConsoleStep("Fetching Config Values...");
            SetRootURL();
            SetFilePath();
            SetCredentials();
            SetLogPath();
            using (StreamWriter w = File.AppendText(LogPath))
            {
                string initLine = "#### Logging Initiated -- " + System.DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss") + " ####";
                int len = initLine.Length;
                string borders = "";
                for(int i = 0; i < len; ++i)
                {
                    borders += "#";
                }
                w.WriteLine(borders);
                w.WriteLine(initLine);
                w.WriteLine(borders);
            }
            BeginLog();
            LogMessage($"Current Environment: {Env}");
            LogMessage($"Current Root URL: {RootURL}");
            LogMessage($"Using Username: {Username}");
            LogMessage($"Using Password: {Password.Substring(0, Password.Length / 5)}");
            LogMessage($"Current Excel File: {ExcelFile}");
            EndLog();
        }
        private void SetConfig()
        {
            UpdateConsoleStep("Initializing Config Parameters...");
            //Required for dotnet run --project <PATH> command to be used to execute the process via batch file
            string workingDirectory = Environment.CurrentDirectory;
            //Console.WriteLine("\n\nWORKING DIRECTORY: " + workingDirectory);
            string projectDirectory = Directory.GetParent(workingDirectory).Parent.FullName;
            //Console.WriteLine(projectDirectory);
            string jsonPath = workingDirectory + "\\appsettings.json";


            //Build Base Config (Global Settings)
            var builder = new ConfigurationBuilder()
                .SetBasePath(Directory.GetCurrentDirectory())
                .AddJsonFile(jsonPath, optional: true, reloadOnChange: true);

            IConfigurationRoot configuration = builder.Build();

            Configuration = configuration;

            //Build Environmental Config
            if(Env == "PROD")
            {
                jsonPath = workingDirectory + "\\appsettings.production.json";
                builder = new ConfigurationBuilder()
                    .SetBasePath(Directory.GetCurrentDirectory())
                    .AddJsonFile(jsonPath, optional: true, reloadOnChange: true);

                IConfigurationRoot configurationEnv = builder.Build();

                ConfigurationEnv = configurationEnv;
            }
            else
            {
                jsonPath = workingDirectory + "\\appsettings.development.json";
                builder = new ConfigurationBuilder()
                    .SetBasePath(Directory.GetCurrentDirectory())
                    .AddJsonFile(jsonPath, optional: true, reloadOnChange: true);

                IConfigurationRoot configurationEnv = builder.Build();

                ConfigurationEnv = configurationEnv;
            }
        }
        private void SetEnv()
        {
            UpdateConsoleStep("Initializing Environment Parameters...");
            string env = "DEV";
            switch (Environment.MachineName.ToUpper())
            {
                case "VAL-H7T4SQ2":
                    env = "DEV";
                    break;
                case "VAL-61LJXT2":
                    env = "DEV";
                    break;
                case "VAVPC-ROBO-02":
                    env = "DEV";
                    break;
                case "VAVPC-ROBO-05":
                    env = "DEV";
                    break;
                case "VAVPC-ROBO-07":
                    env = "DEV";
                    break;
                case "VAVPC-ROBO-01":
                    env = "PROD";
                    break;
                case "VAVPC-ROBO-03":
                    env = "PROD";
                    break;
                case "VAVPC-ROBO-04":
                    env = "PROD";
                    break;
                case "VAVPC-ROBO-06":
                    env = "PROD";
                    break;
            }
            Env = env;
        }
        private void SetRootURL()
        {
            RootURL = ConfigurationEnv.GetSection("UCCX_URL").Value;
        }
        private void SetFilePath()
        {
            ExcelFile = ConfigurationEnv.GetSection("ExcelFile").Value;
        }
        private void SetCredentials()
        {
            Username = Configuration.GetSection("UCCXCredentials")["Username"];
            Password = Configuration.GetSection("UCCXCredentials")["Password"];
        }
        private void SetLogPath()
        {
            UpdateConsoleStep("Initializing Log File...");
            LogPath = ConfigurationEnv.GetSection("Logging").Value + "WFM_UpdateQueue_" + System.DateTime.Now.ToString("MMddyyyy_hh-mm-ss") + ".txt";
        }
        public void BeginLog()
        {
            using (StreamWriter w = File.AppendText(LogPath))
            {
                w.Write("\r\nLog Entry : ");
                w.WriteLine($"{DateTime.Now.ToLongTimeString()} {DateTime.Now.ToLongDateString()}");
            }
        }
        public void LogMessage(string message)
        {
            using (StreamWriter w = File.AppendText(LogPath))
            {
                w.WriteLine($"  :{message}");
            }
        }
        public void EndLog()
        {
            using (StreamWriter w = File.AppendText(LogPath))
            {
                w.WriteLine("----------------------------------------------------------------------------------------------");
            }
        }
        public new void Info()
        {
            Console.WriteLine("ENV: {0}\nROOT URL: {1}\nEXCEL FILE: {2}\nUSERNAME: {3}\nPASSWORD: {4}", Env, RootURL, ExcelFile, Username, Password.Substring(0, 10));
        }
    }
}

using Microsoft.Extensions.Configuration;
using System.IO;
using System;
using System.Collections.Generic;
using System.Text;
using MailKit;
using MimeKit;
using MailKit.Net.Smtp;
using System.Linq;

namespace UCCX_API
{
    class EmailHandler
    {
        public IConfigurationRoot Configuration { get; set; }
        public EmailClasses.IEmailConfiguration EmailConfig { get; set; }
        public string EmailTo { get; set; }
        public string EmailCC { get; set; }
        public string EmailFrom { get; set; }
        public string EmailBody { get; set; }
        public string EmailSubject { get; set; }

        public void BuildEmail(bool fatalError, string message, string logPath, string runSummary, bool updateErrors)
        {
            Console.WriteLine("\n\n");
            //Required for dotnet run --project <PATH> command to be used to execute the process via batch file
            string workingDirectory = Environment.CurrentDirectory;
            //Console.WriteLine("\n\nWORKING DIRECTORY: " + workingDirectory);
            string jsonPath = workingDirectory + "\\appsettings.json";
            //Build Base Config (Global Settings)
            var builder = new ConfigurationBuilder()
                .SetBasePath(Directory.GetCurrentDirectory())
                .AddJsonFile(jsonPath, optional: true, reloadOnChange: true);

            IConfigurationRoot configuration = builder.Build();

            Configuration = configuration;

            EmailClasses.IEmailConfiguration emailConfig = new EmailClasses.EmailConfiguration
            {
                SmtpServer = configuration.GetSection("EmailConfiguration")["SmtpServer"],
                SmtpPort = Convert.ToInt32(configuration.GetSection("EmailConfiguration")["SmtpPort"]),
                SmtpUsername = configuration.GetSection("EmailConfiguration")["SmtpUsername"],
                SmtpPassword = Base64Decode(configuration.GetSection("EmailConfiguration")["SmtpPassword"])
            };
            EmailConfig = emailConfig;
            string mach = Environment.MachineName.ToString().ToUpper();
            if (mach == "VAVPC-ROBO-01" || mach == "VAVPC-ROBO-03" || mach == "VAVPC-ROBO-04" || mach == "VAVPC-ROBO-06")
            {
                EmailTo = configuration.GetSection("EmailConfiguration")["EmailTo"];
                EmailCC = configuration.GetSection("EmailConfiguration")["EmailCC"];
            }
            else
            {
                EmailTo = configuration.GetSection("EmailConfiguration")["EmailCC"];
            }
            EmailFrom = configuration.GetSection("EmailConfiguration")["EmailFrom"];
            if(fatalError == true)
            {
                EmailBody = configuration.GetSection("EmailConfiguration")["FatalErrorEmailBody"];
                EmailSubject = configuration.GetSection("EmailConfiguration")["FatalErrorEmailSubject"].Replace("<DATE>", System.DateTime.Now.ToString("MM/dd/yyyy"));
            }
            else if(updateErrors == true)
            {
                EmailBody = configuration.GetSection("EmailConfiguration")["EmailBody"].Replace("<REPORTING_SUMMARY>", runSummary).Replace("<REPORTING_RESULTS>", message).Replace("<FULL_LOG_PATH>", logPath);
                EmailSubject = configuration.GetSection("EmailConfiguration")["UpdateErrorEmailSubject"].Replace("<DATE>", System.DateTime.Now.ToString("MM/dd/yyyy"));
            }
            else
            {
                EmailBody = configuration.GetSection("EmailConfiguration")["EmailBody"].Replace("<REPORTING_SUMMARY>", runSummary).Replace("<REPORTING_RESULTS>", message).Replace("<FULL_LOG_PATH>", logPath);
                EmailSubject = configuration.GetSection("EmailConfiguration")["EmailSubject"].Replace("<DATE>", System.DateTime.Now.ToString("MM/dd/yyyy"));
            }

            SendEmail();
        }
        public static string Base64Decode(string base64EncodedData)
        {
            var base64EncodedBytes = System.Convert.FromBase64String(base64EncodedData);
            return System.Text.Encoding.UTF8.GetString(base64EncodedBytes);
        }
        public void SendEmail()
        {
            
            var message = new MimeMessage();
            
            // Add Email From
            message.From.Add(new MailboxAddress(EmailFrom));
            
            // Add Email To
            if(EmailTo != null && EmailTo != String.Empty)
            {
                if (EmailTo.Contains(";"))
                {
                    List<string> toList = new List<string>();
                    toList = EmailTo.Split(';').ToList();
                    foreach(string email in toList)
                    {
                        message.To.Add(new MailboxAddress(email));
                    }
                }
                else
                {
                    message.To.Add(new MailboxAddress(EmailTo));
                }
            }
            // Add Email CC
            if(EmailCC != null && EmailCC != String.Empty)
            {
                if (EmailCC.Contains(";"))
                {
                    List<string> toList = new List<string>();
                    toList = EmailCC.Split(';').ToList();
                    foreach (string email in toList)
                    {
                        message.Cc.Add(new MailboxAddress(email));
                    }
                }
                else
                {
                    message.Cc.Add(new MailboxAddress(EmailCC));
                }
            }
            
            // Add Subject, currently just an error subject
            message.Subject = EmailSubject;

            // Add Body, currently just an error body
            message.Body = new TextPart("plain") { Text = EmailBody };

            using (var client = new SmtpClient(new ProtocolLogger("smtp.log")))
            {

                int count = 0;
                int maxTries = 1000;
                bool done = false;
                while (done == false)
                {
                    try
                    {
                        UpdateConsoleStep("Attempting to Connect to SMTP Server...");
                        client.Disconnect(true);
                        client.Timeout = 15000;
                        client.ServerCertificateValidationCallback = (s, c, h, e) => true;
                        client.CheckCertificateRevocation = false;
                        
                        //client.SslProtocols = System.Security.Authentication.SslProtocols.Tls;
                        client.Connect(EmailConfig.SmtpServer, EmailConfig.SmtpPort, MailKit.Security.SecureSocketOptions.Auto);
                        done = true;
                    }
                    catch (Exception e)
                    {
                        UpdateConsoleStep(e.Message);
                        client.Disconnect(true);
                        if(e.Message.Contains("existing connection was forcibly closed"))
                        {
                            if(++count == maxTries)
                            {
                                return;
                            }
                            else
                            {
                                System.Threading.Thread.Sleep(500);
                            }
                        }
                    }
                }

                UpdateConsoleStep("Authenticating User to SMTP Server...");
                client.Authenticate("UiAutoBot@elephant.com", EmailConfig.SmtpPassword);

                UpdateConsoleStep("Sending Email...");
                client.Send(message);

                UpdateConsoleStep("Disconnecting from SMTP Server...");
                client.Disconnect(true);
            }
        }
        public void UpdateConsoleStep(string message)
        {
            int currentLineCursor = Console.CursorTop;
            Console.SetCursorPosition(0, Console.CursorTop);
            Console.Write(new string(' ', Console.WindowWidth));
            Console.SetCursorPosition(0, currentLineCursor);
            Console.Write(message);
        }
    }
}

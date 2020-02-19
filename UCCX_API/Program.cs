using System;
using System.Threading;

namespace UCCX_API
{
    class Program
    {
        static void Main(string[] args)
        {
            EmailHandler emailServices = new EmailHandler();
            try
            {
                //throw new System.Exception("TEST ERROR");
                // Core data is stored within APIHandler
                APIHandler apiHandler = new APIHandler();
                // Creates CredentialManager and APIData objects
                // Credential Manager reads config Parameters based on Environment
                // APIData uses Environment Data to determine where to populate API Data Resources and Skill
                try
                {
                    apiHandler.Init();
                }
                catch (Exception e)
                {
                    Console.Clear();
                    Console.WriteLine("########## ERROR OCCURRED DURING INIT ##########");
                    Thread.Sleep(30000);
                    apiHandler.Init();
                    apiHandler.cm.LogMessage($"Error occurred during Init: {e.Message.ToString()}");
                    apiHandler.cm.LogMessage($"{e.StackTrace.ToString()}");
                }
                // Prints all Data pulled from the API
                //apiHandler.apiData.Info();



                // ExcelData is comprised of the updated info pulled from WFM Excel Sheet
                    // Reads Agents and their desired Queue
                    // Reads Queues and the required skills to be added to the Queue
                ExcelData excelData = new ExcelData();
                try
                {
                    excelData.Init(apiHandler.cm);
                }
                catch (Exception e)
                {
                    apiHandler.cm.BeginLog();
                    apiHandler.cm.LogMessage($"Error occurred during ExcelData Init Sequence: {e.Message.ToString()}");
                    apiHandler.cm.LogMessage($"{e.StackTrace.ToString()}");
                    apiHandler.cm.EndLog();
                    Thread.Sleep(30000);
                    excelData.Init(apiHandler.cm);
                }
                //excelData.Info();
                // Prints Excel Data
                //foreach(ExcelSkill sk in excelData.excelSkills)
                //{
                //    if(sk.SkillResourceGroup != null && sk.SkillTeam != null)
                //    {
                //        Console.WriteLine($"{sk.Name} -- RG: {sk.SkillResourceGroup} -- T: {sk.SkillTeam}");
                //    }
                //    else if (sk.SkillResourceGroup != null)
                //    {
                //        Console.WriteLine($"{sk.Name} -- RG: {sk.SkillResourceGroup}");
                //    }
                //    else
                //    {
                //        Console.WriteLine($"{sk.Name}");
                //    }
                //}

                // Takes in ExcelData Object to determine skills required to update each user via API PUT Request
                    // Tries Twice -- only known issue is a rare error involving Server rejecting Auth due to too many requests that occurred during testing
                try
                {
                    apiHandler.ExcelQueueUpdate(excelData);
                }
                catch (Exception e)
                {
                    apiHandler.cm.LogMessage("");
                    apiHandler.cm.LogMessage($"Error occurred during ExcelQueueUpdate Sequence: {e.Message.ToString()}");
                    apiHandler.cm.LogMessage($"{e.StackTrace.ToString()}");
                    apiHandler.cm.EndLog();
                    Thread.Sleep(30000);
                    try
                    {
                        apiHandler.ExcelQueueUpdate(excelData);
                    }
                    catch (Exception y)
                    {
                        apiHandler.cm.LogMessage("");
                        apiHandler.cm.LogMessage("");
                        apiHandler.cm.LogMessage("----------------------------------- END OF PROCESS REPORT ------------------------------------");
                        apiHandler.cm.LogMessage("ERROR -- WFM Queue Update has encountered an unrecoverable error");
                        apiHandler.cm.LogMessage($">Tasked with updating {excelData.excelAgents.Count.ToString()} agents.");
                        apiHandler.cm.LogMessage($">Successfully Updated 0/{excelData.excelAgents.Count.ToString()} agents.");
                        apiHandler.cm.LogMessage(y.Message.ToString());
                        apiHandler.cm.LogMessage(y.StackTrace.ToString());
                        apiHandler.cm.EndLog();
                    }

                }
                // This override is used to remove all skills from the Agents in the Excel Sheet
                //apiHandler.ExcelQueueUpdate(excelData, true);

                // Reports email to users in Workforce Management with a report detailing the App's actions
                emailServices.BuildEmail(false, apiHandler.reportingMessage, apiHandler.cm.LogPath, apiHandler.cm.LogHeader);

                // Print Info that will be uploaded to Reporting Sandbox Database
                //int count = 1;
                //foreach(System.Data.DataRow dr in apiHandler.AgentsUpdatedDT.Rows)
                //{
                //    var dataArray = dr.ItemArray;
                //    Console.WriteLine($"[{count++}] {dataArray[0]} -- {dataArray[1]}");
                //}

                // Clearly defines console output end
                Console.WriteLine("\n");
                for (int i = 0; i < Console.WindowWidth; ++i)
                {
                    Console.Write("-");
                }
                Console.WriteLine("\n");

            }
            catch (Exception f)
            {
                try
                {
                    emailServices.BuildEmail(true, String.Empty, String.Empty, String.Empty);
                    APIHandler apiHandlerLog = new APIHandler();
                    apiHandlerLog.Init();
                    apiHandlerLog.cm.LogMessage($"FATAL ERROR: {f.Message.ToString()}");
                    apiHandlerLog.cm.LogMessage($"{f.StackTrace.ToString()}");
                }
                catch (Exception m)
                {
                    Console.WriteLine(m);
                }
            }
        }
    }
}

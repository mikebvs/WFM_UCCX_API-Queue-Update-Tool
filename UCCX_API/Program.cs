using System;
using System.Collections.Generic;
using System.Threading;
using System.Linq;

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
                //Console.ReadKey();


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
                    Thread.Sleep(15000);
                    excelData.Init(apiHandler.cm);
                }

                
                //Logs all the skills and their corresponding Role in the Excel Sheet Reference that do not exist in the UCCX Skills API
                try
                {
                    apiHandler.cm.BeginLog();
                    apiHandler.cm.LogMessage("");
                    apiHandler.cm.LogMessage("###############################################################################");
                    apiHandler.cm.LogMessage("## SKILLS THAT ARE NOT AVAILABLE WITHIN THE UCCX SKILLS API ARE LISTED BELOW ##");
                    apiHandler.cm.LogMessage("###############################################################################");
                    apiHandler.cm.LogMessage("");
                    string strPad = "[Role]";
                    apiHandler.cm.LogMessage($"{strPad.PadRight(35, ' ')} >Skill");
                    apiHandler.cm.LogMessage($"-------------------------------------------------------------------------------------------");
                    foreach (ExcelSkill sk in excelData.excelSkills)
                    {
                        foreach(KeyValuePair<string, int> kvp in sk.SkillsAdded)
                        {
                            //Console.WriteLine(kvp.Key + $"({kvp.Value.ToString()})");
                            if(!apiHandler.apiData.SkillsData.Skill.Any(x => x.SkillName.Equals(kvp.Key)))
                            {
                                strPad = "[" + sk.Name + "]";
                                apiHandler.cm.LogMessage($"{strPad.PadRight(35, ' ')} >{kvp.Key}");
                                //Console.WriteLine($"Issue with Role: {sk.Name}");
                                //Console.WriteLine($"\tSkill: {kvp.Key} does not exist in the API Skill Data.");
                                //Console.WriteLine();
                            }
                        }
                    }
                    apiHandler.cm.EndLog();
                }
                catch(Exception e)
                {
                    apiHandler.cm.BeginLog();
                    apiHandler.cm.LogMessage("An error occurred while attempting to log the invalid skills within the Excel Reference Sheet.");
                    apiHandler.cm.LogMessage($"{e.Message.ToString()}");
                    apiHandler.cm.LogMessage($"{e.StackTrace.ToString()}");
                    apiHandler.cm.EndLog();
                }


                try
                {
                    List<string> uccxSkills = new List<string>();
                    uccxSkills.Add("Skill ID, UCCX API Skill Name");
                    string uccxSnapshotFile = apiHandler.cm.Configuration.GetSection("UCCXAPISnapshot")["SnapshotFileLocation"].Replace("<DATETIME>", System.DateTime.Now.ToString("yyyy-mm-dd--HH-mm-ss"));
                    foreach (Skill sk in apiHandler.apiData.SkillsData.Skill)
                    {
                        uccxSkills.Add($"{sk.SkillId.ToString()},{sk.SkillName}");
                        //Console.WriteLine($"{sk.SkillName}");
                        //Console.WriteLine($"[{sk.SkillId.ToString()}] {sk.SkillName}");
                    }
                    string[] uccxSkillsArr = uccxSkills.ToArray();
                    System.IO.File.WriteAllLines(uccxSnapshotFile, uccxSkillsArr);
                    apiHandler.cm.BeginLog();
                    apiHandler.cm.LogMessage($"Current UCCX Skills API Data written to file located at: {uccxSnapshotFile}");
                    apiHandler.cm.EndLog();
                }
                catch (Exception e)
                {
                    apiHandler.cm.BeginLog();
                    apiHandler.cm.LogMessage("An error occurred while attempting to write to file the current snapshot of Skills available via the UCCX Skills API.");
                    apiHandler.cm.LogMessage($"{e.Message.ToString()}");
                    apiHandler.cm.LogMessage($"{e.StackTrace.ToString()}");
                    apiHandler.cm.EndLog();
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
                    Thread.Sleep(120000);
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
                if(apiHandler.failedLogging.Count > 0)
                {
                    //string addFailedQueues = "";
                    if(apiHandler.failedLogging.Count > 0)
                    {
                        string appendReporting = "";
                        appendReporting = "\n\nThe following Agent/Queues were found to be invalid:";
                        foreach (string str in apiHandler.failedLogging)
                        {
                            appendReporting += $"\n{str}";
                        }
                        apiHandler.reportingMessage = appendReporting + "\n\n" + apiHandler.reportingMessage;
                    }
                    emailServices.BuildEmail(false, apiHandler.reportingMessage, apiHandler.cm.LogPath, apiHandler.cm.LogHeader, true);
                }
                else
                {
                    emailServices.BuildEmail(false, apiHandler.reportingMessage, apiHandler.cm.LogPath, apiHandler.cm.LogHeader, false);
                }

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
                    emailServices.BuildEmail(true, String.Empty, String.Empty, String.Empty, false);
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

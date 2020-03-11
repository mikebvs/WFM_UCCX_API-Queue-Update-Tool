using System;
using System.IO;
using System.Collections.Generic;
using System.Text;
using System.Linq;
using System.Net;
using System.Xml;
using System.Diagnostics;
using System.Data;

namespace UCCX_API
{
    class APIHandler
    {
        public APIData apiData { get; set; }
        public CredentialManager cm { get; set; }
        public string reportingMessage { get; set; }
        public int updatesFailed { get; set; }
        public void Init()
        {
            //Creates Console Banner
            Console.WriteLine("########################################################################################");
            Console.WriteLine("#################### WORKFORCE MANAGEMENT QUEUE UPDATE UCCX API TOOL ###################");
            Console.WriteLine("########################################################################################\n\n");
            UpdateConsoleStep("Entering Init State...");
            updatesFailed = 0;
            // Credential Manager stores Config Parameters, includes API Auth Credentials, API Root URL depending on Environment, and Logging paths
            this.cm = new CredentialManager();
            try
            {
                // APIData stores necessary information from the API such as the Agents (userID, refURL, etc) and skills (skillId, refURL, etc)
                this.apiData = new APIData(this.cm);
            }
            catch (Exception e)
            {
                cm.LogMessage($"Error Occurred during APIHandler Init Sequence: {e.Message.ToString()}");
                cm.LogMessage($"{e.StackTrace.ToString()}");
            }
        }
        // Method used to update agent queues based on the results of the Excel File built by WFM
        public void ExcelQueueUpdate(ExcelData excelData, bool wipeData = false)
        {
            long totalTime = 0;
            UpdateConsoleStep("WFM Agent Queue Update Process using the UCCX API...");
            // Variables to track for logging/reporting
            int numFailed = 0;
            int numAgentsProcessed = 0;

            Console.SetCursorPosition(Console.CursorLeft, Console.CursorTop + 1);
            UpdateConsoleStep("\t>Agents Processed: " + numAgentsProcessed.ToString() + "/" + excelData.excelAgents.Count.ToString());
            cm.BeginLog();
            cm.LogMessage("Beginning WFM Agent Queue Update Process using the UCCX API.");
            cm.LogMessage("");

            // Begin iterating through each agent to build and update skillmaps
            foreach (ExcelAgent excelAgent in excelData.excelAgents)
            {
                try
                {
                    Stopwatch isw = new Stopwatch();
                    try
                    {
                        isw.Start();
                    }
                    catch (Exception n)
                    {
                        cm.LogMessage($"Error occurred starting StopWatch object: {n.Message.ToString()} -- {n.StackTrace.ToString()}");
                    }
                    // Determine Agent URL via apiData.ResourcesData
                    string agentUserId = apiData.ResourcesData.Resource.Where(p => p.FirstName + " " + p.LastName == excelAgent.agentName).First().UserID;
                    string agentUrl = $"{cm.RootURL}/resource/{agentUserId}";
                    cm.LogMessage($"Updating {excelAgent.agentName}");
                    //// DEBUG -- Prints the built Agent URL for verification ###############################################
                    //Console.WriteLine(agentUrl);
                    ////#####################################################################################################

                    // Determine which Resource from APIData.ResourcesData corresponds to the current excelAgent being processed
                    Resource agentInfo = apiData.ResourcesData.Resource.Where(p => p.FirstName + " " + p.LastName == excelAgent.agentName).First();
                    // Serialize XML related to the current excelAgent being processed using Resource Object
                    XmlDocument xml = SerializeXml(agentInfo);

                    //Console.WriteLine(xml.OuterXml);
                    //Console.ReadKey();
                    // Isolate Old skillMap Node
                    XmlNode node = xml.SelectSingleNode("/resource/skillMap");

                    // Determine Agent's desired Queue
                    ExcelSkill newQueue = excelData.excelSkills.Where(p => p.Name == excelAgent.Queue).First();

                    // Build Skill Map XML to replace current using Agent's desired Queue
                    XmlDocument xmlSkillMap = new XmlDocument();
                    string skillMapString = BuildSkillMap(newQueue);

                    // BuildSkillMap() Returns "ERROR" if the skill was unable to be found
                    if (skillMapString != "ERROR")
                    {
                        xmlSkillMap.LoadXml(skillMapString);

                        // Create new XmlNode object to replace old skillMap with
                        XmlNode newNode = xmlSkillMap.SelectSingleNode("/skillMap");
                        //// DEBUG -- Prints the updated Skill Map XML ##########################################################
                        //Console.WriteLine("########################### NEW SKILL MAP ###########################\n" + node.OuterXml + "\n\n");
                        ////#####################################################################################################

                        // If wipeData == True, remove all skills from Agents
                        if (wipeData == false)
                        {
                            // Replace skillMap Node with new skillMap Node
                            node.InnerXml = newNode.InnerXml;
                        }
                        else
                        {
                            // Remove all skills from Agents
                            node.InnerXml = "";
                        }
                        if(newQueue.SkillResourceGroup != null)
                        {
                            cm.LogMessage($"Updating Resource Group of {excelAgent.agentName} ({agentUserId}) to {newQueue.SkillResourceGroup}.");
                            xml = UpdateResourceGroup(xml, newQueue);
                        }
                        if(newQueue.SkillTeam != null)
                        {
                            cm.LogMessage($"Updating Team of {excelAgent.agentName} ({agentUserId}) to {newQueue.SkillTeam}.");
                            xml = UpdateTeam(xml, newQueue);
                        }
                        try
                        {
                            // Call Method to make PUT Request to API to update Agent skillMap and Log Action/Results
                            cm.LogMessage($"Attempting to update {excelAgent.agentName} ({agentUserId}) to new Queue: {excelAgent.Queue} -- Agent refURL: {agentUrl}");
                            HttpWebResponse requestResponse = UpdateAgentResource(xml.OuterXml, agentUrl);
                            cm.LogMessage($"Status Code Returned: {requestResponse.StatusCode} -- {requestResponse.StatusDescription}\n");
                            //// DEBUG -- Prints HttpWebResponse from PUT Request ###################################################
                            //Console.WriteLine($"{requestResponse.StatusCode}: {requestResponse.StatusDescription}");
                            ////#####################################################################################################                    
                        }
                        catch (Exception e)
                        {
                            numFailed += 1;
                            // Log Error and update Console
                            LogConsoleAndLogFile($"ERROR: {e.Message.ToString()}", 5);
                            cm.LogMessage($"Source: {e.Source.ToString()}");
                            cm.LogMessage($"Stack Trace: {e.StackTrace.ToString()}");
                            if (e.Message.ToString().Contains("SSL"))
                            {
                                throw new System.Exception("SSL Error", e);
                            }
                        }
                    }
                    else
                    {
                        cm.LogMessage($"Unable to successfully update {excelAgent.agentName}, moving to the next Agent if available.");
                        numFailed += 1;
                    }
                    numAgentsProcessed += 1;
                    if(isw != null)
                    {
                        if (isw.IsRunning)
                        {
                        isw.Stop();
                        totalTime += isw.ElapsedMilliseconds;
                        // End Process Console Output
                            UpdateConsoleStep($"\t>Agents Processed: {numAgentsProcessed.ToString()}/{excelData.excelAgents.Count.ToString()} ({isw.ElapsedMilliseconds.ToString()}ms, {totalTime.ToString()}ms Total)");
                            LogConsoleAndLogFile($"\t>Successfully updated: {(numAgentsProcessed - numFailed).ToString()}.", 1, false, false);
                            LogConsoleAndLogFile($"\t>Failed to Update: {numFailed.ToString()}.", 2, false, false);
                            cm.LogMessage($"Time Elapsed: {isw.ElapsedMilliseconds.ToString()}ms, Total Time Elapsed: {totalTime.ToString()}ms");
                        }
                    }
                    else
                    {
                        totalTime += 5000;
                        UpdateConsoleStep($"\t>Agents Processed: {numAgentsProcessed.ToString()}/{excelData.excelAgents.Count.ToString()} (N/A ms, {totalTime.ToString()}ms Total)");
                        LogConsoleAndLogFile($"\t>Successfully updated: {(numAgentsProcessed - numFailed).ToString()}.", 1, false, false);
                        LogConsoleAndLogFile($"\t>Failed to Update: {numFailed.ToString()}.", 2, false, false);
                        cm.LogMessage($"Time Elapsed: N/A ms, Total Time Elapsed: {totalTime.ToString()}ms");
                    }
                }
                catch (Exception e)
                {
                    cm.LogMessage($"Error Occurred Updating Agent: {excelAgent.agentName} -- {e.Message.ToString()}");
                    cm.LogMessage($"Source: {e.Source.ToString()}");
                    cm.LogMessage($"Stack Trace: {e.StackTrace.ToString()}");
                    if (e.Message.ToString().Contains("SSL"))
                    {
                        throw new System.Exception("SSL Error", e);
                    }
                }
                cm.LogMessage("");
            }
            updatesFailed = numFailed;
            cm.LogMessage("");
            cm.LogMessage("");
            cm.LogMessage("----------------------------------- END OF PROCESS REPORT ------------------------------------");
            EndConsoleLog(excelData.excelAgents.Count, numFailed, totalTime);
            cm.EndLog();
            reportingMessage = $"------------- WFM Queue Update has completed -------------";
            reportingMessage += $"\n|\t>Agents Processed: {numAgentsProcessed.ToString()}/{excelData.excelAgents.Count.ToString()} (Total Elapsed Time: {totalTime.ToString()}ms)";
            reportingMessage += $"\n|\t>Successfully updated: {(numAgentsProcessed - numFailed).ToString()}.";
            reportingMessage += $"\n|\t>Failed to Update: {numFailed.ToString()}.";
            reportingMessage += $"\n----------------------------------------------------------------------------------------------";
        }
        public XmlDocument UpdateTeam(XmlDocument currentDoc, ExcelSkill skill)
        {
            if (skill.SkillResourceGroup != null)
            {
                // Find Resource Group from the API that matches the ExcelSkill's SkillResourceGroup Name
                ResourceGroupData apiRG = apiData.resourceGroupsData.ResourceGroup.Where(p => p.Name == skill.SkillResourceGroup).First();
                // Build XML off of template
                string resourceTemplate = "<resourceGroup name=\"<RESOURCE_NAME>\"><refURL><REF_URL></refURL></resourceGroup>";
                resourceTemplate = resourceTemplate.Replace("<RESOURCE_NAME>", apiRG.Name).Replace("<REF_URL>", apiRG.Self);

                // Create new XML Document to replace InnerXML of currentDoc with
                XmlDocument xmlRG = new XmlDocument();
                xmlRG.LoadXml(resourceTemplate);
                // Isolate the Nodes we want to edit
                XmlNode newNode = xmlRG.SelectSingleNode("/resourceGroup");
                XmlNode oldNode = currentDoc.SelectSingleNode("/resource/resourceGroup");
                // Replace the name attribute of the outerXml
                oldNode.Attributes[0].Value = apiRG.Name;

                // Replace the InnerXml
                oldNode.InnerXml = newNode.InnerXml;
            }
            return currentDoc;
        }
        public XmlDocument UpdateResourceGroup(XmlDocument currentDoc, ExcelSkill skill) 
        {
            if(skill.SkillTeam != null)
            {
                // Find the Team from the API that matches the ExcelSkill's SkillTeam Name
                TeamData apiTD = apiData.teamsData.Team.Where(p => p.Teamname == skill.SkillTeam).First();
                // Build XML off of template
                string teamTemplate = "<team name=\"<TEAM_NAME>\"><refURL><REF_URL></refURL></team>";
                teamTemplate = teamTemplate.Replace("<TEAM_NAME>", apiTD.Teamname).Replace("<REF_URL>", apiTD.Self);

                // Create new XML Document to replace InnerXML of currentDoc with
                XmlDocument xmlTD = new XmlDocument();
                xmlTD.LoadXml(teamTemplate);
                // Isolate the Nodes we want to edit
                XmlNode newNode = xmlTD.SelectSingleNode("/team");
                XmlNode oldNode = currentDoc.SelectSingleNode("/resource/team");
                // Replace the name attribute of the outerXml
                oldNode.Attributes[0].Value = apiTD.Teamname;

                // Replace the InnerXml
                oldNode.InnerXml = newNode.InnerXml;
            }
            return currentDoc;
        }
        public XmlDocument SerializeXml<T>(T serializeObject, bool stripNamespace = true)
        {
            string xmlString = "";
            
            // Create our own namespaces for the output
            System.Xml.Serialization.XmlSerializerNamespaces namespaces = new System.Xml.Serialization.XmlSerializerNamespaces();

            //Add an empty namespace and empty value
            namespaces.Add("", "");

            //System.Xml.Serialization.XmlSerializer x = new System.Xml.Serialization.XmlSerializer(serializeObject.GetType());
            using (StringWriter textWriter = new StringWriter())
            {
                using (XmlWriter writer = XmlWriter.Create(textWriter, new XmlWriterSettings { OmitXmlDeclaration = true }))
                {
                    if (stripNamespace == true)
                    {
                        new System.Xml.Serialization.XmlSerializer(serializeObject.GetType()).Serialize(writer, serializeObject, namespaces);
                    }
                    else
                    {
                        new System.Xml.Serialization.XmlSerializer(serializeObject.GetType()).Serialize(writer, serializeObject);
                    }
                }
                xmlString = textWriter.ToString();
            }
            
            XmlDocument xml = new XmlDocument();
            xml.LoadXml(xmlString);
            return xml;
        }
        // Gets all skills associated with the input Queue name
        private string BuildSkillMap(ExcelSkill newQueue)
        {
            // Template skillMap XML --> Uses Replace in order to build new skillMap string
            string templateSkill = "<skillCompetency><competencelevel>COMPETENCY_LEVEL</competencelevel><skillNameUriPair name=\"SKILL_NAME\"><refURL>REF_URL</refURL></skillNameUriPair></skillCompetency>";
            string skillMap = "";
            string skillsReport = "";
            //// DEBUG -- Prints the Queue Name and all associated Skills/Competency Levels #########################
            //Console.WriteLine(agent.Queue);
            //foreach(KeyValuePair<string, int> kvp in newQueue.SkillsAdded)
            //{
            //    Console.WriteLine("   >Adding " + kvp.Key + ": " + kvp.Value);
            //    Console.WriteLine(templateSkill.Replace("COMPETENCY_LEVEL", kvp.Value.ToString()).Replace("SKILL_NAME", kvp.Key).Replace("REF_URL",addSkill.Self) + "\n\n");
            //}
            ////#####################################################################################################
            bool allSkillsFound = true;
            foreach(KeyValuePair<string, int> kvp in newQueue.SkillsAdded)
            {
                try
                {
                    // Determine Skill refUrl by querying APIData
                    Skill addSkill = apiData.SkillsData.Skill.Where(p => p.SkillName.ToUpper() == kvp.Key.ToUpper()).First();
                    //Append skillMap string with new info
                    skillMap += templateSkill.Replace("COMPETENCY_LEVEL", kvp.Value.ToString()).Replace("SKILL_NAME", addSkill.SkillName).Replace("REF_URL", addSkill.Self);

                    skillsReport += $"{kvp.Key}({kvp.Value.ToString()}), ";
                }
                catch
                {
                    // Log and output to console Error
                    LogConsoleAndLogFile($"The skill {kvp.Key} was unable to be found within the API Skills. Skill name formatting is extremely strict.", 5);
                    allSkillsFound = false;
                    break;
                }

            }
            if(allSkillsFound == true)
            {
                if(skillsReport.Length > 3)
                {
                    skillsReport = skillsReport.Substring(0, skillsReport.Length - 2);
                }
                else
                {
                    skillsReport = skillsReport.Replace(", ", "");
                }
                cm.LogMessage($"Building Skill Map required for {newQueue.Name.ToString()}: {skillsReport.ToString()}");
                // Add XML Outer Node onto new skillMap contents prior to replacing old skillMap XML Node
                skillMap = $"<skillMap>{skillMap}</skillMap>";

                return skillMap;
            }
            else
            {
                return "ERROR";
            }
        }
        private HttpWebResponse UpdateAgentResource(string requestXml, string agentUrl)
        {
            HttpWebRequest request = (HttpWebRequest)WebRequest.Create(agentUrl);

            // Add Basic Authorization Headers
            String encoded = System.Convert.ToBase64String(System.Text.Encoding.GetEncoding("ISO-8859-1").GetBytes(cm.Username + ":" + cm.Password));
            request.Headers.Add("Authorization", "Basic " + encoded);

            // Add Standard Encoding (Do not add into Request Body Header)
            byte[] bytes;
            bytes = System.Text.Encoding.ASCII.GetBytes(requestXml);
            request.ContentType = "text/xml; encoding='utf-8'";
            request.ContentLength = bytes.Length;

            // Method is PUT, not POST
            request.Method = "PUT";
            Stream requestStream = request.GetRequestStream();
            requestStream.Write(bytes, 0, bytes.Length);
            requestStream.Close();

            LogConsoleAndLogFile($"Sending PUT Request to: {agentUrl}", 0);

            // return reponse to action in previous scope
            HttpWebResponse response;
            response = (HttpWebResponse)request.GetResponse();
            return response;
        }
        public void Info()
        {
            Console.WriteLine("\n\n###################################################################");
            Console.WriteLine("###################### CREDENTIAL MANAGER DATA ######################");
            Console.WriteLine("#####################################################################\n");
            cm.Info();

            Console.WriteLine("\n\n###################################################################");
            Console.WriteLine("############################# API DATA ##############################");
            Console.WriteLine("#####################################################################\n");
            apiData.Info();
        }
        public void UpdateConsoleStep(string message)
        {
            int currentLineCursor = Console.CursorTop;
            Console.SetCursorPosition(0, Console.CursorTop);
            Console.Write(new string(' ', Console.WindowWidth));
            Console.SetCursorPosition(0, currentLineCursor);
            Console.Write(message);
        }
        public void LogConsoleAndLogFile(string message, int cursorTop, bool wait = false, bool toLog = true)
        {
            if(cursorTop == 0)
            {
                cm.LogMessage(message);
            }
            else
            {
                int yPlacement = Console.CursorTop;
                Console.SetCursorPosition(0, Console.CursorTop + cursorTop);
                UpdateConsoleStep(message);
                Console.SetCursorPosition(0, yPlacement);
                if(toLog == true)
                {
                    cm.LogMessage(message);
                }
            }
            if(wait == true && cm.Env == "DEV")
            {
                Console.ReadKey();
            }
        }
        public void EndConsoleLog(int totalAgents, int numFailed, long timeElapsed)
        {
            // Console Reporting Formatting
            UpdateConsoleStep("");
            Console.SetCursorPosition(Console.CursorLeft, Console.CursorTop - 1);
            Console.Write(new string(' ', Console.WindowWidth));
            Console.SetCursorPosition(0, Console.CursorTop);
            UpdateConsoleStep("Process Finished...");
            Console.SetCursorPosition(Console.CursorLeft, Console.CursorTop + 1);
            UpdateConsoleStep($"\t>Attempted to update {totalAgents.ToString()} Agents ({timeElapsed.ToString()}ms).\n\n");
            Console.SetCursorPosition(0, Console.CursorTop + 1);
            // Logging to Log File
            cm.LogMessage($"Finished WFM Agent Queue Update Process using the UCCX API.");
            if (numFailed > 0)
            {
                cm.LogMessage($">Attempted to update {totalAgents.ToString()} Agents ({timeElapsed.ToString()}ms).");
                cm.LogMessage($">WARNING: {numFailed.ToString()}/{totalAgents.ToString()} Agents failed to update.");
                cm.LogMessage($">{(totalAgents - numFailed).ToString()}/{totalAgents.ToString()} Agents successfully updated.");
            }
            else
            {
                cm.LogMessage($">{totalAgents.ToString()} Agents successfully updated ({timeElapsed.ToString()}ms).");
            }
        }
    }
}

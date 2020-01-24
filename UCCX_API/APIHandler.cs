using System;
using System.IO;
using System.Collections.Generic;
using System.Text;
using System.Linq;
using System.Net;
using System.Xml;

namespace UCCX_API
{
    class APIHandler
    {
        public APIData apiData { get; set; }
        public CredentialManager cm { get; set; }
        public void Init()
        {
            //Creates Console Banner
            Console.WriteLine("########################################################################################");
            Console.WriteLine("###################### WORKFORCE MANAGEMENT QUEUE UPDATE API TOOL ######################");
            Console.WriteLine("########################################################################################\n\n");
            UpdateConsoleStep("Entering Init State...");
            // Credential Manager stores Config Parameters, includes API Auth Credentials, API Root URL depending on Environment, and Logging paths
            this.cm = new CredentialManager();
            // APIData stores necessary information from the API such as the Agents (userID, refURL, etc) and skills (skillId, refURL, etc)
            this.apiData = new APIData(this.cm);
        }
        // Method used to update agent queues based on the results of the Excel File built by WFM
        public void ExcelQueueUpdate(ExcelData excelData)
        {
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
            foreach(ExcelAgent excelAgent in excelData.excelAgents)
            {
                // Determine Agent URL via apiData.ResourcesData
                string agentUserId = apiData.ResourcesData.Resource.Where(p => p.FirstName + " " + p.LastName == excelAgent.agentName).First().UserID;
                string agentUrl = $"{cm.RootURL}/resource/{agentUserId}";
                cm.LogMessage($"Updating {excelAgent.agentName}");
                //// DEBUG -- Prints the built Agent URL for verification ###############################################
                //Console.WriteLine(agentUrl);
                ////#####################################################################################################

                try
                {

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
                    if(skillMapString != "ERROR")
                    {
                        xmlSkillMap.LoadXml(skillMapString);

                        // Create new XmlNode object to replace old skillMap with
                        XmlNode newNode = xmlSkillMap.SelectSingleNode("/skillMap");
                        //// DEBUG -- Prints the updated Skill Map XML ##########################################################
                        //Console.WriteLine("########################### NEW SKILL MAP ###########################\n" + node.OuterXml + "\n\n");
                        ////#####################################################################################################

                        // Replace skillMap Node with new skillMap Node
                        node.InnerXml = newNode.InnerXml;
                        //// DEBUG -- Removes all skillMap InnerXml Nodes, comment in to reset user skillMaps for testing #######
                        //node.InnerXml = "";
                        ////#####################################################################################################
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
                        }
                    }
                    else
                    {
                        cm.LogMessage($"Unable to successfully update {excelAgent.agentName}, moving to the next Agent if available.");
                        numFailed += 1;
                    }
                    numAgentsProcessed += 1;

                    // End Process Console Output
                    UpdateConsoleStep("\t>Agents Processed: " + numAgentsProcessed.ToString() + "/" + excelData.excelAgents.Count.ToString());
                    LogConsoleAndLogFile($"\t>Successfully updated: {(numAgentsProcessed - numFailed).ToString()}.", 1, false, false);
                    LogConsoleAndLogFile($"\t>Failed to Update: {numFailed.ToString()}.", 2, false, false);
                }
                catch(Exception e)
                {
                    cm.LogMessage($"Error Occurred Serializing XML Data from UCCX API: {e.Message.ToString()}");
                }
                cm.LogMessage("");
            }
            cm.EndLog();
            cm.BeginLog();
            EndConsoleLog(excelData.excelAgents.Count, numFailed);
            cm.EndLog();
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
                    skillsReport = skillsReport.Substring(0, skillsReport.Length - 3);
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
        public void EndConsoleLog(int totalAgents, int numFailed)
        {
            // Console Reporting Formatting
            UpdateConsoleStep("");
            Console.SetCursorPosition(Console.CursorLeft, Console.CursorTop - 1);
            int currentLineCursor = Console.CursorTop;
            Console.Write(new string(' ', Console.WindowWidth));
            Console.SetCursorPosition(0, Console.CursorTop);
            UpdateConsoleStep("Process Finished...");
            Console.SetCursorPosition(Console.CursorLeft, Console.CursorTop + 1);
            UpdateConsoleStep($"\t>Attempted to update {totalAgents.ToString()} Agents.\n\n");
            Console.SetCursorPosition(0, Console.CursorTop + 1);
            // Logging to Log File
            cm.LogMessage($"Finished WFM Agent Queue Update Process using the UCCX API.");
            if (numFailed > 0)
            {
                cm.LogMessage($">Attempted to update {totalAgents.ToString()} Agents.");
                cm.LogMessage($">WARNING: {numFailed.ToString()}/{totalAgents.ToString()} Agents failed to update.");
                cm.LogMessage($">{(totalAgents - numFailed).ToString()}/{totalAgents.ToString()} Agents successfully updated.");
            }
            else
            {
                cm.LogMessage($">{totalAgents.ToString()} Agents successfully updated.");
            }
        }
    }
}

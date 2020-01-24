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
            
            // Begin iterating through each agent to build and update skillmaps
            foreach(ExcelAgent excelAgent in excelData.excelAgents)
            {
                // Determine Agent URL via apiData.ResourcesData
                string agentUserId = apiData.ResourcesData.Resource.Where(p => p.FirstName + " " + p.LastName == excelAgent.agentName).First().UserID;
                string agentUrl = $"{cm.RootURL}/resource/{agentUserId}";
                //// DEBUG -- Prints the built Agent URL for verification ###############################################
                //Console.WriteLine(agentUrl);
                ////#####################################################################################################

                // Request Agent API Info to modify for updated skillMap
                WebRequest apiRequest = WebRequest.Create(agentUrl);
                string encoded = System.Convert.ToBase64String(System.Text.Encoding.GetEncoding("ISO-8859-1").GetBytes(cm.Username + ":" + cm.Password));
                apiRequest.Headers.Add("Authorization", "Basic " + encoded);
                HttpWebResponse apiResponse = (HttpWebResponse)apiRequest.GetResponse();
                string xmlOutput;
                if (apiResponse.StatusCode == HttpStatusCode.OK)
                {
                    using (StreamReader sr = new StreamReader(apiResponse.GetResponseStream()))
                        xmlOutput = sr.ReadToEnd();
                    XmlDocument xml = new XmlDocument();
                    xml.LoadXml(xmlOutput);
                    // Isolate Old skillMap Node
                    XmlNode node = xml.SelectSingleNode("/resource/skillMap");

                    // Determine Agent's desired Queue
                    ExcelSkill newQueue = excelData.excelSkills.Where(p => p.Name == excelAgent.Queue).First();
                    // Build Skill Map XML to replace current using Agent's desired Queue
                    XmlDocument xmlSkillMap = new XmlDocument();
                    xmlSkillMap.LoadXml(BuildSkillMap(newQueue));
                    

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
                        cm.LogMessage($"Attempting to update {excelAgent.agentName} ({agentUserId}) to new Queue: {excelAgent.Queue}\n\tAgent refURL: {agentUrl}");
                        HttpWebResponse requestResponse = UpdateAgentResource(xml.OuterXml, agentUrl);
                        cm.LogMessage($"Status Code Returned: {requestResponse.StatusCode} -- {requestResponse.StatusDescription}\n");
                        //// DEBUG -- Prints HttpWebResponse from PUT Request ###################################################
                        //Console.WriteLine($"{requestResponse.StatusCode}: {requestResponse.StatusDescription}");
                        ////#####################################################################################################                    
                    }
                    catch (Exception e)
                    {
                        numFailed += 1;
                        cm.LogMessage("ERROR: " + e.Message.ToString());
                        // HANDLE ERROR/BAD RESPONSE
                    }
                    numAgentsProcessed += 1;

                    // Console Output
                    UpdateConsoleStep("\t>Agents Processed: " + numAgentsProcessed.ToString() + "/" + excelData.excelAgents.Count.ToString());
                    Console.SetCursorPosition(Console.CursorLeft, Console.CursorTop + 1);
                    UpdateConsoleStep($"\t>Successfully updated: {(numAgentsProcessed - numFailed).ToString()}.");
                    Console.SetCursorPosition(Console.CursorLeft, Console.CursorTop + 1);
                    UpdateConsoleStep($"\t>Failed to Update: {numFailed.ToString()}.");
                    Console.SetCursorPosition(Console.CursorLeft, Console.CursorTop - 2);
                }
                else
                {
                    // HANDLE HTTP STATUS CODES (404 and 403 == break maybe?)
                }
            }
            UpdateConsoleStep("");
            Console.SetCursorPosition(Console.CursorLeft, Console.CursorTop - 1);
            int currentLineCursor = Console.CursorTop;
            Console.Write(new string(' ', Console.WindowWidth));
            Console.SetCursorPosition(0, Console.CursorTop);
            UpdateConsoleStep("Process Finished...");
            Console.SetCursorPosition(Console.CursorLeft, Console.CursorTop + 1);
            UpdateConsoleStep($"\t>Attempted to update {excelData.excelAgents.Count.ToString()} Agents.\n\n");

            // Logging
            cm.LogMessage($"Finished WFM Agent Queue Update Process using the UCCX API.\n\tUpdated {excelData.excelAgents.Count.ToString()} Agents.");
            if(numFailed > 0)
            {
                cm.LogMessage($"\n\tWARNING: {numFailed.ToString()}/{excelData.excelAgents.Count.ToString()} Agents failed to update.");
            }
            cm.LogMessage($"\n\t{(excelData.excelAgents.Count - numFailed).ToString()}/{excelData.excelAgents.Count.ToString()} Agents successfully updated.");
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
            foreach(KeyValuePair<string, int> kvp in newQueue.SkillsAdded)
            {
                // Determine Skill refUrl by querying APIData
                Skill addSkill = apiData.SkillsData.Skill.Where(p => p.SkillName == kvp.Key).First();

                //Append skillMap string with new info
                skillMap += templateSkill.Replace("COMPETENCY_LEVEL", kvp.Value.ToString()).Replace("SKILL_NAME", kvp.Key).Replace("REF_URL", addSkill.Self);

                skillsReport += $"{kvp.Key}({kvp.Value.ToString()}), ";
            }
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
    }
}

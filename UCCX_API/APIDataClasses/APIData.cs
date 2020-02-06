using System;
using System.Collections.Generic;
using System.Text;

namespace UCCX_API
{
    class APIData : APIHandler
    {
        public Resources ResourcesData { get; set; }
        public Skills SkillsData { get; set; }
        public Teams teamsData { get; set; }
        public ResourceGroups resourceGroupsData { get; set; }
        public APIData(CredentialManager cm)
        {
            // Deserialize Agent Data from API
            UpdateConsoleStep("Fetching Resource Data from UCCX API...");
            ResourcesData = ApiWebRequestHelper.GetXmlRequest<Resources>(cm.RootURL + "/resource", cm.Username, cm.Password);
            // Deserialize Skills Data from API
            UpdateConsoleStep("Fetching Skill Data from UCCX API...");
            SkillsData = ApiWebRequestHelper.GetXmlRequest<Skills>(cm.RootURL + "/skill", cm.Username, cm.Password);
            UpdateConsoleStep("Fetching Teams Data from UCCX API...");
            teamsData = ApiWebRequestHelper.GetXmlRequest<Teams>(cm.RootURL + "/team", cm.Username, cm.Password);
            UpdateConsoleStep("Fetching Resource Group Data from UCCX API...");
            resourceGroupsData = ApiWebRequestHelper.GetXmlRequest<ResourceGroups>(cm.RootURL + "/resourceGroup", cm.Username, cm.Password);
        }
        public new void Info()
        {
            Console.WriteLine("\n\n###################################################################");
            Console.WriteLine("########################## RESOURCE DATA ##########################");
            Console.WriteLine("###################################################################\n");
            foreach(Resource rs in ResourcesData.Resource)
            {
                Console.WriteLine($"{rs.FirstName} {rs.LastName} ({rs.UserID}) -- {rs.Extension}\n\trefURL: {rs.Self}");
            }
            Console.WriteLine("\n\n###################################################################");
            Console.WriteLine("########################### SKILLS DATA ###########################");
            Console.WriteLine("###################################################################\n");
            foreach(Skill sk in SkillsData.Skill)
            {
                Console.WriteLine($"[{sk.SkillId}] {sk.SkillName}\n\trefURL: {sk.Self}");
            }
            Console.WriteLine("\n\n###################################################################");
            Console.WriteLine("############################ TEAMS DATA ###########################");
            Console.WriteLine("###################################################################\n");
            foreach(ResourceGroupData rg in resourceGroupsData.ResourceGroup)
            {
                Console.WriteLine($"[{rg.Id}] {rg.Name} -- {rg.Self}");
            }
            Console.WriteLine("\n\n###################################################################");
            Console.WriteLine("####################### RESOURCE GROUP DATA #######################");
            Console.WriteLine("###################################################################\n");
            foreach (TeamData td in teamsData.Team)
            {
                Console.WriteLine($"[{td.TeamId}] {td.Teamname} -- {td.Self}");
                if (td.PrimarySupervisor != null)
                {
                    Console.WriteLine($"\tPrimary Supervisor: {td.PrimarySupervisor.Name}\n");
                }
                if (td.SecondarySupervisors != null)
                {
                    foreach (SecondrySupervisor sc in td.SecondarySupervisors.SecondrySupervisor)
                    {
                        Console.WriteLine($"\tSecondary Supervisor: {sc.Name.ToString()}\n");
                    }
                }
            }
            Console.WriteLine("\n");
        }
    }
}

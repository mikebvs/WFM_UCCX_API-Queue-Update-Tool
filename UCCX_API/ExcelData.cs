using System;
using System.Collections.Generic;
using System.Text;

namespace UCCX_API
{
    class ExcelData : APIHandler
    {
        public ExcelData(CredentialManager cm)
        {
            reader = new Reader(cm.ExcelFile);
            UpdateConsoleStep("Reading Excel Agent Data...");
            excelAgents = reader.ReadAgentData(0);
            UpdateConsoleStep("Reading Excel Queue Data...");
            excelSkills = reader.ReadSkillData(1);
        }
        public Reader reader { get; set; }
        public List<ExcelAgent> excelAgents { get; set; }
        public List<ExcelSkill> excelSkills { get; set; }
        public void Info()
        {
            Console.WriteLine("\n\n###################################################################");
            Console.WriteLine("############################ AGENT DATA #############################");
            Console.WriteLine("#####################################################################\n");
            foreach (ExcelAgent agent in excelAgents)
            {
                agent.Info();
            }
            Console.WriteLine("\n\n###################################################################");
            Console.WriteLine("######################## SKILL (EXCEL) DATA #########################");
            Console.WriteLine("#####################################################################\n");
            foreach (ExcelSkill sk in excelSkills)
            {
                sk.Info();
            }
            Console.WriteLine("\n");
        }
    }
}

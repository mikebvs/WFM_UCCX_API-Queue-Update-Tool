using System;
using System.Threading;

namespace UCCX_API
{
    class Program
    {
        static void Main(string[] args)
        {
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
                Thread.Sleep(60000);
                apiHandler.Init();
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
            }
            // This override is used to remove all skills from the Agents in the Excel Sheet
            //apiHandler.ExcelQueueUpdate(excelData, true);

            // Clearly defines console output end
            for (int i = 0; i < Console.WindowWidth; ++i)
            {
                Console.Write("-");
            }
            Console.WriteLine("\n");

        }
    }
}

using System;
using System.Collections.Generic;
using System.Xml;
using System.Xml.Serialization;
using ExtendedXmlSerializer;
using ExtendedXmlSerializer.Configuration;
using System.Net;
using System.Linq;
using System.IO;
using Microsoft.Extensions.Configuration;

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
            apiHandler.Init();

            // ExcelData is comprised of the updated info pulled from WFM Excel Sheet
                // Reads Agents and their desired Queue
                // Reads Queues and the required skills to be added to the Queue
            ExcelData excelData = new ExcelData(apiHandler.cm);

            // Takes in ExcelData Object to determine skills required to update each user via API PUT Request
            apiHandler.ExcelQueueUpdate(excelData);

            // Clearly defines console output end
            for (int i = 0; i < Console.WindowWidth; ++i)
            {
                Console.Write("-");
            }
            Console.WriteLine("\n");

        }
    }
}

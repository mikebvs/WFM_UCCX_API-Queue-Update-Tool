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
            APIHandler apiHandler = new APIHandler();
            apiHandler.Init();
            //Pulls information from Environment locked config files
            //apiHandler.cm = new CredentialManager();
            //Pulls all resource/skill data from API
            //apiHandler.apiData = new APIData(apiHandler.cm);
            //// DEBUG -- lines below print output for verification #################################################
            //apiHandler.Info();
            ////#####################################################################################################



            //// DEBUG -- lines below print output for verification #################################################
            //// Example of Request to get a single Agent
            //Agent agent = ApiWebRequestHelper.GetXmlRequest<Agent>(apiHandler.cm.RootURL + "/resource/whi01022", apiHandler.cm.Username, apiHandler.cm.Password);
            //agent.Info();
            ////#####################################################################################################

            ExcelData excelData = new ExcelData(apiHandler.cm);
            //// DEBUG -- Prints ExcelSkill and ExcelAgent Data for verification ####################################
            //excelData.Info();
            ////#####################################################################################################

            // Takes in ExcelData Object to determine skills required to update each user via API PUT Request
            apiHandler.ExcelQueueUpdate(excelData);

        }
    }
}

﻿using System;
using System.Collections.Generic;
using System.Text;

namespace UCCX_API
{
    class ExcelAgent
    {
        public string agentName { get; set; }
        public string FirstName { get; set; }
        public string LastName { get; set; }
        public string Queue { get; set; }
        public ExcelAgent(string sheetName, string sheetQueue)
        {
            //Determine First Name, Last Name and Queue name
            agentName = sheetName;
            if (agentName.Contains(" "))
            {
                FirstName = sheetName.Substring(0, sheetName.IndexOf(" "));
                LastName = sheetName.Substring(sheetName.IndexOf(FirstName) + 1);
            }
            else
            {
                FirstName = agentName;
                LastName = "";
            }
            Queue = sheetQueue;
        }
        public void Info()
        {
            Console.WriteLine(agentName + " -- " + Queue);
        }
    }
}

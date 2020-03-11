using System;
using System.IO;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using System.Linq;
using OfficeOpenXml;
using System.Text;

namespace UCCX_API
{
    class Reader
    {
        string filePath { get; set; }
        public Reader(string file)
        {
            filePath = file;
        }
        public List<ExcelSkill> ReadSkillData(int sheetIndex)
        {
            FileInfo file = new FileInfo(filePath);
            List<ExcelSkill> skillData = new List<ExcelSkill>();
            using (ExcelPackage package = new ExcelPackage(file))
            {

                StringBuilder sb = new StringBuilder();
                ExcelWorksheet worksheet = package.Workbook.Worksheets[sheetIndex];
                int rowCount = worksheet.Dimension.Rows;
                int colCount = worksheet.Dimension.Columns;
                //Console.WriteLine("ROWS: " + rowCount.ToString() + "\nCOLUMNS: " + colCount.ToString());

                for (int i = 2; i <= rowCount; i++)
                {
                    string name = String.Empty;
                    string add = String.Empty;
                    string remove = String.Empty;
                    string skillTeam = String.Empty;
                    string skillResource = String.Empty;
                    for (int j = 1; j <= colCount; j++)
                    {
                        if (j == 1 && worksheet.Cells[i, j].Value != null)
                        {
                            name = worksheet.Cells[i, j].Value.ToString().Trim();
                        }
                        else if (j == 2 && worksheet.Cells[i, j].Value != null)
                        {
                            add = worksheet.Cells[i, j].Value.ToString().Replace(" ", "").Trim();
                        }
                        else if (j == 3 && worksheet.Cells[i, j].Value != null)
                        {
                            remove = worksheet.Cells[i, j].Value.ToString().Replace(" ", "").Trim();
                        }
                        else if (j == 4 && worksheet.Cells[i, j].Value != null)
                        {
                            skillResource = worksheet.Cells[i, j].Value.ToString().Replace(" ", "").Trim();
                        }
                        else if (j == 5 && worksheet.Cells[i, j].Value != null)
                        {
                            skillTeam = worksheet.Cells[i, j].Value.ToString().Replace(" ","").Trim();
                        }
                        else if (add == String.Empty || add == null)
                        {
                            continue;
                        }
                    }
                    if (add != String.Empty && add != null && add != "" && add.Length > 2)
                    {
                        //Console.WriteLine("Adding " + name);
                        ExcelSkill skill = new ExcelSkill(name, add, remove);
                        skillData.Add(skill);
                        if(skillResource != String.Empty && skillResource != null)
                        {
                            skill.SkillResourceGroup = skillResource;
                        }
                        if(skillTeam != String.Empty && skillTeam != null)
                        {
                            skill.SkillTeam = skillTeam;
                        }
                    }
                }
            }
            return skillData;
        }
        public List<ExcelAgent> ReadAgentData(int sheetIndex)
        {

            Regex nameRegex = new Regex("[^a-zA-Z0-9 -]");
            FileInfo file = new FileInfo(filePath);

            List<ExcelAgent> agentData = new List<ExcelAgent>();
            using (ExcelPackage package = new ExcelPackage(file))
            {
                StringBuilder sb = new StringBuilder();
                ExcelWorksheet worksheet = package.Workbook.Worksheets[sheetIndex];
                int rowCount = worksheet.Dimension.Rows;
                int colCount = worksheet.Dimension.Columns;

                for (int i = 2; i <= rowCount; i++)
                {
                    string sheetName = String.Empty;
                    string sheetQueue = String.Empty;
                    for (int j = 1; j <= colCount; j++)
                    {
                        if (j == 1 && worksheet.Cells[i, j].Value != null)
                        {
                            // New version of Sheet Name Retrieval, now replaces all non-alphanumeric characters with nothing in the name.
                            // Implemented due to entries such as "<First_Name> <Last_Name>*" Asterisk would cause the process not to find the agent record via API
                            sheetName = nameRegex.Replace(worksheet.Cells[i, j].Value.ToString(), "");
                            // Old version of Sheet Name Retrieval, very basic.
                            //sheetName = worksheet.Cells[i, j].Value.ToString();
                        }
                        else if (j == 2 & worksheet.Cells[i, j] != null && worksheet.Cells[i, j].Value != null)
                        {
                            sheetQueue = worksheet.Cells[i, j].Value.ToString();
                        }
                        else if (sheetQueue == String.Empty || sheetQueue == null)
                        {
                            continue;
                        }
                    }
                    if (sheetQueue != String.Empty && sheetQueue != null)
                    {
                        ExcelAgent agent = new ExcelAgent(sheetName, sheetQueue);
                        agentData.Add(agent);
                    }
                }
            }
            return agentData;
        }
    }
}

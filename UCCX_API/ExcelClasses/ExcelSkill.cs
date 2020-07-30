using System;
using System.Collections.Generic;
using System.Text;
using System.Linq;

namespace UCCX_API
{
    class ExcelSkill
    {
        public string Name { get; set; }
        public Dictionary<string, int> SkillsAdded { get; set; }
        public List<string> SkillsRemoved { get; set; }
        public string SkillResourceGroup { get; set; }
        public string SkillTeam { get; set; }
        public ExcelSkill(string name, string toAdd, string toRemove)
        {
            // Initialize Name
            Name = name;

            // Create Add Dictionary
            List<string> addList = new List<string>();
            addList.AddRange(toAdd.Split(';'));
            Dictionary<string, int> addDictionary = new Dictionary<string, int>();
            foreach (string str in addList)
            {
                int firstParenth = str.IndexOf("(") + 1;
                int lastParenth = str.LastIndexOf(")");
                int difference = lastParenth - firstParenth;
                string valConvert = str.Substring(firstParenth, difference);
                string key = str.Substring(0, firstParenth - 1);
                int val = Convert.ToInt32(valConvert);
                if (addDictionary.ContainsKey(key))
                {
                    Console.WriteLine($"Key already present: {key}, updating the skill value in the Dictionary from {addDictionary.FirstOrDefault(x => x.Key == key).Value.ToString()} to {val.ToString()}.");
                    addDictionary[key] = val;
                }
                else
                {
                    addDictionary.Add(key, val);
                }
            }
            // Initialize Add
            SkillsAdded = addDictionary;

            // Create Remove List
            List<string> removeList = new List<string>();
            removeList.AddRange(toRemove.Split(';'));
            // Initialize Remove
            SkillsRemoved = removeList;
        }
        public void Info()
        {
            Console.WriteLine("\n########### " + Name + " ###########");
            if(SkillResourceGroup != null)
            {
                Console.WriteLine($"  >Resource Group: {SkillResourceGroup}");
            }
            if(SkillTeam != null)
            {
                Console.WriteLine($"  >Team: {SkillTeam}");
            }
            Console.WriteLine("\t__________________________");
            Console.WriteLine("\t|SKILLS ADDED ------------");
            foreach (KeyValuePair<string, int> kvp in SkillsAdded)
            {
                Console.WriteLine("\t|" + kvp.Key + ": " + kvp.Value.ToString());
            }
            Console.WriteLine("\t--------------------------");
            Console.WriteLine("\n");
        }
    }
}

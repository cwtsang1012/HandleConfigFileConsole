using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Configuration;
using System.Collections.Specialized;

namespace ConsoleApplication1
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Practice for retrieving configuration from different section group");
            Console.WriteLine("------------------------------------------------------------------");

            //Open the current configuration files
            var config = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);

            var numOfSectionGrp = Convert.ToInt32(ConfigurationManager.AppSettings["NumOfSectionGrp"]);
            Console.WriteLine("No. of Section Groups defined in App.config: " + numOfSectionGrp);
            //Get Section Group Name Programmatically
            var localSectionGrps = config.SectionGroups.Cast<ConfigurationSectionGroup>().Where(sg => sg.GetType().Name == "ConfigurationSectionGroup").OrderBy(sg => sg.Name);
            foreach (var sectionGrp in localSectionGrps)
            {
                //Get the name of sections under this section group
                var subSectionNames = from s in sectionGrp.Sections.Cast<ConfigurationSection>()
                                  select s.SectionInformation.Name;
                Console.WriteLine(sectionGrp.Name + "(" + String.Join(", ", subSectionNames) + ")");
                foreach (var section in subSectionNames) 
                {
                    RetrievedSectionConfig(sectionGrp.Name, section.ToString());
                }
            }
            Console.WriteLine();

            var numOfSection = Convert.ToInt32(ConfigurationManager.AppSettings["NumOfSection"]);
            Console.WriteLine("No. of Sections defined in App.config: " + numOfSection);
            //Get Section Name Programmatically
            var localSections = config.Sections.Cast<ConfigurationSection>().Where(s => s.SectionInformation.IsDeclared).OrderBy(s => s.SectionInformation.Name);
            foreach (var section in localSections) 
            {
                var secNam = section.SectionInformation.Name;
                //Console.WriteLine(secNam);
                CreateInstance(secNam);
            }
            Console.WriteLine();
            Console.ReadKey();
        }

        /// <summary>
        /// Retrieve value of specific section key
        /// </summary>
        /// <param name="sectionGroupName"></param>
        /// <param name="sectionName"></param>
        public static void RetrievedSectionConfig(string sectionGroupName, string sectionName) 
        {
            var sectionKey = sectionName != "" ? sectionGroupName + "/" + sectionName : sectionName;
            NameValueCollection sectionSettings = ConfigurationManager.GetSection(@sectionKey) as NameValueCollection;
            string sectionId = sectionSettings["SectionID"].ToString();
            string description = sectionSettings["Description"].ToString();
            Console.WriteLine(">> " + sectionName + ": Id = " + sectionId + "; Description = " + description);
        }

        public static void CreateInstance(string sectionName) 
        {
            NameValueCollection sectionSettings = ConfigurationManager.GetSection(sectionName) as NameValueCollection;
            string sectionId = sectionSettings["SectionID"].ToString();
            string description = sectionSettings["Description"].ToString();
            string path = sectionSettings["Path"].ToString();
            
            //new instance of class using generic method
            var ns = typeof(Program).Namespace;
            Type oType = System.Type.GetType(ns + "." + sectionName);
            Test test = (Test)System.Activator.CreateInstance(oType, sectionId, description, path);
            test.Print();
        }
    }
}


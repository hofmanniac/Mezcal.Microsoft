using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Query;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Mezcal.Commands;

namespace Mezcal.Microsoft.CommonDataService
{
    public class CDSImportSolution : ICommand
    {
        public JObject Prompt(CommandEngine commandEngine)
        {
            return null;
        }

        public void Process(JObject command, Context context)
        {
            var source = command["source"].ToString();

            source = context.ReplaceVariables(source);

            this.Import(source, context);
        }

        private void Import(string solutionPath, Context context)
        {
            byte[] fileBytes = File.ReadAllBytes(solutionPath);

            ImportSolutionRequest impSolReqWithMonitoring = new ImportSolutionRequest()
            {
                CustomizationFile = fileBytes,
                ImportJobId = Guid.NewGuid()
            };

            Console.WriteLine($"Importing {solutionPath}");

            var cdsConnection = (CDSConnection)context.DefaultConnection;
            cdsConnection.Execute(impSolReqWithMonitoring);

            var job = cdsConnection.Retrieve("importjob",
               impSolReqWithMonitoring.ImportJobId, new ColumnSet(new System.String[] { "data", "solutionname" }));

            System.Xml.XmlDocument doc = new System.Xml.XmlDocument();
            doc.LoadXml(job["data"].ToString()); // check

            String ImportedSolutionName =
               doc.SelectSingleNode("//solutionManifest/UniqueName").InnerText;

            String SolutionImportResult =
               doc.SelectSingleNode("//solutionManifest/result/@result").Value;

            Console.WriteLine("Report from the ImportJob data");
            Console.WriteLine("Solution Unique name: {0}", ImportedSolutionName);
            Console.WriteLine("Solution Import Result: {0}", SolutionImportResult);
            Console.WriteLine("");

            System.Xml.XmlNodeList optionSets = doc.SelectNodes("//optionSets/optionSet");

            foreach (System.Xml.XmlNode node in optionSets)
            {
                string OptionSetName = node.Attributes["LocalizedName"].Value;
                string result = node.FirstChild.Attributes["result"].Value;

                if (result == "success")
                {
                    Console.WriteLine("{0} result: {1}", OptionSetName, result);
                }
                else
                {
                    string errorCode = node.FirstChild.Attributes["errorcode"].Value;
                    string errorText = node.FirstChild.Attributes["errortext"].Value;

                    Console.WriteLine("{0} result: {1} Code: {2} Description: {3}", OptionSetName,
                    result, errorCode, errorText);
                }
            }

        }
    }
}

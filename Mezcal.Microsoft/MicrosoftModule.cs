using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Mezcal.Commands;
using Mezcal.Connections;
using Mezcal.Microsoft.BingSearch;
using Mezcal.Microsoft.CommonDataService;
using Newtonsoft.Json.Linq;

namespace Mezcal.Microsoft
{
    public class MicrosoftModule : IModule
    {
        //public IConnection ProvideConnection(JObject envConfig)
        //{
        //    IConnection result = null;

        //    if (envConfig.Type == "Microsoft.CommonDataService")
        //    {
        //        result = new CDSConnection(envConfig);
        //    }

        //    return result;
        //}

        public ICommand ResolveCommand(JObject joCommand)
        {
            ICommand result = null;

            var commandName = JSONUtil.GetCommandName(joCommand);

            if (commandName == null) { return null; }
            if (commandName == "cds-entity") { result = new CDSCreateEntity(); }
            else if (commandName == "cds-field") { result = new CDSCreateField(); }
            else if (commandName == "cds-fields") { result = new CDSCreateFields(); }
            else if (commandName == "cds-import-solution") { result = new CDSImportSolution(); }
            else if (commandName == "cds-import-data") { result = new CDSImportData(); }
            else if (commandName == "cds-update-data") { result = new CDSUpdateData(); }
            else if (commandName == "cds-retrieve-data") { result = new CDSRetrieveData(); }
            else if (commandName == "cds-retrieve") { result = new CDSRetrieve(); }
            else if (commandName == "bing-websearch") { result = new SearchBing(); }
            //else if (commandName == "excel-load") { result = new Office.LoadExcel(); }
            else if (commandName == "sharepoint-download") { result = new Office.SharePointDownload(); }
            else if (commandName == "word-export") { result = new Office.WordExport(); }
            else if (commandName == "excel-export") { result = new Office.ExcelExport(); }

            return result;
        }
    }
}

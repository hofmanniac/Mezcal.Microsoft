using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Mezcal.Commands;
using Mezcal.Connections;
using Mezcal.Microsoft.CommonDataService;

namespace Mezcal.Microsoft
{
    public class MicrosoftModule : IModule
    {
        public IConnection ProvideConnection(ConnectionConfig envConfig)
        {
            IConnection result = null;

            if (envConfig.Type == "Microsoft.CommonDataService")
            {
                result = new CommonDataService.CDSConnection(envConfig);
            }

            return result;
        }

        public ICommand ResolveCommand(string commandName)
        {
            ICommand result = null;

            if (commandName == null) { return null; }
            if (commandName == "cds-create-entity") { result = new CDSCreateEntity(); }
            else if (commandName == "cds-import-solution") { result = new CDSImportSolution(); }
            else if (commandName == "cds-import-data") { result = new CDSImportData(); }
            else if (commandName == "cds-retrieve-data") { result = new CDSRetrieveData(); }
            else if (commandName == "cds-retrieve") { result = new CDSRetrieve(); }

            return result;
        }
    }
}

using Microsoft.Xrm.Sdk.Query;
using Mezcal.Commands;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Mezcal.Microsoft.CommonDataService
{
    public class CDSRetrieveData : ICommand
    {
        public void Process(JObject command, Context context)
        {
            string entity = CommandEngine.GetCommandArgument(command, "entity");
            string optionset = null;
            if (entity.Length == 0) { optionset = CommandEngine.GetCommandArgument(command, "Option Set"); }
            string into = CommandEngine.GetCommandArgument(command, "into");
            string env = CommandEngine.GetCommandArgument(command, "env");

            ColumnSet columns = new ColumnSet(true);
            if (command["fields"] != null)
            {
                string fieldlist = command["fields"].ToString();
                if (fieldlist.Length > 0)
                {
                    List<string> fields = fieldlist.Split(',').Select(p => p.Trim()).ToList();
                    columns = new ColumnSet(fields.ToArray());
                }
            }

            var cdsConnection = (CDSConnection)context.GetConnection(env);

            if (optionset == null)
            {
                Console.WriteLine("Loading CDS Entity Data {0} into {1}", entity, into);
             
                var sub = cdsConnection.RetrieveEntityData(entity, columns);
                context.Store(into, sub);
            }
            else
            {
                Console.WriteLine("Loading CDS OptionSet {0} into {1}", entity, into);

                var sub = cdsConnection.RetrieveOptionSet(entity, optionset);
                context.Store(into, sub);
            }
        }

        public JObject Prompt(CommandEngine commandEngine)
        {
            JObject command = new JObject();
            command.Add("command", "retrieve-data");

            commandEngine.PromptForArgument(command, "entity");
            commandEngine.PromptForArgument(command, "fields");
            commandEngine.PromptForArgument(command, "into");
            commandEngine.PromptForArgument(command, "env");

            return command;
        }
    }
}

using Microsoft.Xrm.Sdk.Query;
using Mezcal.Commands;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Mezcal.Connections;

namespace Mezcal.Microsoft.CommonDataService
{
    public class CDSRetrieveData : ICommand
    {
        public void Process(JObject command, Context context)
        {
            var entity = JSONUtil.GetToken(command, "#cds-retrieve-data");
            if (entity == null) { entity = JSONUtil.GetToken(command, "entity"); }
            if (entity == null) { return; }

            //string optionset = null;
            //if (entity.Length == 0) { optionset = CommandEngine.GetCommandArgument(command, "Option Set"); }
            string into = CommandEngine.GetCommandArgument(command, "into");
            //string env = CommandEngine.GetCommandArgument(command, "env");

            CDSConnection cdsConnection = CDSConnection.FromCommand(command, context);
            if (cdsConnection == null) { return; }

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

            //var cdsConnection = (CDSConnection)context.GetConnection(env);

            //if (optionset == null)
            //{
            Console.WriteLine("Loading CDS Entity Data {0} into {1}", entity, into);

            var sub = cdsConnection.RetrieveEntityData(entity.ToString(), columns);
            context.Store(into, sub);
            //}
            //else
            //{
            //Console.WriteLine("Loading CDS OptionSet {0} into {1}", entity, into);

            //    var sub = cdsConnection.RetrieveOptionSet(entity, optionset);
            //    context.Store(into, sub);
            //}
        }

        public JObject Prompt(CommandEngine commandEngine)
        {
            JObject command = new JObject();
            command.Add("command", "cds-retrieve-data");

            commandEngine.PromptForArgument(command, "entity");
            commandEngine.PromptForArgument(command, "fields");
            commandEngine.PromptForArgument(command, "into");
            commandEngine.PromptForArgument(command, "env");

            return command;
        }
    }
}

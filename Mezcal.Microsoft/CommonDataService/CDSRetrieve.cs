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
    public class CDSRetrieve : ICommand
    {
        public JObject Prompt(CommandEngine commandEngine)
        {
            return null;
        }

        public void Process(JObject command, Context context)
        {
            string entityName = command["entity"].ToString();
            string id = command["id"].ToString();
            Guid recordId = Guid.Parse(id);
            ColumnSet columnSet = new ColumnSet(true);
            string into = command["into"].ToString();
            string env = Connections.JSONUtil.GetText(command, "env");

            var cdsConnection = (CDSConnection)context.GetConnection(env);

            var entity = cdsConnection.Retrieve(entityName, recordId, columnSet);

            // todo - convert to json
            context.Store(into, entity);
        }
    }
}

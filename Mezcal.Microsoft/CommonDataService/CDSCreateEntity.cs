using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Mezcal.Commands;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Sdk.Messages;

namespace Mezcal.Microsoft.CommonDataService
{
    public class CDSCreateEntity : ICommand
    {
        public JObject Prompt(CommandEngine commandEngine)
        {
            return null;
        }

        public void Process(JObject command, Context context)
        {
            string schemaName = command["schemaname"].ToString();
            string displayName = command["displayname"].ToString();
            string collectionName = command["collectionname"].ToString();
            string primaryFieldName = command["primaryfieldname"].ToString();

            EntityMetadata entity = new EntityMetadata();
            entity.SchemaName = schemaName;
            entity.DisplayName = new Label(displayName, 1033);
            entity.DisplayCollectionName = new Label(collectionName, 1033);
            //Description = new Label("", 1033);
            entity.OwnershipType = OwnershipTypes.UserOwned;
            entity.IsActivity = false;

            this.Create(entity, primaryFieldName, context);
        }

        public void Create(EntityMetadata entity, string primaryFieldName, Context context)
        {
            if (entity.SchemaName == null) { return; }

            var cdsConnection = (CDSConnection)context.GetConnection();
            var em = cdsConnection.GetEntityMetadata(entity.SchemaName);

            if (em == null)
            {
                var createRequest = new CreateEntityRequest();
                createRequest.Entity = entity;

                StringAttributeMetadata atr = new StringAttributeMetadata();
                atr.SchemaName = primaryFieldName;
                atr.RequiredLevel = new AttributeRequiredLevelManagedProperty(AttributeRequiredLevel.None);
                atr.MaxLength = 100;
                atr.FormatName = StringFormatName.Text;
                atr.DisplayName = new Label("Name", 1033);
                createRequest.PrimaryAttribute = atr;

                Console.WriteLine("Creating entity...");
                cdsConnection.Execute(createRequest);
                Console.WriteLine("Entity created.");
            }
        }
    }
}

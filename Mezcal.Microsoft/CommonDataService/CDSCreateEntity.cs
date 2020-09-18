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
using Mezcal.Connections;
using System.Net.PeerToPeer.Collaboration;

namespace Mezcal.Microsoft.CommonDataService
{
    public class CDSCreateEntity : ICommand
    {
        private CDSConnection _cdsConnection;

        public JObject Prompt(CommandEngine commandEngine)
        {
            return null;
        }

        public void Process(JObject command, Context context)
        {
            var config = command["config"];
            if (config == null) { return; }
            var joConfig = context.FindItemByName(config.ToString());

            this._cdsConnection = new CDSConnection(joConfig);
            if (this._cdsConnection == null) { return; }

            this.CreateEntity(command, context);       
        }

        public void CreateEntity(JObject command, Context context)
        {
            var schemaName = JSONUtil.GetText(command, "#cds-entity");
            if (schemaName == null) { schemaName = JSONUtil.GetText(command, "schemaname"); }
            if (schemaName == null) { return; }

            var joEntity = context.FindItemByName(schemaName);
            if (joEntity == null) { return; }

            string displayName = joEntity["displayname"].ToString();
            string collectionName = joEntity["collectionname"].ToString();
            string primaryFieldName = joEntity["primaryfieldname"].ToString();

            EntityMetadata entity = new EntityMetadata();
            entity.SchemaName = schemaName;
            entity.DisplayName = new Label(displayName, 1033);
            entity.DisplayCollectionName = new Label(collectionName, 1033);
            //Description = new Label("", 1033);
            entity.OwnershipType = OwnershipTypes.UserOwned;
            entity.IsActivity = false;

            this.Create(entity, primaryFieldName);
        }

        public void Create(EntityMetadata entity, string primaryFieldName)
        {
            if (entity.SchemaName == null) { return; }

            var em = this._cdsConnection.GetEntityMetadata(entity.SchemaName);

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
                this._cdsConnection.Execute(createRequest);
                Console.WriteLine("Entity created.");
            }
        }
    }
}

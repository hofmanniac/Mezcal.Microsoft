using Mezcal.Commands;
using Mezcal.Connections;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Mezcal.Microsoft.CommonDataService
{
    public class CDSCreateField: ICommand
    {
        private CDSConnection _cdsConnection;

        public void Process(JObject command, Context context)
        {
            this._cdsConnection = CDSConnection.FromCommand(command, context);
            if (this._cdsConnection == null) { return; }

            this.CreateField(command);
        }

        public void SetConnection(CDSConnection cdsConnection)
        {
            this._cdsConnection = cdsConnection;
        }

        public JObject Prompt(CommandEngine commandEngine)
        {
            throw new NotImplementedException();
        }

        public void CreateField(JToken field)
        {
            var schemaName = JSONUtil.GetText(field, "#cds-field");
            if (schemaName == null) { schemaName = JSONUtil.GetText(field, "schemaname"); }
            if (schemaName == null) { return; }

            var displayName = field["displayname"];
            var fieldType = field["fieldtype"].ToString();

            Console.WriteLine("Creating field " + displayName.ToString());
            if (fieldType == "text" || fieldType == "TextType") { this.CreateTextField(field); }
            else if (fieldType == "money") { this.CreateMoneyField(field); }
            else if (fieldType == "lookup") { this.CreateLookupField(field); }
            else if (fieldType == "datetime") { this.CreateDateTimeField(field); }
            else if (fieldType == "boolean") { this.CreateBooleanField(field); }
            else if (fieldType == "integer") { this.CreateIntegerField(field); }
            else if (fieldType == "optionset") { this.CreateOptionSetField(field); }
        }

        private void CreateOptionSetField(JToken field)
        {
            CreateAttributeRequest req = new CreateAttributeRequest();
            req.EntityName = field["entity"].ToString();

            var am = new PicklistAttributeMetadata();
            am.SchemaName = field["schemaname"].ToString();
            am.RequiredLevel = new AttributeRequiredLevelManagedProperty(AttributeRequiredLevel.None);
            am.DisplayName = new Label(field["displayname"].ToString(), 1033);
            am.Description = new Label("", 1033);

            OptionSetMetadata os = new OptionSetMetadata();
            os.IsGlobal = false;
            foreach (var option in field["options"])
            {
                Label label = new Label(option["displayname"].ToString(), 1033);
                int? value = JSONUtil.GetInt32(option, "value");
                os.Options.Add(new OptionMetadata(label, value));
            }
            am.OptionSet = os;

            req.Attribute = am;

            this._cdsConnection.Execute(req);
        }

        private void CreateTextField(JToken field)
        {
            var entitySchemaName = JSONUtil.GetText(field, "entity");
            var displayName = JSONUtil.GetText(field, "displayname");
            var fieldSchemaName = JSONUtil.GetText(field, "schemaname");

            var req = new CreateAttributeRequest();
            req.EntityName = entitySchemaName;

            var format = JSONUtil.GetText(field, "format");
            if (format == null) { format = "single"; }

            int? maxlength = JSONUtil.GetInt32(field, "maxlength");

            if (format == "single")
            {
                var am = new StringAttributeMetadata();
                am.SchemaName = field["schemaname"].ToString();
                am.RequiredLevel = new AttributeRequiredLevelManagedProperty(AttributeRequiredLevel.None);

                maxlength = maxlength == null ? 100 : maxlength;
                am.MaxLength = maxlength;

                am.FormatName = StringFormatName.Text;
                am.DisplayName = new Label(displayName, 1033);
                am.Description = new Label("", 1033);
                req.Attribute = am;
            }
            else if (format == "multi")
            {
                var am = new MemoAttributeMetadata();
                am.SchemaName = fieldSchemaName;
                am.RequiredLevel = new AttributeRequiredLevelManagedProperty(AttributeRequiredLevel.None);

                maxlength = maxlength == null ? 2000 : maxlength;
                am.MaxLength = maxlength;

                am.DisplayName = new Label(displayName, 1033);
                am.Description = new Label("", 1033);
                req.Attribute = am;
            }

            this._cdsConnection.Execute(req);
        }

        private void CreateIntegerField(JToken field)
        {
            var entitySchemaName = JSONUtil.GetText(field, "entity");
            var displayName = JSONUtil.GetText(field, "displayname");
            var fieldSchemaName = JSONUtil.GetText(field, "schemaname");

            CreateAttributeRequest req = new CreateAttributeRequest
            {
                EntityName = entitySchemaName,
                Attribute = new IntegerAttributeMetadata
                {
                    SchemaName = fieldSchemaName,
                    RequiredLevel = new AttributeRequiredLevelManagedProperty(AttributeRequiredLevel.None),
                    //MaxValue = field.MaxValue == null ? null : field.MaxValue,
                    //MinValue = field.MinValue == null ? null : field.MinValue,
                    DisplayName = new Label(displayName, 1033),
                    Description = new Label("", 1033)
                }
            };

            this._cdsConnection.Execute(req);
        }

        private void CreateMoneyField(JToken field)
        {
            var entitySchemaName = JSONUtil.GetText(field, "entity");
            var displayName = JSONUtil.GetText(field, "displayname");
            var fieldSchemaName = JSONUtil.GetText(field, "schemaname");

            CreateAttributeRequest req = new CreateAttributeRequest
            {
                EntityName = entitySchemaName,
                Attribute = new MoneyAttributeMetadata
                {
                    SchemaName = fieldSchemaName,
                    RequiredLevel = new AttributeRequiredLevelManagedProperty(AttributeRequiredLevel.None),
                    PrecisionSource = 2,
                    DisplayName = new Label(displayName, 1033),
                    Description = new Label("", 1033),

                }
            };

            this._cdsConnection.Execute(req);
        }

        private void CreateBooleanField(JToken field)
        {
            var entitySchemaName = JSONUtil.GetText(field, "entity");
            var displayName = JSONUtil.GetText(field, "displayname");
            var fieldSchemaName = JSONUtil.GetText(field, "schemaname");

            CreateAttributeRequest req = new CreateAttributeRequest();
            req.EntityName = entitySchemaName;

            var am = new BooleanAttributeMetadata();
            am.SchemaName = fieldSchemaName;
            am.RequiredLevel = new AttributeRequiredLevelManagedProperty(AttributeRequiredLevel.None);
            am.DisplayName = new Label(displayName, 1033);
            am.Description = new Label("", 1033);

            am.OptionSet = new BooleanOptionSetMetadata(
                new OptionMetadata(new Label("Yes", 1033), 1),
                new OptionMetadata(new Label("No", 1033), 0));

            req.Attribute = am;

            this._cdsConnection.Execute(req);
        }

        private void CreateDateTimeField(JToken field)
        {
            var entitySchemaName = JSONUtil.GetText(field, "entity");
            var displayName = JSONUtil.GetText(field, "displayname");
            var fieldSchemaName = JSONUtil.GetText(field, "schemaname");
            
            CreateAttributeRequest req = new CreateAttributeRequest();
            req.EntityName = entitySchemaName;
            var dta = new DateTimeAttributeMetadata();
            dta.SchemaName = fieldSchemaName;
            dta.RequiredLevel = new AttributeRequiredLevelManagedProperty(AttributeRequiredLevel.None);

            var datetimeoption = JSONUtil.GetText(field, "datetimeoption");
            if (datetimeoption == null) { datetimeoption = "dateonly"; }
            if (datetimeoption == "dateonly") { dta.Format = DateTimeFormat.DateOnly; }
            else if (datetimeoption == "datetime") { dta.Format = DateTimeFormat.DateAndTime; }

            dta.DisplayName = new Label(displayName, 1033);
            dta.Description = new Label("", 1033);
            req.Attribute = dta;

            this._cdsConnection.Execute(req);
        }

        /// <summary>
        /// Create the LookupField in CDS 
        /// </summary>
        /// <param name="entity">
        /// Uses: Entity.CollectionName
        /// </param>
        /// <param name="field">
        /// Uses: LookupField.SchemaName, .LookupToEntity, .LookupToField
        /// </param>
        public void CreateLookupField(JToken field)
        {
            var entitySchemaName = JSONUtil.GetText(field, "entity");
            var displayName = JSONUtil.GetText(field, "displayname");
            var fieldSchemaName = JSONUtil.GetText(field, "schemaname");

            var targetentity = JSONUtil.GetText(field, "target-entity");
            var targetfield = JSONUtil.GetText(field, "target-field");

            var relationshipname = JSONUtil.GetText(field, "relname");

            var em = this._cdsConnection.GetEntityMetadata(entitySchemaName);

            CreateOneToManyRequest req = new CreateOneToManyRequest();

            // define the general lookup metadata
            var la = new LookupAttributeMetadata();
            la.Description = new Label("", 1033);
            la.DisplayName = new Label(displayName, 1033);
            la.LogicalName = fieldSchemaName;
            la.SchemaName = fieldSchemaName;
            la.RequiredLevel = new AttributeRequiredLevelManagedProperty(
                AttributeRequiredLevel.Recommended);
            req.Lookup = la;

            // define the 1:N relationship
            var rel = new OneToManyRelationshipMetadata();

            // 1:N associated menu config
            var amc = new AssociatedMenuConfiguration();
            amc.Behavior = AssociatedMenuBehavior.UseCollectionName;
            amc.Group = AssociatedMenuGroup.Details;
            amc.Label = em.DisplayCollectionName;
            amc.Order = 10000;
            rel.AssociatedMenuConfiguration = amc;

            // 1:N cascade behavior config
            var cc = new CascadeConfiguration();
            cc.Assign = CascadeType.NoCascade;
            cc.Delete = CascadeType.RemoveLink;
            cc.Merge = CascadeType.NoCascade;
            cc.Reparent = CascadeType.NoCascade;
            cc.Share = CascadeType.NoCascade;
            cc.Unshare = CascadeType.NoCascade;
            rel.CascadeConfiguration = cc;

            // 1:N entity reference
            rel.ReferencedEntity = targetentity;
            rel.ReferencedAttribute = targetfield;
            rel.ReferencingEntity = entitySchemaName;

            if (relationshipname == null) { relationshipname = this.GetNextRelationshipName(em, field); }
            rel.SchemaName = relationshipname;

            req.OneToManyRelationship = rel;

            this._cdsConnection.Execute(req);
        }

        private string GetNextRelationshipName(EntityMetadata em, JToken specField)
        {
            string result = null;

            var targetentity = JSONUtil.GetText(specField, "target-entity");

            int num = 0;
            while (num < 1000) // hopefully will find a unique name after 1000 times! :)
            {
                num++;
                string candidateName = String.Format("{0}_{1}_{2}", targetentity, em.SchemaName, num);
                var rel = em.ManyToOneRelationships.FirstOrDefault(r => r.SchemaName == candidateName);
                if (rel == null) { result = candidateName; break; }
            }

            return result;
        }

    }
}

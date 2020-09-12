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
        private Context _context;
        private CDSConnection _cdsConnection;

        public void Process(JObject command, Context context)
        {
            this._context = context;

            this._cdsConnection = CDSConnection.FromCommand(command, context);
            if (this._cdsConnection == null) { return; }
        }

        public JObject Prompt(CommandEngine commandEngine)
        {
            throw new NotImplementedException();
        }

        /*
        private void CreateField(JToken entity, JToken field)
        {
            var displayName = entity["displayname"];
            var fieldType = entity["fieldtype"].ToString();

            Console.WriteLine("Creating field " + displayName.ToString());
            if (fieldType == "StringType" || fieldType == "TextType") { this.CreateTextField(entity, field); }
            else if (fieldType == "MoneyType") { this.CreateMoneyField(entity, field); }
            else if (fieldType == "LookupType") { this.CreateLookupField(entity, field); }
            else if (fieldType == "DateTimeType") { this.CreateDateTimeField(entity, field); }
            else if (fieldType == "BooleanType") { this.CreateBooleanField(entity, field); }
            else if (fieldType == "IntegerType") { this.CreateIntegerField(entity, field); }
            else if (fieldType == "OptionSetType") { this.CreateOptionSetField(entity, field); }
        }

        private void CreateOptionSetField(JToken cdsEntity, JToken field)
        {
            CreateAttributeRequest req = new CreateAttributeRequest();
            req.EntityName = cdsEntity["schemaname"].ToString();

            var am = new PicklistAttributeMetadata();
            am.SchemaName = field["schemaname"].ToString();
            am.RequiredLevel = new AttributeRequiredLevelManagedProperty(AttributeRequiredLevel.None);
            am.DisplayName = new Label(field["displayName"].ToString(), 1033);
            am.Description = new Label("", 1033);

            OptionSetMetadata os = new OptionSetMetadata();
            os.IsGlobal = false;

            foreach (var option in field["options"])
            {
                Label label = new Label(option["displayName"].ToString(), 1033);
                int? value = JSONUtil.GetInt32(option, "value");
                os.Options.Add(new OptionMetadata(label, value));
            }
            am.OptionSet = os;

            req.Attribute = am;

            this._cdsConnection.Execute(req);
        }

        private void CreateTextField(JToken cdsEntity, JToken field)
        {
            var entitySchemaName = JSONUtil.GetText(cdsEntity, "schemaname");
            var displayName = JSONUtil.GetText(field, "displayname");
            var fieldSchemaName = JSONUtil.GetText(field, "schemaname");

            var req = new CreateAttributeRequest();
            req.EntityName = entitySchemaName;

            if (field["format"].ToString() == "single")
            {
                var am = new StringAttributeMetadata();
                am.SchemaName = field["schemaname"].ToString();
                am.RequiredLevel = new AttributeRequiredLevelManagedProperty(AttributeRequiredLevel.None);

                int? maxlength = JSONUtil.GetInt32(field, "maxlength");
                am.MaxLength = maxlength == 0 ? 100 : maxlength;
                
                
                am.FormatName = StringFormatName.Text;
                am.DisplayName = new Label(displayName, 1033);
                am.Description = new Label("", 1033);
                req.Attribute = am;
            }
            else if (field["format"].ToString() == "multi")
            {
                var am = new MemoAttributeMetadata();
                am.SchemaName = fieldSchemaName;
                am.RequiredLevel = new AttributeRequiredLevelManagedProperty(AttributeRequiredLevel.None);
                am.MaxLength = field.MaxLength == 0 ? 100 : field.MaxLength;
                am.DisplayName = new Label(displayName, 1033);
                am.Description = new Label("", 1033);
                req.Attribute = am;
            }

            this._cdsConnection.Execute(req);
        }

        private void CreateIntegerField(JToken cdsEntity, JToken field)
        {
            CreateAttributeRequest req = new CreateAttributeRequest
            {
                EntityName = cdsEntity.Entity.SchemaName,
                Attribute = new IntegerAttributeMetadata
                {
                    SchemaName = field.SchemaName,
                    RequiredLevel = new AttributeRequiredLevelManagedProperty(AttributeRequiredLevel.None),
                    MaxValue = field.MaxValue == null ? null : field.MaxValue,
                    MinValue = field.MinValue == null ? null : field.MinValue,
                    DisplayName = new CDS.Label(field.DisplayName, 1033),
                    Description = new CDS.Label("", 1033)
                }
            };

            this._cdsConnection.Execute(req);
        }

        private void CreateMoneyField(JToken cdsEntity, JToken field)
        {
            CreateAttributeRequest req = new CreateAttributeRequest
            {
                EntityName = cdsEntity.Entity.SchemaName,
                Attribute = new MoneyAttributeMetadata
                {
                    SchemaName = field.SchemaName,
                    RequiredLevel = new AttributeRequiredLevelManagedProperty(AttributeRequiredLevel.None),
                    PrecisionSource = 2,
                    DisplayName = new CDS.Label(field.DisplayName, 1033),
                    Description = new CDS.Label("", 1033),

                }
            };

            this._cdsConnection.Execute(req);
        }

        private void CreateBooleanField(JToken cdsEntity, JToken field)
        {
            CreateAttributeRequest req = new CreateAttributeRequest();
            req.EntityName = cdsEntity.Entity.SchemaName;

            var am = new BooleanAttributeMetadata();
            am.SchemaName = field.SchemaName;
            am.RequiredLevel = new AttributeRequiredLevelManagedProperty(AttributeRequiredLevel.None);
            am.DisplayName = new CDS.Label(field.DisplayName, 1033);
            am.Description = new CDS.Label("", 1033);

            am.OptionSet = new BooleanOptionSetMetadata(
                new OptionMetadata(new CDS.Label("Yes", 1033), 1),
                new OptionMetadata(new CDS.Label("No", 1033), 0));

            req.Attribute = am;

            this._cdsConnection.Execute(req);
        }

        private void CreateDateTimeField(JToken cdsEntity, JToken field)
        {
            CreateAttributeRequest req = new CreateAttributeRequest();
            req.EntityName = cdsEntity.Entity.SchemaName;
            var dta = new DateTimeAttributeMetadata();
            dta.SchemaName = field.SchemaName;
            dta.RequiredLevel = new AttributeRequiredLevelManagedProperty(AttributeRequiredLevel.None);

            if (field.DateTimeOption == DateTimeOption.DateOnly) { dta.Format = DateTimeFormat.DateOnly; }
            else if (field.DateTimeOption == DateTimeOption.DateAndTime) { dta.Format = DateTimeFormat.DateAndTime; }

            dta.DisplayName = new CDS.Label(field.DisplayName, 1033);
            dta.Description = new CDS.Label("", 1033);
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
        public void CreateLookupField(JToken cdsEntity, JToken field)
        {
            CreateOneToManyRequest req = new CreateOneToManyRequest();

            // define the general lookup metadata
            var la = new LookupAttributeMetadata();
            la.Description = new CDS.Label("", 1033);
            la.DisplayName = new CDS.Label(field.DisplayName, 1033);
            la.LogicalName = field.SchemaName;
            la.SchemaName = field.SchemaName;
            la.RequiredLevel = new AttributeRequiredLevelManagedProperty(
                AttributeRequiredLevel.Recommended);
            req.Lookup = la;

            // define the 1:N relationship
            var rel = new OneToManyRelationshipMetadata();

            // 1:N associated menu config
            var amc = new AssociatedMenuConfiguration();
            amc.Behavior = AssociatedMenuBehavior.UseCollectionName;
            amc.Group = AssociatedMenuGroup.Details;
            amc.Label = new CDS.Label(cdsEntity.Entity.CollectionName, 1033);
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
            rel.ReferencedEntity = field.LookupToEntity;
            rel.ReferencedAttribute = field.LookupToField;
            rel.ReferencingEntity = cdsEntity.Entity.SchemaName;

            string relName = null;
            if (field.RelationshipName != null) { relName = field.RelationshipName; }
            //else { relName = entity.GetNextRelationshipName(field.LookupToEntity); }
            else { relName = this.GetNextRelationshipName(cdsEntity.Entity, field); }

            rel.SchemaName = relName;

            req.OneToManyRelationship = rel;

            this._cdsConnection.Execute(req);
        }

        private string GetNextRelationshipName(JToken entity, JToken specField)
        {
            string result = null;

            this.Connect();

            var em = this.GetEntityMetadata(entity.SchemaName);

            int num = 0;
            while (num < 1000) // hopefully will find a unique name after 1000 times! :)
            {
                num++;
                string candidateName = String.Format("{0}_{1}_{2}", specField.LookupToEntity, entity.SchemaName, num);
                var rel = em.ManyToOneRelationships.FirstOrDefault(r => r.SchemaName == candidateName);
                if (rel == null) { result = candidateName; break; }
            }

            return result;
        }

        */
    }
}

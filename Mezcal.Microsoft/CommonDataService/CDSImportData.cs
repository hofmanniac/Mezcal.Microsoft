using Mezcal.Commands;
using Mezcal.Connections;
using Microsoft.Office.Interop.Word;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Metadata;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Linq;

namespace Mezcal.Microsoft.CommonDataService
{
    public class CDSImportData : ICommand
    {
        //private Context _context = null;

        public JObject Prompt(CommandEngine commandEngine)
        {
            return null;
        }

        public void Process(JObject command, Context context)
        {
            string entity = command["entity"].ToString();
            string source = command["source"].ToString();
            JArray map = (JArray)command["map"];

            CDSConnection cdsConnection = CDSConnection.FromCommand(command, context);
            if (cdsConnection == null) { return; }

            JArray set = (JArray)context.Fetch(source);
            //this._context = context;

            var entities = CDSConnection.ConvertToCDSEntities(entity, set, map, cdsConnection);
            this.CreateRecords(entities, cdsConnection);
        }

        public void CreateRecords(List<Entity> cdsEntities, CDSConnection cdsConnection)
        {
            var num = cdsEntities.Count;
            var i = 0;

            foreach (var item in cdsEntities)
            {
                //Entity cdsEntity = new Entity(item.LogicalName);
                i++;

                Console.Write($"Adding record {i} of {num}...");
                cdsConnection.Create(item);
                Console.WriteLine("Added." + JSONUtil.SingleLine(JArray.FromObject(item.Attributes)));
            }             
        }

        //public void LoadData(string entityName, JArray data, JArray map, CDSConnection cdsConnection)
        //{
        //    var em = cdsConnection.GetEntityMetadata(entityName);
        //    if (em == null)
        //    {
        //        Console.WriteLine("Unable to load entity metadata for " + entityName + " - skipping.");
        //        return;
        //    }

        //    foreach (var item in data)
        //    {
        //        Entity cdsEntity = new Entity(entityName);
        //        var record = (JObject)item;

        //        foreach (var field in record)
        //        {
        //            string targetName = field.Key;

        //            if (map != null)
        //            {
        //                var mapEntry = map.FirstOrDefault(m => m["source"].ToString() == field.Key);
        //                if (mapEntry == null) { continue; }
        //                var target = mapEntry["target"];
        //                if (target != null) { targetName = target.ToString(); }
        //                else { targetName = mapEntry["source"].ToString(); }
        //            }

        //            if (targetName == null) { continue; }

        //            var atr = em.Attributes.FirstOrDefault(a => a.LogicalName == targetName);
        //            if (atr == null) { continue; }

        //            JToken valToken = record[field.Key];
        //            if (valToken == null) { continue; }

        //            if (atr.AttributeTypeName.Value == "LookupType")
        //            {
        //                string lookupTo = null;
        //                string id = null;

        //                if (valToken.HasValues == false)
        //                {
        //                    var erAtr = (LookupAttributeMetadata)atr;
        //                    lookupTo = erAtr.Targets[0];
        //                    id = valToken.ToString();
        //                }
        //                //else
        //                //{
        //                //    var jLookupTo = valToken["lookupto"];
        //                //    if (jLookupTo != null)
        //                //    {
        //                //        lookupTo = jLookupTo.ToString();
        //                //        id = valToken[lookupTo + "id"].ToString();                        
        //                //    }
        //                //}

        //                if (lookupTo != null && id != null)
        //                {
        //                    EntityReference er = new EntityReference(lookupTo);
        //                    er.Id = new Guid(id);

        //                    cdsEntity.Attributes.Add(targetName, er);
        //                }
        //            }
        //            else if (atr.AttributeTypeName.Value == "PicklistType")
        //            {
        //                string val = null;

        //                if (valToken.HasValues == false) { val = valToken.ToString(); }
        //                else { val = valToken["value"].ToString(); }

        //                int numval = Int32.Parse(val);
        //                var osv = new OptionSetValue(numval);

        //                cdsEntity.Attributes.Add(targetName, osv);
        //            }
        //            else if (atr.AttributeTypeName.Value == "IntegerType")
        //            {
        //                string val = valToken.ToString();
        //                int numval = Int32.Parse(val);
        //                //var osv = new CDS.OptionSetValue(numval);

        //                cdsEntity.Attributes.Add(targetName, numval);
        //            }
        //            else if (atr.AttributeTypeName.Value == "DateTimeType")
        //            {
        //                string val = valToken.ToString();
        //                var valDateTime = DateTime.Parse(val);
        //                cdsEntity.Attributes.Add(targetName, valDateTime);
        //            }
        //            else if (atr.AttributeTypeName.Value == "BooleanType")
        //            {
        //                string val = valToken.ToString();
        //                var valBoolean = Boolean.Parse(val);
        //                cdsEntity.Attributes.Add(targetName, valBoolean);
        //            }
        //            else if (atr.AttributeTypeName.Value == "MoneyType")
        //            {
        //                string val = valToken.ToString();
        //                var valMoney = new Money(Decimal.Parse(val));
        //                cdsEntity.Attributes.Add(targetName, valMoney);
        //            }
        //            else if (atr.AttributeTypeName.Value == "UniqueidentifierType")
        //            {
        //                Guid val = new Guid(valToken.ToString());
        //                cdsEntity.Attributes.Add(targetName, val);
        //            }
        //            else
        //            {
        //                string val = valToken.ToString();
        //                cdsEntity.Attributes.Add(targetName, val);
        //            }

        //        }

        //        Console.Write("Adding record...");
        //        cdsConnection.Create(cdsEntity);
        //        Console.WriteLine("Added.");
        //    }
        //}
    }
}

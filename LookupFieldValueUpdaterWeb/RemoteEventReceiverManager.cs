using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.ServiceModel;
using System.ServiceModel.Channels;

namespace LookupFieldValueUpdaterWeb.Services
{
    public class RemoteEventReceiverManager
    {
        private const string RECEIVER_NAME_ADDED = "LookupFieldValueUpdater";
        private const string LIST_TITLE = "documents";

        public void AssociateRemoteEventsToHostWeb(ClientContext clientContext)
        {

            var rerList = clientContext.Web.Lists.GetByTitle(LIST_TITLE);
            clientContext.Load(rerList);
            clientContext.ExecuteQuery();

            bool rerExists = false;
            if (!rerExists)
            {
                //Get WCF URL where this message was handled
                OperationContext op = OperationContext.Current;
                Message msg = op.RequestContext.RequestMessage;

                EventReceiverDefinitionCreationInformation receiver = new EventReceiverDefinitionCreationInformation();

                receiver.EventType = EventReceiverType.ItemAdded;
                receiver.ReceiverUrl = msg.Headers.To.ToString();
                receiver.ReceiverName = EventReceiverType.ItemAdded.ToString();
                receiver.Synchronization = EventReceiverSynchronization.Synchronous;
                receiver.SequenceNumber = 400;
                rerList.EventReceivers.Add(receiver);
                clientContext.ExecuteQuery();
                System.Diagnostics.Trace.WriteLine("Added ItemAdded receiver at " + receiver.ReceiverUrl);

                receiver = new EventReceiverDefinitionCreationInformation();
                receiver.EventType = EventReceiverType.ItemUpdated;
                receiver.ReceiverUrl = msg.Headers.To.ToString();
                receiver.ReceiverName = EventReceiverType.ItemUpdated.ToString();
                receiver.Synchronization = EventReceiverSynchronization.Synchronous;
                receiver.SequenceNumber = 400;
                rerList.EventReceivers.Add(receiver);
                clientContext.ExecuteQuery();

                System.Diagnostics.Trace.WriteLine("Added ItemUpdated receiver at " + receiver.ReceiverUrl);
            }
        }

        public void RemoveEventReceiversFromHostWeb(ClientContext clientContext)
        {
            List myList = clientContext.Web.Lists.GetByTitle(LIST_TITLE);
            clientContext.Load(myList, p => p.EventReceivers);
            clientContext.ExecuteQuery();

            var rer = myList.EventReceivers.Where(e => e.ReceiverName == RECEIVER_NAME_ADDED).FirstOrDefault();

            try
            {
                System.Diagnostics.Trace.WriteLine("Removing receiver at "
                        + rer.ReceiverUrl);

                var rerList = myList.EventReceivers.Where(e => e.ReceiverUrl == rer.ReceiverUrl).ToList<EventReceiverDefinition>();

                foreach (var rerFromUrl in rerList)
                {
                    //This will fail when deploying via F5, but works
                    //when deployed to production
                    rerFromUrl.DeleteObject();
                }
                clientContext.ExecuteQuery();
            }
            catch (Exception oops)
            {
                System.Diagnostics.Trace.WriteLine(oops.Message);
            }

            clientContext.ExecuteQuery();
        }

        public void ItemUpdatedToListEventHandler(ClientContext clientContext, Guid listId, int listItemId)
        {
            try
            {
                UpdateDescription(clientContext, listId, listItemId);
            }
            catch (Exception oops)
            {
                System.Diagnostics.Trace.WriteLine(oops.Message);
            }
        }

        public void ItemAddedToListEventHandler(ClientContext clientContext, Guid listId, int listItemId)
        {
            try
            {
                UpdateDescription(clientContext, listId, listItemId);
            }
            catch (Exception oops)
            {
                System.Diagnostics.Trace.WriteLine(oops.Message);
            }
        }

        public void ItemAddedToListEventHandlerAsync(ClientContext clientContext, Guid listId, int listItemId)
        {
            try
            {
                //UpdateDescription(clientContext, listId, listItemId);
            }
            catch (Exception oops)
            {
                System.Diagnostics.Trace.WriteLine(oops.Message);
            }
        }

        void UpdateDescription(ClientContext clientContext, Guid listId, int listItemId)
        {
            List list = clientContext.Web.Lists.GetById(listId);
            ListItem item = list.GetItemById(listItemId);
            ContentType itemContentType = item.ContentType;
            clientContext.Load(item);
            clientContext.Load(itemContentType);
            FieldCollection fields = list.Fields;
            clientContext.Load(fields);
            clientContext.ExecuteQuery();

            if (!item.ContentType.Name.Equals("Rapid Delivery Document Set NL"))
                return;

            List<LookupFieldSet> lookupDefinitions = new List<LookupFieldSet>();

            LookupFieldSet lookupFieldDefinition = new LookupFieldSet();
            lookupFieldDefinition.LookupField = "BusinessConsultant";
            lookupFieldDefinition.FieldMappings.Add(new LookupFieldMapping("BusinessConsultantFirstName", "FirstName"));
            lookupFieldDefinition.FieldMappings.Add(new LookupFieldMapping("BusinessConsultantLastName", "Title"));
            lookupFieldDefinition.FieldMappings.Add(new LookupFieldMapping("BusinessConsultantEmail", "Email"));
            lookupFieldDefinition.FieldMappings.Add(new LookupFieldMapping("BusinessConsultantMobile", "CellPhone"));
            lookupDefinitions.Add(lookupFieldDefinition);

            lookupFieldDefinition = new LookupFieldSet();
            lookupFieldDefinition.LookupField = "FunctionalConsultant";
            lookupFieldDefinition.FieldMappings.Add(new LookupFieldMapping("FunctionalConsultantFirstName", "FirstName"));
            lookupFieldDefinition.FieldMappings.Add(new LookupFieldMapping("FunctionalConsultantLastName", "Title"));
            lookupFieldDefinition.FieldMappings.Add(new LookupFieldMapping("FunctionalConsultantEmail", "Email"));
            lookupFieldDefinition.FieldMappings.Add(new LookupFieldMapping("FunctionalConsultantMobile", "CellPhone"));
            lookupDefinitions.Add(lookupFieldDefinition);

            lookupFieldDefinition = new LookupFieldSet();
            lookupFieldDefinition.LookupField = "ProjectManager";
            lookupFieldDefinition.FieldMappings.Add(new LookupFieldMapping("ProjectManagerFirstName", "FirstName"));
            lookupFieldDefinition.FieldMappings.Add(new LookupFieldMapping("ProjectManagerLastName", "Title"));
            lookupFieldDefinition.FieldMappings.Add(new LookupFieldMapping("ProjectManagerEmail", "Email"));
            lookupFieldDefinition.FieldMappings.Add(new LookupFieldMapping("ProjectManagerMobile", "CellPhone"));
            lookupDefinitions.Add(lookupFieldDefinition);

            Web lookupListWeb = null;
            List parentList = null;

            Dictionary<string, string> itemValues = new Dictionary<string, string>();
            foreach (var lookupDefinition in lookupDefinitions)
            {
                if (item[lookupDefinition.LookupField] == null) continue;
                FieldLookupValue itemField = item[lookupDefinition.LookupField] as FieldLookupValue;
                string lookupFieldValue = itemField.LookupValue;
                int lookupFieldId = itemField.LookupId;

                FieldLookup lookupField = clientContext.CastTo<FieldLookup>(fields.GetByInternalNameOrTitle(lookupDefinition.LookupField));
                clientContext.Load(lookupField);
                clientContext.ExecuteQuery();

                Guid parentWeb = lookupField.LookupWebId;
                string parentListGUID = lookupField.LookupList;

                if (lookupListWeb == null) lookupListWeb = clientContext.Site.OpenWebById(parentWeb);
                if (parentList == null)
                {
                    parentList = lookupListWeb.Lists.GetById(new Guid(parentListGUID));
                    clientContext.Load(parentList);
                    clientContext.ExecuteQuery();
                }

                ListItem lookupListItem = parentList.GetItemById(lookupFieldId);
                clientContext.Load(lookupListItem);
                clientContext.ExecuteQuery();

                foreach(var fieldMap in lookupDefinition.FieldMappings)
                {
                    if (item[fieldMap.ItemField] == null || !item[fieldMap.ItemField].ToString().Equals(lookupListItem[fieldMap.LookupListField].ToString()))
                    {
                        itemValues.Add(fieldMap.ItemField, lookupListItem[fieldMap.LookupListField].ToString());
                    }
                }
            }

            if (!itemValues.Any())
                return;

            foreach(KeyValuePair<string,string> itemValue in itemValues)
            {
                item[itemValue.Key] = itemValue.Value;
            }
            item.Update();
            clientContext.ExecuteQuery();
        }

        bool FieldExists(Dictionary<string, object> FieldValues, string Field)
        {
            foreach (var field in FieldValues)
            {
                if (field.Key.Equals(Field))
                    return true;
            }
            return false;
        }
    }

    public class LookupFieldSet
    {
        public string LookupField { get; set; }
        public List<LookupFieldMapping> FieldMappings { get; set; }

        public LookupFieldSet()
        {
            FieldMappings = new List<LookupFieldMapping>();
        }

    }

    public class LookupFieldMapping
    {
        public string ItemField { get; set; }
        public string LookupListField { get; set; }

        public LookupFieldMapping(string localLookupField, string parentSourceField)
        {
            ItemField = localLookupField;
            LookupListField = parentSourceField;
        }
    }

    public static class StringExtensions
    {
        public static string Replace(this string originalString, string oldValue, string newValue, StringComparison comparisonType)
        {
            int startIndex = 0;
            while (true)
            {
                startIndex = originalString.IndexOf(oldValue, startIndex, comparisonType);
                if (startIndex == -1)
                    break;

                originalString = originalString.Substring(0, startIndex) + newValue + originalString.Substring(startIndex + oldValue.Length);

                startIndex += newValue.Length;
            }

            return originalString;
        }

    }
}
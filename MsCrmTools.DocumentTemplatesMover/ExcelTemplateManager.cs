using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using NuGet;
using McTools.Xrm.Connection;
using System.ComponentModel;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.ServiceModel;
using Microsoft.Xrm.Tooling.Connector;
using Microsoft.Xrm.Sdk.Metadata;

namespace MsCrmTools.DocumentTemplatesMover
{
    enum ViewType
    {
        Personal,
        System
    }

    class ViewMetadataInfo
    {
        public string LayoutXml { get; private set;}
        public string FetchXml { get; private set; }

        public string primaryFieldLogicalName { get; set; } 
        public string primaryKeyLogicalName { get; set; }

        public ViewMetadataInfo(string logicalName, int entityTypeCode, Dictionary<string, string> columns, List<Tuple<OneToManyRelationshipMetadata, List<string>>> attributesToFind)
        {
            primaryKeyLogicalName = columns.Keys.First();
            List<KeyValuePair<string, string>> columnsToAddToView = columns.Where(o => o.Key != "checksumLogicalName" && !o.Key.Contains(".")).ToList();
            var orderBy = $"<order attribute='{primaryFieldLogicalName}' descending='false'/>";
            var attributeXml = string.Join("", (from KeyValuePair<string,string> columnMapping  in columnsToAddToView
                                select $"<attribute name='{columnMapping.Key}'/>").ToList());

            var linkEntitiesXml = string.Join("",(from Tuple<OneToManyRelationshipMetadata, List<string>> relationShip in attributesToFind
                                  select $"<link-entity name='{relationShip.Item1.ReferencedEntity}' from='{relationShip.Item1.ReferencedAttribute}' to='{relationShip.Item1.ReferencingAttribute}' link-type='inner' alias='{relationShip.Item1.ReferencedEntity}'>{string.Join("", (relationShip.Item2.Select(o => $"<attribute name='{o}' />")))}</link-entity>").ToList());

            var cellXml = string.Join("", (from KeyValuePair<string, string> columnMapping in columnsToAddToView
                                           select $"<cell name='{columnMapping.Key}' width='208'/>").ToList());
            var linkEntitiesCellXml = string.Join("", (from Tuple<OneToManyRelationshipMetadata, List<string>> relationShip in attributesToFind
                                                       select $"{string.Join("", (relationShip.Item2.Select(o => $"<cell name='{relationShip.Item1.ReferencedEntity}.{o}' />")))}").ToList());

            //jump='{primaryFieldLogicalName}' attribute may not be needed
            FetchXml = $"<fetch version='1.0' mapping='logical' top='10'><entity name='{logicalName}'>{attributeXml}<filter type='and'><condition attribute='statecode' operator='eq' value='0'/></filter>{linkEntitiesXml}</entity></fetch>";
            LayoutXml = $"<grid name='resultset' object='{entityTypeCode}' select='1' icon='1' preview='1'><row name='result' id='{primaryKeyLogicalName}'>{cellXml}{linkEntitiesCellXml}</row></grid>";
        }

    }

    class ExcelTemplateManager
    {
        public ExcelTemplateManager() { }
        public ExcelTemplateManager(string name) { }
        private int destinationEntityTypeCode;

        public ExcelTemplateManager(ConnectionDetail connectionDetail1, ConnectionDetail connectionDetail2)
        {
            this.sourceServer = connectionDetail1;
            this.destinationServer = connectionDetail2;
            relationShipGuidMappings = new Dictionary<string, Guid>();
        }

        private BackgroundWorker bg;
        private PluginControl pluginControl;
        private ConnectionDetail destinationServer;
        private ConnectionDetail sourceServer;

        private Dictionary<string, Guid> relationShipGuidMappings;

        internal void Transform(Entity originalEntityTemplate, PluginControl pluginControl, BackgroundWorker worker)
        {
            bg = worker;
            pluginControl.LogInfo("Fetching Template");
            ExcelTemplate originalTemplate = null;
            try
            {
                originalTemplate = ExcelTemplate.FromEntityRecord(originalEntityTemplate);
            }
            catch(Exception ex)
            {
                pluginControl.LogError($"Failed to get originalEntityTemplate: {originalEntityTemplate.GetAttributeValue<string>("name")}. Reason: {ex.Message}");
                throw new Exception($"Failed to read original originalEntityTemplate: {originalEntityTemplate.GetAttributeValue<string>("name")}", ex);
            }

            ExcelTemplate destinationEnvTemplate = GenerateEmptyTemplateCopy(originalTemplate);

            originalTemplate.ReplaceColumnMappings(destinationEnvTemplate, relationShipGuidMappings);

            originalEntityTemplate = originalTemplate.ToEntity();
        }

        public int? GetEntityTypeCode(CrmServiceClient service, string entity)
        {
            RetrieveEntityRequest request = new RetrieveEntityRequest();

            request.LogicalName = entity;
            request.EntityFilters = EntityFilters.Entity;

            RetrieveEntityResponse response = (RetrieveEntityResponse)service.Execute(request);
            EntityMetadata metadata = response.EntityMetadata;

            return metadata.ObjectTypeCode;
        }

        private EntityMetadata GetEntityMetadata(CrmServiceClient service, ExcelTemplate template)
        {
            RetrieveEntityRequest request = new RetrieveEntityRequest();

            request.LogicalName = template.LogicalName;
            request.EntityFilters = EntityFilters.Relationships | EntityFilters.Attributes | EntityFilters.Entity;

            RetrieveEntityResponse response = (RetrieveEntityResponse)service.Execute(request);
            EntityMetadata metadata = response.EntityMetadata;

            return metadata;
        }

        private ExcelTemplate GenerateEmptyTemplateCopy(ExcelTemplate template)
        {
            ExcelTemplate defaultTemplate = null;
            int destinationEntityTypeCode = (int)GetEntityTypeCode(destinationServer.GetCrmServiceClient(), template.LogicalName);

            //retrieve template's primary entity information
            EntityMetadata metadata = GetEntityMetadata(destinationServer.GetCrmServiceClient(), template);

            var lookupAttributes = metadata.Attributes.Where(o => o.AttributeType == AttributeTypeCode.Lookup);
            //List of relationships and attributes to include from that relationship in the view
            List<Tuple<OneToManyRelationshipMetadata, List<string>>> attributesToFind = new List<Tuple<OneToManyRelationshipMetadata, List<string>>>();
            List<IGrouping<string,string>> relatedEntities = template.ColumnMappings.Keys.Where(o => o.Contains(".")).GroupBy(o => o.Substring(0, o.IndexOf("."))).ToList();
            foreach (IGrouping<string,string> grpAttributes in relatedEntities)
            {
                string relatedEntity = template.ColumnMappings.FirstOrDefault(o => o.Key.StartsWith(grpAttributes.Key)).Value;
                
                //Get lookup name in brackets
                string relatedEntityDisplayInfo = relatedEntity.Substring(relatedEntity.IndexOf("("));  
                string lookupDisplayName = relatedEntity.Substring(relatedEntity.IndexOf("(") + 1);
                lookupDisplayName = lookupDisplayName.Substring(0, lookupDisplayName.IndexOf(")"));

                //Get relationship metadata
                LookupAttributeMetadata referencingAttribute = null;
                try
                {
                    referencingAttribute = (LookupAttributeMetadata)lookupAttributes.Single(o => o.DisplayName.UserLocalizedLabel.Label == lookupDisplayName);
                }
                catch(Exception ex)
                {
                    throw new Exception($"Could not find lookup \"{lookupDisplayName}\"! It may have been renamed or deleted");
                }
                //string relatedEntityLogicalName = lookupAttributes.Single(o => o.DisplayName.UserLocalizedLabel.Label == relatedEntityDisplayInfo);
                OneToManyRelationshipMetadata relationShip = metadata.ManyToOneRelationships.First(o => o.ReferencedEntity == referencingAttribute.Targets.FirstOrDefault());
                List<string> columnsInRelationshipToFetch = new List<string>();
                columnsInRelationshipToFetch.AddRange(grpAttributes.Select(o => o.Substring(o.IndexOf('.')+1)));

                attributesToFind.Add(new Tuple<OneToManyRelationshipMetadata, List<string>>(relationShip, columnsInRelationshipToFetch));
            }
            
            ViewMetadataInfo info = new ViewMetadataInfo(template.LogicalName, destinationEntityTypeCode, template.ColumnMappings, attributesToFind);

            OrganizationRequest req = new OrganizationRequest("ExportTemplateToExcel");
            req.Parameters = new ParameterCollection()
            {
                ["EntityLocalizedDisplayName"] = metadata.DisplayName.UserLocalizedLabel.Label,
                ["FetchXml"] = info.FetchXml,
                ["LayoutXml"] = info.LayoutXml
            };

            try
            {
                OrganizationResponse resp = destinationServer.ServiceClient.Execute(req);
                if(resp != null && resp.Results.Count > 0 && resp.Results["ExcelFile"] != null)
                {
                    Stream newTemplateStream = new MemoryStream((byte[])resp.Results["ExcelFile"]);
                    defaultTemplate = ExcelTemplate.FromStream(newTemplateStream, template.LogicalName, template.TemplateName);
                }
            }
            catch(FaultException e)
            {
                throw new Exception($"Failed to create default originalEntityTemplate: {e.Message}");
            }

            return defaultTemplate;
        }

        private FileStream GenerateDefaultTemplate(string logicalName, Guid viewId, ViewType viewType)
        {
            FileStream fs = null;

            return fs;
        }       
    }

    public class ExcelTemplate
    {
        private ExcelTemplate() { }
        public string TemplateName { get; private set; }        
        public string LogicalName { get; private set; }
        public Dictionary<string, string> ColumnMappings { get; private set; }

        private String base64Content;

        public static ExcelTemplate FromEntityRecord(Entity template)
        {
            ExcelTemplate tmp = new ExcelTemplate() { LogicalName = template.GetAttributeValue<string>("associatedentitytypecode") }; 
            tmp.ColumnMappings = new Dictionary<string, string>();
            tmp.base64Content = template.GetAttributeValue<string>("content");
            tmp.TemplateName = template.GetAttributeValue<string>("name");

            tmp.PopulateColumnMappings();

            return tmp;
        }

        public static ExcelTemplate FromFile(string filePath)
        {
            return new ExcelTemplate();
        }

        public static ExcelTemplate FromStream(Stream fs, string logicalName, string templateName)
        {
            ExcelTemplate template = new ExcelTemplate()
            {
                base64Content = Convert.ToBase64String(fs.ReadAllBytes()),
                ColumnMappings = new Dictionary<string, string>(),
                LogicalName = logicalName,
                TemplateName = templateName
            };
            template.PopulateColumnMappings();

            return template;
        }


        public void PopulateColumnMappings()
        {
            ColumnMappings = new Dictionary<string, string>();

            string mappingSheet = ExtractMapFromFile(WriteToDisk(base64Content, Environment.CurrentDirectory));            
            ParseD365Map(mappingSheet);
        }

        private string ExtractMapFromFile(Stream fileStream)
        {
            SharedStringItem item = null;
            //Single-Sheet Spreadsheet -> xl\worksheets\sheet.xml
            //Multi-Sheet Spreadsheet -> xl\sharedStrings.xml
            //default tempalte generated auto: sheet.xml, destination sheet check if it's single or multi
            SpreadsheetDocument doc = SpreadsheetDocument.Open(fileStream, false);

            string map = null;

            if(doc.WorkbookPart.SharedStringTablePart != null)
            {
                SharedStringTable tbl = doc.WorkbookPart.SharedStringTablePart.SharedStringTable;
                item = (SharedStringItem)tbl.SingleOrDefault(o => !string.IsNullOrEmpty(FindD365Map(o.InnerText) ));
                map = item.InnerText;
            }
            else
            {
                WorksheetPart hiddenSheet = doc.WorkbookPart.WorksheetParts.FirstOrDefault(o => o.Worksheet.SheetProperties.CodeName == "hiddenDataSheet");
                SheetData dataSheet = (SheetData)hiddenSheet.Worksheet.ChildElements.SingleOrDefault(o => o.GetType() == typeof(SheetData));
                foreach (var sheet in dataSheet.ChildElements)
                {
                    if(!string.IsNullOrEmpty(FindD365Map(sheet.InnerText)))
                    {
                        map = sheet.InnerText;
                        break;
                    }
                }
            }

            doc.Close();
            fileStream.Close();

            if (string.IsNullOrEmpty(map))
            {
                throw new Exception("Template Map was not found!");
            }

            return map;
        }

        public string FindD365Map(string content)
        {
            return content.StartsWith($"{LogicalName}:") ? content : null;
        }

        private Stream WriteToDisk(string base64Content, string filePath)
        {
            try
            {
                string fileName = $"{filePath}\\{this.TemplateName}.xlsx";
                byte[] buffer = Convert.FromBase64String(base64Content);
                MemoryStream fs = new MemoryStream(buffer);
                fs.Position = 0;

                return fs;
            }
            catch(IOException e)
            {
                throw new Exception($"Failed to write originalEntityTemplate to disk: {e.Message}");
            }
        }

        public void ReplaceColumnMappings(ExcelTemplate plainTemplate, Dictionary<string,Guid> relationShipMappings)
        {
            //Get Guids of related entity attributes of primary template
            //Get Guids of related entity attributes of plainTemplate
            //Use mapping to replace Guids
            //ColumnMappings
            plainTemplate.PopulateColumnMappings();


        }

        private void ParseD365Map(string content)
        {
            string map = content.Replace($"{LogicalName}:", "");
            //Remove base64 data
            map = map.Substring(map.IndexOf("=:") + 2);
            List<string> columnDefs = map.Split(new [] {@"&"}, StringSplitOptions.RemoveEmptyEntries).ToList();
            foreach(string i in columnDefs)
            {
                string[] columnMap = i.Split('=');
                try
                {
                    ColumnMappings.Add(columnMap[0], System.Web.HttpUtility.UrlDecode(columnMap[1]));
                }
                catch(ArgumentException e)
                {
                    //TODO: possible duplciate column in mappings produce argument exception because of dictionary
                }
            }
        }

        private List<Guid> GetRelatedEntityGuid(string map)
        {
            List<Guid> relatedEntityGuid = null;
            Stream st = null;
            if (st != null)
            {
                string content = new StreamReader(st).ReadToEnd();
                

            }
            else
            {
                //neither file exists
                throw new Exception("File is not a valid originalEntityTemplate!");
            }

            return relatedEntityGuid;
        }

        public Entity ToEntity()
        {
            Entity entity = new Entity(LogicalName);

            entity["documenttype"] = new OptionSetValue(2);
            //entity["associatedentitytypecode"] = "";
            entity["name"] = TemplateName;
            entity["content"] = base64Content;


            return entity;
        }
    }
}

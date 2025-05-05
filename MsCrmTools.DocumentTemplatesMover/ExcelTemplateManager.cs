using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO.Compression;
using System.Windows.Forms;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using NuGet;
using System.Text.RegularExpressions;
using McTools.Xrm.Connection;
using System.Collections.ObjectModel;
using System.ComponentModel;
using XrmToolBox.Extensibility;
using DocumentFormat.OpenXml.Office2010.CustomUI;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.Crm.Sdk.Messages;
using System.ServiceModel;

namespace MsCrmTools.DocumentTemplatesMover
{
    enum ViewType
    {
        Personal,
        System
    }

    class ExcelTemplateManager
    {
        public ExcelTemplateManager() { }
        public ExcelTemplateManager(string name) { }

        public ExcelTemplateManager(ConnectionDetail connectionDetail1, ConnectionDetail connectionDetail2)
        {
            this.sourceServer = connectionDetail1;
            this.destinationServer = connectionDetail2;
        }

        private BackgroundWorker bg;
        private PluginControl pluginControl;
        private ConnectionDetail destinationServer;
        private ConnectionDetail sourceServer;

        internal void Transfer(Entity template, PluginControl pluginControl, BackgroundWorker worker)
        {
            bg = worker;
            pluginControl.LogInfo("Fetching Template");
            ExcelTemplate originalTemplate = null;
            try
            {
                originalTemplate = ExcelTemplate.FromEntityRecord(template);
            }
            catch(Exception ex)
            {
                pluginControl.LogError($"Failed to get template: {template.GetAttributeValue<string>("name")}. Reason: {ex.Message}");
                throw new Exception($"Failed to read original template: {template.GetAttributeValue<string>("name")}", ex);
            }

            ExcelTemplate destinationEnvTemplate = GenerateDefaultTemplate(originalTemplate);

            originalTemplate.ReplaceColumnMappings(destinationEnvTemplate);
        }

        private ExcelTemplate GenerateDefaultTemplate(ExcelTemplate template)
        {
            ExcelTemplate defaultTemplate = null;

            OrganizationRequest req = new OrganizationRequest("ExportTemplateToExcel");
            req.Parameters = new ParameterCollection()
            {
                ["EntityLocalizedDisplayName"] = "Case Intervention",
                ["FetchXml"] = "",
                ["LayoutXml"] = ""
            };
            Entity newTemplate = null;

            try
            {
                OrganizationResponse resp = destinationServer.ServiceClient.Execute(req);

            }
            catch(FaultException e)
            {
                throw new Exception($"Failed to create default template: {e.Message}");
            }


            if(newTemplate!= null)
            {
                defaultTemplate = ExcelTemplate.FromEntityRecord(newTemplate);
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

        public void PopulateColumnMappings()
        {
            ColumnMappings = new Dictionary<string, string>();

            string mappingSheet = ExtractMapFromFile(WriteToDisk(base64Content, Environment.CurrentDirectory), "");            
            ParseD365Map(mappingSheet);
        }

        private string ExtractMapFromFile(Stream fileStream, string filePathToExport)
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
                foreach(var sheet in doc.WorkbookPart.Workbook.Sheets)
                {
                    if(!string.IsNullOrEmpty(FindD365Map(sheet.InnerText)))
                    {
                        map = sheet.InnerText;
                        break;
                    }
                }
            }
            //ZipArchive archive = new ZipArchive( new ZipArchive(fileStream);
            //ZipArchiveEntry worksheet = archive.GetEntry(@"xl\worksheets\sheet.xml");
            //return worksheet.Open();

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

        private FileStream WriteToDisk(string base64Content, string filePath)
        {
            try
            {
                string fileName = $"{filePath}\\{this.TemplateName}.xlsx";
                FileStream fs = new FileStream(fileName, FileMode.Create, FileAccess.ReadWrite);
                byte[] buffer = Convert.FromBase64String(base64Content);
                fs.Write(buffer, 0, buffer.Length);
                fs.Flush();
                fs.Position = 0;

                return fs;
            }
            catch(IOException e)
            {
                throw new Exception($"Failed to write template to disk: {e.Message}");
            }
        }

        public void ReplaceColumnMappings(ExcelTemplate plainTemplate)
        {
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
                ColumnMappings.Add(columnMap[0], System.Web.HttpUtility.UrlDecode(columnMap[1]));
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
                throw new Exception("File is not a valid template!");
            }

            return relatedEntityGuid;
        }
    }
}

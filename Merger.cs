using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Word;
using System.IO;

namespace word_mailmerge
{
    public class Merger
    {
        Application wordApp;
        Object oMissing = System.Reflection.Missing.Value;
        Object oFalse = false;
        Object oTrue = true;
        string rootFolder;
        string templateDocument;
        string docName = "cost_and_charges.docx";
        string sourceDataCsv;

        public Merger(ref Application app)
        {
            rootFolder = Settings.Default.RootFolder;
            sourceDataCsv = $"{rootFolder}\\{Settings.Default.SourceDataFile}";
            templateDocument = $"{rootFolder}\\{Settings.Default.TemplateDocument}";            //Annual Cost and Charges Report 2019_AT.docm";
            docName = Settings.Default.DocumentOutName;
            wordApp = app;
        }

        public List<string> ReadCsv(string filename)
        {
            var res = new List<string>();
            using (StreamReader rd = new StreamReader(filename))
            {
                while (!rd.EndOfStream)
                {
                    res.Add(rd.ReadLine());
                }
            }
            return res;
        }

        public Document MergeCore(string[] headers, string[] dataFields, ref Application wordApp, object templateName)
        {
            Document wordDoc = new Document();

            if (dataFields.Length == 1)
                return wordDoc;

            Object oMissing = System.Reflection.Missing.Value;
            Object oTemplatePath = (object)templateName;

            wordDoc = wordApp.Documents.Add(ref oTemplatePath, ref oMissing, ref oMissing, ref oMissing);
            foreach (Field myMergeField in wordDoc.Fields)
            {
                Range rngFieldCode = myMergeField.Code;
                String fieldText = rngFieldCode.Text;

                if (fieldText.StartsWith(" MERGEFIELD "))
                {
                    Int32 endMerge = (" MERGEFIELD ").Length;
                    Int32 fieldNameLength = fieldText.Length - endMerge;
                    String fieldName = fieldText.Substring(endMerge, fieldNameLength);

                    fieldName = fieldName.Trim();

                    for (int ix = 0; ix < headers.Length; ix++)
                    {
                        if (headers[ix] == fieldName)
                        {
                            myMergeField.Select();
                            wordApp.Selection.TypeText(dataFields[ix]);
                        }
                    }
                }
            }
            return wordDoc;
        }

        public List<SmtpMailDetails> Merge(int mailBatchId)
        {
            var datarows = ReadCsv(sourceDataCsv);
            string[] headers = new string[0];
            var mailingList = new List<SmtpMailDetails>();

            foreach ( var row in datarows)
            {
                var attachments = new List<string>();
                var mailToList = new List<Tuple<string, string>>();

                if (headers.Length == 0)
                {
                    headers = datarows[0].Split(';');
                    continue;
                }

                var dataFields = row.Split(';');
                if (dataFields.Length == 1 && dataFields[0] == "")
                    continue;

                var wordDoc = MergeCore(headers, dataFields, ref wordApp, templateDocument);

                var docGuid = Guid.NewGuid();
                //var documentName = $"{Guid.NewGuid()}_{docName}.docx";
                var documentName = $"{wordDoc.DocID}_{docName}.docx";

                string filePath = $"{rootFolder}\\{documentName}";
                attachments.Add(filePath);
                wordDoc.SaveAs(filePath);

                mailToList.Add(new Tuple<string, string>(dataFields[0], $"{dataFields[1]} {dataFields[2]}"));

                var details = new SmtpMailDetails();
                details.smtp_mail_batch_id = mailBatchId;
                details.mime_mail_to_list = mailToList;
                details.mime_attachment_list = attachments;
        
                mailingList.Add(details);
                wordDoc.Close();
            }
            return mailingList;
        }
    }
}

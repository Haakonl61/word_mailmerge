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

        _Document wrdDoc;
        Object oMissing = System.Reflection.Missing.Value;
        Object oFalse = false;
        Object oTrue = true;
        string rootFolder = Settings.Default.RootFolder;
        Object oTemplateName = rootFolder + "\\State University.docx";
        string docName = "cost_and_charges.docx";
        Object oDotName = rootFolder + "\\State University.dotx";
        Object oDataName = rootFolder + "\\DataDoc.doc";
        Object oMergedName = rootFolder + "\\MergedDoc.docx";
        string sourceDataCsv = Settings.Default.SorceDataPath;

        public Merger(ref Application app)
        {
            wordApp = app;
        }

        private void InsertLines(int LineNum)
        {
            int iCount;

            // Insert "LineNum" blank lines.
            for (iCount = 1; iCount <= LineNum; iCount++)
            {
                wordApp.Selection.TypeParagraph();
            }
        }

        private void FillRow(_Document oDoc, int Row, string Text1,
        string Text2, string Text3, string Text4)
        {
            // Insert the data into the specific cell.
            oDoc.Tables[1].Cell(Row, 1).Range.InsertAfter(Text1);
            oDoc.Tables[1].Cell(Row, 2).Range.InsertAfter(Text2);
            oDoc.Tables[1].Cell(Row, 3).Range.InsertAfter(Text3);
            oDoc.Tables[1].Cell(Row, 4).Range.InsertAfter(Text4);
        }

        private void CreateMailMergeDataFile()
        {
            _Document oDataDoc;
            int iCount;

            Object oName = "C:\\source_files\\word_mailmerge\\DataDoc.doc";
            Object oHeader = "FirstName;LastName;Address;CityStateZip";
            wrdDoc.MailMerge.CreateDataSource(ref oName, ref oMissing,
            ref oMissing, ref oHeader, ref oMissing, ref oMissing,
            ref oMissing, ref oMissing, ref oMissing);

            // Open the file to insert data.
            oDataDoc = wordApp.Documents.Open(ref oName, ref oMissing,
            ref oMissing, ref oMissing, ref oMissing, ref oMissing,
            ref oMissing, ref oMissing, ref oMissing, ref oMissing,
            ref oMissing, ref oMissing, ref oMissing, ref oMissing,
            ref oMissing/*, ref oMissing */);

            for (iCount = 1; iCount <= 2; iCount++)
            {
                oDataDoc.Tables[1].Rows.Add(ref oMissing);
            }
            // Fill in the data.
            FillRow(oDataDoc, 2, "Steve", "DeBroux",
            "4567 Main Street", "Buffalo, NY  98052");
            FillRow(oDataDoc, 3, "Jan", "Miksovsky",
            "1234 5th Street", "Charlotte, NC  98765");
            FillRow(oDataDoc, 4, "Brian", "Valentine",
            "12348 78th Street  Apt. 214",
            "Lubbock, TX  25874");
            // Save and close the file.
            oDataDoc.Save();
            oDataDoc.Close(ref oFalse, ref oMissing, ref oMissing);
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

                var wordDoc = MergeCore(headers, dataFields, ref wordApp, oTemplateName);

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
                //wordApp.Documents.Open(oMergedName);
                wordDoc.Close();
            }
            return mailingList;
        }

        public void mergeit() //object sender, System.EventArgs e)
        {
            Selection wrdSelection;
            MailMerge wrdMailMerge;
            MailMergeFields wrdMergeFields;
            Table wrdTable;
            string StrToAdd;

            // Create an instance of Word  and make it visible.
            wordApp = new Application();
            wordApp.Visible = true;

            // Add a new document.
            wrdDoc = wordApp.Documents.Add(ref oMissing, ref oMissing,
            ref oMissing, ref oMissing);
            wrdDoc.Select();

            wrdSelection = wordApp.Selection;
            wrdMailMerge = wrdDoc.MailMerge;

            // Create a MailMerge Data file.
            CreateMailMergeDataFile();

            // Create a string and insert it into the document.
            StrToAdd = "State University\r\nElectrical Engineering Department";
            wrdSelection.ParagraphFormat.Alignment =
            WdParagraphAlignment.wdAlignParagraphCenter;
            wrdSelection.TypeText(StrToAdd);

            InsertLines(4);

            // Insert merge data.
            wrdSelection.ParagraphFormat.Alignment =
            WdParagraphAlignment.wdAlignParagraphLeft;
            wrdMergeFields = wrdMailMerge.Fields;
            wrdMergeFields.Add(wrdSelection.Range, "FirstName");
            wrdSelection.TypeText(" ");
            wrdMergeFields.Add(wrdSelection.Range, "LastName");
            wrdSelection.TypeParagraph();

            wrdMergeFields.Add(wrdSelection.Range, "Address");
            wrdSelection.TypeParagraph();
            wrdMergeFields.Add(wrdSelection.Range, "CityStateZip");

            InsertLines(2);

            // Right justify the line and insert a date field
            // with the current date.
            wrdSelection.ParagraphFormat.Alignment =
            WdParagraphAlignment.wdAlignParagraphRight;

            Object objDate = "dddd, MMMM dd, yyyy";
            wrdSelection.InsertDateTime(ref objDate, ref oFalse, ref oMissing,
            ref oMissing, ref oMissing);

            InsertLines(2);

            // Justify the rest of the document.
            wrdSelection.ParagraphFormat.Alignment =
            WdParagraphAlignment.wdAlignParagraphJustify;

            wrdSelection.TypeText("Dear ");
            wrdMergeFields.Add(wrdSelection.Range, "FirstName");
            wrdSelection.TypeText(",");
            InsertLines(2);

            // Create a string and insert it into the document.
            StrToAdd = "Thank you for your recent request for next " +
            "semester's class schedule for the Electrical " +
            "Engineering Department. Enclosed with this " +
            "letter is a booklet containing all the classes " +
            "offered next semester at State University.  " +
            "Several new classes will be offered in the " +
            "Electrical Engineering Department next semester.  " +
            "These classes are listed below.";
            wrdSelection.TypeText(StrToAdd);

            InsertLines(2);

            // Insert a new table with 9 rows and 4 columns.
            wrdTable = wrdDoc.Tables.Add(wrdSelection.Range, 9, 4,
            ref oMissing, ref oMissing);
            // Set the column widths.
            wrdTable.Columns[1].SetWidth(51, WdRulerStyle.wdAdjustNone);
            wrdTable.Columns[2].SetWidth(170, WdRulerStyle.wdAdjustNone);
            wrdTable.Columns[3].SetWidth(100, WdRulerStyle.wdAdjustNone);
            wrdTable.Columns[4].SetWidth(111, WdRulerStyle.wdAdjustNone);
            // Set the shading on the first row to light gray.
            wrdTable.Rows[1].Cells.Shading.BackgroundPatternColorIndex =
            WdColorIndex.wdGray25;
            // Bold the first row.
            wrdTable.Rows[1].Range.Bold = 1;
            // Center the text in Cell (1,1).
            wrdTable.Cell(1, 1).Range.Paragraphs.Alignment =
            WdParagraphAlignment.wdAlignParagraphCenter;

            // Fill each row of the table with data.
            FillRow(wrdDoc, 1, "Class Number", "Class Name",
            "Class Time", "Instructor");
            FillRow(wrdDoc, 2, "EE220", "Introduction to Electronics II",
            "1:00-2:00 M,W,F", "Dr. Jensen");
            FillRow(wrdDoc, 3, "EE230", "Electromagnetic Field Theory I",
            "10:00-11:30 T,T", "Dr. Crump");
            FillRow(wrdDoc, 4, "EE300", "Feedback Control Systems",
            "9:00-10:00 M,W,F", "Dr. Murdy");
            FillRow(wrdDoc, 5, "EE325", "Advanced Digital Design",
            "9:00-10:30 T,T", "Dr. Alley");
            FillRow(wrdDoc, 6, "EE350", "Advanced Communication Systems",
            "9:00-10:30 T,T", "Dr. Taylor");
            FillRow(wrdDoc, 7, "EE400", "Advanced Microwave Theory",
            "1:00-2:30 T,T", "Dr. Lee");
            FillRow(wrdDoc, 8, "EE450", "Plasma Theory",
            "1:00-2:00 M,W,F", "Dr. Davis");
            FillRow(wrdDoc, 9, "EE500", "Principles of VLSI Design",
            "3:00-4:00 M,W,F", "Dr. Ellison");

            // Go to the end of the document.
            Object oConst1 = WdGoToItem.wdGoToLine;
            Object oConst2 = WdGoToDirection.wdGoToLast;
            wordApp.Selection.GoTo(ref oConst1, ref oConst2, ref oMissing, ref oMissing);
            InsertLines(2);

            // Create a string and insert it into the document.
            StrToAdd = "For additional information regarding the " +
            "Department of Electrical Engineering, " +
            "you can visit our Web site at ";
            wrdSelection.TypeText(StrToAdd);
            // Insert a hyperlink to the Web page.
            Object oAddress = "http://www.ee.stateu.tld";
            Object oRange = wrdSelection.Range;
            wrdSelection.Hyperlinks.Add(oRange, ref oAddress, ref oMissing,
            ref oMissing, ref oMissing, ref oMissing);
            // Create a string and insert it into the document
            StrToAdd = ".  Thank you for your interest in the classes " +
            "offered in the Department of Electrical " +
            "Engineering.  If you have any other questions, " +
            "please feel free to give us a call at " +
            "555-1212.\r\n\r\n" +
            "Sincerely,\r\n\r\n" +
            "Kathryn M. Hinsch\r\n" +
            "Department of Electrical Engineering \r\n";
            wrdSelection.TypeText(StrToAdd);

            // Perform mail merge.
            wrdMailMerge.Destination = WdMailMergeDestination.wdSendToNewDocument;
            wrdMailMerge.Execute(ref oTrue);

            // Close the original form document.
            wrdDoc.Saved = true;
            wrdDoc.Close(ref oFalse, ref oMissing, ref oMissing);

            // Release References.
            wrdSelection = null;
            wrdMailMerge = null;
            wrdMergeFields = null;
            wrdDoc = null;
            wordApp = null;

            // Clean up temp file.
           // System.IO.File.Delete("C:\\source_files\\word_mailmerge\\DataDoc.doc");
        }

    }
}

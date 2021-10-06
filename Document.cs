using Microsoft.Office.Interop.Word;
using System.Reflection;
using System.Windows;

namespace RenderTableCreator
{
    public class Document
    {
        internal _Application wordApp;
        private Microsoft.Office.Interop.Word.Document document;
        private object missing = Missing.Value;

        public Document()
        {
            wordApp = new Microsoft.Office.Interop.Word.Application();
            document = wordApp.Documents.Add(ref missing, ref missing, ref missing, ref missing);
        }

        public void AddHeader(WdParagraphAlignment alignment, WdColorIndex fontColor, int fontSize, string text)
        {
            foreach (Section section in document.Sections)
            {
                Range headerRange = section.Headers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                headerRange.Fields.Add(headerRange, WdFieldType.wdFieldPage);
                headerRange.ParagraphFormat.Alignment = alignment;
                headerRange.Font.ColorIndex = fontColor;
                headerRange.Font.Size = fontSize;
                headerRange.Text = text;
            }
        }

        public void AddFooter(WdParagraphAlignment alignment, WdColorIndex fontColor, int fontSize, string text)
        {
            foreach (Section section in document.Sections)
            {
                Range footerRange = section.Headers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                footerRange.ParagraphFormat.Alignment = alignment;
                footerRange.Font.ColorIndex = fontColor;
                footerRange.Font.Size = fontSize;
                footerRange.Text = text;
            }
        }

        public void AddHeading(string style, string text)
        {
            Paragraph paragraph = document.Content.Paragraphs.Add(ref missing);
            object headingStyle = style;
            paragraph.Range.set_Style(ref headingStyle);
            paragraph.Range.Text = text;
            paragraph.Range.InsertParagraphAfter();
        }

        public void AddParagraph(string text)
        {
            Paragraph paragraph = document.Content.Paragraphs.Add(ref missing);
            object headingStyle = "Normal";
            paragraph.Range.set_Style(ref headingStyle);
            paragraph.Range.Text = text;
            paragraph.Range.InsertParagraphAfter();
        }

        public Table CreateTable(int columns, int rows)
        {
            Paragraph paragraph = document.Content.Paragraphs.Add(ref missing);

            Table table = document.Tables.Add(paragraph.Range, rows, columns, ref missing, ref missing);
            table.Borders.Enable = 1;
            return table;
        }

        public void SaveDocument(string _filename)
        {
            object filename = _filename;
            try
            {
                document.SaveAs2(ref filename);
            }
            catch
            {
                MessageBox.Show("Failed to save document");

            }
            finally
            {
                document.Close(ref missing, ref missing, ref missing);
                document = null;
                wordApp.Quit(ref missing, ref missing, ref missing);
                wordApp = null;
            }
        }
    }
}

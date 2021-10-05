using System.Reflection;
using System.Windows;
using Microsoft.Office.Interop.Word;

namespace RenderTableCreator
{
    internal class Document
    {
        private Microsoft.Office.Interop.Word.Application wordApp;
        private Microsoft.Office.Interop.Word.Document document;
        private object missing = Missing.Value;

        public Document()
        {
            wordApp = new();
            document = wordApp.Documents.Add(ref missing, ref missing, ref missing, ref missing);
            AddHeader(WdParagraphAlignment.wdAlignParagraphCenter, WdColorIndex.wdBlue, 10, "Property of College Kings");
        }

        private void AddHeader(WdParagraphAlignment alignment, WdColorIndex fontColor, int fontSize, string text)
        {
            foreach (Section section in document.Sections)
            {
                Microsoft.Office.Interop.Word.Range headerRange = section.Headers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                headerRange.Fields.Add(headerRange, WdFieldType.wdFieldPage);
                headerRange.ParagraphFormat.Alignment = alignment;
                headerRange.Font.ColorIndex = fontColor;
                headerRange.Font.Size = fontSize;
                headerRange.Text = text;
            }
        }

        private void AddFooter(WdParagraphAlignment alignment, WdColorIndex fontColor, int fontSize, string text)
        {
            foreach (Section section in document.Sections)
            {
                Microsoft.Office.Interop.Word.Range footerRange = section.Headers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                footerRange.ParagraphFormat.Alignment = alignment;
                footerRange.Font.ColorIndex = fontColor;
                footerRange.Font.Size = fontSize;
                footerRange.Text = text;
            }
        }

        private void AddHeading(string style, string text)
        {
            Paragraph paragraph = document.Content.Paragraphs.Add(ref missing);
            object headingStyle = style;
            paragraph.Range.set_Style(ref headingStyle);
            paragraph.Range.Text = text;
            paragraph.Range.InsertParagraphAfter();
        }

        private void AddParagraph(string text)
        {
            document.Content.SetRange(0, 0);
            document.Content.Text = text;
        }

        private void CreateTable()
        {
            Paragraph paragraph = document.Content.Paragraphs.Add(ref missing);

            Table table = document.Tables.Add(paragraph.Range, 5, 5, ref missing, ref missing);
            table.Borders.Enable = 1;
            foreach (Row row in table.Rows)
            {
                foreach (Cell cell in row.Cells)
                {
                    // Header row
                    if (cell.RowIndex == 1)
                    {
                        cell.Range.Text = $"Column {cell.ColumnIndex.ToString()}";
                        cell.Range.Font.Bold = 1;
                        cell.Range.Font.Size = 10;

                        cell.Shading.BackgroundPatternColor = WdColor.wdColorGray25;
                        cell.VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter;
                        cell.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                    }
                    else
                    {
                        cell.Range.Text = (cell.RowIndex - 2 + cell.ColumnIndex).ToString();
                    }
                }
            }
        }

        private void SaveDocument()
        {
            object filename = @"c:\temp1.docx";
            document.SaveAs2(ref filename);
            document.Close(ref missing, ref missing, ref missing);
            document = null;
            wordApp.Quit(ref missing, ref missing, ref missing);
            wordApp = null;
            MessageBox.Show("Document created successfully !");
        }
    }
}

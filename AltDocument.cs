using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;

namespace RenderTableCreator
{
    internal class AltDocument
    {
        private Spire.Doc.Document document; 



        internal AltDocument()
        {
            document = new Spire.Doc.Document();
        }

        internal void AddHeading(String _text, BuiltinStyle _style = BuiltinStyle.Heading1)
        {
            Section section;
            if (_style == BuiltinStyle.Title)
                section = document.AddSection();
            else
                section = document.LastSection;

            Paragraph p1 = section.AddParagraph();
            p1.ApplyStyle(_style); 
            p1.Text = _text;
            
        }
        internal void AddParagraph(String _text)
        {
            Section section = document.LastSection; 
            Paragraph p1 = section.AddParagraph();
            p1.ApplyStyle(BuiltinStyle.Normal);
            p1.Text = _text; 
        }

        internal void AddTable(String[] _headings, List<RenderItem> _tableData)
        {
            int maxRows = _tableData.Count;
            int maxCols = _headings.Length;
            //int maxHeaderCols = maxCols;
                        
            Section section = document.LastSection;
            Table table = section.AddTable(true);
            table.ResetCells(maxRows + 2, maxCols);     // Add 1 for the header row.

            // Process Header Row
            for(int c = 0; c < _headings.Length; c++)
            {
                Paragraph p1 = table.Rows[0].Cells[c].AddParagraph();
                TextRange tr = p1.AppendText(_headings[c]);
            }

            for(int r = 0; r < maxRows; r++)
            {               
                
                for(int c = 0; c < maxCols; c++)
                {
                    // ImageName 
                    Paragraph p1 = table.Rows[r + 1].Cells[c].AddParagraph();
                    p1.AppendText(_tableData[r].ImageName);
                    c++;

                    // Image Description 
                    Paragraph p2 = table.Rows[r + 1].Cells[c].AddParagraph();
                    p2.AppendText(_tableData[r].Description);
                    c++;

                    // Occurences 
                    Paragraph p3 = table.Rows[r + 1].Cells[c].AddParagraph();
                    p3.AppendText(_tableData[r].RefCount.ToString()); 

                }
                
            }            
        }
        internal void SaveDocument(string _filename)
        {
            FileFormat ff = FileFormat.Docx2013; 

            if(0 == String.Compare(System.Environment.GetEnvironmentVariable("RTC_PDF_OUTPUT"), "1", true))
            {
                _filename = System.IO.Path.ChangeExtension(_filename, ".pdf");
                ff = FileFormat.PDF;
            }

            document.SaveToFile(_filename, ff);
        }

    }
}

using Microsoft.Office.Interop.Word;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows;

namespace RenderTableCreator
{
    public partial class MainWindow : System.Windows.Window
    {
        private List<RenderItem> renderItems = new();
        private Dictionary<string, RenderItem> scenes = new();

        private string selectedFile;
        private string renderTableFile;
        private string sceneName;

        internal static string errorText = "ERRORS FOUND IN TRANSCRIPT. FIX THEM AND TRY AGAIN:";
        internal static string warnText = "WARNINGS:";

        private List<string> outputLog = new();
        private List<string> notes = new();

        public MainWindow()
        {
            InitializeComponent();
        }

        private void AddLog(string text)
        {
            outputLog.Add(text);
            WindowOutput.Text = string.Join("\n", outputLog.ToArray());
        }

        private void BrowseFiles_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new();
            openFileDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            openFileDialog.Filter = "RenPy files (*.rpy)|*.rpy|All files (*.*)|*.*";
            openFileDialog.RestoreDirectory = true;

            if (openFileDialog.ShowDialog() == true)
            {
                selectedFile = openFileDialog.FileName;
                ChosenFile.Text = $"Selected File: {selectedFile}";
                CreateRenderTableButton.Visibility = Visibility.Visible;
                renderTableFile = Path.ChangeExtension(selectedFile.Trim(), ".docx");

                sceneName = renderTableFile.Split('\\').Last().Split('.').First().Replace("scene", "Scene ");
            }
        }

        private void CreateRenderTableButton_Click(object sender, RoutedEventArgs e)
        {
            List<string> speakers = new();

            StreamReader file = new(selectedFile);

            string line;
            int lineNumber = 0;
            bool inNotes = true;

            while ((line = file.ReadLine()) != null)
            {
                lineNumber++;
                line = line.Trim();

                if (line.StartsWith("#") && inNotes)
                {
                    notes.Add(line[1..].Trim());
                    continue;
                }
                else { inNotes = false; }

                CreateRenderItem(line, lineNumber);
            }

            if (errorText == "ERRORS FOUND IN TRANSCRIPT. FIX THEM AND TRY AGAIN:" && warnText == "WARNINGS:")
            {
                SuccessfulConvert();
            }
            else { FailedConvert(); }
        }

        private void CreateRenderItem(string line, int lineNumber)
        {
            if (line.StartsWith("scene") || line.StartsWith("show"))
            {
                string[] lineArgs = line.Split(' ');
                string imageName = lineArgs[1];
                if (imageName == "black") { return; }
                if (lineArgs.Length <= 2) { return; }

                string imageDesc = string.Join(' ', lineArgs[3..]);

                if (scenes.ContainsKey(imageName) && imageDesc != scenes[imageName].Description)
                {
                    errorText += $"\n{imageName}: Conflicting description found at line {lineNumber}";
                }
                else if (imageDesc == string.Empty)
                {
                    warnText += $"\n{imageName}: Missing description";
                }
                else
                {
                    renderItems.Add(new RenderItem(imageName, imageDesc, lineNumber));
                }
            }
        }

        private void CreateDocument()
        {
            Document document = new();
            document.AddHeading("Title", $"{sceneName} Render Table");

            document.AddHeading("Heading 1", "Scene Notes:");
            document.AddParagraph(string.Join("\n", notes));

            document.AddHeading("Heading 1", "Render Table:");

            // Create render table
            Table table = document.CreateTable(2, renderItems.Count);

            renderItems.Sort(delegate (RenderItem x, RenderItem y)
            {
                return x.ImageName.CompareTo(y.ImageName);
            });

            int index = 0;
            foreach (Row row in table.Rows)
            {
                if (row.Index == 1)
                {
                    row.Range.Font.Bold = 1;
                    row.Cells[1].Range.Text = "Scene";
                    row.Cells[2].Range.Text = "Description";
                }
                else
                {
                    row.Cells[1].Range.Text = renderItems[index].ImageName;
                    row.Cells[2].Range.Text = renderItems[index].Description;
                }
                index++;
            }

            document.SaveDocument(renderTableFile);
            AddLog("Render Table Created Successfully");
        }

        private void SuccessfulConvert()
        {
            CreateDocument();
        }

        private void FailedConvert()
        {
            AddLog("Failed to create render table");
            if (errorText != "ERRORS FOUND IN TRANSCRIPT. FIX THEM AND TRY AGAIN:")
            {
                AddLog(errorText);
            }
            if (warnText != "WARNINGS:")
            {
                AddLog(warnText);
            }
        }
    }
}
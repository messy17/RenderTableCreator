using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows;
using Spire.Doc.Documents; 

namespace RenderTableCreator
{
    public partial class MainWindow : System.Windows.Window
    {
        private List<RenderItem> renderList = new();
        private SortedDictionary<string, RenderItem> scenes = new();

        private string selectedFile;
        private string renderTableFile;
        private string sceneName;

        // BUGFIX - Need to enforce consistent version numbers
        // which prevents having a mix of scene v15s2_9a and v14s15_9a 
        // in the same file. Inconsistent version numbers will report an error
        // and block creating the table
        private string version = String.Empty; 

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
            // Reset the state each time the create render table button is clicked
            
            WindowOutput.Clear();
            outputLog.Clear(); 
            notes.Clear();
            scenes.Clear(); 
            errorText = "ERRORS FOUND IN TRANSCRIPT. FIX THEM AND TRY AGAIN:";
            warnText = "WARNINGS:";

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
            if(line.StartsWith("scene") || line.StartsWith("show"))
            {
                string[] lineArgs = line.Split(' ');
                string imageName = lineArgs[1];

                // BUGFIX: Enforce version consistency by checking the version 
                // of each scene to the first scene. Error on any inconsisency. 
                if(String.IsNullOrEmpty(version))
                {
                    version = imageName.Substring(0, 3);
                }
                else
                {
                    string currentVersion = imageName.Substring(0, 3); 
                    if(0 !=String.Compare(version, currentVersion, StringComparison.OrdinalIgnoreCase))
                    {
                        errorText += $"\n{imageName}: Conflicting version found at line {lineNumber}.\nThe version should be {version}";
                    }
                }
                // END BUGFIX 
                
                if (0 == String.Compare(imageName, "black", true))
                    return;

                string imageDesc = String.Empty; 
                
                if (lineArgs.Length > 2)
                    imageDesc = string.Join(' ',lineArgs[3..]);

                // Normalize the description
                imageDesc = imageDesc.Replace('#', ' ').Trim();

                // current scene not in the list; New Scene
                if(!scenes.ContainsKey(imageName))
                {
                    // New scene without a render description; Error case
                    if (String.IsNullOrEmpty(imageDesc))
                    {
                        warnText += $"\n{imageName}: Missing description at line {lineNumber}";
                    }
                    // New scene with a proper render description; add to list
                    else
                    {
                        scenes.Add(imageName, new RenderItem(
                            imageName,
                            imageDesc,
                            lineNumber));
                    }
                }
                // current scene is in the list; Potential reuse
                else
                {
                    // Scene Reuse; legit use case 
                    if(String.IsNullOrEmpty(imageDesc))
                    {
                        scenes[imageName].RefCount++;
                    }                   
                    //Existing scene with a different render description than previous; Error case
                    else if(!imageDesc.Equals(scenes[imageName].Description))
                    {
                        int originalLineNumber = scenes[imageName].LineNumber;
                        
                        errorText += $"\n{imageName}: Conflicting description found at line {lineNumber} with orignal description at line {originalLineNumber}.";
                    }

                }
            }
        }
        
        private void CreateDocument()
        {
            AltDocument document = new();
  
            document.AddHeading(String.Format("{0} Render Table", sceneName), BuiltinStyle.Title);
            document.AddHeading("Scene Notes:", BuiltinStyle.Heading1);
            document.AddParagraph(String.Join("\n", notes));
            document.AddParagraph(String.Format("Render count: {0}", renderList.Count));
            document.AddHeading("Render Table:", BuiltinStyle.Heading1);
            document.AddParagraph(String.Empty);
            
            String[] tableHeadings = { "Scene", "Description", "Occurrences" };
            document.AddTable(tableHeadings, renderList);

            document.SaveDocument(renderTableFile);
            AddLog(String.Format("Render Table Created Successfully for {0}", sceneName));

        }                

        private void SuccessfulConvert()
        {            
            renderList = scenes.Values.ToList();
            OrderList(ref renderList); 
            CreateDocument(); 
        }

        private void FailedConvert()
        {
            AddLog(String.Format("Failed to create render table for {0}.", sceneName));
            if (errorText != "ERRORS FOUND IN TRANSCRIPT. FIX THEM AND TRY AGAIN:")
            {
                AddLog(errorText);
            }
            if (warnText != "WARNINGS:")
            {
                AddLog(warnText);
            }
        }

        private void OrderList(ref List<RenderItem> list)
        {
            bool change = false; 
            do
            {

                for (int previous = 0; previous < list.Count - 1; previous++)
                {
                    int current = previous + 1;

                    int cInt = GetImageNumber(list[current].ImageName);
                    int pInt = GetImageNumber(list[previous].ImageName);

                    if (cInt < pInt)
                    {
                        SwampListEntries(previous, current, ref list);
                        change = true; 

                    }
                    else if (cInt == pInt)
                    {
                        char cLetter = list[current].ImageName[list[current].ImageName.Length - 1];
                        char pLetter = list[previous].ImageName[list[previous].ImageName.Length - 1];

                        if (cLetter < pLetter)
                        {
                            SwampListEntries(previous, current, ref list);
                            change = true; 
                        }
                    }

                }
            } while (change);
            
        }

        private static void SwampListEntries(int posA, int posB,  ref List<RenderItem> list)
        {
            RenderItem placholder = list[posB];
            list[posB] = list[posA];
            list[posA] = placholder;
        }

        private static int GetImageNumber(string imageName)
        {            
            int retval = -1;

            if (String.IsNullOrEmpty(imageName))
                return retval; 

            int underScorePos = imageName.IndexOf("_");

            if (underScorePos == -1)
                return retval; 
            
            String intStr = imageName.Substring(underScorePos + 1, imageName.Length - (underScorePos + 1));
            if (!Int32.TryParse(intStr, out retval))
            {
                intStr = intStr.Remove(intStr.Length - 1);

                if (!Int32.TryParse(intStr, out retval))
                    retval = -1;     
            }

            return retval; 

        }
    }
}
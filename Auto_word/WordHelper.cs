using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;
using Word = Microsoft.Office.Interop.Word;

namespace Auto_word
{
    internal class WordHelper
    {
        private FileInfo _fileInfo;
        Thread thread;
        Thread thread2;
        Dictionary<string, List<string>> items;
        Dictionary<string, List<string>> filesToMerge;
        //List<string> filesToMerge;

        public WordHelper(string filename)
        {
            if (File.Exists(filename))
            {
                _fileInfo = new FileInfo(filename);
                thread = new Thread(Process);
                //thread2 = new Thread(Merge);

            }
            else
            {
                throw new ArgumentException("File not found");
            }

        }

        internal void threadStart(Dictionary<string, List<string>> items)
        {
            thread.Start(items);
        }

        internal void threadStart_2(Dictionary<string, List<string>> items)
        {
            thread.Start(items);
        }

        public static void Merge(string[] filesToMerge, string outputFilename, bool insertPageBreaks,string documentTemplate)
        {
            object defaultTemplate = documentTemplate;
            object missing = System.Type.Missing;
            object pageBreak = Word.WdBreakType.wdSectionBreakNextPage;
            object outputFile = outputFilename;

            // Create a new Word application
            Word._Application wordApplication = new Word.Application();

            try
            {
                // Create a new file based on our template
                Word.Document wordDocument = wordApplication.Documents.Add(
                                              ref defaultTemplate
                                            , ref missing
                                            , ref missing
                                            , ref missing);

                // Make a Word selection object.
                Word.Selection selection = wordApplication.Selection;

                //Count the number of documents to insert;
                int documentCount = filesToMerge.Length;

                //A counter that signals that we shoudn't insert a page break at the end of document.
                int breakStop = 0;

                // Loop thru each of the Word documents
                foreach (string file in filesToMerge)
                {
                    breakStop++;
                    // Insert the files to our template
                    selection.InsertFile(
                                                file
                                            , ref missing
                                            , ref missing
                                            , ref missing
                                            , ref missing);

                    //Do we want page breaks added after each documents?
                    if (insertPageBreaks && breakStop != documentCount)
                    {
                        selection.InsertBreak(ref pageBreak);
                    }
                }

                // Save the document to it's output file.
                wordDocument.SaveAs(
                                ref outputFile
                            , ref missing
                            , ref missing
                            , ref missing
                            , ref missing
                            , ref missing
                            , ref missing
                            , ref missing
                            , ref missing
                            , ref missing
                            , ref missing
                            , ref missing
                            , ref missing
                            , ref missing
                            , ref missing
                            , ref missing);

                // Clean up!
                wordDocument = null;
            }
            catch (Exception ex)
            {
                //I didn't include a default error handler so i'm just throwing the error
                throw ex;
            }
            finally
            {
                // Finally, Close our Word application
                wordApplication.Quit(ref missing, ref missing, ref missing);
            }
        }
    

    internal void Process(object? obj)
        {
            items = (Dictionary<string, List<string>>)obj;

            Word.Application app = null; 
            try
            {

                

                for (int i = 0; i < items["<POS>"].Count; i++)
                {
                    app = new Word.Application();
                    Object file = _fileInfo.FullName;

                    Object missing = Type.Missing;

                    app.Documents.Open(file);

                    foreach (var item in items)
                    {
                        Word.Find find = app.Selection.Find;
                        find.Text = item.Key;
                        find.Replacement.Text = item.Value[i];

                        Object wrap = Word.WdFindWrap.wdFindContinue;
                        Object replace = Word.WdReplace.wdReplaceAll;

                        find.Execute(FindText: Type.Missing,
                            MatchCase: false,
                            MatchWholeWord: false,
                            MatchWildcards: false,
                            MatchSoundsLike: missing,
                            MatchAllWordForms: false,
                            Forward: true,
                            Wrap: wrap,
                            Format: false,
                            ReplaceWith: missing, Replace: replace
                            );

                    }

                    //filesToMerge.Add(_fileInfo.DirectoryName + "\\" +i.ToString() + " " + _fileInfo.Name);
                    Object newFileName = Path.Combine(_fileInfo.DirectoryName, i.ToString() + " " + _fileInfo.Name);
                    app.ActiveDocument.SaveAs2(newFileName);
                    app.ActiveDocument.Close();


                }
                //Merge(filesToMerge, _fileInfo.DirectoryName + "\\" +  )
                

                //return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            finally
            {
                if(app != null)
                {
                    app.Quit();
                }
                
            }
               //return false;
            
        }
    }
}

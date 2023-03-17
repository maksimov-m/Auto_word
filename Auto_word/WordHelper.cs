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
        Dictionary<string, List<string>> items;

        public WordHelper(string filename)
        {
            if (File.Exists(filename))
            {
                _fileInfo = new FileInfo(filename);
                thread = new Thread (Process);

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
                        
                        var wrap = Word.WdFindWrap.wdFindContinue;
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

                    Object newFileName = Path.Combine(_fileInfo.DirectoryName, i.ToString() + " " + _fileInfo.Name);
                    app.ActiveDocument.SaveAs2(newFileName);
                    app.ActiveDocument.Close();


                }
                
                

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

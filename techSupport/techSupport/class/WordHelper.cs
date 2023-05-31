using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Diagnostics;
using Word = Microsoft.Office.Interop.Word;
using Microsoft.Office.Interop.Word;
using System.Windows.Shapes;

namespace WordHelperClass
{
    class WordHelper
    {
        private FileInfo _fileInfo;

        public WordHelper(string fileName) 
        { 
            if ( File.Exists(fileName) ) 
            { 
                _fileInfo = new FileInfo(fileName);
            }
            else 
            {
                throw new ArgumentException("Файл не обнаружен!");
            }
        }

        internal bool Process(Dictionary<string, string> items)
        {
            Word.Application app = null;

            try 
            {
                app = new Word.Application();
                Object file = _fileInfo.FullName;
                Object missing = Type.Missing;
                app.Documents.Open(file);

                foreach ( var item in items ) 
                {
                    Word.Find find = app.Selection.Find;
                    find.Text = item.Key;
                    find.Replacement.Text = item.Value;

                    object wrap = Word.WdFindWrap.wdFindContinue;
                    object replace = Word.WdReplace.wdReplaceAll;

                    find.Execute(FindText: Type.Missing,
                        MatchCase: false,
                        MatchWholeWord: false,
                        MatchWildcards: false,
                        MatchSoundsLike: missing,
                        MatchAllWordForms: false,
                        Forward: true,
                        Wrap: wrap,
                        Format: false,
                        ReplaceWith: missing, Replace: replace);
                }

                string testpath = System.IO.Path.Combine(_fileInfo.DirectoryName, DateTime.Now.ToString("yyyy.MM.dd HH-mm") + ' ' + _fileInfo.Name);
                //Object newFileName = System.IO.Path.Combine(_fileInfo.DirectoryName, DateTime.Now.ToString("yyyy.MM.dd HH-mm") + ' ' + _fileInfo.Name);

                Object newFileName = testpath;
                app.ActiveDocument.SaveAs2(newFileName);
                app.ActiveDocument.Close();

                System.Diagnostics.Process.Start(testpath);
                return true;
            }
            catch (Exception ex) { MessageBox.Show("Ошибка, попробуйте еще раз!", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Error); }
            finally 
            {
                if ( app != null )
                {
                    app.Quit();
                }
            }

            return false;
        }
    }
}

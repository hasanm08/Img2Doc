using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace img_to_word
{
    class Program
    {
        static void Main(string[] args)
        {
            List<string> vs = new List<string>();
            for (int i = 41; i > 0; i--)
            {
                vs.Add(@"C:\Users\hasanm08\Downloads\Telegram Desktop\System Digital\" + i.ToString() + ".JPG");
            }
            imgtodoc(vs);
        }

        private static void imgtodoc(List<string> imgs)
        {
            // first we are creating application of word.
            Microsoft.Office.Interop.Word.Application WordApp = new Microsoft.Office.Interop.Word.Application();
            // now creating new document.
            WordApp.Documents.Add();
            // see word file behind your program
            WordApp.Visible = true;
            // get the reference of active document
            Microsoft.Office.Interop.Word.Document doc = WordApp.ActiveDocument;
            foreach (string item in imgs)
            {
                doc.InlineShapes.AddPicture(item, Type.Missing, Type.Missing, Type.Missing);
                Console.WriteLine(item + " Done!");
            }
            // file is saved.
            doc.SaveAs(@"C:\Users\hasanm08\Downloads\Telegram Desktop\System Digital\hello.doc", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            // application is now quit.
            WordApp.Quit(Type.Missing, Type.Missing, Type.Missing);
        }
    }
}

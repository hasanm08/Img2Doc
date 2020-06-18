using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Diagnostics;
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
            Console.WriteLine("Pleasse Enter num of pics :");
            int n = int.Parse(Console.ReadLine());
            for (int i = n; i > 0; i--)
            {
                vs.Add(@"C:\Users\L340\Desktop\signal\Signal (" + i.ToString() + ").JPG");
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

            doc.Content.Text += "\n\n\nCreatedBy ImageToDoc C# App \n Please Report any problems";
            foreach (string item in imgs)
            {

                doc.InlineShapes.AddPicture(item, Type.Missing, Type.Missing, Type.Missing);
                Console.WriteLine(item + " Done!");
            }
            // file is saved.
            doc.SaveAs(@"C:\Users\L340\Desktop\signal\hello.doc", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            // application is now quit.
            WordApp.Quit(Type.Missing, Type.Missing, Type.Missing);
            ProcessStartInfo info = new ProcessStartInfo(@"C:\Users\L340\Desktop\signal\hello.doc");
            info.Verb = "Print";
            info.CreateNoWindow = true;
            info.WindowStyle = ProcessWindowStyle.Hidden;
            Process.Start(info);
        }
    }
}

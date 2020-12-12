using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace img_to_word
{
    class Program
    {
        public static string Path = "";
        static void Main(string[] args)
        {
            List<string> vs = new List<string>();
            Console.WriteLine("Please Enter specified folder Contains jpg images : ");
            Path = Console.ReadLine();
            string imageName = "ImageToDoc08";
            DirectoryInfo d = new DirectoryInfo(Path);
            var infos = d.EnumerateFiles("*.jpg").OrderBy(f => f.CreationTime);
            int tmp = 0;
            Console.WriteLine("Just Wait ...");

            foreach (FileInfo f in infos)
            {
                try
                {
                    File.Move(f.FullName, f.FullName.Replace(f.Name, imageName+" ("+(++tmp)+").jpg"));
                }
                catch { }
            }
            for (int i = tmp; i > 0; i--)
            {
                string path =Path+ @"\"+imageName+" (" + i.ToString() + ").jpg";
                Image img = Image.FromFile(path);
                if (img.Width > img.Height)
                {
                    //Rotate the image in memory
                    img.RotateFlip(RotateFlipType.Rotate90FlipNone);

                    //Delete the file so the new image can be saved
                    File.Delete(path);

                    //save the image out to the file
                    img.Save(path);

                    //release image file
                    img.Dispose();
                }

                vs.Add(path);
            }
            
            Img2Doc(vs);
            Console.ReadKey(false);
        }

        private static void Img2Doc(List<string> imgs)
        {
            // first we are creating application of word.
            Application WordApp = new Application();
            // now creating new document.
            WordApp.Documents.Add();
            // see word file behind your program
            WordApp.Visible = false;
            // get the reference of active document
            Document doc = WordApp.ActiveDocument;

            doc.Content.Text += "\n\n\nCreatedBy ImageToDoc C# App \n Please Report any problems\n hasanm08.github.io \n https://github.com/hasanm08/Img2Doc";
            foreach (string item in imgs)
            {

                doc.InlineShapes.AddPicture(item, Type.Missing, Type.Missing, Type.Missing);
                Console.WriteLine(item + " Done!");
            }
            // file is saved.
            doc.SaveAs(Path+@"\Output08.doc", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            // application is now quit.
            WordApp.Quit(Type.Missing, Type.Missing, Type.Missing);
            Console.WriteLine("Word File Is Ready : " + Path + @"\Output08.doc");
            Console.WriteLine("\n\nDo You Want to Print? 1. yes 2. no");
            try
            {
                int n = int.Parse(Console.ReadLine());
                if (n==1)
                {
                    ProcessStartInfo info = new ProcessStartInfo(Path + @"\Output08.doc")
                    {
                          Verb = "Print",
                          CreateNoWindow = true,
                          WindowStyle = ProcessWindowStyle.Hidden
                    };
                    Process.Start(info);
                    Console.WriteLine("Have A good day :) \n press any key to exit");
                    return;
                }
                else
                {
                    throw new Exception();
                }
            }
            catch (Exception)
            {
                Console.WriteLine("Have A good day :) \n press any key to exit");
                return;
            } 
            
        }
    }
}

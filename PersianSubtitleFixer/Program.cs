using System;
using System.Text;
using System.IO;
using Word = Microsoft.Office.Interop.Word;
using Core = Microsoft.Office.Core;
using Marshal = System.Runtime.InteropServices.Marshal;
using System.Collections.Generic;

namespace PersianSubtitleFixer
{
    class Program
    {

        static void Main(string[] args)
        {
            Console.WriteLine("For fixing an entire directory, enter 1");
            Console.WriteLine("For fixing separate files, enter 2");
            int userSelection = Convert.ToInt32(Console.ReadLine());
            List<string> filenames;
            if (userSelection == 1)
                filenames = DirectoryFiles();
            else
                filenames = SeparateFiles();
            Word.Application app = new Word.Application();
            app.Visible = false;
            app.DisplayAlerts = Word.WdAlertLevel.wdAlertsNone;
            //
            foreach (string filename in filenames)
            {
                if (File.Exists(filename))
                {
                    Word.Document document = app.Documents.Open(filename, Visible: false, ReadOnly: true);
                    string filePath = Path.GetDirectoryName(filename) + "\\";
                    //
                    switch (document.TextEncoding)
                    {
                        case Core.MsoEncoding.msoEncodingUTF8:
                        case Core.MsoEncoding.msoEncodingUnicodeBigEndian:
                        case Core.MsoEncoding.msoEncodingUnicodeLittleEndian:
                            break;
                        default:
                            document = app.Documents.Open(filename, Encoding: Core.MsoEncoding.msoEncodingArabicAutoDetect);
                            break;
                    }
                    string full = waitForIt(getText, document);
                    string pureFileName = Path.GetFileNameWithoutExtension(filename);
                    string outputFileName = filePath + pureFileName + "-fixed.srt";
                    Console.WriteLine(outputFileName);
                    File.WriteAllText(outputFileName, full, Encoding.UTF8);
                    waitForIt(closeDocument, document);
                    releaseObject(document, false);
                    //System.Threading.Thread.Sleep(1000);
                }
                else
                {
                    Console.WriteLine("file: " + filename + " does not exist.");
                }
            }
            app.Quit();
            releaseObject(app, true);
            Console.WriteLine("All of your files were fixed successfully, enter a key to exit...");
            Console.ReadKey();
        }

        static void releaseObject(object obj, bool gc)
        {
            Marshal.ReleaseComObject(obj);
            obj = null;
            if (gc)
                GC.Collect();
        }

        static List<string> SeparateFiles()
        {
            Console.WriteLine("Drag your files and drop them here, when you're finished, enter q");
            List<string> filenames = new List<string>();
            string name = Console.ReadLine();
            while (name != "q")
            {
                filenames.Add(name.Replace("\"", ""));
                name = Console.ReadLine();
            }
            return filenames;
        }

        static List<string> DirectoryFiles()
        {
            Console.WriteLine("Type the list of the directories you want to fix, when you're finished, enter q");
            List<string> filenames = new List<string>();
            string directory = Console.ReadLine();
            while (directory != "q")
            {
                directory = directory.Replace("\"", "");
                if (Directory.Exists(directory))
                {
                    string[] filePaths = Directory.GetFiles(directory, "*.srt", SearchOption.TopDirectoryOnly);
                    filenames.AddRange(filePaths);
                }
                directory = Console.ReadLine();
            }
            return filenames;
        }

        static string getText(Word.Document doc)
        {
            return doc.Content.Text;
        }

        static string closeDocument(Word.Document doc)
        {
            doc.Close();
            // This last line is to avoid duplicate code.
            return "";
        }

        static string waitForIt(Func<Word.Document, string> func, Word.Document doc)
        {
            while (true)
            {
                try
                {
                    return func(doc);
                }
                catch (Exception)
                {
                    System.Threading.Thread.Sleep(500);
                    continue;
                }
            }
        }
    }
}

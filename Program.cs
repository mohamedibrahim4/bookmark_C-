using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
//using static System.Net.Mime.MediaTypeNames;

namespace bookmark
{
    class Program
    {
        static void Main(string[] args)
        {
             Application app = new Application();
            Document doc = app.Documents.Open("C:/Users/muhamed/Desktop/bookmark/Requirerments.docx");

            Console.WriteLine("number of bookmarks:"+doc.Bookmarks.Count);

            if (doc.Bookmarks.Exists("EDIT"))
            {
                object oBookMark = "EDIT";
                object range = doc.Bookmarks.get_Item(ref oBookMark).Range.Text;
                object range2 = doc.Bookmarks.get_Item(ref oBookMark).Range;

                doc.Bookmarks.get_Item(ref oBookMark).Range.Text = range.ToString();

                

              

                doc.Bookmarks.Add("Test", range2);


            }
            if (doc.Bookmarks.Exists("F"))
            {
                object oBookMark = "F";
                object rangeb = doc.Bookmarks.get_Item(ref oBookMark).Range.Text;
                object rangeb2 = doc.Bookmarks.get_Item(ref oBookMark).Range;
                doc.Bookmarks.get_Item(ref oBookMark).Range.Text = rangeb.ToString();


                doc.Bookmarks.Add("TestF", rangeb2);

            }


            doc.Save();

            doc.ExportAsFixedFormat("myNewPdf.pdf", WdExportFormat.wdExportFormatPDF);

            ((_Document)doc).Close();
            ((_Application)app).Quit();
        }
    }
}

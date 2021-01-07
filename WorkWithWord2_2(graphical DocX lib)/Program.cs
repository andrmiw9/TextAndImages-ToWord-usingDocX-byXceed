using System;
using System.Collections.Generic;
using System.Linq;
using System.Drawing;
using System.Drawing.Imaging;
using System.Threading.Tasks;
using System.Threading;
using System.Windows.Forms;
using Xceed.Document.NET;
using Xceed.Words.NET;

namespace WorkWithWord2_2_graphical_DocX_lib_
{
    static class Programm
    {
        /// <summary>
        /// Главная точка входа для приложения.
        /// </summary>
        //  Маркин Андрей 2020
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new Entering());
        }
        #region defining public var-s
        public static string docdic = @"C:\Users\Sinitza\Documents\APlaces\";
        //public static string imgdic = @"C:\Users\Sinitza\Documents\APlaces\";
        public static string templatedic = @"C:\Users\Sinitza\Documents\APlaces\Templates\";
        public static string docname = "Places1.docx";
        public static string templatename = "Template1.docx";
        //public static string picname = "ex2.jpg";
        //комфортная ширина картинки = 520, 292 = 520*1.777777 (16/9 = 1.777777)
        #endregion
        public static void EditT(string img1name, string img2name, string h1 = "1", string h2 = "2", string h3 = "3")
        {
            //if (h1.IndexOf('1') == 0 && h2.IndexOf('2') == 0)     //error checking
            // Console.WriteLine("Editting Template...");
            
            using (var doc1 = DocX.Load(templatedic + templatename))
            {
                doc1.ReplaceText("oneplkj", (h1));
                doc1.ReplaceText("twohgis", (h2));
                doc1.ReplaceText("threejsyv", (h3));
                try
                {
                    var img1 = doc1.AddImage(img1name);
                    var pic1 = img1.CreatePicture(292, 520);        //комфортная ширина картинки = 520, 292 = 520/1.777777 (16/9 = 1.777777)
                    var img2 = doc1.AddImage(img2name);
                    var pic2 = img2.CreatePicture(292, 520);        //комфортная ширина картинки = 520, 292 = 520/1.777777 (16/9 = 1.777777)
                    doc1.ReplaceTextWithObject("fournzes", pic1, false, System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                    doc1.ReplaceTextWithObject("fivemtrl", pic2, false, System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                }
                catch (Exception e)
                {
                    MessageBox.Show(e.ToString());
                    Thread.Sleep(2000);
                    //Environment.Exit(0);
                }
                doc1.SaveAs(templatedic + "TimeTemplate.docx");
                //Console.WriteLine("Finished editting template!");
            }
        }
        public static void AddT()
        {
            //Console.WriteLine("Adding template with appending...");
            using (var doc1 = DocX.Load(docdic + docname))
            {
                using (var doc2 = DocX.Load(templatedic + "TimeTemplate.docx"))
                {
                    doc1.InsertDocument(doc2, true);
                    //var templatepath = templatedic + templatename;
                    //doc2.ApplyTemplate(templatepath);
                    doc1.Save();
                    MessageBox.Show("Finished adding template");
                    //Console.WriteLine("Finished adding template!");
                }
            }
        }
    }
}

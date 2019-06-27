using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Word;
namespace IHKDocScanner
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Das Programm überprüft das Dokument auf korrekte Formatierungen bzgl. Rändern, Inhaltsverzeichnis, Schriftgröße, -Art und -Format und Überschriften.");
            Console.WriteLine("Es wird nur Word mit der Benutzersprache Deutsch unterstützt.");
            Console.WriteLine("Die Absatzangaben beziehen sich auf das gesamte Dokument.");
            RunProgram();

        }
        static void RunProgram()
        {
            Console.WriteLine("Dateipfad des Dokuments angeben.");
            String fileName = Console.ReadLine();
            Application wordApp = new Application();
            String fileTest = "C:\\IHKDocTest\\ihk.docx";
            Document wordDoc = null;

            try
            {
                wordDoc = wordApp.Documents.Open(fileTest);
            }
            catch
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("Datei nicht gefunden");
                Console.ForegroundColor = ConsoleColor.Gray;
                RunProgram();
            }
            //---------------------------------------------------------

            
            //-----------------------------------------
            TextFormatting tf = new TextFormatting(wordDoc);
            GlobalFormating glF = new GlobalFormating(wordDoc);
            ParagraphFormatting stC = new ParagraphFormatting(wordDoc);

            Console.WriteLine();
            Console.WriteLine("globale Einstellungen");
            glF.checkMargin();
            glF.checkPageCount();
            glF.checkTableOfContents();
            if (!glF.CheckFooter())
                glF.CheckPageNumbers();

            Console.WriteLine();

            for (int par = 1; par < wordDoc.Paragraphs.Count; par++)
            {
                Paragraph paragraph = wordDoc.Paragraphs[par];
                int pageNumber = paragraph.Range.Information[WdInformation.wdActiveEndAdjustedPageNumber];

                tf.SetClassAttributes(par, paragraph, pageNumber);
                stC.SetClassAttributes(par, paragraph, pageNumber);

                if (stC.CheckHeading())
                    continue;
                stC.CheckWidow();
                tf.CheckFontSize();
                tf.CheckFont();
                tf.CheckLineSpacing();
                tf.CheckWordFormat();
            }

            wordApp.Documents.Close();
            wordApp.Quit();
            Console.ReadLine();
        }


        public static void checkHeader(Document wordDoc, Range rng)

        {
            //WdPageNumberStyle.wd
            int countSec = wordDoc.Sections.Count;
            Console.WriteLine("Sections: "+countSec);
            List<String> testHeader = new List<String>();
            foreach(Section asection in wordDoc.Sections)
            {
                foreach(HeaderFooter aHeader in asection.Footers)
                {
                    testHeader.Add(aHeader.Range.Text);
                    if (aHeader.PageNumbers == null)
                        Console.WriteLine("es existieren nummern!!!!!!!!!!!!!!!!!!!!!");
                }
            }
            foreach(String head in testHeader)
            {
                Console.WriteLine("footerListe " + head);
            }
            for(int i = 1; i <= wordDoc.Sections.Count; i++)
            {
                Section section = rng.Sections[i];
                if(section != null)
                {
                    HeaderFooter headOrFoot = section.Headers[WdHeaderFooterIndex.wdHeaderFooterFirstPage];
                    if(headOrFoot.PageNumbers.Count>0)
                    {
                        Console.WriteLine("FirstPage: Nummern vorhanden");
                        Console.WriteLine("inhalt: " + headOrFoot.Range.Text);
                    }
                    else
                    {
                        Console.WriteLine("FirstPage: keine Nummern vorhanden");
                        Console.WriteLine("inhalt: " + headOrFoot.Range.Text);
                    }

                    headOrFoot = section.Headers[WdHeaderFooterIndex.wdHeaderFooterEvenPages];

                    if (headOrFoot.PageNumbers.Count > 0)
                    {
                        Console.WriteLine("EvenPages: Nummern vorhanden");
                        Console.WriteLine("inhalt: " + headOrFoot.Range.Text);
                    }
                    else
                    {
                        Console.WriteLine("EvenPages: keine Nummern vorhanden");
                        Console.WriteLine("inhalt: " + headOrFoot.Range.Text);
                    }

                     headOrFoot = section.Headers[WdHeaderFooterIndex.wdHeaderFooterPrimary];
                    if (headOrFoot.PageNumbers.Count > 0)
                    {
                        Console.WriteLine("PrimaryPages: Nummern vorhanden");
                        Console.WriteLine("inhalt: " + headOrFoot.Range.Text);
                    }
                    else
                    {
                        Console.WriteLine("PrimaryPages = keine Nummern vorhanden");
                        Console.WriteLine("inhalt: " + headOrFoot.Range.Text);
                    }
                }
            }

            if (wordDoc.Sections[1].Headers[WdHeaderFooterIndex.wdHeaderFooterPrimary].IsHeader) 

            {
                Console.WriteLine("er sagt ist header or footer");
            }
            Sections sections = wordDoc.Sections;
            List<HeadersFooters> test = new List<HeadersFooters>();

            if (test == null)
            {
                Console.WriteLine("kein Header");
            }
            Console.WriteLine("Header vorhanden?!");
        }
    }
}

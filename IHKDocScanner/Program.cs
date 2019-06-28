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
            Document wordDoc = null;
            
            try
            {
                wordDoc = wordApp.Documents.Open(fileName);
            }
            catch
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("Datei nicht gefunden");
                Console.ForegroundColor = ConsoleColor.Gray;
                RunProgram();
            }

            TextFormatting tf = new TextFormatting(wordDoc);
            GlobalFormating glF = new GlobalFormating(wordDoc);
            ParagraphFormatting stC = new ParagraphFormatting(wordDoc);

            Console.WriteLine();
            Console.WriteLine("Globale Einstellungen");
            glF.checkMargin();
            glF.checkPageCount();
            glF.checkTableOfContents();
            glF.CheckFooter();
            glF.CheckPageNumbers();
            
            Console.WriteLine();

            //Es wird durch jeden Absatz des Dokuments iteriert und die Formatierungen überprüft.
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
    }
}

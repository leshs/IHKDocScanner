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
            String fileNameOld = null;
            RunProgram(fileNameOld);
        }

        public static void RunProgram(string fileNameOld)
        {
            //Der input muss zu einem CHar[] gemacht werden und die letzten beiden Chars geprüft werden, ob sie "-H" sind. Wenn ja, ist der Dateiname Char[] - 3; ansonsten das gesamte char[];
            Application wordApp = new Application();
            Document wordDoc = null;
            string fileName;
            bool showHinweise = true;

            while (true)
            {
                Console.WriteLine("Hinweise anzeigen j/n?");
                ConsoleKey ck = Console.ReadKey().Key;
                Console.WriteLine();
                if (ck == ConsoleKey.J)
                {
                    break;
                }
                else if (ck == ConsoleKey.N)
                {
                    showHinweise = false;
                    break;
                }
            }

            if (fileNameOld == null)
            {
                Console.WriteLine("Dateipfad des Dokuments angeben.");
                fileName = Console.ReadLine();
            }
            else
            {
                Console.WriteLine(fileNameOld);
                fileName = fileNameOld;
            }
          
            try
            {
                wordDoc = wordApp.Documents.Open(fileName);
            }
            catch
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("Datei nicht gefunden oder Eingabe fehlerhaft");
                Console.ForegroundColor = ConsoleColor.Gray;
                RunProgram(null);
            }

            TextFormatting tf = new TextFormatting(wordDoc);
            GlobalFormating glF = new GlobalFormating(wordDoc);
            ParagraphFormatting stC = new ParagraphFormatting(wordDoc);

            tf.SetShowHinweise(showHinweise);

            Console.WriteLine();
            Console.WriteLine("Globale Einstellungen");

            //Globale Einstellungen überprüfen
            glF.checkMargin();
            glF.checkPageCount();
            glF.checkTableOfContents();
            glF.CheckFooter();
            glF.CheckPageNumbers();

            Console.WriteLine();

            //Iteration durch Absätze und Überprüfung von Absatz-Formatierung und Text-Formatierung
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

            int warnings = stC.GetWarnings() + glF.GetWarnings();
            int errors = tf.GetErrors() + glF.GetErrors();
            int notifications = tf.GetNotifications();

            PrintSummary(warnings, errors, notifications, showHinweise);
            NextAction(fileName);
        }

        //Ausgeben der Zusammenfassung der Fehlermeldungen (Fehler, Warnung, Hinweis)
        public static void PrintSummary(int warningCount, int errorCount, int notificationCount, bool showHinweise)
        {
            Console.WriteLine();
            Console.WriteLine("Prüfung abgeschlossen");
            Console.WriteLine("Fehler: " + errorCount);
            Console.WriteLine("Warnungen: " + warningCount);
            if(showHinweise)
                Console.WriteLine("Hinweise: " + notificationCount);
            Console.WriteLine();
        }

        public static void NextAction(string fileName)
        {
            Console.WriteLine("Um das Dokument erneut zu prüfen 'j' drücken");
            Console.WriteLine("Um ein anderes Dokument zu prüfen 'n' drücken");
            Console.WriteLine("Um das Programm zu schließen Escape drücken");

            ConsoleKey nextActionKey = Console.ReadKey().Key;
            Console.WriteLine();

            if(nextActionKey == ConsoleKey.J)
            {
                RunProgram(fileName);
            } else if (nextActionKey == ConsoleKey.N)
            {
                RunProgram(null);
            } else if (nextActionKey == ConsoleKey.Escape)
            {
                Environment.Exit(0);
            } else
            {
                Console.WriteLine("Eingabe nicht erkannt");
                Console.WriteLine();
                NextAction(fileName);
            }
        }
    }
}

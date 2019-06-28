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
            string[] oldInput = null;
            RunProgram(oldInput);
        }

        public static void RunProgram(string[] oldInput)
        {

            Application wordApp = new Application();
            Document wordDoc = null;
            bool showHinweise = true;
            string[] inputArr = null;
            if(oldInput == null)
            {
                Console.WriteLine("Dateipfad des Dokuments angeben.");
                Console.WriteLine("Mit dem Zusatz -H kann die Anzeige von Hinweisen ausgeschaltet werden.");
                string input = Console.ReadLine();
                inputArr = input.Split(' ');
            }
            else
            {
                inputArr = oldInput;
            }
            
            string fileName = inputArr[0];

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

            //Überprüfen, ob Hinweise angezeigt werden sollen
            if (inputArr.Length > 1)
            {
                if (inputArr[1] == "-H")
                    showHinweise = false;
            }
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
            NextAction(inputArr);
        }

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

        /* Methode zum bestimmen der nächsten Aktion - 3 Optionen:
         * Schließen der Konsole
         * ReRun mit gleichem Dokument
         * Rerun mit neuem Dokument
         */
        public static void NextAction(string[] inputArr)
        {
            Console.WriteLine("Um das Dokument erneut zu prüfen 'j' drücken");
            Console.WriteLine("Um ein anderes Dokument zu prüfen 'n' drücken");
            Console.WriteLine("Um das Programm zu schließen Escape drücken");

            ConsoleKey nextActionKey = Console.ReadKey().Key;

            if(nextActionKey == ConsoleKey.J)
            {
                RunProgram(inputArr);
            } else if (nextActionKey == ConsoleKey.N)
            {
                Console.WriteLine();
                RunProgram(null);
            } else if (nextActionKey == ConsoleKey.Escape)
            {
                Environment.Exit(0);
            } else
            {
                Console.WriteLine("Eingabe nicht erkannt");
                Console.WriteLine();
                NextAction(inputArr);
            }
        }
    }
}

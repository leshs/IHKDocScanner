using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;


namespace IHKDocScanner
{
    class GlobalFormating
    {
        private Range Rng;
        private PageSetup PSetup;
        private Document Document;
        private HeaderFooter FooterEven;
        private HeaderFooter FooterFirst;
        private HeaderFooter FooterPrimary;

        public GlobalFormating(Document document)
        {
            Rng = document.Range();
            PSetup = Rng.PageSetup;
            Document = document;
            FooterEven = Document.Sections.First.Footers[WdHeaderFooterIndex.wdHeaderFooterEvenPages];
            FooterFirst = Document.Sections.First.Footers[WdHeaderFooterIndex.wdHeaderFooterFirstPage];
            FooterPrimary = Document.Sections.First.Footers[WdHeaderFooterIndex.wdHeaderFooterPrimary];
        }

        //Seitenanzahl überüfen
        public void checkPageCount()
        {
            int numberOfPages = Rng.get_Information(WdInformation.wdNumberOfPagesInDocument);

            if (numberOfPages > 15)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("Fehler: Das Dokument darf höchstens 15 Seiten lang sein. Die Seitenzahl beträgt " + numberOfPages + ".");
                Console.ForegroundColor = ConsoleColor.Gray;
            }
            else
            {
                Console.WriteLine("Das Dokument ist " + numberOfPages + " Seiten lang.");
            }
        }

        //Rand des Dokumentes überprüfen.
        public void checkMargin()
        {
            float marginLeft = PSetup.LeftMargin;
            float marginRight = PSetup.RightMargin;

            if (marginLeft < 70 || marginLeft > 71)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("Fehler: Der linke Rand muss 2,5cm betragen.");
                Console.ForegroundColor = ConsoleColor.Gray;
            }
            else
            {
                Console.WriteLine("Der rechte Rand ist korrekt formatiert.");
            }

            if (marginRight < 42 || marginRight > 43)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("Fehler: Der linke Rand muss 1,5cm betragen.");
                Console.ForegroundColor = ConsoleColor.Gray;
            }
            else
            {
                Console.WriteLine("Der linke Rand ist korrekt formatiert.");
            }
        }

        //Überprüfen, ob ein Inhaltsverzeichnis existiert
        public void checkTableOfContents()
        {
            TablesOfContents tbc = Document.TablesOfContents;

            if (tbc.Count < 1)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("Fehler: Kein Inhaltsverzeichnis vorhanden.");
                Console.ForegroundColor = ConsoleColor.Gray;
            }
        }

        //Die Methode prüft, ob sich gerade und ungerade Fußzeilen unterscheiden.
        public void CheckFooter()
        {
            if (!FooterEven.Exists)
            {
                Console.ForegroundColor = ConsoleColor.Yellow;
                Console.WriteLine("Warnung: Die Fußzeilen auf geraden und ungeraden Seiten können sich unterscheiden.");
                Console.ForegroundColor = ConsoleColor.Gray;
                Console.WriteLine("In Textverarbeitungsprogrammen kann unter dem Menüpunkt Fußzeilen eingestellt werden, ob sich gerade und ungerade Seiten unterscheiden sollen");
            }
        }

        //Überprüft, ob dynamische Seitennummern vorhanden sind
        public void CheckPageNumbers()
        {
            int pageNumbersEven = FooterEven.PageNumbers.Count;
            int pageNumbersPrimary = FooterPrimary.PageNumbers.Count;

            if (!(pageNumbersEven > 0) && !(pageNumbersPrimary > 0))
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("Fehler: Fehlende dynamische Seitennummern in der Fußleiste.");
                Console.ForegroundColor = ConsoleColor.Gray;
                Console.WriteLine("In Textverarbeitungsprogrammen können dynamische Seitennummern unter dem Menüpunkt Einfügen eingestellt werden.");
            }
        }
    }
}

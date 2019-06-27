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
                Console.WriteLine("Fehler: Das Dokument darf höchstens 15 Seiten lang sein. Die Seitenzahl beträgt " + numberOfPages +".");
                Console.ForegroundColor = ConsoleColor.Gray;
            } else
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
            } else
            {
                Console.WriteLine("Der rechte Rand ist korrekt formatiert.");
            }

            if (marginRight < 42 || marginRight > 43)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("Fehler: Der linke Rand muss 1,5cm betragen.");
                Console.ForegroundColor = ConsoleColor.Gray;
            } else
            {
                Console.WriteLine("Der linke Rand ist korrekt formatiert.");
            }
        }

        //Überprüfen, ob ein Inhaltsverzeichnis existiert
        public void checkTableOfContents()
        {
            TablesOfContents tbc = Document.TablesOfContents;
            if (tbc.Count < 1) {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("Fehler: Kein Inhaltsverzeichnis vorhanden.");
                Console.ForegroundColor = ConsoleColor.Gray;
            }
        }
        /*
        public void checkPageNumbers(Document docs)
        {
            for (int i = 1; i <= docs.Sections.Count; i++)
            {

                try
                {
                    Section section = docs.Sections[i];
                    if (section != null)
                    {
                        headOrFooter = section.Footers[.WdHeaderFooterIndex.wdHeaderFooterPrimary];
                        hasNumberPages = HeaderOrFooterHasPageNumber(headOrFooter);
                        if (hasNumberPages)
                            break;


                        headOrFooter = section.Footers[Word.WdHeaderFooterIndex.wdHeaderFooterEvenPages];
                        hasNumberPages = HeaderOrFooterHasPageNumber(headOrFooter);
                        if (hasNumberPages)
                            break;

                        headOrFooter = section.Footers[Word.WdHeaderFooterIndex.wdHeaderFooterFirstPage];
                        hasNumberPages = HeaderOrFooterHasPageNumber(headOrFooter);
                        if (hasNumberPages)
                            break;

                        headOrFooter = section.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary];
                        hasNumberPages = HeaderOrFooterHasPageNumber(headOrFooter);
                        if (hasNumberPages)
                            break;

                        headOrFooter = section.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterEvenPages];
                        hasNumberPages = HeaderOrFooterHasPageNumber(headOrFooter);
                        if (hasNumberPages)
                            break;

                        headOrFooter = section.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterFirstPage];
                        hasNumberPages = HeaderOrFooterHasPageNumber(headOrFooter);
                        if (hasNumberPages)
                            break;
                    }
                }
        }
        */

        //Die Methode prüft, ob Fußzeilen existieren. Die erste Seite wird dabei ignoriert.
        public bool CheckFooter()
        {
            /*
            Unterscheidung nach Footer-Art
            */


            if (!FooterEven.Exists && !FooterFirst.Exists && !FooterPrimary.Exists)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.Write("Fehler: Keine Fußzeile vorhanden: ");
                Console.ForegroundColor = ConsoleColor.Gray;
                Console.WriteLine("Die Seitennummern sollen dynamisch in einer Fußzeile generiert werden.");
                return false;
            }
            return true;
        }
        
        //Überprüft, ob dynamische Seitennummern vorhanden sind
        public void CheckPageNumbers()
        {
            /*
             * Differenzierung zwischen unterschiedlichen Footern einfügen
             */
            int pageNumbersEven = FooterEven.PageNumbers.Count;
            int pageNumbersFirst = FooterFirst.PageNumbers.Count;
            int pageNumbersPrimary = FooterPrimary.PageNumbers.Count;

            if (!(pageNumbersEven > 0) && !(pageNumbersFirst > 0) || !(pageNumbersPrimary > 0))
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.Write("Fehler: Keine dynamischen Seitennummern in der Fußleiste vorhanden.");
                Console.ForegroundColor = ConsoleColor.Gray;
            }
        }
    }
}

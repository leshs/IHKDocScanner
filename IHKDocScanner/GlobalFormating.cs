﻿using Microsoft.Office.Interop.Word;
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
        public GlobalFormating(Document document)
        {
            Rng = document.Range();
            PSetup = Rng.PageSetup;
            Document = document;
        }

        //Seitenanzahl überüfen
        public void checkPageCount()
        {
            int numberOfPages = Rng.get_Information(WdInformation.wdNumberOfPagesInDocument);
            if (numberOfPages > 15)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("Fehler: Das Dokument darf höchstens 15 Seiten lang sein. Die Seitenzahl beträgt " + numberOfPages);
                Console.ForegroundColor = ConsoleColor.White;
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
                Console.ForegroundColor = ConsoleColor.White;
            } else
            {
                Console.WriteLine("Der rechte Rand ist korrekt formatiert");
            }

            if (marginRight < 42 || marginRight > 43)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("Fehler: Der linke Rand muss 1,5cm betragen.");
                Console.ForegroundColor = ConsoleColor.White;
            } else
            {
                Console.WriteLine("Der linke Rand ist korrekt formatiert");
            }
        }

        //Überprüfen, ob ein Inhaltsverzeichnis existiert
        public void checkTableOfContents()
        {
            TablesOfContents tbc = Document.TablesOfContents;
            if (tbc.Count < 1) {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("Fehler: Kein Inhaltsverzeichnis vorhanden");
                Console.ForegroundColor = ConsoleColor.White;
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
        

    }
}

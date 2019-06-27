using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace IHKDocScanner
{
    class TextFormatting
    {
        private Document Doc;
        private Paragraphs Par;
        private Range DocRng;
        private int ParagraphNumber;
        private int PageNumber;
        private Paragraph Paragraph;
        public TextFormatting(Document doc)
        {
            Doc = doc;
            DocRng = doc.Range();
            Par = doc.Paragraphs;
        }

        public void SetParagraphNumber(int parNumber)
        {
            this.ParagraphNumber = parNumber;
        }

        public void SetParagraph(Paragraph paragraph)
        {
            this.Paragraph = paragraph;
        }

        public void SetPageNumber(int pageNumber)
        {
            this.PageNumber = pageNumber;
        }

        //Überprüft den Zeilenabstand
        public void CheckLineSpacing()
        {
                float lineSpace = Paragraph.LineSpacing;
                if(lineSpace != 18)
                {
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine("Fehler: Der Zeilenabstand von Absatz " + ParagraphNumber + " auf Seite " + PageNumber + " beträgt nicht 1,5");
                    Console.ForegroundColor = ConsoleColor.White;
                }
        }

        //Schriftart überprüfen
        public void CheckFont()
        {
                Range parRng = Paragraph.Range;
                    if (parRng.Font.Name != "Arial")
                    {
                        Console.ForegroundColor = ConsoleColor.Red;
                        Console.WriteLine("Fehler: Die Schriftart von Absatz " + ParagraphNumber + " auf Seite " + PageNumber + " ist nicht Arial");
                        Console.ForegroundColor = ConsoleColor.White;
                    }
            
        }

        public void CheckFontSize()
        {
            StyleCheck sc = new StyleCheck(Doc);
                Range parRng = Paragraph.Range;
                if (sc.CheckHead(Paragraph, PageNumber, ParagraphNumber))
                return;
                
                if (parRng.Font.Size != 12)
                {
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine("Fehler: Die Schriftgröße von Absatz " + ParagraphNumber + " auf Seite " + PageNumber + " ist nicht 12");
                    Console.ForegroundColor = ConsoleColor.White;
                }          
        }
    }
}

using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace IHKDocScanner
{
    class ParagraphFormatting
    {
        private int ParagraphNumber;
        private int PageNumber;
        private Paragraph Paragraph;
        private Range Rng;
        private PageSetup PSetup;

        public ParagraphFormatting(Document document)
        {
            Rng = document.Range();
            PSetup = Rng.PageSetup;
        }

        public void SetClassAttributes(int parNumber, Paragraph paragraph, int pageNumber)
        {
            this.ParagraphNumber = parNumber;
            this.Paragraph = paragraph;
            this.PageNumber = pageNumber;
        }

        //Überprüft, ob es sich um eine Überschrift handelt. Wenn ja, werden die übrigen Formatierungs-Prüfungen ignoriert.
        //Der Style der Überschrift wird auf der Konsole ausgegeben.
        public bool CheckHeading()
        {
            Style style = Paragraph.get_Style();
            
            String styleName = style.NameLocal;
            String[] styleNameArr = styleName.Split(' ');
            if(styleName == "Standard")
            {
                return false;
            } else if(styleNameArr[0] == "Verzeichnis" || styleNameArr[0] == "Inhaltsverzeichnisüberschrift")
            {
                Console.WriteLine("Absatz " + ParagraphNumber + " auf Seite " + PageNumber + " ist eine Überschrift im Inhaltsverzeichnis. Style: " + styleName);
                return true;
            } else if(styleNameArr[0] == "Überschrift")
            {
                Console.WriteLine("Absatz " + ParagraphNumber + " auf Seite "+ PageNumber + " ist eine Überschrift. Style: " + styleName);
                return true;
            } else
            {
                return false;
            }
        }

        public void CheckWidow()
        {
            if (Paragraph.WidowControl == 0)
            {
                Console.ForegroundColor = ConsoleColor.Yellow;
                Console.WriteLine("Warnung: Die Absatzkontrolle bei Absatz " + ParagraphNumber + " auf Seite " + PageNumber + " ist deaktiviert");
                Console.ForegroundColor = ConsoleColor.Gray;
            }
        }
    }
}

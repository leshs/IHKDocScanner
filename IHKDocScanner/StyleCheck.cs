using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace IHKDocScanner
{
    class StyleCheck
    {
        private Range Rng;
        private PageSetup PSetup;
        public StyleCheck(Document document)
        {
            Rng = document.Range();
            PSetup = Rng.PageSetup;
        }
        public bool CheckHead(Paragraph paragraph, int pageNumber, int paragraphNumber)
        {
            Style style = paragraph.get_Style();
            String styleName = style.NameLocal;
            if(styleName == "Standard")
            {
                return false;
            } else
            {
                Console.WriteLine("Absatz " + paragraphNumber + " auf Seite "+ pageNumber + " ist eine Überschrift. Style: " + styleName);
                return true;
            }
        }
    }
}

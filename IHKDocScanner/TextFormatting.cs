﻿using Microsoft.Office.Interop.Word;
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
        private bool ShowHinweise;
        private int Errors = 0;
        private int Notifications = 0;

        public TextFormatting(Document doc)
        {
            Doc = doc;
            DocRng = doc.Range();
            Par = doc.Paragraphs;
        }

        public int GetErrors()
        {
            return Errors;
        }

        public int GetNotifications()
        {
            return Notifications;
        }
        public void SetClassAttributes(int parNumber, Paragraph paragraph, int pageNumber)
        {
            ParagraphNumber = parNumber;
            Paragraph = paragraph;
            PageNumber = pageNumber;
        }

        public void SetShowHinweise(bool showHinweise)
        {
            ShowHinweise = showHinweise;
        }

        //Überprüft den Zeilenabstand
        public void CheckLineSpacing()
        {
            float lineSpace = Paragraph.LineSpacing;
            if (lineSpace != 18)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("Fehler: Der Zeilenabstand von Absatz " + ParagraphNumber + " auf Seite " + PageNumber + " beträgt nicht 1,5");
                Console.ForegroundColor = ConsoleColor.Gray;
                Errors++;
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
                Console.ForegroundColor = ConsoleColor.Gray;
                Errors++;
            }
        }

        //Überprüft die Schriftgröße
        public void CheckFontSize()
        {
            Range parRng = Paragraph.Range;
            if (parRng.Font.Size != 12)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("Fehler: Die Schriftgröße von Absatz " + ParagraphNumber + " auf Seite " + PageNumber + " ist nicht 12");
                Console.ForegroundColor = ConsoleColor.Gray;
                Errors++;
            }
        }

        //Überprüft die Formatierung: Fett, unterstrichen, kursiv; Wenn dies vorkommt, wird eine Warnung ausgegeben.
        public void CheckWordFormat()
        {
            bool bold = false;
            bool italic = false;
            bool underline = false;
            Range rng = Paragraph.Range;

            foreach (Range word in rng.Words)
            {
                if (word.Bold == -1)
                {
                    bold = true;
                }
                else if (word.Italic == -1)
                {
                    italic = true;
                }
                else if (((WdUnderline)word.Underline) == WdUnderline.wdUnderlineSingle)
                {
                    underline = true;
                }
            }

            Console.ForegroundColor = ConsoleColor.Yellow;
            String message = "Hinweis: In Absatz " + ParagraphNumber + " auf Seite " + PageNumber + " befinden sich Wörter mit der Formatierung:";
            if (bold)
                message = message + " fett";
            if (italic)
                message = message + " kursiv";
            if (underline)
                message = message + " unterstrichen";
            if (bold || italic || underline)
            {
                if(ShowHinweise)
                {
                    Console.WriteLine(message);
                    Notifications++;
                }
            }

            Console.ForegroundColor = ConsoleColor.Gray;
        }
    }
}

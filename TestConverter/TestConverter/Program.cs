using System;
using System.Windows.Forms;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Word = Microsoft.Office.Interop.Word;
using System.Collections;
using System.Text.RegularExpressions;
using System.Xml.Linq;

namespace TestConverter
{
    class Program
    {
        [STAThread]
        static void Main(string[] args)
        {
            
            string finalHTML;
            List<string> listString = new List<string>();
            object True = true;
            Word.Application app = new Word.Application();
            Word.Document doc = new Word.Document();
            string fileName = @"D:\Programming\C#\Word To HTML Converter\hello world.docx";
            doc = app.Documents.Open(fileName, Type.Missing, True);

            //foreach (Word.Paragraph para in doc.Paragraphs)
            //{
            //    para.Range.Copy();
            //    stripClasses(Clipboard.GetText(TextDataFormat.Html));
            //}
            Word.ListParagraphs listpara = doc.ListParagraphs;
            IEnumerator ienum = doc.ListParagraphs.GetEnumerator();
            List<Tuple<string, Word.WdListType, int>> items = new List<Tuple<string, Word.WdListType, int>>();

            while (ienum.MoveNext())
            {
                Word.Range r = ((Word.Paragraph)ienum.Current).Range;
                items.Add(new  Tuple<string, Word.WdListType, int>(r.Text, r.ListFormat.ListType, r.ListFormat.ListLevelNumber));
            }

            items.Reverse();
            listString = createList(items);
            doc.SelectAllEditableRanges();
            doc.Range().Copy();

            string returnHTMLText = null;




            if (Clipboard.ContainsText(TextDataFormat.Html))
            {
                Console.WriteLine("html");
                returnHTMLText = Clipboard.GetText(TextDataFormat.Html);
                //Console.WriteLine(returnHTMLText);
                stripClasses(returnHTMLText, listString);

            }
            else
            {
                Console.WriteLine("no html");
                //Console.WriteLine(doc.);
                //returnHTMLText = Clipboard.GetText(TextDataFormat.Html);                
            }

            doc.Close();
            app.Quit();
            Console.WriteLine("closed");
            while (true) ;

        }

        private static List<string> createList(List<Tuple<string, Word.WdListType, int>> items)
        {
            int currentIndex = 1;
            List<string> listStrings = new List<string>();
            StringBuilder str = new StringBuilder();
            Word.WdListType currentType = items[0].Item2;
            
            str.Append(closeOpenStyle(items[0].Item2, true));
            foreach (Tuple<string, Word.WdListType, int> item in items)
            {
                //check if the list type has changed. Append the close tag and the open tag
                if (currentType != item.Item2)
                {
                    str.Append(closeOpenStyle(currentType, false));
                    listStrings.Add(str.ToString());
                    str.Clear();
                    str.Append(closeOpenStyle(item.Item2, true));
                    currentType = item.Item2;
                }

                //check the level
                if (item.Item3 > currentIndex)
                {
                    str.Append(closeOpenStyle(currentType, true));
                    currentIndex = item.Item3;
                }
                else if (item.Item3 < currentIndex)
                {
                    str.Append(closeOpenStyle(currentType, false));
                    currentIndex = item.Item3;
                }

                //append the list item
                str.Append("<li>" + item.Item1.Replace("\r","") + "</li>\n");                
            }

            str.Append(closeOpenStyle(currentType, false));
            listStrings.Add(str.ToString());
            return listStrings;
        }

        private static string closeOpenStyle(Word.WdListType listType, bool isOpening)
        {
            if (listType == Word.WdListType.wdListBullet)
            {
                if (isOpening)
                    return "<ul>\n";
                else
                    return "</ul>\n";
            }
            else
            {
                if (isOpening)
                    return "<ol>\n";
                else
                    return "</ol>\n";                
            }
        }

        public static string stripClasses(string html, List<string> listItems)
        {
            StringBuilder finalhtml = new StringBuilder(html);
            finalhtml.Remove(html.IndexOf("<!--EndFragment-->"), html.Length - 1 - html.IndexOf("<!--EndFragment-->"));
            finalhtml.Remove(0, html.IndexOf("<!--StartFragment-->") + "<!--StartFragment-->".Length);

            //remove <o:b> tags
            finalhtml.Replace("<o:p>", "");
            finalhtml.Replace("</o:p>", "");


            
            

            while (finalhtml.ToString().IndexOf("<p class=") != -1)
            {
                int start = finalhtml.ToString().IndexOf("<p class=");
                int end = finalhtml.ToString().IndexOf(">", start);

                finalhtml.Remove(start, end - start);
                finalhtml.Insert(start, "<p");
            }

            while (finalhtml.ToString().IndexOf("<b style=") != -1)
            {
                int start = finalhtml.ToString().IndexOf("<b style=");
                int end = finalhtml.ToString().IndexOf(">", start);

                finalhtml.Remove(start, end - start);
                finalhtml.Insert(start, "<b");
            }

            while (finalhtml.ToString().IndexOf("<i style=") != -1)
            {
                int start = finalhtml.ToString().IndexOf("<i style=");
                int end = finalhtml.ToString().IndexOf(">", start);

                finalhtml.Remove(start, end - start);
                finalhtml.Insert(start, "<i");
            }

            while (finalhtml.ToString().IndexOf("<table class=") != -1)
            {
                int start = finalhtml.ToString().IndexOf("<table class=");
                int end = finalhtml.ToString().IndexOf(">", start);

                finalhtml.Remove(start, end - start);
                finalhtml.Insert(start, "<table");
            }

            while (finalhtml.ToString().IndexOf("<tr style=") != -1)
            {
                int start = finalhtml.ToString().IndexOf("<tr style=");
                int end = finalhtml.ToString().IndexOf(">", start);

                finalhtml.Remove(start, end - start);
                finalhtml.Insert(start, "<tr");
            }

            while (finalhtml.ToString().IndexOf("<td width=") != -1)
            {
                int start = finalhtml.ToString().IndexOf("<td width=");
                int end = finalhtml.ToString().IndexOf(">", start);

                finalhtml.Remove(start, end - start);
                finalhtml.Insert(start, "<td");
            }

            //replace the lists
            int listItemPos = 0;

            while (finalhtml.ToString().IndexOf("<p><![if !supportLists]>") != -1)
            {
                //check how many items there are in the list (this tells us how many paragraphs to get rid of
                int itemCount = Regex.Matches(listItems[listItemPos], @"<li>").Count;
                //int itemCount = listItems[listItemPos].Count(f => f.Equals("<li>"));
                
                int mainStart = -1;
                //remove that many blocks of list items.
                for (int i = 0; i < itemCount; i++)
                {
                    int start = finalhtml.ToString().IndexOf("<p><![if !supportLists]>");
                    int end = finalhtml.ToString().IndexOf("</p>", start);
                    if (mainStart == -1)
                    {
                        mainStart = start;
                    }
                    finalhtml.Remove(start, end - start + "</p>".Length + 4);
                }

                //insert the new string
                finalhtml.Insert(mainStart, listItems[listItemPos]);
                listItemPos++;
                mainStart = -1;
            }

            finalhtml.Replace("<b>", "<strong>");
            finalhtml.Replace("</b>", "</strong>");

            finalhtml.Replace("<i>", "<em>");
            finalhtml.Replace("</i>", "</em>");

            finalhtml.Replace("<p>&nbsp;</p>", "");
            Console.WriteLine();
            Console.WriteLine();
            Console.WriteLine();
            Console.WriteLine();
            Console.WriteLine(finalhtml);

            //XDocument doc = new XDocument();
            //doc.Add(finalhtml.ToString().Trim().Replace("\n",""));
            finalhtml = new StringBuilder(finalhtml.ToString().Trim());
            finalhtml.Insert(0, "<body>");
            finalhtml.Insert(finalhtml.Length, "</body>");
            //HtmlAgilityPack.HtmlDocument doc = new HtmlAgilityPack.HtmlDocument();
            //doc.LoadHtml(finalhtml.ToString());
            
            XDocument doc = XDocument.Parse(finalhtml.ToString());
            Clipboard.SetText(finalhtml.ToString().Trim());
            return finalhtml.ToString().Trim();
        }
    }
}

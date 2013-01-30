using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Word=Microsoft.Office.Interop.Word;
using System.Xml.Linq;
using System.Windows;
using System.Collections;
using System.Text.RegularExpressions;
namespace Word_To_Markup_Converter.Module
{
    public abstract class MarkupGenerator
    {
        protected string finalHTML;
        protected string boldTagStart;
        protected string boldTagEnd;
        protected string italicTagStart;
        protected string italicTagEnd;
        protected string h1TagStart;
        protected string h1TagEnd;
        protected string h2TagStart;
        protected string h2TagEnd;
        protected string h3TagStart;
        protected string h3TagEnd;
        protected string h4TagStart;
        protected string h4TagEnd;
        protected string h5TagStart;
        protected string h5TagEnd;
        protected string pTagStart;
        protected string pTagEnd;
                 
        [STAThread]
        public string generateMarkup(string fileName)
        {
            Word.Application app = new Word.Application();
            Word.Document doc = new Word.Document();
            
            doc = app.Documents.Open(fileName, Type.Missing, true);
            Word.ListParagraphs listpara = doc.ListParagraphs;
            IEnumerator ienum = doc.ListParagraphs.GetEnumerator();

            List<string> listString = new List<string>();            
            
            
            List<Tuple<string, Word.WdListType, int>> items = new List<Tuple<string, Word.WdListType, int>>();

            while (ienum.MoveNext())
            {
                Word.Range r = ((Word.Paragraph)ienum.Current).Range;
                items.Add(new  Tuple<string, Word.WdListType, int>(r.Text, r.ListFormat.ListType, r.ListFormat.ListLevelNumber));
            }

            items.Reverse();
            if (items.Count > 0)
            {
                listString = createList(items);
            }               
            doc.SelectAllEditableRanges();
            doc.Range().Copy();

            string returnHTMLText = null;




            if (Clipboard.ContainsText(TextDataFormat.Html))
            {
                Console.WriteLine("html");
                returnHTMLText = Clipboard.GetText(TextDataFormat.Html);
                //Console.WriteLine(returnHTMLText);
                finalHTML = stripClasses(returnHTMLText, listString);
                

            }
            else
            {
                Console.WriteLine("no html");
                returnHTMLText = "";
                //Console.WriteLine(doc.);
                //returnHTMLText = Clipboard.GetText(TextDataFormat.Html);                
            }

            doc.Close();
            app.Quit();
            Console.WriteLine("closed");
            return finalHTML;

        }

        private List<string> createList(List<Tuple<string, Word.WdListType, int>> items)
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

        private string closeOpenStyle(Word.WdListType listType, bool isOpening)
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

        private string stripClasses(string html, List<string> listItems)
        {
            StringBuilder formattedHTML = new StringBuilder(html);
            formattedHTML.Remove(html.IndexOf("<!--EndFragment-->"), html.Length - 1 - html.IndexOf("<!--EndFragment-->"));
            formattedHTML.Remove(0, html.IndexOf("<!--StartFragment-->") + "<!--StartFragment-->".Length);

            //remove <o:[whatever]> tags
            formattedHTML.Replace("<o:p>", "");
            formattedHTML.Replace("</o:p>", "");                    

            //start replacing all the weird formatting that word puts in. Iterate through each para and get rid of things only if the p doesn't contain <pre> tags
            //TODO: place the logic for working with pre tag based stuff.
            while (formattedHTML.ToString().IndexOf("<p class=") != -1)
            {
                int start = formattedHTML.ToString().IndexOf("<p class=");
                int end = formattedHTML.ToString().IndexOf(">", start);
                formattedHTML.Replace("\r\n", " ", start, formattedHTML.ToString().IndexOf("</p>", start) + "</p>".Length + 1 - start);
                formattedHTML.Remove(start, end + 1 - start);
                formattedHTML.Insert(start, "<p>");                
                while (formattedHTML.ToString().IndexOf("<b style=") != -1)
                {
                    start = formattedHTML.ToString().IndexOf("<b style=");
                    end = formattedHTML.ToString().IndexOf(">", start);

                    formattedHTML.Remove(start, end + 1 - start);
                    formattedHTML.Insert(start, boldTagStart);
                    formattedHTML.Replace("</b>", boldTagEnd, start, formattedHTML.ToString().IndexOf("</b>", start) + "</b>".Length + 1 - start);
                }

                while (formattedHTML.ToString().IndexOf("<i style=") != -1)
                {
                    start = formattedHTML.ToString().IndexOf("<i style=");
                    end = formattedHTML.ToString().IndexOf(">", start);

                    formattedHTML.Remove(start, end + 1 - start);
                    formattedHTML.Insert(start, "<i>");
                    formattedHTML.Insert(start, italicTagStart);
                    formattedHTML.Replace("</i>", italicTagEnd, start, formattedHTML.ToString().IndexOf("</i>", start) + "</i>".Length + 1 - start);
                }
                
                //formattedHTML.Replace("<b>", boldTagStart);
                //formattedHTML.Replace("</b>", boldTagEnd);

                
            }            

            while (formattedHTML.ToString().IndexOf("<table class=") != -1)
            {
                int start = formattedHTML.ToString().IndexOf("<table class=");
                int end = formattedHTML.ToString().IndexOf(">", start);

                formattedHTML.Remove(start, end - start);
                formattedHTML.Insert(start, "<table");

                while (formattedHTML.ToString().IndexOf("<tr style=") != -1)
                {
                    start = formattedHTML.ToString().IndexOf("<tr style=");
                    end = formattedHTML.ToString().IndexOf(">", start);

                    formattedHTML.Remove(start, end - start);
                    formattedHTML.Insert(start, "<tr");
                }

                while (formattedHTML.ToString().IndexOf("<td width=") != -1)
                {
                    start = formattedHTML.ToString().IndexOf("<td width=");
                    end = formattedHTML.ToString().IndexOf(">", start);

                    formattedHTML.Remove(start, end - start);
                    formattedHTML.Insert(start, "<td");
                }

            }

            

            //replace the lists
            int listItemPos = 0;

            while (formattedHTML.ToString().IndexOf("<p><![if !supportLists]>") != -1)
            {
                //check how many items there are in the list (this tells us how many paragraphs to get rid of
                int itemCount = Regex.Matches(listItems[listItemPos], @"<li>").Count;
                //int itemCount = listItems[listItemPos].Count(f => f.Equals("<li>"));
                
                int mainStart = -1;
                //remove that many blocks of list items.
                for (int i = 0; i < itemCount; i++)
                {
                    int start = formattedHTML.ToString().IndexOf("<p><![if !supportLists]>");
                    int end = formattedHTML.ToString().IndexOf("</p>", start);
                    if (mainStart == -1)
                    {
                        mainStart = start;
                    }
                    formattedHTML.Remove(start, end - start + "</p>".Length);
                }

                //insert the new string
                formattedHTML.Insert(mainStart, listItems[listItemPos]);
                listItemPos++;
                mainStart = -1;
            }

            

            formattedHTML.Replace("<p>&nbsp;</p>", "");
            Console.WriteLine();
            Console.WriteLine();
            Console.WriteLine();
            Console.WriteLine();
            Console.WriteLine(formattedHTML);           
            
            //Clipboard.SetText(doc.ToString().Replace("<br />", "<br>").Trim());
            return formattedHTML.ToString().Replace("<br />", "<br>").Trim();
        }
        
    }
}

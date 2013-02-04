using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Word = Microsoft.Office.Interop.Word;
using System.Xml.Linq;
namespace Word_To_Markup_Converter.Module
{
    public class HTMLGenerator : MarkupGenerator
    {
        public HTMLGenerator() : base()
        {
            boldTagStart = "<strong>";
            boldTagEnd = "</strong>";
            italicTagStart = "<em>";
            italicTagEnd = "</em>";
            h1TagStart = "<h1>";
            h1TagEnd = "</h1>";
            h2TagStart = "<h2>";
            h2TagEnd = "</h2>";
            h3TagStart = "<h3>";
            h4TagEnd = "</h3>";
            pTagStart = "<p>";
            pTagEnd = "</p>";
        }
        public string generateMarkup(string fileName, string headerFile, string footerFile, string title)
        {
            
            base.generateMarkup(fileName);
            StringBuilder html = new StringBuilder(finalHTML);

            html.Insert(0, "<body>");
            html.Insert(html.Length, "</body>");
            XDocument xdoc = XDocument.Parse(html.ToString().Replace("\v", "<br />").Replace("\r", ""));
            finalHTML = xdoc.ToString().Trim();

            Word.Application app = new Word.Application();
            Word.Document doc = new Word.Document();
            if (headerFile != "")
            {
                doc = app.Documents.Open(headerFile, Type.Missing, true);
                html.Insert(0, doc.Range().Text);
                doc.Close();
            }
            else
            {
                html.Insert(0, String.Format(@"
<!DOCTYPE html>
 <head>
  <title> {0} </title>
 </head>
 ", title));
            }

            if (footerFile != "")
            {
                doc = app.Documents.Open(footerFile, Type.Missing, true);
                html.Insert(html.Length, doc.Range().Text);
                doc.Close();
            }
            else
            {
                html.Insert(html.Length, @"

</html>");
            }
            app.Quit();
            return finalHTML;
        }
    }
}

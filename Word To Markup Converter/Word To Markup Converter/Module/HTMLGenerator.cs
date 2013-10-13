using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using System.IO;
using System.Text.RegularExpressions;

namespace Word_To_Markup_Converter.Module
{
    public class HTMLGenerator : MarkupGenerator
    {
        public HTMLGenerator()
        {
            pTag = new Tuple<string, string>("<p>", "</p>\n");

            boldTag = new Tuple<string, string>("<strong>", "</strong>");
            italicTag = new Tuple<string, string>("<em>", "</em>");

            unorderedListTag = new Tuple<string, string>("<ul>\n", "</ul>\n");
            orderedListTag = new Tuple<string, string>("<ol>\n", "</ol>\n");

            unorderedListItemTag = orderedListItemTag = new Tuple<string, string>("<li>", "</li>\n");

            header1Tag = new Tuple<string, string>("<h1>", "</h1>\n");
            header2Tag = new Tuple<string, string>("<h2>", "</h2>\n");
            header3Tag = new Tuple<string, string>("<h3>", "</h3>\n");
            header4Tag = new Tuple<string, string>("<h4>", "</h4>\n");
            header5Tag = new Tuple<string, string>("<h5>", "</h5>\n");
            header6Tag = new Tuple<string, string>("<h6>", "</h6>\n");
        }

        public void generateMarkup(String documentPath, String headerPath, String footerPath, String documentTitle)
        {
            base.generateMarkup(documentPath);
            StreamReader headerReader = new StreamReader(headerPath);
            StreamReader footerReader = new StreamReader(footerPath);

            docText.Insert(0, headerReader.ReadToEnd()).Append(footerReader.ReadToEnd());
            createTitle(documentTitle);
        }

        public void generateMarkup(String documentPath, String documentTitle)
        {
            base.generateMarkup(documentPath);
            docText.Insert(0, Properties.Resources.DEFAULT_HEADER).Append(Properties.Resources.DEAFULT_FOOTER);
            createTitle(documentTitle);
        }

        protected void createTitle(String documentTitle)
        {
            docText = new StringBuilder(Regex.Replace(docText.ToString(), @"<title>.*<\/title>", "<title>" + documentTitle + "</title>"));
        }

        protected override void formatLink(StringBuilder textToAppend, string link)
        {
            textToAppend.Insert(0, "<a href=\"" + link + "\">");
            textToAppend.Append("</a>");
        }
    }
}

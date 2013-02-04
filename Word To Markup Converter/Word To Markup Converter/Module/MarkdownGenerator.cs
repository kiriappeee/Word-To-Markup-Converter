using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Word_To_Markup_Converter.Module
{
    public class MarkdownGenerator : MarkupGenerator
    {
        public MarkdownGenerator() : base()
        {
            boldTagStart = boldTagEnd = "*";
            italicTagStart = italicTagEnd = "_";
            h1TagStart = h1TagEnd = "#";
            h2TagStart = h2TagEnd = "##";
            h3TagStart = h3TagEnd = "###";
            pTagStart = "\n";
            pTagEnd = "";
            ulTagStart = "\n";
            ulTagEnd = "";
            uliTagStart = "\t* ";
            uliTagEnd = "";
        }


        public string generateMarkup(string fileName)
        {
            base.generateMarkup(fileName);
            StringBuilder markdown = new StringBuilder(finalHTML);
            return markdown.ToString();
        }
    }
}

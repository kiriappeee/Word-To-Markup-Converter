using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace Word_To_Markup_Converter.Module
{
    public class MarkdownGenerator : MarkupGenerator
    {
        public MarkdownGenerator()
        {
            pTag = new Tuple<string, string>("\n", "\n");

            boldTag = new Tuple<string, string>("**", "**");
            italicTag = new Tuple<string, string>("_", "_");

            unorderedListTag = new Tuple<string, string>("\n", "\n");
            orderedListTag = new Tuple<string, string>("\n", "\n");

            unorderedListItemTag = new Tuple<string, string>("* ", "\n");
            orderedListItemTag = new Tuple<string, string>("1. ", "\n");

            header1Tag = new Tuple<string, string>("#", "#\n");
            header2Tag = new Tuple<string, string>("##", "##\n");
            header3Tag = new Tuple<string, string>("###", "###\n");
            header4Tag = new Tuple<string, string>("####", "####\n");
            header5Tag = new Tuple<string, string>("#####", "#####\n");
            header6Tag = new Tuple<string, string>("######", "######\n");
        }


        protected override void formatListItemOpener(StringBuilder textToAppend, int currentListLevel, int currentListType)
        {
            string tabLevel = new string('\t', currentListLevel + 1);
            if (currentListType == LIST_TYPE_UNORDERED)
            {
                textToAppend.Insert(0, unorderedListItemTag.Item1);
            }
            else if (currentListType == LIST_TYPE_ORDERED)
            {
                textToAppend.Insert(0, orderedListItemTag.Item1);
            }
        }

        protected override void formatItalic(StringBuilder textToAppend)
        {

            textToAppend.Insert(Regex.Match(textToAppend.ToString(), @"[^\s]").Index, italicTag.Item1);
            string whiteSpaceEnd = new string(' ', textToAppend.ToString().Length - textToAppend.ToString().TrimEnd().Length);
            string temp = textToAppend.ToString().TrimEnd();
            textToAppend.Clear().Append(temp);
            textToAppend.Append(italicTag.Item2).Append(whiteSpaceEnd);
        }

        protected override void formatBold(StringBuilder textToAppend)
        {
            textToAppend.Insert(Regex.Match(textToAppend.ToString(), @"[^\s]").Index, boldTag.Item1);
            string whiteSpaceEnd = new string(' ', textToAppend.ToString().Length - textToAppend.ToString().TrimEnd().Length);
            string temp = textToAppend.ToString().TrimEnd();
            textToAppend.Clear().Append(temp);
            textToAppend.Append(boldTag.Item2).Append(whiteSpaceEnd);
        }
    }
}

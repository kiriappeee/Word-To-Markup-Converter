using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using System.Windows;
using System.Collections;
using System.Text.RegularExpressions;
using System.Xml;
using Ionic.Zip;
using System.IO;
using System.Web;

namespace Word_To_Markup_Converter.Module
{
    public abstract class MarkupGenerator
    {
        public StringBuilder docText = new StringBuilder();

        protected Tuple<string, string> boldTag;
        protected Tuple<string, string> italicTag;
        protected Tuple<string, string> pTag;
        protected Tuple<string, string> header1Tag;
        protected Tuple<string, string> header2Tag;
        protected Tuple<string, string> header3Tag;
        protected Tuple<string, string> header4Tag;
        protected Tuple<string, string> header5Tag;
        protected Tuple<string, string> header6Tag;
        protected Tuple<string, string> unorderedListTag;
        protected Tuple<string, string> orderedListTag;
        protected Tuple<string, string> unorderedListItemTag;
        protected Tuple<string, string> orderedListItemTag;

        protected const int LIST_TYPE_UNORDERED = 0;
        protected const int LIST_TYPE_ORDERED = 1;

        public virtual void generateMarkup(String documentPath)
        {
                        
            string extractPath = documentPath + System.IO.Path.GetFileName(documentPath) + " open xml";

            ZipFile zip = ZipFile.Read(documentPath);
            foreach (ZipEntry e in zip)
            {
                e.Extract(extractPath, true);
            }

            string xmlDocPath = extractPath + "\\word\\document.xml";
            string xmlrefDocPath = extractPath + @"\word\_rels\document.xml.rels";
            string xmlnumrefPath = extractPath + @"\word\numbering.xml";
            int currentListLevel = -1;
            int currentListType = -1;
            Stack<Tuple<int, int>> listStack = new Stack<Tuple<int, int>>();

            XmlDocument doc = new XmlDocument();
            XmlDocument docref = new XmlDocument();
            XmlDocument numref = new XmlDocument();
            
            XmlNamespaceManager docNamespaceManager = new XmlNamespaceManager(doc.NameTable);
            XmlNamespaceManager docrefNameSpaceManager = new XmlNamespaceManager(docref.NameTable);
            XmlNamespaceManager numrefNameSpaceManager = new XmlNamespaceManager(numref.NameTable);
            
            docNamespaceManager.AddNamespace("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            docrefNameSpaceManager.AddNamespace("x", "http://schemas.openxmlformats.org/package/2006/relationships");
            numrefNameSpaceManager.AddNamespace("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

            doc.Load(xmlDocPath);
            docref.Load(xmlrefDocPath);
            
            if(File.Exists(xmlnumrefPath))
            {
                numref.Load(xmlnumrefPath);
            }
            
            // get the body of the document.
            XmlNode body = doc.SelectSingleNode("//w:body", docNamespaceManager);

            //get all the children of the body
            XmlNodeList bodyItems = body.ChildNodes;
            StringBuilder textToAppend = new StringBuilder();
            StringBuilder paraTextToAppend = new StringBuilder();
            foreach (XmlNode item in bodyItems)
            {
                //check if the item is a paragraph or a table
                if (item.LocalName == "p")
                {
                    buildBasicParagraph(item, docNamespaceManager, textToAppend, docref, docrefNameSpaceManager);

                    //check if we are dealing with a special kind of paragraph
                    if (item.SelectSingleNode("w:pPr", docNamespaceManager) != null)
                    {                                            
                        //logic for dealing with special paragraphs
                        if (item.SelectSingleNode("w:pPr/w:pStyle", docNamespaceManager) != null && item.SelectSingleNode("w:pPr/w:pStyle", docNamespaceManager).Attributes.GetNamedItem("w:val").Value.Contains("Heading"))
                        {                            
                            //logic for header
                            formatHeader(textToAppend, item.SelectSingleNode("w:pPr/w:pStyle", docNamespaceManager).Attributes.GetNamedItem("w:val").Value);
                            
                            while (listStack.Count != 0)
                            {
                                currentListLevel = listStack.Peek().Item1;
                                currentListType = listStack.Peek().Item2;
                                formatListCloser(textToAppend, currentListLevel, currentListType);
                                formatListItemCloser(textToAppend, currentListLevel, currentListType);
                                listStack.Pop();
                            }
                        }

                        //logic for dealing with list items
                        if (item.SelectSingleNode("w:pPr/w:pStyle", docNamespaceManager) != null && item.SelectSingleNode("w:pPr/w:pStyle", docNamespaceManager).Attributes.GetNamedItem("w:val").Value.Contains("ListParagraph"))
                        {
                            //get the index of the item being inserted. 
                            int insertListLevel = Convert.ToInt16(item.SelectSingleNode("w:pPr/w:numPr/w:ilvl", docNamespaceManager).Attributes.GetNamedItem("w:val").Value);
                            int insertListType = getListType(Convert.ToInt16(item.SelectSingleNode("w:pPr/w:numPr/w:numId", docNamespaceManager).Attributes.GetNamedItem("w:val").Value), numref, numrefNameSpaceManager);

                            //list insertion has not begun yet
                            if (listStack.Count == 0)
                            {
                                currentListLevel = insertListLevel;
                                currentListType = insertListType;
                                listStack.Push(new Tuple<int, int>(currentListLevel, currentListType));
                                formatListItemOpener(textToAppend, currentListLevel, currentListType);
                                formatListOpener(textToAppend, currentListLevel, currentListType);
                            }
                            else
                            {
                                //logic for when coming out of a sublist block
                                if (currentListLevel > insertListLevel)
                                {
                                    while (currentListLevel != insertListLevel || currentListType != insertListType)
                                    {
                                        formatListCloser(textToAppend, currentListLevel, currentListType);
                                        formatListItemCloser(textToAppend, currentListLevel, currentListType);
                                        listStack.Pop();
                                        currentListLevel = listStack.Peek().Item1;
                                        currentListType = listStack.Peek().Item2;
                                    }
                                }
                               
                                if (currentListLevel == insertListLevel)
                                {
                                    int tempListLevel = currentListLevel;
                                    int tempListType = currentListType;
                                    
                                    //check if the list is the same list or whether a new one is opening
                                    if (currentListType == insertListType)   //same list
                                    {
                                        formatListItemOpener(textToAppend, currentListLevel, currentListType);
                                    }
                                    else
                                    {
                                        formatListCloser(textToAppend, currentListLevel, currentListType);
                                        listStack.Pop();
                                        currentListLevel = insertListLevel;
                                        currentListType = insertListType;
                                        listStack.Push(new Tuple<int, int>(currentListLevel, currentListType));
                                    }
                                    formatListItemCloser(textToAppend, tempListLevel, tempListType);
                                }
                                else if (currentListLevel < insertListLevel)
                                {
                                    currentListLevel = insertListLevel;
                                    currentListType = insertListType;
                                    listStack.Push(new Tuple<int, int>(currentListLevel, currentListType));
                                    formatListItemOpener(textToAppend, currentListLevel, currentListType);
                                    formatListOpener(textToAppend, currentListLevel, currentListType);
                                }                                
                            }                            
                        }

                    }
                    else
                    {
                        while (listStack.Count != 0)
                        {
                            currentListLevel = listStack.Peek().Item1;
                            currentListType = listStack.Peek().Item2;
                            formatListCloser(textToAppend, currentListLevel, currentListType);
                            formatListItemCloser(textToAppend, currentListLevel, currentListType);
                            listStack.Pop();                            
                        }

                        formatParagraph(textToAppend);
                    }

                    
                    docText.Append(textToAppend.ToString());
                }
                else if (item.LocalName == "tbl")
                {
                    //insert the logic for dealing with any kind of table. The reason I didn't write a blanket else 
                    //is because I don't know if there could be any other kind of nodes
                }

                
            }
            textToAppend = new StringBuilder();
            while (listStack.Count != 0)
            {
                currentListLevel = listStack.Peek().Item1;
                currentListType = listStack.Peek().Item2;
                formatListItemCloser(textToAppend, currentListLevel, currentListType);
                formatListCloser(textToAppend, currentListLevel, currentListType);
                listStack.Pop();
            }
            docText.Append(textToAppend.ToString());
            Directory.Delete(extractPath, true);
        }

        protected virtual int getListType(short numid, XmlDocument doc, XmlNamespaceManager ns)
        {
            if (doc.SelectSingleNode("//w:abstractNum[@w:abstractNumId=\"" +
                doc.SelectSingleNode("//w:num[@w:numId=\"1\"]", ns).SelectSingleNode("w:abstractNumId", ns).Attributes.GetNamedItem("w:val").Value + "\"]", ns).SelectSingleNode
                ("w:lvl", ns).SelectSingleNode("w:numFmt", ns).Attributes.GetNamedItem("w:val").Value == "bullet")
                return LIST_TYPE_UNORDERED;
            else
                return LIST_TYPE_ORDERED;
        }

        

        #region logic for types of content

        protected virtual void buildBasicParagraph(XmlNode item, XmlNamespaceManager namespaceManager, StringBuilder textToAppend, XmlDocument docref, XmlNamespaceManager docrefNameSpaceManager)
        {
            StringBuilder paraTextToAppend = new StringBuilder();
            textToAppend.Clear();       //use a string builder here since there can be multiple tags to iterate through in a single paragraph. 
            XmlNodeList paraNodes = item.ChildNodes;    //hyperlinks are stored within a hyper link tag. therefore iterating through just w:r tags isn't good enough
            foreach (XmlNode childNode in paraNodes)
            {
                XmlNode textNode;
                String link = "";
                paraTextToAppend.Clear();
                if (childNode.LocalName == "hyperlink")
                {
                    textNode = childNode.SelectSingleNode("w:r", namespaceManager);
                    link = docref.SelectSingleNode("//x:Relationship[@Id='" + childNode.Attributes.GetNamedItem("r:id").Value + "']", docrefNameSpaceManager).Attributes.GetNamedItem("Target").Value;
                }
                else if (childNode.LocalName == "r")
                {
                    textNode = childNode;
                }
                else
                {
                    textNode = null;
                }

                if (textNode != null)
                {
                    //get the text within the particular block                
                    paraTextToAppend.Append(HttpUtility.HtmlEncode(textNode.SelectSingleNode("w:t", namespaceManager).InnerXml));
                    //search for any formatting and apply it
                    if (textNode.SelectSingleNode("w:rPr", namespaceManager) != null)
                    {
                        XmlNodeList styleNodes = textNode.SelectSingleNode("w:rPr", namespaceManager).ChildNodes;
                        foreach (XmlNode styleNode in styleNodes)
                        {
                            //this method can be extended in the future to incorporate any other styling that might come along such as strike through lines. 
                            if (styleNode.LocalName == "b")
                                formatBold(paraTextToAppend);
                            else if (styleNode.LocalName == "i")
                                formatItalic(paraTextToAppend);
                        }
                    }
                    if (!link.Equals(string.Empty))
                    {
                        formatLink(paraTextToAppend, link);
                    }
                    textToAppend.Append(paraTextToAppend.ToString());
                }
                
            }
        }

        protected virtual void paragraphLogic(StringBuilder textToAppend)
        {
            
        }

        #endregion


        #region format methods


        protected virtual void formatParagraph(StringBuilder textToAppend)
        {
            textToAppend.Insert(0, pTag.Item1);
            textToAppend.Append(pTag.Item2);
        }

        protected virtual void formatListItemOpener(StringBuilder textToAppend, int currentListLevel, int currentListType)
        {
            if (currentListType == LIST_TYPE_UNORDERED)
            {
                textToAppend.Insert(0, unorderedListItemTag.Item1);
            }
            else if (currentListType == LIST_TYPE_ORDERED)
            {
                textToAppend.Insert(0, orderedListItemTag.Item1);
            }
        }

        protected virtual void formatListItemCloser(StringBuilder textToAppend, int currentListLevel, int currentListType)
        {
            if (currentListType == LIST_TYPE_UNORDERED)
            {                
                textToAppend.Insert(0, unorderedListItemTag.Item2);
            }
            else if (currentListType == LIST_TYPE_ORDERED)
            {
                textToAppend.Insert(0, orderedListItemTag.Item2);
            }
        }

        protected virtual void formatListCloser(StringBuilder textToAppend, int insertListLevel, int listType)
        {
            if (listType == LIST_TYPE_UNORDERED)
            {
                textToAppend.Insert(0, unorderedListTag.Item2);
            }
            else if (listType == LIST_TYPE_ORDERED)
            {
                textToAppend.Insert(0, orderedListTag.Item2);
            }
        }

        protected virtual void formatListOpener(StringBuilder textToAppend, int insertListLevel, int listType)
        {
            if (listType == LIST_TYPE_UNORDERED)
            {
                textToAppend.Insert(0, unorderedListTag.Item1);
            }
            else if (listType == LIST_TYPE_ORDERED)
            {
                textToAppend.Insert(0, orderedListTag.Item1);
            }
        }

        protected virtual void formatHeader(StringBuilder textToAppend, string headerType)
        {
            switch(headerType)
            {
                case "Heading1":
                    textToAppend.Insert(0, header1Tag.Item1);
                    textToAppend.Append(header1Tag.Item2);
                    break;
                case "Heading2":
                    textToAppend.Insert(0, header2Tag.Item1);
                    textToAppend.Append(header2Tag.Item2);
                    break;
                case "Heading3":
                    textToAppend.Insert(0, header3Tag.Item1);
                    textToAppend.Append(header3Tag.Item2);
                    break;
                case "Heading4":
                    textToAppend.Insert(0, header4Tag.Item1);
                    textToAppend.Append(header4Tag.Item2);
                    break;
                case "Heading5":
                    textToAppend.Insert(0, header5Tag.Item1);
                    textToAppend.Append(header5Tag.Item2);
                    break;
                case "Heading6":
                    textToAppend.Insert(0, header6Tag.Item1);
                    textToAppend.Append(header6Tag.Item2);
                    break;
            }
        }

        protected virtual void formatItalic(StringBuilder textToAppend)
        {
            textToAppend.Insert(0, italicTag.Item1);
            textToAppend.Append(italicTag.Item2);
        }

        protected virtual void formatBold(StringBuilder textToAppend)
        {
            textToAppend.Insert(0, boldTag.Item1);
            textToAppend.Append(boldTag.Item2);
        }

        protected virtual void formatLink(StringBuilder textToAppend, String link)
        {

        }

        #endregion
    }       
    
}

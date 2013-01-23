using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Word;
namespace Word_To_Markup_Converter.Module
{
    public abstract class MarkupGenerator
    {
        //using the skeleton pattern to go through the word doc and call on several abstract methods specific to each markup type (bold, italic)
        public void generateMarkup(string fileName)
        { 
            
        }
    }
}

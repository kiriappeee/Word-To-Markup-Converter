
#Word to Markup Generator#

##Introduction and project status##

The word to markup converter is a software that automatically generates a file in the markup language of one’s choice from a (Microsoft office) word document. Currently, only doc, and docx formats are supported. While the project has moved beyond proof of concept after a complete core rewrite it remains unsuitable for production use. 

##Using the software##

The Word to Markup Generator is being designed (no I haven’t finished it yet) to be as user friendly as possible.  The steps to use it are given below

* Select the document you want under the document to convert field
* Select the type of markup you want
* Select the file you wish to save to
* _For HTML only._ Select the header and footer template. If these are not selected then a very basic header and footer section will be added before and after the body of the html file. 
* _For HTML only._ Input the title of the document. This will automatically search for the title field if an HTML header has been given and replace it. _**This feature is not yet built in. For now a title is specified only for the default header.**_


##Feature List (Done and todo)##

Features that are done

* Basic paragraph support 
* Formatting support for bold and italic (strong and em)
* List support. Multiple levels and unordered + ordered list support including nested lists. (this was actually a very big deal to complete)
* Heading support 


Features that are in the immediate pipeline

* Links within text (this is super super important)
* Tables (this is super important)
* Support for &lt;pre&gt; tags (this is important… You get the idea)


##Code status##

The code right now is badly in need of a bit of refactoring in the generateMarkup method since all the features are iterated through in there and over time as more word features get covered the method is just going to get longer and longer. This shouldn’t be too difficult at all and will probably be done over the next week or whenever I get to make the next few releases. 

Having said that, the code has also come a long way and is extremely easy to extend now and as such I will be adding more features and more markup languages very soon. How soon is not something I’m willing to predict but I just hope it will really be sooner than later. Most important though is that I get a documentation of sorts to explain how to extend the language capabilities of the software to include any other kind of markup. 

###Tests###

Dirty secret time. I’ve never really worked on a project seriously with tests so this might take a little longer to get done. Having said that I will probably start working on it by next week although it won’t be included with the commits until I know what the heck it is I’m doing. 

##License##

This software is licensed under the Microsoft Reciprocal License and a copy of the license has been provided. For a less restrictive license I will be providing a standard library at some point that can be integrated with any other software under either the Microsoft Public License or the MIT license. 

##Credits##

Interface – Mahapps.Metro project was used for the look and feel of the software

DotNetZip Library used for unpacking the word document files 

Both of the above code bases have been provided under the Microsoft Public License (Ms-PL)

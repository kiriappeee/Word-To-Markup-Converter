#Word to Markup Generator#

##Introduction and project status##


The word to markup converter is a software that automatically generates a file in the markup language of one’s choice from a (Microsoft office) word document. Currently, only doc, and docx formats are supported and the current status of the project was more a proof of concept. This is to be revamped soon to provide heavy extensibility as well as support for other word processing software based documents. 

##Using the software##


Please keep in mind that this software is currently in proof of concept and therefore should not be used in any kind of production or serious scenarios. The code base will most likely be re written to provide a better extendable model to handle other markup languages in the future. With that said, the software aims to be as user friendly as possible. Usage of software is as follows



<li>Select the document you want under the document to convert field</li>
<li>Select the type of markup you want</li>
<li>Select the file you wish to save to</li>
<li>For HTML only. Select the header and footer template. If these are not selected then a very basic header and footer section will be added before and after the body of the html file. </li>
<li>For HTML only. Input the title of the document. This will automatically search for the title field if an HTML header has been given and replace it. This feature is not yet built in. For now a title is specified only for the default header. </li>
</ul>
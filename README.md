# CobieConvertExample
# Simple project to demonstrate COBie -> CAFM import template
This demonstrates how to use Xbim CobieExpress to read a cobie file then load a CAFM system template file spreadsheet and populate it with data from the cobie file.
I have used the ParieSky testing data from the Dormitory project and included it in the Sample data.
The CAFM System template is one I have made up but is typical of an import spreadsheet.
This program demonstrates opening the CObie file with Xbim.CobieExpress which reads the data into an Object model.
We then open the Template file using Epplus but you can use other libraries such as Spire. 
We set the template file workbook we want to poplutate and then use Linq to get the data and write to each row's cell. 

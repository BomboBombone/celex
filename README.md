# Celex
A python Program used to automate the sorting of excel files.

## Why should I use Celex?
**Celex** is a program developed as an interacive interface to **automate** some tasks that would otherwise be extremely tedious and time-consuming.  
When working in any commercial department in a big or small company, it is often the case that orders come in **excel** format.  
Nonetheless every company uses a different standard to pass order informations to the managment (which tells the actual workers what to produce).  
This means that between a company and another there will be almost always a different standard to make or receive orders.  
To actually convert an order to a file ready to be sent to one's own managment sector it takes time and effort, filtering useful information to be passed on  
and sorting it out.  

## This is where Celex comes in
**Celex** automates the filtering process, giving you the freedom to filter files by name, or content (useful for messy people like me), then choose  
which columns are of your interest, them being already included in the source file, or to be created in the output file.  
You can also specify a dictionary of values that need to be substituted in the process of filtering the excel file (particularly useful to automate  
value translation, for example in orders where multiple words can be used to indicate the same element, and substituting them all to a single one makes  
it easier to process it for the managment)  


# Current features:
Filter files by name or content.  
Open one or multiple selected files in the specified editor.  
Set keywords to look for a "value keyword" or "keyword value" string. Ex: keyword = 'MG' -> string = "100x40x60 MG" or string = "MG 100x40x60"  
Set column names to filter for or to generate. Ex: col = ["Prices", "Quantity"] / input file col = ["Prices", "Measures"] -> output file col = ["Prices", "Quantity"]

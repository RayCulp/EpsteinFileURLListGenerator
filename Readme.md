# What the macro in this file does

Takes a list of unresolved base URLs, a list of possible extensions, and generates a list of href URLs that can be saved as an HTML file and processed using a download manager like DownThemAll

# How to use

1.	Save the workbook locally
2.	Paste a list of unresolved base URLs into the sheet "unresolved"
3.	Paste a list of extensions you would like to try in the sheet "extensions"
4.	Run macro "GenerateURLList"
5.	Go drink favorite beverage. This will take a while.
6.	When list of URLs has finished generating, make sure "merged" sheet is selected, then save as CSV. Leave "Field delimiter" and "String delimiter" empty.
7.	Change file extension of resulting file to .html
8.	Open html file in browser (I’m using Waterfox)
9.	Configure DownThemAll to only download one file at a time
10.	Select all URLs on page and use DownThemAll to try to download them


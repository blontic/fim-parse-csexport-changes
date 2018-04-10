# fim-parse-csexport-changes

Takes an XML file created by CSEXPORT, and produces a CSV file more suitable for opening in Excel.
Supports both single-valued attributes and multi-valued attributes

- Original Script By Carol Wapsphere (http://www.wapshere.com/missmiis/using-powershell-to-parse-a-csexport-file)..
- Script Rewrite By Jorge de Almeida Pinto (https://jorgequestforknowledge.wordpress.com/2013/02/08/parsing-a-csexport-generated-xml-file-into-a-scoped-csv-file/)  
- Script Customised To Only Output Changes [ADD/DELETE] And Dynamically Create Column Headers By Clayton Brady (https://github.com/blontic/fim-parse-csexport-changes)  

.EXAMPLE

.\FIMSync-Engine-Parse-ChangesOnly.ps1 -sourceXML C:\MAExport.xml -targetCSV C:\MAExport.csv

![Alt text](image.png?raw=true "Sample Export")
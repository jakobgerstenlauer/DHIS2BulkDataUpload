### Summary
With this app users can upload a list of individual events using a MS Excel spreadsheet file.
	
### How to use the bulk data upload app
First, select the program and your organizational unit. 
If the program applies to NGOs you additionally have to select your NGO.
Next, download the spreadsheet template and fill in your data in the first sheet.
Have a look at the second sheet which explains template structure and expected values.
Be sure to supply only values in the correct format and with appropriate value types.
Then, select the updated spreadsheet and upload it to the DHIS system.
You do not have to upload the spreadsheet with the data in the same session.
However, if you upload it later on, be sure to choose the correct program and org unit.
Check if there are any error messages.
Fix the spreadsheet before giving it another try.		
Note that the upload may take some time.
	
### Organization
index.html:			Main html page
manifest.webapp:	Configuration file with the
img:				Folder with images.
scripts:			Folder with javascript scripts.
-scripts.js:		File with the main functions defined for this app.
-Blob.js:			A blob implementation by Eli Grey and Devin Samarin:
					http://purl.eligrey.com/github/Blob.js/blob/master/Blob.js
-FileSaver.js:		A file saver implementation by Eli Grey and Devin Samarin:
					http://purl.eligrey.com/github/FileSaver.js/blob/master/FileSaver.js
-ouwt.js:			Common file with the functions needed to make the orgUnit tree work. Generic functions: getPolygon(), listenerFunction().
-md5-min.js:		File with the md5 hash-function, implemented by Paul Johnston:
					http://pajhome.org.uk/crypt/md5/instructions.html
-funcxl.js:			File with functions that process spreadsheets. Generic code in to_formulae(workbook).
-xlsx.core.min:		Parser and writer functions of various spreadsheet formats:
					https://github.com/SheetJS/js-xlsx/blob/master/README.md#parsing-workbooksJavascript
styles:				Folder with stylesheets
-styles.css:		File with the styles defined for this app
-tableStyles:		Common file with the styles of the main table
-treeStyles:		Common file with the styles used on the orgUnit tree
	
	
### Updated DHIS Version:
If the DHIS Version is updated, the following maintenance steps have to be carried out:
1) Update of ouwt.js: Get the latest version of this script
and copy getPolygon() in line 81 and the listenerFunction() in line 101 
from the old to the new version.
2) Update of scripts.js: Update the path to the event API in line 557 (marked with //TODO).
			
### Installation: 
This app is installed throught the DHIS2 menu normally.

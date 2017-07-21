### Summary
With this app users can upload a list of individual events using a MS Excel spreadsheet file.
	
### Credits:
This App started as a fork from the DHIS2 "Excel Data Importer" App version 2.2 developed by the Hisp Vietnam group. 
This App allows upload of event data to DHIS2 given a json template describing the data structure. The current App extends this functionality 
by offering dynamic template creation for programs and data sets. I am very thankful to Eric Mourin Marin who introduced me to DHIS2 and often helped me in
the practical implementation of the app. Marc Garnica Caparros shared his experience using the Excel Data Upload App and some
customized Javascript code with me.

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
- index.html:			Main html page
- manifest.webapp:	Configuration file.
- img:				Folder with images.
- scripts:			Folder with javascript scripts.
- scripts.js:		File with the main functions defined for this app.
- Blob.js:			A blob implementation by Eli Grey and Devin Samarin:
					http://purl.eligrey.com/github/Blob.js/blob/master/Blob.js
- FileSaver.js:		A file saver implementation by Eli Grey and Devin Samarin:
					http://purl.eligrey.com/github/FileSaver.js/blob/master/FileSaver.js
- ouwt.js:			Common file with the functions needed to make the orgUnit tree work. Generic functions: getPolygon(), listenerFunction().
- md5-min.js:		File with the md5 hash-function, implemented by Paul Johnston:
					http://pajhome.org.uk/crypt/md5/instructions.html
- funcxl.js:			File with functions that process spreadsheets. Generic code in to_formulae(workbook).
- xlsx.core.min:		Parser and writer functions of various spreadsheet formats:
					https://github.com/SheetJS/js-xlsx/blob/master/README.md#parsing-workbooksJavascript
- styles:				Folder with stylesheets
- styles.css:		File with the styles defined for this app
- tableStyles:		Common file with the styles of the main table
- treeStyles:		Common file with the styles used on the orgUnit tree
	
###Overview of API Calls
|API|API Call|JavaScript function|Purpose|
|-----|-----------------------------------|-----------|-----------------------------------|
|me|me.json?paging=FALSE&fields=userCredentials,displayName|queryUserRoles()|Retrieve the ID of the user to then query the userRoles API.|
|userRoles|"userRoles/""+roleId+"".json?paging=FALSE&fields=programs,dataSets"|queryUserRoles()|Retrieve the datasets and programs the user is authorized for.|
|dataSets|/dataSets.json?paging=false&field=dataSets|queryDataSetsApi()|We want to get a list of all datasets (ID and display name) the user is authorized to edit. Based on this information a drop down list is populated.|
|dataSets|"/dataSets/""+dataSet_id+"".json?paging=false&fields=dataSetElements,sections,periodType,categoryCombo"|queryDataSet()|Retrieve the data elements, sections, and categorycombos of this data set.|
|categoryCombos|"/categoryCombos/""+categoryComboId+"".json?paging=false&fields=categoryOptionCombos"|queryCategoryCombo()|Queries category option combinations for individual data elements or for the data set as a whole.|
|categoryOptionCombos|"categoryOptionCombos/""+categoryOptionComboId+"".json?paging=false&fields=displayName"|queryCategoryOptionCombo()|Queries ID and display name for a given category option combination.|
|programs|/programs.json?fields=id,displayName,programStages,categoryCombo&filter=attributeValues.attribute.name:eq:WISCC|queryProgramsApi()|Query program names, IDs, and category combinations. We also want to get a list of “program stages” which we can use to query the program stages api in the next step.|
|categoryCombos|"/categoryCombos/""+idCategoryCombo+"".json?fields=categories"|queryCategoryCombosApi()|Nested calls to three APIs in order to retrieve the name and id of NGOs associted to a program for non-official organisations.|
|categories|"/categories/""+val.id+"".json?fields=categoryOptions"|queryCategoryCombosApi()|Nested calls to three APIs in order to retrieve the name and id of NGOs associted to a program for non-official organisations.|
|categoryOptions|"/categoryOptions/""+idCategoryOptionCombo+"".json?fields=id,shortName"|queryCategoryOptionCombos()|Nested calls to three APIs in order to retrieve the name and id of NGOs associted to a program for non-official organisations.|
|programStages|"/programStages/""+ program_stage_id +"".json?&paging=false&fields=programStageDataElements,programStageSections"|retrieveProgramStageDataElements()|Retrieves data elements of program stage endpoint.|
|programStageSections|"/programStageSections/""+ sectionId +"".json?&paging=false&fields=programStageDataElements,displayName"|queryProgramStageSectionsInnerCall()|Retrieve the data elements associated to each program stage section.|
|sections|"/sections/""+ sectionId +"".json?&paging=false&fields=dataElements,displayName"|queryDataSetSections()|Retrieves the data elements (ID and display name) associated to a data set section.|
|programStageDataElements|"/programStageDataElements/""+ dataElement +"".json?&paging=false&fields=dataElement,compulsory"|queryProgramStageDataElementsInnerCall()|Retrieves Id, label, and compulsory property associated to a data element from the program stage data element endpoint.|
|dataElements|"/dataElements/""+ dataElementId +"".json?&paging=false&fields=formName,valueType,description,optionSetValue,optionSet"|queryDataElement()|Reads label, value type, description, and hasOptionSet property of a given data element from the dataElements API.|
|optionSets|"/optionSets/""+ optionSetId +"".json?&paging=false&fields=options"|queryOptionsInnerCall()|Read option IDs for given option set.|
|options|"/options/""+ optionId +"".json?&paging=false&fields=displayName"|queryOption()|Retrieves the text value for a given option ID.|
|dataValueSets|"/dataValueSets?dryRun=true&importStrategy=""+importStrategy"|importDataFromDataSet()|Upload data for data sets dry run|
|dataValueSets|"/dataValueSets?dryRun=false&importStrategy=""+importStrategy"|importDataFromDataSet()|Upload data for data sets |
|events|/events?dryRun=true|importData()|Upload data for programs|
|events|/events?dryRun=false|importData()|Upload data for programs dry run|
	
### Updated DHIS Version:
If the DHIS Version is updated, the following maintenance steps have to be carried out:
1) Update of ouwt.js: Get the latest version of this script
and copy getPolygon() in line 81 and the listenerFunction() in line 101 
from the old to the new version.
2) Update of scripts.js: Update the path to the event API in line 557 (marked with //TODO).
			
### Installation: 
This app is installed throught the DHIS2 menu normally.

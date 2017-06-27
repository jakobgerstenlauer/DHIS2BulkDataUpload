//Update zip file: ~/workspace/Bulk Data Upload$ 7z a -tzip BulkDataUpload.zip index.html scripts/ styles/
var baseUrl;
var apiBaseUrl;

//Do we import data rows if there are missing values?
var noMissingValuesAllowed = false;

function genCharArray(charA, charZ) {
    var a = [], i = charA.charCodeAt(0), j = charZ.charCodeAt(0);
    for (; i <= j; ++i) {
        a.push(String.fromCharCode(i));
    }
    return a;
}

var letters = genCharArray('A', 'Z'); // ["A", ..., "Z"]
var orgUnit_id_metadata;
var program_id_metadata;
var preImportValidationSummary = [];
var importSummary = [];
var isReload=0;
var isNGO = false; 
var eventHtmlString = "";
var sheetEndColumns = []; //end columns of all sheets
var sheetEndRows = []; // collected in xlsx.js
var orgUnitIdScheme = "UID";
var dataElementIdScheme = "UID";
var idScheme = "UID";
var dataValues = [];
var eventDataValues = [];
var errorString = "";
var hasErrors = false;

//Map used to go from PROGRAMID to Category Combo
var programsIDtoCategoryCombo = new Map();

//define logging levels
var traceLevelsIDtoNAME = new Map();
traceLevelsIDtoNAME.set(1,"trace");
traceLevelsIDtoNAME.set(2,"debug");
traceLevelsIDtoNAME.set(3,"info");
traceLevelsIDtoNAME.set(4,"error");

//obligatory data elements: key: column, value: label
var obligatoryDataElementsRowLabelMap = new Map();
obligatoryDataElementsRowLabelMap.set(1,"ReportingDate");
obligatoryDataElementsRowLabelMap.set(2,"Latitude");
obligatoryDataElementsRowLabelMap.set(3,"Longitude");

//Stores optionSet IDs for which a query of the options has already been sent in queryDataElement().
var queriedOptionSets = new Set();

// key: option ID,
// value: displayName (e.g. string "Yes")
var options = new Map();

//global counter for the options that still have to be queried
var optionsToQuery=0;

// key: optionSet ID
// value: array of option IDs
var optionMap = new Map();

//Map used to go from data set ID to the name of the data set
var dataSetsIDtoNAME = new Map();

//Map used to go from PROGRAMID to PROGRAM NAME
var programsIDtoNAME = new Map();

//Map used to go from PROGRAMID to PROGRAM STAGE
var programsIDtoSTAGE = new Map();

//Map used to go from name of non-official org to id
var nonOffOrgNametoID = new Map();

//program stage data elements
var dataElements = new Array();

//Id of program stage data elements
var dataElementIDs = new Set();

//IDs of program stage sections
programStageSectionID = new Array();

//Map associating section ID wih display name:
//key: section ID 
//value: display name for section
var sectionDisplayNameMap = new Map();

//Association between sections and program stage data elements
//key: section ID
//value: array of program stage data element IDs
var sectionDataElementMap = new Map();

//This Map is necessary to translate from program stage data element ID to the data element ID:
//key: program stage data element ID
//value: data element ID
var programStageDataElementMap = new Map();

/*
 * In javascript classes are possible but not supported by some browsers.
 * Therefore I use several Maps as workaround to describe the following properties of data elements:
 * All of the following Maps have the IDs of the dataElementIDs set as keys.
 * 
 * 1) Is the data element compulsory? Map dataElementsCompulsory
 * 2) What is the label of the data elment? Map dataElementsLabel
 * 3) What is the value type? Map dataElementsValueType
 * 4) A more detailed description of the data element: Map dataElementsDescription
 * 5) Is there an option set associated with the data element? Map dataElementsHasOptionSet
 * 6) If so, what is the option set id for the data element? Map dataElementsOptionSet
 */

//Is the data element compulsory? (boolean)
var dataElementsCompulsory = new Map();

//labels of all data elements
var dataElementsLabel = new Map();
	
//value type of all data elements
var dataElementsValueType = new Map();

//decriptions of all data elements
var dataElementsDescription = new Map();

//option set id for all data elements
var dataElementsOptionSet = new Map();

//boolean Map 
var dataElementsHasOptionSet = new Map();

//The set of programs the user can access.
let userPrograms = new Set();

//The set of datasets the user can access.
let userDataSets = new Set();

//The set of non-official organisations.
let nonOfficialOrgs = new Set();

//The number of organisational units.
var numOfOrgUnits = 0; 

//A regular expression used to find programmes for non-official organisations (NGOs).
var re = new RegExp("UCR");

var isTrue = new RegExp("true");

//The id of the program.
var program_id;

//The name of the program selected by the user.
var program_name;

//The program stage Id of the program selected by the user.
var program_stage_id;

//The administrative organisational unit, third level, of the program selected by the user.
var org_unit;

//The non-official organisational unit of the user (a NGO).
var non_off_org_unit;

//The Id of the non-official organisational unit of the user (a NGO).
var non_off_org_unit_id;

//Has the drop down list of programes already been created? 0: false 1:true
var programListCreated=0;

/**
 * Hide the button with the given ID.
 * @param id
 * @returns
 */
function hideSelectButton (id) {
    document.getElementById(id).hidden = true;
}

/**
 * Show button with the selected ID.
 * @param id
 * @returns
 */
function showSelectButton (id) {
    document.getElementById(id).hidden = false;
}

/**
 * Retrieve a specific element of the html document and set its value to the given value.
 * @param id
 * @param val
 * @returns No return value.
 */
function setSelectValue (id, val) {
    document.getElementById(id).value = val;
}

/**
 * Retrieve a specific element of the html document and set its name to the given value.
 * @param id
 * @param val
 * @returns No return value.
 */
function setSelectName (id, val) {
    document.getElementById(id).name = val;
}

/**
 * Returns the value of a specific element of the html document.
 * @param id
 * @returns
 */
function getSelectValue (id) {
    return document.getElementById(id).value;
}

/**
 * Clears the text area.
 * @returns No return value.
 */
function eraseText() {
	document.getElementById("myTextArea").value = "";
	logging="";
}

/**
 * Stores an object in the session storage.
 * @param name
 * @param object
 * @returns No return value.
 */
function store(name, object){
	
	var isStringName = "isString"+name;
	
	if(typeof object === "string"){
		sessionStorage.setItem(name, object);
		sessionStorage.setItem(isStringName, true);		
	}else{
		sessionStorage.setItem(name, JSON.stringify(object));
		sessionStorage.setItem(isStringName, false);		
	}
}

/**
 * Restores the original object from session storage.
 * @param name
 * @returns
 */
function reStore(name){
	var isStringName = "isString"+name;
	var object = sessionStorage.getItem(name);
	//if the original object was of type String
	if(sessionStorage.getItem(isStringName)){		
		return object;
	}else{		
	    return JSON.parse(object);
	}
}

/**
 * This function is run before any reload of the page.
 */
onbeforeunload = function()
{
	store("reloaded", true);
	
	// Store log messages
	store("logging", logging);
	
	// Store metadata
	store("org_unit_id", org_unit_id);
	store("org_unit_name", org_unit_name);
	store("org_unit_polygon", org_unit_polygon);
	store("program_id", program_id);
	store("program_name", program_name);
	store("resultArray", resultArray);
	store("metaDataArray", metaDataArray);	
	store("traceLevel", getSelectValue ("traceLevelDropDown"));
	store("isNGO", isNGO);
	
	if(isNGO){
		store("non_off_org_unit", non_off_org_unit);
		store("non_off_org_unit_id", non_off_org_unit_id);
	}
};

/**
 * This function is run after reload of the page.
 */
window.onload = function(event)
{
	isReloaded = sessionStorage.getItem("reloaded");
	
	if(isReloaded){
		
		// Restore log messages
		logging = reStore("logging");
		var TheTextBox = document.getElementById("myTextArea");
		TheTextBox.value = logging;
		
		// Restore metadata and read data
		org_unit_id =      	reStore("org_unit_id");
		org_unit_name =    	reStore("org_unit_name");
		org_unit_polygon = 	reStore("org_unit_polygon");
		program_id = 	   	reStore("program_id");
		program_name = 		reStore("program_name");
		resultArray = 		reStore("resultArray");
		metaDataArray = 	reStore("metaDataArray");	
		
		//reset buttons
		setSelectValue('traceLevelDropDown', reStore("traceLevel"));
		showSelectButton ('programList');
		setSelectValue('programList', reStore("program_id"));
		setSelectName('programList', reStore("program_name"));
		
		//If the selected program is for NGOs reset the NGO button too:
		if(reStore("isNGO")){
			hideSelectButton ('orgList');
			setSelectValue('orgList', reStore("non_off_org_unit_id"));
			setSelectName('orgList', reStore("non_off_org_unit"));
		}
	}
};

/**
 * Reads properties from properties file "manifest.webapp".
 * 
 * Here it is assumed that the path of the properties file and the script file are identical!
 * This function calls queryUserRoles().
 ****/
function readProperties() {
	localStorage.setItem("reloaded", false);
	$.getJSON("manifest.webapp", function( json ) {
		baseUrl = json.activities.dhis.href;
		//TODO Update to new version when current DHIS version = 30!		
		apiBaseUrl = baseUrl + "/api/26";					
	})
	.done(function(){
		queryUserRoles();
	});		
}

/**
* Stores all programs the user has access rights to in the set userPrograms.
* Stores all data sets the user has access rights to in the set userDataSets.
* Calls queryProgramsApi() and queryDataSetsApi() before terminating.
*
***/
function queryUserRoles() {
	//$("#rightBar").show();
	//http://who-dev.essi.upc.edu:8082/api/me.json?paging=FALSE&fields=userCredentials
    $.getJSON(apiBaseUrl+"/me.json?paging=FALSE&fields=userCredentials", 
	function (json) {      
		$.each( json.userCredentials.userRoles, function( key, val ) { 
			var roleId = val.id
			//http://who-dev.essi.upc.edu:8082/api/userRoles/rtxnH4ZGLQh.json?paging=FALSE&fields=programs
			$.getJSON(apiBaseUrl+"/userRoles/"+roleId+".json?paging=FALSE&fields=programs,dataSets", 
			function (json) {      		
				$.each( json.programs, function( key, val ) {
					userPrograms.add(val.id);										
				})
				$.each( json.dataSets, function( key, val ) {
					userDataSets.add(val.id);										
				})
			})
		})
	}).done(function(){
		queryProgramsApi();
		queryDataSetsApi();
	})		
}

/**
 *  Queries dataSets API.
 *  Here it is assumed that the path of the properties file and the script file are identical! 
 *  In the first step, we query the dataSets endpoint: /api/dataSets.json
 *  We want to get a list of data elements contained in the data set.
 */
function queryDataSetsApi() {
	
	var dataSetCounter = 1000;
	var authorizedDataSets = 0;
	
	//$("#rightBar").show();
    $.getJSON(apiBaseUrl+"/dataSets.json?paging=false&field=dataSets", 
	function (json) {
    	dataSetCounter = json.dataSets.length;
    	$.each( json.dataSets, function( key, val ) {			
			//check if user has access rights for this data set			
			if(userDataSets.has(val.id)){
				dataSetsIDtoNAME.set(val.id, val.displayName);
				console.log("data set id: "+val.id+" display name: "+val.displayName);
				authorizedDataSets++;
			}		
			dataSetCounter--;
		})
	}).done(function(){	
		if(authorizedDataSets===0){
			add("You are not authorized to edit any data sets!",1);
			return;
		}
		if((dataSetCounter===0)&&(authorizedDataSets > 0)){
			tryToCreateDataSetDropDown();
		}else{
			sleep(1000);
			if(programCounter===0){
				tryToCreateDataSetDropDown();
			}else{
				sleep(2000);
				tryToCreateDataSetDropDown();
			}
		}
	});		
}

function tryToCreateDataSetDropDown(){
	var sel = document.getElementById('dataSetList');
	if(sel && (dataSetsIDtoNAME.length > 0) ){
		createDataSetDropDown();
	}else{
		sleep (1000);
		tryToCreateDataSetDropDown();
	}
}

/**
 * Creates a drop-down list with all the programs the user has access to.
 * 
 * This function sleeps until the html element program list is loaded.
 * Once it is loaded a drop down menu is created based on the available
 * program options which depend on the user.
 */
function createDataSetDropDown() {	
		if(document.getElementById('dataSetList')){
				var sel = document.getElementById('dataSetList');
				for (const [id,name] of dataSetsIDtoNAME.entries()) {
						var opt = document.createElement('option');		
						//console.log(name);	
						opt.innerHTML = name;
						//console.log(id);	
						opt.value = id;
						sel.appendChild(opt);
				}		
				dataSetListCreated=1;			
		}
}
	
/**
 *  Queries programs API.
 *  Here it is assumed that the path of the properties file and the script file are identical! 
 *  In the first step, we query the program endpoint: /api/programs.json
 *  We want to get a list of “program stages” which we can use to query the program stages api in the next step.
 *  //http://who-dev.essi.upc.edu:8082/api/programs.jsn?&paging=false&fields=id,displayName,attributeValues[value,attribute    [id,name]]&filter=attributeValues.attribute.name:eq:WISCC
 */
function queryProgramsApi() {
	
	//count down of how many programs have been processed
	var programCounter = 100000;
	
	
	//$("#rightBar").show();
    $.getJSON(apiBaseUrl+"/programs.json?fields=id,displayName,programStages,categoryCombo&filter=attributeValues.attribute.name:eq:WISCC", 
	function (json) {
    	programCounter = json.programs.length;
		$.each( json.programs, function( key, val ) {			
			//check if user has access rights for this program			
			if(userPrograms.has(val.id)){	
				if(typeof val.programStages[0] === 'undefined'){	
					console.log("Undefined program stage for program: "+val.displayName+" id: "+val.id);					
				}else{
					programsIDtoCategoryCombo.set(val.id, val.categoryCombo.id);
					//PROGRAM ID TO NAME (JAVAscript HASHTABLES)
					programsIDtoNAME.set(val.id, val.displayName);
					//Check if user has access rights for this program				
					//PROGRAM ID TO PROGRAM STAGE (JAVAscript HASHTABLES)
					programsIDtoSTAGE.set(val.id, val.programStages[0].id);
				}
			}
			programCounter++;
		})
	}).done(function(){		
		if(programCounter===0){
			tryToCreateDropDown();
		}else{
			sleep(1000);
			if(programCounter===0){
				tryToCreateDropDown();
			}else{
				sleep(2000);
				tryToCreateDropDown();
			}
		}
	});		
}
/**
 * Tries to create the drop down list for programs and triggers a reload if the html element 'programList' does not yet exist.
 * @returns
 */
function tryToCreateDropDown(){
	var sel = document.getElementById('programList');
	if(sel){
		createDropDown();
	}else{//trigger a reload
		location.reload();
	}
}

/**
 * Sleep for a certain time.
 * 
 * @param Time in milliseconds
 * TODO: This function is not supported by Internet Explorer 11.576 
 */
function sleep (time) {
  return new Promise((resolve) => setTimeout(resolve, time));
}

/**
 * Creates a drop-down list with all the programs the user has access to.
 * 
 * This function sleeps until the html element program list is loaded.
 * Once it is loaded a drop down menu is created based on the available
 * program options which depend on the user.
 */
function createDropDown() {	
		if(document.getElementById('programList')){
				var sel = document.getElementById('programList');
				for (const [id,name] of programsIDtoNAME.entries()) {
						var opt = document.createElement('option');		
						//console.log(name);	
						opt.innerHTML = name;
						//console.log(id);	
						opt.value = id;
						sel.appendChild(opt);
				}		
				programListCreated=1;			
		}
}

/**
 *  Query the data elements of the selected program.
 *  
 *  This function is triggered once the user selects a program.
 *  It collects relevant information about the selected program and calls 
 *  queryCategoryCombosApi() which retrieves all relevant data elements of this program. 
 */
function queryProgramStageApi() {
	
	//$("#rightBar").show();
	programSelected = true;
	
	//Delete all elements of the organisation drop down list 
	//in order to avoid duplicates!
	var sel = document.getElementById('orgList');
	if(typeof sel === 'undefined'){	
		console.log("drop down for non-official orgs yet undefined!")
	}else{
		var length = sel.options.length;		
		for (i = 0; i < length; i++) {
		  sel.options[i] = null;
		}
	}
	
	//Make the get spreadsheet button invisible.
	//This is necessary if the user had previously selected a program. 
	$("#getSpreadsheet").prop("hidden",true);
	$("#uploadSpreadsheet").prop("disabled",true);
	$("#orgList").prop("hidden",true);
	
	//get the id of the selected program
	program_id=$("#programList").val();
	//console.log(program_id);

	//get the name of the selected program
	program_name=programsIDtoNAME.get(program_id);
	//console.log(program_name);
	
	//get the corresponding program stage id
	program_stage_id=programsIDtoSTAGE.get(program_id);
	
	//console.log(program_stage_id);
	retrieveProgramStageDataElements(program_stage_id);
	
	//The button for spreadsheet download is only shown once the user has selected his 
	//non-official organisation (if applicable) and all necessary information has been retrieved from the web api.
	if(re.test(program_name)){		
		var categoryCombo = programsIDtoCategoryCombo.get(program_id);
		//console.log(categoryCombo); //rkfrcCeCb15
		queryCategoryCombosApi(categoryCombo);
		$("#orgList").prop("hidden",false);		
		isNGO=true;
	}else{
		orgUnitSelected = false;
		//$("#rightBar").hide();
		$("#uploadSpreadsheet").prop("disabled",false);
		$("#getSpreadsheet").prop("hidden",false);
		document.getElementById("getSpreadsheet").onclick = function fun() {
	        console.log("Activated getSpreadsheet button!");
	        getSpreadsheet();  
	    }			
	}		
}

/**
 * Nested calls to three different APIs in order to retrieve the name and id of NGOs.
 * 
 * Queries first the /categoryCombos API and second the /categories API.
 * Calls queryCategoryOptionCombos() with the id of the retrieved category.
 * 
 * @param idCategoryCombo Id of the category combo with which the /categoryCombos API will be queried.
 * @returns
 */
function queryCategoryCombosApi(idCategoryCombo) {
	//make the select button invisible
	$("#orgList").prop("hidden",false);	
	//$("#rightBar").show();
	
	//Delete all elements of the organisation drop down list 
	//in order to avoid duplicates!
	var sel = document.getElementById('orgList');
	var length = sel.options.length;
	for (i = 0; i < length; i++) {
	  sel.options[i] = null;
	}
	
    $.getJSON(apiBaseUrl+"/categoryCombos/"+idCategoryCombo+".json?fields=categories", 
	function (json) {
    	$.each( json.categories, function( key, val ) {
    	    	$.getJSON(apiBaseUrl+"/categories/"+val.id+".json?fields=categoryOptions",
						function (json) {
    	    				var counterNGOs=0;
					    	var totalNumNGOs=json.categoryOptions.length;
					    	//console.log(json.categoryOptions);
					    	//console.log(json.categoryOptions.length);
    	    				$.each( json.categoryOptions, function( key, val ) {			    		
					    		queryCategoryOptionCombos(val.id).
					    		then(orgUnit => { 				
									var sel = document.getElementById('orgList');					
									var opt = document.createElement('option');
									opt.value = orgUnit;
									opt.innerHTML = orgUnit;
									sel.appendChild(opt);
									counterNGOs++;
									//once all NGOs have been retrieved.
									if(counterNGOs===totalNumNGOs){
										document.getElementById("getSpreadsheet").onclick = function fun() {
									        console.log("Activated getSpreadsheet button!");
									        getSpreadsheet();  
									    }
										//stop showing the loading image 
										//$("#rightBar").hide();
										//show the get spreadsheet button
										$("#getSpreadsheet").prop("hidden",false);	
										//enable the input field for file upload
										$("#uploadSpreadsheet").prop("disabled",false);																				
									}
								})			
							.catch(error => { 
								console.log("No value available for category option combo: " + item); 
							});			
					    	})
						})    	
    	})
	})
}

/**
 * Queries /categoryOptions API in order to retrieve the name and id of NGOs.
 * 
 * @param idCategoryOptionCombo
 * @returns
 */
function queryCategoryOptionCombos(idCategoryOptionCombo) {
	return new Promise(
			function (resolve, reject) {
				   	$.getJSON(apiBaseUrl+"/categoryOptions/"+idCategoryOptionCombo+".json?fields=id,shortName", 
					function (json) { 
				   				nonOffOrgNametoID.set(json.shortName,json.id);
								nonOfficialOrgs.add(json.shortName,1);
								//console.log(json.shortName,1);
								resolve(json.shortName,1);		
					})
			}
	)
}

/**
 * Retrieves data elements of program stage endpoint. 
 * 
 * Queries the program stage endpoint.
 * http://who-dev.essi.upc.edu:8082/api/programStages/JP8t81g0uIT
 * This is an example how to query for a given program stage.
 * http://who-dev.essi.upc.edu:8082/api/programStages/T2FtodAmxMa?fields=programStageDataElements
 * @return No return value.
 */
function retrieveProgramStageDataElements(program_stage_id){
	return new Promise(
			function (resolve, reject) {
				//get the id of all program stage data elements
				$.getJSON(apiBaseUrl+"/programStages/"+ program_stage_id +".json?&paging=false&"+
						"fields=programStageDataElements,programStageSections", function (json) 
						{
							$.each( json.programStageDataElements, function( key, val ) {
								dataElements[key] = val.id;
							});
							$.each( json.programStageSections, function( key, val ) {
								programStageSectionID[key] = val.id;
							});	
						}).done(function() {	
							queryProgramStageDataElements();
							queryProgramStageSections();
						}).done(function() {	
							resolve(1);
						});

			});
}

/**
 * Query the program stage section endpoint for all program stage sections (elements of the array programStageSectionID).
 * Retrieve the data elements associated to each section.
 * @return No return value.
 */
function queryProgramStageSections() {
	for (var i = 0; i< programStageSectionID.length; i++){
		console.log("Query program stage section with ID: "+programStageSectionID[i])
		queryProgramStageSectionsInnerCall(programStageSectionID[i], i);
	};
}

/**
 * Reads Id and label of data element of program stage data element with index i 
 * from /programStageDataElements API.
 * 
 * @param dataElement An array with data element Ids.
 * @param i Index of array dataElement. 
 * @returns
 */
function queryProgramStageSectionsInnerCall(sectionId, i){
$.getJSON(apiBaseUrl+"/programStageSections/"+ sectionId +".json?&paging=false&"+
		"fields=programStageDataElements,displayName", function (json) {	
			console.log(json);
			//program stage data element ids
	        var arrayOfDataElementIDs = []; 	        
	        $.each( json.programStageDataElements, function( key, val ) {
	        	arrayOfDataElementIDs.push(val.id);
			});	
	        
	        console.log("section id: " + sectionId + " display name: " + json.displayName);
	        sectionDisplayNameMap.set(sectionId, json.displayName);	        
	        
	        console.log("Add new array of data element IDs to sectionDataElementMap: "+arrayOfDataElementIDs.toString())
			sectionDataElementMap.set(sectionId, arrayOfDataElementIDs);
		});
}

/**
 * Query the program stage endpoint for all program stage data elements (elements of the array dataElements).
 * @return No return value.
 */
function queryProgramStageDataElements() {
	for (var i = 0; i< dataElements.length; i++){
		queryProgramStageDataElementsInnerCall(dataElements[i], i);
	};
	if(programListCreated===0){
		createDropDown();
	}
}

/**
 * Reads Id and label of data element of program stage data element with index i 
 * from /programStageDataElements API.
 * 
 * @param dataElement An array with program stage data element Ids.
 * @param i Index of array dataElement. 
 * @returns
 */
function queryProgramStageDataElementsInnerCall(dataElement, i){
$.getJSON(apiBaseUrl+"/programStageDataElements/"+ dataElement +".json?&paging=false&"+
		"fields=dataElement,compulsory", function (json) {
	        programStageDataElementMap.set(dataElement,json.dataElement.id);
			dataElementIDs.add(json.dataElement.id);
			dataElementsCompulsory.set(dataElement,json.compulsory);
			queryDataElement(json.dataElement.id, i);
		});
}

/**
 * Reads label of data element i from /dataElements API.
 * 
 * @param dataElementId This is the data element ID.
 * @param i
 * @returns
 */
function queryDataElement(dataElementId, i) {	
	$.getJSON(apiBaseUrl+"/dataElements/"+ dataElementId +".json?&paging=false&"+
	"fields=formName,valueType,description,optionSetValue,optionSet", function (json) {			
		
		if(!isNullOrUndefinedOrEmptyString(json.formName)){
			dataElementsLabel.set(dataElementId,json.formName);
		}
		
		if(!isNullOrUndefinedOrEmptyString(json.valueType)){
			dataElementsValueType.set(dataElementId,json.valueType);
		}
		
		if(!isNullOrUndefinedOrEmptyString(json.description)){
			dataElementsDescription.set(dataElementId,json.description);	
		}
		
		if(!isNullOrUndefinedOrEmptyString(json.optionSetValue)){
			dataElementsHasOptionSet.set(dataElementId,json.optionSetValue);
		}
		
		//Test if the json object representing the data element has an option set:
		if(json.optionSetValue==true){
			if(json.hasOwnProperty("optionSet")){
				var optionSetId = json.optionSet.id;
				add("Line 539: dataElementId: " + dataElementId 
						+" optionSetId: " + optionSetId , 1)
				dataElementsOptionSet.set(dataElementId,optionSetId);
				//check if a query for the options has already been sent for this option set:
				if(!queriedOptionSets.has(optionSetId)){
					queryOptions(optionSetId);	
					queriedOptionSets.add(optionSetId)
				}
			}else{//Correct value of this Map if there is no optionSet property!
				dataElementsHasOptionSet.set(dataElementId,false);
			}			
		}
	});
}

/**
 * Read option IDs for given option set.
 * 
 * @param optionSetId
 */
function queryOptions(optionSetId) {	
	add("Line 562: optionSetId: " + optionSetId, 1)
	//Here, I have to wait for the response!
	queryOptionsInnerCall(optionSetId).then(arrayOfOptionIDs=> {
		optionMap.set(optionSetId, arrayOfOptionIDs);		
		add("Added "+arrayOfOptionIDs.length+" options to option set: "+optionSetId, 1);
	});
}

/**
 * Queries the optionSets Api for the options of a given optionset ID.
 * @param optionSetId
 * @returns
 */
function queryOptionsInnerCall(optionSetId){
	return new Promise(
			function (resolve, reject) {
					$.getJSON(apiBaseUrl+"/optionSets/"+ optionSetId +".json?&paging=false&"+
							"fields=options", function (json) {	
								add("Line 575: "+JSON.stringify(json),1);
								
								var local_arrayOfOptionIDs = new Array(json.options.length);		
								for (var i = 0; i < json.options.length; i++) {
								    var object = json.options[i];
								    local_arrayOfOptionIDs[i]=object.id;		    
								    //console.log(JSON.stringify(object));	
								    if(!options.has(object.id)){
								    	optionsToQuery++;
										queryOption(object.id);
									};
								}
								resolve(local_arrayOfOptionIDs);			 
							});
				}
			)
}

/**
 * Retrieves the text value for a given option ID.
 * 
 * @param optionId
 * @returns
 */
function queryOption(optionId) {	
	console.log(JSON.stringify(optionId));	
	$.getJSON(apiBaseUrl+"/options/"+ optionId +".json?&paging=false&"+
	"fields=displayName", function (json) {	
		//console.log(JSON.stringify(json));
		console.log("key: "+optionId+" value: "+json.displayName)
		options.set(optionId, json.displayName);
		optionsToQuery--;
	});
}

/**
 * Creates a template spreadsheet.
 *  
 * @returns
 */
function getSpreadsheet() {

	console.log("Start getSpreadsheet().")
	var numOfElements = dataElementIDs.size;
	var dataElementsSectionLabel_Array = new Array(numOfElements);
	var dataElementsLabel_Array = new Array(numOfElements);
	var dataElementsValueType_Array = new Array(numOfElements);
	var dataElementsIDs_Array = new Array(numOfElements);
	var dataElementsDescription_Array = new Array(numOfElements);
	var dataElementsCompulsory_Array = new Array(numOfElements);
		
	for (var [key, value] of sectionDisplayNameMap.entries()) {
		  console.log(key + ' = ' + value);
	}
	
	//Here we order the data elements according to the order of the sections of the program
	var i = 0;
	for (var [key, value] of sectionDataElementMap.entries()) {
		 var arrayOfDataElementIDs = value;		 
		 console.log("dataElements: "+ arrayOfDataElementIDs.toString())
		    		 
		 for(let programStageDataElement of arrayOfDataElementIDs){		
			dataElementsSectionLabel_Array[i]= sectionDisplayNameMap.get(key);
			//translate from the program stage data element ID to the data element ID
		 	var dataElement = programStageDataElementMap.get(programStageDataElement);		 	
		 	console.log("i: "+i+" program stage data element: "+programStageDataElement+" dataElement: "+dataElement+" label: "+ dataElementsLabel.get(dataElement)+" description:"+dataElementsDescription.get(dataElement))
		    dataElementsLabel_Array[i]=dataElementsLabel.get(dataElement);
			dataElementsValueType_Array[i]=dataElementsValueType.get(dataElement);
			dataElementsIDs_Array[i]=dataElement;
			dataElementsDescription_Array[i]=dataElementsDescription.get(dataElement);
			dataElementsCompulsory_Array[i]=dataElementsCompulsory.get(dataElement);
			i++;
		}	
	}
	
	//if a non-official unit has been selected	
	if(orgUnitSelected){
		non_off_org_unit=$("#orgList").val();
		non_off_org_unit_id=nonOffOrgNametoID.get(non_off_org_unit)
	}else{
		non_off_org_unit="not applicable";
		non_off_org_unit_id="not applicable";
	}
	
	if(programSelected && regionalUnitSelected){
		
	//first row with header containing informative labels for all data elements	  
	var output_array_sheet_0 = [
		["This template spreadsheet was created by the Bulk Data Upload App for DHIS2."],
		[""],
		["How to use the app:"],
		[""],
		["First, select the program and your organizational unit."],
		["If the program applies to NGOs you additionally have to select your NGO."],
		["Next, download the spreadsheet template and fill in your data in the first sheet."],
		["Have a look at the third sheet (\"Legend\") which explains template structure and expected values."],
		["Be sure to supply only values in the correct format and with appropriate value types."],
		["Then, select the updated spreadsheet and upload it to the DHIS system."],
		["You do not have to upload the spreadsheet with the data in the same session."],
		["However, if you upload it later on, be sure to choose the correct program and org unit."],
		["The app will reject the spreadsheet if program id and org unit do not coincide with the metadata in the \"Metadata\" sheet."],
		["Check if there are any error messages."],
		["If there are errors, you may change the log level to \"debug\" using the button blow the text window to gain more information about possible inconsistencies in the data."],
		["Fix the spreadsheet before giving it another try."],		
		["Note that the data upload may take some time and your browser may warn you that the app is unresponsive."],
		["Please ignore this browser warning and keep waiting for the app to respond."]
	];
		
	//first row with header containing informative labels for all data elements	  
	var output_array_sheet_1 = [
		//dataElementsSectionLabel_Array
		[].concat.apply([],["ReportingDate","Latitude","Longitude",dataElementsLabel_Array]),
		[,,,,,,,,]
	];
	
	//console.log(output_array_sheet_1);
	
	var output_array_sheet_2 = [
		["The first row of spreadsheet 1 contains descriptive labels of all columns."],
		[""],
		["Fixed column:"],
		["Reporting Date","","","Enter the date time when the data was recorded in the following format: <2016-12-01T00:00:00.000> (first December 2016)."],
		[""],
		["Generic columns:"],
		[""],
		["data element ID:","Section:","Label:","Description:","","Compulsory?","Value type:","Option set Id:","Possible values:"]
	];

	for(j = 0; j<dataElementsLabel_Array.length; j++){
		var dataElement = dataElementsIDs_Array[j];
		if(dataElementsHasOptionSet.get(dataElement)){
			var option_set_id = dataElementsOptionSet.get(dataElement);
			console.log(option_set_id);
			//array of IDs of optional values
			var ids_of_options = optionMap.get(option_set_id);
			console.log(ids_of_options);
			var numOptions = ids_of_options.length;
			console.log(numOptions);
			var new_row = new Array(8+numOptions);
			new_row[0]=dataElement;
			new_row[1]=dataElementsSectionLabel_Array[j];
			new_row[2]=dataElementsLabel_Array[j];
			new_row[3]=dataElementsDescription_Array[j];
			new_row[4]="";
			new_row[5]=dataElementsCompulsory_Array[j];
			new_row[6]=dataElementsValueType_Array[j];		
			new_row[7]=option_set_id;
			for(k = 0; k < numOptions; k++){
				new_row[8+k]=options.get(ids_of_options[k]);
			}
			output_array_sheet_2.push(new_row);
		}else{
			var new_row = new Array(7);
			new_row[0]=dataElement;
			new_row[1]=dataElementsSectionLabel_Array[j];
			new_row[2]=dataElementsLabel_Array[j];
			new_row[3]=dataElementsDescription_Array[j];
			new_row[4]="";
			new_row[5]=dataElementsCompulsory_Array[j];
			new_row[6]=dataElementsValueType_Array[j];	
			output_array_sheet_2.push(new_row);			
		}
	}	
	
	var output_array_sheet_3 = [
		
		// creating the header of the table	  
		// create first table row
		// ProgramId,GfOWfC9blOI,ProgramStage,JP8t81g0uIT,,,,
		[].concat.apply([],["ProgramId", "ProgramStage", "ProgramDescription","OrganisationalUnit","OrgUnitId","UnofficialOrganisationalUnit", "IdUnofficialOrgUnit", dataElementsLabel_Array]),
	  
		// create second table row
		//Description,Health ministry officers manage collective dwelling inspections,,,,,,
		[].concat.apply([],[program_id, program_stage_id, program_name, org_unit_name, org_unit_id, non_off_org_unit, non_off_org_unit_id, dataElementsIDs_Array])

	];
	
	var str = program_name;
	str=str.replace("  ", "_");
	str=str.replace(" ", "_");
	str=str.replace("(", "_");
	str=str.replace(")", "_");
	var fileName = "WISCC_Data_Upload_"+str+"_.xlsx"
	var ws0_name = "Readme";	
	var ws1_name = "WISCC_Data_Upload Data Entry Template";	
	var ws2_name = "Legend";	
	var ws3_name = "Metadata - Do Not Change!";	
	
	var workbook = new Workbook();
	ws0 = sheet_from_array_of_arrays(output_array_sheet_0);
	ws1 = sheet_from_array_of_arrays(output_array_sheet_1);
	ws2 = sheet_from_array_of_arrays(output_array_sheet_2);
	ws3 = sheet_from_array_of_arrays(output_array_sheet_3);
	
	/* add worksheet 0 to workbook */
	workbook.SheetNames.push(ws0_name);
	workbook.Sheets[ws0_name] = ws0;
	
	/* add worksheet 1 to workbook */
	workbook.SheetNames.push(ws1_name);
	workbook.Sheets[ws1_name] = ws1;
	
	/* add worksheet 2 to workbook */
	workbook.SheetNames.push(ws2_name);
	workbook.Sheets[ws2_name] = ws2;
	
	/* add worksheet 3 to workbook */
	workbook.SheetNames.push(ws3_name);
	workbook.Sheets[ws3_name] = ws3;
	
	var wb_out = XLSX.write(workbook, {bookType:'xlsx', bookSST:true, type: 'binary'});	
	saveAs(new Blob([s2ab(wb_out)],{type:"application/octet-stream"}), fileName)	
	}else{
		add("Error: Can not create spreadsheet. Either the org. unit or the program was not selected!",3);
	}
}

/**
 * Sends Json collection of events to the events API and processes the import summary reply by the server.
 * 
 * @param isTest Should the function just do a dry run?
 * @returns
 */
function importData(){

	return new Promise(
			function (resolve, reject) {

				$.ajax({
					method: "POST",
					type: 'post',
					url: apiBaseUrl + "/events?dryRun=true",
					contentType: "application/json; charset=utf-8",
					data: JSON.stringify(eventDataValues),
					dataType: 'json',
					headers:{
						'Accept': 'application/json',
						'Content-Type': 'application/json'
					},	
					async: false
				}).done(function(res) {						
					add(res.message,3);
					add(res.httpStatus,3);

					var importSummaryArray = res.response.importSummaries;
					var successfulImports = 0;
					for (var i = 0; i < importSummaryArray.length; i++){
						if(importSummaryArray[i].status === "SUCCESS"){
							successfulImports++;
						}
					}		
					if(successfulImports==importSummaryArray.length){
						add("All "+ importSummaryArray.length +" row imports were successful in the dry run!", 3)
						add("Now the real import of data starts!", 3)
						
						$.ajax({
						method: "POST",
						type: 'post',
						url: apiBaseUrl + "/events?dryRun=false",
						contentType: "application/json; charset=utf-8",
						data: JSON.stringify(eventDataValues),
						dataType: 'json',
						headers:{
							'Accept': 'application/json',
							'Content-Type': 'application/json'
						},
						async: false
					}).done(function(res) {						
						add(res.message,3);
						add(res.httpStatus,3);

						add("Total number of data elements imported: " + res.response.imported, 3);

						var ignoredValues = res.response.ignored;
						if(ignoredValues>0){
							add("Total number of data elements ignored: " + ignoredValues, 3);
							add("There are several errors that have to be fixed! " + ignoredValues, 3);	
							onbeforeunload();
							reject("There are "+ ignoredValues +" errors that have to be fixed! ");
						}

						//write import summary for each row up to max_length
						var max_length = 100;
						if(res.response.importSummaries.length < max_length){
							for(var i = 0; i < res.response.importSummaries.length;i++){
								add("row: "+i+" data elements imported: "+res.response.importSummaries[i].importCount.imported, 3);
							}
						}else{
							add("Only the import results for the first "+max_length+" of "+res.response.importSummaries.length+" are shown:", 3);
							for(var i = 0; i < max_length;i++){
								add("row: "+i+" values imported: "+res.response.importSummaries[i].importCount.imported, 3);
							}				
						}	

						onbeforeunload();
						resolve("Successful data upload!");
					})
					.fail(function (request, textStatus, errorThrown) {
						try
						{			
							add("The following request could not be processed:"+JSON.stringify(eventDataValues), 4)
							add("Event data import response:", 3);
							if(isNullOrUndefined(request)){
								if(isNullOrUndefined(errorThrown)){
									onbeforeunload();
									reject();
								}else{
									onbeforeunload();
									reject(errorThrown);
								}
							}else{
								console.log(request);
								if(isNullOrUndefined(textStatus)){
									console.log(textStatus);
								}
								if(isNullOrUndefined(errorThrown)){
									onbeforeunload();
									reject();
								}else{
									console.log(errorThrown);
									onbeforeunload();
									reject(errorThrown);
								}
							}
						}
						catch(ex)
						{
							add("Something went wrong while fetching event import error summary", 4);
							add(ex, 4);
							console.log(ex);
							reject("Something went wrong while fetching event import error summary");
						}			
					})
					}else{
						reject("Error: Only "+ successfulImports + " out of " + importSummaryArray.length +" imports were successful!")
					}

				})
				.fail(function (request, textStatus, errorThrown) {
					try
					{			
						add("The following request could not be processed:"+JSON.stringify(eventDataValues), 4)
						add("Event data import response:", 3);
						if(isNullOrUndefined(request)){
							if(isNullOrUndefined(errorThrown)){
								onbeforeunload();
								reject();
							}else{
								onbeforeunload();
								reject(errorThrown);
							}
						}else{
							console.log(request);
							if(isNullOrUndefined(textStatus)){
								console.log(textStatus);
							}
							if(isNullOrUndefined(errorThrown)){
								onbeforeunload();
								reject();
							}else{
								console.log(errorThrown);
								onbeforeunload();
								reject(errorThrown);
							}
						}
					}
					catch(ex)
					{
						add("Something went wrong while fetching event import error summary", 4);
						add(ex, 4);
						console.log(ex);
						onbeforeunload();
						reject("Something went wrong while fetching event import error summary");
					}			
				})

			})
}

/**
 * Is this label an obligatory label?
 * @param label
 * @returns
 */
function isObligatoryLabel(label){	
	return obligatoryDataElementsLabel.includes(label);
}

/**
 * Checks if the Json array of arrays containing the input data 
 * has duplicate rows using the MD5 hashing function.
 * @returns Boolean, are there any duplicate rows?
 */
function hasDuplicates(){	
	set_of_hashes = new Set();
	var hasDuplicates = false;
	resultArray.forEach( function (arrayItem)
	{		
			hash = hex_md5(JSON.stringify(arrayItem));		
			add("hash for row: "+hash, 1);
			if(set_of_hashes.has(hash)){
				hasDuplicates = true;
				add("hash: "+hash+" is duplicated!", 1);
			}else{
				set_of_hashes.add(hash);
			}
	})
	if(hasDuplicates) add("Error! There are duplicate rows in the data!", 4);		
	return hasDuplicates;
}

/**
 * Checks if a variable is null or undefined or an empty string.
 * @param variable
 * @returns Boolean 
 */
function isNullOrUndefinedOrEmptyString(variable){
	if(isNullOrUndefined(variable))return true;
	var string = String(variable);
	return  string.length===0 || !string.trim();
}

/**
 * Checks if a variable is null or undefined.
 * @param variable
 * @returns Boolean
 */
function isNullOrUndefined(variable) { 
	return variable === null || variable === undefined; 
}

/**
 * Checks if the program and org unit metadata in the excel sheet 
 * is consistent with the selected values in the drop down lists.
 * @returns Boolean value if meta data is consistent.
 */
function isMetaDataValid(){
	orgUnit_id_metadata = metaDataArray[0].OrgUnitId;
	console.log("org unit id excel: " + orgUnit_id_metadata);
	console.log("org unit id form: " + org_unit_id);
	
	program_id_metadata = metaDataArray[0].ProgramId;
	console.log("program id excel: " + program_id_metadata);
				
	//get the id of the selected program
	var program_id_form=$("#programList").val();
	console.log("program id form: " + program_id_form);
	//get the id of the selected org unit: org_unit_id
	
	//test if the ids of program and org unit match with metadata in third sheet
	if(!(program_id_metadata === program_id_form)){
		add("Error! The selected program id: "+program_id_form+" does not match the id in the spreadsheet: " +program_id_metadata+" !", 4);
		console.log("Error! The selected program id: "+program_id_form+" does not match the id in the spreadsheet: " +program_id_metadata+" !");
		return false;
	}
	if(!(orgUnit_id_metadata === org_unit_id)){
		add("Error! The selected org unit id: "+org_unit_id+" does not match the id in the spreadsheet: " +orgUnit_id_metadata+" !", 4);
		console.log("Error! The selected org unit id: "+org_unit_id+" does not match the id in the spreadsheet: " +orgUnit_id_metadata+" !");
		return false;
	}
	return true;
}

/**
 * 
 * @param file The excell file which should be processed.
 * @param isTest Is this a test run?
 * @returns hasErrors A boolean indicating if errors occurred.
 */
function handleFile(f) {
	
	var reader = new FileReader();
	
	reader.onload =
		(function(theFile) {
			return function(e) {
				var data = e.target.result;			
				var wb;				
				var arr = fixdata(data);
				wb = X.read(btoa(arr), {type: 'base64'});	
				readWorkbook(wb);	
				if(hasDuplicates() || !isMetaDataValid()){
					console.log("Error! The metadata is not consistent!");
					add("Error! The metadata is not consistent!", 4);
				}else{	
					processData().then("File was processed.");					
				}
			};
		})(f);
	
	reader.readAsArrayBuffer(f);	
}

/**
 * Processing the data in the excel sheets.
 * 
 * ResultArray is defined and populated in funcxl.js
 * The metadata is read from the third sheet.
 * The data itself is read from first sheet.
 * An example json object:
 *  {
 *	  "program": "eBAyeGv0exc",
 *	  "orgUnit": "DiszpKrYNg8",
 *	  "eventDate": "2013-05-17",
 *	  "status": "COMPLETED",
 *	  "storedBy": "admin",
 *	  "coordinate": {
 *	    "latitude": 59.8,
 *	    "longitude": 10.9
 *	  },
 *	  "dataValues": [
 *	    { "dataElement": "qrur9Dvnyt5", "value": "22" },
 *	    { "dataElement": "oZg33kd9taw", "value": "Male" },
 *	    { "dataElement": "msodh3rEMJa", "value": "2013-05-18" }
 *	  ]
 *	}
 * Source: https://docs.dhis2.org/master/en/developer/html/dhis2_developer_manual_full.html#webapi_events
 * 
 * @param isTest Is this a test run?
 * @returns hasErrors Was the data incorrect and thus not sent to the event API?
 */	  
function processData(){
	return new Promise(
			function (resolve, reject) {

				//Should the geolocation be checked?
				var CheckGeoLocation = document.getElementById("CheckGeoLocation").value == 1;
				var hasErrors = false;
				var rejected = false;		

				//Define a regex pattern for the date time information for the reporting date:
				//2016-12-01T00:00:00.000
				var DateTimePattern = /[1-2][0-9]{3}-[0-1][0-9]-[0-3][0-9]T[0-9]{2}:[0-9]{2}:[0-9]{2}.[0-9]{3}/;
				
				//Define a more simple alternative regex pattern which only describes the date
				//2016-12-01
				var AlternativeDateTimePattern = /[1-2][0-9]{3}-[0-1][0-9]-[0-3][0-9]/;
				
				dataValues = [];
				eventDataValues = {};
				errorString = "";
				isAggDataAvailable = false;	
				eventDataValues.events = [];
				var lineNr=0;

				//Iterate over all the rows, which are the inner arrays of the json array of arrays.
				resultArray.forEach( function (arrayItem){		
					
					lineNr++;
					var eventDataValue = {};
					eventDataValue.status="COMPLETED";
					eventDataValue.storedBy="admin";
					eventDataValue.program = program_id;
					eventDataValue.orgUnit = org_unit_id;	
					var point = new Array();
										
					if(CheckGeoLocation){						
						//check the first three columns (obligatory values):
						for (const [column, label] of obligatoryDataElementsRowLabelMap.entries()) {

							if(!arrayItem.hasOwnProperty(label)){
								//if not, then we have to reject this line / this event 
								add("The value of "+label+" in column "+column+
										" is undefined for input line "+ lineNr +" "
										+JSON.stringify(arrayItem), 4);
								add("Please read the log messages attentively and fix the problem! ", 4);
								add("You may have to set the log level to \"trace\" or \"debug\".", 4);
								rejected=true;		
								hasErrors=true;	
							}						
							switch(label){
							case "ReportingDate": 
								//Does the date entered match the regex pattern?
								//If not, reject the input data!
								if((! DateTimePattern.test(arrayItem.ReportingDate)) && (! AlternativeDateTimePattern.test(arrayItem.ReportingDate))){
									rejected=true;		
									hasErrors=true;
									add("Invalid reporting date/time entered: "+ arrayItem.ReportingDate ,4);
									add("Row: "+lineNr+"->The reporting date has to be entered in the following format: 2016-12-01T00:00:00.000 !",4);
									break;
								}	

							case "Latitude":  
								if(isNaN(arrayItem.Latitude) || (Math.abs(parseInt(arrayItem.Latitude))>90.0)) {
									rejected=true;		
									hasErrors=true;
									add("Row: "+lineNr+"->The entered value "+arrayItem.Latitude+" for latitude is not a valid number!",4);
									break;
								}else{
									point[0]=arrayItem.Latitude;
								}
							case "Longitude": 
								if(isNaN(arrayItem.Longitude) || (Math.abs(parseInt(arrayItem.Longitude))>180.0)) {
									rejected=true;		
									hasErrors=true;
									add("Row: "+lineNr+"->The entered value "+arrayItem.Longitude+" for longitude is not a valid number!",4);
									break;
								}else{
									point[1]=arrayItem.Longitude;
								}						
							}	
						}		
						
						add("Location: "+point,1)
						//If no polygon has been supplied for org-unit, its parent organisation
						//and the grandparent organisation, I have to skip this test.
						//Instead a warning is printed, because it would be better to do this test.
						if(isNullOrUndefined(org_unit_polygon)){
							add("Row: "+lineNr+"->Warning: No polygon information has been supplied!",3)
							add("Warning: The supplied location can not be validated!",3)
						}else{ 
							//Check if the location is within any of the polygons
							//of the org. unit:
							if(!insideAnyPolygon(point, org_unit_polygon)){
								add("Row: "+lineNr+"->Invalid location! The location "+point+" is not located within the polygon of the org unit "+org_unit_name+" !", 4)
								add("The polygon of this org-unit is:"+org_unit_polygon, 2)
								add("Row: "+lineNr+"->Fatal error! The data import is canceled!", 4);
								rejected=true;
								return hasErrors;
							}
						}			
						if(rejected){
							add("Row: "+lineNr+"->Fatal error! The data import is canceled!", 4);
							add("Please read the log messages attentively and fix the problem! ", 3);
							add("You may have to set the log level to \"trace\" or \"debug\".", 3);
						}			
						
					//If the geolocation is not provided only the reporting date will be checked.
					}else{
						if(!arrayItem.hasOwnProperty("ReportingDate")){
							//if not, then we have to reject this line / this event 
							add("The value of ReportingDate in column "+column+
									" is undefined for input line "+ lineNr +" "
									+JSON.stringify(arrayItem), 4);
							add("Please read the log messages attentively and fix the problem! ", 4);
							add("You may have to set the log level to \"trace\" or \"debug\".", 4);
							rejected=true;		
							hasErrors=true;	
						}
						if((! DateTimePattern.test(arrayItem.ReportingDate)) && (! AlternativeDateTimePattern.test(arrayItem.ReportingDate))){
							rejected=true;		
							hasErrors=true;
							add("Invalid reporting date/time entered: "+ arrayItem.ReportingDate ,4);
							add("Row: "+lineNr+"->The reporting date has to be entered in the following format: 2016-12-01T00:00:00.000 !",4);
						}	
					}
					
					//This is the event timestamp.
					eventDataValue.eventDate = arrayItem.ReportingDate;
					eventDataValue.eventDate = eventDataValue.eventDate.replace(/['"]+/g,'');
					eventDataValue.coordinate = {};
					eventDataValue.coordinate.latitude = arrayItem.Latitude;
					eventDataValue.coordinate.longitude = arrayItem.Longitude;						
					eventDataValue.dataValues = [];						
					
					//Count missing data elements per row.
					var missingDataElement=0;						
					
					//here all option values have to be available
					while(optionsToQuery>0){
						sleep(1000);		
					}						
					if(optionsToQuery>0){
						add("Row: "+lineNr+"->Error: Some values of optionals are not yet available!",4);
						rejected=true;		
						hasErrors=true;
					}

					for(let dataElement of dataElementIDs)
					{
						var dv = {};
						var label = dataElementsLabel.get(dataElement);
						var valueType = dataElementsValueType.get(dataElement);
						var optionSetId = dataElementsOptionSet.get(dataElement);
						if(dataElementsHasOptionSet.get(dataElement) && isNullOrUndefinedOrEmptyString(optionSetId)){							
							add("Row: "+lineNr+"->Error! The option set is not defined: ", optionSetId, 4);
						}								
						dv.dataElement = dataElement;							
						//Test if the json object representing the row has the property label:
						if(arrayItem.hasOwnProperty(label)){								
							var rawData = arrayItem[label];															
							//Depending on the type of value, 
							//do some cleaning of the data:
							switch (valueType) {								 
							case "COORDINATE":
							case "LONG_TEXT":
							case "TEXT":
								add("before cleaning: "+rawData, 1);
								//Remove all inner quotes and escapes from strings
								if(typeof rawData === "string"){
									rawData = rawData.replace(/['"]+/g,'');
								}else{
									rawData = rawData.replace(/['"]+/g,'');
								}
								add("Row: "+lineNr+"->after cleaning: "+rawData, 1);
								break;
							case "INTEGER_POSITIVE":
								add("Row: "+lineNr+"->Before cleaning: "+rawData, 1);									 
								//Remove negative or zero values for data type INTEGER_POSITIVE
								if(rawData<=0)rawData=void 0;
								add("Row: "+lineNr+"->After cleaning: "+rawData, 1);
								break;
							case "TRUE_ONLY":
								add("Row: "+lineNr+"->Before cleaning: "+rawData, 1);
								if(typeof rawData === "string"){
									rawData = rawData.replace(/['"]+/g,'');
								}else{
									rawData = rawData.replace(/['"]+/g,'');
								}
								//Using regular expressions, replace all char sequences with letters 
								// T/t R/r U/u E/e with "true"
								rawData.replace(/true/gi, "true");
								if(!(isTrue.test(rawData))){
									rawData=void 0;
								}
								add("Row: "+lineNr+"->After cleaning: "+rawData, 1);
								break;
							default:
								add("Row: "+lineNr+"->No cleaning operation defined for data type: "+valueType, 1);
							}								
							//check if value is within set of valid options for option sets:
							if(dataElementsHasOptionSet.get(dataElement))
							{
								add("Row: "+lineNr+"->Data element with ID:\""+dataElement+"\" label \"" + label + "\" and value type \"" + valueType +
										"\" has option set with ID: \""+ optionSetId +"\"", 1);
								if(optionMap.has(optionSetId)){
									add("Row: "+lineNr+"->Option map has "+optionMap.size +" valid values.", 1);
									var optionSet = optionMap.get(optionSetId);
									add("Row: "+lineNr+"->Option set has "+optionSet.length+" valid values:", 1);
									for (var i = 0; i < optionSet.length; i++) {
										if(options.has(optionSet[i])){
											add("Row: "+lineNr+"> Option "+i+" Id: "+optionSet[i]+" Value: "+ options.get(optionSet[i]), 1);
										}else{
											add("Row: "+lineNr+"->Option "+i+" Id: "+optionSet[i], 4);
										}
									}										
									var valueInOptionSet = false;									
									for (var i = 0; i < optionSet.length; i++) {											
										if(options.has(optionSet[i])){
											var option = options.get(optionSet[i]);
										}else{
											add("Row: "+lineNr+"->Option with ID: "+optionSet[i]+" is not available",4);
											add("Row: "+lineNr+"->Available options are: ",4);
											for (const [k,v] of options.entries()) {
												add("key: "+k+"\tvalue: "+v, 4);
											}	
											rejected=true;		
											hasErrors=true;
											//return true;	
										}											
										switch (valueType) {								 
										case "LONG_TEXT":
										case "TEXT":
											rawData = String(rawData);
											//If the text string matches the option (upper/lower case is ignored)
											if(rawData.toUpperCase() === option.toUpperCase()){
												add("Row: "+lineNr+"->Value: "+rawData+" matches option "+option+"!",1);
												rawData = option;
												valueInOptionSet = true;
												break;
											}else{
												add("Row: "+lineNr+"->Value: " + rawData + " does NOT match option "+option+"!",1);	
												break;
											}
										default:  
											//If the text string matches the option (upper/lower case is ignored)
											if(rawData === option){
												add("Row: "+lineNr+"->Value: "+rawData+" of data type "+ valueType +" matches option "+option+"!",1);
												rawData = option;
												valueInOptionSet = true;
												break;
											}else{
												add("value: " + rawData +" of data type "+ valueType + " does NOT match option "+option+"!",1);
												break;
											}
										}										   
									}
									if(valueInOptionSet == false){
										add("Row: "+lineNr+"->Invalid value \""+rawData+"\" for option set: "+optionSetId, 4);
										rejected=true;		
										hasErrors=true;
									}
								}else{
									add("Row: "+lineNr+"->Error! No options defined for option set with ID: "+optionSetId, 4);
								}
							}
							dv.value = rawData;
							if(!rejected){
								eventDataValue.dataValues.push(dv);
							}
						}else{	
							missingDataElement++;
							//Abort if an obligatory data element is missing
							//or if all data elements are missing.
							if((dataElementsCompulsory.get(dataElement)==true)||(missingDataElement==dataElementsLabel.size)){
								if(dataElementsCompulsory.get(dataElement)==true){
									add("Row: "+lineNr+"->The value of the compulsory "+label+" in column "+column+
											" is undefined for input line "+ lineNr +" "
											+JSON.stringify(arrayItem), 4);
								}else{
									add("Row: "+lineNr+"->No single data element is supplied in input line "+ lineNr +" "
											+JSON.stringify(arrayItem), 4);
								}
								rejected=true;		
								hasErrors=true;	
							}								
						}
					}
					if(!rejected){
						eventDataValues.events.push(eventDataValue);
					}
				});			
								
				add("Processed events: "+resultArray.length, 3);

				if(!rejected){					
					importData().then(resolve());					
				}else{
					reject("The data upload was rejected as a whole. No data was uploaded");
				}
			}
	)	
}

/**
 * Function checks if a given point is inside a polygon.
 * Sources:
 * https://github.com/substack/point-in-polygon
 * http://stackoverflow.com/questions/22521982/js-check-if-point-inside-a-polygon
 * http://www.ecse.rpi.edu/Homepages/wrf/Research/Short_Notes/pnpoly.html
 * @param point The coordinates of the point.
 * @param vs The coordinates of the polygon.
 * @returns Boolean, is the point located within the polygon?
 */
function inside(point, vs) {
    // ray-casting algorithm based on
    // http://www.ecse.rpi.edu/Homepages/wrf/Research/Short_Notes/pnpoly.html
    var x = point[0], y = point[1];
    if(typeof x == 'string'){
  	  x = JSON.parse(x);	  
    }
    if(typeof y == 'string'){
  	  y = JSON.parse(y);	  
    }
    var inside = false;
    for (var i = 0, j = vs.length - 1; i < vs.length; j = i++) {
        var xi = vs[i][0], yi = vs[i][1];
        var xj = vs[j][0], yj = vs[j][1];
        var intersect = ((yi > y) != (yj > y))
            && (x < (xj - xi) * (y - yi) / (yj - yi) + xi);
        if (intersect) inside = !inside;
    }
    return inside;
}

/**
 * Is the point within any of the polygons?
 * "Polygons" may consist of several disjunct areas.
 * This function is a wrapper for the inside() function, 
 * calling it for each inner array decribing a polygon.
 * @param point The coordinates of the point.
 * @param vs The coordinates of the polygon.
 * @returns
 */
function insideAnyPolygon(point, vs) {    
  if(typeof vs == 'string'){
	  vs = JSON.parse(vs);	  
  }
  if(typeof point == 'string'){
	  point = JSON.parse(point);	  
  }
  var isInside = false;    
  var isArray = Array.isArray(vs);
  var lengthArray = 0;
  if(isArray){
  	lengthArray = vs.length;
  }    
  for (var dim = 0; dim < lengthArray; dim++) {
	  	console.log("polygon dimension: "+dim)
	  	//console.log(vs[dim])
	  	isInside = isInside || inside(point, vs[dim]);
	  	console.log("Point is inside polygon: "+isInside)    
	  	if(isInside) return isInside;
  }    
  return isInside;
}

/**
 * 
 * Here, I have to define which is the last valid column for each spreadsheet
 * The second sheet contains the data, here the nr of columns is always 4 + nr of data elements.
 * This function has to be udpated whenever the column layout of the template changes. 
**/
function getSheetEndColumn() {
	
	if(sheetEndColumns.length==0){
		//Here, I have to define which is the last valid column for each spreadsheet
		//The second sheet contains the data, here the nr of columns is always 4 + nr of data elements:
		//Update the column index if the layout of the template changes!
		columnIndex=4+dataElementIDs.size;	
		var div = Math.floor(columnIndex/26);
		var rem = columnIndex % 26;	
		var lastColumn = "";
		
		if(div==0){
			lastColumn=letters[rem];
		}else{
			lastColumn=letters[div].concat(letters[rem])
		}		
		console.log("div: "+ div +"rem: "+ rem ,"letter:"+ lastColumn)
		
		sheetEndColumns.push(lastColumn);
		//The second sheet always contains a legend with two columns.
		//sheetEndColumns.push('B');
		//The third sheet always contains a legend with two columns.
		//sheetEndColumns.push(letter[div].concat(letter[rem]));
	}
	return lastColumn
}
 
/**
 * Converts an array of arrays into a spreadsheet.
 * 
 * @param data The json array. 
 * @returns
 */
function sheet_from_array_of_arrays(data) {
	var ws = {};
	var range = {s: {c:10000000, r:10000000}, e: {c:0, r:0 }};
	for(var R = 0; R != data.length; ++R) {
		for(var C = 0; C != data[R].length; ++C) {
			if(range.s.r > R) range.s.r = R;
			if(range.s.c > C) range.s.c = C;
			if(range.e.r < R) range.e.r = R;
			if(range.e.c < C) range.e.c = C;
			var cell = {v: data[R][C] };
			if(cell.v == null) continue;
			var cell_ref = XLSX.utils.encode_cell({c:C,r:R});
			
			if(typeof cell.v === 'number') cell.t = 'n';
			else if(typeof cell.v === 'boolean') cell.t = 'b';
			else if(cell.v instanceof Date) {
				cell.t = 'n'; cell.z = XLSX.SSF._table[14];
				cell.v = datenum(cell.v);
			}
			else cell.t = 's';
			
			ws[cell_ref] = cell;
		}
	}
	if(range.s.c < 10000000) ws['!ref'] = XLSX.utils.encode_range(range);
	return ws;
}
  
/**
 * Creates a new workbook.
 * @returns A new workbook.
 */
function Workbook() {
	if(!(this instanceof Workbook)) return new Workbook();
	this.SheetNames = [];
	this.Sheets = {};
}

/**
 * Converts a string to an array buffer.
 * @param s
 * @returns
 */
function s2ab(s) {
	var buf = new ArrayBuffer(s.length);
	var view = new Uint8Array(buf);
	for (var i=0; i!=s.length; ++i) view[i] = s.charCodeAt(i) & 0xFF;
	return buf;
}

//JQuery syntax for: wait until the html document has been loaded,
//then run the function readProperties(). 
$(document).ready(readProperties());

var X = XLSX;
var XW = {
	/* worker message */
	msg: 'xlsx',
	/* worker scripts */
	rABS: './xlsxworker2.js',
	norABS: './xlsxworker1.js',
	noxfer: './xlsxworker.js'
};

var rABS = typeof FileReader !== "undefined" && typeof FileReader.prototype !== "undefined" && typeof FileReader.prototype.readAsBinaryString !== "undefined";
if(!rABS) {
	document.getElementsByName("userabs")[0].disabled = true;
	document.getElementsByName("userabs")[0].checked = false;
}

var use_worker = typeof Worker !== 'undefined';
if(!use_worker) {
	document.getElementsByName("useworker")[0].disabled = true;
	document.getElementsByName("useworker")[0].checked = false;
}

var transferable = use_worker;
if(!transferable) {
	document.getElementsByName("xferable")[0].disabled = true;
	document.getElementsByName("xferable")[0].checked = false;
}

var wtf_mode = false;

function fixdata(data) {
	var o = "", l = 0, w = 10240;
	for(; l<data.byteLength/w; ++l) o+=String.fromCharCode.apply(null,new Uint8Array(data.slice(l*w,l*w+w)));
	o+=String.fromCharCode.apply(null, new Uint8Array(data.slice(l*w)));
	return o;
}

var resultArray = [];
var metaDataArray = [];

/**
 * 
 * @param workbook The excell workbook which should be processed.
 * @param isTest Is this a test run?
 * @returns Returns a function which processes the excel sheet.
 */
function readWorkbook(workbook, isTest) {	
	//the data of the first sheet (data itself plus header)
	var result = [];	
	//the metadata from the third sheet
	var metadata = [];	
	lastColumn = getSheetEndColumn();		
	//only read the first four sheets
	for(var currentSheetNumber = 0; currentSheetNumber < 4; currentSheetNumber++) {
		sheetName = workbook.SheetNames[currentSheetNumber]
		//read data from second sheet
		if(currentSheetNumber==1){
			resultArray = X.utils.sheet_to_json(workbook.Sheets[sheetName], currentSheetNumber, lastColumn);			
		}
		//ignore sheet nr 3, it only contains a legend
		//read metadata from sheet nr 4
		if(currentSheetNumber==3){			
			metaDataArray = X.utils.sheet_to_json(workbook.Sheets[sheetName], currentSheetNumber, lastColumn);
		}
	}
}

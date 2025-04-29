
//-----------------------------------------------------------
var xlValues = -4163;
var xlPart = 2;
var xlWhole = 1;
var xlByRows = 1;
var xlNext = 1;
var xlCellTypeLasstCell = 11;

var fso = new ActiveXObject("Scripting.FileSystemObject") 
var  dt = new Date();
var sDt = dt.getFullYear() 
+ ('0' + (dt.getMonth() + 1)) .slice( - 2) 
+ ('0' + dt.getDate()) .slice( - 2) 
+ ('0' + dt.getHours()) .slice( - 2) 
+ ('0' + dt.getMinutes()) .slice( - 2) 
+ ('0' + dt.getSeconds()) .slice( - 2);
//-----------------------------------------------------------
var currentpath = new ActiveXObject("WScript.Shell") .CurrentDirectory 

/*
var logFilePath = currentpath + "/dutysheet_log_" + sDt + ".txt";
var logFile = fso.CreateTextFile(logFilePath, true);
// var currentpath = new ActiveXObject("Scripting.FileSystemObject") .GetFolder(".") .Path
logFile.WriteLine("logFilePath:" + logFilePath);
*/
//--------------------------------------------
var tbl_info = {  lnm:"C6"  };
var ctrl_info = {  folder:"D:/svn/optiplex5040/sakura_macro/test",book:"link_test.xlsx", sht:"log_check"  };
//--------------------------------------------
make_XL();

var ctrl_book=open_excel(ctrl_info.folder,ctrl_info.book);
var ctrl_sht=ctrl_book.Worksheets(ctrl_info.sht);

var  wk_sht = XL.ActiveSheet;
wk_sht.PasteSpecial();
var rng=XL.Selection;
var all_col=rng.EntireColumn;
all_col.ColumnWidth=50;
make_range_square(rng,0);

var max_setting_row=ctrl_sht.UsedRange.Rows(ctrl_sht.UsedRange.Rows.Count).Row;


var show_txt="";
var style_list=[];
var  col_nm=1;
var row=1;
while(row<=max_setting_row){
	var tmp=ctrl_sht.Cells(row,col_nm).Value;
	var reg_txt=tmp?tmp.toString().replace(/(^\s*)|(\s*$)/,""):null;
	if(reg_txt){
		show_txt+= reg_txt + "\n";
		var cell=ctrl_sht.Cells(row,col_nm);
		var regexp=new RegExp(reg_txt);
		style_list.push({ chk:regexp,  c_color:cell.Interior.Color,  f_color:cell.Font.Color,   bold:cell.Font.Bold });
	}
	row++;
}

var  all_wk_rng=wk_sht.UsedRange;

for (var i = 1; i <= all_wk_rng.Rows.Count; i++) {
    for (var j = 1; j <= all_wk_rng.Columns.Count; j++) {
        var cell = all_wk_rng.Cells(i, j);
        var tmp = cell.Value;
	var txt=tmp?tmp.toString().replace(/(^\s*)|(\s*$)/,""):null;
        if (txt) {
		for (  idx in style_list){
			var obj=style_list[idx];
			var mat=txt.match(obj.chk);
			if(mat){
				cell.Interior.Color=obj.c_color;
				cell.Font.Color=obj.f_color;
				cell.Font.Bold=obj.bold;
			}
		}
        }
    }
}

AddTail(show_txt);




function manage_link(_sht,_rng, _sht_nm) {
	if(_rng.Hyperlinks.Count>0){
		_rng.Hyperlinks.Delete();
	}

	_sht.Hyperlinks.Add(_rng,"", _sht_nm+"!A1",  _sht_nm);
	sht_nm_list[_sht_nm]=20;

}

function search_folder(_list, _path, _sub, _ext) {

	var fsofolder = fso.GetFolder(_path + "/" + _sub);
	var  folders = new Enumerator(fsofolder.SubFolders);
	for (; !folders.atEnd(); folders.moveNext()) {
		var sub_dir = folders.item() .Name;
		search_folder(_list, _path, _sub ? _sub + "/" + sub_dir:sub_dir, _ext);
	}
	var  files = new Enumerator(fsofolder.Files);
	for (; !files.atEnd(); files.moveNext()) {
		var fname = files.item() .Name;
		if (_ext == "*") {
			_list.push(_sub ? _sub + "/" + fname:fname);
		}
		else {
			var	mat = fname.match(/^(\S+\.)(\w+)$/);
			if (mat) {
				var ext = mat[2].toLowerCase();
				if (typeof(_ext) == 'string') {
					if (ext == _ext) {
						_list.push(_sub ? _sub + "/" + fname:fname);
					}
				}
				else if (typeof(_ext) == "object") {
					if (_ext[ext] > 0) {
						_list.push(_sub ? _sub + "/" + fname:fname);
					}
				}
			}
		}
	}
}

function   CreateFolders(_folderspec) 
{
	var parFsoFolder = fso.GetParentFolderName(_folderspec);
	if ( !fso.FolderExists(parFsoFolder)) {
		CreateFolders(parFsoFolder);
	}
	if ( !fso.FolderExists(_folderspec)) {
		fso.CreateFolder(_folderspec);
	}
	return;
}


function  make_XL() {
	try
	{
		XL = GetObject("", "Excel.Application");
	}
	catch(e) 
	{
		XL = new ActiveXObject("Excel.Application");
		XL.Visible = true;
	}
}

function open_excel(path, book) {
	make_XL();
	try
	{
		var excel_book = XL.workbooks(book);
	}
	catch(e) 
	{
		var fso = new ActiveXObject("Scripting.FileSystemObject");
		if (typeof(path) == "string") 
		{
			excel_book = XL.workbooks.Open(path + "\\" + book);
		}
		else {
			for (var idx in path) {
				if (fso.FileExists(path[idx] + "\\" + book)) {
					excel_book = XL.workbooks.Open(path[idx] + "\\" + book);
					break;
				}
			}
		}
		if ( !excel_book) {
			throw new Error("=====[" + book + "]");
		}
	}
	return excel_book;
}
function getFromSlashEscapeStr(str) {
	var arr = str.split("");
	var flgEsc = 0;
	var ret = "";
	var idx = 0;
	while (idx < arr.length) {
		if (flgEsc == 0) {
			if (arr[idx] == "\\") {
				flgEsc = 1;
			}
			else {
				ret += arr[idx];
			}
		}
		else {
			switch(arr[idx]) 
			{
				case "\\":  ret += arr[idx]; break;
				case "t": ret += "\t"; break;
				case "r": ret += "\r"; break;
				case "n": ret += "\n"; break;
				default : ret += arr[idx];
			}
			flgEsc = 0;
		}
		++idx;
	}
	return ret;
}
function msg_box(msg, title) 
{
	if ( !title) {
		title = "ソース移行"
	}
	var WSHShell = new ActiveXObject("WScript.Shell");
	WSHShell. Popup(msg, 0, title, 1);
}
function inputBox_SC(prompt, title, def) 
{
	var result;
	var objScr;
	// objScr = new ActiveXObject("MSScriptControl.ScriptControl");
	objScr = new ActiveXObject("ScriptControl");
	objScr.language = "VBScript";
	objScr.addCode(
		"Function getInput()" + 
		'    getInput = InputBox("' + prompt + '", "' + title + '", "' + def + '")' + 
		"End Function");
	result = objScr.eval("getInput");
	objScr = null;
	return result;
}
function MsgBox(prompt, buttons, title) 
{
	if ( !title) {
		title = "ソース移行"
	}
	var result;
	var objScr;
	objScr = new ActiveXObject("MSScriptControl.ScriptControl");
	objScr.language = "VBScript";
	objScr.addCode(
		"Function vbsMsgbox()" + 
		'    vbsMsgbox = MsgBox("' + prompt + '", ' + buttons + ', "' + title + '")' + 
		"End Function");
	result = objScr.eval("vbsMsgbox");
	objScr = null;
	return result;
}
function convert_RGB(r, g, b) {
	var color = 0;
	color += r;
	color += g << 8;
	color += b << 16;
	return color;
}

function make_range_square(range, inner_idx) 
{
	var xlDiagonalDown = 5;
	var xlDiagonalUp = 6;
	var xlInsideVertical = 11;
	var xlInsideHorizontal = 12;

	var xlNone = -4142;

	var xlEdgeLeft = 7;
	var xlEdgeTop = 8;
	var xlEdgeBottom = 9;
	var xlEdgeRight = 10;

	var xlContinuous = 1;
	var xlAutomatic = -4105;
	var xlThin = 2;
	var list_all = [ [], [xlInsideVertical], [xlInsideHorizontal], [ xlInsideVertical, xlInsideHorizontal  ]  ]; // , xlDiagonalDown, xlDiagonalUp

	list1 = list_all[inner_idx];

	for (var idx in list1) 
	{
		range.Borders(list1[idx]) .LineStyle = xlContinuous;
		range.Borders(list1[idx]) .ColorIndex = xlAutomatic;
		range.Borders(list1[idx]) .TintAndShade = 0;
		range.Borders(list1[idx]) .Weight = xlThin;
	}

	for (var idx in list = [ xlEdgeLeft, xlEdgeTop, xlEdgeBottom, xlEdgeRight ]) {
		range.Borders(list[idx]) .LineStyle = xlContinuous;
		range.Borders(list[idx]) .ColorIndex = xlAutomatic;
		range.Borders(list[idx]) .TintAndShade = 0;
		range.Borders(list[idx]) .Weight = xlThin;
	}
}


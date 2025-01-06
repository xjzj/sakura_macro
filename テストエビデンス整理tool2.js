

var msoShapeFlowchartAlternateProcess=62;
var msoShapeRoundedRectangularCallout=106;


var msoShapeFlowchartProcess=61;
var msoThemeColorAccent1=5;
var msoThemeColorAccent2 = 6;
var msoThemeColorAccent4 = 8;
var msoThemeColorBackground1 = 14;

//-----------------------------------------------------------

var txt="test aaaaaaaa";
//-----------------------------------------------------------

make_XL();
var ev_sheet =XL.ActiveSheet;

var col_y=ev_sheet.Range("Y1").Column;
var tmp=ev_sheet.Cells(1,col_y).Value;
var pos=tmp?Number(tmp):400
var locations=[];
var iColor=[ msoThemeColorAccent2,msoThemeColorAccent1 ];
var cnt=ev_sheet.Shapes.Count;
// msg_box("ev_sheet.Name:" + ev_sheet.Name);
// msg_box("cnt:" + cnt);
var idx=0;
while(idx<cnt){
	idx++;
	if(ev_sheet.Shapes(idx).Type==13){
		var sp=ev_sheet.Shapes(idx);
		var ii=(idx-1+cnt)%2;
		locations.push({ top:sp.Top, height:sp.Height , color:iColor[ii] });
	}

}

idx=0;
var px=188;
var py=pos;
for(var i in locations ){
	if(locations[i]){
		py=locations[i].top + (locations[i].height)/2;
	}
	var shape1=ev_sheet.Shapes.AddShape(msoShapeFlowchartAlternateProcess, px, py, 400, 100);
	shape1.Fill.Visible=false;
	shape1.Line.ForeColor.ObjectThemeColor=locations[i].color;
	shape1.Line.Weight=3;
	shape1.Line.Transparency=0.30;

	var shape2=ev_sheet.Shapes.AddShape(msoShapeRoundedRectangularCallout, px+600, py, 400, 100);
	shape2.Fill.Visible=true;
	shape2.Fill.ForeColor.ObjectThemeColor=msoThemeColorBackground1;

	shape2.Line.ForeColor.ObjectThemeColor=locations[i].color;
	shape2.Line.Weight=1;
	shape2.TextFrame2.TextRange.Characters.Text=txt;
	shape2.Adjustments.Item(1)=-1.0396;
	shape2.Adjustments.Item(2)=-0.0396;
	shape2.TextFrame2.TextRange.Font.Fill.ForeColor.RGB=convert_RGB(0,0,0);
	py+=400;
	idx++;
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



var  show_text="";
show_text+="前提:"+get_date() + "\r\n";

var parPath = "";
var workPath = "";
var workSubDir = "";

var fso = new ActiveXObject("Scripting.FileSystemObject") 
var objShell = new ActiveXObject("Shell.Application") 
var objFolder = objShell.BrowseForFolder(0, "Select Folder", 0, "");
var parFolder = objFolder.ParentFolder;
if (parFolder) {
	parPath = parFolder.Self.Path;
}
if (objFolder) {
	workPath = objFolder.Self.Path;
	// workSubDir = objFolder.Self.Name;
}

var tbl_info = { foler:"D:/tmp/test_get_tbl_name/テスト", book:"tbl_list.xlsx",  jnm:"K5", nm:"BB5" };

var list_book = open_excel(tbl_info.folder, tbl_info.book);
var list_sheet = list_book.Worksheets(1);

var col_jnm = list_sheet.Range(tbl_info.jnm).Column;
var col_nm = list_sheet.Range(tbl_info.nm).Column;

var  work_path="C:/work/working";

path_list = [];
// search_folder(path_list, parPath + "/" + workSubDir, "", "*");
search_folder(path_list, workPath, "", "xlsx");

show_text+= '----------'+ "\r\n";
show_text+="処理開始:" + get_date() + "\r\n";
show_text+="全件:" +   path_list.length+ "\r\n";
show_text+= '----------'+ "\r\n";

AddTail(show_text);
var err_txt="";
var row=1;

make_XL();


list_sheet.Cells(row, 1)="種類";
list_sheet.Cells(row, 2)="ファイル名";
list_sheet.Cells(row, 3)="テーブル名(日本語)";
list_sheet.Cells(row, 4)="テーブル名(英文字)";
list_sheet.Cells(row, 5)="リンク";

var  head_range = list_sheet.Range(list_sheet.Cells(row, 1), list_sheet.Cells(row, 5));
head_range.Font.Bold = true;
head_range.Interior.Color = convert_RGB(200, 210, 240);
row++;


var st_list={}
for( var i in path_list){
	var obj=path_list[i];
	var subdir=obj.sub;
	var fname=obj.nm;
	var mat=fname.match(/^DEDA6-000000-FED183_(\S+).xlsx$/);
	if( mat ){
		var flg=0;
		get_tbl(workPath + "/"+subdir, fname, subdir);
		AddTail("完了:" + i + ":" + fname+"\r\n");

	}
	else{
		err_txt+="err1●"+subdir + ":"  +fname + "\r\n" ;
	}
	
}
make_range_square(list_sheet.Range(list_sheet.Cells(1, 1), list_sheet.Cells(row-1, 4)), 3);


AddTail(err_txt);
show_text="処理終了:" + get_date() + "\r\n";
AddTail(show_text);

AddTail("-----------------------------------------\r\n");
function get_tbl(folder,fxlsx, _sub){
	var tbl_book = open_excel(folder, fxlsx);
	if( tbl_book.Sheets.Count > 1){
		err_txt+="err2●"+folder + ":"  +fxlsx + "\r\n" ;
	}
	var tbl_sheet = tbl_book.Worksheets(1);
	var  jnm=tbl_sheet.Cells(5, col_jnm).Value;
	var  enm=tbl_sheet.Cells(5, col_nm).Value;
	var obj=st_list[enm];
	if(!obj){
		list_sheet.Cells(row, 1)=_sub;
		list_sheet.Cells(row, 2)=fxlsx;
		list_sheet.Cells(row, 3)=jnm;
		list_sheet.Cells(row, 4)=enm;
		list_sheet.Cells(row, 5)="リンク";
		// list_sheet.Cells(row, 6)=enm;
	                                        
	                                        
		 tbl_sheet.Copy(null, list_book.Sheets(list_book.Sheets.Count));
		 list_book.Sheets(list_book.Sheets.Count).Name=enm;
		//
		tbl_book.Close(false);
		 list_sheet.Hyperlinks.Add(list_sheet.Cells(row,5),"",enm+"!"+tbl_info.jnm,jnm);
		 // list_sheet.Hyperlinks.Add(list_sheet.Cells(row,6),"",enm+"!"+tbl_info.nm,enm);
		 st_list[enm]={ j:jnm, e:enm , f:fxlsx, d:_sub  }
		row++;
	}
	else{
		tbl_book.Close(false);
		err_txt+="err3●"+enm + "[" +folder + ":" +fxlsx +":" +jnm +  "][" +obj.d + ":" +obj.f + ":" +obj.j + "]"  + "\r\n" ;
	}

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
			_list.push({ sub:_sub, nm:fname  });
		}
		else {
			var	mat = fname.match(/^(\S+\.)(\w+)$/);
			if (mat) {
				var ext = mat[2].toLowerCase();
				if (typeof(_ext) == 'string') {
					if (ext == _ext) {
						_list.push({ sub:_sub, nm:fname  });
					}
				}
				else if (typeof(_ext) == "object") {
					if (_ext[ext] > 0) {
						_list.push({ sub:_sub, nm:fname  });
					}
				}
			}
		}
	}
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

function convert_RGB(r, g, b) {
	var color = 0;
	color += r;
	color += g << 8;
	color += b << 16;
	return color;
}

function make_range_square(range, inner_idx, _flg ) 
{
	if(_flg==undefined) 
	{
		_flg=true;
	}

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
	
	var _LienStyle=xlContinuous;
	if(!_flg){
		_LienStyle=xlNone;
	}
	
	var list_all = [ [], [xlInsideVertical], [xlInsideHorizontal], [ xlInsideVertical, xlInsideHorizontal  ]  ]; // , xlDiagonalDown, xlDiagonalUp

	list1 = list_all[inner_idx];

	for (var idx in list1) 
	{
		// logFile.WriteLine("_LienStyle:" + _LienStyle +"]inner_idx[" + inner_idx +"]_flg[" + _flg  + "]" );
		range.Borders(list1[idx]) .LineStyle = _LienStyle;
		if(_flg){
			range.Borders(list1[idx]) .ColorIndex = xlAutomatic;
			range.Borders(list1[idx]) .TintAndShade = 0;
			range.Borders(list1[idx]) .Weight = xlThin;
		}
	}
	if(_flg){
		for (var idx in list = [ xlEdgeLeft, xlEdgeTop, xlEdgeBottom, xlEdgeRight ]) {
			range.Borders(list[idx]) .LineStyle = _LienStyle;
			range.Borders(list[idx]) .ColorIndex = xlAutomatic;
			range.Borders(list[idx]) .TintAndShade = 0;
			range.Borders(list[idx]) .Weight = xlThin;
		}
	}
}

function get_date(){
	var  dt = new Date();
	var sDt = dt.getFullYear() 
	+ ('0' + (dt.getMonth() + 1)) .slice( - 2) 
	+ ('0' + dt.getDate()) .slice( - 2) 
	+ ('0' + dt.getHours()) .slice( - 2) 
	+ ('0' + dt.getMinutes()) .slice( - 2) 
	+ ('0' + dt.getSeconds()) .slice( - 2);
	return  sDt;
}

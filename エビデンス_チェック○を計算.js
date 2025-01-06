
var sOut = "";
var sTest = "";
make_XL();

var ev_sheet = XL.ActiveSheet;
var sel = XL.Selection;
if (sel) {
	var col = sel.Column;
	var row = sel.Row;
	for (var j = col; j <= ev_sheet.UsedRange.Columns.Count; j++) {
		var cnt = 0;
		var tmp = ev_sheet.Cells(row, j).Value;
		var  cs_nm =tmp?tmp.toString().replace(/(^\s*)|(\s*$)/g, "") :"";
		for (var i = row; i <= ev_sheet.UsedRange.Rows.Count; i++) {
			var _tmp = ev_sheet.Cells(i, j) .Value;
			var cont = _tmp?_tmp.toString().replace(/(^\s*)|(\s*$)/g, "") :"";
			if (cont == "›" || cont == "Z") {
				cnt++;
			}

		}
		if (cnt > 0) {
			sOut +=cs_nm + "\t" + cnt + "\r\n";
		}
	}
}
AddTail(sOut);


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

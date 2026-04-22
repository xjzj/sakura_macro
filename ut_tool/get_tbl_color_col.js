
var Cmd;
//-----------------------------------------------------------
function inpath(file)
{
	var ForReading = 1, ForWriting = 2;
	var wsh = new ActiveXObject("wscript.shell");
	var env=wsh.Environment("SYSTEM");
	var path=env.item("SAKURA_SCRIPT") + "/";
	
	var FileOpener = new ActiveXObject( "Scripting.FileSystemObject");
	var FilePointer = FileOpener.OpenTextFile(path + file, ForReading, true);
	Cmd = FilePointer.ReadAll();
}

//-----------------------------------------------------------
inpath("inc.js");
eval(Cmd);
var fso=new ActiveXObject("Scripting.FileSystemObject");

make_XL();

var  tbl_chk_info = {   col_end:"Z1"  };

var show_tmp="";
var flg=0;
var wk_book=XL.ActiveWorkbook;
var wk_sht=XL.ActiveSheet;

var bookName=wk_book.Name;
var shtName=wk_sht.Name;

var  tbl_addr_list={};
var def_color=convert_RGB(200, 240, 255);
var col_end = wk_sht.Range(tbl_chk_info.col_end).Column;
var max_row=wk_sht.UsedRange.Rows(wk_sht.UsedRange.Rows.Count).Row;
var col=1;
while(col<=col_end){
	var row=1;
	while(row<=max_row){
		var tmp1=wk_sht.Cells(row,col).Value;
		var color=wk_sht.Cells(row,col).Interior.Color;
		if( def_color== color ){
			var color2=wk_sht.Cells(row+1,col).Interior.Color;
			if( def_color== color2 ){
				var jnm=wk_sht.Cells(row,col).Value;
				var enm=wk_sht.Cells(row+1,col).Value;
				var addr=wk_sht.Cells(row, col).Address;
				tbl_addr_list[enm]={ jnm:jnm, addr:addr  };
				break;
			}
		}
		row++;
	}
	col++;
}
for(var tbl in tbl_addr_list ){
	var obj=tbl_addr_list[tbl];
	show_tmp+='"' + tbl + '":"' + obj.addr + '",' + "\n"
}

AddTail( show_tmp + '\n'  );


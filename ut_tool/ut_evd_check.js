
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
var show_txt="";
show_txt+="-----------------------------------------------\r\n";
make_XL();

var chk_data_dup={};
var tbl_folder="C:/work/doc/TBL/txt";
var fmt_folder="C:/work/doc/UT/txt";


var chk_tbl_list={
	"N5_SKISKJ":"$C$22",
	"N5_SKDGISSKS":"$D$22",
	"N5_SKIFKJ":"$E$22",
	"N5_KONYU_SHINSEI_JSN":"$F$22",

	"NULL":"----------------"
};

var data_tp={  'TPC_T_0487_C04':'購入申請情報',  'UT_購入申請情報_正常':'購入申請情報',  'ut_検収情報_data':'検収情報' };

/*
var tbl_nm_list={
	N5_BUIBYK            : "物品要求"         ,
	N5_KONYU_SHINSEI_JSN : "購入申請情報受信" ,
	NULL                 : ""
}
*/

var path_list = [];
search_folder(path_list, tbl_folder, "", "txt");
/*
var all_tbl_list={};
var fobj=new file_obj();
for( var i in path_list){
	var obj=path_list[i];
	var subdir=obj.sub;
	var fname=obj.nm;
	
	var fnm_arr=fname.split('=');
	var tbl_nm=fnm_arr[1].slice(0,-4);
	if(!all_tbl_list[tbl_nm]){
		all_tbl_list[tbl_nm]={  order:[], pty:{}, jnm:tbl_nm_list[tbl_nm]  };
	}
	var order_list=all_tbl_list[tbl_nm].order;
	var pty_map=all_tbl_list[tbl_nm].pty;
	
	show_txt += "tbl_nm=" + tbl_nm + "\n";
	
	var chk=10;
	if(chk){
		var lines=fobj.readFile("utf-8", tbl_folder + '/' + fname );
		var idx=0;
		while(idx < lines.length){
			var line=lines[idx];
			var arr=line.split("\t");
			
			var num=arr[0];
			var enm=arr[1];
			var jnm=arr[2];
			var tp=arr[3];
			var len=arr[4];
			var nnl=arr[5];
			order_list.push(enm);
			pty_map[enm]={  num:num, jnm:jnm, tp:tp, len:len, nnl:nnl };
			idx++;
		}
	}
}
*/

var ctrl_info={   folder:"D:/svn/optiplex5040/scene/sinagawa_seaside2/tbl_tool", book:"仕様分析.xlsx" , sht:"tbl_check" };

var ctrl_book=open_excel(ctrl_info.folder, ctrl_info.book);
var ctrl_sht=ctrl_book.Worksheets(ctrl_info.sht);

var max_setting_row=ctrl_sht.UsedRange.Rows(ctrl_sht.UsedRange.Rows.Count).Row;

var tbl_itm_list={};
var row_flg=0;
var row=1;
while(row<=max_setting_row){
	var tmp=ctrl_sht.Cells(row,1).Value;
	var str=tmp?tmp.toString().replace(/(^\s*)|(\s*$)/g,""):"";
	if(str){
		row_flg=1;
		var color=ctrl_sht.Cells(row,1).Interior.Color;
		tbl_itm_list[str]={ color:color};
	}
	else{
		if(row_flg==1){
			break;
		}
	}
	row++;
}

var wk_sht=XL.ActiveSheet;

var sel=XL.Selection;
var sel_col=sel.Column;
var sel_row=sel.Row;


var max_wk_col=wk_sht.UsedRange.Columns(wk_sht.UsedRange.Columns.Count).Column;


var tmp=wk_sht.Cells(sel_row,sel_col).Value;
var tbl_nm=tmp?tmp.toString().replace(/(^\s*)|(\s*$)/g,""):"";

if ( tbl_nm.length > 27 ){
	var csv_prefix=tbl_nm.slice(0,14);
	if( csv_prefix=='TPC_T_0487_C04' ){
		tbl_nm=csv_prefix;
	}
}


//-----------------------------------------------------------
var chk_col=1;
var chk_addr=chk_tbl_list[tbl_nm];
if(chk_addr){
	chk_col= ctrl_sht.Range(chk_addr).Column;
}

var tbl_itm_list=[];
var row_flg=0;
var row=1;
while(row<=max_setting_row){
	var tmp=ctrl_sht.Cells(row,chk_col).Value;
	var str=tmp?tmp.toString().replace(/(^\s*)|(\s*$)/g,""):"";
	if(str){
		row_flg=1;
		var color=ctrl_sht.Cells(row,chk_col).Interior.Color
		tbl_itm_list[str]={  color:color };
	}
	else{
		if(row_flg==1){
			break;
		}
	}
	row++;
}
//-----------------------------------------------------------
var tbl_obj=null;
var pty_map=null;
if(tbl_nm){
	show_txt+='●' + tbl_nm + "\n";
	tbl_obj=read_data_fmt(path_list,tbl_nm)
	//  tbl_obj=all_tbl_list[tbl_nm];
	//  if(tbl_obj){
	//  	pty_map=tbl_obj.pty;
	//  }
}
if(tbl_obj){
	wk_sht.Cells(sel_row-1, sel_col)=tbl_obj.jnm;
	wk_sht.Cells(sel_row-1, sel_col).Interior.Color=convert_RGB(200,240, 255);
	wk_sht.Cells(sel_row, sel_col).Interior.Color=convert_RGB(200,240, 255);
	wk_sht.Cells(sel_row-1, sel_col).Font.Bold=true;
	wk_sht.Cells(sel_row, sel_col).Font.Bold=true;
	pty_map=tbl_obj.pty;
}
var line_cnt=1;
var tmp_cnt=inputBox_SC('行数', 'chek line', line_cnt);
line_cnt=Number(tmp_cnt);

show_txt += 'line_cnt:' + line_cnt + "\n";

var flg_start=0;
var col=sel_col+1;

var  tbl_val_list="";

while(col<=max_wk_col){
	var tmp=wk_sht.Cells(sel_row, col).Value;
	var str=tmp?tmp.toString().replace(/(^\s*)|(\s*$)/g,""):"";
	if(str){
		var flg=0;
		var chk_obj=tbl_itm_list[str];
		if(chk_obj){
			wk_sht.Cells(sel_row, col).Interior.Color=convert_RGB(80, 220, 137);
			var chk_idx=0;
			while(chk_idx<line_cnt){
				wk_sht.Cells(sel_row+1+chk_idx,col).Interior.Color=chk_obj.color;
				chk_idx++;
			}
			flg=1;
		}
		if(pty_map){
			var obj=pty_map[str];
			if(obj){
				wk_sht.Cells(sel_row-1, col)=obj.jnm;
				if(flg){
					wk_sht.Cells(sel_row-1, col).Font.Bold=true;
					wk_sht.Cells(sel_row-1, col).Interior.Color=convert_RGB(210,255,255);
				}
				else{
					wk_sht.Cells(sel_row-1, col).Interior.Color=convert_RGB(240, 255,255);
				}
				if( obj.nnl == '●'){
					wk_sht.Cells(sel_row, col).Font.Bold=true;
					wk_sht.Cells(sel_row, col).Font.Color=convert_RGB(200,0,0);
				}
				var cell=wk_sht.Cells(sel_row,col);
				if(cell.Comment){
					cell.Comment.Delete();
				}
				cell.AddComment();
				cell.Comment.Visible=false;
				cell.Comment.Text( obj.tp + '(' + obj.len + ')' );
			}
		}
		flg_start=1;
	}
	else{
		if(flg_start){
			break;
		}
	}
	col++;
}

show_txt += "end";
AddTail(show_txt);


function  read_data_fmt(_path_list, _tbl_nm ) {
	
	var tbl_jnm="";
	var fname="";
	var file_tp=0;
	var file_folder="";
	var tmp_dat=data_tp[_tbl_nm];
	if(tmp_dat){
		tbl_jnm=tmp_dat;
		fname=tbl_jnm+ ".txt";
		file_tp=2;
		file_folder=fmt_folder;
	}
	else{
		for( var i  in  _path_list){
			var obj=_path_list[i];
			var subdir=obj.sub;
			var _fname=obj.nm;
			
			var fnm_arr=_fname.split('=');
			var tbl_nm=fnm_arr[1].slice(0,-4);
			var _tbl_jnm=fnm_arr[0];
			if(tbl_nm==_tbl_nm){
				tbl_jnm=_tbl_jnm;
				fname=_fname;
				file_tp=1;
				file_folder=tbl_folder;
				break;
			}
		}
	}
	if(file_tp==0){
		return null;
	}
	
	var all_tbl={ pty:{}, jnm:tbl_jnm  };
	
	var pty_map=all_tbl.pty;
	var itm_pty={};
	itm_pty[1]={  num:0, enm:1, jnm:2, tp:3, len:4, nnl:5 };
	itm_pty[2]={  num:0, enm:3, jnm:1, tp:4, len:5, nnl:7 };
	
	
	var fobj= new file_obj();
	
	var lines=fobj.readFile("utf-8", file_folder + '/' + fname );
	var idx=0;
	if(file_tp>1){
		idx=1;
	}
	AddTail("file_tp:"+ file_tp);
	
	var pty_obj=itm_pty[file_tp];
	while(idx<lines.length){
		var line=lines[idx];
		var arr=line.split("\t");
		var num=arr[pty_obj.num];
		var enm=arr[pty_obj.enm];
		var jnm=arr[pty_obj.jnm];
		var tp=arr[pty_obj.tp];
		var len=arr[pty_obj.len];
		var nnl=arr[pty_obj.nnl];
		pty_map[enm]={ num:num, jnm:jnm, tp:tp, len:len, nnl:nnl };
		idx++;
	}
	return all_tbl;
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





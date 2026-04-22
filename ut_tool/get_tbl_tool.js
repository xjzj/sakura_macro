
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

var nnl_flg_list={
	"〇":"●",
	"○":"●",
	"":""
}

//------------------------------------------------------------
var fso=new ActiveXObject("Scripting.FileSystemObject");
var mk_tbl_head_list={ "N5":1, "CM":1 };
var not_mk_tbl_tail_list={ "SEQUENCE":1 };

var tbl_folder="C:/work/doc/TBL/txt";

var tbl_info = {  num:"A9", enm:"C9", jnm:"R9", tp:"AI9", len:"AO9", nnl:"AS9"  };


var debug="";
var tbl_list="";
var show_txt="";
make_XL();

var wk_book=XL.ActiveWorkbook;
var wk_sht=XL.ActiveSheet;

var bookName=wk_book.Name;
var shtName=wk_sht.Name;

if(  bookName != "仕様分析.xlsx"  ){
	throw new Error("exit");
}
show_txt=bookName + "\n";
fso_check_create_path(tbl_folder);
if( shtName == 'all' ) {
	var st_cnt=wk_book.Worksheets.Count;
	var idx=1;
	while(idx<=st_cnt){
		var _sht=wk_book.Worksheets(idx);
		var shtNm=_sht.Name;
		var arr_nm=shtNm.split('_');
		var head=arr_nm[0];
		var tail=arr_nm[1];
		if(mk_tbl_head_list[head] && (!not_mk_tbl_tail_list[tail]) ){
			mk_tbl(_sht);
		}
		else{
			show_txt+='●' + shtNm + "\n";
		}
		idx++;
	}
}
else{
	mk_tbl(wk_sht);
}
show_txt+='end' + "\n";

var asc_cd=32;
var str_list="";
while(asc_cd<126){
	str_list+=String.fromCharCode(asc_cd);
	asc_cd++;
}
AddTail(show_txt);
AddTail(str_list);


function mk_tbl(_wk_sht){
	show_txt+=_wk_sht.Name + "\n";
	var tbl_jnm=_wk_sht.Range("K5").Value;
	var tbl_enm=_wk_sht.Range("BB5").Value;
	
	var col_num=_wk_sht.Range(tbl_info.num).Column;
	var col_enm=_wk_sht.Range(tbl_info.enm).Column;
	var col_jnm=_wk_sht.Range(tbl_info.jnm).Column;
	var col_tp=_wk_sht.Range(tbl_info.tp).Column;
	var col_len=_wk_sht.Range(tbl_info.len).Column;
	var col_nnl=_wk_sht.Range(tbl_info.nnl).Column;
	
	
	var row_num=_wk_sht.Range(tbl_info.num).row;
	var row=row_num;
	var tbl_list=[];
	debug+= "_wk_sht.UsedRange.Rows.Count:" + _wk_sht.UsedRange.Rows.Count + "\n";
	var max_row=2+_wk_sht.UsedRange.Rows.Count;
	for( ;row<=max_row; row++){
		

		var num=_wk_sht.Cells(row,col_num).Value;
		var enm=_wk_sht.Cells(row,col_enm).Value;

		var jnm=_wk_sht.Cells(row,col_jnm).Value;
		
		var tp=_wk_sht.Cells(row,col_tp).Value;
		var len=_wk_sht.Cells(row,col_len).Value;
		var tmp_nnl=_wk_sht.Cells(row,col_nnl).Value;
		var nnl=tmp_nnl?tmp_nnl.toString().replace(/(^\s*)|(\s*$)/g,""):"";
		
		debug+=enm+":" + jnm + "\n";
		if(enm){
			var line=num+ "\t" + enm + "\t" + jnm + "\t" + tp + "\t" + len + "\t" + nnl_flg_list[nnl];
			tbl_list.push(line);
		}
		else{
			break;
		}

		
	}
	var file_name=tbl_jnm + "=" + tbl_enm + ".txt";
	var fobj= new file_obj();
	fobj.writeFile("utf-8", tbl_folder + '/' + file_name, tbl_list);
}



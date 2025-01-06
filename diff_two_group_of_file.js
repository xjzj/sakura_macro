
var Cmd;
//-----------------------------------------------------------
function inpath(file) 
{
	var ForReading = 1, ForWriting = 2;
	var wsh = new ActiveXObject("wscript.shell");
	var env = wsh.Environment("SYSTEM");
	var path = env.item("SAKURA_SCRIPT") + "/";

	var FileOpener = new ActiveXObject("Scripting.FileSystemObject");
	var FilePointer = FileOpener.OpenTextFile(path + file, ForReading, true);
	Cmd = FilePointer.ReadAll();
}

//-----------------------------------------------------------
//-----------------------------------------------------------
inpath("inc.js");
eval(Cmd);

d = new Date();
sdt= d.getYear() +  ( "0" + (d.getMonth() + 1)).slice(-2)  +  ( "0" + d.getDate()).slice(-2);  


var  empty_line_reg = /^(\s*)$/;
var  valid_line_reg = /^\s*(\S+|\S.*\S)\s*$/;



// sOut += "apk fetch "
var  pid = 0;
var cmd_list = new Array();
cmd_list [ pid++ ] = { func:check_diff , pattern:empty_line_reg , att:1 } ;
cmd_list [ pid++ ] = { func:check_diff , pattern:valid_line_reg , att:2 } ;

var  test_tmp = "------------------------------\r\n";
// var   funcs=select_cmd(cmd_list);

// sOut += "apk fetch ";
var  run_flg=0;
do_select_cmd(cmd_list);
// run_flg=1;
// do_select_cmd(cmd_list);
finish();

// AddTail(test_tmp);

ReDraw(0);

function check_diff(_sline, _pat, _cnt, att, mat) 
{
	// test_tmp +=  "_sline[" + _sline + "]_pat[" + _pat + "]_cnt[" + _cnt + "]att[" + att + "]mat[" + mat + "]\r\n";
	if ( !('num' in check_diff)) {
		check_diff.valuid_line_cnt = 0;
		check_diff.empty_line_cnt = 0;
		check_diff.num=0;
		check_diff.valuid_lines={};
		check_diff.empty_lines={};
	}
	var line_cnt=_cnt+1;
	check_diff.num++;
	
	if (att == 1) {
		check_diff.valuid_line_cnt = line_cnt+1;
		if(  check_diff.empty_line_cnt>0  ){
			if( !check_diff.empty_lines[check_diff.empty_line_cnt]){
				check_diff.empty_lines[check_diff.empty_line_cnt]={ cnt:1, flg:false};
			}
			else{
				check_diff.empty_lines[check_diff.empty_line_cnt].cnt++;
			}
		}
	}
	else  if (att == 2) {
		if( check_diff.empty_lines[check_diff.empty_line_cnt]){
			check_diff.empty_lines[check_diff.empty_line_cnt].flg=true;
		}
		check_diff.empty_line_cnt = line_cnt+1;
		if( !check_diff.valuid_lines[check_diff.valuid_line_cnt]){
			check_diff.valuid_lines[check_diff.valuid_line_cnt]={};
		}
		var  valuid_lines=check_diff.valuid_lines[check_diff.valuid_line_cnt];
		var  diff_key=mat[1];
		valuid_lines[diff_key]=1;
	}
	else{
		throw "_cnt[" + _cnt +  "]_sline[" + _sline  + "]";   // 扔出一个错误。
	}
	return;
}

function finish() 
{
	var split_num=0;
	var max_empty_cnt=0;
	for ( var i in  check_diff.empty_lines){
		var obj=check_diff.empty_lines[i];
		var num_i=Number(i);
		if( obj.flg && max_empty_cnt< obj.cnt){
			max_empty_cnt=obj.cnt;
			split_num=num_i;
		}
	}
	
	var before={};
	var after={};
	for ( var i in  check_diff.valuid_lines){
		var maps=check_diff.valuid_lines[i];
		if(  i  <=split_num  ){
			// AddTail("before:" +split_num+ ":"  +  i  +"------------\n");
			for( var key in maps ){
				before[key]=i;
				// AddTail(key + "\n");
			}
		}
		else{
			// AddTail("after:" +split_num+ ":" +  i  + "------------\n");
			for( var key in maps ){
				after[key]=i;
				// AddTail(key + "\n");
			}
		}
	}
	
	var com_list=[];
	var before_list=[];
	var after_list=[];
	for( var key in before ){
		if(after[key]){
			com_list.push(key);
		}
		else{
			before_list.push(key);
		}
	}
	for( var key in after ){
		if(!before[key]){
			after_list.push(key);
		}
	}
	AddTail("********(" +  split_num   +    ")**************************\n");
	var  com_txt="";
	var  before_txt="";
	var  after_txt="";
	var idx=0;
	for(  ; idx<com_list.length  ; idx++ ){
		var  line=com_list[idx];
		com_txt+= line+ "\n";
	}
	idx=0;
	for(  ; idx<before_list.length  ; idx++ ){
		var  line=before_list[idx];
		before_txt+= line+ "\n";
	}
	idx=0;
	for(  ; idx<after_list.length  ; idx++ ){
		var  line=after_list[idx];
		after_txt+= line+ "\n";
	}

	AddTail(com_txt);
	AddTail("--" + sdt  + "-before_txt-------------\n");
	AddTail(before_txt);
	AddTail("--" + sdt  + "-after_txt-------------\n");
	AddTail(after_txt);
	AddTail("----------------------------\n");
	return  null;
}








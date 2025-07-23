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
inpath("inc.js");
eval(Cmd);


var rep_cd_list=['@','?','~','%','^','`','$','&','#','!','<','>'];
function code_replace(str,ch,map,all_str){
	var chr=map.a[ch];
	if( chr){
		var nstr = str.split(ch).join(chr);
		return nstr ;
	}

	for(var i in rep_cd_list){
		var  chr=rep_cd_list[i];
		if( (chr!=ch) && (!map.b[chr])){
			var p=all_str.indexOf(chr);
			if(p==-1){
				map.a[ch]=chr;
				map.b[chr]=ch;
				var nstr = str.split(ch).join(chr);
				return nstr ;
			}
		}
	}
	return 'ali char not useable!!!!!!!!';
}
var com_code_list={};

function  all_code_replace(from_str,_cnt, _code_list,change_list, _all_str){
	if(!_code_list[_cnt]){
		_code_list[_cnt]={ a:{}, b:{}};
	}
	var i=0;
	var to_str="";
	while(i< change_list.length){
		var ch=change_list[i];
		to_str=code_replace(from_str,ch,_code_list[_cnt], _all_str);
		from_str=to_str
		i++;
	}
	return to_str;
}

function  all_code_restore(from_str,_cnt, _code_list){
	var to_str=from_str;
	if( _code_list[_cnt] ) {
		var char_map=_code_list[_cnt].b;
		for( var  chr in char_map){
			var ch=char_map[chr];
			to_str = from_str.split(chr).join(ch);
			from_str=to_str
		}
	}
	return to_str;
}


function  com_code_bakup( str,cnt  )
{
	var new_line=str;
	var mt=new_line.match(/(\/\*\s+)(\S.*\S)(\s+\*\/)/ );
	if(mt){
		var left=RegExp.leftContext;
		var right=RegExp.rightContext;
		var cmt=mt[2];
		var ncmt=all_code_replace(cmt,cnt,com_code_list,[' ',';', ':', '(', ')' ],str);
		new_line=left+mt[1] + ncmt + mt[3] + right;
	}
	var mt=new_line.match(/(TO_DATE\([^,]+,[^)]+\))(,)/ );
	if(!mt){
		mt=new_line.match(/(LEFT\([^,]+,[^)]+\))(\s*=)/ );
	}
	if(mt){
		var left=RegExp.leftContext;
		var right=RegExp.rightContext;
		var tmp=mt[1];
		var new_str=all_code_replace(tmp,cnt,com_code_list,[' ', '(', ')' ],str);
		new_line=left+ new_str + mt[2] + right;
	}
	
	var mt=new_line.match(/(unsigned|struct)\s+(\w+)/ );
	if(mt){
		var left=RegExp.leftContext;
		var right=RegExp.rightContext;
		var tmp=mt[1] +' ' +mt[2];
		var new_str=all_code_replace(tmp,cnt,com_code_list,[' '],str);
		new_line=left+ new_str + right;
	}

	return new_line;
}
function  com_code_restore( str,cnt  )
{
	var new_line=str;
	new_line= all_code_restore(str, cnt, com_code_list);
	return new_line;
}
var prefix_list=[com_code_bakup];
var tailfix_list=[com_code_restore];


//-----------------------------------------------------------
doAlignmentDefault.name="doAlignmentDefault";
doAlignmentDefine.name="doAlignmentDefine";
doAlignmentMd5.name="doAlignmentMd5";
doAlignmentInit.name="doAlignmentInit";


//sample.js
//空行をカウントする
var debug="";
// var WSHShell = new ActiveXObject("WScript.Shell");

var reg_var_init = /^\s*(VARCHAR|double|char|int|short|\S+)\s+(\S.*\S|\S)\s*;.*$/;

var reg_js = /^\s*cmd_list\[.*$/;
var reg_md5 = /^(\S.*\S)\s+:\s*(\w+)\s*$/;
var reg_define = /^#define\s(\w+)\s+(\S+).*$/;
var reg_default = /^\s*(\S+)\s*/;

var reg_func_arg_list = /^\s*\*[^:]+:[^:]+:[^:]+:[^:]+$/;
var reg_func_arg_list2 = /^\s*\*[^:]+:[^:]+/;

var reg_var_def = /^\s*(char|int)\s+(\S.*\S)\s*;.*$/;

var  pid = 0;
var cmd_list = new Array();
cmd_list[pid++] = { func:doAlignmentDefault , pattern:null         , att:0 };
cmd_list[pid++] = { func:doAlignmentDefine  , pattern:reg_define   , att:0 };
cmd_list[pid++] = { func:doAlignmentMd5     , pattern:reg_md5      , att:0 };
cmd_list[pid++] = { func:doAlignmentDefault , pattern:reg_js       , att:0 };
cmd_list[pid++] = { func:doAlignmentInit    , pattern:reg_var_init , att:0 };

cmd_list[pid++] = { func:doAlignmentFuncArg    , pattern:reg_func_arg_list , att:0 };
cmd_list[pid++] = { func:doAlignmentFuncArg    , pattern:reg_func_arg_list2 , att:0 };

cmd_list[pid++] = { func:doAlignmentDef    , pattern:reg_var_def , att:0 };

do_select_cmd(cmd_list, 1);
// AddTail(debug);
// run_debug();
function  run_debug(){

	var vlist=judge_cmd.vlist;
	var idx_v=0;
	while(idx_v<vlist.length){
		var func_obj=vlist[idx_v];
		debug+= func_obj.func.name + ":" + func_obj.cnt + "\n";
		var pat_list=func_obj.pat_list;
		var idx_pat=0;
		while(idx_pat<pat_list.length){
			var pat_obj=pat_list[idx_pat];
			debug+= "\tpattern:[" + pat_obj.pat + "]att:[" + pat_obj.att + "]cnt:[" + pat_obj.cnt +"]" + "]line:[" + pat_obj.line.join(',') +"]" + "\n";
			idx_pat++;
		}
		idx_v++;
	
	}

	var idx=0;
	list=align.arr_line;
	debug+=list.length + "\n";
	while( idx<list.length){
		var obj=list[idx];
		if(obj && obj.gdb ){
			debug+= "idx:" + idx + ":" + obj.gdb.reg + "\n";
		}
		idx++;
	}
	AddTail(debug);
}



function doAlignmentFuncArg() {
	debug += "doAlignmentFuncArg" + "\r\n"
	// delimiter,  front_con, front_split, behind_con,  behind_split, section1, section2
	var strtok =[ " \t", ",;[]+", '*:', "*", "=:", "\"", ""  ];
	var reg = /^\s*\*/;
	rule=[{ tok:strtok, reg:reg, cflg:2, gflg:0 }];
	
	reg_list1=[];
	reg_list1.push([/^(\s*\*)\s(\S)/,"$1$2"]);
	reg_list2=[];
	reg_list2.push([/^(\s*\*)/,"$1 "]);
	doAlignmentRule( rule , reg_list1 ,reg_list2);

}

function doAlignmentDef() {

	debug += "doAlignmentInit" + "\r\n"
	// delimiter,  front_con, front_split, behind_con,  behind_split, section1, section2
	//var strtok1 = [ " \t", ";+", '=', "+", "=", "\"", ""  ];
	 var strtok1 = [ " \t", "+", '=;', "+", ";=", "\"", ""  ];
	// var reg1 = /^\s*(char|int)\s+(\S.*\S)\s*;.*$/;
	var reg1 = /^\s*(\S+)\s*/;

	// var strtok2 = [ " \t", ";+", '=', "+", "=", "\"", ""  ];
	var reg2 =  /^\s*(int)\s+(\S.*\S)\s*;.*$/;
	rule=[];
	rule[0] = {tok:strtok1, reg:reg1, cflg:2, gflg:0 };
	// rule[1] = {tok:strtok2, reg:reg2, cflg:2, gflg:0 };

	doAlignmentRule( rule , prefix_list ,tailfix_list);

}

function doAlignmentInit() {
	// msg_box( "doAlignmentInit" )
	debug += "doAlignmentInit" + "\r\n"
	// delimiter,  front_con, front_split, behind_con,  behind_split, section1, section2
	//var strtok1 = [ " \t", ";+", '=', "+", "=", "\"", ""  ];
	 var strtok1 = [ " \t", "+", '=;', "+", ";=", "\"", ""  ];
	// var reg1 = /^\s*(char|int)\s+(\S.*\S)\s*;.*$/;
	var reg1 = reg_var_init;

	// var strtok2 = [ " \t", ";+", '=', "+", "=", "\"", ""  ];
	var reg2 =  /^\s*(int)\s+(\S.*\S)\s*;.*$/;
	rule=[];
	rule[0] = {tok:strtok1, reg:reg1, cflg:2, gflg:0 };
	// rule[1] = {tok:strtok2, reg:reg2, cflg:2, gflg:0 };

	doAlignmentRule( rule , prefix_list ,tailfix_list);
}



function doAlignmentDefine() {
	// msg_box( "doAlignmentDefine" )
	debug += "doAlignmentDefine" + "\r\n"
	// delimiter,  front_con, front_split, behind_con,  behind_split, section1, section2
	var strtok = [ " \t", "", '=', "", "=", "\"", ""  ];
	
	var code_list={}; 
	function  def_replace1( str,cnt  )
	{
		var new_line=str;
		var mt=str.match(/(\/\*\s*)(\S.*\S)(\s*\*\/)/ );
		if(mt){
			var cmt=mt[2];
			if(!code_list[cnt]){
				code_list[cnt]={};
			}
			var ncmt=code_replace(cmt,' ',code_list[cnt]);
			ncmt=code_replace(ncmt,'(',code_list[cnt]);
			ncmt=code_replace(ncmt,')',code_list[cnt]);
			var arr=str.split('/*');
			new_line=arr[0]+mt[1] + ncmt + mt[3];
		}
		return new_line;
	}
	function  def_replace2( str,cnt  )
	{
		var new_line=str;
		var mt=str.match(/(\/\*\s*)(\S.*\S)(\s*\*\/)/ );
		if(mt){
			var cmt=mt[2];
			var char_map=code_list[cnt];
			for( var  chr in char_map){
				var ch=char_map[chr];
				cmt = cmt.split(chr).join(ch);
			}
			var arr=str.split('/*');
			new_line=arr[0]+mt[1] + cmt + mt[3];
		}
		return new_line;
	}
		
	var rlist1=[def_replace1];
	var rlist2=[def_replace2];

	rule=[ {tok:strtok,reg:reg_define,  cflg:2, gflg:0  }];
	doAlignmentRule(rule,rlist1,rlist2);
	// doAlignmentRule(rule,rlist1,null);
	// doAlignmentRule(rule);
}

function doAlignmentDefault() {
	debug += "doAlignmentDefault" + "\r\n"
	// msg_box( "doAlignmentDefault" )
	// delimiter,  front_con, front_split, behind_con,  behind_split, section1, section2
	// var strtok = [ " \t", "", '=,;[]', "", "=,[];", "\"", ""  ];
	var strtok = [ " \t", ",", '=;[],', "", ",[];", "\"", ""  ];
	rule=[{tok:strtok,reg:reg_default,  cflg:2, gflg:0  }];
	
	var reg_list1=[];
	reg_list1.push([/const\s+char/,"const-char"]);
	var reg_list2=[];
	reg_list2.push([/const\-char/,"const char"]);
	
	
	doAlignmentRule(rule,reg_list1,reg_list2);
}

function doAlignmentMd5() {
	debug += "doAlignmentMd5" + "\r\n"
	// msg_box( "doAlignmentMd5" )
	// delimiter,  front_con, front_split, behind_con,  behind_split, section1, section2
	var strtok = [ " \t", "", '=', "", "=", "\"", ""  ];
	rule=[{tok:strtok,reg:reg_md5,  cflg:2, gflg:0  }];
	doAlignmentRule(rule);
}

function doAlignmentRule( rule_, _rlst1, _rlst2) {
	if (IsTextSelected) {
		AlignmentRule_(rule_, lines.length, outputEditor, _rlst1, _rlst2);

	} else {
		var iMax = GetLineCount(0);
		AlignmentRule_(rule_,  iMax, outputEditor, _rlst1, _rlst2);
	}
}

// doAlignment( nMaxLineCnt, outputEditor);
//-----------------------------------------------------------------------
function AlignmentRule_(rule_,  CountMax, fucWrite, _rlst1, _rlst2) {
	var iCnt = 0;
	var al = align(rule_);
	while (iCnt < CountMax) { //全行をループ
		var sline = GetLine(iCnt);
		if(_rlst1){
			sline=r_change(sline,_rlst1,iCnt);
		}
		al.check(sline);
		iCnt++;
	}
	iCnt = 0;
	while (iCnt < al.getsize()) {
		var line= al.getnext();
		if(_rlst2){
			line=r_change(line,_rlst2,iCnt);
		}
		sOut += line;
		iCnt++;
		if (iCnt < al.getsize()) {
			sOut += "\n";
		}
	}

	if (iCnt > 0) {
		fucWrite(sOut);
	}
}

function  r_change(_line,_rlst, _cnt){
	var idx=0;
	while(idx<_rlst.length){
		var obj=_rlst[idx];
		if(typeof(obj) == 'object' ){
			var tmp=_line.replace(obj[0],obj[1]);
			_line=tmp;
		}
		else if(typeof(obj) == 'function'){
			var tmp=obj(_line,_cnt);
			_line=tmp;
		}
		idx++;
	}
	return _line
}






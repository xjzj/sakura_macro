
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


ReDraw(0);
var reg1=new RegExp("^\\s*(\\S+\\s+\\S+).*\\s*$");
var reg2=new RegExp("^(\\w:|\\w\\w+)\\\\\\w+.*$");
var reg99=new RegExp("^\\s*.*\\S+(\\s*)$");


var cmd_list = new Array();

var pid=0;

cmd_list[pid++]={ func:make_tab_between_word ,  pattern:reg1,  att:1  };
cmd_list[pid++]={ func:back_slash_to_slash ,  pattern:reg2,  att:2     };
cmd_list[pid++]={ func:delete_space_crlf ,  pattern:reg99, att:3     };


do_select_cmd(  cmd_list );


function make_tab_between_word(_sline, _pat, _cnt, att, mat )
{
	// msg_box( att )
	var  sTmp="";
	var flg=0;
	var arr=_sline.split("");
	var cnt=0;
	var idx = 0;
	while(idx<arr.length){
		if(arr[idx]== " " || arr[idx] == "\t" ){
			if ( sTmp != "" ) {
				flg++;
			}
		}
		else{
			if ( flg>0){
				//  if ( sTmp != "$CSH" && sTmp != "${CSH}"
				//  			&& sTmp != "/bin/sh" && sTmp != "/bin/csh" && sTmp != "/bin/csh -f" ){
				//  	sTmp += "\t";
				//  }
				//  else{
				//  	sTmp += " ";
				//  	
				//  }
				sTmp += "\t";
			}
			sTmp += arr[idx];
			flg = 0;
		}
		++idx;
	}
	sOut += sTmp + "\n";
}


function back_slash_to_slash( _sline, _pat, _cnt, att, mat )
{
	// msg_box( att )
	var item = _sline.match(_pat);
	if ( item != null ){
		sOut += _sline.split("\\").join('/') + "\n";
	}
	else{
		sOut += _sline + "\n";
	}
}

function delete_space_crlf( _sline, _pat, _cnt, att, mat )
{
	// msg_box( att )
	var item = _sline.match(/^\s*(\S.*\S|\S)\s*$/);
	if ( item != null ){
		sOut += item[1].split("\\").join('/') + "\n";
	}
	else{
		var item = _sline.match(/^(\s*)$/);
		if ( item != null ){
			;;
		}
		else{
			sOut += _sline.split("\\").join('/') + "\n";
		}
	}
}



function delete_end_space( _sline, _pat, _cnt, att, mat )
{
	// msg_box( att )
	var item = _sline.match(/^(.*\S)\s*$/);
	if ( item != null ){
		sOut += item[1] + "\n";
	}
	else{
		var item = _sline.match(/^(\s*)$/);
		if ( item != null ){
			;;
		}
		else{
			sOut += _sline + "\n";
		}
	}
}
































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

//-----------------------------------------------------------
//sample.js
//空行をカウントする

// var WSHShell = new ActiveXObject("WScript.Shell");

var reg_var_init = /^\s*(char|int)\s+(\S.*\S)\s*;.*$/;

var reg_js = /^\s*cmd_list\[.*$/;
var reg_md5 = /^(\S.*\S)\s+:\s*(\w+)\s*$/;
var reg_define = /^#define\s(\w+)\t+(\S+)s*$/;
var reg_default = /^\s*(\S+)\s*/;
var  pid = 0;
var cmd_list = new Array();
cmd_list[pid++] = { func:doAlignmentDefault , pattern:null         , att:0 };
cmd_list[pid++] = { func:doAlignmentDefine  , pattern:reg_define   , att:0 };
cmd_list[pid++] = { func:doAlignmentMd5     , pattern:reg_md5      , att:0 };
cmd_list[pid++] = { func:doAlignmentDefault , pattern:reg_js       , att:0 };
cmd_list[pid++] = { func:doAlignmentInit    , pattern:reg_var_init , att:0 };

do_select_cmd(cmd_list, 1);

function doAlignmentInit() {
	// msg_box( "doAlignmentInit" )
	// delimiter,  front_con, front_split, behind_con,  behind_split, section1, section2
	var strtok1 = [ " \t", ";+", '=', "+", "=", "\"", ""  ];
	var reg1 = /^\s*(char)\s+(\S.*\S)\s*;.*$/;

	var strtok2 = [ " \t", ";+", '=', "+", "=", "\"", ""  ];
	var reg2 =  /^\s*(int)\s+(\S.*\S)\s*;.*$/;
	rule=[];
	rule[0] = {tok:strtok1, reg:reg1, cflg:1, gflg:1 };
	rule[1] = {tok:strtok2, reg:reg2, cflg:1, gflg:1 };

	doAlignmentRule( rule);
}

function doAlignmentDefine() {
	// msg_box( "doAlignmentDefine" )
	// delimiter,  front_con, front_split, behind_con,  behind_split, section1, section2
	var strtok = [ "\t", "", '=', "", "=", "\"", ""  ];
	rule=[ {tok:strtok,reg:reg_define,  cflg:1, gflg:1  }];
	doAlignmentRule(rule);
}

function doAlignmentDefault() {
	// msg_box( "doAlignmentDefault" )
	// delimiter,  front_con, front_split, behind_con,  behind_split, section1, section2
	// var strtok = [ " \t", "", '=,;[]', "", "=,[];", "\"", ""  ];
	var strtok = [ " \t", ",", '=;[]', "", ",[];", "\"", ""  ];
	rule=[{tok:strtok,reg:reg_default,  cflg:2, gflg:0  }];
	doAlignmentRule(rule);
}

function doAlignmentMd5() {
	// msg_box( "doAlignmentMd5" )
	// delimiter,  front_con, front_split, behind_con,  behind_split, section1, section2
	var strtok = [ " \t", "", '=', "", "=", "\"", ""  ];
	rule=[{tok:strtok,reg:reg_md5,  cflg:2, gflg:0  }];
	doAlignmentRule(rule);
}

function doAlignmentRule( rule_) {
	if (IsTextSelected) {
		AlignmentRule_(rule_, lines.length, outputEditor);

	} else {
		var iMax = GetLineCount(0);
		AlignmentRule_(rule_,  iMax, outputEditor);
	}
}

// doAlignment( nMaxLineCnt, outputEditor);
//-----------------------------------------------------------------------
function AlignmentRule_(rule_,  CountMax, fucWrite) {
	var iCnt = 0;
	var al = align(rule_);
	while (iCnt < CountMax) { //全行をループ
		var sline = GetLine(iCnt);
		al.check(sline);
		iCnt++;
	}
	iCnt = 0;
	while (iCnt < al.getsize()) {
		sOut += al.getnext();
		iCnt++;
		if (iCnt < al.getsize()) {
			sOut += "\n";
		}
	}

	if (iCnt > 0) {
		fucWrite(sOut);
	}
}




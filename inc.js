
var gCmd;

function getEnvVar(field) 
{
	var wsh = new ActiveXObject("WScript.shell");
	// var	val = WSHShell.ExpandEnvironmentStrings( "%" + field + "%" );
	var env = wsh.Environment("SYSTEM");
	var val = env.item(field);
	return val;
}

function runCommand(cmd) 
{
	var wsh = new ActiveXObject("WScript.shell");
	var oe = wsh.Exec(cmd);
	// var path= env.ExpandEnvironmentStrings ("%SAKURA_SCRIPT%｡ｯ) ;
	var r = oe.StdOut.ReadAll();
	return r;
}
function inc(js_file) 
{
	var ForReading = 1, ForWriting = 2;
	var FileOpener = new ActiveXObject("Scripting.FileSystemObject");
	var FilePointer = FileOpener.OpenTextFile(getEnvVar('SAKURA_SCRIPT') + "/" + js_file, ForReading, true);
	gCmd = FilePointer.ReadAll();
}

function msgbox(msg,title) 
{
	if(!title) {
		title="修正"
	}
	if(typeof(msg) == 'object' ){
		msg=msg.join(':');
	}
	var WSHShell = new ActiveXObject("WScript.Shell");
	WSHShell. Popup(msg, 0, title, 1);
}
function msg_box(msg) 
{
	msgbox(msg);
}

function isword(ch) {
	var flg = false;
	if (ch >= "0" && ch <= "9") {
		flg = true;

	}

	if (ch >= "a" && ch <= "z") {
		flg = true;

	}

	if (ch >= "A" && ch <= "Z") {
		flg = true;

	}

	if (ch == "_") {
		flg = true;

	}
	return flg;
}

function repeat(str, n) {
	if (typeof(n) == 'string') {
		n = Number(n);
	}
	if (n < 0) {
		n = 0;
	}
	var arr = new Array(n + 1);
	return arr.join(str); // "" + str + "" + str + ""  + str + "" ....
}

String.prototype.trim = function() {
	return this.replace(/(^\s*)|(\s*$)/g, "");
}

function get_basedir(path) {
	var npath = path.replace(/\\/g, "/");
	var foldersAndFile = npath.split("/");
	var folders = foldersAndFile.slice(0, foldersAndFile.length - 1);
	var folerpath = folders.join("\\");

	// msg_box("folerpath["+folerpath+"]")
	if (folerpath == "") {
		folerpath = ".";
	}
	return (folerpath+"\\");
}

var _tab_len = 4
function getTabStrLen(str) {
	var arr = str.split("");
	var idx = 0;
	var len = 0;
	while (idx < arr.length) {

		if (arr[idx] < "~") {
			if (arr[idx] == "\t") {
				len += _tab_len - (len % _tab_len);
			} else {
				++len;
			}
		} else {
			len += 2;
		}
		++idx;
	}
	return len;
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

function sDate(sep) {
	var d, s = "";
	d = new Date();
	s += getFixedNum(d.getYear(), 4) + sep;
	s += getFixedNum(d.getMonth() + 1, 2) + sep;
	s += getFixedNum(d.getDate(), 2);
	return (s);
}
function getFixedNum(num, len) {
	var sNum = num.toString();
	var ret = repeat("0", (len - sNum.length)) + sNum;
	return ret;
}


var sep_width = 0;

function  strtok(str, delimiter, front_con, front_split, behind_con, behind_split, section1, section2) {
	if (section1 == undefined) {
		section1 = "\"\'";
	}
	if (section2 == undefined) {
		section2 = "#";
	}
	var arr_word = new Array();
	var arr = str.split("");
	var flag = 0;
	var split_flg = 0;
	var next_split_flg = -1;
	var sec = '#';
	var tmp_word = "";
	var idx = 0;
	var idx_word = 0;
	var len = 0;
	while (idx < arr.length) {
		if (flag == 0) {
			var  op = 0
			var  pos = delimiter.indexOf(arr[idx]) 
			if (pos != -1) {
				split_flg = 1;
				if (next_split_flg >-1) {
					split_flg = next_split_flg;
					next_split_flg = -1;
				}
				op = 1;
			}
			else {
				if (next_split_flg >-1) {
					//echo( next_split_flg+ ":" + tmp_word);
					split_flg = next_split_flg;
					next_split_flg = -1;
				}
				pos = front_con.indexOf(arr[idx]) 
				if (pos != -1) {
					split_flg = 0;
					// op=2;
				}
				pos = front_split.indexOf(arr[idx]) 
				if (pos != -1) {
					split_flg = 1;
					// op=2;
				}
				pos = behind_con.indexOf(arr[idx]) 
				if (pos != -1) {
					next_split_flg = 0;
					// op=2;
				}

				pos = behind_split.indexOf(arr[idx]) 
				if (pos != -1) {
					next_split_flg = 1;
					//echo( arr[idx] + ":" + tmp_word);
					// op=2;
				}
				pos = section1.indexOf(arr[idx]) 
				if (pos != -1) {
					flag = 1;
					sec = arr[idx];
				}
				else {
					pos = section2.indexOf(arr[idx]) 
					if (pos != -1) {
						flag = 2;
						sec = arr[idx];
					}
				}
			}
			// echo( "57:" +idx);
			if (op == 0) {
				if (split_flg == 1) {
					if (tmp_word != "") {
						// echo( "60:" +tmp_word);
						arr_word[idx_word] = tmp_word;
						idx_word++;
						tmp_word = "";
					}
					split_flg = 0;
				}
				tmp_word += arr[idx];
			}
		}
		else if (flag == 1) {
			var pos = section1.indexOf(arr[idx]) 
			if (pos != -1 && sec == arr[idx]) {
				flag = 0;
			}
			tmp_word += arr[idx];
		}
		else if (flag == 2) {
			if (arr.length == (idx + 1)) {
				flag = 0;
			}
			tmp_word += arr[idx];
		}
		//echo( "85:" +flag+":" +  idx);
		++idx;
	}
	//echo( "86:" +tmp_word);
	if (tmp_word != "") {
		arr_word[idx_word] = tmp_word;
		idx_word++;
	}
	return arr_word;
}

function align(rule_) {
	if ( !( 'max_line' in align)) {
		align.max_line = 0;
		// align.strtok = null;
		align.cnt = 0;
		align.arr_line = new Array();
		
		align.arr_width = null;
		align.arr_width_list = new Array();
		align.rule = rule_;
		align.num = 0;
	}
	function checkMaxColumn(line) {

		var chk_flg=0
		var get_flg=0
		var _strtok=null;
		for (var i in align.rule) {
			var it = align.rule[i];
			var mt = line.match(it.reg); //
			// msg_box( "line:["   + line +     "]mt="+mt  );
			if (mt) {
				if ( !align.arr_width_list[i] ){
					align.arr_width_list[i]=[];
				}
				align.arr_width=align.arr_width_list[i];
				align.num=i;
				chk_flg = it.cflg;
				get_flg = it.gflg;
				_strtok = it.tok;
				break;
			}
		}

		if (chk_flg == 0) {
			align.arr_line[align. max_line] ={ arr:line, flg:get_flg };
		}
		else {
			var sp = line. match(/^(\s*)(\S.*)$/);
			if (sp != null) {
				if (align. arr_width[0] == undefined) {
					align. arr_width[0] = sp[1];
				}
				line = sp[2];
			}
			// var sline = line. replace(/\s+/g, " ");
			// var arr_ln = sline.split(" ");
			// strtok(str, delimiter,  front_con, front_split, behind_con,  behind_split, section1, section2){
			var  arr_ln = null;
			if (_strtok) {
				var _tok = _strtok
				// msg_box( "_tok:" + _tok);
				arr_ln = strtok(line, _tok[0], _tok[1], _tok[2], _tok[3], _tok[4], _tok[5], _tok[6]);
			}
			else {
				arr_ln = strtok(line, "\t", '.', '=', "", "=", "\"", "");
			}

			// msgbox(arr_ln)
			align.arr_line[align.max_line] ={ arr:arr_ln, flg:get_flg, num:align.num };
			var idx = 0;
			while (idx < arr_ln.length) {
				if (align.arr_width[idx + 1] == undefined || align. arr_width[idx + 1] <= getTabStrLen(arr_ln[idx])) 
				{
					var tlen = getTabStrLen(arr_ln[idx]);
					// if( arr_ln.length < 0 ){			//1111111111111111111111111
					if ((chk_flg == 1) && (arr_ln.length < 20)) { //1111111111111111111111111
						if (tlen % _tab_len != 0) {
							tlen = tlen + (_tab_len - tlen%_tab_len);
						} else {
							tlen = tlen + _tab_len
						}
					}
					else {
						//tlen++;
						tlen += 1;
					}
					align.arr_width[idx + 1] = tlen;
				}
				++idx;
			}
		}
		align.max_line++;
	}

	function getMaxLine() 
	{
		return align.max_line;
	}
	function getColumnText() {
		var arr = align. arr_line[align. cnt].arr;
		var get_flg = align. arr_line[align. cnt].flg;
		var num = align. arr_line[align. cnt].num;
		align.arr_width=align.arr_width_list[num];
		var idx = 0;
		var line = "";

		// msg_box("align. cnt="+ align. cnt  +  "arr:" + arr   );
		if (typeof arr == 'string') {
			line = arr;
		} else {

			while (idx < arr.length) {
				if (idx > 0) {
					var len_ = align.arr_width[idx] + sep_width - getTabStrLen(arr[idx - 1]);
					if (get_flg == 0) {
						line += repeat(" ", len_) + arr[idx];
					}
					else if (get_flg == 1) {
						var tablen_ = parseInt(len_/_tab_len);
						var splen_ = len_%_tab_len;
						if (splen_ > 0) {
							tablen_++;
						}
						line += repeat("\t", tablen_);
						line += arr[idx];
					}
					else {
						line += " ---" + arr[idx];
					}
				}
				else {
					line = align. arr_width[idx] + arr[idx];
				}
				++idx;
			}
		}
		align. cnt++;

		return line;
	}
	return {check:checkMaxColumn, getnext:getColumnText, getsize:getMaxLine};
}

//-------------------------------------------------------------------
function convert_RGB(r, g, b) {
	var color = 0;
	color += r;
	color += g << 8;
	color += b << 16;
	return color;
}
//-------------------------------------------------------------------


var sSeparator = "\n---------------------------------------\n";
var sOut = "";
var nMaxLineCnt = 0;
var lines;

function reset_editor() {
	if (IsTextSelected) {
		var txt = GetSelectedString(0);
		lines = txt.split("\n");
		nMaxLineCnt = lines.length;
	} else {
		sOut += sSeparator;
		nMaxLineCnt = GetLineCount(0);
	}
}

reset_editor();

function GetLine(iCnt) {
	var sline;
	if (IsTextSelected) {
		sline = lines[iCnt].replace(/\r\n|\r|\n$/, "");
	} else {
		sline = GetLineStr(iCnt + 1) .replace(/\r\n|\r|\n$/, "");
	}
	return sline;
}

function outputEditor(sOutPut) {
	if (IsTextSelected) {
		InsText(sOutPut);
	} else {
		sOutPut += "\n" + sSeparator;
		AddTail(sOutPut);
	}
}

function dojob0(dolist, funcWrite) {
	var iCnt = 0;
	while (iCnt < nMaxLineCnt) {
		var sline = GetLine(iCnt);
		dolist.func(sline, dolist.pattern, iCnt);
		iCnt++;
	}
	if (iCnt > 0 && funcWrite) {
		funcWrite(sOut);
	}
}


function dojob(dolist, funcWrite) {
	var iCnt = 0;
	while (iCnt < nMaxLineCnt) {
		var sline = GetLine(iCnt);
		var idx = 0;
		var idx_null = -1;
		var item = false;
		while (idx < dolist.pat_list.length) {
			if(dolist.pat_list[idx].pat){
				item = sline.match(dolist.pat_list[idx].pat) 
				if (item != null) {
					break;
				}
			}
			else{
				idx_null=idx;
			}
			idx++;
		}
		if (idx < dolist.pat_list.length) {
			dolist.func(sline, dolist.pat_list[idx].pat, iCnt, dolist.pat_list[idx].att, item, dolist.cnt);
		}
		else {
			if( idx_null>-1 ){
				// AddTail("idx_null:" + idx_null  + ";_obj.pat:"+ dolist.pat_list[idx_null].pat + ";dolist.pat_list.length:"+ dolist.pat_list.length + "\r\n");
				dolist.func(sline, dolist.pat_list[idx_null].pat, iCnt, dolist.pat_list[idx_null].att, null, dolist.cnt);
			}
			else{
				// AddTail("sline:" + sline  + ";iCnt:"+ iCnt + ";dolist.cnt:"+ dolist.cnt + "\r\n");
				dolist.func(sline, null, iCnt, null, null, dolist.cnt);
			}
		}
		iCnt++;
	}
	if (iCnt > 0 && funcWrite) {
		funcWrite(sOut);
	}
}

function  judge_cmd(_list, _cht) 
{
	if ( !('line_cnt' in judge_cmd)) {
		judge_cmd.line_cnt = 0;
		judge_cmd.max_cnt = _cht;
		judge_cmd.joblist = _list;
		judge_cmd.vlist = [];
	}
	function  count_func(_obj) 
	{
		var ret = true;
		var idx = 0;
		while (idx < judge_cmd.vlist.length) {
			if (_obj.func == judge_cmd.vlist[idx].func) {
				judge_cmd.vlist[idx].cnt++;
				ret = false;
				var ret_pat = true;
				var idx_pat = 0;
				while (idx_pat < judge_cmd.vlist[idx].pat_list.length) {
					if (_obj.pattern == judge_cmd.vlist[idx].pat_list[idx_pat].pat) {
						judge_cmd.vlist[idx].pat_list[idx_pat].cnt++;
						ret_pat = false;
						break;
					}
					idx_pat++;
				}
				// msgbox("ret:" + ret  + ";_obj.pat:"+ _obj.pat+ ";judge_cmd.vlist[idx].pat_list.length:"+ judge_cmd.vlist[idx].pat_list.length)
				if (ret_pat) {
					judge_cmd.vlist[idx].pat_list.push( { pat:_obj.pattern, att:_obj.att, cnt:1 } );
				}
				break;
			}
			++idx;
		}
		if (ret) {
			judge_cmd.vlist.push( { func:_obj.func, pat_list:[ { pat:_obj.pattern, att:_obj.att, cnt:1 } ], cnt:1 } );
		}
	}
	function  get_funclist() 
	{
		var vlist = null;
		// msgbox("judge_cmd.line_cnt:" + judge_cmd.line_cnt  + ";nMaxLineCnt:"+ nMaxLineCnt)
		// if ( judge_cmd.line_cnt== nMaxLineCnt){
		var idx = 0;
		var cnt = 0;
		var max_idx = -1;
		while (idx < judge_cmd.vlist.length) {
			if (judge_cmd.vlist[idx].cnt > cnt) {
				cnt = judge_cmd.vlist[idx].cnt;
				max_idx = idx;
			}
			++idx;
		}
		if (max_idx < 0) {
			// throw new Error("[matching pattern ]" + " not exsit" );
			return null;
		}
		var funcs = judge_cmd.vlist[max_idx];
		funcs.pat_list.length = 0;
		//--------------------------------
		var idx = 0;
		var jblist = judge_cmd.joblist;
		while (idx < jblist.length) {
			if (jblist[idx].func == funcs.func) {
				var flg = true;
				var idx_pat = 0;
				while (idx_pat < funcs.pat_list.length) {
					if (jblist[idx].pattern == funcs.pat_list[idx_pat].pat) {
						flg = false;
						break;
					}
					idx_pat++;
				}
				if (flg) {
					funcs.pat_list.push( { pat:jblist[idx].pattern, att:jblist[idx].att, cnt:0 } );
				}
			}
			++idx;
		}
		//--------------------------------
		// }
		return funcs;
	}
	function  check_line(_sline) 
	{
		judge_cmd.line_cnt++;
		var mflg=false;
		var idx = 0;
		var idx_null = 0;
		while (idx < judge_cmd.joblist.length) {
			if (judge_cmd.joblist[idx]) {
				// msgbox("check_line(_sline:[" + _sline + "]pattern:[" +judge_cmd.joblist[idx].pattern + "]"   );
				if(!judge_cmd.joblist[idx].pattern ){
					// count_func(judge_cmd.joblist[idx]);
					idx_null=idx;
				}
				else{
					// AddTail("check_line(_sline:[" + _sline + "]pattern:[" +judge_cmd.joblist[idx].pattern + "]" + "\r\n"  );
					var item = _sline.match(judge_cmd.joblist[idx].pattern);
					if (item) {
						// AddTail("idx:" + idx + "\r\n");
						count_func(judge_cmd.joblist[idx]);
						mflg=true;
						//  break;
					}
				}
			}
			++idx;
		}
		if(!mflg 
				&& judge_cmd.max_cnt==judge_cmd.line_cnt
				&& judge_cmd.vlist.length==0 ){
			// AddTail("idx_null:" + idx_null+ "\r\n");
			
			count_func(judge_cmd.joblist[idx_null]);
		}
	}
	return {get:get_funclist, check:check_line};
}


function select_cmd(_list) 
{
	var ret_var = null;
	var iCnt = 0;
	var do_judge = judge_cmd(_list);
	while (iCnt < nMaxLineCnt) 
	{
		var sline = GetLine(iCnt);
		do_judge.check(sline);
		iCnt++;
	}

	return do_judge.get();
}

function do_select_cmd(_list, mode) 
{
	if (mode == undefined) {
		mode = 0;
	}
	var iCnt = 0;
	var check_cnt = 50;
	if (nMaxLineCnt < check_cnt) {
		check_cnt = nMaxLineCnt;
	}
	var do_judge = judge_cmd(_list,check_cnt);

	// msgbox("check_cnt:" + check_cnt);
	while (iCnt < check_cnt) 
	{
		var sline = GetLine(iCnt);
		do_judge.check(sline);
		iCnt++;
	}
	var do_vlist = do_judge.get();
	if (do_vlist) {
		// msg_box( "1" )
		if (mode == 0) {
			dojob(do_vlist, outputEditor);
		}
		else {
			do_vlist.func();
		}
	}
	else {
		// msg_box( "2" )
		if (mode != 0) {
			_list[0].func();
		}
	}
}

//-----------------------------------------------
function getpair(str_pair) {
	var tt = null;
	if (tt == null) {
		tt = str_pair.match(/^\s*"([^\"]+)"\s+"([^\"]+)"\s*/);
	}
	if (tt == null) {
		tt = str_pair.match(/^\s*'([^\']+)'\s+'([^\']+)'\s*/);
	}
	if (tt == null) {
		tt = str_pair.match(/^\s*"([^\"]+)"\s+'([^\']+)'\s*/);
	}
	if (tt == null) {
		tt = str_pair.match(/^\s*'([^\']+)'\s+"([^\"]+)"\s*/);
	}
	if (tt == null) {
		tt = str_pair.match(/^\s*(\S+)\s+(\S+)\s*/);
	}
	return tt;
}

function enumFiles(_target, callback, field_name) {
	var fso = new ActiveXObject("Scripting.FileSystemObject");
	var ext = field_name.split('.') .pop() .toLowerCase();
	if (ext == "") {
		return;
	}
	if (typeof(_target) == "string") {
		var path_name = _target + "/" + field_name;
		if (fso.FileExists(path_name)) {
			callback[ext]( { path: path_name, file:field_name, dir:_target } );
			return;
		}

		if ( !fso.FolderExists(_target)) 
			return;
		var dir = fso.GetFolder(_target);
		_enum(dir);
	}
	else {
		for (var idx in _target) {
			var path_name = _target[idx] + "/" + field_name;
			if (fso.FileExists(path_name)) {
				callback[ext]( { path: path_name, file:field_name, dir:_target[idx] } );
				return;
			}

			if ( !fso.FolderExists(_target[idx])) 
				return;
			var dir = fso.GetFolder(_target[idx]);
			var ret = _enum(dir);
			if (ret > 0) {
				break;
			}
		}
	}
	function _enum(dir) {
		var e = new Enumerator(dir.SubFolders);
		for (; !e.atEnd(); e.moveNext()) {
			var sdir = e.item();
			var file_name = "";
			var iLine = 0;

			file_name = field_name;
			var path_name = sdir.Path + "/" + file_name;
			if (fso.FileExists(path_name)) {
				callback[ext]( { path: path_name, file:field_name, dir:sdir.Path } );
				return 1;
			}
			var ret = _enum(sdir);
			if (ret > 0) {
				return ret;
			}
		}
		return 0;
	}
}

function open_file(fileinfo) {
	FileOpen(fileinfo.path);
}
function open_excel_file(fileinfo) {
	open_excel(fileinfo.dir, fileinfo.file);
}
function open_pdf_file(fileinfo) {
	var sap = new ActiveXObject("Shell.Application");
	var file_path=fileinfo.path.replace(/\//g, "\\");
	sap.ShellExecute(file_path);
}

function lock_XL() {
	xlCalculationAutomatic = -4105;
	xlNormalView = 1;
	XL.ScreenUpdating = true;
	XL.EnableEvents = true;
	XL.AskToUpdateLinks = true;
	XL.DisplayAlerts = true;
	XL.Calculation = xlCalculationAutomatic;
}

function unlock_XL() {
	xlCalculationAutomatic = -4105;
	xlNormalView = 1;
	XL.ScreenUpdating = true;
	XL.EnableEvents = true;
	XL.AskToUpdateLinks = true;
	XL.DisplayAlerts = true;
	XL.Calculation = xlCalculationAutomatic;
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

function  fso_check_create_path(path) {
	if (fso.FolderExists(path)) {
		; ;
	}
	else {
		fso.CreateFolder(path);
	}
}

function sleep(waitMsec) {
	var startMsec = new Date();
	while (new Date() - startMsec < waitMsec);
}

function  Clipboard() 
{
	// IE のインスタンスを作成
	this.internetExplorer = new ActiveXObject('internetExplorer.Application');
	this.internetExplorer.Navigate('about:blank');
	while (this.internetExplorer.Busy) 
		sleep(10);
	// クリップボードを取得
	this.clipboard = this.internetExplorer.Document.parentWindow.clipboardDate;

	// クリップボードより文字列を取得するメソッド
	this.getText = function() 
	{
		return this.clipboard.getData('text');
	}

	this.setText = function(s) 
	{
		while (this.internetExplorer.Busy) 
			sleep(500);
		this.clipboard.setData('text', s);
	}
	this.release = function() 
	{
		this.internetExplorer.Quit();
		return true;
	}
	return this;
}

function make_range_square(range, inner_idx) 
{
	if (_flg == undefined) {
		_flg = true;
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

function txt_type(){
	/* StreamTypeEnum Values
	*/
	this.adTypeBinary = 1;
	this.adTypeText = 2;

	/* LineSeparatorEnum Values
	*/
	this.adLF = 10;
	this.adCR = 13;
	this.adCRLF = -1;

	/* StreamWriteEnum Values
	*/
	this.adWriteChar = 0;
	this.adWriteLine = 1;

	/* SaveOptionsEnum Values
	*/
	this.adSaveCreateNotExist = 1;
	this.adSaveCreateOverWrite = 2;

	/* StreamReadEnum Values
	*/
	this.adReadAll = -1;
	this.adReadLine = -2;
}

// "utf-8"
function  readFile(code, path){
	txt_type.call(this);

	var stream;
	stream = new ActiveXObject("ADODB.Stream");
	stream.type = this.adTypeText;
	stream.charset = code;
	stream.LineSeparator = this.adLF;
	stream.open();

	var tmp_lines = new Array();
	stream.loadFromFile(path);
	while ( !stream.EOS) {
		var line = stream.readText(this.adReadLine);
		var _sline=line.replace(/\r\n|\r|\n$/, "");
		tmp_lines.push(_sline);
		// msg_box("test:"+line);
	}
	stream.close();
	return tmp_lines;
}

function  writeFile(code, path,list){
	txt_type.call(this);
	var stream;

	stream = new ActiveXObject("ADODB.Stream");
	stream.type = this.adTypeText;
	stream.charset = code;
	stream.LineSeparator = this.adLF;
	stream.open();
	var idx=0;
	while(idx<list.length){
		var line=list[idx];
		stream.WriteText(line, this.adWriteLine);
		idx++;
	}
	stream.SaveToFile(path , this.adSaveCreateOverWrite);
	stream.close();

}







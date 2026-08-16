
/*
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
*/

//-----------------------------------------------------------
var fso = new ActiveXObject("Scripting.FileSystemObject") 
var  dt = new Date();
var sDt = dt.getFullYear() 
+ ('0' + (dt.getMonth() + 1)) .slice( - 2) 
+ ('0' + dt.getDate()) .slice( - 2) 
+ ('0' + dt.getHours()) .slice( - 2) 
+ ('0' + dt.getMinutes()) .slice( - 2) 
+ ('0' + dt.getSeconds()) .slice( - 2);
//-----------------------------------------------------------
var parPath = "";
var workSubDir = "";
var workDir = "";
var objShell = new ActiveXObject("Shell.Application") 
var objFolder = objShell.BrowseForFolder(0, "Select Folder", 0, "");

var parFolder = objFolder.ParentFolder;
if (parFolder) {
	parPath = parFolder.Self.Path;
}
if (objFolder) {
	workSubDir = objFolder.Self.Name;
	workDir  = objFolder.Self.Path;
}
//  var test_tmp='parPath:[' +parPath + ']workSubDir[' +workSubDir + ']workDir[' +workDir + ']';
// AddTail(test_tmp);
var ex_dir={};
ex_dir['.svn']=10;
var dir_obj={};
search_folder(dir_obj, workDir, "", "*", null,ex_dir);

var  indent_str='\t';
var  lines=[];
output_info(dir_obj, 0);

var fio= new file_obj();
var path=workDir + '/' + 'file_time_' + sDt + '.txt';
fio.writeFile("utf-8",path, lines);
ReDraw(0);


function output_info(_dir_obj, indent) {
	var  str_indent= repeat(indent_str, indent);
	var dir_list=[];
	var file_list=[];
	var idx=0;
	while(idx < _dir_obj.odrs.length ){
		var fnm=_dir_obj.odrs[idx];
		var obj=_dir_obj.objs[fnm];
		if(obj.tp==1){
			dir_list.push(fnm);
		}
		else{
			file_list.push(fnm);
		}
		idx++;
	}
	function compare(a,b){
		if ( a>b ){
			return 1;
		}
		else if( a==b){
			return 0
		}
		else{
			return -1;
		}
	}
	dir_list.sort(compare);
	file_list.sort(compare);
	for(var i in file_list){
		var fnm=file_list[i];
		var obj=_dir_obj.objs[fnm];
		var line=str_indent + fnm + '\t' +obj.time;
		lines.push(line);
	}
	for(var i in dir_list){
		var fnm=dir_list[i];
		var line=str_indent + fnm;
		lines.push(line);
		var obj=_dir_obj.objs[fnm];
		output_info(obj.subdir, (indent+1));
	}
	return;
}
function search_folder(_map, _path, _sub, _ext, _ex_ext, _ex_dir) {
	var order=[];
	var objs={};

	var fsofolder = fso.GetFolder(_path + "/" + _sub);
	var  folders = new Enumerator(fsofolder.SubFolders);
	for (; !folders.atEnd(); folders.moveNext()) {
		var chk_flg=true;
		var sub_dir = folders.item() .Name;
		if(_ex_dir && _ex_dir[sub_dir] ){
			chk_flg=false;
		}
		if(chk_flg){
			order.push(sub_dir);
			objs[sub_dir]={  tp:1, subdir:{} };
			search_folder(objs[sub_dir].subdir, _path, _sub ? _sub + "/" + sub_dir:sub_dir, _ext);
		}
	}
	var  files = new Enumerator(fsofolder.Files);
	for (; !files.atEnd(); files.moveNext()) {
		var fname = files.item() .Name;
		if (fname) {
			// var ext = mat[2].toLowerCase();
			var ext = fso.GetExtensionName(fname).toLowerCase();
			if(_ex_ext &&  _ex_ext[_ext] ){
				;;
			}
			else{
				if (typeof(_ext) == 'string') {
					if (_ext == "*" || ext == _ext) {
						order.push(fname);
						objs[fname]={  tp:2, time:get_time_str(files.item().DateLastModified)  };
					}
				}
				else if (typeof(_ext) == "object") {
					if (_ext[ext] > 0) {
						order.push(fname);
						objs[fname]={  tp:2, time:get_time_str(files.item().DateLastModified)  };
					}
				}
			}
		}

	}
	
	_map['odrs']=order;
	_map['objs']=objs;
}

function get_time_str(time_stmp) {
	var dt1=new Date(time_stmp);
	var dt_str = dt1.getFullYear() 
		+ ('0' + (dt1.getMonth() + 1)) .slice( - 2) 
		+ ('0' + dt1.getDate()) .slice( - 2) ;
		var tm_str=dt_str + ('0' + dt1.getHours()) .slice( - 2) 
		+ ('0' + dt1.getMinutes()) .slice( - 2) 
		+ ('0' + dt1.getSeconds()) .slice( - 2);

	return tm_str; 
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


function file_obj(){
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
	
	// "utf-8"
	file_obj.prototype.readFile = function (code, path){
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
	file_obj.prototype.writeFile = function (code, path,list){
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
}






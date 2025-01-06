
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

var path=GetFilename();
var iLine=GetSelectLineFrom();
var line=GetLineStr(iLine);
var folders=get_basedir(line);
// msg_box("folders["+folders+"]")


var dir=get_basedir(path);


var open_path=dir+folders;

//msg_box("folders["+folders+"]")

var chars=folders.split("");
if (chars[1]==":"||( chars[0]!="." && chars[1]=="\\") ){
	open_path=folders;
}
// msg_box("open_path["+open_path+"]")
runCommand("explorer "+open_path);




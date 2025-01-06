
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

open_type={};
open_type['txt']=open_file;
open_type['js']=open_file;

open_type['xls']=open_excel_file;
open_type['xlsx']=open_excel_file;

open_type['pdf']=open_pdf_file;

var file_name;
if (IsTextSelected) {
	file_name=ExpandParameter('$C');
} else {
	var iLine=GetSelectLineFrom();
	var line=GetLineStr(iLine);
	file_name=line.trim();
}
// msg_box("file_name["+file_name+"]")

enumFiles( [
				"D:\\svn\\nas-48-71-3B\\sakura\\macro",
				"D:\\svn\\nas-27-59-fd"] ,
			open_type, file_name);

var insert="��ғ�"
insert=inputBox_SC("�ŏ��̍s�Ƃ��Ēǉ����镶����","�ŏ��̍s", insert );
SelectAll();
var txt = GetSelectedString(0);
lines = txt.split("\n");
var nMax = lines.length;
var  map={};
var iCnt = 0;
while (iCnt < nMax) {
	var str = lines[iCnt].replace(/\r\n|\r|\n$/, "");
	var tmp=str.replace(/(^\s*)|(\s*$)/g, "");
	if(tmp!="" ){
		var nm=tmp.replace(/\\/g, "/");
		map[nm]=10;
	}
	iCnt++;
}

var fso = new ActiveXObject("Scripting.FileSystemObject") 
var objShell = new ActiveXObject("Shell.Application") 
var objFolder = objShell.BrowseForFolder(0, "Select Folder", 0, "");
var path = objFolder.Self.Path;
// msg_box("path:" +path);
path_list = [];
search_folder(path_list, path , "", "*");

var adTypeText = 2;
var adSaveCreateNotExist = 1;
var adReadAll = -1;
var stream;
stream = new ActiveXObject("ADODB.Stream");
stream.type = adTypeText;
/* charset �̒l�̗�:
*  _autodetect, euc-jp, iso-2022-jp, shift_jis, unicode, utf-8,...
*/
stream.charset = "shift_jis";         // ����  �R�[�h�`��
var  result1="-----------[�����Ώ�]------------------------------\n";
var  result2="-----------[�����ΏۈȊO]--------------------------\n";
for (var  i  in  path_list) {
	var obj = path_list[i];
	if(map[obj.sub]){             // ���� �t�@�C���p�X
	// if(map[obj.nm]){                 // ���� �t�@�C����
		var fullpath=path + "\\"+ obj.sub;
		stream.open();
		stream.loadFromFile( fullpath );
		var text = stream.readText(adReadAll);
		stream.close();
		var old=fullpath + ".old";
		if (fso.FileExists(old)){
			fso.DeleteFile(old);
		}
		var file=fso.GetFile(fullpath);
		file.Move(fullpath + ".old");
		stream.open();
		stream.writeText(insert + "\n" +text );
		stream.saveToFile(fullpath, adSaveCreateNotExist);
		stream.close();
		result1+=obj.sub + "\n";
	}
	else{
		result2+=obj.sub + "\n";
	}
}
AddTail(result1);
AddTail(result2);

msg_box("����")

function search_folder(_list, _path, _sub,_ext) {

	var fsofolder = fso.GetFolder(_path + "/" + _sub ); 
	var  folders = new Enumerator(fsofolder.SubFolders);
	for (; !folders.atEnd(); folders.moveNext()) {
		var sub_dir = folders.item() .Name;
		search_folder(_list, _path , _sub ? _sub + "/" +sub_dir:sub_dir ,  _ext);
	}
	var  files = new Enumerator(fsofolder.Files);
	for (; !files.atEnd(); files.moveNext()) {
		var fname = files.item().Name;
		var add_flg=0;
		if(_ext == "*"){
			add_flg=1;
		}
		else {
			var	mat = fname.match(/^(\S+\.)(\w+)$/);
			if (mat) {
				var ext=mat[2].toLowerCase();
				if( typeof(_ext) == 'string'){
					 if( ext ==_ext ){
					 	add_flg=1;
					 }
				}
				else if (typeof(_ext)=="object"){
					if( _ext[ext] > 0 ){
					 	add_flg=1;
					}
				}
			}
		}
		if( add_flg > 0){
			var sub_path=_sub ? _sub + "/" + fname:fname;
			_list.push( { sub:sub_path, nm:fname });
		}
	}
}

function msg_box(msg, title) 
{
	if ( !title) {
		title = "�C��"
	}
	var WSHShell = new ActiveXObject("WScript.Shell");
	WSHShell. Popup(msg, 0, title, 1);
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




var path_from="/"
var path_to="/"

SelectAll();
var txt = GetSelectedString(0);
lines = txt.split("\n");
var nMax = lines.length;
var  map={};
var path_list=[];
var iCnt = 0;
while (iCnt < nMax) {
	var str = lines[iCnt].replace(/\r\n|\r|\n$/, "");
	var tmp=str.replace(/(^\s*)|(\s*$)/g, "");
	if(tmp!="" ){
		// var nm=tmp.replace(/\\/g, "/");
		var arr_tmp=tmp.split(':');
		var path_tmp=arr_tmp[0];
		var arr=path_tmp.split('\\');
		var _tmp=arr.join('/');
		var arr_dir=_tmp.split('/');
		var _file=arr_dir.pop();
		
		var dir=arr_dir.slice(2);
		var path=dir.join('/');

		path_list.push({ dir:path,  file:_file });
	}
	iCnt++;
}


var fso = new ActiveXObject("Scripting.FileSystemObject") 

for (var  i  in  path_list) {
	var obj = path_list[i];
	var  mk_dir=path_to + '/' + obj.dir;
	var sub_path=obj.dir  + '/' +  obj.file;
	CreateFolders(mk_dir);
	object.CopyFile ( path_from + '/' + sub_path , path_to + '/' +sub_path, true );
}


function   CreateFolders(_folderspec) 
{
	var parFsoFolder = fso.GetParentFolderName(_folderspec);
	if ( !fso.FolderExists(parFsoFolder)) {
		CreateFolders(parFsoFolder);
	}
	if ( !fso.FolderExists(_folderspec)) {
		fso.CreateFolder(_folderspec);
	}
	return;
}







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


var reg1 = new RegExp( "^\\s*[A-Z][A-z]\\w*.*$" );
var reg2 = new RegExp( "^\\s*[A-z]+\\s*$" );
var reg3 = new RegExp( "^.*[A-z]+_[A-z]\\w*.*$" );




var  pid=0;
var cmd_list = new Array();

cmd_list[pid++]=    {   func:auto_make_change           ,   pattern:reg3                ,   att:1   };
cmd_list[pid++]=    {   func:auto_make_change           ,   pattern:reg1                ,   att:2   };
cmd_list[pid++]=    {   func:auto_make_change           ,   pattern:reg2                ,   att:3   };

do_select_cmd( cmd_list  );




ReDraw(0);


function to_upper( _sline, _pat, _cnt)
{
	if ( _cnt>0 )
	{
		sOut += "\n";
	}
	sOut += make_field_Upper(_sline);
}

function to_lower( _sline, _pat, _cnt)
{
	if ( _cnt>0 )
	{
		sOut += "\n";
	}
	sOut += make_field_Lower(_sline);
}
//---------------------------------------
//    cls_fuc_list()
//---------------------------------------
function  auto_make_change(_sline, _pat, _cnt)
{
	var flg_=0;
	var flg_upper=0;
	var flg_lower=0;

	var idx=0;
	var arr=_sline.split("");
	while(idx<arr.length){
		if ( arr[idx]>="A" && arr[idx]<="Z")
		{
			flg_upper=1;
		}
		else if ( arr[idx]>="a" && arr[idx]<="z")
		{
			flg_lower=1;
		}
		else if ( arr[idx] == "_" )
		{
			flg_=1;
		}
		++idx;
	}
	
	// msg_box( "flg_:" + flg_ + ";flg_upper:" + flg_upper + ";flg_lower:" + flg_lower  );
	if ( flg_ == 1 || (  ( flg_upper+flg_lower) ==1 ) )
	{
		sOut +=  make_field_Upper(_sline);
	}
	else{
	
		sOut +=  make_field_Lower(_sline);
	}
	sOut += "\n";
	return;
}



function make_field_Upper( field ){
	var arr=field.split("");
	var flg=0;
	var idx=0;
	var len=0;
	var new_word1="";
	while(idx<arr.length){
		if (flg==0 && (arr[idx]>="A" && arr[idx]<="z")){
			new_word1+=arr[idx].toUpperCase();
			flg=1;
		}
		else if(idx>0 && arr[idx-1]=="_"){
			new_word1 += arr[idx].toUpperCase();
		}
		else if(arr[idx]!="_"){
			new_word1 += arr[idx].toLowerCase();
		}
		if ( arr[idx] < "A" || arr[idx] > "z" ){
			flg=0;
		}
		++idx;
	}
	return new_word1;
}

function  make_field_Lower(field)
{
	var arr=field.split("");
	var flg=0;
	var idx=0;
	var len=0;
	var new_word1="";
	while(idx<arr.length){
		if (flg==0){
			new_word1 += arr[idx].toLowerCase();
		}
		else if ( flg==1 && ( arr[idx] >="A" && arr[idx]<="Z" )){
			new_word1 += "_" + arr[idx].toLowerCase();
		}
		else{
			new_word1 += arr[idx];
		}
		if ( ( arr[idx] >="A" && arr[idx]<="z"  ) ){
			flg=1;
		}
		else{
			flg=0;
		}
		++idx;
	}
	return new_word1.toUpperCase();
}

function make_head_Lower(field)
{
	if (field==""){
		return "";
	}
	var arr=field.split("");
	arr[0]=arr[0].toLowerCase();
	return arr.join("");
}
function make_head_Upper(field)
{
	if (field==""){
		return "";
	}
	var arr=field.split("");
	arr[0]=arr[0].toUpperCase();
	return arr.join("");
}




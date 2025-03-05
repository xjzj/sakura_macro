

var start=GetSelectLineFrom();
var end=GetSelectLineTo();

var txt_list=[];
var idx=start;

var max_col=0;
while(idx<=end){
	var line=GetLineStr(idx);
	line=line.replace(/\r\n|\r|\n$/, "");
	
	var path=GetFilename();
	var arr_path=path.split("\\");
	
	var new_arr=arr_path.slice(2);
	
	
	var s_line=new_arr.join("/") +":"+idx +":"+line;
	var len=getTabStrLen(s_line);
	if(len>max_col){
		max_col=len;
	}
	
	
	txt_list.push(new_arr.join("/") +":"+idx +":"+line );
	idx++;
}

// var  obj={ cmd:1, row:txt_list.length+2 , col:max_col+10 };


InputBox('path' , 'txt list', txt_list )
// RadioGroupBox( gIeDoc, 'path' , 'txt list', txt_list, obj, 10  );

function InputBox(prompt, title, list){
	var txt="";
	var idx=0;
	var strShowTxt="";
	while(idx<list.length){
		var line=list[idx];
		if(idx>0){
			strShowTxt+='" & vbNewLine & "';
		}
		strShowTxt+=line;
		idx++;
	}
	inputBox_SC(prompt, title, strShowTxt);
}

function getTabStrLen(str) {
	var _tab_len = 4
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

function inputBox_SC(prompt, title, def)
{

    var result;
    var objScr;
    
    objScr =new ActiveXObject("MSScriptControl.ScriptControl");
    objScr.language="VBScript";
    objScr.addCode(
        "Function getInput()" +
        '    getInput = InputBox("' + prompt + '", "' + title +  '", "' + def + '")' +
        "End Function");
    result = objScr.eval("getInput");
    objScr = null;
    return result;
}




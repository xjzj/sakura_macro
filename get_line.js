
var w_h=1980;
var h_v=1080;

//IEを起動
var objIE = new ActiveXObject('internetExplorer.Application');
objIE.Toolbar = 0
objIE.Height = 250
objIE.Width = 500
objIE.Left = parseInt((w_h - objIE.Width)/2)
objIE.Top = parseInt((h_v - objIE.Height)/2)
objIE.Visible = false;



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

var  obj={ cmd:1, row:txt_list.length+2 , col:max_col+10 };


objIE.Height = obj.row*30+50;
objIE.Width =  obj.col*10;
objIE.Left = parseInt((w_h - objIE.Width)/2)
objIE.Top = parseInt((h_v - objIE.Height)/2)

objIE.Navigate('about:blank');
var  gIeDoc = objIE.Document;
objIE.Visible = true;



RadioGroupBox( gIeDoc, 'path' , 'txt list', txt_list, obj, 10  );
	

objIE.Visible = false;
objIE.Quit();      //  IEを終了


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


function RadioGroupBox( ieDoc, sTitle, sCaption, lines, _obj, FontSize)
{
	// 解像度の取得
	// Call GetResolution(h,w)

	ieDoc.write("<TITLE>" + sTitle + "</TITLE>");
	ieDoc.write("<BODY STYLE=overflow-y:hidden BGColor=ButtonFace>");
	ieDoc.write("<TABLE WIDTH=100% HEIGHT=100% BORDER=0 STYLE=font-size:" + FontSize + "pt>");
	
	
	ieDoc.write("<SPAN>" +  sCaption.replace(/\r\n|\r|\n$/, "<BR>") + "</SPAN>");
	ieDoc.write("<FORM NAME=\"Form\">");
	ieDoc.write("<TR><TD ALIGN=Left VALIGN=Top>");
	
	ieDoc.write("<INPUT TYPE=HIDDEN NAME=\"Button\" value=0>");

	


	// ieDoc.write("</TD>");   // 
	// ieDoc.write("<TD  ALIGN=Left VALIGN=Top>");   // 
	ieDoc.write("              ");
	
	
	ieDoc.write("<textarea rows=\""+ _obj.row+ "\" cols=\""+ _obj.col + "\"  readonly>");
	for(var i in  lines){
		ieDoc.write( lines[i] +  "\r\n" );
	}
	ieDoc.write("</textarea>");
	
	/*
                <select id="dropdown1" name="dropdown1">
                    <option value="option1">Option 1</option>
                    <option value="option2">Option 2</option>
                    <option value="option3">Option 3</option>
                </select>
	
	*/
	
	// ieDoc.write("</TD></TR><TR><TD ALIGN=Center VALIGN=Bottom>");   // 
	
	// ieDoc.write("<TD  ALIGN=Left VALIGN=Top>");   //  ALIGN=Center
	ieDoc.write("&nbsp;&nbsp;<INPUT   TYPE=BUTTON VALUE=\"FINISH\" onclick=\"Form.Button.value=1\">");
	ieDoc.write("</TD><TD  ALIGN=right VALIGN=Top>");   //  ALIGN=Center

	ieDoc.write("</TD></TR>");   // 
	// ieDoc.Write("<INPUT TYPE=BUTTON VALUE=""キャンセル"" ｏnclick=""Form.Button.value=2;Radio.value=-1"">");
	ieDoc.write("</FORM></TABLE></BODY>");
	// On Error Resume Next 'エラー無視 (ウィンドウの[×]終了対応)
	
	if(!ieDoc.Form.Button){
		_obj.cmd=0;
		return;
	}

	//  ユーザ操作取得
	var flg=0
	while(flg==0)
	{
		sleep(100);
		// WScript.Sleep(100);
		flg = parseInt(ieDoc.Form.Button.value)
		if(objIE== null)
		{
			break;
		}
	}
	// var ret = parseInt(ieDoc.Form.Radio.value);  //  取得データは文字のため
	_obj.cmd=flg;

	return;
}



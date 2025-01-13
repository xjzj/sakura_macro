

var reg_outline = /^([^(]+)\((\d+),\d+\): (.+)$/;

var  pid = 0;
var cmd_list = new Array();
cmd_list[pid++] = { func:parse_outline , pattern:reg_outline  , att:0 };


var max_pat_idx=0;
var pat_cnt_map={};
pat_cnt_map[max_pat_idx]=0;

if (IsTextSelected) {
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

	//  var txt = GetSelectedString(0);
	//  var lines = txt.split("\n");
	//  nMaxLineCnt = lines.length;
	
	var start=GetSelectLineFrom();
	var end=GetSelectLineTo();

	// var iCnt = 0;
	var iCnt = start;
	var line_list=[];
	// while (iCnt < nMaxLineCnt) {
	while(iCnt<=end){
		// var tmp = lines[iCnt];
		var tmp=GetLineStr(iCnt);
		var line=tmp.replace(/\r\n|\r|\n$/, "");

		var _idx=check_pattern(line);
		if(_idx>-1){
			line_list.push(line);
			if(!pat_cnt_map[_idx]){
				pat_cnt_map[_idx]=1;
			}
			else{
				++pat_cnt_map[_idx];
			}
			if( max_pat_idx!=_idx && pat_cnt_map[_idx]>pat_cnt_map[max_pat_idx] ){
				max_pat_idx=_idx;
			}
		}
		iCnt++;
	}

	var max_col=0;
	var cmd_obj=cmd_list[max_pat_idx];
	var txt_list=[];
	iCnt = 0;
	while (iCnt < line_list.length) {
		var line=line_list[iCnt];
		var tag=cmd_obj.func(line);
		var len=getTabStrLen(tag);
		if(len>max_col){
			max_col=len;
		}
		
		txt_list.push(tag);
		iCnt++;
	}
	
	var  obj={ cmd:1, row:txt_list.length+2 , col:max_col+30 };

	objIE.Height = obj.row*20+0;
	objIE.Width =  obj.col*9;
	objIE.Left = parseInt((w_h - objIE.Width)/2)
	objIE.Top = parseInt((h_v - objIE.Height)/2)

	objIE.Navigate('about:blank');
	var  gIeDoc = objIE.Document;
	objIE.Visible = true;

	RadioGroupBox( gIeDoc, 'path' , 'txt list', txt_list, obj, 10  );

	objIE.Visible = false;
	objIE.Quit();      //  IEを終了
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



function check_pattern( _line )
{
	var idx=0;
	while (idx < cmd_list.length) {
		var pat=cmd_list[idx].pattern;
		var mat=_line.match(pat);
		if(mat){
			break;
		}
		idx++;
	}
	if( idx== cmd_list.length ){
		idx=-1;
	}
	return idx;
}

function  parse_outline (line)
{
	var mat=line.match(reg_outline);
	var path=mat[1];
	var num=mat[2];
	var content=mat[3];
	
	var tp="m";

	var def_list=content.split('::');
	
	var def=content;
	var cmt=content;
	
	if(def_list.length >1){
		def=def_list.pop();
		var tmp=def_list[0];
		var _list=tmp.split("\\");
		if( def=="定義位置" ){
			def=_list.pop();
			if( _list.length>0){
				cmt=_list.join('::');
			}
			else{
				cmt="";
			}
			cmt+= "::" + def + '::定義位置';
		}
		else{
			cmt=_list.join('::');
		}
	}
	else{
		cmt = '';
	}

	var arr_path=path.split("\\");
	var new_arr=arr_path.slice(2);
	var tag=def + "\t" + new_arr.join("\\") + "\t" +  num + ';"' +  "\t"   + tp  + "\t" +  cmt ;
	return tag;
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



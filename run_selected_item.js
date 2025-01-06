//-----

//-----------------------------------------------------------
/*
var fso = new ActiveXObject("Scripting.FileSystemObject")

var objShell = new ActiveXObject("Shell.Application")
var objFolder = objShell.BrowseForFolder(0, "Select Folder", 0, "");
var workpath = objFolder.Self.Path;
*/




// var currentpath = new ActiveXObject("WScript.Shell") .CurrentDirectory 

var currentpath = 'D:/RegEditor/table/JStable';






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

// タグ記述

var  script_list=[ 
	"run_kvm_client.js"
];

/*
var max_len=0
var connect_list=[];
for( var i in vm_list ){
	var vm=vm_list[i];
	for( var j in port_list ){
		var port=port_list[j];
		var conn=vm+":" + port;
		if( conn.length > max_len ){
			max_len=conn.length ;
		}
		connect_list.push(conn);
	}
}
*/

var  obj={ cmd:1,  it:0 };

while(obj.cmd==1){
	objIE.Navigate('about:blank');
	var  gIeDoc = objIE.Document;
	objIE.Visible = true;
	// var ret=RadioGroupBox('connect vm' , 'vm list', connect_list, 0, 10, 10*max_len,  20*connect_list.length +120  );
	RadioGroupBox( gIeDoc, 'run script' , 'script list', script_list, obj, 10  );

	// AddTail( connect_list[ret] );
	if(obj.cmd==1){
		var  srcipt=script_list[obj.it];

		// runCommand(remote_viewer + '  spice://' + connect_list[ret]);
		runCommand('wscript  ' + currentpath + "/" +  srcipt );
	}
}

objIE.Visible = false;
objIE.Quit();      //  IEを終了


function runCommand(cmd) 
{
	var wsh = new ActiveXObject("WScript.shell");
	var oe = wsh.Exec(cmd);
	// var path= env.ExpandEnvironmentStrings ("%SAKURA_SCRIPT%’) ;
	// var r = oe.StdOut.ReadAll();
	// return r;
}

function RadioGroupBox( ieDoc, sTitle, sCaption, scripts,  _obj, FontSize)
{
	// 解像度の取得
	// Call GetResolution(h,w)

	ieDoc.write("<TITLE>" + sTitle + "</TITLE>");
	ieDoc.write("<BODY STYLE=overflow-y:hidden BGColor=ButtonFace>");
	ieDoc.write("<TABLE WIDTH=100% HEIGHT=100% BORDER=0 STYLE=font-size:" + FontSize + "pt>");
	
	
	ieDoc.write("<SPAN>" +  sCaption.replace(/\r\n|\r|\n$/, "<BR>") + "</SPAN>");
	ieDoc.write("<FORM NAME=\"Form\">");
	ieDoc.write("<TR><TD ALIGN=Left VALIGN=Top>");
	
	ieDoc.write("<INPUT TYPE=HIDDEN NAME=\"Button\" VALUE=\"0\">");

	
	
	ieDoc.write("<INPUT TYPE=HIDDEN NAME=\"Radio\"  VALUE=\"" + 0 + "\">");
	var idx=0;
	for(var i in  scripts){
		var chk="";
		if( idx== 0){
			chk= "checked"
		}
		ieDoc.write("　<INPUT TYPE=RADIO Name=\"RG\" onclick=\"Radio.value=" + i + "\" " + chk + ">" + scripts[i]+ "</INPUT><BR>");
		idx++;
	}
	

	/*
	ieDoc.write("<select id=\"host\" name=\"host\">");
	var idx=0;
	for(var i in  hosts){
		var chk="";
		if( idx== _obj.host){
			chk= "selected"
		}
		ieDoc.write("　<option value=" + i + " " + chk + " >" + i + ":" + hosts[i] +  "</option>" );
		idx++;
	}
	ieDoc.write("</select>");

	// ieDoc.write("</TD>");   // 
	// ieDoc.write("<TD  ALIGN=Left VALIGN=Top>");   // 
	ieDoc.write("              ");
	
	
	ieDoc.write("<select id=\"port\" name=\"port\">");
	var idx=0;
	for(var i in  ports){
		var chk="";
		if( idx== _obj.port){
			chk= "selected"
		}
		ieDoc.write("　<option value=" + i + " " + chk + " >" + i + ":" + ports[i] +  "</option>" );
		idx++;
	}
	ieDoc.write("</select>");
	*/
	
	/*
                <select id="dropdown1" name="dropdown1">
                    <option value="option1">Option 1</option>
                    <option value="option2">Option 2</option>
                    <option value="option3">Option 3</option>
                </select>
	
	*/
	
	// ieDoc.write("</TD></TR><TR><TD ALIGN=Center VALIGN=Bottom>");   // 
	ieDoc.write("</TD></TR><TR><TD  ALIGN=Left >");   //  ALIGN=Center
	ieDoc.write("<INPUT TYPE=BUTTON VALUE=\"　　RUN　\" onclick=\"Form.Button.value=1\">　");

	// ieDoc.write("</TD><TD ALIGN=Center VALIGN=Bottom>");   // 
	//  ieDoc.write("</TD><TD ALIGN=Center>");   // 
	ieDoc.write("              ");
	ieDoc.write("<INPUT TYPE=BUTTON VALUE=\"　　EXIT　　\" onclick=\"Form.Button.value=2\">　");
	

	
	ieDoc.write("</TD></TR>");   // 
	// ieDoc.Write("<INPUT TYPE=BUTTON VALUE=""キャンセル"" ｏnclick=""Form.Button.value=2;Radio.value=-1"">");
	ieDoc.write("</FORM></TABLE></BODY>");
	// On Error Resume Next 'エラー無視 (ウィンドウの[×]終了対応)

	//  ユーザ操作取得
	var flg=0
	while(flg==0)
	{
		sleep(100);
		flg = parseInt(ieDoc.Form.Button.value)
		if(objIE== null)
		{
			break;
		}
	}
	// var ret = parseInt(ieDoc.Form.Radio.value);  //  取得データは文字のため
	
	var iIt = parseInt(ieDoc.Form.Radio.value);  //  取得データは文字のため
	// var iPort = parseInt(ieDoc.Form.port.value);  //  取得データは文字のため

	_obj.cmd=flg;
	_obj.it=iIt;

	return;
}

function sleep(waitMsec) {
	var startMsec = new Date();
	while (new Date() - startMsec < waitMsec);
}



function get_driver_info(path) {
	var flg = false;
	var obj = null;
	var ForReading = 1, ForWriting = 2;
	var FilePointer = fso.OpenTextFile(path, ForReading, true);
	while (!FilePointer.AtEndOfStream) {
		var line = FilePointer.ReadLine();
		var mat = line.match(/^\s*DriverVer\s*=\s*([^,]+)\s*,\s*([0-9.]+)\s*/);
		if (mat) {
			var  date_arr = mat[1].split('/');
			var  ver = mat[2];
			if (date_arr.length ==3) {
				var  yer = date_arr.pop();
				date_arr.unshift(yer);
				var _dt = date_arr.join('');
				flg = true;
				obj = { dt:_dt, ver:ver };
				break;
			}
		}
	}
	FilePointer.Close();
	return obj;
}


var Cmd;
//-----------------------------------------------------------
function inpath(file)
{
	var ForReading = 1, ForWriting = 2;
	var wsh = new ActiveXObject("wscript.shell");
	var env=wsh.Environment("SYSTEM");
	var path=env.item("SAKURA_SCRIPT") + "\\";
	// AddTail ( path );
	var FileOpener = new ActiveXObject( "Scripting.FileSystemObject");
	var FilePointer = FileOpener.OpenTextFile(path + file, ForReading, true);
	Cmd = FilePointer.ReadAll();
}

//-----------------------------------------------------------
inpath("inc.js");
eval(Cmd);

/*
// aaa=11;
// aaa=aaa?aaa.toString():"--"
// msg_box( aaa.toString() )
var tmp="Aaa_Bbb_Ccc";
var tt=make_field_Lower(tmp);

var ss={}
ss["11"]=12
if(!ss["11"]){
	msg_box("not define 11")
}

tp=-9999
if ( tp==0){

	msg_box(tt);

}

tt=null
aa=typeof(tt);
msg_box("aa="+aa);
*/

// var test_book=open_excel("XXXX","test.xlsx");
// var test_sheet=test_book.Worksheets("Sheet1");
// var test_range= test_sheet.Range("B3:D7")
// make_range_square(test_range,3);






FileOpen( getEnvVar("SAKURA_SCRIPT") + "/" + "sakura_list.txt" );



function  make_field_Lower(field)
{
	var arr=field.split("");
	var flg=1;
	var idx=1;
	var len=0;
	var new_word1=arr[0];
	while(idx<arr.length){
		if ( arr[idx] <= "Z" ){
			flg=0;
		}
		else if( arr[idx] >= "a" ){
			flg=1;
		}
		else{
			flg=2;
		}
		
		
		if (flg==0){
			new_word1 += "_";
		}
		if ( flg != 2){
			new_word1 += arr[idx];
		}

		

		++idx;
	}
	return new_word1.toUpperCase();
}


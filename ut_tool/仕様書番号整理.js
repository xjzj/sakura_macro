
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

var  brackets_dec =  /^\s*(\uff08[\uff10-\uff19\uff0e\.0-9]+\uff09)([^\uff09\uff10-\uff19\uff0e\.0-9].*)$/;  // （１）（２）（３）
var  brackets_katakana =  /^\s*(\uff08[\uff71-\uff9d\uff08\uff09]+\uff09)([^\uff71-\uff9d].*)$/;  // （１）（２）（３）

var  katakana=[ "--", "ｱ", "ｲ", "ｳ", "ｴ", "ｵ", "ｶ", "ｷ", "ｸ", "ｹ", "ｺ", "ｻ", "ｼ", "ｽ", "ｾ", "ｿ", "ﾀ", "ﾁ", "ﾂ", "ﾃ", "ﾄ", "ﾅ", "ﾆ", "ﾇ", "ﾈ", "ﾉ", "ﾊ", "ﾋ", "ﾌ", "ﾍ", "ﾎ", "ﾏ", "ﾐ", "ﾑ", "ﾒ", "ﾓ", "ﾔ", "ﾕ", "ﾖ", "ﾗ", "ﾘ", "ﾙ", "ﾚ", "ﾛ", "ﾜ", "ﾝ" ];
var  full_dec=[ "０","１", "２", "３", "４", "５", "６", "７", "８", "９", "１０", "１１", "１２", "１３", "１４", "１５", "１６", "１７", "１８", "１９", "２０", "２１", "２２", "２３", "２４", "２５", "２６", "２７", "２８", "２９", "３０", "３１", "３２", "３３", "３４", "３５", "３６", "３７", "３８", "３９", "４０", "４１", "４２", "４３", "４４", "４５", "４６", "４７", "４８", "４９", "５０", "５１", "５２", "５３", "５４", "５５", "５６", "５７", "５８", "５９", "６０", "６１", "６２", "６３", "６４", "６５", "６６", "６７", "６８", "６９", "７０", "７１", "７２", "７３", "７４", "７５", "７６", "７７", "７８", "７９", "８０", "８１", "８２", "８３", "８４", "８５", "８６", "８７", "８８", "８９", "９０", "９１", "９２", "９３", "９４", "９５", "９６", "９７", "９８", "９９"  ];
// var  head_list=[ "B1", "C1", "F1", "G1"];
var  chk_list=[ 
           { rc:"B1" , flg:0 , exp:null              , head:full_dec , deli:null } , // ２．
           { rc:"C1" , flg:1 , exp:null              , head:full_dec , deli:null } , // ２．１
           { rc:"D1" , flg:1 , exp:null              , head:full_dec , deli:null } , // ２．２．２
           
           { rc:"E1" , flg:1 , exp:brackets_dec      , head:full_dec , deli: [ "（" , "）" ] } , // （１）
           { rc:"F1" , flg:1 , exp:brackets_katakana , head:katakana , deli: [ "（" , "）" ] } , // （ｱ）

           null
];

var  chk_list_for_num={};
var  chk_num_list=[];



make_XL();
// var wk_book=XL.ActiveWorkbook;
var wk_sht=XL.ActiveSheet;

var order_cnt=0;
var num_order={};
var head_nums=[];
var idx=0;
var len=chk_list.length-1;
while( idx<len ){
	var obj=chk_list[idx];
	var col=wk_sht.Range(obj.rc).Column;
	chk_list_for_num[col]=obj;
	num_order[col]=order_cnt; order_cnt++;
	chk_num_list.push(col);
	head_nums.push(0);
	idx++;
}

var max_row=wk_sht.UsedRange.Rows(wk_sht.UsedRange.Rows.Count).Row;
var  max_col_str=chk_list[len-1].rc;
var max_col=wk_sht.Range(max_col_str).Column;

var sel=XL.Selection;
var row=sel.Row;

var  first=1;
while(row<=max_row){
	var col=1;
	while(col<=max_col){
		var tmp=wk_sht.Cells(row,col).Value;
		var chk_val=tmp?tmp.toString().replace(/(^\s*)|(\s*$)/g,""):"";
		if( chk_val ){
			var del_flg=wk_sht.Cells(row,col).Font.Strikethrough;
			if( del_flg ){
				;;
			}
			else{
				var obj=chk_list_for_num[col];
				if(obj){
					if(first ){
						if(!obj.exp){
							var ret=get_first(chk_val);
							if(ret.flg){
								first=0;
								var i=0;
								while(i < ret.lst.length ){
									head_nums[i]=ret.lst[i];
									i++;
								}
							}
						}
						break;
					}
					var ret=chk_for_new_str(chk_val, obj, col);
					if(ret.flg){
						wk_sht.Cells(row,col)=ret.str;
						wk_sht.Cells(row,col).Interior.Color = convert_RGB(255, 255, 0);
					}
				}
			}
			break;
		}
		col++;
	}
	row++;
}

function chk_for_new_str(str, _obj, _col ){

	var old_head="";
	var str_head;
	var body="";
	var new_line="";
	var order=num_order[_col];
	
	if( !_obj.exp  ){
		var mat=str.match(/^\s*([\uff10-\uff19\uff0e\.0-9]+)([^\uff10-\uff19\uff0e\.0-9].*)$/);
		if( mat ){
			old_head=mat[1];
			body=mat[2];
		}
		else{
			body=str;
			if(_obj.flg){
				return { flg:0 , str:"" };
			}
		}
		
		
		var head=[];var i=0;
		while(i < head_nums.length ){
			var num=head_nums[i];
			if(order<=i){
				num++;
				if(order<i){
					num=0;
				}
				head_nums[i]=num;
			}
			if(order>=i){
				head.push(full_dec[num]);
			}
			
			i++;
		}
		if(head.length==1){
			head.push("");
		}

		str_head=head.join('．');
	}
	else{
		var deli_flg=0;
		var mat=str.match(_obj.exp);
		if( mat ){
			old_head=mat[1];
			body=mat[2];
		}
		else{
			deli_flg=1;
			body=str;
			if(_obj.flg){
				return { flg:0 , str:"" };
			}
		}
	
		var num=head_nums[order];
		num++;
		head_nums[order]=num;
		head_nums[order+1]=0;
		var ch=_obj.head[num];
		
		// str_head= "\uff08" + full_dec[num] + "\uff09";
		str_head= _obj.deli.join(ch);
		if(deli_flg){
			str_head+= "　";
		}
	}
	
	var chg_flg=1;
	if( str_head == old_head ){
		chg_flg=0;
	}
	new_line= str_head+ body;

	return { flg:chg_flg , str:new_line };

}


function get_first(str){

	var ret=0;
	var num_list=[];
	var mat=str.match(/^\s*([\uff10-\uff19\uff0e\.0-9]+)([^\uff10-\uff19\uff0e\.0-9].*)$/);
	if(mat){
		var tmp=mat[1];
		var i=0;
		var num_str="";
		while(i<tmp.length){
			var asc=tmp.charCodeAt(i);
			if(asc>=0xFF10){
				asc-=0xFF10;
			}
			if(asc>=0x30 ){
				asc-=0x30;
			}
			var deli_flg=1;
			if(asc>=0 && asc<=9){
				num_str+=asc;
				deli_flg=0;
			}
			if(deli_flg){
				var num=Number(num_str);
				num_list.push(num);
				num_str="";
				ret=1;
			}
			i++;
		}
		if(num_str!=""){
			var num=Number(num_str);
			num_list.push(num);
			ret=1;
		}
	}

	return { flg:ret, lst:num_list };
}




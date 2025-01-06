
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
var xlValues=-4163;
var xlPart=2;
var xlWhole=1;
var xlByRows=1;
var xlNext=1;
var xlCellTypeLasstCell=11;

eval(Cmd);

case_list={};
case_list[  "bulkProcessing"                                        ]=  {   book:"test_case.xlsx"   ,   st:"マトリクス１"   ,   tbl:null    ,   item_start:"E6" ,   val:"L6"    ,   start:"P6"  ,   end:"AD24"  };
case_list[  "VrfyStatusListDaoSelectAuthSubRepExecCanPriceInfoTest" ]=  {   book:"test_case.xlsx"   ,   st:"マトリクス３"   ,   tbl:"I6"    ,   item_start:"N6" ,   val:"AB6"   ,   start:"AL6" ,   end:"BU86"  , test_data:{ st:"テストデータ", tbl:"B10"   }};

case_list[  "VrfyStatusListDaoSearchNo6Test"                        ]=  {   book:"test_case.xlsx"   ,   st:"マトリクス２"   ,   tbl:"K5"    ,   item_start:"P5" ,   val:"AA5"   ,   start:"AI5" ,   end:"AZ50"  ,

vars:{
"今日の前":"testData.no6Time.getTodayBefore()",
"今日":"testData.no6Time.getToday()",
"今日の後":"testData.no6Time.getTodayAfter()",
"次15営業日の前":"testData.no6Time.getTodayNext15Before()",
"次15営業日":"testData.no6Time.getTodayNext15()",
"次15営業日の後":"testData.no6Time.getTodayNext15After()",
"前7営業日の前":"testData.no6Time.getTodayBefore7Before()",
"前7営業日":"testData.no6Time.getTodayBefore7()",
"前7営業日の後":"testData.no6Time.getTodayBefore7After()"
}

};



var export_tbl_sql = /^Inser into EXPORT_TABLE\s*\([^\)]+\)\s*values\s*\(\s*'([^']+)'\s*,\s*'([^']+)'\s*,\s*'([^']+)'\s*,[^,]+,[^,]+,\s*'([^']+)'\s*\);$/ ;
var export_tbl_sql2 = /^Inser into EXPORT_TABLE\s*\([^\)]+\)\s*values\s*\(\s*'([^']+)'\s*,\s*'([^']+)'\s*,\s*'([^']+)'\s*,[^,]+,[^,]+,\s*(null)\s*\);$/ ;   // '


var reg_print = /^.+(\.|\s+)(\w+)\s+--\s*(\S.*\S)\s*$/;
var reg_print2 = /^\s*(\S[^\t]*)\t([^\t]+)\t([^\t]+)\t([^\t]+)\t([^\t]+)\s*$/;
var reg_print3 = /^\s*(\S[^\t]*)\t([^\t]+)\t([^\t]+)\t([^\t]+)\s*$/;
var reg_print4 = /^\s*(\S[^\t]*)\t([^\t]+)\t([^\t]+)\t([^\t]+)\t([^\t]+)\t([^\t]+)\s*$/;
var reg_print5 = /^\s*(\S[^\t]*)\t([^\t]+)\t([^\t]+)\s*$/;


var reg_drder1=/^(\s*)(\w+)\s+(\w+)\s*= \s*new\s+(\w+)\s*\(\s*\)\s*;\s*$/;
var reg_drder11=/^(\s*)(\w+)\s+(\w+)\s*= \s*make(\w+)\s*\(\s*\)\s*;\s*$/;
var reg_drder2=/^(\s*)(\w+)(\.set\w+\s*\(.*\S)\s*$/;
var reg_drder3=/^(\s*\w+\.add)\s*\(\s*(\w+)\s*\)\s*;\s*$/;


var  reg_dto1 = /^\s*\/\*\*\s+(.*)\s+\*\/\s*$/;
var  reg_dto2 = /^\s*(public|private)\s+(String|Long|Integer|Date|BigDecima)\s+(\w+)\s*;\s*$/;
var  reg_dto3 = /^\s*(public|private)\s+List<(String|Long|Integer|Date|BigDecima)>\s+(\w+)\s*;\s*$/;




var reg_make_ut_data0=          /^\s*\/\/\s*TBL:(\w+)\s*$/;
var reg_make_ut_data1=          /^\s*\/\/\s*\S.*\(([^\)]+)\)\s*$/;
var reg_make_ut_data2=          /^(\s*)([A-z]+)(\d+)\.(\w+)\((.+)\);$/;
var reg_make_ut_data3=          /^(\s*)(\w+)()\.(\w+)\((.+)\);$/;
var reg_make_ut_data_assert=    /^(\s*assertThat.+\.isEqualTo)\(([^\)]+)\);$/;





var  pid=0;
var cmd_list = new Array();

cmd_list[pid++]=    {   func:dto_list           ,   pattern:reg_dto1                ,   att:1   };
cmd_list[pid++]=    {   func:dto_list           ,   pattern:reg_dto2                ,   att:2   };
cmd_list[pid++]=    {   func:dto_list           ,   pattern:reg_dto3                ,   att:3   };

cmd_list[pid++]=    {   func:order_data         ,   pattern:reg_drder1              ,   att:1   };
cmd_list[pid++]=    {   func:order_data         ,   pattern:reg_drder11             ,   att:4   };
cmd_list[pid++]=    {   func:order_data         ,   pattern:reg_drder2              ,   att:2   };
cmd_list[pid++]=    {   func:order_data         ,   pattern:reg_drder1              ,   att:3   };


cmd_list[pid++]=    {   func:print_fuc          ,   pattern:reg_print               ,   att:1   };
cmd_list[pid++]=    {   func:print_fuc          ,   pattern:reg_print2              ,   att:2   };
cmd_list[pid++]=    {   func:print_fuc          ,   pattern:reg_print3              ,   att:2   };
cmd_list[pid++]=    {   func:print_fuc          ,   pattern:reg_print4              ,   att:2   };
cmd_list[pid++]=    {   func:print_fuc          ,   pattern:reg_print5              ,   att:2   };


cmd_list[pid++]=    {   func:export_tbl_fuc     ,   pattern:export_tbl_sql          ,   att:1   };
cmd_list[pid++]=    {   func:export_tbl_fuc     ,   pattern:export_tbl_sql2         ,   att:1   };

cmd_list[pid++]=    {   func:make_ut_data_set   ,   pattern:reg_make_ut_data0       ,   att:-1  };
cmd_list[pid++]=    {   func:make_ut_data_set   ,   pattern:reg_make_ut_data1       ,   att:1   };
cmd_list[pid++]=    {   func:make_ut_data_set   ,   pattern:reg_make_ut_data2       ,   att:2   };
cmd_list[pid++]=    {   func:make_ut_data_set   ,   pattern:reg_make_ut_data3       ,   att:3   };
cmd_list[pid++]=    {   func:make_ut_data_set   ,   pattern:reg_make_ut_data_assert ,   att:4   };


do_select_cmd( cmd_list  );


function make_ut_data_set( _sline, _pat, _cnt, att, mat )
{
	if ( !('num' in make_ut_data_set )) {
		make_ut_data_set.num=0;
		// var case_name="bulkProcessing";
		// var case_name="VrfyStatusListDaoSearchNo6Test";
		var case_name="VrfyStatusListDaoSelectAuthSubRepExecCanPriceInfoTest";
		var case_no=1;
		make_ut_data_set.next_print={ line:"", flg:0 };

		make_ut_data_set.field={ top_it:"", cnt:0, tbl:"", nm:"", case_no:case_no };
		make_ut_data_set.obj_num={};
		make_ut_data_set.data_map={};
		

		var case_book=open_excel("XXXX",case_list[case_name].book);
		var case_sheet=case_book.Worksheets(case_list[case_name].st);
		//-----------------------------------------------------
		var data_tbl_col=null;
		var data_tbl_range=null;
		var data_sheet=null;
		
		if( case_list[case_name].test_data ){
			data_sheet = case_book.Worksheets( case_list[case_name].test_data.st );
			data_tbl_range =  data_sheet.Range( case_list[case_name].test_data.tbl );
			
			var lastrow =  data_sheet.Cells.SpecialCells( xlCellTypeLasstCell ).Row;
			var lastcol =  data_sheet.Cells.SpecialCells( xlCellTypeLasstCell ).Column;
			data_tbl_col=data_tbl_range.Column;
			make_ut_data_set.test_data={ sheet:data_sheet, tbl_rg:data_tbl_range, tbl_col:data_tbl_col, row_end:lastrow, col_end:lastcol, cache:{} };
		}
		
		function excelRange(dat,tbl,item){
			if ( !dat.cache[tbl] ){
				var tbl_range=dat.sheet.Range( dat.tbl_rg, dat.sheet.Cells( dat.row_end, dat.tbl_col ));
				var range=tbl_range.Find( tbl, tbl_range.Cells(1,1), xlValues, xlWhole, xlByRows, xlNext, false,false,false );
				dat.cache[tbl]={rg:range};
			}
			var nm=make_field_Lower(item);
			if ( !dat.cache[tbl][nm] ){
				var name_range=dat.sheet.Range( dat.cache[tbl].rg, dat.sheet.Cells( dat.cache[tbl].rg.Row, dat.col_end ));
				var range=name_range.Find(nm, name_range.Cells(1,1), xlValues, xlWhole, xlByRows, xlNext, false,false,false );
				dat.cache[tbl][nm]={rg:range};
			}
			return dat.cache[tbl][nm].rg;
		}

		
		function make_data( dat, tbl, item, no, val_map, row ){
			if ( dat ){
				var mp=val_map[row];
				var tmp=mp.val;
				if ( mp.tp == 99 ){
					val=Number(tmp)+no;
				}
				else if( mp.tp == 3 ){
					val=tmp.slice(0,4) + "/" + tmp.slice(4,6) + "/" + tmp.slice(6,8) + " " + tmp.slice(8,10) +":" +  tmp.slice(10,12) +":" + tmp.slice(12,14);
				}
				else if( mp.tp == 2 ){
					val=tmp.slice(0,4) + "/" + tmp.slice(4,6) + "/" + tmp.slice(6,8);
				}
				else{
					if ( tmp=="null" ){
						val="(NULL)"
					}
					else{
						val=tmp;
					}
				}
				var nm_range=excelRange( dat, tbl, item);
				dat.sheet.Cells(nm_range.Row+no, nm_range.Column ) = val;
				if ( mp.tp==99 ){
					dat.sheet.Cells(nm_range.Row+no,nm_range.Column).Interior.Color=convert_RGB(150,150,150);
				}
				else{
					dat.sheet.Cells(nm_range.Row+no,nm_range.Column).Interior.Color=convert_RGB(255,200,200);
				}
			}
		}
		//-----------------------------------------------------
		function get_data(dat, ss, ee){
			if ( dat.sheet ){
				var fld=make_ut_data_set.field;
				var tbl_nm=fld.tbl;
				var item=fld.nm;
				if( !tbl || tbl=="" ){
					throw new Error("[// TBL:XXXXXXXXXXXXX]" + " not exsit" );
				}
				var nm_range=excelRange( dat, tbl, item);
				var tmp=nm_range.Offset(1,0).Value;
				if ( tmp.length==10 && tmp.match(/^\d{4}\/\d{2}\/\d{2}$/) ){
					val='DateTimeUils.date("' + tmp + '","yyyy/MM/dd")' ;
				}
				else{
					val=ss+tmp+ee;
				}
				
				return val;
			}
		}
		//-----------------------------------------------------
		var tbl_col=0;
		if (case_list[case_name].tbl){
			tbl_col=case_sheet.Range(case_list[case_name].tbl).Column;
		}
		var item_col=case_sheet.Range(case_list[case_name].item_start).Column;
		var val_col=case_sheet.Range(case_list[case_name].val).Column;
		
		var row_start= case_sheet.Range(case_list[case_name].start).Row;
		var row_end= case_sheet.Range(case_list[case_name].end).Row;
		var col_start= case_sheet.Range(case_list[case_name].start).Column;
		var col_end= case_sheet.Range(case_list[case_name].end).Column;
		//-----------------------------------------------------
		var val_map={};
		var val_range=case_sheet.Range( case_sheet.Cells(row_start,val_col),case_sheet.Cells(row_end,val_col));
		for(var i=1;i<=val_range.cells.count;i++){
			var nm=val_range.cells(i).Value;
			nm=nm?nm.toString().trim():"";
			var row=val_range.cells(i).Row;
			if ( nm !="" ){
				var cmt=nm;
				var val="";
				var tp=0;
				
				var mt=nm.match(/^\s*([^:：. 	]+)[:：.].*$/);
				if (mt){
					val=mt[1];
					
					// msg_box( "val_map[row].val:" + val_map[row].val + ";val_map[row].cmt:" + val_map[row].cmt )
				}
				if (val==""){
					if( case_list[case_name].vars && case_list[case_name].vars[nm] ){
						val=case_list[case_name].vars[nm];
						tp=-1;
					}
					else{
						val=nm;
						var mt=nm.match( /^\w+$/ );
						if( mt ){
							tmp=nm.toLowerCase();
							if( tmp=="null" ){
								val=tmp;
								tp=1;
							}
						}
					}
				}
				if ( val.length==8 && val.match( /^([0-9]+)$/ ) ){
					tp=2;  //date
				}
				if ( val.length==14 && val.match( /^([0-9]+)$/ ) ){
					tp=3;  //date
				}
				var mt=val.match( /^(\d+)\+CaseNo$/ );
				if ( mt ){
					val=mt[1];
					tp=99;
				}
				
				val_map[row]={ val:val, cmt:cmt, tp:tp};
			}
		}
		//-----------------------------------------------------
		var map=make_ut_data_set.data_map;
		var tbl_nm=""
		var item="";
		var item_range=case_sheet.Range( case_sheet.Cells(row_start,item_col),case_sheet.Cells(row_end,item_col));
		for(var i=1;i<=item_range.cells.count;i++){
			var row=val_range.cells(i).Row;
			if ( tbl_col>0){
				var tmp=case_sheet.Cells(row,tbl_col).Value;
				tmp=tmp?tmp.toString().trim():"";
				if (  tmp!="" ){
					tbl_nm=tmp;
				}
			}
			
			var nm=item_range.cells(i).Value;
			nm=nm?nm.toString().trim():"";
			if ( nm !="" ){
				var  mt=nm.match(/^[^\(]*\((\w+)\)\s*$/);
				if (mt){
					if( !map[tbl_nm] ){
						map[tbl_nm]={};
					}
					item=get_key(mt[1]);
					map[tbl_nm][item]={};
				}
			}
			var case_no=1;
			for( var col=col_start;col<=col_end; col++ ){
				var flg=case_sheet.Cells(row,col).Value;
				if( flg=="○" || flg=="〇" ){
					map[tbl_nm][item][case_no]=val_map[row];
					//-----------------------------------------------------
					make_data( make_ut_data_set.test_data, tbl_nm, item, case_no, val_map, row )
					//-----------------------------------------------------

					// msg_box( "item[" + item + "]case_no[" + case_no + "]val_map[row].val:" + val_map[row].val + ";val_map[row].cmt:" + val_map[row].cmt )
				}
				case_no++;
			}
		}
		//-----------------------------------------------------
	}
	
	function get_fix(str)
	{
		var func_flg=0;
		var ss="";
		var ee="";
		var arr=str.split("");
		if ( arr[arr.length-1]==")" ){
			var mt=str.match( /^\s*(new\s+BigDecimal\(|DateTimeUtils.date\().+(\))$/ );
			if ( mt ){
				ss=mt[1];
				ee=mt[2];
			}
			func_flg=1;
		}
		else{
			if ( arr[0]=='"' ){
				ss=arr[0];
				ee=ss;
			}
			else if( arr[arr.length-1] > "9" ){
				ss="";
				ee=arr[arr.length-1];
			}
		}
		return { flg:func_flg, ss:ss, ee:ee };
	}
	function get_value(str)
	{
		var val="";
		var map=make_ut_data_set.data_map;
		var fld=make_ut_data_set.field;
		var obj=null;
		if ( 	map[fld.tbl] && 
				map[fld.tbl][fld.nm] && 
				map[fld.tbl][fld.nm][fld.case_no]){
			obj=map[fld.tbl][fld.nm][fld.case_no];
		}
		else{
			obj={ val:"", tp:-9999, cmt:"", fmt:"" };
		}
		var fix=get_fix(str);
		if( fix.flg==1 && ( 2<=obj.tp && obj.tp<=3 ) ){
			val=fix.ss+'"'+obj.val+'", "' + obj.fmt + '"' + fix.ee;
		}
		else if ( obj.tp==99 ){
			var tmp=obj.val ;
			val= fix.ss+ (Number(tmp)+fld.case_no) + fix.ee ;
		}
		else{
			val= fix.ss+ obj.val + fix.ee ;
		}
		return { val:val, cmt:"		// "+obj.cmt };
	}
	//------------------------------------------------
	
	if ( att==-1 ){
		var flg=0;
		var fld=make_ut_data_set.field;
		fld.tbl=mat[1];
		var nxt= make_ut_data_set.next_print;
		nxt.line=_sline;
		nxt.flg=2;
		// sOut += _sline + "\r\n";
		return;
	}
	else if ( att==1 ){
		var flg=0;
		var fld=make_ut_data_set.field;
		var nm=get_key(mat[1]);
		if ( fld.top_it=="" ){
			fld.top_it=fld.tbl+":"+nm;
			fld.cnt++;
			flg=1;
		}
		else if(  fld.top_it==(fld.tbl+":"+nm) ){
			fld.case_no++;
			fld.cnt++;
			flg=1;
		}
		fld.nm=nm;
		if ( flg==1 ){
			//  sOut += "==================================" + fld.case_no + "     cnt:" + fld.cnt + "\r\n";
			
			sOut += "			break;"  + "\r\n";
			
			
			
			sOut += "			//========" + fld.case_no  + "\r\n";
			sOut += "			case " + fld.case_no  + ":"+ "\r\n";
		}
		
		var nxt= make_ut_data_set.next_print;
		if ( nxt.flg>0 ){
			sOut += _sline + "\r\n";
			nxt.flg=0;
		}

		sOut += _sline + "\r\n";
		return;
	}
	else if( att==2 ||att==3  ){
	
		var obj=mat[2];
		var obj_num=make_ut_data_set.obj_num;
		var fld=make_ut_data_set.field;
		if ( obj_num[obj] ){
			if ( fld.m == fld.top_it ){
				obj_num[obj]++;
			}
		}
		else{
			obj_num[obj]=1;
		}
	
		//---------------------------
		make_ut_data_set.test_data=null;
		if (  make_ut_data_set.test_data  ){
			var fix=get_fix(mat[5]);
			val=get_data(make_ut_data_set.test_data, fix.ss, fix.ee );
			sOut += mat[1] + obj + "." + mat[4] + "(" + val + "); " + "\r\n";
			return;
		}
		//---------------------------
		var ret=get_value(mat[5]);
		if ( att == 2 ){
			sOut += mat[1] + obj + obj_num[obj] + "." + mat[4] + "(" + ret.val + "); " + ret.cmt + "\r\n";
		}
		else{
			sOut += mat[1] + obj + "." + mat[4] + "(" + ret.val + "); " + ret.cmt; // + "\r\n";
		}
		
		return;

	}
	else if(att==4){
		var ret=get_value(mat[2]);
		sOut += mat[1] + "(" + ret.val + "); " + ret.cmt; // + "\r\n";
		return;
	}
	sOut += _sline +  "\r\n";
	return;
}



function dto_list( _sline, _pat, _cnt, att, mat )
{
	if ( !('num' in dto_list )) {
		dto_list.num=0;
		dto_list.order=[];
		dto_list.structs={};
		dto_list.nm="";
		dto_list.map={};
		var map=dto_list.map;
		map["String"]       =   {   tp:"CHAR"   ,   len:1       };
		map["Long"]         =   {   tp:"NUMBER" ,   len:15      };
		map["Integer"]      =   {   tp:"NUMBER" ,   len:1       };
		map["Date"]         =   {   tp:"DATE"   ,   len:""      };
		map["BigDecimal"]   =   {   tp:"NUMBER" ,   len:"1,2"   };
	}
	dto_list.num++;
	if ( att == 1 ){
		dto_list.nm=mat[1];
		return;
	}
	if ( att == 2 || att == 3 ) {
		var typ=mat[2];
		var val=make_field_Lower(mat[3]).toUpperCase();
		var line = dto_list.nm + "\t" + val + "\t" ;
		var map=dto_list.map;
		line += map[typ].tp + "\t" + map[typ].len + "\t" + "Yes";
		dto_list.order.push(val);
		dto_list.structs[val]=line;
		
		var get="	/**" + "\r\n";
		// get += "	* @return " + mat[3]  +  "\r\n";
		get += "	* @return the " + mat[3]  +  "\r\n";
		get += "	*/" + "\r\n";
		get += "	public " + typ + " get" +  make_field_Lower(mat[3])  + "() {\r\n";
		get += "		return " + mat[3]  +  ";\r\n";
		get += "	}" +  "\r\n";
		
		var set="	/**" + "\r\n";
		// set += "	* @param " + mat[3]  + " セットする " + mat[3]  +  "\r\n";
		set += "	* @param " + mat[3]  + " the " + mat[3]  + " to set " + "\r\n";
		set += "	*/" + "\r\n";
		set += "	public void set"  +  make_field_Lower(mat[3]) + "(" + typ + " " + mat[3] +  ") {" + "\r\n";
		set += "		this." + mat[3] + " = " + mat[3] +  ";\r\n";
		get += "	}" +  "\r\n";

		sOut += get +  "\r\n";
		sOut += set +  "\r\n";
	}
	
	if (  dto_list.num == nMaxLineCnt ){
		for ( var it in dto_list.order ){
			var val =dto_list.order[it];
			// msg_box(tbl);
			sOut += dto_list.structs[val];
		}
	}
	return ;
}

function order_data( _sline, _pat, _cnt, att, mat )
{
	if ( !('num' in order_data )){
		order_data.num=0;
		order_data.structs={};
		order_data.items={};
	}
	
	
	// msg_box( mat )
	if ( (att == 4 || att ==1 ) && mat[2]==mat[4]  ){
		
		// msg_box( mat[3] )
		var _idx=0;
		var struct=mat[2];
		var it=mat[3];
		var mt=it.match(/^\s*(\D+)(\d+)\s*/);
		if (mt){
			nm=mt[1];
			_idx=Number(mt[2]);
		}
		if ( order_data.structs[struct] ){
			order_data.structs[struct].num++;
			order_data.structs[struct].item[it]=order_data.structs[struct].num;
		}
		else{
			order_data.structs[struct]={ num:_idx, nm:nm, item:{} };
			order_data.structs[struct].item[it]=_idx;
		}
		order_data.items[it]=struct;
		if ( att == 1 ) {
			sOut += mat[1] + struct + " " + order_data.structs[struct].nm + order_data.structs[struct].item[it] + " = new " + struct + "():\r\n";
		}
		else if ( att == 4 ) {
			sOut += mat[1] + struct + " " + order_data.structs[struct].nm + order_data.structs[struct].item[it] + " = make " + struct + "():\r\n";
		}
		
		// msg_box( _cnt + ":" + _sline )
		// msg_box( it + ":" + struct )
	}
	else if (  att == 2 ) {
		// msg_box( _sline )
		try{
			var it=mat[2];
			var struct=order_data.items[it];
			sOut += mat[1] +  order_data.structs[struct].nm + order_data.structs[struct].item[it] + mat[3] + "\r\n";
		}
		catch(e){
			msg_box( att + ":" + _cnt + ":" + _sline );
			msg_box( it + ":" + struct );
			
			msg_box(e);
			throw e;
		}
	}
	else if ( att == 3 ) {
		// msg_box( _sline )
		try{
			var it=mat[2];
			var struct=order_data.items[it];
			sOut += mat[1] + "(" + order_data.structs[struct].nm + order_data.structs[struct].item[it]  + ")\r\n";
		}
		catch(e){
			msg_box( att + ":" + _cnt + ":" + _sline );
			msg_box( it + ":" + struct );
			
			msg_box(e);
			throw e;
		}
	}
	else{
		sOut += _sline + "\r\n";
	}
	return;

}

function print_fuc( _sline, _pat, _cnt, att, mat ){
	if (  !('num' in print_fuc ) ){
		print_fuc.flg=0;
		print_fuc.num=0;
		print_fuc.sOut00="****************************************************\r\n";
		print_fuc.sOut01="****************************************************\r\n";
		print_fuc.sOut02="****************************************************\r\n";
		print_fuc.sOut03="****************************************************\r\n";
		print_fuc.sOut1="****************************************************\r\n";
		print_fuc.sOut2="****************************************************\r\n";
		print_fuc.sOut3="****************************************************\r\n";
		print_fuc.sOut4="****************************************************\r\n";
		print_fuc.sOut40="****************************************************\r\n";
		print_fuc.sOut5="****************************************************\r\n";
		print_fuc.order_tbl=0;
		print_fuc.d=new Date(2019,2,20,10,20);
		print_fuc.asc1=49;
		print_fuc.no_allow="@=?*:;\",()<>/\\";
		// print_fuc.no_asc={};
		print_fuc.flg_used={};
		print_fuc.flg_it={
			REORD_ORD_TYP:1,
			CNCL_CORR_KBN:1,
			ORD_TYP_CD:1,
			LIQUIDATION_KBN:1,
			TRD_TYP_CD:1,
			IFA_CLIENT_FLUG:[0,1],
			ORD_STS_CD:1,
			STOP_PRICE_STS:1,
			VALUE_TYPE:1,
			EXEC_COND_CD:1,
			DELAY_RESULT_CD:1,
			STOP_PRICE_COND_KBN:1,
			FORCE_KBN:1,
			SPA_KBN:1,
			FX_FLUG:1,
			TAX_KBN:1,
			COMMISSION_ZERO_ETF_FLG:1,
			DISP_VALID_FLG:1
		};
		
		print_fuc.tps={};
		print_fuc.values={};
		print_fuc.values_a={};
		print_fuc.items={};
		
		
		print_fuc.select_order={};
	}
	// msg_box("_cnt:" + _cnt)
	//------------------------------------------
	print_fuc.num++;
	var tmp=_sline.match(/^\s*([-\/]+)\s*(\w*)(\s*|\s+.*)$/);
	if ( tmp ){
		print_fuc.order_tbl=tmp[2];
		print_fuc.sOut3+="// TBL:" + print_fuc.order_tbl + "\r\n";
		return;
	}
	if ( print_fuc.order_tbl!=0 ){
		var tmp = _sline.match(/^\s*(\w+)\s*$/);
		if ( tmp ) {
			var val=tmp[1];
			if (  print_fuc.select_order[print_fuc.order_tbl] ){
				;;
			}
			else{
				print_fuc.select_order[print_fuc.order_tbl]={ order:[], map:{} };
			}
			print_fuc.select_order[print_fuc.order_tbl].order.push(val);
			return;
		}
		if ( att==2 ){
			var item=mat[2];
			if ( print_fuc.select_order[print_fuc.order_tbl] ){
				print_fuc.select_order[print_fuc.order_tbl].map[item]=(  _sline + "\r\n" );
				return;
			}
		}
	}
	//------------------------------------------
	//------------------------------------------
	function get_str(_it,_len, _len2, _cnt, tp )
	{
		if ( typeof(_len)=="string"){
			_len=Number(_len);
		}
		if ( print_fuc.flg_it[_it]){
			// msg_box("_cnt:" + _cnt)
			var itm=print_fuc.flg_it[_it];
			var _ret="";
			if ( typeof(itm)!= "object" ){
				;;
				item=[0,1,2];
			}
			var _mod=itm.length;
			do
			{
				var idx=_cnt%_mod;
				_ret=itm[idx].toString();
				_cnt++;
			}
			while( print_fuc.flg_used[_it]==_ret )
			print_fuc.flg_used[_it]=_ret;
			print_fuc.values[_ret]=1;
			
			return ( '"' + _ret + '"' );
		}
		
		var idx=0;
		while ( _len == 1 ){
			if ( print_fuc.values_a.length >60 ){
				print_fuc.values_a.clear();
			}
			idx++;
			var str=String.fromCharCode(print_fuc.asc1++);
			if ( !print_fuc.values_a[str] || print_fuc.values_a[str] < 1 ){
				print_fuc.values_a[str]=idx;
				print_fuc.values[str]=idx;
				ret='"' + str + '"';
				return ret;
			}
		}
		
		var ret="";
		var idx=0;
		while(1){
			var str="";
			var lft=0;
			var al=idx+1;
			var al_bak=al;
			do
			{
				lft=al%74;
				al=parseInt(al/74);
				if ( lft>0 ){
					var tmp=String.fromCharCode(47+lft);
					var pos=print_fuc.no_allow.indexOf(tmp);
					if ( pos!=-1 ){
						al=al_bak+1;
						al_bak=al;
						str="";
						continue;
					}
					// if ( print_fuc.no_asc[tmp]){
					// 	continue;
					// }
					str=str+tmp;
				}
				// msg_box( "lft:" + lft + ";str:" + str );
			}
			while(al>0)
			
			var new_str = repeat( str, _len );
			
			if ( new_str.length > _len ){
				new_str=new_str.slice(0,_len);
			}
		
			if ( !print_fuc.values[str] || print_fuc.values[str]<1 ){
				print_fuc.values[str]=idx+2;
				ret=new_str;
				break;
			}		
			idx++;
		}
		ret='"' + ret + '"';
		return ret;
	}
	//-------------------------------------------
	function get_date(_it,_len, _len2, _cnt, tp )
	{
		var ret="";
		var idx=1;
		var day=print_fuc.d;
		while(1){
			idx++;
			
			var date= new Date(	day.getFullYear()
								,day.getMonth()
								,day.getDate()+1
								,day.getHours()
								,day.getMinutes());
			var str="";
			str +=   getFixedNum(    date.getFullYear()  ,   4);
			str +=   getFixedNum(    date.getMonth()     ,   2);
			str +=   getFixedNum(    date.getDate()      ,   2);
			str +=   getFixedNum(    date.getHours()     ,   2);
			str +=   getFixedNum(    date.getMinutes()   ,   2);
			str +="00";
			if ( !print_fuc.values[str] || print_fuc.values[str]<1 ){
				print_fuc.values[str]=idx;
				ret=str;
				break;
			}
			day=date;
		}
		ret= 'DateTimeUtils.date("' + ret + '","yyyyMMddHHmmss")';
		return ret;
	}
	//-------------------------------------------
	function get_num(_it,_len, _len2, _cnt, tp )
	{
		if ( typeof(_len)=="string" ){
			_len=Number(_len);
		}
		if ( typeof(_len2)=="string" ){
			_len2=Number(_len2);
		}
		if ( _len>19 ){
			_len=19;
		}
		
		// msg_box("0-name:" + _it + ";_cnt:" + _cnt +  ";_len:" + _len + ":_len2:" + _len2 );
		if ( print_fuc.flg_it[_it]){
			// msg_box("_cnt:" + _cnt)
			var itm=print_fuc.flg_it[_it];
			var _ret="";
			if ( typeof(itm)!= "object" ){
				;;
				itm=[0,1,2];
			}
			var _mod=itm.length;
			do
			{
				var idx=_cnt%_mod;
				// msg_box("idx:" + idx +  ";_mod:" + _mod  )
				_ret=itm[idx].toString();
				_cnt++;
			}
			while( print_fuc.flg_used[_it]==_ret )
			print_fuc.flg_used[_it]=_ret;
			print_fuc.values[_ret]=1;
			
			return  _ret ;
		}
		
		var ret="";
		var idx=1;
		while(1){
			var str=idx.toString();
			var new_str=repeat(str,_len);
			if ( new_str.length > _len ){
				new_str=new_str.slice(0,_len);
			}
			if ( !print_fuc.values[str] || print_fuc.values[str]<1 ){
				print_fuc.values[str]=idx;
				ret=new_str;
				break;
			}		
			idx++;
		}
		
		if ( tp=="FLOAT"  && _len2>0 ){
			ret=ret.slice(0,(_len-_len2)) + "." + ret.slice((_len-_len2)) + "d";
			return ret;
		}
		
		if ( _len2>0 || _len==0 ){
			ret='new BigDecimal(' + ret + ')'
		}
		else{
			if ( ret.length > _len ){
				ret=ret.slice(0,_len);
			}
			if ( _len>7 ){
				ret=ret+"L";
			}
		}
		return ret;
	}
	
	
	//---------------------------------------------------------------------------------
	function judge_tps(tps,field )
	{
		if ( !field || field == "" ){
			return tps.tp;
		}
		if ( tps.tp != "NUMBER" ){
			return tps.tp;
		}
		
		var num=0;
		var num2=0;
		var ret="Integer";
		if ( typeof(field) == "number" ){
			num=field;
		}
		else{
			var arr=field.split(",");
			num=Number(arr[0]);
			if (arr[1]){
				num2=Number(arr[1]);
			}
		}
		if ( num2>0){
			ret="BigDecimal"
		}
		else if( num >7 ){
			ret="Long";
		}
		return ret;
	}
	
	if (  print_fuc.flg == 0 ){
		print_fuc.tps['VARCHAR2']   =   {   func:get_str    ,   tp:"String"         };
		print_fuc.tps['CHAR']       =   {   func:get_str    ,   tp:"String"         };
		print_fuc.tps['DATE']       =   {   func:get_date   ,   tp:"java.sql.Date"  };
		print_fuc.tps['NUMBER']     =   {   func:get_num    ,   tp:"NUMBER"         };
		print_fuc.tps['FLOAT']      =   {   func:get_num    ,   tp:"BigDecimal"     };
		print_fuc.tps['FLOAT']      =   {   func:get_str    ,   tp:"Double"         };
		print_fuc.flg = 1;
	}
	
	var cmt="";
	var nm="";
	var tp="";
	var len="";
	var len2="";
	var yn="";
	
	
	if ( att==1 ){
		nm=mat[2];
		cmt=mat[3];
	}
	if ( att==2 ){
		nm=mat[2];
		cmt=mat[1];
		tp=mat[3];
		var tmp=mat[4];
		if( tmp ){
			var arr=tmp.split(",");
			len=arr[0];
			if (arr[1]){
				len2=arr[1];
			}
		}
		yn=mat[5];
	}
	
	try{
		if (  att > 0  ){
			var unm= make_field_Upper(nm);
			print_fuc.sOut1 += 'System.out.println("' + cmt + '(' + mat[2] + ')' + ':" + data.get' + unm + '());' + "\r\n"  ;
			print_fuc.sOut2 += 'System.out.println("' + cmt + '(' + make_head_Lower(unm) + ')' + ':" + data.get' + unm + '());' + "\r\n" ;
			
			print_fuc.sOut3 +=	'// ' + cmt + '(' + mat[2] + ')' + "\r\n" ;
			print_fuc.sOut4 +=	'// ' + cmt + '(' + mat[2] + ')' + "\r\n" ;
			print_fuc.sOut40 +=	'// ' + cmt + '(' + make_head_Lower(unm) + ')' + "\r\n" ;
			print_fuc.sOut5 +=	'// ' + cmt + '(' + make_head_Lower(unm) + ')' + "\r\n" ;
			
			
			print_fuc.sOut02 +=	 cmt + '(' + mat[2] + ')' + "\r\n" ;
			print_fuc.sOut03 +=	 cmt + '(' + make_head_Lower(unm) + ')' + "\r\n" ;
			
			
			print_fuc.sOut4 +=	'assertThat(result.get(0).get' + unm + '()).isEqualTo(allData1.get' + unm + '());' + "\r\n" + "\r\n"  ;
			print_fuc.sOut40 +=	'assertThat(result.get(0).get' + unm + '()).isEqualTo(allData1.get' + unm + '());' + "\r\n" + "\r\n"  ;
			
			print_fuc.sOut5 +=  'assertNull( data.get' + unm + '());'  + "\r\n" + "\r\n"  ;
		
			print_fuc.sOut00 +=  judge_tps(print_fuc.tps[tp],mat[4]) + " " + make_head_Lower(unm) + ";" + ' // ' + cmt  + "\r\n"  ;
			print_fuc.sOut01 += make_head_Lower(unm) + " = " + 'dto.get' + unm + '();' + ' // ' + cmt + "\r\n"  ;
			
			if ( att == 2 ) {
				var dat="dto";
				if ( print_fuc.order_tbl == "VRFY_STATUS_LIST" ){
					dat="insStatus1";
				}
				else if( print_fuc.order_tbl == "VRFY_ORD_DELAY_LIST" ){
					dat="insOrd1";
				}
				else if( print_fuc.order_tbl == "VRFY_RPT_ORD_LIST" ){
					dat="insRpt1";
				}
				
				// if ( _cnt == 89 ){
				// 	msg_box( "_cnt:" + _cnt + ";mat[2]:" + mat[2] + ";len:" + len );
				// }
				
				print_fuc.sOut3 +=	dat + '.set' + unm + '(' + print_fuc.tps[tp].func(nm,len,len2,_cnt,tp) + ');'  + "\r\n" + "\r\n"  ;

			}
			return;
		
		}
	}
	catch(e){
		msg_box(_cnt + ":" + _sline );
		msg_box( "tp:" + tp + ";len:" + len + ";_pat:" + _pat );
		msg_box(e);
		throw e;
	}
	
	sOut += _sline + "\r\n";
	
	if (  print_fuc.num == nMaxLineCnt ){
		for ( var tbl in print_fuc.select_order ){
			sOut += "-----------------------------" + tbl +  "\r\n"  ;
			var idx =0;
			while( idx < print_fuc.select_order[tbl].order.length  ){
				var it=print_fuc.select_order[tbl].order[idx];
				if ( print_fuc.select_order[tbl].map[it] ){
					sOut += print_fuc.select_order[tbl].map[it] ;
				}
				++idx;
			}
		}
	}
	
	if(  print_fuc.num == nMaxLineCnt ){
		sOut    +=  print_fuc.sOut00    +   "\r\n"  ;
		sOut    +=  print_fuc.sOut01    +   "\r\n"  ;
		sOut    +=  print_fuc.sOut02    +   "\r\n"  ;
		sOut    +=  print_fuc.sOut03    +   "\r\n"  ;


		sOut    +=  print_fuc.sOut1     +   "\r\n"  ;
		sOut    +=  print_fuc.sOut2     +   "\r\n"  ;
		sOut    +=  print_fuc.sOut3     +   "\r\n"  ;
		sOut    +=  print_fuc.sOut4     +   "\r\n"  ;
		sOut    +=  print_fuc.sOut40    +   "\r\n"  ;
		sOut    +=  print_fuc.sOut5     +   "\r\n"  ;
	}
}

function export_tbl_fuc( _sline, _pat, _cnt, att, mat )
{
	if (  !('num' in export_tbl_fuc ) ){
		export_tbl_fuc.num=1;
	}
	
	if ( att == 1 ){
		var tp="";
		var sz="";
		var item=mat[2].match(/^\s*(\w+)\(([^\)]+)\)\s*$/);
		if ( item ){
			tp=item[1];
			sz=item[2];
			var mt=sz.match(/^(\d+)\s*BYTE\s*$/);
			if ( mt ){
				sz = mt[1];
			}
		}
		else{
			tp=mat[2];
		}
		
		sOut += mat[4] + '\t' + mat[1] + '\t' + tp + '\t' + sz + '\t' + mat[3] + "\n" ;
		return;
	}
	return;
}


ReDraw(0);


function get_key( field ){

	var flg_=0;
	var flg_upper=0;
	var flg_lower=0;
	var idx=0;
	var arr=field.split("");
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
	if ( flg_ == 1 || ( flg_upper+flg_lower) ==1 ){
		field=make_field_Upper(field);
	}
	return make_head_Lower(field);
}

function make_field_Upper( field ){
	var arr=field.split("");
	var flg=0;
	var idx=0;
	var len=0;
	var new_word1="";
	while(idx<arr.length){
		if ( arr[idx] == "_"  ){
			;;
		}
		else if( flg==0 ){
			if (arr[idx]>="A" && arr[idx]<="z"){
				new_word1+=arr[idx].toUpperCase();
				flg=1;
			}
			else{
				new_word1 += arr[idx]
			}
		}
		else{
			new_word1 += arr[idx].toLowerCase();
		}
	
		/*
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
		*/
		
		if ( arr[idx] < "A" || arr[idx] > "z" || arr[idx] == "_"){
			flg=0;
		}
		++idx;
	}
	return new_word1;
}

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



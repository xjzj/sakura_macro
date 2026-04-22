
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

var xlValidateList=3;
var xlValidAlertStop=1;
var xlBetween=1;

var folder="c:/work/doc/UT";

var itm_map={ "集計対象":10, "識別キー":10, "通番":10, "品管STS":10, "連携対象":10, "Redmine":10, "関連監理番号":10, "状態1":10, "発行日":10, "発行者":10, "品質管理単位":10, "検出処理ID":10, "処理名":10, "C/S":10, "試験項目番号":10, "検出AP":10, "故障概要（題名）":10, "故障内容／状況（付加情報）":10, "故障検出日":10, "故障検出者":10, "解析状況":10, "故障原因":10, "故障AP":10, "解析日":10, "解析者":10, "処置内容":10, "修正AP":10, "修正APSVN Rev":10, "修正ドキュメント":10, "修正ドキュメントSVN Rev":10, "設計修正要否":10, "修正日":10, "修正者":10, "試験日":10, "確認者":10, "確認結果":10, "故障解決日":10, "発生箇所":10, "発生現象":10, "影響度":10, "検出契機":10, "バグ種別":10, "バグ箇所":10, "混入工程":10, "混入現象":10, "混入原因":10, "レビュー未発現原因":10, "摘出すべき試験工程":10, "試験未発現原因":10, "分類":10, "設計修正ドキュメント":10, "設計修正依頼先":10, "設計修正依頼日":10, "設計修正ドキュメントSVN Rev":10, "設計書修正日":10, "設計書修正者":10, "設計書修正確認者":10, "横展開要否":10, "状態2":10, "横展開対象処理":10, "横展開観点":10, "横展開実施結果":10 };

//-------------------------------------------------------------------
var ctrl_info={ folder:"C:/work/doc", book:"仕様分析.xlsx", sht:"故障反映"};

var fso=new ActiveXObject("Scripting.FileSystemObject");

var debug="";
var tbl_list="";
make_XL();
var wk_book=XL.ActiveWorkbook;
var wk_sht=XL.ActiveSheet;

var bookName=wk_book.Name;
var shtName=wk_sht.Name;

function input_bug()
{
	var fobj= new file_obj();
	var bug_input_list=fobj.readFile("utf-8", folder+'/' + "bug_input.txt");
	var setting_txt=bug_input_list.join("\n");
	debug += setting_txt + "\n";
	eval(setting_txt);
	var ctrl_book=open_excel(ctrl_info.folder, ctrl_info.book);
	var ctrl_sht= ctrl_book.Worksheets(ctrl_info.sht);
	
	var row=input_list.start_row;
	var ctl_col=input_list.col;
	var ctl_end_row=input_list.end_row;
	
	var sel=XL.Selection;
	
	var wk_col=8;
	var wk_row= sel.Row;
	
	while(row<=ctl_end_row){
		var txt=ctrl_sht.Cells(row,ctl_col).Value;
		wk_sht.Cells(wk_row,wk_col)=txt;
		row++;
		wk_col++;
	}
	var bug_num=wk_sht.Cells(wk_row,3).Value;
	ctrl_sht.Cells(1,ctl_col)=bug_num;

}




function  mk_bug_list()
{
	var input_list={start_row:0, end_row:0};
	debug += "mk_bug_list" + "\n";
	var path=folder+'/' + "bug_input.txt";
	var fobj =new file_obj();
	if (fso.FolderExists(path)){
		var bug_input_list=fobj.readFile("utf-8", path);
		var setting_txt=bug_input_list.join("\n");
		debug += setting_txt + "\n";
		eval(setting_txt);
	}
	
	get_setting();
	var sel=XL.Selection;
	var col=sel.Column;
	var row=sel.Row;
	var max_row=wk_sht.UsedRange.Rows(wk_sht.UsedRange.Rows.Count).Row;
	
	
	var all_col=sel.EntireColumn;
	all_col.ColumnWidth=50;
	
	var start_row=0;
	
	var end_row=0;
	
	debug += "sel_addr:" + sel.Address + "\n";
	debug += "col:" + col + "\n";
	debug += "row:" + row + "\n";
	debug += "max_row:" + max_row + "\n";

	while(row<=max_row){
		var tmp=wk_sht.Cells(row,1).Value;
		var val_txt=tmp?tmp.toString().replace(/(^\s*)|(\s*$)/g,""):"";
		
		if( start_row==0 ){
			if( val_txt=="状態1" ) {
				var tmp=wk_sht.Cells(row,col).Value;
				var val_status=tmp?tmp.toString().replace(/(^\s*)|(\s*$)/g,""):"";
				if( val_status ){
					debug += "val_status:" + val_status + "\n";
					start_row=row;
					break;
				}

			}
		}
		
		// debug += "val_txt:" + val_txt + "\n";
		if( val_txt && itm_map[val_txt]){
			var flg=true;
			var input_flg=false;
			// debug += "val_txt:" + val_txt + "\n";
			var obj=setting_list[val_txt];
			if( typeof(obj)=='number' ){
				if( obj > 10 ){
					wk_sht.Cells(row,col)=obj;
					wk_sht.Cells(row,coll).Interior.Color = convert_RGB(200, 230, 255);
				}
				else{
					flg=false;
				}
			}
			else if( typeof(obj)=='function' ){
				wk_sht.Cells(row, col)=obj();
				wk_sht.Cells(row, col).Interior.Color = convert_RGB(200, 230, 255);
			}
			else if( typeof(obj)=='object' ){
				
				if( obj.list){
					var valid=wk_sht.Cells(row,col).Validation;
					valid.Delete();
					valid.Add(xlValidateList, xlValidAlertStop, xlBetween, obj.list.join(','));
					// valid.IgnoreBlank=true;
					valid.InCellDropdown=true;
					// valid.IMEMode=0;
					valid.ShowInput=true;
					// valid.ShowError=true;
					wk_sht.Cells(row,col)=obj.list[obj.val];
				}
				else{
					wk_sht.Cells(row,col)=obj.val;
				}
				wk_sht.Cells(row,col).Interior.Color = obj.color;
			}
			else if( typeof(obj)=='string' ){
				wk_sht.Cells(row,col)=obj;
				wk_sht.Cells(row,col).Interior.Color = convert_RGB(255,255,220);
			}
			
			if( start_row==0 && flg ){
				start_row=row;
			}
			
		}
		
		if(  "横展開実施結果" == val_txt ){
			end_row=row;
			break;
		}
		row++;
	}
	if( end_row==0 ){
		if(start_row==0){
			start_row=input_list.start_row;
		}
		end_row=input_list.end_row;
	}
	else{
		var all_range=wk_sht.Range(wk_sht.Cells(start_row,col), wk_sht.Cells(end_row,col));
		make_range_square(all_range,3);
		
	}
	
	
	bug_input_list=[];
	bug_input_list.push("var input_list={");
	bug_input_list.push("	col:" + col+ ",");
	bug_input_list.push("	start_row:" + start_row + ",");
	bug_input_list.push("	end_row:"+end_row);
	bug_input_list.push("};");
	// var fobj= new file_obj();
	AddTail(folder + '/' + "bug_input.txt");
	fobj.writeFile("utf-8", folder + '/' + "bug_input.txt", bug_input_list);
}
var excel_list={};
excel_list["仕様分析.xlsx:故障反映"]={  func:mk_bug_list };
excel_list["【グループ共通IT】RUR_UT故障管理表_vr20250425.xlsm:故障管理表"]={ func:input_bug };
excel_list["仕様分析.xlsx:Sheet2"]={ func:input_bug };


var wk_book_sht=bookName + ":" + shtName;


var setting_list;
debug += "\n";
debug += "wk_book_sht:" + wk_book_sht + "\n";
for(var key in  excel_list){
	debug += "key:" + key + "\n";
}

var do_function=excel_list[wk_book_sht].func;

if(do_function){
	// debug += "do_function:" + do_function + "\n";
	do_function();
	AddTail( debug + "\n");
}
else{
	AddTail("setting not work" + "\n");
}
AddTail("***************************" + "\n");


function get_date(){
	var dt = new Date();
	var sDt = dt.getFullYear() + '/'
	+ ('0' + (dt.getMonth() + 1 )).slice(-2) + "/"
	+ ('0' + dt.getDate()).slice(-2);
	return sDt;
}


function convert_RGB(r, g, b) {
	var color=0;
	color += r;
	color += g <<8;
	color += b << 16;
	return color;
}



function get_setting(){
	setting_list={
		"集計対象":5,
		"識別キー":5,
		"通番":5,
		"品管STS":5,
		"連携対象":5,
		"Redmine":5,
		"関連監理番号":5,
		"状態1":{val:0, list:["改修対応中","試験中","完了"], color:convert_RGB(255, 200, 170)},
		"発行日":get_date,
		"発行者":"CJS 金　忠傑",
		"品質管理単位":{ val:"一般計画工事-物品・受渡-物品", color:convert_RGB(255,255,220)},
		"検出処理ID":{val:0, list:["327610","327620"], color:convert_RGB(255, 230, 200)},
		"処理名":{val:0, list:["現調品_購入申請情報取込（バッチ）","現調品_検収情報取込（バッチ）"], color:convert_RGB(255, 230, 200)},
		"C/S":"S",
		"試験項目番号":"MCL_327610_現調品_購入申請情報取込（バッチ）_【1次1期】.xlsx",
		"検出AP":"S327610000.pc",
		"故障概要（題名）":{ val:"", color:convert_RGB(200,220,180)},
		"故障内容／状況（付加情報）":{ val:"", color:convert_RGB(200,220,180)},
		"故障検出日":get_date,
		"故障検出者":"CJS 金　忠傑",
		"解析状況":{ val:"", color:convert_RGB(200,200,200)},
		"故障原因":{ val:"", color:convert_RGB(200,230,200)},
		"故障AP":{val:0, list:["S327610000.pc","S327620000.pc"], color:convert_RGB(255, 230, 200)},
		"解析日":get_date,
		"解析者":"CJS 金　忠傑",
		"処置内容":{ val:"", color:convert_RGB(200,220,180)},
		"修正AP":"S327610000.pc",
		"修正APSVN Rev":{ val:"", color:convert_RGB(200,220,180)},
		"修正ドキュメント":{ val:"", color:convert_RGB(200,200,200)},
		"修正ドキュメントSVN Rev":{ val:"-", color:convert_RGB(200,200,200)},
		"設計修正要否":{val:2, list:["ED書・ID書","ID書のみ","対象なし"], color:convert_RGB(255, 230, 200)},
		"修正日":get_date,
		"修正者":"CJS 金　忠傑",
		"試験日":get_date,
		"確認者":"CJS 金　忠傑",
		"確認結果":"OK",
		"故障解決日":get_date,
		"発生箇所":"開発プログラム",
		"発生現象":{ val:"出力異常", color:convert_RGB(200,200,200)},
		"影響度":{ val:"中", color:convert_RGB(200,200,200)},
		"検出契機":{ val:"デバッグツール", color:convert_RGB(200,200,200)},
		"バグ種別":{ val:"作込みバグ", color:convert_RGB(200,200,200)},
		"バグ箇所":{ val:"処理", color:convert_RGB(200,200,200)},
		"混入工程":{ val:"C_コーディング", color:convert_RGB(200,200,200)},
		"混入現象":{ val:"コーディング誤り", color:convert_RGB(200,200,200)},
		"混入原因":{val:8, list:[
			"ユーザ要件検討不足",
			"実現方式検討不足",
			"設計条件確認不足",
			"修正確認不足",
			"連絡不足",
			"業務知識不足",
			"設計技術不足",
			"作業標準理解不足",
			"単純ミス",
			"その他"], color:convert_RGB(255, 230, 200)},
		"レビュー未発現原因":{val:7, list:[
			"レビュー未実施",
			"レビュー体制不備",
			"レビュー方法不備",
			"レビュー時間不足",
			"チェック項目漏れ",
			"指摘事項修正不備",
			"該当原因なし"], color:convert_RGB(255, 230, 200)},
		"摘出すべき試験工程":{val:0, list:["ＵＴ_単体試験","IT1"], color:convert_RGB(255, 230, 200)},
		"試験未発現原因":"-",
		"分類":{ val:"APバグ", color:convert_RGB(200,200,200)},
		"設計修正ドキュメント":{val:7, list:[""
			,"DIDA2-327610-FID03H.xlsm"
			,"DIDA2-327610-FIDA13_01.xlsx"
			,"DIDA2-327610-FIDA13_02.xlsx"
			,"DIDA2-327610-FIDA13_03.xlsx"
			,"DIDA2-327610-FIDA14_01.xlsx"
			,"DIDA2-327610-FIDA14_02.xlsx"
			,"DIDA2-327610-FIDA14_03.xlsx"
				], color:convert_RGB(255, 230, 200)},
		"設計修正依頼先":10,
		"設計修正依頼日":10,
		"設計修正ドキュメントSVN Rev":10,
		"設計書修正日":10,
		"設計書修正者":{val:2, list:["","CJS 官　開宣"], color:convert_RGB(255, 230, 200)},
		"設計書修正確認者":10,

		"横展開要否":{ val:"不要", color:convert_RGB(200,200,200)},
		"状態2":{ val:"完了", color:convert_RGB(200,200,200)},
		"横展開対象処理":5,
		"横展開観点":5,
		"横展開実施結果":5
	};
}









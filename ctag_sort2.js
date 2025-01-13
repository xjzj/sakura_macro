


var order=[];
var tag_map={};

var iLine = 0;
var iMax = GetLineCount(0);	//論理行数を取得
while (++iLine <= iMax){	//全行をループ
	var tmp=GetLineStr(iLine);
	var line=tmp.replace(/\r\n|\r|\n$/, "");
	var arr=line.split('\t');
	var key=arr[0] +"\t" +arr[1] +"\t" +arr[2];
	if(!tag_map[key]){
		tag_map[key]=[];
		order.push(key);
	}
	var list=tag_map[key];
	list.push( { arr:arr.slice(3) , flg:1 });
}

var list_txt="---------------------------------------\r\n";


var txt_del="";
var txt_chk=""
for( var i in order  ){
	var  key= order[i]
	var list=tag_map[key];
	if( list.length > 1 ){
		var idx=0;
		
		var check={};
		var chk_cnt=0;
		while(idx< list.length ){
			var arr=list[idx].arr;
			var chk_key=arr.join(":")
			if(!check[chk_key]){
				check[chk_key]=1;
				chk_cnt++;
				if( chk_cnt>1 ){
					txt_chk+=key + "\t" +  chk_key + "\r\n";
				}
			}
			else{
				list[idx].flg++;
				txt_del+=key + "\t" +  chk_key + "\r\n";
			}
			idx++;
		}
	}
}

list_txt+="--del●" + "\r\n";
list_txt+=txt_del;
list_txt+="--chk●" + "\r\n";
list_txt+=txt_chk;
AddTail(list_txt);


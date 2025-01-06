
var  uniq_map={};


var   all_cnt=GetLineCount(0);


var sTxt="-------------------------\r\n";
var  all_list=[];
var idx=0
while(idx<all_cnt){
	var line=GetLineStr(idx+1);
	// sTxt+="[" + line + "]" + "\r\n";
	var item=line.replace(/(^\s*)|(\s*$)/g, "");
	if(!uniq_map[item]){
		uniq_map[item]=10;
		all_list.push(item);
	}
	idx++;
}

all_list.sort();
// 数组排序
var sOut="-------------------------\r\n";
for(var i in all_list){
	sOut+=  all_list[i] + "\r\n";
}

// 反转数组
sOut+="-------------------------\r\n";
all_list.reverse();
for(var i in all_list){
	sOut+=  all_list[i] + "\r\n";
}

SelectAll();
InsText(sOut);
// InsText(sTxt);


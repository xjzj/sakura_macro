
var debug="";
var sql_word_list={
	'SELECT':{ nm:'SELECT',  end:{'FROM':1}, lf:{'SELECT':1,',':1}, indent:1  },
	'INSERT':{nm:'INSERT', omit:{'(':1}, end:{')':1}, lf:{'INTO':1, ',':1,  '(':1,  ')':1  } , indent:1},
	'UPDATE':{nm:'UPDATE',  end:'SET', lf:{'UPDATE':1, 'SET':1, ',':1,  '(':1,  ')':1  } , indent:1},
	// 'INTO':{omit:0,  end:'VALUES', lf:',' },
	'VALUES':{nm:'VALUES',  end:{')':1}, lf:{'INTO':1, ',':1,  '(':1,  ')':0  } , indent:1},
	
	'UNION':{nm:'UNION',  end:{';':1}, lf:{'ALL':1, ',':1,  '(':1,  ')':1  } , indent:1},
	'JOIN':{nm:'JOIN',  end:{'ON':1}, lf:{'ALL':1, ',':1,  '(':1,  ')':1  } , indent:1},

	'SET':{nm:'SET',  end:'WHERE', lf:{'SET':1, ',':1,  '(':1,  ')':1  } , indent:1},
	'FROM':{nm:'FROM', omit:0,  end:'WHERE', lf:{'FROM':1,',':1},  indent:1},
	'WHERE':{nm:'WHERE', omit:0,  end:'', lf:{'WHERE':1,',':1},  indent:1},
	'GROUP':{nm:'GROUP', omit:0,  end:'', lf:{'BY':1,',':1},  indent:1},
	'ORDER':{nm:'ORDER', omit:0,  end:'', lf:{'BY':1,',':1},  indent:1},
	'AND':{nm:'AND', omit:0,  end:'WHERE', lf:',' , indent:1},
	'OR':{nm:'OR', omit:0,  end:'WHERE', lf:',' , indent:1},
	'EXISTS':{nm:'EXISTS', omit:0,  end:'(', lf:',' , indent:1}
}

if( IsTextSelected ){
	var add_txt='';
	add_txt+='\n========================\n' ;
	var tmp_txt=GetSelectedString(0);
	var    one_line=to_one_line(tmp_txt);
	add_txt+=one_line;
	add_txt+='\n========================\n' ;
	var    sql_txt=on_line_sql_format(one_line);
	add_txt+=sql_txt;
	add_txt+='\n========================\n' ;
	// add_txt+=debug;
	AddTail( add_txt);
}



function to_one_line(str){
	var arr=str.split('');
	var txt="";
	var ret_str="";
	idx=0;
	var join_flg=0;
	while(idx < arr.length){
		var ch=arr[idx];
		if( ch == " " || ch =="\t" || ch =="\n" || ch =="\r"){
			if( txt!="" ){
				if(join_flg==0){
					ret_str+=' ';
				}
				ret_str+=txt;
				txt="";
				join_flg=0;
			}
		}
		else{
			if( ch=='(' && txt==""){
				join_flg=1;
			}
			txt+=ch;
		}
		idx++;
	}
	if( txt!=""){
		ret_str+=txt;
		txt="";
	}
	return ret_str;
}

function on_line_sql_format(str){
	var arr=str.split('');
	var field="";
	var ret_str="";
	idx=0;
	var indent_str='    ';
	var pre_status_flg=0;  //  0: 2:from  3:WHERE
	var status_flg=0;  //  0: 2:from  3:WHERE
	var func_flg=0;  //  0: 2:from  3:WHERE
	var pre_flg=0;  // 
	var lf_flg=0;  //  
	
	
	var indent_cnt=0;
	//--------------------------------
	var line_start_flg=1;  //
	function do_lf(s){
		if(line_start_flg==0){
			ret_str+="\n";
			line_start_flg=1;
			debug+="do lf" + "\r\n";
		}
	}
	function do_join(s){
		if( (s!='' &&  s!=' ') || ((line_start_flg==0) &&  s==' ') ){
			ret_str+=s;
			line_start_flg=0;
		}
	}
	//--------------------------------
	var stat_obj={end:{}, lf:{},indent:0};  //  
	var stat_pipe=[];  // 
	function push_stack(s){
		stat_pipe.push(stat_obj);
		stat_obj=s;
	}
	function pop_stack(){
		var ret_s=stat_obj;
		if(  stat_pipe.length>0 ){
			stat_obj=stat_pipe.pop();
		}
		// return ret_s;
	}
	//--------------------------------
	while(idx < arr.length){
		var do_lf_flg=false;

		var ch=arr[idx];
		if(is_word(ch)){
			field+=ch;
		}
		else{
			var chk_end=0;
			var field_upper=field.toUpperCase();
			if(stat_obj.end[field_upper]){
				do_lf();
				pop_stack();
			}
			var obj=sql_word_list[field_upper];
			if(obj){
				push_stack(obj);
				debug+='---------field[' + field + ']nm[' + stat_obj.nm +']do_lf_flg[' + do_lf_flg +']ch[' +    ch+ ']line_start_flg[' +    line_start_flg+ ']' + "\n";
				if(stat_obj.lf[field_upper]){
					do_lf_flg=1;
				}

				do_lf();
				ret_str+=field_upper;
				line_start_flg=0;
				// ret_str+="\n";
				// if( status_flg<4 ){
				// 	lf_flg=1;
				// }
				if(stat_obj.lf[ch]){
					do_lf_flg=1;
				}
			}
			else{
				if( field !=""){
					pre_flg=1;
				}
				/*
				if( field !="" && status_flg<4 ){
					pre_flg=1;
					ret_str+=indent_str + field;
				}
				if( ch == ','  && status_flg<3 ){
					lf_flg=1;
				}
				*/
				if(line_start_flg && (ch!=' ' ||field !="") ){
					ret_str+=repeat( indent_str, stat_obj.indent);
					debug+='repeat==field[' + field + ']nm[' + stat_obj.nm + ":"+ stat_obj.indent +']do_lf_flg[' + do_lf_flg + ']ch[' +    ch+ ']line_start_flg[' + line_start_flg+ ']' + "\n";
					/*
					try{
					}
					catch(e){
						debug+= 'err-------------------:' + 'field[' + field + ']nm[' + stat_obj.nm +']do_lf_flg[' + do_lf_flg + ']ch[' +    ch+ ']' + "\n";
					}
					*/
				}
				
				do_join(field);
				
				debug+='========field[' + field + ']nm[' + stat_obj.nm + ":"+ stat_obj.indent +']do_lf_flg[' + do_lf_flg + ']ch[' +    ch+ ']line_start_flg[' + line_start_flg+ ']' + "\n";
				if(( !stat_obj.omit ||!stat_obj.omit[ch] ) &&ch=='('  ){
					var obj={nm:field, omit:0,  end:{')':1}, lf:{')':1} , indent:0};
					push_stack(obj);
				}
				if(stat_obj.lf[ch]){
					do_lf_flg=1;
				}
				if( ch==')'  ){
					if(stat_obj.end[ch]){
						pop_stack();
					}
				}

			}
			if(stat_obj.lf[field_upper]){
				do_lf_flg=1;
			}
			debug+='field[' + field + ']nm[' + stat_obj.nm + ":"+ stat_obj.indent +']do_lf_flg[' + do_lf_flg + ']ch[' + ch+  ']line_start_flg[' + line_start_flg+']' + "\n";
			field='';
			do_join(ch);
			
			if(stat_obj.end[ch]){
				// debug+='--1----field[' + field + ']nm[' + stat_obj.nm + ":"+ stat_obj.cnt +']do_lf_flg[' + do_lf_flg + ']ch[' +    ch+ ']' + "\n";
				do_lf();
				pop_stack();
			}
			if( do_lf_flg){
				// debug+='--2----field[' + field + ']nm[' + stat_obj.nm + ":"+ stat_obj.cnt +']do_lf_flg[' + do_lf_flg + ']ch[' +    ch+ ']' + "\n";
				do_lf();
			}
		}
		idx++;
	}
	if( field !=""){
		ret_str+= field;
	}
	return ret_str;
}





function is_word(c){
	if( ('a' <= c && c <='z') || ( '0' <= c && c <='9' )  ||
	    ('A' <= c && c <='Z') || (  c =='_' ) ){
		    return true;
	    }
	return false;
}
function is_sign(c){
	var sign_list={
		',':1,
		'+':1,
		'(':1,
		')':1
	}
	if( sign_list[c]  ){
		return true;
	}
	return false;
}


function repeat(str, n) {
	if (typeof(n) == 'string') {
		n = Number(n);
	}
	if (n < 0) {
		n = 0;
	}
	var arr = new Array(n + 1);
	return arr.join(str); // "" + str + "" + str + ""  + str + "" ....
}

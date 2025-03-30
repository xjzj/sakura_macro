
var debug="";
var sql_word_list={
	'SELECT':{ nm:'SELECT',  end:{'FROM':1}, lf:{'SELECT':1,',':1}, indent:1  },
	'INSERT':{nm:'INSERT', omit:{'(':1}, end:{')':1}, lf:{'INTO':1, ',':1,  '(':1,  ')':1  } ,blf:{')':1  } , indent:1},
	'UPDATE':{nm:'UPDATE',  end:{'SET':1}, lf:{'UPDATE':1, 'SET':1, ',':1,  '(':1,  ')':1  } , indent:1},
	// 'INTO':{omit:0,  end:'VALUES', lf:',' },
	'VALUES':{nm:'VALUES',  end:{')':1}, lf:{'INTO':1, ',':1,  '(':1,  ')':0  } , blf:{')':1  }   , indent:1},
	'UNION':{nm:'UNION',  end:{'SELECT':1}, lf:{'ALL':1, ',':1,  '(':1,  ')':1  } , indent:1},
	'JOIN':{nm:'JOIN',  end:{'ON':1}, lf:{'ALL':1, ',':1,  '(':1,  ')':1  } , indent:1},
	'SET':{nm:'SET',  end:{'WHERE':1}, lf:{'SET':1, ',':1,  '(':1,  ')':1  } , indent:1},
	'FROM':{nm:'FROM',  nest:1, omit:0,  end:{'WHERE':1,'ORDER':1,'GROUP':1,';':1}, cut:{')':1},lf:{'FROM':1,',':1},  indent:1},
	'WHERE':{nm:'WHERE', omit:0,  end:{'AND':1,'OR':1,'UNION':1},cut:{')':1}, lf:{'WHERE':1,',':1},  indent:1},
	'GROUP':{nm:'GROUP', omit:0,  end:{},cut:{')':1}, lf:{'BY':1,',':1},  indent:1},
	'ORDER':{nm:'ORDER', omit:0,  end:'',cut:{')':1}, lf:{'BY':1,',':1},  indent:1},
	'AND':{nm:'AND', omit:0, nest:1,cut:{')':1}, end:{'AND':1,'OR':1,'GROUP':1}, lf:',' , indent:1},
	'OR':{nm:'OR', omit:0, nest:1,cut:{')':1}, end:{'AND':1,'OR':1,'GROUP':1}, lf:',' , indent:1},
	'EXISTS':{nm:'EXISTS', end:{')':1},  lf:{',':1,  '(':1,  ')':0  } ,blf:{')':1  }   ,  indent:1}
}

if( IsTextSelected ){
	var add_txt='';
	add_txt+='\n========================一行\n' ;
	var tmp_txt=GetSelectedString(0);
	var    one_line=to_one_line(tmp_txt);
	add_txt+=one_line;
	add_txt+='\n========================形成\n' ;
	var    sql_txt=on_line_sql_format(one_line);
	add_txt+=sql_txt;
	add_txt+='\n========================\n' ;
	add_txt+=debug;
	AddTail( add_txt);
}



function to_one_line(str){
	//---------------
	var ret_str="";
	var txt="";
	var join_flg=0;
	function set_txt(){
		if( txt!="" ){
			// debug+= txt + "\n";
			if(join_flg==0){
				ret_str+=' ';
			}
			ret_str+=txt;
			txt="";
			join_flg=0;
		}
	}
	//---------------
	var arr=str.split('');
	idx=0;
	while(idx < arr.length){
		var ch=arr[idx];
		if( ch == " " || ch =="\t" || ch =="\n" || ch =="\r"){
			set_txt()
		}
		else{
			if( ch=='(' && txt==""){
				join_flg=1;
			}
			txt+=ch;
		}
		idx++;
	}
	set_txt();
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
	var  stak_level=0;
	var stat_obj={nm:'-'  ,end:{}, lf:{},indent:0, cnt:0, level:0 };  //  
	var stat_pipe=[];  // 
	function push_stack(s){
		var  indent=s.indent+stat_obj.indent;
		stak_level=stat_obj.level;
		stat_pipe.push(stat_obj);
		stak_level++;
		var obj={ nm:s.nm, end:s.end, cut:s.cut, lf:s.lf, blf:s.blf,  omit:s.omit,  nest:s.nest,  indent:indent, cnt:0, level:stak_level };
		stat_obj=obj;
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
				// debug+='pop_stack start   field[' + field + ']nm[' + stat_obj.nm + ":"+ stat_obj.level + ":"+ + stat_obj.indent +']do_lf_flg[' + do_lf_flg + ']ch[' + ch+  ']line_start_flg[' + line_start_flg+']' + "\n";
				pop_stack();
				// debug+='pop_stack  end  field[' + field + ']nm[' + stat_obj.nm + ":"+ stat_obj.level + ":"+ + stat_obj.indent +']do_lf_flg[' + do_lf_flg + ']ch[' + ch+  ']line_start_flg[' + line_start_flg+']' + "\n";

			}
			var obj=sql_word_list[field_upper];
			if(obj){
				// obj.indent+=stat_obj.indent;
				// debug+='push_stack start   field[' + field + ']nm[' + stat_obj.nm + ":"+ stat_obj.level + ":"+ + stat_obj.indent +']do_lf_flg[' + do_lf_flg + ']ch[' + ch+  ']line_start_flg[' + line_start_flg+']' + "\n";
				push_stack(obj);
				// debug+='push_stack  end  field[' + field + ']nm[' + stat_obj.nm + ":"+ stat_obj.level + ":"+ + stat_obj.indent +']do_lf_flg[' + do_lf_flg + ']ch[' + ch+  ']line_start_flg[' + line_start_flg+']' + "\n";

				// debug+='---------field[' + field + ']nm[' + stat_obj.nm +']do_lf_flg[' + do_lf_flg +']ch[' +    ch+ ']line_start_flg[' +    line_start_flg+ ']' + "\n";
				if(stat_obj.lf[field_upper]){
					do_lf_flg=1;
				}

				do_lf();
				ret_str+=repeat( indent_str, (stat_obj.indent-1));
				do_join(field_upper);
				// ret_str+=field_upper;
				// line_start_flg=0;
				
				// ret_str+="\n";
				// if( status_flg<4 ){
				// 	lf_flg=1;
				// }
				//  if(stat_obj.lf[ch]){
				//  	do_lf_flg=1;
				//  }
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
					// debug+='repeat==field[' + field + ']nm[' + stat_obj.nm + ":"+ stat_obj.indent +']do_lf_flg[' + do_lf_flg + ']ch[' +    ch+ ']line_start_flg[' + line_start_flg+ ']' + "\n";
					/*
					try{
					}
					catch(e){
						debug+= 'err-------------------:' + 'field[' + field + ']nm[' + stat_obj.nm +']do_lf_flg[' + do_lf_flg + ']ch[' +    ch+ ']' + "\n";
					}
					*/
				}
				
				do_join(field);
				
				
				/*
				// debug+='========field[' + field + ']nm[' + stat_obj.nm + ":"+ stat_obj.indent +']do_lf_flg[' + do_lf_flg + ']ch[' +    ch+ ']line_start_flg[' + line_start_flg+ ']' + "\n";
				if(( !stat_obj.omit ||!stat_obj.omit[ch] ) &&ch=='('  ){
					var obj;
					if( field=='' ){
						obj={nm:field, omit:0,  end:{')':1}, lf:{')':1} , indent:(stat_obj.indent+1)};
					}
					else{
						obj={nm:field, omit:0,  end:{')':1}, lf:{')':0} , indent:(stat_obj.indent+1)};
					}
					push_stack(obj);
				}
				*/
				
				// if( ch==')'  ){
				// 	if(stat_obj.end[ch]){
				// 		pop_stack();
				// 	}
				// }

			}
			
			
			if( (	(stat_obj.nest&& (stat_obj.cnt==0) )||
				((stat_obj.cnt>0)&&( !stat_obj.omit ||!stat_obj.omit[ch] ))) 
				&&ch=='('  ){
				var obj;
				var nm_field=field;
				var nest_flg=0;
				if(stat_obj.nest&& (stat_obj.cnt==0) ){
					nm_field=field + ":NEST";
					nest_flg=1;
				}
				if(  field== ''){
					nest_flg=1;
				}
				if( nest_flg){
					obj={nm:nm_field, omit:0,  end:{')':1}, blf:{')':1}, lf:{')':1,'(':1} , indent:(stat_obj.indent+1)};
				}
				else{
					obj={nm:nm_field, omit:0,  end:{')':1},   lf:{')':0,'(':0} , indent:(stat_obj.indent+1)};
				}
				push_stack(obj);
			}

			if(stat_obj.blf){
				if(stat_obj.blf[field_upper]){
					do_lf();
				}
			}

			if(stat_obj.lf[field_upper]){
				do_lf_flg=1;
			}
			debug+='field[' + field + ']nm[' + stat_obj.nm + ":"+ stat_obj.level + ":"+ + stat_obj.indent +']do_lf_flg[' + do_lf_flg + ']ch[' + ch+  ']line_start_flg[' + line_start_flg+']' + "\n";
	
			if(stat_obj.end[ch]||( stat_obj.cut && stat_obj.cut[ch])){
				debug+='end or cut  field[' + field + ']nm[' + stat_obj.nm + ":"+ stat_obj.level + ":"+ + stat_obj.indent +']do_lf_flg[' + do_lf_flg + ']ch[' + ch+  ']line_start_flg[' + line_start_flg+']' + "\n";
				// do_lf();
				while(1){
					if(stat_obj.lf[ch]){
						do_lf_flg=1;
						// debug+='--2----field[' + field + ']nm[' + stat_obj.nm + ":"+ stat_obj.nm + ":"+ stat_obj.cnt +']do_lf_flg[' + do_lf_flg + ']ch[' +    ch+ ']' + "\n";
					}
					if(stat_obj.end[ch]){
						debug+='end field[' + field + ']nm[' + stat_obj.nm + ":"+ stat_obj.level + ":"+ + stat_obj.indent +']do_lf_flg[' + do_lf_flg + ']ch[' + ch+  ']line_start_flg[' + line_start_flg+']' + "\n";
						if(stat_obj.lf[ch]){
							do_lf_flg=1;
						}
						if(stat_obj.blf){
							if(stat_obj.blf[ch]){
								do_lf();
							}
							if(line_start_flg ){
								ret_str+=repeat( indent_str, stat_obj.indent-1);
							}
						}
						pop_stack();
						break;
					}
					else if( stat_obj.cut && stat_obj.cut[ch] ){
						debug+='cut field[' + field + ']nm[' + stat_obj.nm + ":"+ stat_obj.level + ":"+ + stat_obj.indent +']do_lf_flg[' + do_lf_flg + ']ch[' + ch+  ']line_start_flg[' + line_start_flg+']' + "\n";
						pop_stack();
					}
					else{
						break;
					}
				}
			}
			else{
				if(stat_obj.lf[ch]){
					do_lf_flg=1;
				}
				if(stat_obj.blf){
					if(stat_obj.blf[ch]){
						do_lf();
					}
					if(line_start_flg ){
						ret_str+=repeat( indent_str, stat_obj.indent-1);
					}
				}
			}
			field='';
			
			do_join(ch);
			if( do_lf_flg){
				// debug+='--2----field[' + field + ']nm[' + stat_obj.nm + ":"+ stat_obj.cnt +']do_lf_flg[' + do_lf_flg + ']ch[' +    ch+ ']' + "\n";
				do_lf();
			}
			stat_obj.cnt++;
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

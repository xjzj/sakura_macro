
var Cmd;
//-----------------------------------------------------------
function inpath(file) 
{
	var ForReading = 1, ForWriting = 2;
	var wsh = new ActiveXObject("wscript.shell");
	var env = wsh.Environment("SYSTEM");
	var path = env.item("SAKURA_SCRIPT") + "/";

	var FileOpener = new ActiveXObject("Scripting.FileSystemObject");
	var FilePointer = FileOpener.OpenTextFile(path + file, ForReading, true);
	Cmd = FilePointer.ReadAll();
}

//-----------------------------------------------------------
inpath("inc.js");
eval(Cmd);



//  AOK's JavaScript Library
//  IE6.0(Win2000/XP) ��ư���ǧ���Ƥ��ޤ���	Last Update: Nov.20, 2006
//  ADODB.Stream ��Ȥä��ե�������ɤ߽�
/* StreamTypeEnum Values
*/
var adTypeBinary = 1;
var adTypeText = 2;

/* LineSeparatorEnum Values
*/
var adLF = 10;
var adCR = 13;
var adCRLF = -1;

/* StreamWriteEnum Values
*/
var adWriteChar = 0;
var adWriteLine = 1;

/* SaveOptionsEnum Values
*/
var adSaveCreateNotExist = 1;
var adSaveCreateOverWrite = 2;

/* StreamReadEnum Values
*/
var adReadAll = -1;
var adReadLine = -2;

/* charset ���ͤ���:
*  _autodetect, euc-jp, iso-2022-jp, shift_jis, unicode, utf-8,...
*/

/* filename: �ɤ߹���ե�����Υѥ�
* charset:  ʸ��������
* �����:   ʸ����
*/
function adoLoadText(filename, charset) {
	var stream, text;
	stream = new ActiveXObject("ADODB.Stream");
	stream.type = adTypeText;
	stream.charset = charset;
	stream.open();
	stream.loadFromFile(filename);
	text = stream.readText(adReadAll);
	stream.close();
	return text;
}

/* filename: �ɤ߹���ե�����Υѥ�
* charset:  ʸ��������
* �����:   ��ñ�̤�ʸ���������
*/
function adoLoadLinesOfText(filename, charset) {
	var stream;
	var lines = new Array();
	stream = new ActiveXObject("ADODB.Stream");
	stream.type = adTypeText;
	stream.charset = charset;
	stream.open();
	stream.loadFromFile(filename);
	while ( !stream.EOS) {
		lines.push(stream.readText(adReadLine));
	}
	stream.close();
	return lines;
}

/* filename: �񤭽Ф��ե�����Υѥ�
* charset:  ʸ��������
*/
function adoSaveText(filename, text, charset) {
	var stream;
	stream = new ActiveXObject("ADODB.Stream");
	stream.type = adTypeText;
	stream.charset = charset;
	stream.open();
	stream.writeText(text);
	stream.saveToFile(filename, adSaveCreateOverWrite);
	stream.close();
}




/*
	Property get isUp As Boolean
    Sub die(msg)
	Sub Caption(msg)
	Function alert(msg)
	Function Message(msg As String)
	Function MakeStr(va,  ascii As Boolean = True)
	Function MakeUnk(va, size)
	Property get LoadedFile As String
	Sub t(data As String)
	Sub ClearLog()
	Function PatchString(va, str,  isUnicode = False)
	Function PatchByte(va, newVal)
	Function intToHex(x)
	Function GetAsm(va)
	Function InstSize(offset)
	Function XRefsTo(offset)
	Function XRefsFrom(offset)
	Function GetName(offset)
	Function FunctionName(functionIndex)
	Function HideBlock(offset, leng)
	Function ShowBlock(offset, leng)
	Sub Setname(offset, name)
	Sub AddComment(offset, comment)
	Function GetComment(offset)
	Sub AddCodeXRef(offset, tova)
	Sub AddDataXRef(offset, tova)
	Sub DelCodeXRef(offset, tova)
	Sub DelDataXRef(offset, tova)
	Function FuncVAByName(name)
	Function RenameFunc(oldname, newName) As Boolean
	Function Find(startea, endea, hexstr) 
	Function Decompile(va) As String
	Function Jump(va As Long)
	Function JumpRVA(rva As Long)
	Function refresh()
	Function Undefine(offset)
	Function ShowEA(offset)
	Function HideEA(offset)
	Sub RemoveName(offset)
	Sub MakeCode(offset)
	Function FuncIndexFromVA(va)
	Function NextEA(va)
	Function PrevEA(va)
	Function funcCount() As Long
	Function NumFuncs() As Long
	Function FunctionStart(functionIndex)
	Function FunctionEnd(functionIndex)
	Function ReadByte(va)
	Function OriginalByte(va)
	Function ImageBase() As Long
	Function ScreenEA() As Long
	Function EnableIDADebugMessages( enabled)
	Function QuickCall(msg As Long,  arg1 as Long) As Long
#	'Sub AddProgramComment(cmt)
#	' Function ScreenEA()
#	'Function GetAsmBlock(start, leng)
#	'Function GetBytes(start, leng)
#	'Sub AnalyzeArea(startat, endat)
	Function AskValue( prompt,  defVal) As String
	Sub Exec(cmd)
	Function ReadFile(filename) As Variant
	Sub WriteFile(path As String, it As Variant)
	Sub AppendFile(path, it)
	Function FileExists(path As String) As Boolean
	Function DeleteFile(fpath As String) As Boolean
	Function getClipboard()
	Function setClipboard(x)
	Function OpenFileDialog() As String
	Function SaveFileDialog() As String
	Function BenchMark() As Long
#	'Function refListToArray(x) As Long() 
#	'Function InstSize(offset)
*/

function idaClass(){

	this.hInst = 0
	
	this.die = function(msg){
		return resolver('ida.die', arguments.length,0, msg);
	}

	this.Caption = function(msg){
		return resolver('ida.Caption', arguments.length,0, msg);
	}

	this.alert = function(msg){
		return resolver('ida.alert', arguments.length,0, msg);
	}

	this.Message = function(msg){
		return resolver('ida.Message', arguments.length,0, msg);
	}

	this.MakeStr = function(va, ascii){
		return resolver('ida.MakeStr', arguments.length,0, va, ascii);
	}

	this.MakeUnk = function(va, size){
		return resolver('ida.MakeUnk', arguments.length,0, va, size);
	}

	this.t = function(data){
		return resolver('ida.t', arguments.length,0, data);
	}

	this.ClearLog = function(){
		return resolver('ida.ClearLog', arguments.length,0);
	}

	this.PatchString = function(va, str, isUnicode){
		return resolver('ida.PatchString', arguments.length,0, va, str, isUnicode);
	}

	this.PatchByte = function(va, newVal){
		return resolver('ida.PatchByte', arguments.length,0, va, newVal);
	}

	this.intToHex = function(x){
		return resolver('ida.intToHex', arguments.length,0, x);
	}

	this.GetAsm = function(va){
		return resolver('ida.GetAsm', arguments.length,0, va);
	}

	this.InstSize = function(offset){
		return resolver('ida.InstSize', arguments.length,0, offset);
	}

	this.XRefsTo = function(offset){
		return resolver('ida.XRefsTo', arguments.length,0, offset);
	}

	this.XRefsFrom = function(offset){
		return resolver('ida.XRefsFrom', arguments.length,0, offset);
	}

	this.GetName = function(offset){
		return resolver('ida.GetName', arguments.length,0, offset);
	}

	this.FunctionName = function(functionIndex){
		return resolver('ida.FunctionName', arguments.length,0, functionIndex);
	}

	this.HideBlock = function(offset, leng){
		return resolver('ida.HideBlock', arguments.length,0, offset, leng);
	}

	this.ShowBlock = function(offset, leng){
		return resolver('ida.ShowBlock', arguments.length,0, offset, leng);
	}

	this.Setname = function(offset, name){
		return resolver('ida.Setname', arguments.length,0, offset, name);
	}

	this.AddComment = function(offset, comment){
		return resolver('ida.AddComment', arguments.length,0, offset, comment);
	}

	this.GetComment = function(offset){
		return resolver('ida.GetComment', arguments.length,0, offset);
	}

	this.AddCodeXRef = function(offset, tova){
		return resolver('ida.AddCodeXRef', arguments.length,0, offset, tova);
	}

	this.AddDataXRef = function(offset, tova){
		return resolver('ida.AddDataXRef', arguments.length,0, offset, tova);
	}

	this.DelCodeXRef = function(offset, tova){
		return resolver('ida.DelCodeXRef', arguments.length,0, offset, tova);
	}

	this.DelDataXRef = function(offset, tova){
		return resolver('ida.DelDataXRef', arguments.length,0, offset, tova);
	}

	this.FuncVAByName = function(name){
		return resolver('ida.FuncVAByName', arguments.length,0, name);
	}

	this.RenameFunc = function(oldname, newName){
		return resolver('ida.RenameFunc', arguments.length,0, oldname, newName);
	}

	this.Find = function(startea, endea, hexstr){
		return resolver('ida.Find', arguments.length,0, startea, endea, hexstr);
	}

	this.Decompile = function(va){
		return resolver('ida.Decompile', arguments.length,0, va);
	}

	this.Jump = function(va){
		return resolver('ida.Jump', arguments.length,0, va);
	}

	this.JumpRVA = function(rva){
		return resolver('ida.JumpRVA', arguments.length,0, rva);
	}

	this.refresh = function(){
		return resolver('ida.refresh', arguments.length,0);
	}

	this.Undefine = function(offset){
		return resolver('ida.Undefine', arguments.length,0, offset);
	}

	this.ShowEA = function(offset){
		return resolver('ida.ShowEA', arguments.length,0, offset);
	}

	this.HideEA = function(offset){
		return resolver('ida.HideEA', arguments.length,0, offset);
	}

	this.RemoveName = function(offset){
		return resolver('ida.RemoveName', arguments.length,0, offset);
	}

	this.MakeCode = function(offset){
		return resolver('ida.MakeCode', arguments.length,0, offset);
	}

	this.FuncIndexFromVA = function(va){
		return resolver('ida.FuncIndexFromVA', arguments.length,0, va);
	}

	this.NextEA = function(va){
		return resolver('ida.NextEA', arguments.length,0, va);
	}

	this.PrevEA = function(va){
		return resolver('ida.PrevEA', arguments.length,0, va);
	}

	this.funcCount = function(){
		return resolver('ida.funcCount', arguments.length,0);
	}

	this.NumFuncs = function(){
		return resolver('ida.NumFuncs', arguments.length,0);
	}

	this.FunctionStart = function(functionIndex){
		return resolver('ida.FunctionStart', arguments.length,0, functionIndex);
	}

	this.FunctionEnd = function(functionIndex){
		return resolver('ida.FunctionEnd', arguments.length,0, functionIndex);
	}

	this.ReadByte = function(va){
		return resolver('ida.ReadByte', arguments.length,0, va);
	}

	this.OriginalByte = function(va){
		return resolver('ida.OriginalByte', arguments.length,0, va);
	}

	this.ImageBase = function(){
		return resolver('ida.ImageBase', arguments.length,0);
	}

	this.ScreenEA = function(){
		return resolver('ida.ScreenEA', arguments.length,0);
	}

	this.EnableIDADebugMessages = function(enabled){
		return resolver('ida.EnableIDADebugMessages', arguments.length,0, enabled);
	}

	this.QuickCall = function(msg, arg1){
		return resolver('ida.QuickCall', arguments.length,0, msg, arg1);
	}

	this.AskValue = function(prompt, defVal){
		return resolver('ida.AskValue', arguments.length,0, prompt, defVal);
	}

	this.Exec = function(cmd){
		return resolver('ida.Exec', arguments.length,0, cmd);
	}

	this.ReadFile = function(filename){
		return resolver('ida.ReadFile', arguments.length,0, filename);
	}

	this.WriteFile = function(path, it){
		return resolver('ida.WriteFile', arguments.length,0, path, it);
	}

	this.AppendFile = function(path, it){
		return resolver('ida.AppendFile', arguments.length,0, path, it);
	}

	this.FileExists = function(path){
		return resolver('ida.FileExists', arguments.length,0, path);
	}

	this.DeleteFile = function(fpath){
		return resolver('ida.DeleteFile', arguments.length,0, fpath);
	}

	this.getClipboard = function(){
		return resolver('ida.getClipboard', arguments.length,0);
	}

	this.setClipboard = function(x){
		return resolver('ida.setClipboard', arguments.length,0, x);
	}

	this.OpenFileDialog = function(){
		return resolver('ida.OpenFileDialog', arguments.length,0);
	}

	this.SaveFileDialog = function(){
		return resolver('ida.SaveFileDialog', arguments.length,0);
	}

	this.BenchMark = function(){
		return resolver('ida.BenchMark', arguments.length,0);
	}

}

idaClass.prototype = {
	get isUp(){
		return resolver('ida.isUp.get', 0, this.hInst);
	},

	/*set Enabled(val){
		return resolver('list.Enabled.let', 1, this.hInst, val);
	},*/

	get LoadedFile(){
		return resolver('ida.LoadedFile.get', 0, this.hInst);
	}
}

var ida = new idaClass()


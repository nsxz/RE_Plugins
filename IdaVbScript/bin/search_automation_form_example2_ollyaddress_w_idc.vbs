
'Top level objects
 'txtSearch = search textbox
 'txtComment = comment textrbox
 'lv is listview     (.additem, and .clear most used functions)
 '  lv.listitems(x).text is address in hex
 '  lv.listitems(x).subitems(1) = disasm
 '  lv.listitems(x).subitems(2) = comment
 'pb is progressbar  (.min, .max and .value are most used)
 'list1 is vb list box (.additem , .clear most used)
 'fso = clsFileSystem
 'cmndlg = clsCmnDlg
 'clipboard = vb clipboard object (.clear, .settext most used)

'Specific Form Elements
 'form.text1 = search textbox
 'form.text2 = comment textbox
 'form.command1_click = click search button proc
 'form.command2_click = add comment button click proc
 'form.pb  is a progressbar 

'Form functions
 'form.DoSearch(Optional parameter = "") As Long  'parameter = search string, retval = found count
 'form.AddComments(Optional comment = "") 'comment = text to add for each list item found
 'form.GetAsmCode(offset) As String
 'form.InstructionLength(offset) As Long
 'form.Set_Comment(offset, comm As String) as Long   (1 = success, 0=fail)
 'form.ScanForInstruction(offset As Long, find_inst As String, scan_x_lines As Long) As Long 'returns ea
 'form.AddXRef(ref_to As Long, ref_from As Long)
 'form.SelAll() select alls list tiems
 'form.Setname(offset As Long, name As String)
 'form.FunctionatVA(va as long) as cfunction  - cfunction properties: .startea, .endea, .index, .name, .length
 'Function DoUndefine(offset As Long)

function main()

	t = cmndlg.OpenDialog(4) 'all files
	if fso.fileexists(t) then t = fso.readfile(t)
	
	'assumes file is olly long address dump list
	'00AFB0C4  7C809CAD  kernel32.MultiByteToWideChar
	'00AFB0C8  7C80A0C7  kernel32.WideCharToMultiByte
	'00AFB0CC  7C8221CF  kernel32.GetTempPathA
	'00AFB0D0  7C810626  kernel32.CreateRemoteThread

	t = split(t,vbcrlf)
	
	pb.value = 0
	pb.max = ubound(t)+2
    idc = ""
    
	for each x in t
		pb.value = pb.value + 1
		if len(x) > 0 then 
			
			y = split(x,"  ")     
			if ubound(y) > 1 then 
				 my_hash = ucase(trim(y(0)))
				 my_name = trim(y(2)) 
				 my_hash = replace(my_hash, "000BA", "0BA") 
				 
				if len(my_hash) > 0 and len(my_name) > 0 then 
					if form.DoSearch("ds:" & my_hash) > 0  then   'call dword ptr ds:0AFB008
						form.AddComments my_name						
						for each li in lv.listitems 'for each search result
							va = clng("&h" & li.text)  
							if va <> 0 then 
								form.AddXRef va, clng("&h" & my_hash)
								form.DoUndefine clng("&h" & my_hash)              'addxref may auto analyze and convert addr to code...
					    		form.Setname clng("&h" & my_hash), cstr(my_name)  'in which case we could not setname...
								form.Set_Comment va,  cstr(my_name)               'this is probably redundant I forget doesnt hurt..
								idc = idc & "MakeComm( 0x" & hex(va) & ", """ & my_name & """);" & vbcrlf  
							end if 
						next
	
						list1.additem my_name & " " & lv.listitems.count & " items found"
					end if
				end if 
			end if 
			
		end if
	next

	pb.value = 0
	
	clipboard.clear
	clipboard.settext idc
	msgbox "Done! (reusable idc copied to clipboard)"

end function

 
 

 
	


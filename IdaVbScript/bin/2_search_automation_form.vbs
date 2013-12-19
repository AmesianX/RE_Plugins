
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



'input file format sample
' CcCanIWrite =  5f650566
' CcCopyRead =  4553982b


'one based these are to be created by this script at their related array index = va 
'(make a new segment base at 0 or modify logic below)

api = ",DbgPrint,ExAcquireFastMutexUnsafe ,ExAllocatePool ,ExFreePool ,ExReleaseFastMutexUnsafe ,FsRtlAllocatePool ,FsRtlFastUnlockAll ,FsRtlInitializeFileLock ,FsRtlIsNameInExpression ,FsRtlProcessFileLock ,IoAllocateIrp ,IoAllocateMdl ,IoCompleteRequest ,IoCreateDevice ,IoCreateDriver ,IoGetCurrentProcess ,IoGetRequestorProcess ,KeDelayExecutionThread ,KeGetCurrentThread ,KeInitializeApc ,KeInitializeEvent ,KeInsertQueueApc ,KeQuerySystemTime ,KeSetEvent ,KeStackAttachProcess ,KeWaitForSingleObject ,MmMapLockedPagesSpecifyCache ,MmProbeAndLockPages ,NtLockFile ,ObDereferenceObject ,ObMakeTemporaryObject ,ObReferenceObjectByHandle ,ObfDereferenceObject ,ObfReferenceObject ,PsGetCurrentProcessId ,PsLookupProcessByProcessId ,PsLookupThreadByThreadId ,PsSetLoadImageNotifyRoutine ,RtlEqualUnicodeString ,RtlFillMemoryUlong ,RtlImageNtHeader ,RtlInitUnicodeString ,RtlRandom ,ZwClose ,ZwCreateEvent ,ZwCreateFile ,ZwCreateSection ,ZwDeviceIoControlFile ,ZwFlushVirtualMemory ,ZwFreeVirtualMemory ,ZwFsControlFile ,ZwMapViewOfSection ,ZwOpenFile ,ZwOpenKey ,ZwOpenProcess ,ZwQueryInformationFile,ZwQueryInformationProcess ,ZwQuerySystemInformation ,ZwQueryValueKey ,ZwQueryVolumeInformationFile ,ZwReadFile ,ZwSetInformationFile ,ZwSetValueKey ,ZwUnmapViewOfSection ,ZwWriteFile ,_alldiv ,_allmul ,_snprintf ,_snwprintf ,_strlwr ,_strnicmp ,strchr ,strcmp ,strlen ,strncpy,strrchr,strstr ,wcslen ,wcsncmp,wcsncpy"

imports = split(api,",")

function main()

	setNames()

	t = cmndlg.OpenDialog(4) 'all files
	if fso.fileexists(t) then t = fso.readfile(t)
	t = split(t,vbcrlf)
	
	pb.value = 0
	pb.max = ubound(t)+2

	for each x in t
		pb.value = pb.value + 1
		if len(x) > 0 and instr(x,"=") > 0 then 
			
			y = split(x,"=")     'ex: CcCanIWrite =  5f650566
			my_name = trim(y(0)) 'name
			my_hash = trim(y(1)) 'hash
			
			if form.DoSearch(my_hash) > 0  then 
				form.AddComments my_name
				addreferences(form.text2)
				list1.additem my_name & " " & lv.listitems.count & " items found"
			end if
			
		end if
	next

	pb.value = 0
	msgbox "Done!"

end function


function addreferences(apiName)

	for each li in lv.listitems

		next_inst = form.ScanForInstruction( clng("&h" & li.text) , "call eax", 50)
		
		if next_inst <> 0 then 
			import_addr = getImportedPad(apiName)
			if import_addr <> 0 then 
				form.AddXRef next_inst, import_addr
				form.Set_Comment next_inst, apiName
			end if 
		end if 
	next

end function

function setNames()

	for i=1 to ubound(imports)
		form.Setname i, trim(imports(i))
	next

end function


function getImportedPad(apiName)

	for i=1 to ubound(imports)
	   if trim(lcase(imports(i))) = trim(lcase(apiname)) then 
			getImportedPad = i
			exit for
	   end if 
    next

end function

	


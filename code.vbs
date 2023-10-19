' VBSCRIPT CODE FOR WIMEDITOR 
' === Start Of File ===
' = VISUAL =
' === Initialize Application Windows ===
Sub appInit()
	Const appW = 432
	Const appH = 464
	window.resizeTo appW, appH
	window.moveTo screen.width/2 - appW/2, screen.height/2 - appH/2
End Sub

' === All Button Mouse Over And Leave Event ===
Sub bttMove(bttId)
	btId = Left(bttId, Len(bttId) - 1)
	btKw = LCase(Mid(btId,3,3))
	document.getElementById(btId).src="./imgs/" & btKw & Right(bttId,1) & ".jpg"
	If Right(bttId,1)="L" Then 
		document.title="WINDOWS IMAGE EDITOR"
	Else
		Select Case btKw
			Case "cap"
				document.title="WINDOWS IMAGE EDITOR [CAPTURE IMAGE]"
			Case "app"
				document.title="WINDOWS IMAGE EDITOR [APPLY IMAGE]"
			Case "mou"
				document.title="WINDOWS IMAGE EDITOR [MOUNT IMAGE]"
			Case "unm"
				document.title="WINDOWS IMAGE EDITOR [UNMOUNT IMAGE]"
			Case "add"
				document.title="WINDOWS IMAGE EDITOR [APPEND IMAGE]"
			Case "del"
				document.title="WINDOWS IMAGE EDITOR [DELETE IMAGE]"
			Case "exp"
				document.title="WINDOWS IMAGE EDITOR [EXPORT IMAGE]"
			Case "spl"
				document.title="WINDOWS IMAGE EDITOR [SPLIT IMAGE]"
			Case "inf"
				document.title="WINDOWS IMAGE EDITOR [IMAGE INFORMATION]"
		End Select
	End If
End Sub
' = END VISUAL =

' = SUPPORT FUNCTION =
' === Open Folder Browse Dialog and get selected path ===
Function getFolder(title, isPC, newFolder, showFiles)
	Set oShell = CreateObject("Shell.Application")
	Dim vPC, vOpt
	If isPC Then vPC=0 Else vPC=17
	vOpt = 0
	If (Not newFolder) Then vOpt = vOpt + 512
	If showFiles Then vOpt = vOpt + 16384
	Set dDlg = oShell.BrowseForFolder(0, title, vOpt, vPC)
	If (Not dDlg Is Nothing) Then
		tPath=dDlg.items.item.path
		If Left(tPath,2)="::" Then
			getFolder=""
		Else
			getFolder=tPath
		End If
	Else
		getFolder=""
	End If
End Function

' === Open File Browse Dialog and get selected file path ===
Function getFile(fTitle,fFilter)
	Set FSO = CreateObject("Scripting.FileSystemObject")
	sIniDir = "::{20D04FE0-3AEA-1069-A2D8-08002B30309D}"'FSO.GetSpecialFolder(Desktop)
	Set oShell = CreateObject("WScript.Shell").Exec("mshta.exe ""about:<object id=d classid=clsid:3050f4e1-98b5-11cf-bb82-00aa00bdce0b></object><script>moveTo(0,-9999);eval(new ActiveXObject('Scripting.FileSystemObject').GetStandardStream(0).Read("&Len(sInitDir)+Len(fFilter)+Len(fTitle)+41&"));function window.onload(){var p=/[^\16]*/;new ActiveXObject('Scripting.FileSystemObject').GetStandardStream(1).Write(p.exec(d.object.savefiledlg(iniDir,null,filter,title)));close();}</script><hta:application showintaskbar=no />""")
	oShell.StdIn.Write "var iniDir='" & sInitDir & "';var filter='" & fFilter & "';var title='" & fTitle & "';"
	res=oShell.StdOut.ReadAll
	If res<>"" Then
		fInd=Instr(1,res,Chr(0))
		getFile = Left(res,fInd-1)
	Else
		getFile=""
	End If
End Function

' === Run a command ===
Function exeRun(command)
	CreateObject("WScript.Shell").Run(command)
End Function

' === Check file name or folder name is available for Windows
Function checkName(name)
	checkName=((InStr(1, name, Chr(&H2F)) + InStr(1, name, Chr(&H5C)) + InStr(1, name, Chr(&H3A)) +_
				InStr(1, name, Chr(&H2A)) + InStr(1, name, Chr(&H3F)) + InStr(1, name, Chr(&H22)) +_
				InStr(1, name, Chr(&H3C)) + InStr(1, name, Chr(&H3E)) + InStr(1, name, Chr(&H7C)))=0)
End Function

' === Run command line and read output
Function readCmd(cmdLine)
	Set objShell = CreateObject("WScript.Shell")
	Set objExecObject = objShell.Exec("cmd /c " & cmdLine)
	strText = ""
	Do While Not objExecObject.StdOut.AtEndOfStream
		strText = strText & chr(13) & objExecObject.StdOut.ReadLine()
	Loop
	readCmd=strText
End Function

' === Read all images in image file ===
Function readWim(wimPath)
	strCmd=readCmd("dism /get-imageinfo /imagefile:" & Chr(34) & wimPath & Chr(34))
	indStr=1
	Dim arrImg()
	ReDim arrImg(0)
	stArr=False
	Do While (InStr(indStr, strCmd, "Index :")>0)
		sInd=InStr(indStr, strCmd, "Index :")
		eInd=InStr(sInd, strCmd, Chr(13))
		sName=InStr(eInd, strCmd, "Name :")
		eName=InStr(sName, strCmd, Chr(13))
		sDesc=InStr(eName, strCmd, "Description :")
		eDesc=InStr(sDesc, strCmd, Chr(13))
		sSize=InStr(eDesc, strCmd, "Size :")
		eSize=InStr(sSize, strCmd, "bytes")
		intInd=Trim(Mid(strCmd, sInd+7, eInd-sInd-7))
		strName=Trim(Mid(strCmd, sName+6, eName-sName-6))
		strDesc=Trim(Mid(strCmd, sDesc+13, eDesc-sDesc-13))
		lngSize=Trim(Replace(Mid(strCmd, sSize+6, eSize-sSize-6),",",""))
		If stArr Then ReDim Preserve arrImg(UBound(arrImg)+1)
		arrImg(UBound(arrImg))=intInd & "|" & strName & "|" & strDesc & "|" & lngSize
		stArr=True
		indStr=eSize
	Loop
	readWim=arrImg
End Function

' === Convert byte to Megabyte
Function convertByte(vByte,vDigit)
	Dim resStr
	If (vByte<921) Then
		resStr=vByte & " bytes"
	ElseIf (vByte<943718) Then
		resStr=Round(vByte/1024, vDigit) & " kB"
	ElseIf (vByte<966367642) Then
		resStr=Round(vByte/(1024*1024),vDigit) & " MB"
	Else
		resStr=Round(vByte/(1024*1024*1024),vDigit) & " GB"
	End If
	convertByte=resStr
End Function

' === Class of Index, Name, Description and Size of image in WIM file ===
Class imageWim
	Private prIndex, prName, prDescription, prSize
	Public Default Function Init(arrStr)
		sLine=1
		eLine=InStr(sLine, arrStr, "|")
		prIndex=CInt(Mid(arrStr, sLine, eLine-sLine))
		sLine=eLine+1
		eLine=InStr(sLine, arrStr, "|")
		prName=Mid(arrStr, sLine, eLine-sLine)
		sLine=eLine+1
		eLine=InStr(sLine, arrStr, "|")
		prDescription=Mid(arrStr, sLine, eLine-sLine)
		sLine=eLine+1
		prSize=Mid(arrStr, sLine, Len(arrStr)-sLine+1)
		Set Init=Me
	End Function
	Public Property Get Index()
		Index=prIndex
	End Property
	Public Property Get Name()
		Name=prName
	End Property
	Public Property Get Description()
		Description=prDescription
	End Property
	Public Property Get Size()
		Size=prSize
	End Property
End Class

' = END SUBPORT FUNCTION =

' = CONTROL EVENTS =
' === Capture Clicked ===
Sub captureNow()
	Set fso=CreateObject("Scripting.FileSystemObject")
	capDir=getFolder("Select folder that you want capture (Capture all contents in folder without folder):", true, false, false)
	Dim dirPath
	dirPath = False
	If capDir<>"" Then
		Do
			capPath=getFolder("Select folder store image file:", true, true, false)
			extPath=Left(capPath, Len(capDir))
			If extPath=capDir Then
				MsgBox "Can not save image file in captured folder!", 16, "Warning"
				dirPath=True
			Else
				dirPath=False
			End If
		Loop While (dirPath)
		If capPath<>"" Then
			Dim capFile
			Dim condFile
			condFile=False
			Do
				capFile=InputBox("Enter your image file name:","Input",fso.GetBaseName(capDir))
				If checkName(capFile) Then
					condFile=True
				Else
					MsgBox "File name can not content " & Chr(&H2F) &  Chr(&H5C) & Chr(&H3A) &_
					Chr(&H2A) & Chr(&H3F) & Chr(&H22) & Chr(&H3C) & Chr(&H3E) & Chr(&H7C), 16, "Warning"
					condFile=False
				End If
				If (fso.FileExists(capPath & "\" & capFile & ".wim") Or fso.FileExists(capPath & "\" & capFile & ".esd")) Then
					MsgBox "File name " & capFile & " is exists, chose another name", 0+16, "Warning"
					condFile=False
				End If
			Loop Until (condFile)
			If capFile<>"" Then
				capType=MsgBox("Capture in Electronic Software Download format (*.esd)?", 4+32, "Confirm")
				capExt=".wim"
				If capType=6 Then capExt=".esd"
				bootOpt=MsgBox("Do you want make bootable image file?", 4+32, "Confirm")
				bootStr=""
				If bootOpt=6 Then bootStr=" /bootable"
				compOpt=MsgBox("Do you want compress this image file?", 1+32, "Confirm")
				compStr=""
				If compOpt=1 Then
					maxOpt=MsgBox("High compress?", 4+32, "Confirm")
					If maxOpt=6 Then compStr=" /compress:max" Else compStr=" /compress:fast"
				Else
					compStr=" /compress:none"
				End If
				wimName=InputBox("Enter your image index name:", "Confirm", fso.GetBaseName(capDir))
				If wimName="" Then wimName=fso.GetBaseName(capDir)
				wimDesc=InputBox("Enter your image index description:", "Confirm", "This is " & fso.GetBaseName(capDir))
				descStr=""
				If wimDesc="" Then descStr="" Else descStr=" /description:" & Chr(34) & wimDesc & Chr(34)
				cmd="dism.exe /capture-image /imagefile:" & Chr(34) & capPath & "\" & capFile & capExt & Chr(34) &_
				" /capturedir:" & Chr(34) & capDir & Chr(34) & " /name:" & Chr(34) & wimName & Chr(34) & descStr &_
				compStr & bootStr & " /checkintegrity /verify"
				exCmd=readCmd(cmd)
				If Right(exCmd,8)="dism.log" Then
					MsgBox "Error when capture image file! Try again!", 16, "Error"
				Else
					MsgBox "Capture image file complete!", 32, "Congratulation"
				End If
			End If
		End If
	End If
End Sub

' === Apply Clicked ===
Sub applyNow()
	appWim=getFile("Select your image file","All Image files (*.wim;*.esd;*.swm)|*.wim;*.esd;*.swm|Windows Imaging Format files (*.wim)|*.wim|Electronic Software Download files (*.esd)|*.esd|Split Windows Imaging Format files (*.swm)|*.swm|")
	If appWim<>"" Then
		imgs=readWim(appWim)
		If UBound(imgs)>0 Then
			strDisp="File: " & appWim & " have " & (Ubound(imgs)+1) & " images:" & Chr(13) & Chr(13)
			For i=0 To UBound(imgs)
				Set img=(New imageWim)(imgs(i))
				strDisp=strDisp & img.Index & ". " & img.Name & " [" & convertByte(img.Size,2) & "]" & Chr(13) &_
				img.Description & Chr(13) & Chr(13)
			Next
			strDisp = strDisp & "Enter index number to select image:"
			Dim conFill, conNumb, conMore, conLess
			conFill=False:conNumb=False:conMore=False:conLess=False
			Do Until (conFill And conNumb And conMore And conLess)
				strInd=InputBox(strDisp,"Enter your image index",1)
				conFill=(strInd<>"")
				conNumb=(IsNumeric(strInd))
				If IsNumeric(strInd) Then
					conMore=(CInt(strInd)>0)
					conLess=(CInt(strInd)<=UBound(imgs)+1)
				Else
					conMore=False:conLess=False
				End If
				If Not (conFill And conNumb And conMore And conLess) Then
					If Not conFill Then
						Exit Do
					Else
						MsgBox "Invalid value!", 0+48, "Warning"
					End If
				End If
			Loop
			If Not conFill Then Exit Sub
		End If
		appDir=getFolder("Select folder that image file applied:", true, true, false)
		If appDir<>"" Then
			Set fso=CreateObject("Scripting.FileSystemObject")
			swmPat=""
			If fso.GetExtensionName(appWim)="swm" Then swmPat=" /swmfile:"  & Chr(34) & fso.GetParentFolderName(appWim) &_
			"\" & fso.GetBaseName(appWim) & "*.swm" & Chr(34)
			str="dism.exe /apply-image /imagefile:" & Chr(34) & appWim & Chr(34) &_
			swmPat & " /index:" & strInd & " /applydir:" & Chr(34) & appDir & Chr(34) & " /checkintegrity /verify"
			exCmd=readCmd(str)
			If Right(exCmd,8)="dism.log" Then
				MsgBox "Error when apply image file! Try again!", 16, "Error"
			Else
				MsgBox "Apply image file complete!", 32, "Congratulation"
			End If
		End If
	End If
End Sub

' === Mount Clicked ===
Sub mountNow()
	Set fso=CreateObject("Scripting.FileSystemObject")
	mouWim=getFile("Select your image file","All Mounted Image files (*.wim;*.esd)|*.wim;*.esd|image Format files(*.wim)|*.wim|Electric Software Download files (*.esd)|*.esd|")
	If mouWim<>"" Then
		imgs=readWim(mouWim)
		Dim strInd:strInd=1
		Dim strDrvLb:strDrvLb=""
		If UBound(imgs)>0 Then
			strDisp="File: " & appWim & " have " & (Ubound(imgs)+1) & " images:" & Chr(13) & Chr(13)
			For i=0 To UBound(imgs)
				Set img=(New imageWim)(imgs(i))
				strDisp=strDisp & img.Index & ". " & img.Name & " [" & convertByte(img.Size,2) & "]" & Chr(13) &_
				img.Description & Chr(13) & Chr(13)
			Next
			strDisp = strDisp & "Enter index number to select image:"
			Dim conFill, conNumb, conMore, conLess
			conFill=False:conNumb=False:conMore=False:conLess=False
			Do Until (conFill And conNumb And conMore And conLess)
				strInd=InputBox(strDisp,"Enter your image index",1)
				conFill=(strInd<>"")
				conNumb=(IsNumeric(strInd))
				If IsNumeric(strInd) Then
					conMore=(CInt(strInd)>0)
					conLess=(CInt(strInd)<=UBound(imgs)+1)
				Else
					conMore=False:conLess=False
				End If
				If Not (conFill And conNumb And conMore And conLess) Then
					If Not conFill Then
						Exit Do
					Else
						MsgBox "Invalid value!", 0+48, "Warning"
					End If
				End If
			Loop
			If Not conFill Then Exit Sub
		End If
		Dim nxtDrv
		wDrv=Left(fso.GetSpecialFolder(WindowsFolder),3)
		For i=65 To 90
			If (Not fso.DriveExists(Chr(i)) And Not fso.FolderExists(wDrv & "mDrives\" & Chr(i))) Then
				nxtDrv=Chr(i)
				Exit For
			End If
		Next
		Set img=(New imageWim)(imgs(CInt(strInd)-1))
		strDrvLb=img.Name
		If (img.Name="" Or img.Name="<undefined>") Then strDrvLb=fso.GetBaseName(mouWim)
		wDir=wDrv & "mDrives"
		If (Not fso.FolderExists(wDir)) Then fso.CreateFolder(wDir)
		mDir=wDir & "\" & nxtDrv
		If fso.FolderExists(mDir) Then fso.DeleteFolder mDir, True
		fso.CreateFolder(mDir)
		Set netW=CreateObject("WScript.Network")
		compName=netW.ComputerName
		mountStr="dism.exe /mount-image /imagefile:" & Chr(34) & mouWim & Chr(34) & " /index:" & strInd & " /mountdir:" &_
		Chr(34) & mDir & Chr(34) & " /readonly /checkintegrity"
		shrStr="net.exe share " & Chr(34) & strDrvLb & Chr(34) & "=" & Chr(34) & mDir & Chr(34) & " /cache:none"
		mapStr="net use " & nxtDrv & ": " & Chr(34) & "\\" & compName & "\" & strDrvLb & Chr(34)
		Set cmdFile=fso.CreateTextFile(wDir & "\command.cmd", True)
		cmdFile.WriteLine(mountStr)
		cmdFile.WriteLine(shrStr)
		cmdFile.WriteLine(mapStr)
		cmdFile.Close
		exeRun(wDir & "\command.cmd")
	End If
End Sub

' === Unmount Clicked ===
Sub unmountNow()
	Set fso=CreateObject("Scripting.FileSystemObject")
	wDrv=Left(fso.GetSpecialFolder(WindowsFolder),3)
	If fso.FolderExists(wDrv & "mDrives") Then
		Dim lstDrvs
		lstDrvs=""
		cntDrv=0
		For Each vDrv In fso.GetFolder(wDrv & "mDrives").SubFolders
			lstDrvs = lstDrvs & Chr(13) & fso.GetBaseName(vDrv.Path)
			cntDrv=cntDrv+1
		Next
		If cntDrv<>0 Then
			Dim trueDrv
			Do
				lblDrv=InputBox("All virtual drive:" & lstDrvs & Chr(13) & "Enter letter of your unmounted virtual drive:", "Enter your virtual drive letter", "A")
				If fso.FolderExists(wDrv & "mDrives\" & lblDrv) Then
					trueDrv = True
				Else
					trueDrv = False
					MsgBox "Invalid virtual drive letter!", 0+48, "Warning"
				End If
			Loop Until (trueDrv)
			If lblDrv<>"" Then
				cfmDel=MsgBox("Are you sure unmount virtual drive " & UCase(lblDrv), 4+32, "Confirm")
				If cfmDel=6 Then
					Dim disMap, disShr
					If fso.FolderExists(lblDrv & ":\") Then
						Set netW=CreateObject("WScript.Network")
						compName=netW.ComputerName
						disMap="net.exe use " & lblDrv & ": /delete /y"
					Else
						disMap=""
					End If
					disShr="net.exe share " & wDrv & "mDrives\" & lblDrv & " /delete"
					unDir = "dism.exe /unmount-image /mountdir:" & wDrv & "mDrives\" & lblDrv & " /discard /checkintegrity"
					Set cmdFile=fso.CreateTextFile(wDrv & "mDrives\command.cmd", True)
					cmdFile.WriteLine(disMap)
					cmdFile.WriteLine(disShr)
					cmdFile.WriteLine(unDir)
					cmdFile.WriteLine("rd /s /q " & wDrv & "mDrives\" & lblDrv)
					cmdFile.Close
					exeRun(wDrv & "mDrives\command.cmd")
				End If
			Else
				Exit Sub
			End If
		Else
			MsgBox "No virtual drive to unmount :)", 0+64, "Information"
		End If
	Else
		MsgBox "No virtual drive to unmount :)", 0+64, "Information"
	End If
End Sub

' === Add Clicked ===
Sub addNow()
	Set fso=CreateObject("Scripting.FileSystemObject")
	addWim=getFile("Select your image file","All Image files (*.wim;*.esd)|*.wim;*.esd|image Format files(*.wim)|*.wim|Electronic Software Download files (*.esd)|*.esd|")
	If addWim<>"" And fso.FileExists(addWim) Then
		imgs=readWim(addWim)
		addDir=getFolder("Select captured folder", true, true, false)
		If addDir<>"" And fso.FolderExists(addDir) Then
			condExists=True
			Do While condExists
				wimName=InputBox("Enter your image index name:", "Confirm", fso.GetBaseName(addDir))
				If wimName="" Then wimName=fso.GetBaseName(addDir)
				condExists=False
				For i=0 To UBound(imgs)
					Set img=(New imageWim)(imgs(i))
					If UCase(img.Name)=UCase(wimName) Then 
						condExists=True
						MsgBox wimName & " is same name of image in image file! Enter another name:)", 0+16, "Warning"
					End If
				Next
			Loop
			wimDesc=InputBox("Enter your image index description:", "Confirm", "This is " & fso.GetBaseName(addDir))
			descStr=""
			If wimDesc="" Then descStr="" Else descStr=" /description:" & Chr(34) & wimDesc & Chr(34)
			bootOpt=MsgBox("Do you want make bootable image file?", 4+32, "Confirm")
			bootStr=""
			If bootOpt=6 Then bootStr=" /bootable"
			strCmd="dism.exe /append-image /imagefile:" & Chr(34) & addWim & Chr(34) & " /capturedir:" &_
			Chr(34) & addDir & Chr(34) & " /name:" & Chr(34) & wimName & Chr(34) & descStr & bootStr & " /checkintegrity /verify"
			exCmd=readCmd(strCmd)
			If Right(exCmd,8)="dism.log" Then
				MsgBox "Error when add image! Try again!", 16, "Error"
			Else
				MsgBox "Add image complete!", 32, "Congratulation"
			End If
		End If
	End If
End Sub

' === Delete Clicked ===
Sub deleteNow()
	delWim=getFile("Select your image file","All Image files (*.wim;*.esd)|*.wim;*.esd|image Format files(*.wim)|*.wim|Electronic Software Download files (*.esd)|*.esd|")
	If delWim<>"" Then
		imgs=readWim(delWim)
		If UBound(imgs)>0 Then
			strDisp="File: " & delWim & " have " & (Ubound(imgs)+1) & " images:" & Chr(13) & Chr(13)
			For i=0 To UBound(imgs)
				Set img=(New imageWim)(imgs(i))
				strDisp=strDisp & img.Index & ". " & img.Name & " [" & convertByte(img.Size,2) & "]" & Chr(13) &_
				img.Description & Chr(13) & Chr(13)
			Next
			strDisp = strDisp & "Enter index number to select image:"
			Dim conFill, conNumb, conMore, conLess, strInd
			conFill=False:conNumb=False:conMore=False:conLess=False
			Do Until (conFill And conNumb And conMore And conLess)
				strInd=InputBox(strDisp,"Enter your image index",1)
				conFill=(strInd<>"")
				conNumb=(IsNumeric(strInd))
				If IsNumeric(strInd) Then
					conMore=(CInt(strInd)>0)
					conLess=(CInt(strInd)<=UBound(imgs)+1)
				Else
					conMore=False:conLess=False
				End If
				If Not (conFill And conNumb And conMore And conLess) Then
					If Not conFill Then
						Exit Do
					Else
						MsgBox "Invalid value!", 0+48, "Warning"
					End If
				End If
			Loop
			If Not conFill Then Exit Sub
			Set img=(New imageWim)(imgs(CInt(strInd)-1))
			cfmDel=MsgBox("Are you sure delete " & img.Name & "?", 4+32, "Confirm")
			If cfmDel=6 Then
				strCmd="dism.exe /delete-image /imagefile:" & Chr(34) & delWim & Chr(34) &_
				" /index:" & strInd & " /checkintegrity"
				exCmd=readCmd(strCmd)
				If Right(exCmd,8)="dism.log" Then
					MsgBox "Error when delete image! Try again!", 16, "Error"
				Else
					MsgBox "Delete image complete!", 32, "Congratulation"
				End If
			End If
		Else
			MsgBox "File: " & delWim & " have 1 image. Can not delete this image!", 0+16, "Warning"
		End If
	End If
End Sub

' === Export Clicked ===
Sub exportNow()
	expWim=getFile("Select your image file","All Image files (*.wim;*.esd;*.swm)|*.wim;*esd;*.swm|Windows Imaging Format files (*.wim)|*.wim|Electronic Software Download files (*.esd)|*.esd|Split Windows Imaging Format files (*.swm)|*.swm|")
	If expWim<>"" Then
		imgs=readWim(expWim)
		If UBound(imgs)>0 Then
			strDisp="File: " & expWim & " have " & (Ubound(imgs)+1) & " images:" & Chr(13) & Chr(13)
			For i=0 To UBound(imgs)
				Set img=(New imageWim)(imgs(i))
				strDisp=strDisp & img.Index & ". " & img.Name & " [" & convertByte(img.Size,2) & "]" & Chr(13) &_
				img.Description & Chr(13) & Chr(13)
			Next
			strDisp = strDisp & "Enter index number to select image:"
			Dim conFill, conNumb, conMore, conLess, strInd
			conFill=False:conNumb=False:conMore=False:conLess=False
			Do Until (conFill And conNumb And conMore And conLess)
				strInd=InputBox(strDisp,"Enter your image index",1)
				conFill=(strInd<>"")
				conNumb=(IsNumeric(strInd))
				If IsNumeric(strInd) Then
					conMore=(CInt(strInd)>0)
					conLess=(CInt(strInd)<=UBound(imgs)+1)
				Else
					conMore=False:conLess=False
				End If
				If Not (conFill And conNumb And conMore And conLess) Then
					If Not conFill Then
						Exit Do
					Else
						MsgBox "Invalid value!", 0+48, "Warning"
					End If
				End If
			Loop
			If Not conFill Then Exit Sub
			Set img=(New imageWim)(imgs(CInt(strInd)-1))
			expPath=getFolder("Select folder store image file:", true, true, false)
			Set fso=CreateObject("Scripting.FileSystemObject")
			If expPath<>"" Then
				condFile=False
				Do
					expFile=InputBox("Enter your image file name:","Input","")
					If checkName(expFile) Then
						condFile=True
					Else
						MsgBox "File name can not content " & Chr(&H2F) &  Chr(&H5C) & Chr(&H3A) &_
						Chr(&H2A) & Chr(&H3F) & Chr(&H22) & Chr(&H3C) & Chr(&H3E) & Chr(&H7C), 16, "Warning"
						condFile=False
					End If
					If (fso.FileExists(expPath & "\" & expFile & ".wim") Or fso.FileExists(expPath & "\" & expFile & ".esd")) Then
						MsgBox "File name " & expFile & " is exists, chose another name", 0+16, "Warning"
						condFile=False
					End If
				Loop Until (condFile)
				If expFile<>"" Then
					expType=MsgBox("Capture in Electronic Software Download format (*.esd)?", 4+32, "Confirm")
					expExt=".wim"
					If expType=6 Then expExt=".esd"
					bootOpt=MsgBox("Do you want make bootable image file?", 4+32, "Confirm")
					bootStr=""
					If bootOpt=6 Then bootStr=" /bootable"
					compOpt=MsgBox("Do you want compress this image file?", 1+32, "Confirm")
					compStr=""
					If compOpt=1 Then
						maxOpt=MsgBox("High compress?", 4+32, "Confirm")
						If maxOpt=6 Then compStr=" /compress:max" Else compStr=" /compress:fast"
					Else
						compStr=" /compress:none"
					End If
					wimName=InputBox("Enter your image index name:", "Confirm", img.Name)
					If wimName="" Then wimName=img.Name
					strSwm=""
					If fso.GetExtensionName(expWim)="swm" Then strSwm=" /swmfile:" & Chr(34) & fso.GetParentFolderName(expWim) & "\" & fso.GetBaseName(expWim) & "*.swm" & Chr(34)
					strCmd="dism.exe /export-image /sourceimagefile:" & Chr(34) & expWim & Chr(34) & strSwm & " /sourceindex:" & strInd &_
					" /destinationimagefile:" & Chr(34) & expPath & "\" & expFile & expExt & Chr(34) & " /destinationname:" &_
					Chr(34) & wimName & Chr(34) & compStr & bootStr & " /checkintegrity"
					exCmd=readCmd(strCmd)
					If Right(exCmd,8)="dism.log" Then
						MsgBox "Error when export image file! Try again!", 16, "Error"
					Else
						MsgBox "Export image file complete!", 32, "Congratulation"
					End If
				End If
			End If
		End If
		
	End If
End Sub

' === Split Clicked ===
Sub splitNow()
	Set fso=CreateObject("Scripting.FileSystemObject")
	splWim=getFile("Select your image file","All Image files (*.wim;*.esd)|*.wim;*esd|Windows Imaging Format files (*.wim)|*.wim|Electronic Software Download files (*.esd)|*.esd|")
	imgSz=CLng(fso.GetFile(splWim).Size)
	If imgSz<10*1024*1024 Then
		MsgBox "Your image file is very small, chose another large file size :)", 0+32, "Information"
		Exit Sub
	End If
	If splWim<>"" Then
		swmPath=getFolder("Select your split image path", true, true, false)
		If swmPath<>"" Then
			Dim swmFile
			Dim condFile
			condFile=False
			Do
				swmFile=InputBox("Enter your split image file name:","Input",fso.GetBaseName(splWim))
				If checkName(swmFile) Then
					condFile=True
				Else
					MsgBox "File name can not content " & Chr(&H2F) &  Chr(&H5C) & Chr(&H3A) &_
					Chr(&H2A) & Chr(&H3F) & Chr(&H22) & Chr(&H3C) & Chr(&H3E) & Chr(&H7C), 16, "Warning"
					condFile=False
				End If
				If fso.FileExists(swmPath & "\" & swmFile & ".swm") Then
					MsgBox "File name " & swmFile & " is exists, chose another name", 0+16, "Warning"
					condFile=False
				End If
			Loop Until (condFile)
			maxSz=Int(0.9*imgSz/(1024*1024))
			Dim condSize: condSize=False
			Do
				swmSize=InputBox("Enter your split image file size (MB)","Input",maxSz)
				If IsNumeric(swmSize) Then
					If CInt(swmSize)>0 And CInt(swmSize)<=maxSz Then
						condSize=True
					Else
						condSize=False
						MsgBox "Enter size more than 0(MB) and  less than or equals " & maxSz & "(MB)!", 0+16, "Warning"
					End If
				Else
					condSize=False
					MsgBox "Enter a numeric!", 0+16, "Warning"
				End If
			Loop Until (condSize)
			strCmd="dism.exe /split-image /imagefile:" & Chr(34) & splWim & Chr(34) & " /swmfile:" & Chr(34) & swmPath &_
			"\" & swmFile & ".swm" & Chr(34) & " /filesize:" & swmSize
			exCmd=readCmd(strCmd)
			If Right(exCmd,8)="dism.log" Then
				MsgBox "Error when split image file! Try again!", 16, "Error"
			Else
				MsgBox "Split image file complete!", 32, "Congratulation"
			End If
		End If
	End If
End Sub

' === Get image file's information
Sub infoNow()
	infWim=getFile("Select your image file","All Image files (*.wim;*.esd;*.swm)|*.wim;*esd;*.swm|Windows Imaging Format files (*.wim)|*.wim|Electronic Software Download files (*.esd)|*.esd|Split Windows Imaging Format files (*.swm)|*.swm|")
	If infWim<>"" Then
		arrWim=readWim(infWim)
		cntWim=" image" : If UBound(arrWim)>0 Then cntWim=" images"
		strDisp="File: " & infWim & " have " & (UBound(arrWim)+1) & cntWim & Chr(13) & Chr(13)
		For i=0 To UBound(arrWim)
			Set imgWim = (New imageWim)(arrWim(i))
			strDisp = strDisp & "Index : " & imgWim.Index & Chr(13)
			strDisp = strDisp & "Name : " & imgWim.Name & Chr(13)
			strDisp = strDisp & "Description : " & imgWim.Description & Chr(13)
			strDisp = strDisp & "Size : " & convertByte(imgWim.Size,2) & Chr(13) & Chr(13)
		Next
		MsgBox Trim(strDisp), 0+64, "Information"
	End If
End Sub
' = CONTROL EVENTS =

' === End Of File ===
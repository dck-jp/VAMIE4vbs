Option Explicit
' VAMIE (VBScript Auto Mation for Internet Explorer)
' LastModified : 2014/8/25
' Created By D*isuke YAMAKWA

' VBA版との差異
'	Win32API をコールしていたところは、他の手段に置き換え
'	DoEvents削除

Class VAMIE
	Private ie
	
	Private Sub Sleep(time)   'VBAの場合は、Win32API呼び出し
		WScript.Sleep time
	End Sub
	
	Public Property Let Visible(setBoolean)
	    ie.Visible = setBoolean
	End Property
	Public Property Get Visible
	    Visible = ie.Visible
	End Property
	Public Property Let FullScreen(setBoolean)
	    ie.FullScreen = setBoolean
	End Property
	Public Property Get FullScreen
	    FullScreen = ie.FullScreen
	End Property
	Public Property Get Document 'VAMIEに用意されたメソッドでは目的の動作が出来ない時用(Documentクラスを直接操作したい場合用)
	    Document = ie.Document
	End Property
	Public Property Get LocationURL
		LocationURL = ie.LocationURL
	End Property
	Public Property Get LocationName
		LocationName = ie.LocationName
	End Property

	Sub Class_Initialize
	    Set ie = CreateObject("InternetExplorer.Application")
	    ie.Visible = True
	End Sub
	Sub Class_Terminate
		If flagQuitWhenTerminate Then ie.Quit
		Set ie = Nothing
	End Sub
	Dim flagQuitWhenTerminate ' デストラクタ用
	Public Property Let AutoQuit(setBoolean)
		flagQuitWhenTerminate = setBoolean
	End Property
	
	'--------------------	'--------------------	'--------------------
	Public Sub Navigate(url)
	    ie.Navigate url
	    WaitLoading
	End Sub
	Public Sub NavigateWithNoWait(url)	' WaitLoadingを挟むと無限ループするようなページ対策
	    ie.Navigate url
	End Sub
	Public Sub Quit()
	    ie.Quit
	End Sub
	Public Sub Refresh()
	    ie.Refresh
	End Sub
	Sub ResizeTo(width,height)
		If LocationURL = Empty then 
			msgbox("VAMIE 制限 : ReizeToメソッドはページを表示した後で呼び出してください")
			exit sub
		end if

		call ExecuteJavaScript("window.resizeTo(" & width & "," & height & ");")
	End Sub
	
	'DOM要素操作用メソッド群 ----------------------------------------------------
	Sub Exists(element)
		dim test : Set test = element
		If test <> Empty then
			Exists = True
		Else
			Exists = False
		End if
	End Sub
	Function Find(arr) ' 簡易DOMセレクタ 【引数の与え方】例： VAMIE.Find(Array("id","hoge","class","fuga",1, "tag","table",2))
		Dim parent_obj : Set parent_obj = ie.Document
		Dim child_obj 
		dim dom_id, tag_name, index_num, name_ 

		Dim cur : cur = 0
		Dim continue_flag : continue_flag = True
		Do While continue_flag = True
		        Select Case arr(cur):
		            Case "id"
		                dom_id = arr(cur + 1)
		                Set child_obj = parent_obj.getElementById(dom_id)
		                cur = cur + 2
		            Case "tag"
		                tag_name = arr(cur + 1)
		                index_num = arr(cur + 2)
		                Set child_obj = parent_obj.GetElementsByTagName(tag_name)(index_num)
		                cur = cur + 3
		            Case "name"
		                name_ = arr(cur + 1)
		                index_num = arr(cur + 2)
		                Set child_obj = parent_obj.getElementsByName(name_)(index_num)
		                cur = cur + 3
		            Case "class"
		                name_ = arr(cur + 1)
		                index_num = arr(cur + 2)
		                Set child_obj = parent_obj.getElementsByClassName(name_)(index_num)
		                cur = cur + 3
		        End Select
		        
		        Set parent_obj = child_obj
		        
		        If cur > UBound(arr) Then
		            continue_flag = False
		        End If
		Loop
		
		Set Find = parent_obj
	End Function

	Function FindById(dom_id)
	    Set FindById = ie.Document.getElementById(dom_id)' 注：旧IEのgetElementByIdはnameも参照する
	End Function
	Function FindsByName(name)
	    Set FindsByName = ie.Document.GetElementsByName(name)
	End Function
	Function FindsByTag(tag_name)
	    Set FindsByTag = ie.Document.GetElementsByTagName(tag_name)
	End Function
	Function FindsByClass(className)
	    Set FindsByClass = ie.Document.GetElementsByClassName(className)
	End Function

	Function GetInnerText(element) 'テキストを取得
	    GetInnerText = element.innerText
	End Function
	Function GetInnerHTML(element) 'HTMLコードを取得
	    GetInnerHTML = element.innerHTML
	End Function

	Sub SetValue(element, val)' テキストボックスへの入力など
	    element.value = val
	    WaitLoading
	End Sub
	Sub Click(element)' 送信ボタンやリンクをクリック
	    element.Click
	    WaitLoading
	End Sub
	Sub SetCheckBox(element, checked_flag)' チェックボックスの状態をセットします
	    If Not (element.Checked = checked_flag) Then
		Call Click(element)
	    End If
	End Sub
	Sub SelectListBox(element, label)' セレクトボックスを文言ベースで選択します
	    If Len(label) < 1 Then Exit Sub

	    Dim opts : Set opts = element.Options
	    Dim i : For i = 0 To opts.Length - 1
	        If opts(i).innerText = label Then
	            opts(i).Selected = True
	            Exit Sub
	        End If
	    Next
	End Sub
	Sub SetRadioButton(element, value)' ラジオボタンを値ベースで選択します
	    If Len(value) < 1 Then Exit Sub

	    Dim radios: Set radios = element
	    Dim i: For i = 0 To radios.Length - 1
	        If radios(i).value = CStr(value) Then
	            radios(i).Click
	            Sleep 100
	        End If
	    Next
	End Sub

	' -----------------------------------------------------------------------------
	Public Sub WaitLoading()
	    Do While ie.Busy = True Or ie.ReadyState <> 4
	        Sleep 100
	    Loop
	    Sleep 100
	End Sub
	Public Sub Wait(millisecond)
	    Sleep millisecond
	End Sub

	' おまけ ---------------------------------------------------------------------
	Function GetIEVersion()
	    Dim FS: Set FS = CreateObject("Scripting.FileSystemObject")
	    Dim hoge: hoge = Fix(val(FS.GetFileVersion(ie.FullName)))
	    GetIEVersion = hoge
	End Function

	Sub DisableConfirmFunction()'confirm()呼び出し時に確認ダイアログを表示させない
	    Dim ele: Set ele = ie.Document.createElement("SCRIPT")
	    
	    ele.Type = "text/javascript"
	    ele.text = "function confirm() { return true; }"
	    
	    Call ie.Document.body.appendChild(ele)
	End Sub
	
	Sub Activate() 'SendKeys用
		Dim wLoc, wSvc, wEnu, wIns
		Set wLoc = CreateObject("WbemScripting.SWbemLocator")
		Set wSvc = wLoc.ConnectServer
		Set wEnu = wSvc.InstancesOf("Win32_Process")
		Dim pId
		For Each wIns in wEnu
		    If Not IsEmpty(wIns.ProcessId) And wIns.Description = "iexplore.exe" Then
		        pId = wIns.ProcessId
		    End If
		Next

		dim wsh : Set wsh = CreateObject("Wscript.Shell")
		While not wsh.AppActivate(pId) 
			Sleep 100 
		Wend 
	End Sub

	Sub SendKeys(keys) '値を代入するのではなく、キー入力をエミュレートしたい場合
		dim wsh : Set wsh = CreateObject("Wscript.Shell")
		wsh.SendKeys keys
	End Sub

	Public Sub ExecuteJavaScript(jsCode)
		Call ie.Document.Script.setTimeout("javascript:" & jsCode, 1) ' 第2引数:実行までの待機時間[msec]
		WaitLoading
	End Sub
End Class

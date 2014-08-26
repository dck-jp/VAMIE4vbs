Option Explicit
Call Main()

Public Sub Import(ByVal strFileName) ' 外部スクリプトをロードするための関数
	Dim objFso : Set objFso = WScript.CreateObject("Scripting.FileSystemObject")
	DIm objFile : Set objFile = objFso.OpenTextFile(strFileName, 1, False) ' 1 = ReadOnly
	ExecuteGlobal objFile.ReadAll()
	objFile.Close()
End Sub

Sub Main()
	Import "VAMIE.vbs"
	
	dim ie : Set ie = New VAMIE
	ie.AutoQuit = True 'インスタンス破棄時にIEを閉じる

	ie.Activate
	Call ie.Navigate("www.google.com")
	Call ie.SetValue(ie.FindsByName("q")(0), "SeleniumVBA")
	Call ie.Click(ie.FindsByName("btnK")(0))

	Call ie.ResizeTo(1024,700)
	Call ie.Wait(2000)
End Sub

Sub Sample()
	dim ie : Set ie = New VAMIE

	ie.Visible = True 'デフォルトTrue
	ie.Document 'Documentクラスを直接操作したいとき用のプロパティ(Frameページの操作とか)
	ie.AutoQuit = True 'インスタンス破棄時にIEを閉じる　デフォルトFalse

	ie.FullScreen = True
	Call ie.ResizeTo(200,300) 'ウィンドウリサイズ

	ie.Activate 'ウィンドウをアクティブに
	Call ie.SendKeys("test") 'アクティブウィンドウにSendKeys ...のはず

	Call ie.Navigate("http://www.google.com/") 'ページを表示して読み込みが終わるまで待機
	Call ie.NavigateWithNoWait("http://www.google.com/") 'ページを表示　（待機なし）　※待機状態が常に続くページ対策
	ie.Refresh

	ie.LocationURL '現在のURLを取得
	ie.LocationName '現在のページのタイトルを取得

	dim DOM_Element
	Set DOM_Element = ie.FindsByID("id") 'idを指定してDOM_Elementを取得
	Set DOM_Element = ie.FindsByName("name")(0) 'Find[s]はDOM_Elementの配列を返す、のでDOM_Elementを取得したいときは添え字を
	Set DOM_Element = ie.FindsByClass("class")(0)
	Set DOM_Element = ie.FindsByTag("tag")(0)
	Set DOM_Element = ie.Find(Array("id", "res", "tag", "li", 0, "tag", "h3", 0))(0) 'DOMセレクタ的なやつ。使えるキーワードは,id, name, tag, class

	If ie.Exists(DOM_Element) Then
		Call ie.GetInnerText(DOM_Element) 'テキスト取得
		Call ie.GetInnerHTML(DOM_Element) 'HTMLコード取得	

		Call ie.SetValue(DOM_Element) '値を入力ていうか代入（キー入力のエミュレーションはSendKeys()で）
		Call ie.Click(DOM_Element) 'クリックとかSubmit
		Call ie.SetCheckBox(DOM_Element,true) 'チェックボックスのON/OFF設定
		Call ie.SelectListBox(DOM_Element, "label名") 'リストボックスにおいて、label名と一致するアイテムを選択
		Call ie.SetRadioButton(DOM_Element, 3) 'ラジオボタンを値ベースで選択

		Call ie.Wait(2000) '指定ミリ秒停止
		ie.WaitLoading '読み込みが終わるまで待機

		ie.GetIEVersion() 'IEのバージョンを文字列で取得
		ie.DisableConfirmFunction() 'JS で実装されたconfirm関数を空に。呼び出し時に確認ダイアログを表示させない
		Call ie.ExecuteJavaScript("window.resizeTo(10,10);") '任意のJavaScriptコードを実行

		ie.Quit 'IEを閉じる
	End If
End Sub
Sub ボタン5_Click()
Attribute ボタン5_Click.VB_Description = "値札ラベルの印刷"
Attribute ボタン5_Click.VB_ProcData.VB_Invoke_Func = "p¥n14"

'ダイアログ入力のIMEをオフにする
Dim n As Long
If IMEStatus = vbIMEHiragana Then SendKeys "{Kanji}"

'印刷枚数を入力するダイアログ  入力された変数は「n」
n = Val(InputBox("印刷枚数を入力してください", Default:="1"))

'入力された値が1以上であればプリンタ「Brother QL-700」を指定して印刷する
If n < 1 Then Exit Sub
Sheets("Print").PrintOut ActivePrinter:="Brother QL-700", Copies:=n

End Sub
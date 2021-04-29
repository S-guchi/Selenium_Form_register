Option Explicit
Sub ServiceNow()
    Dim Driver As New Selenium.WebDriver
    Dim ks As New Keys
    Dim wb As Workbook: Set wb = ThisWorkbook
    Dim Ws1 As Worksheet: Set Ws1 = wb.Worksheets("Main")
    Dim Ws2 As Worksheet: Set Ws2 = wb.Worksheets("メール内容")
    Dim Id As String: Id = Ws1.Range("A3").Value
    Dim Pw As String: Pw = Ws1.Range("B3").Value
        Dim Url As String: Url = "ServiceNowのページ"
    Dim Title As String: Title = ""
    Dim flg As Long: flg = Ws2.Cells(Ws2.Rows.Count, "E").End(xlUp).Row + 1
    Dim MailBody() As Variant
    Dim flgResult() As String
    Dim SCSNo As String: SCSNo = ""
    Dim i As Long
    '解決情報
    Dim kaiketusya As String: kaiketusya = Ws1.Range("C3").Value2
    Dim taisyo As String: taisyo = Ws1.Range("D3").Value2
    Dim closeDate As String: closeDate = Ws1.Range("E3")
    Dim gaiyou As String: gaiyou = ""
On Error GoTo Error:
    If flg <= Ws2.Cells(Ws2.Rows.Count, "A").End(xlUp).Row Then
        MailBody = Ws2.Range("A" & flg & ":D" & Ws2.Cells(Ws2.Rows.Count, "A").End(xlUp).Row + 1).Value
        ReDim flgResult(UBound(MailBody) - 1)
        With Driver
            'ページ起動
            .Start "chrome", Url
            .Get "/"
            .Window.Maximize
            '安全ではないサービスって表示される画面
            .FindElementById("details-button").Click
            .FindElementById("final-paragraph").Click
            'ServiceNowへのログイン画面
            .FindElementByName("src-TUdXLUFE").Click
            .FindElementById("username").SendKeys (Id)
            .FindElementById("password").SendKeys (Pw)
            .FindElementByClass("btn").Submit
            'ServiceNowの画面
            .FindElementsByXPath("/html/body/nav/div/div/button")(1).Click
            For i = LBound(MailBody) To UBound(MailBody) - 1
                If SCSNo <> MailBody(i, 1) Then
                    .FindElementById("sysparm_search").Clear
                    .FindElementById("sysparm_search").SendKeys (MailBody(i, 1))
                    .Wait (2000)
                    .SendKeys (ks.Enter)
                    SCSNo = MailBody(i, 1)
                End If
                'ServiceNowのケース画面
                'iflame切り替え
                .SwitchToFrame ("gsft_main")
                .Wait (2500)
                '承認をクリック、なかったらエラー7
                .FindElementByXPath("//*[@id=""accept""]").Click
                If MailBody(i, 2) <> "" _
                Or MailBody(i, 3) <> "" _
                Or MailBody(i, 4) <> "" Then
                    Title = .FindElementByXPath("//*[@id=""sn_customerservice_case.short_description""]").Value
                    .FindElementByXPath("//*[@id=""tabs2_list""]/span[3]/span").Click
                    .SendKeys (ks.Tab)
                    .SendKeys (ks.Tab)
                    .SendKeys (ks.Enter)
                    .Wait (1000)
                    .FindElementByXPath("//*[@id=""sn_customerservice_task.state""]").SendKeys ("クローズ済み")
                    .FindElementByXPath("//*[@id=""sn_customerservice_task.short_description""]").SendKeys (Title)
                    .FindElementByXPath("//*[@id=""sn_customerservice_task.description""]").SendKeys (MailBody(i, 2))
                    .FindElementByXPath("//*[@id=""sn_customerservice_task.work_notes""]").SendKeys (MailBody(i, 3))
                    .FindElementByXPath("//*[@id=""sn_customerservice_task.u_important_notes""]").SendKeys (MailBody(i, 4))
                    .FindElementByXPath("//*[@id=""sysverb_insert_bottom""]").Click
                End If
                If SCSNo <> MailBody(i + 1, 1) Then
                '概要取得
                    gaiyou = .FindElementByXPath("//*[@id=""sn_customerservice_case.description""]").Value
                '解決情報クリック
                    .FindElementByXPath("//*[@id=""tabs2_section""]/span[5]/span[1]").Click
                '解決者入力
                    .FindElementByXPath("//*[@id=""sys_display.sn_customerservice_case.resolved_by""]").Clear
                    .FindElementByXPath("//*[@id=""sys_display.sn_customerservice_case.resolved_by""]").SendKeys (kaiketusya)
                '概要入力
                    .FindElementByXPath("//*[@id=""sn_customerservice_case.u_event_summary""]").Clear
                    .FindElementByXPath("//*[@id=""sn_customerservice_case.u_event_summary""]").SendKeys (gaiyou)
                '日付入力
                    .FindElementByXPath("//*[@id=""sn_customerservice_case.resolved_at""]").Clear
                    .FindElementByXPath("//*[@id=""sn_customerservice_case.resolved_at""]").SendKeys (Replace(Now, "/", "-"))
                'クローズ日時入力
                    .FindElementByXPath("//*[@id=""sn_customerservice_case.u_days_to_close""]").Clear
                    .FindElementByXPath("//*[@id=""sn_customerservice_case.u_days_to_close""]").SendKeys (closeDate)
                '対処入力
                    .SendKeys (ks.Tab)
                    .SendKeys (ks.Tab)
                    .SendKeys (ks.Tab)
                    .FindElementByXPath("//*[@id=""sn_customerservice_case.close_notes""]").Clear
                    .FindElementByXPath("//*[@id=""sn_customerservice_case.close_notes""]").SendKeys (taisyo)
                '顧客にクローズ依頼をクリック
                    .FindElementByXPath("//*[@id=""proposeSolution""]").Click
                    .SwitchToAlert.Accept
                    .Wait (2000)
                    gaiyou = ""
                End If
                flgResult(i - 1) = "〇"
                .SwitchToParentFrame
            Next i
            .Close
        End With
        Ws2.Range("E" & flg & ":E" & flg + UBound(flgResult)).Value = WorksheetFunction.Transpose(flgResult)
    End If
    MsgBox "処理終了です。"
    wb.Save
    Set Ws1 = Nothing
    Set Ws2 = Nothing
    Set Driver = Nothing
    Exit Sub
Error:
    Select Case Err.Number
    Case 26
        Driver.SwitchToAlert.Accept
        Resume Next
    Case 7
        Resume Next
    Case Else
        MsgBox "読み込みエラーが発生しました。途中までの処理を反映させます。"
    End Select
    Ws2.Range("E" & flg & ":E" & flg + UBound(flgResult)).Value = WorksheetFunction.Transpose(flgResult)
    wb.Save
    Set Ws1 = Nothing
    Set Ws2 = Nothing
    Set Driver = Nothing
End Sub

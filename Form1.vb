Public Class Form1
    Private Sub ButtonClick(sender As Object, e As EventArgs) Handles Button1.Click, Button2.Click, Button3.Click, Button4.Click, Button5.Click, Button6.Click, Button7.Click, Button8.Click, Button9.Click, Button0.Click
        TextBox1.Text = TextBox1.Text & CType(sender, Button).Text
        Label1.Text = Label1.Text & CType(sender, Button).Text
    End Sub
    Private Sub Button12_Click(sender As Object, e As EventArgs) Handles Button12.Click
        If Fix(Val(TextBox1.Text)) = Val(TextBox1.Text) Then TextBox1.Text = TextBox1.Text & "."
        If Mid(TextBox1.Text, 1, 1) = "." Then TextBox1.Text = "0" & TextBox1.Text
        If Mid(TextBox1.Text, 1, 1) = "-" And Mid(TextBox1.Text, 2, 1) = "." Then TextBox1.Text = "-0" & Mid(TextBox1.Text, 2)
        Button12.Focus()
        Label1.Text = Label1.Text & "."
    End Sub
    Private Sub Button13_Click(sender As Object, e As EventArgs) Handles Button13.Click
        TextBox1.Text = ""
        Button13.Focus()
        Label1.Text = Label1.Text & "+"
    End Sub
    Private Sub Button14_Click(sender As Object, e As EventArgs) Handles Button14.Click
        TextBox1.Text = ""
        Button14.Focus()
        Label1.Text = Label1.Text & "-"
    End Sub
    Private Sub Button15_Click(sender As Object, e As EventArgs) Handles Button15.Click
        TextBox1.Text = ""
        Button15.Focus()
        Label1.Text = Label1.Text & "*"
    End Sub
    Private Sub Button16_Click(sender As Object, e As EventArgs) Handles Button16.Click
        TextBox1.Text = ""
        Button16.Focus()
        Label1.Text = Label1.Text & "/"
    End Sub
    Private Sub Button17_Click(sender As Object, e As EventArgs) Handles Button17.Click
        Dim Zuo = "", ZuoShu = "", You = "", YouShu = ""
        On Error Resume Next
        TextBox1.Text = Label1.Text
        Label1.Text = Label1.Text & "="
        While InStr(TextBox1.Text, "*") > 0
            Zuo = Mid(TextBox1.Text, 1, InStr(TextBox1.Text, "*") - 1)
            For z = 1 To 100
                If InStrRev(Microsoft.VisualBasic.Right(Zuo, z), "+") = 0 And InStrRev(Microsoft.VisualBasic.Right(Zuo, z), "-") = 0 And InStrRev(Microsoft.VisualBasic.Right(Zuo, z), "/") = 0 Then
                    ZuoShu = Microsoft.VisualBasic.Right(Zuo, z)
                Else
                    Exit For
                End If
            Next
            You = Mid(TextBox1.Text, InStr(TextBox1.Text, "*") + 1)
            For y = 1 To 100
                If InStr(Mid(You, 1, y), "+") = 0 And InStr(Mid(You, 1, y), "-") = 0 And InStr(Mid(You, 1, y), "*") = 0 And InStr(Mid(You, 1, y), "/") = 0 Then
                    YouShu = Mid(You, 1, y)
                Else
                    Exit For
                End If
            Next
            TextBox1.Text = Mid(TextBox1.Text, 1, Len(Zuo) - Len(ZuoShu)) & Val(ZuoShu) * Val(YouShu) & Microsoft.VisualBasic.Right(TextBox1.Text, Len(You) - Len(YouShu))
        End While
        While InStr(TextBox1.Text, "/") > 0
            Zuo = Mid(TextBox1.Text, 1, InStr(TextBox1.Text, "/") - 1)
            For z = 1 To 100
                If InStrRev(Microsoft.VisualBasic.Right(Zuo, z), "+") = 0 And InStrRev(Microsoft.VisualBasic.Right(Zuo, z), "-") = 0 Then
                    ZuoShu = Microsoft.VisualBasic.Right(Zuo, z)
                Else
                    Exit For
                End If
            Next
            You = Mid(TextBox1.Text, InStr(TextBox1.Text, "/") + 1)
            For y = 1 To 100
                If InStr(Mid(You, 1, y), "+") = 0 And InStr(Mid(You, 1, y), "-") = 0 And InStr(Mid(You, 1, y), "/") = 0 Then
                    YouShu = Mid(You, 1, y)
                Else
                    Exit For
                End If
            Next
            TextBox1.Text = Mid(TextBox1.Text, 1, Len(Zuo) - Len(ZuoShu)) & Val(ZuoShu) / Val(YouShu) & Microsoft.VisualBasic.Right(TextBox1.Text, Len(You) - Len(YouShu))
        End While
        While InStr(Mid(TextBox1.Text, 2), "+") > 0 or InStr(Mid(TextBox1.Text, 2), "-") > 0 '避开首位-号
            For z = 1 To 100
                If InStr(Mid(TextBox1.Text, 2, z), "+") > 0 or InStr(Mid(TextBox1.Text, 2, z), "-") > 0 Then
                    ZuoShu = Mid(TextBox1.Text, 1, z)
                    Exit For
                End If
            Next
            You = Mid(TextBox1.Text, Len(ZuoShu) + 2)
            For y = 1 To 100
                If InStr(Mid(You, 1, y), "+") = 0 And InStr(Mid(You, 1, y), "-") = 0 Then
                    YouShu = Mid(You, 1, y)
                Else
                    Exit For
                End If
            Next
            If Mid(TextBox1.Text, Len(ZuoShu) + 1, 1) = "+" Then
                TextBox1.Text = Val(ZuoShu) + Val(YouShu) & Microsoft.VisualBasic.Right(TextBox1.Text, Len(You) - Len(YouShu))
            ElseIf Mid(TextBox1.Text, Len(ZuoShu) + 1, 1) = "-" Then
                TextBox1.Text = Val(ZuoShu) - Val(YouShu) & Microsoft.VisualBasic.Right(TextBox1.Text, Len(You) - Len(YouShu))
            End If
        End While
        Button17.Focus()
    End Sub
    Private Sub ButtonGenClick(sender As Object, e As EventArgs) Handles Button18.Click, Button20.Click
        If TextBox1.Text <> "" Then TextBox1.Text = TextBox1.Text ^ (1 / CType(sender, Button).Tag)
        Label1.Text = ""
    End Sub
    Private Sub ButtonFangClick(sender As Object, e As EventArgs) Handles Button19.Click, Button21.Click
        If TextBox1.Text <> "" Then TextBox1.Text = TextBox1.Text ^ CType(sender, Button).Tag
        If Len(Label1.Text) > Len(TextBox1.Text) Then
            Label1.Text = Microsoft.VisualBasic.Left(Label1.Text, Len(Label1.Text) - Len(TextBox1.Text)) & TextBox1.Text
        Else
            Label1.Text = TextBox1.Text
        End If
    End Sub
    Private Sub Button22_Click(sender As Object, e As EventArgs) Handles Button22.Click
        On Error Resume Next
        If TextBox1.Text <> "" Then TextBox1.Text = "1/" & TextBox1.Text
        If Len(Label1.Text) > Len(TextBox1.Text) Then
            Label1.Text = Microsoft.VisualBasic.Left(Label1.Text, Len(Label1.Text) - Len(TextBox1.Text)) & TextBox1.Text
        Else
            Label1.Text = TextBox1.Text
        End If
    End Sub
    Private Sub Button23_Click(sender As Object, e As EventArgs) Handles Button23.Click
        TextBox1.Text = ""
        Label1.Text = ""
    End Sub
    Private Sub Form1_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        Dim KeyCode As Short = e.KeyCode
        Dim Shift As Short = e.KeyData \ &H10000
        Dim i As Object
        For i = 0 To 9
            If KeyCode = i + 96 Or KeyCode = i + 48 And Shift <> 1 Then
                TextBox1.Text = TextBox1.Text & i
                Controls("Button" & i).Focus()
                Label1.Text = Label1.Text & i
            End If
        Next
        If KeyCode = 46 And Shift <> 1 Then '按Delete键
            Label1.Text = Microsoft.VisualBasic.Left(Label1.Text, Len(Label1.Text) - Len(TextBox1.Text))
            TextBox1.Text = ""
        End If
        If KeyCode = 8 And Shift <> 1 Then Button10_Click(Button10, New EventArgs()) '按Backspace键
        If KeyCode = 110 Or KeyCode = 190 And Shift <> 1 Then Button12_Click(Button12, New EventArgs) '按.键
        If KeyCode = 107 Or KeyCode = 187 And Shift = 1 Then Button13_Click(Button13, New EventArgs()) '按+键
        If KeyCode = 109 Or KeyCode = 189 And Shift <> 1 Then Button14_Click(Button14, New EventArgs()) '按-键
        If KeyCode = 106 Or KeyCode = 56 And Shift = 1 Then Button15_Click(Button15, New EventArgs()) '按*键
        If KeyCode = 111 Or KeyCode = 191 And Shift <> 1 Then Button16_Click(Button16, New EventArgs()) '按/键
        If KeyCode = 13 Or KeyCode = 187 And Shift <> 1 Then Button17_Click(Button17, New EventArgs()) '按=键
    End Sub
    Private Sub Button24_Click(sender As Object, e As EventArgs) Handles Button24.Click
        TextBox1.Text = TextBox1.Text & "%"
        Label1.Text = Label1.Text & "/100"
    End Sub
    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click
        On Error Resume Next
        If TextBox1.Text <> "" Then
            TextBox1.Text = Microsoft.VisualBasic.Left(TextBox1.Text, Len(TextBox1.Text) - 1)
            Label1.Text = Microsoft.VisualBasic.Left(Label1.Text, Len(Label1.Text) - 1)
        End If
        Button10.Focus()
    End Sub
    Private Sub TextBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox1.KeyDown
        Dim KeyCode As Short = e.KeyCode
        Dim Shift As Short = e.KeyData \ &H10000
        TextBox1.ReadOnly = IIf(KeyCode = 46, False, True) '更改只读属性以粘贴数字
    End Sub
    Private Sub PictureBox1_Click(sender As Object, e As EventArgs) Handles PictureBox1.Click
        MsgBox("核心代码：" & vbCrLf &
               "While InStr(TextBox1.Text,  * ) > 0" & vbCrLf &
               "   Zuo = Mid(TextBox1.Text, 1, InStr(TextBox1.Text,  * ) - 1)" & vbCrLf &
               "   For z = 1 To 100" & vbCrLf &
               "      If InStrRev(Microsoft.VisualBasic.Right(Zuo, z),  + ) = 0 And...Then" & vbCrLf &
               "         ZuoShu = Microsoft.VisualBasic.Right(Zuo, z)" & vbCrLf &
               "      Else" & vbCrLf &
               "         Exit For" & vbCrLf &
               "      End If" & vbCrLf &
               "   Next" & vbCrLf &
               "   You = Mid(TextBox1.Text, InStr(TextBox1.Text,  * ) + 1)" & vbCrLf &
               "   For y = 1 To 100" & vbCrLf &
               "      If InStr(Mid(You, 1, y),  + ) = 0 And...Then" & vbCrLf &
               "         YouShu = Mid(You, 1, y)" & vbCrLf &
               "      Else" & vbCrLf &
               "         Exit For" & vbCrLf &
               "      End If" & vbCrLf &
               "   Next" & vbCrLf &
               "   TextBox1.Text = Mid(TextBox1.Text, 1, Len(Zuo) - Len(ZuoShu)) & Val(ZuoShu)" & vbCrLf &
               "                                * Val(YouShu) & Microsoft.VisualBasic.Right(TextBox1.Text, " & vbCrLf &
               "                                Len(You) - Len(YouShu))" & vbCrLf &
               "End While" & vbCrLf & "赵然" & vbCrLf & "2019年12月16日" & vbCrLf & "其他作品：" & vbCrLf & "文学集《雪华》ISBN978-7-5108-2812-6",, "四则混合运算计算器")
        Button11.Visible = True
    End Sub
    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click
        Button11.Visible = False
    End Sub
End Class
Public Class Calculator

    Dim numQueueList As New Queue() '計算対象
    Dim calcModeQueueList As New Queue() '0:None 1:Add 2:Sub 3:Multi 4:Div
    Dim operQueueList As New Queue() 'ボタン操作
    Dim intDigit As Integer = 8 '計算桁の設定
    Dim errLock As Boolean = False 'Errの時はTrue

#Region "function"

    '少数点を除いて、文字列長さの検測
    Private Function GetLength(text As String) As Integer
        Dim length As Integer = 0
        If text.Contains(".") Then
            length = text.Length - 1
        Else
            length = text.Length
        End If
        Return length
    End Function

    '計算対象キュー
    Private Sub NumEnqueue(key As Double)
        If numQueueList.Count >= 2 Then
            numQueueList.Dequeue()
        End If

        numQueueList.Enqueue(key)
    End Sub

    '操作記録キュー
    Private Sub OperEnqueue(key As String)
        If operQueueList.Count >= 2 Then
            operQueueList.Dequeue()
        End If

        operQueueList.Enqueue(key)
    End Sub

    '演算式キュー
    Private Sub CalcModeEnqueue(key As Integer)
        If calcModeQueueList.Count >= 2 Then
            calcModeQueueList.Dequeue()
        End If

        calcModeQueueList.Enqueue(key)
    End Sub

#End Region

#Region "Display"

    'LCD画面で数字の入力
    Private Sub Display_Show_Add(text As String)
        If operQueueList.Count > 0 Then
            If operQueueList.Peek = "Oper" Or operQueueList.Peek = "Eque" Then
                LabelDisplay.Text = ""
            End If
        End If
        '桁以上の場合は入力禁止
        If GetLength(LabelDisplay.Text) >= intDigit Then
            Return
        End If

        '押した数字を最末へ追加する
        LabelDisplay.Text = LabelDisplay.Text + text
        If LabelDisplay.Text.StartsWith("0") And Not LabelDisplay.Text.StartsWith("0.") Then
            Dim realDisplay As Double = LabelDisplay.Text
            LabelDisplay.Text = realDisplay
        End If
    End Sub

    'LCD画面で小数点の入力
    Private Sub Display_Show_Dot()
        If operQueueList.Count > 0 Then
            If operQueueList.Peek = "Oper" Or operQueueList.Peek = "Eque" Then
                LabelDisplay.Text = "0"
            End If
        End If

        '桁以上の場合は入力禁止
        If GetLength(LabelDisplay.Text) >= intDigit Then
            Return
        End If

        '小数点は既に存在しないことを判断する
        If Not LabelDisplay.Text.Contains(".") Then
            '小数点を最末へ追加する
            LabelDisplay.Text = LabelDisplay.Text + "."
        End If
    End Sub

    'LCD画面のクリア
    Private Sub Display_Clear()
        LabelDisplay.Text = "0"
        numQueueList.Clear()
        operQueueList.Clear()
        calcModeQueueList.Clear()
        errLock = False
    End Sub

#End Region

#Region "Process"

    '計算処理
    Private Sub Process(oper As Integer)
        Dim result As Double = 0.0

        Dim tmpNum As Array = numQueueList.ToArray()

        Select Case oper
            Case 1 '+
                result = tmpNum(0) + tmpNum(1)

            Case 2 '-
                result = tmpNum(0) - tmpNum(1)

            Case 3 '×
                result = tmpNum(0) * tmpNum(1)

            Case 4 '÷
                If tmpNum(1) = 0 Then
                    LabelDisplay.Text = "Err"
                    errLock = True
                    Return
                Else
                    result = tmpNum(0) / tmpNum(1)
                End If

            Case Else
                result = LabelDisplay.Text

        End Select

        '計算結果の桁を処理する
        LabelDisplay.Text = ChangeDigit(result)

        If Double.TryParse(LabelDisplay.Text, tmpNum(1)) Then
            NumEnqueue(LabelDisplay.Text)
        End If
    End Sub

    '計算結果の桁処理
    Private Function ChangeDigit(input As Double) As String
        Dim output As String = input

        '有効桁内の場合
        If GetLength(output) <= intDigit Then
            Return input
        End If

        If output.Contains(".") Then
            '有効桁外、小数の場合
            output = output.Substring(0, intDigit + 1)
            Dim tmp As Double = output
            output = tmp
        Else
            '有効桁外、整数の場合
            output = "Err"
            errLock = True
        End If

        Return output
    End Function

#End Region

#Region "Button"

    'ボタン0のイベント
    Private Sub Button0_Click(sender As Object, e As EventArgs) Handles Button0.Click
        'Err後操作不可にする
        If errLock Then
            Return
        End If

        OperEnqueue("0")
        Display_Show_Add("0")
        ButtonEque.Focus()
    End Sub

    'ボタン1のイベント
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        'Err後操作不可にする
        If errLock Then
            Return
        End If

        OperEnqueue("1")
        Display_Show_Add("1")
        ButtonEque.Focus()
    End Sub

    'ボタン2のイベント
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        'Err後操作不可にする
        If errLock Then
            Return
        End If

        OperEnqueue("2")
        Display_Show_Add("2")
        ButtonEque.Focus()
    End Sub

    'ボタン3のイベント
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        'Err後操作不可にする
        If errLock Then
            Return
        End If

        OperEnqueue("3")
        Display_Show_Add("3")
        ButtonEque.Focus()
    End Sub

    'ボタン4のイベント
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        'Err後操作不可にする
        If errLock Then
            Return
        End If

        OperEnqueue("4")
        Display_Show_Add("4")
        ButtonEque.Focus()
    End Sub

    'ボタン5のイベント
    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        'Err後操作不可にする
        If errLock Then
            Return
        End If

        OperEnqueue("5")
        Display_Show_Add("5")
        ButtonEque.Focus()
    End Sub

    'ボタン6のイベント
    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        'Err後操作不可にする
        If errLock Then
            Return
        End If

        OperEnqueue("6")
        Display_Show_Add("6")
        ButtonEque.Focus()
    End Sub

    'ボタン7のイベント
    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        'Err後操作不可にする
        If errLock Then
            Return
        End If

        OperEnqueue("7")
        Display_Show_Add("7")
        ButtonEque.Focus()
    End Sub

    'ボタン8のイベント
    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        'Err後操作不可にする
        If errLock Then
            Return
        End If

        OperEnqueue("8")
        Display_Show_Add("8")
        ButtonEque.Focus()
    End Sub

    'ボタン9のイベント
    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        'Err後操作不可にする
        If errLock Then
            Return
        End If

        OperEnqueue("9")
        Display_Show_Add("9")
        ButtonEque.Focus()
    End Sub

    'ボタンCのイベント
    Private Sub ButtonC_Click(sender As Object, e As EventArgs) Handles ButtonC.Click
        OperEnqueue("C")
        Display_Clear()
        ButtonEque.Focus()
    End Sub

    'ボタン[.]のイベント
    Private Sub ButtonDot_Click(sender As Object, e As EventArgs) Handles ButtonDot.Click
        'Err後操作不可にする
        If errLock Then
            Return
        End If

        OperEnqueue(".")
        Display_Show_Dot()
        ButtonEque.Focus()
    End Sub

    'ボタン[+]のイベント
    Private Sub ButtonAdd_Click(sender As Object, e As EventArgs) Handles ButtonAdd.Click
        'Err後操作不可にする
        If errLock Then
            Return
        End If

        OperEnqueue("Oper")
        CalcModeEnqueue(1)
        NumEnqueue(LabelDisplay.Text)

        If numQueueList.Count = 2 And calcModeQueueList.Count = 2 And Not operQueueList.Peek = "Oper" And Not operQueueList.Peek = "Eque" Then
            Process(calcModeQueueList.Peek)
        End If
        ButtonEque.Focus()
    End Sub

    'ボタン[-]のイベント
    Private Sub ButtonSub_Click(sender As Object, e As EventArgs) Handles ButtonSub.Click
        'Err後操作不可にする
        If errLock Then
            Return
        End If

        OperEnqueue("Oper")
        CalcModeEnqueue(2)
        NumEnqueue(LabelDisplay.Text)

        If numQueueList.Count = 2 And calcModeQueueList.Count = 2 And Not operQueueList.Peek = "Oper" And Not operQueueList.Peek = "Eque" Then
            Process(calcModeQueueList.Peek)
        End If
        ButtonEque.Focus()
    End Sub

    'ボタン[×]のイベント
    Private Sub ButtonMulti_Click(sender As Object, e As EventArgs) Handles ButtonMulti.Click
        'Err後操作不可にする
        If errLock Then
            Return
        End If

        OperEnqueue("Oper")
        CalcModeEnqueue(3)
        NumEnqueue(LabelDisplay.Text)

        If numQueueList.Count = 2 And calcModeQueueList.Count = 2 And Not operQueueList.Peek = "Oper" And Not operQueueList.Peek = "Eque" Then
            Process(calcModeQueueList.Peek)
        End If
        ButtonEque.Focus()
    End Sub

    'ボタン[÷]のイベント
    Private Sub ButtonDiv_Click(sender As Object, e As EventArgs) Handles ButtonDiv.Click
        'Err後操作不可にする
        If errLock Then
            Return
        End If

        OperEnqueue("Oper")
        CalcModeEnqueue(4)
        NumEnqueue(LabelDisplay.Text)

        If numQueueList.Count = 2 And calcModeQueueList.Count = 2 And Not operQueueList.Peek = "Oper" And Not operQueueList.Peek = "Eque" Then
            Process(calcModeQueueList.Peek)
        End If
        ButtonEque.Focus()
    End Sub

    'ボタン[=]のイベント
    Private Sub ButtonEque_Click(sender As Object, e As EventArgs) Handles ButtonEque.Click
        'Err後操作不可にする
        If errLock Then
            Return
        End If

        OperEnqueue("Eque")
        NumEnqueue(LabelDisplay.Text)
        Dim tmpCalcMode As Array = calcModeQueueList.ToArray()
        If numQueueList.Count = 2 And Not operQueueList.Peek = "Oper" And Not operQueueList.Peek = "Eque" And calcModeQueueList.Count > 0 Then
            Process(tmpCalcMode(tmpCalcMode.Length - 1))
            numQueueList.Clear()
            calcModeQueueList.Clear()
        End If
    End Sub

    'ボタン[±]のイベント
    Private Sub ButtonSymbol_Click(sender As Object, e As EventArgs) Handles ButtonSymbol.Click
        'Err後操作不可にする
        If errLock Then
            Return
        End If

        Dim display As Double = LabelDisplay.Text
        display = display * -1
        LabelDisplay.Text = display
        ButtonEque.Focus()
    End Sub

#End Region

#Region "Shortcut"

    'キーボードのショットカット
    Private Sub Calculator_KeyDown(sender As Object, e As KeyEventArgs) Handles MyBase.KeyDown
        If e.KeyCode = Keys.NumPad1 Then
            Button1_Click(sender, e)
        ElseIf e.KeyCode = Keys.NumPad2 Then
            Button2_Click(sender, e)
        ElseIf e.KeyCode = Keys.NumPad3 Then
            Button3_Click(sender, e)
        ElseIf e.KeyCode = Keys.NumPad4 Then
            Button4_Click(sender, e)
        ElseIf e.KeyCode = Keys.NumPad5 Then
            Button5_Click(sender, e)
        ElseIf e.KeyCode = Keys.NumPad6 Then
            Button6_Click(sender, e)
        ElseIf e.KeyCode = Keys.NumPad7 Then
            Button7_Click(sender, e)
        ElseIf e.KeyCode = Keys.NumPad8 Then
            Button8_Click(sender, e)
        ElseIf e.KeyCode = Keys.NumPad9 Then
            Button9_Click(sender, e)
        ElseIf e.KeyCode = Keys.NumPad0 Then
            Button0_Click(sender, e)
        ElseIf e.KeyCode = Keys.Decimal Then
            ButtonDot_Click(sender, e)
        ElseIf e.KeyCode = Keys.Delete Or e.KeyCode = Keys.Back Then
            ButtonC_Click(sender, e)
        ElseIf e.KeyCode = Keys.Enter Then
            ButtonEque_Click(sender, e)
        ElseIf e.KeyCode = Keys.Add Then
            ButtonAdd_Click(sender, e)
        ElseIf e.KeyCode = Keys.Subtract Then
            ButtonSub_Click(sender, e)
        ElseIf e.KeyCode = Keys.Multiply Then
            ButtonMulti_Click(sender, e)
        ElseIf e.KeyCode = Keys.Divide Then
            ButtonDiv_Click(sender, e)
        End If
    End Sub

#End Region

End Class

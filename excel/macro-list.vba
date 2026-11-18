' Макрос множественного выбора для листа «Монтажный план».
' Куда вставлять: Excel → Сервис → Макрос → Редактор Visual Basic →
' двойной клик по листу с планом → вставить весь код. Файл сохранять как .xlsm.
'
' Работает в колонках «Локация» и «Исполнитель» (или старых «Зона»/«Место», «Подрядчик»):
' выбор второго значения из выпадающего списка дописывает его через " / ",
' повторный выбор убирает, «Все…» заменяет остальное.

Option Explicit

Private Const SEP As String = " / "
Private prevAddr As String
Private prevVal As String

' Запоминаем значение ячейки, когда на неё встаём
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    If Target.Cells.Count = 1 Then
        prevAddr = Target.Address
        prevVal = CStr(Target.Value)
    Else
        prevAddr = ""
    End If
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
    Dim hdrRow As Long, colLoc As Long, colExec As Long
    Dim newVal As String, oldVal As String, result As String
    Dim parts() As String, i As Long, found As Boolean

    If Target.Cells.Count > 1 Then Exit Sub

    ' Колонки ищем по заголовкам (новые названия, затем старые)
    colLoc = FindCol("Локация", hdrRow)
    If colLoc = 0 Then colLoc = FindCol("Зона", hdrRow)
    If colLoc = 0 Then colLoc = FindCol("Место", hdrRow)
    colExec = FindCol("Исполнитель", hdrRow)
    If colExec = 0 Then colExec = FindCol("Подрядчик", hdrRow)

    If hdrRow = 0 Or Target.Row <= hdrRow Then GoTo Remember
    If Target.Column <> colLoc And Target.Column <> colExec Then GoTo Remember
    If Target.Address <> prevAddr Then GoTo Remember

    newVal = CStr(Target.Value)
    oldVal = prevVal

    If newVal = "" Or oldVal = "" Or InStr(newVal, SEP) > 0 Then
        result = newVal                          ' очистка, первое значение или вставка готового текста
    ElseIf Left(newVal, 3) = "Все" Or Left(oldVal, 3) = "Все" Then
        result = newVal                          ' «Все…» заменяет остальное и наоборот
    Else
        parts = Split(oldVal, SEP)
        For i = 0 To UBound(parts)
            If Trim(parts(i)) = newVal Then
                found = True                     ' повторный выбор убирает значение
            ElseIf Trim(parts(i)) <> "" Then
                If result <> "" Then result = result & SEP
                result = result & Trim(parts(i))
            End If
        Next i
        If Not found Then
            If result <> "" Then result = result & SEP
            result = result & newVal
        End If
    End If

    If result <> newVal Then
        On Error GoTo Done
        Application.EnableEvents = False
        Target.Value = result
    End If
Done:
    Application.EnableEvents = True
Remember:
    prevAddr = Target.Address
    prevVal = CStr(Target.Value)
End Sub

' Ищет колонку по точному названию заголовка в первых 5 строках
Private Function FindCol(header As String, ByRef hdrRow As Long) As Long
    Dim r As Long, c As Long
    For r = 1 To 5
        For c = 1 To 30
            If StrComp(Trim(CStr(Me.Cells(r, c).Value)), header, vbTextCompare) = 0 Then
                FindCol = c
                hdrRow = r
                Exit Function
            End If
        Next c
    Next r
End Function

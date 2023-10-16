Attribute VB_Name = "Module1"
Sub 登錄成績()
Attribute 登錄成績.VB_ProcData.VB_Invoke_Func = " \n14"
'
' 登錄成績 巨集
'

'
    Dim cell As Range
    
    Dim ID As String
    
    
    ID_former = InputBox("前半ID:")
    
    While (StrPtr(ID_former) <> 0) And (ID_former <> vbNullString)
        ID_latter = InputBox(ID_former & vbNewLine & "後半ID:")
        Do While (StrPtr(ID_latter) <> 0) And (ID_latter <> vbNullString)
            
            
            
            ID = ID_former & ID_latter
            
            Set cell = Range("B2:B200").Find(ID, LookIn:=xlValues, LookAt:=xlPart)
            
            If Not cell Is Nothing Then
                cell.Interior.Color = RGB(220, 130, 130)
                cell.Borders.ColorIndex = 3
                
                
                cell.Select
                grade = InputBox(ActiveCell.Offset(0, -1) & vbNewLine & "成績:")
                
                If StrPtr(grade) <> Null Or grade <> vbNullString Then
                
                
                grade = CDbl(grade)
                ActiveCell.Offset(0, 1) = grade
                
                
                End If
                
                cell.Interior.Color = xlColorIndexNone
                cell.Borders.ColorIndex = xlColorIndexNone

                'MsgBox "位置：" & cell.Address & vbNewLine & "內容：" & cell.Value
                
            Else
                MsgBox "找不到學號"
            
            End If
            ID_latter = InputBox(ID_former & vbNewLine & "後半ID:")
        Loop
        
        ID_former = InputBox("前半ID:")
    Wend
End Sub

Attribute VB_Name = "Module1"
Sub �n�����Z()
Attribute �n�����Z.VB_ProcData.VB_Invoke_Func = " \n14"
'
' �n�����Z ����
'

'
    Dim cell As Range
    
    Dim ID As String
    
    
    ID_former = InputBox("�e�bID:")
    
    While (StrPtr(ID_former) <> 0) And (ID_former <> vbNullString)
        ID_latter = InputBox(ID_former & vbNewLine & "��bID:")
        Do While (StrPtr(ID_latter) <> 0) And (ID_latter <> vbNullString)
            
            
            
            ID = ID_former & ID_latter
            
            Set cell = Range("B2:B200").Find(ID, LookIn:=xlValues, LookAt:=xlPart)
            
            If Not cell Is Nothing Then
                cell.Interior.Color = RGB(220, 130, 130)
                cell.Borders.ColorIndex = 3
                
                
                cell.Select
                grade = InputBox(ActiveCell.Offset(0, -1) & vbNewLine & "���Z:")
                
                If StrPtr(grade) <> Null Or grade <> vbNullString Then
                
                
                grade = CDbl(grade)
                ActiveCell.Offset(0, 1) = grade
                
                
                End If
                
                cell.Interior.Color = xlColorIndexNone
                cell.Borders.ColorIndex = xlColorIndexNone

                'MsgBox "��m�G" & cell.Address & vbNewLine & "���e�G" & cell.Value
                
            Else
                MsgBox "�䤣��Ǹ�"
            
            End If
            ID_latter = InputBox(ID_former & vbNewLine & "��bID:")
        Loop
        
        ID_former = InputBox("�e�bID:")
    Wend
End Sub

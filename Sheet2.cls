VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_Change(ByVal Target As Range)
    'Add date and money to date'
    If ((Target.Column = 6) Or (Target.Column = 11)) And (Target.Row > 1 And Target.Row < 999) Then
    
        Target.Offset(0, -2).Value = Date
        
        If (Target.Column = 6) Then
            Target.Offset(0, 1).Value = Cells(4, 1).Value
        Else
            Target.Offset(0, 1).Value = Cells(6, 1).Value
        End If
        
        If IsEmpty(Target.Offset(0, -1).Value) Then
            Target.Offset(0, -1).Value = "Automatically Generated"
        End If
        
    End If
End Sub


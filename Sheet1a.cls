VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub VALUE()
Dim x As Integer, xd As Integer
    fb = "FB"
    bp = "BP"
xd = Range("D65534").End(xlUp).Row
For x = 2 To xd
    If Range("D" & x).VALUE <> "" And InStr(1, Range("E" & x).VALUE, bp) <> 0 Then
        Range("M" & x).VALUE = "型材"
    Else
        If Range("D" & x).VALUE <> "" And InStr(1, Range("E" & x).VALUE, fb) <> 0 Then
            Range("M" & x).VALUE = "数切"
        Else
            If Range("D" & x).VALUE <> "" And Range("E" & x).VALUE = "Assembly" Then
                Range("M" & x).VALUE = ""
            Else
                If Range("D" & x).VALUE = "" Or Range("D" & x).VALUE = "COG:" Then
                    Range("M" & x).VALUE = ""
                Else
                    Range("M" & x).VALUE = "数切"
                End If
            End If
        End If
    End If
    
    
Next
End Sub

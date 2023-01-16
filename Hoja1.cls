VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Hoja1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Ines_Click()

Dim contador As Integer

contador = 0


For i = 2 To 2000 Step 1
    If Range("d" & i) = "INES" Then
        contador = contador + 1
    End If
Next
    

MsgBox "Ha trabajado " & contador & " veces."


End Sub

Private Sub Ivan_Click()

Dim contador As Integer

contador = 0


For i = 2 To 2000 Step 1
    If Range("d" & i) = "IVAN" Then
        contador = contador + 1
    End If
Next
    

MsgBox "Ha trabajado " & contador & " veces."


End Sub

Private Sub Jose_Click()

Dim contador As Integer

contador = 0


For i = 2 To 2000 Step 1
    If Range("d" & i) = "JOSE" Then
        contador = contador + 1
    End If
Next
    

MsgBox "Ha trabajado " & contador & " veces."


End Sub

Private Sub Juanfran_Click()

Dim contador As Integer

contador = 0


For i = 2 To 2000 Step 1
    If Range("d" & i) = "JUANFRAN" Then
        contador = contador + 1
    End If
Next
    

MsgBox "Ha trabajado " & contador & " veces."


End Sub

Private Sub Laura_Click()

Dim contador As Integer

contador = 0


For i = 2 To 2000 Step 1
    If Range("d" & i) = "LAURA" Then
        contador = contador + 1
    End If
Next
    

MsgBox "Ha trabajado " & contador & " veces."


End Sub

Private Sub Melvin_Click()

Dim contador As Integer

contador = 0


For i = 2 To 2000 Step 1
    If Range("d" & i) = "MELVIN" Then
        contador = contador + 1
    End If
Next
    

MsgBox "Ha trabajado " & contador & " veces."

End Sub

Private Sub Pablo_Click()

Dim contador As Integer

contador = 0


For i = 2 To 2000 Step 1
    If Range("d" & i) = "PABLO" Then
        contador = contador + 1
    End If
Next
    

MsgBox "Ha trabajado " & contador & " veces."


End Sub

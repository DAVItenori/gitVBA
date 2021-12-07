VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   8220
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13005
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btCancelar_Click()

End Sub

Private Sub btDeletar_Click()

Dim nlin As Integer

If btEditar.Value = True Then
    nlin = ListBox1.ListIndex
    If nlin = -1 Then
        MsgBox "Selecione um item para deletar!"
        Exit Sub
        ElseIf ListBox1.Value = 0 Then
        MsgBox "Selecione um item para deletar!"
        Exit Sub
   End If
    Call Deletar
Else
MsgBox "Coloque no modo edição!"
End If
End Sub

Private Sub btok_Click()

Dim nlin As Integer

If btEditar.Value = True Then
    nlin = ListBox1.ListIndex
    If nlin = -1 Then
        MsgBox "Selecione um item para editar!"
        Exit Sub
        ElseIf ListBox1.Value = 0 Then
        MsgBox "Selecione um item para editar!"
        Exit Sub
   End If
    Call Editar
Else
    Call Inserir
End If




End Sub

Private Sub ComboBox1_Change()

End Sub

Private Sub Label10_Click()

End Sub

Private Sub TextBox3_Change()

End Sub

Private Sub TextBox4_Change()

End Sub

Private Sub TextBox5_Change()

End Sub

Private Sub TextBox6_Change()

End Sub

Private Sub ListBox1_Change()

Dim nlin As Integer
nlin = ListBox1.ListIndex
If nlin = -1 Then Exit Sub
If bloqueado = True Then Exit Sub
If ListBox1.Value = 0 Then

txtUnidade.Value = ""
txtNome.Value = ""
txtValor.Value = ""
txtDataRef.Value = ""
txtDataAcordo.Value = ""
TextDia.Value = ""
txtMes.Value = ""
txtAno.Value = ""
txtParcela.Value = ""

Else

txtUnidade.Value = ListBox1.List(nlin, 0)
txtNome.Value = ListBox1.List(nlin, 1)
txtValor.Value = ListBox1.List(nlin, 2)
txtDataRef.Value = FormatDateTime(ListBox1.List(nlin, 3), vbShortDate)
txtDataAcordo.Value = FormatDateTime(ListBox1.List(nlin, 4), vbShortDate)
TextDia.Value = ListBox1.List(nlin, 5)
txtMes.Value = ListBox1.List(nlin, 6)
txtAno.Value = ListBox1.List(nlin, 7)
txtParcela.Value = ListBox1.List(nlin, 8)

End If
End Sub

Private Sub Registro_Click()

End Sub

Private Sub txtValor_Change()

End Sub

Private Sub UserForm_Initialize()
Call Atualizar_ListBox



End Sub

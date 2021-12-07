Attribute VB_Name = "Módulo2"
Option Explicit
Global bloqueado As Boolean


Sub Inserir()
bloqueado = True
Dim tabela As ListObject
Dim n As Integer, Id As Integer



Set tabela = Planilha1.ListObjects(1)
Id = Range("Id").Value


n = tabela.Range.Rows.Count
tabela.Range(n, 10).Value = Id

tabela.Range(n, 1).Value = UserForm1.txtUnidade.Value
tabela.Range(n, 2).Value = UserForm1.txtNome.Value
tabela.Range(n, 3).Value = UserForm1.txtValor.Value
tabela.Range(n, 4).Value = UserForm1.txtDataRef.Value
tabela.Range(n, 5).Value = UserForm1.txtDataAcordo.Value
tabela.Range(n, 6).Value = UserForm1.TextDia.Value
tabela.Range(n, 7).Value = UserForm1.txtMes.Value
tabela.Range(n, 8).Value = UserForm1.txtAno.Value
tabela.Range(n, 9).Value = UserForm1.txtParcela.Value

UserForm1.ListBox1.RowSource = ""
tabela.ListRows.Add
Range("Id").Value = Id + 1

Call Atualizar_ListBox
Call LimparCampos
MsgBox "Cadastrado com sucesso!", vbInformation, "Informação"
bloqueado = False

End Sub

Sub Editar()
bloqueado = True
Dim tabela As ListObject
Dim n As Integer, l As Integer

Set tabela = Planilha1.ListObjects(1)

n = UserForm1.ListBox1.Value
l = tabela.Range.Columns().Find(n, , , xlWhole).Row

tabela.Range(l, 1).Value = UserForm1.txtUnidade.Value
tabela.Range(l, 2).Value = UserForm1.txtNome.Value
tabela.Range(l, 3).Value = UserForm1.txtValor.Value
tabela.Range(l, 4).Value = UserForm1.txtDataRef.Value
tabela.Range(l, 5).Value = UserForm1.txtDataAcordo.Value
tabela.Range(l, 6).Value = UserForm1.TextDia.Value
tabela.Range(l, 7).Value = UserForm1.txtMes.Value
tabela.Range(l, 8).Value = UserForm1.txtAno.Value
tabela.Range(l, 9).Value = UserForm1.txtParcela.Value

Call Atualizar_ListBox
Call LimparCampos
MsgBox "O Registro foi atualizado"
bloqueado = False


End Sub
Sub Atualizar_ListBox()
bloqueado = True
Dim tabela As ListObject
Set tabela = Planilha1.ListObjects(1)

UserForm1.ListBox1.RowSource = tabela.DataBodyRange.Address(, , , True)

bloqueado = False
End Sub

Sub Deletar()
bloqueado = True
Dim n As Integer, l As Integer
Dim tabela As ListObject


Set tabela = Planilha1.ListObjects(1)

n = UserForm1.ListBox1.Value
l = tabela.Range.Columns().Find(n, , , xlWhole).Row

UserForm1.ListBox1.RowSource = ""
tabela.Range.Rows(l).Delete

Call Atualizar_ListBox
MsgBox "O Registro foi DELETADO"
bloqueado = False

End Sub


Sub LimparCampos()

UserForm1.txtUnidade.Value = ""
UserForm1.txtNome.Value = ""
UserForm1.txtValor.Value = ""
UserForm1.txtDataRef.Value = ""
UserForm1.txtDataAcordo.Value = ""
UserForm1.TextDia.Value = ""
UserForm1.txtMes.Value = ""
UserForm1.txtAno.Value = ""
UserForm1.txtParcela.Value = ""

End Sub

Sub EXIBIR()
UserForm1.Show

End Sub

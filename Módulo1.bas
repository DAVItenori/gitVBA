Attribute VB_Name = "Módulo1"

Sub UserForm1()
Const wdreplaceall = 2

Set objWord = CreateObject("Word.Application")

objWord.Visible = True

Set arqRecibo = objWord.documents.Open(ThisWorkbook.Path & "\HONORARIO_REFERENCIA.docx")
Set conteudoDoc = arqRecibo.Application.Selection

For i = 1 To 9

    conteudoDoc.Find.Text = Cells(1, i).Value
    conteudoDoc.Find.Replacement.Text = Cells(2, i).Value
    conteudoDoc.Find.Execute Replace:=wdreplaceall
Next

arqRecibo.saveas2 ("C:\Users\asus\Desktop\Sistema de Cadastro\Recibos\ReciboHonorarios - " & Cells(2, 2).Value & ".docx")
arqRecibo.Close
objWord.Quit

Set objWord = Nothing
Set arqRecibo = Nothing
Set conteudoDoc = Nothing


End Sub



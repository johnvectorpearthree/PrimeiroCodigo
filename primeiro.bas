Attribute VB_Name = "M�dulo1"
Sub primeiro()
'O comando DIM(Dimension) � utilizado para declarar var�avel.
'A var�avel nome foi tipada como String(texto).
Dim nome As String

'O comando InputBox abre uma caixa de entrada de dados, assim o us�ario digita o nome e aloca na var�avel nome.
nome = InputBox("Digite o seu nome")

'O comando Range permite selecionar uma c�lula na planilha do Excel,
'assim selecionamos a c�lula "A1" e adicionamos o valor que foi digitado na caixa de entrada,
'usando a var�avel nome.
Range("A1").Value = nome
End Sub

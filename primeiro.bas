Attribute VB_Name = "Módulo1"
Sub primeiro()
'O comando DIM(Dimension) é utilizado para declarar varíavel.
'A varíavel nome foi tipada como String(texto).
Dim nome As String

'O comando InputBox abre uma caixa de entrada de dados, assim o usúario digita o nome e aloca na varíavel nome.
nome = InputBox("Digite o seu nome")

'O comando Range permite selecionar uma célula na planilha do Excel,
'assim selecionamos a célula "A1" e adicionamos o valor que foi digitado na caixa de entrada,
'usando a varíavel nome.
Range("A1").Value = nome
End Sub

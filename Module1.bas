Attribute VB_Name = "Module1"
'Declara a função que lê o arquivo ini.
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal NomeSeção As String, ByVal NomeItem As String, ByVal ValorDefault As String, ByVal ValorRetornado As String, ByVal TamanhoBuffer As Long, ByVal NomeArquivoINI As String) As Long
'Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Atualizando As Integer

'Declara a função que escreve o arquivo ini.
Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

Public Sub CentraForm(Formulario As Form)
'Esta rotina centraliza qualquer formulário que é passado para ela como "Form".

    Formulario.Move (Screen.Width \ 2) - (Formulario.Width \ 2), (Screen.Height \ 2) - (Formulario.Height \ 2)

End Sub

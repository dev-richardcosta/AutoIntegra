Attribute VB_Name = "modArquivos"
Option Explicit

Public Sub PrepararEstruturaPastas()
    On Error GoTo Falhou
    Dim fso As Object, pasta As Variant
    Set fso = CreateObject("Scripting.FileSystemObject")
    For Each pasta In Array(PASTA_MODELOS, PASTA_TEMP, PASTA_GERADOS, PASTA_CONFIG, PASTA_LOGS)
        If Not fso.FolderExists(CaminhoBase() & "\" & CStr(pasta)) Then fso.CreateFolder CaminhoBase() & "\" & CStr(pasta)
    Next
    Exit Sub
Falhou: MostrarErro "Não foi possível preparar as pastas", Err.Number, Err.Description: Err.Raise Err.Number
End Sub

Public Sub LimparPastaTemp()
    On Error GoTo Falhou
    Dim fso As Object, arquivo As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    For Each arquivo In fso.GetFolder(CaminhoBase() & "\" & PASTA_TEMP).Files: fso.DeleteFile arquivo.Path, True: Next
    RegistrarLog "Pasta Temp limpa."
    Exit Sub
Falhou: MostrarErro "Não foi possível limpar Temp", Err.Number, Err.Description: Err.Raise Err.Number
End Sub

Public Function CopiarModelosParaTemp() As Collection
    On Error GoTo Falhou
    Dim fso As Object, arquivo As Object, saida As New Collection
    Set fso = CreateObject("Scripting.FileSystemObject")
    For Each arquivo In fso.GetFolder(CaminhoBase() & "\" & PASTA_MODELOS).Files
        If EhModeloWord(CStr(arquivo.Name)) Then
            fso.CopyFile arquivo.Path, CaminhoBase() & "\" & PASTA_TEMP & "\" & arquivo.Name, True
            saida.Add CaminhoBase() & "\" & PASTA_TEMP & "\" & arquivo.Name
            RegistrarLog "Modelo copiado: " & arquivo.Name
        End If
    Next
    Set CopiarModelosParaTemp = saida
    Exit Function
Falhou: MostrarErro "Não foi possível copiar os modelos", Err.Number, Err.Description: Err.Raise Err.Number
End Function

Public Function EhModeloWord(ByVal nomeArquivo As String) As Boolean
    Dim extensao As String: extensao = LCase$(Mid$(nomeArquivo, InStrRev(nomeArquivo, ".")))
    EhModeloWord = (extensao = ".docx" Or extensao = ".docm" Or extensao = ".dotx")
End Function

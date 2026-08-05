Attribute VB_Name = "modMain"
Option Explicit

Public Sub GerarDocumentos()
    On Error GoTo Falhou
    Dim caminhos As Collection, campos As Object, valores As Object, exportarPDF As Boolean
    Application.ScreenUpdating = False: Application.DisplayAlerts = wdAlertsNone
    PrepararEstruturaPastas: LimparPastaTemp
    Set caminhos = CopiarModelosParaTemp()
    If caminhos.Count = 0 Then MsgBox "Nenhum modelo Word foi encontrado na pasta '" & PASTA_MODELOS & "'.", vbInformation, NOME_PROJETO: GoTo Encerrar
    Set campos = LocalizarMarcadores(caminhos)
    If campos.Count = 0 Then MsgBox "Nenhum marcador {{CAMPO}} foi encontrado.", vbInformation, NOME_PROJETO: GoTo Encerrar
    Set valores = SolicitarValores(campos)
    exportarPDF = frmPrincipal.ExportarPDFSelecionado
    Unload frmPrincipal
    If valores Is Nothing Then GoTo Encerrar
    GerarDocumentosPreenchidos caminhos, valores, exportarPDF
    MsgBox CStr(caminhos.Count) & " documento(s) gerado(s) em '" & PASTA_GERADOS & "'.", vbInformation, NOME_PROJETO
Encerrar:
    Application.DisplayAlerts = wdAlertsAll: Application.ScreenUpdating = True
    Exit Sub
Falhou: MostrarErro "Gerador de documentos", Err.Number, Err.Description: Resume Encerrar
End Sub

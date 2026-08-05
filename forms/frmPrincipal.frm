VERSION 5.00
Begin VB.UserForm frmPrincipal
   Caption         =   "Gerador de Documentos"
   ClientHeight    =   7650
   ClientWidth     =   5850
   StartUpPosition =   1  'CenterOwner
End
Begin MSForms.CheckBox chkExportarPDF
   Caption         =   "Exportar também em PDF"
   Height          =   255
   Left            =   180
   TabIndex        =   0
   Top             =   6600
   Width           =   2200
End
Begin MSForms.CommandButton cmdGerar
   Caption         =   "GERAR DOCUMENTOS"
   Height          =   375
   Left            =   180
   TabIndex        =   1
   Top             =   7020
   Width           =   1800
End
Begin MSForms.CommandButton cmdCancelar
   Caption         =   "Cancelar"
   Height          =   375
   Left            =   2070
   TabIndex        =   2
   Top             =   7020
   Width           =   1200
End
Attribute VB_Name = "frmPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mCampos As Object, mValores As Object, mConfirmado As Boolean, mMontado As Boolean

Public Sub Configurar(ByVal campos As Object)
    Set mCampos = campos: Set mValores = CriarDicionario(): mConfirmado = False: mMontado = False
End Sub

Public Function ObterValores() As Object
    If mConfirmado Then Set ObterValores = mValores
End Function

Public Property Get ExportarPDFSelecionado() As Boolean
    ExportarPDFSelecionado = Me.chkExportarPDF.Value
End Property

Private Sub UserForm_Initialize()
    Me.Width = LARGURA_FORMULARIO: Me.Height = ALTURA_FORMULARIO
    Me.ScrollBars = fmScrollBarsVertical: Me.KeepScrollBarsVisible = fmScrollBarsVertical
End Sub

Private Sub UserForm_Activate()
    If Not mMontado And Not mCampos Is Nothing Then MontarCampos
End Sub

Private Sub MontarCampos()
    Dim chave As Variant, etiqueta As MSForms.Label, caixa As MSForms.TextBox
    Dim y As Single
    y = 12
    For Each chave In mCampos.Keys
        Set etiqueta = Me.Controls.Add("Forms.Label.1", "lbl_" & CStr(chave), True)
        etiqueta.Caption = FormatarRotulo(CStr(chave)): etiqueta.Left = 12: etiqueta.Top = y: etiqueta.Width = 150
        Set caixa = Me.Controls.Add("Forms.TextBox.1", "txt_" & CStr(chave), True)
        caixa.Left = 168: caixa.Top = y - 3: caixa.Width = 180: caixa.Height = 18
        y = y + ALTURA_CAMPO
    Next
    Me.chkExportarPDF.Left = 12: Me.chkExportarPDF.Top = y + 6
    Me.cmdGerar.Left = 12: Me.cmdGerar.Top = y + 34
    Me.cmdCancelar.Left = 183: Me.cmdCancelar.Top = y + 34
    Me.ScrollHeight = y + 78
    mMontado = True
End Sub

Private Sub cmdGerar_Click()
    On Error GoTo Falhou
    Dim chave As Variant, caixa As MSForms.TextBox
    For Each chave In mCampos.Keys
        Set caixa = Me.Controls("txt_" & CStr(chave)): mValores(CStr(chave)) = caixa.Text
    Next
    mConfirmado = True: Me.Hide
    Exit Sub
Falhou: MostrarErro "Falha ao ler o formulário", Err.Number, Err.Description
End Sub

Private Sub cmdCancelar_Click()
    mConfirmado = False: Me.Hide
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then Cancel = True: mConfirmado = False: Me.Hide
End Sub

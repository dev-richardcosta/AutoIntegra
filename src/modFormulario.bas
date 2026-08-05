Attribute VB_Name = "modFormulario"
Option Explicit

Public Function SolicitarValores(ByVal campos As Object) As Object
    frmPrincipal.Configurar campos
    frmPrincipal.Show vbModal
    Set SolicitarValores = frmPrincipal.ObterValores
End Function

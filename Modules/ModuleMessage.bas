Attribute VB_Name = "ModMessage"
Option Explicit

Public Function ErrorHandler(ByVal NumberError_ As Long, ByVal ErrorDescription_ As String, ByVal ErrorPlace As String, Optional ByVal Query_ As String)
Dim iMessage As String

iMessage = "-------------------------------------------------------------------------------" & vbCrLf
iMessage = iMessage & "                                         GoInvest" & vbCrLf
iMessage = iMessage & "-------------------------------------------------------------------------------" & vbCrLf
iMessage = iMessage & "Error #" & NumberError_ & ": '" & ErrorDescription_ & vbCrLf
iMessage = iMessage & "-------------------------------------------------------------------------------" & vbCrLf
iMessage = iMessage & Query_ & vbCrLf
iMessage = iMessage & "-------------------------------------------------------------------------------" & vbCrLf

MsgBox iMessage, vbCritical, "GoInvest"
End Function

Public Function FillFooter() As String
Dim iReturn As String

iReturn = " - Servidor: " & pServidorBanco & " | Banco de Dados: " & pNomeBanco & " | Usu�rio: " & pUsuarioBanco & " | V." & pVersao & " | "

FillFooter = iReturn
End Function

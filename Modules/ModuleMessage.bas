Attribute VB_Name = "ModuleMessage"
Option Explicit

Public Function ErrorHandler(ByVal NumberError_ As String, ByVal ErrorDescription_ As String, Optional ByVal Query_ As String) As String
Dim sRetorno As String

sRetorno = "-------------------------------------------------------------------------------" & vbCrLf
sRetorno = sRetorno & "                                         GoInvest" & vbCrLf
sRetorno = sRetorno & "-------------------------------------------------------------------------------" & vbCrLf
sRetorno = sRetorno & "Error #" & NumberError_ & ": '" & ErrorDescription_ & vbCrLf
sRetorno = sRetorno & "-------------------------------------------------------------------------------" & vbCrLf
sRetorno = sRetorno & Query_ & vbCrLf
sRetorno = sRetorno & "-------------------------------------------------------------------------------" & vbCrLf

ErrorHandler = sRetorno
End Function

Public Function FillFooter() As String
Dim iReturn As String

iReturn = "| Servidor: " & gServer & " | Banco de Dados: " & gConnection & " | Usuário: " & gUser & " | V." & gVersion & " | "

FillFooter = iReturn
End Function

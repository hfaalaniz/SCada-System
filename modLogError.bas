Attribute VB_Name = "mdlErrores"
Option Explicit

Public Sub Err_Handler(Optional ByVal vblnDisplayError As Boolean = True, _
                       Optional ByVal vstrErrNumber As String = vbNullString, _
                       Optional ByVal vstrErrDescription As String = vbNullString, _
                       Optional ByVal vstrModuleName As String = vbNullString, _
                       Optional ByVal vstrProcName As String = vbNullString)

  Dim strTemp As String
  Dim lngFN   As Long

   On Error Resume Next
   '// Purpose: Error handling - On Error

   '// Show Error Message
   If vblnDisplayError Then
      strTemp = "Error ocurrido: "
      If LenB(vstrErrNumber) Then strTemp = strTemp & vstrErrNumber & vbNewLine Else strTemp = strTemp & vbNewLine
      If LenB(vstrErrDescription) Then strTemp = strTemp & "Descripción: " & vstrErrDescription & vbNewLine
      If LenB(vstrModuleName) Then strTemp = strTemp & "Módulo: " & vstrModuleName & vbNewLine
      If LenB(vstrProcName) Then strTemp = strTemp & "Función: " & vstrProcName
      MsgBox strTemp, vbCritical, App.Title & " - ERROR"
   End If

   '// Write error log
   lngFN = FreeFile
   Open App.Path & "\ErrorLog.txt" For Append As #lngFN
   Write #lngFN, Now, vstrErrNumber, vstrErrDescription, vstrModuleName, vstrProcName, _
      App.Title & " v" & App.Major & "." & App.Minor & "." & App.Revision, _
      Environ("username"), Environ("computername")
   Close #lngFN

End Sub

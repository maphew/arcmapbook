Attribute VB_Name = "ErrorHandling"
Option Explicit
'
' FILE AUTOMATICALLY GENERATED BY ESRI ERROR HANDLER ADDIN
' DO NOT EDIT OR REMOVE THIS FILE FROM THE PROJECT
'
Dim pErrorLog As New ErrorHandlerUI.ErrorDialog


Private Sub DisplayVersion2Dialog(sProcedureName As String, sErrDescription As String)
  Beep
  MsgBox "An error has occured in the application.   Record the call stack sequence" & vbCrLf & "and the description of the error." & vbCrLf & vbCrLf & _
         "Error Call Stack Sequence " & vbCrLf & vbTab & sProcedureName & vbCrLf & sErrDescription, vbExclamation + vbOKOnly, "Unexpected Program Error"
End Sub

Private Sub DisplayVersion3Dialog(sProcedureName As String, sErrDescription As String, parentHWND As Long, raiseException As Boolean)
  Beep
  MsgBox "An error has occured in the application.   Record the call stack sequence" & vbCrLf & "and the description of the error." & vbCrLf & vbCrLf & _
         "Error Call Stack Sequence " & vbCrLf & vbTab & sProcedureName & vbCrLf & sErrDescription, vbExclamation + vbOKOnly, "Unexpected Program Error"
End Sub

Private Sub DisplayVersion4Dialog(sProcedureName As String, sErrDescription As String, parentHWND As Long)
  pErrorLog.AppendErrorText "Record Call Stack Sequence - Bottom line is error line." & vbCrLf & vbCrLf & vbTab & sProcedureName & vbCrLf & sErrDescription
  pErrorLog.Visible = True


End Sub

Public Sub HandleError(ByVal bTopProcedure As Boolean, _
                       ByVal sProcedureName As String, _
                       ByVal lErrNumber As Long, _
                       ByVal sErrSource As String, _
                       ByVal sErrDescription As String, _
                       Optional ByVal version As Long = 1, _
                       Optional ByVal parentHWND As Long = 0, _
                       Optional ByVal reserved1 As Variant = 0, _
                       Optional ByVal reserved2 As Variant = 0, _
                       Optional ByVal reserved3 As Variant = 0)
  ' Generic Error handling Function - This function should be called with
  ' the following Arguments
  '
  ' Boolean    -in-  True if called from a top level procedure - Event / Method / Property
  ' String     -in-  Name of function called from
  ' Long       -in-  Error Number (retrieved from Err object)
  ' String     -in-  Error Source (retrieved from Err object)
  ' String     -in-  Error Description (retrieved from Err object)
  ' Long       -in-  Version of Function (optional Default 1)
  ' parentHWND -in-  Parent Hwnd for error dialogs, NULL is valid
  ' reserved1  -in-
  ' reserved2  -in-
  ' reserved3  -in-
  
  
  ' Clear the error object
  Err.Clear

  ' Static variable used to control the call stack formatting
  Static entered As Boolean

  If (bTopProcedure) Then
    ' Top most procedure in call stack so report error to user
    ' Via a dialog
    If (Not entered) Then
      sErrDescription = vbCrLf & "Error Number " & vbCrLf & vbTab & CStr(lErrNumber) & vbCrLf & "Description" & vbCrLf & vbTab & sErrDescription & vbCrLf & vbCrLf
    End If
    entered = False
    If (version = 4) Then
      DisplayVersion4Dialog sProcedureName, sErrDescription, parentHWND
    ElseIf (version = 3) Then
      Dim raiseError As Boolean
      DisplayVersion3Dialog sProcedureName, sErrDescription, parentHWND, raiseError
      If (raiseError) Then Err.Raise lErrNumber, sErrSource, vbTab & sProcedureName & vbCrLf & sErrDescription
    ElseIf (version = 2) Then
      DisplayVersion2Dialog sProcedureName, sErrDescription
    Else
      Beep
      MsgBox "An error has occured in the application.   Record the call stack sequence" & vbCrLf & "and the description of the error." & vbCrLf & vbCrLf & _
             "Error Call Stack Sequence " & vbCrLf & vbTab & sProcedureName & vbCrLf & sErrDescription, vbExclamation + vbOKOnly, "Unexpected Program Error"
    End If
  Else
    ' An error has occured but we are not at the top of the call stack
    ' so append the callstack and raise another error
    If (Not entered) Then sErrDescription = vbCrLf & "Error Number " & vbCrLf & vbTab & CStr(lErrNumber) & vbCrLf & "Description" & vbCrLf & vbTab & sErrDescription & vbCrLf & vbCrLf
    entered = True
    Err.Raise lErrNumber, sErrSource, vbTab & sProcedureName & vbCrLf & sErrDescription
  End If
End Sub

Public Function GetErrorLineNumberString(ByVal lLineNumber As Long) As String
  ' Test the line number if it is non zero create a string
  If (lLineNumber <> 0) Then GetErrorLineNumberString = "Line : " & lLineNumber
End Function

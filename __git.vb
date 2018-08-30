Attribute VB_Name = "__git"
'***************************************************************************
'Module __git
'***************************************************************************
'Purpose: Export als VBA-Code to files to have them in GIT
'
'Author: Markus Grabosch, 2018
'Version: 1.0
'
'Reference:
'You must reference the "Microsoft Visual Basic for Applications Extensibility 5.3".
'
'Usage:
'Simply add the module to any of your VBA-Projects and put the desired path to
'the const "c_ExportPath" and run the Main-Sub everytime you want to export your code.
'***************************************************************************
Option Compare Database
Option Explicit

'***************************************************************************
'CONST
Private Const c_ExportPath As String = "E:\path\to\my\project" 'ExportPath, must end with "\"
'***************************************************************************

Public Sub ExportSourceFiles()
'***************************************************************************
'Purpose: Main-Sub to start export process
'Input: N/A
'Output: N/A
'
'Author: Markus Grabosch, 2018
'***************************************************************************

Dim Module As VBComponent
Dim Suffix As String

For Each Module In Application.VBE.ActiveVBProject.VBComponents
  Select Case Module.Type
    Case vbext_ct_ClassModule
      Suffix = ".cls"
    Case vbext_ct_StdModule
      Suffix = ".vb"
    Case vbext_ct_MSForm
      Suffix = ".frm"
    Case 100 'Reports and Forms share the samt type...
      If Left(Module.Name, 5) = "Form_" Then
        Suffix = ".frm"
      ElseIf Left(Module.Name, 7) = "Report_" Then
        Suffix = ".rep"
      Else
        Suffix = ".bas"
      End If
    Case Else
      Suffix = ".vb"
  End Select
  Module.Export c_ExportPath + Module.Name + Suffix
Next
 
End Sub


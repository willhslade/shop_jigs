Attribute VB_Name = "exportAllBAS"
Option Explicit

Sub WriteAllBas()
' Write all VBA modules as .bas files to the directory of ThisWorkbook.
' Implemented to make version control work smoothly for identifying changes.
' Designed to be called every time this workbook is saved,
'   if code has changed, then will show up as a diff
'   if code has not changed, then file will be same (no diff) with new date.
' Following https://stackoverflow.com/questions/55956116/mass-importing-modules-references-in-vba
'            which references https://www.rondebruin.nl/win/s9/win002.htm

Dim cmp As VBComponent, cmo As CodeModule
Dim fn As Integer, outName As String
Dim sLine As String, nLine As Long
Dim dirExport As String, outExt As String
Dim fileExport As String
Dim filePath As String

   On Error GoTo MustTrustVBAProject
   Set cmp = ThisWorkbook.VBProject.VBComponents(1)
   On Error GoTo 0
   
    filePath = ThisWorkbook.path
    If UCase(Left(filePath, 4)) = "HTTP" Then
        filePath = GetLocalPath(filePath)
    End If
   
   dirExport = filePath + Application.PathSeparator + "VBA" + Application.PathSeparator
   For Each cmp In ThisWorkbook.VBProject.VBComponents
      Select Case cmp.Type
         Case vbext_ct_ClassModule:
            outExt = ".cls"
         Case vbext_ct_MSForm
            outExt = ".frm"
         Case vbext_ct_StdModule
            outExt = ".bas"
         Case vbext_ct_Document
            Set cmo = cmp.CodeModule
            If Not cmo Is Nothing Then
               If cmo.CountOfLines = cmo.CountOfDeclarationLines Then ' Ordinary worksheet or Workbook, no code
                  outExt = ""
               Else ' It's a Worksheet or Workbook but has code, export it
                  outExt = ".cls"
               End If
            End If ' cmo Is Nothing
         Case Else
            Stop ' Debug it
      End Select
      
      If outExt <> "" Then
'        If Len(Dir(dirExport)) = 0 Then
'            MkDir dirExport
'        End If
         fileExport = dirExport + cmp.name + outExt
         If Dir(fileExport) <> "" Then Kill fileExport   ' From Office 365, Export method does not overwrite existing file
         cmp.Export fileExport
      End If
   Next cmp
   Exit Sub
    
MustTrustVBAProject:
   MsgBox "Must trust VB Project in Options, Trust Center, Trust Center Settings ...", vbCritical + vbOKOnly, "WriteAllBas"
End Sub

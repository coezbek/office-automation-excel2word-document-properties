Const e2wModeExport As Integer = 1        ' Use this integer constant if you want the ExportWordFromShape method to make a copy of the template.
Const e2wModeSetupTemplate As Integer = 2 ' Use this integer constant if you want to insert DocProperty example fields in the template.

Option Explicit                           ' This will make VBA raise an error if you haven't defined a variable

'
' This Macro is the entry point to the Excel2Word module. You should call it from a button and customize
' the code below to inform it about the following:
'
'  1.) Where to find the word document that you want to use as a template for exporting. This should be the name of an OLE object that contains a Word file.
'  2.) Which data you want to export. This can either be cell references such as B5 or Named Ranges you defined.
'
' The default configuration looks for a shape called "TemplateShape" and will export all named ranges to Word.
'
Sub ExportExcel2Word()
     
  Call ExportWordFromShape(ActiveSheet.Shapes("TemplateShape"), Join(NamedRangesNames(), ", "), e2wModeExport)
  
End Sub

'
' Call this to prepare the word template stored in the "TemplateShape" with the same properties that will be exported.
'
' If you make any changes to the origin of the document shape or use other
' names to export, also make the same change below.
'
Sub SetupWordTemplate()
      
  Call ExportWordFromShape(ActiveSheet.Shapes("TemplateShape"), Join(NamedRangesNames(), ", "), e2wModeSetupTemplate)
  
End Sub


'
' Opens the given template word document contained in `shapeContainingWord` and does the following:
'
' Depending on e2wMode:
'
'   - e2wModeExport: Make a copy of document and in this copy create/update the given fieldlist.
'   - e2wModeSetupTemplate: Open the document template itself and append the given fields
'
' In both cases at the end a document is open for editing:
'
'   - e2wModeExport: The copy for you to edit with the current values exported from Excel.
'   - e2wModeSetupTemplate: The template file stored in the given `shapeContainingWord`
'
' Also see DocumentSetDocPropFromFieldList
'
Public Sub ExportWordFromShape(shapeContainingWord As Shape, ByVal fieldlist As String, e2wMode As Integer)

    Dim oWordApp As Object
    
    ' Open Word Document from given Shape and save using a temporary name
    Dim objWordDocument As Object
    Dim objOLE As OLEObject
    shapeContainingWord.OLEFormat.Activate
    
    Set objOLE = shapeContainingWord.OLEFormat.Object
    Set objWordDocument = objOLE.Object
    
    If e2wMode = e2wModeExport Then
        Dim templateFilename As String
        
        templateFilename = GetTempFile
        objWordDocument.SaveAs2 Filename:=templateFilename
        objWordDocument.Close saveChanges:=False
        
        ' Get new word instance
        On Error Resume Next
        Set oWordApp = GetObject(, "Word.Application")
        If Err.Number <> 0 Then
            Set oWordApp = CreateObject("Word.Application")
        End If
        Err.Clear
        On Error GoTo 0
    
        ' Create new document based on the template
        Set objWordDocument = oWordApp.Documents.Add(templateFilename, NewTemplate:=False, DocumentType:=0)
    
        ' Delete file
        Kill templateFilename
            
        ' Set document title to get a default file name
        Dim dlgProp As Variant
        Set dlgProp = oWordApp.Dialogs(wdDialogFileSummaryInfo)
        dlgProp.Title = "990909 " & Format(Now, "yyyy.mm.dd") & " " & Evaluate("last_name").Text & ", " & Evaluate("first_name").Text
        dlgProp.Execute
            
        oWordApp.Visible = True
        
    End If
            
    objWordDocument.Application.ScreenUpdating = False
    DocumentSetDocPropFromFieldList objWordDocument, fieldlist, e2wMode
    objWordDocument.Application.ScreenUpdating = True
       
    Application.SendKeys ("%{TAB}")
    DoEvents
    
End Sub

'
' Get an array with the names of all named ranges in the Active Workbook.
' Internal named ranges (starting with _ are filtered out)
'
Function NamedRangesNames() As String()
    
    Dim ary() As String
    
    If ActiveWorkbook.Names.Count > 0 Then
        ReDim ary(ActiveWorkbook.Names.Count)
        
        Dim i As Long
        i = 0
            
        Dim n As Name
        For Each n In ActiveWorkbook.Names
            If Not startsWith(n.Name, "_") Then ' Exclude Build-In Fields
                ary(i) = n.Name
                i = i + 1
            End If
        Next
        
        ReDim Preserve ary(i - 1)
    End If
    
    NamedRangesNames = ary
End Function

'
' Uses the given list of fields to create/update DocumentProperties in the given document.
'
' The fieldList is split by commas and each element is interpreted using the Evaluate() VBA function.
'
' The name of the created DocumentProperties is based on the fieldlist elements with a prefix of "xls_"
'
' Depending on e2wMode this function will either:
'
'   - e2wModeExport: Just create/update the given properties
'   - e2wModeSetupTemplate: Also append fields showing these document properties to the document
'
Sub DocumentSetDocPropFromFieldList(document As Word.document, ByVal fieldlist As String, e2wMode As Integer)

  Dim fields() As String
  fields = Split(fieldlist, ",")
  
  If e2wMode = e2wModeSetupTemplate Then
    document.Content.InsertAfter Text:=Chr(13) & Chr(10) & "Fields from Excel:" & Chr(13) & Chr(10)
  End If
    
  Dim propName As Variant
  For Each propName In fields
  
    Dim propId As String
  
    propName = Trim(propName)
  
    Dim propValue As String
    propValue = Evaluate(propName).Text
    
    propId = propName
    If Not startsWith(propId, "xls_") Then
        propId = "xls_" & propId
    End If
    
    If e2wMode = e2wModeExport Then
    
        updateCustomDocumentProperty document, propId, propValue, msoPropertyTypeString
    
    ElseIf e2wMode = e2wModeSetupTemplate Then
      
        updateCustomDocumentProperty document, propId, propName & ", e.g. " & propValue, msoPropertyTypeString
        
        document.Content.InsertAfter Text:=propName & ": "
        document.Characters.Last.Select
        
        AddDocPropertyField document.ActiveWindow.Selection.Range, propId
        document.Content.InsertAfter Text:=Chr(13) & Chr(10)
    End If
  
  Next
    
  UpdateAllFields document

End Sub

'
' Add a DocProperty Field with the given name at the position of the given range (replacing it)
'
Public Sub AddDocPropertyField(ByVal r As Word.Range, ByVal propName As String)

    r.fields.Add Range:=r, _
                 Type:=wdFieldEmpty, _
                 Text:="DOCPROPERTY """ & propName & """", _
                 PreserveFormatting:=True

End Sub


'
' Updates all Fields in the given document
' Includes the main document range and all headers and footers.
' Does not update any fields which are in shapes, text boxes and potentially
' other exotic places (see for instance: https://stackoverflow.com/a/33762199)
'
Public Sub UpdateAllFields(doc As document)

  Dim section As Word.section
  Dim header As Word.HeaderFooter
  
  doc.Range.fields.Update

  For Each section In doc.Sections
    For Each header In section.Headers
      header.Range.fields.Update
    Next
  
    For Each header In section.Footers
      header.Range.fields.Update
    Next
  Next
    
End Sub


'
' Update a DocProperty in the given document with a new value or create new from given value.
'
' See https://stackoverflow.com/a/14863333
Public Sub updateCustomDocumentProperty(oDoc As document, ByVal strPropertyName As String, _
  varValue As Variant, docType As Office.MsoDocProperties)

    On Error Resume Next
    oDoc.CustomDocumentProperties(strPropertyName).Value = varValue
    
    If Err.Number > 0 Then
        On Error GoTo 0
        
        oDoc.CustomDocumentProperties.Add _
            Name:=strPropertyName, _
            LinkToContent:=False, _
            Type:=docType, _
            Value:=varValue
    End If
    
End Sub

'
' Return true if the given string `str` starts with the given prefix
' See https://stackoverflow.com/a/20805609
'
Public Function startsWith(ByVal str As String, prefix As String) As Boolean
    startsWith = Left(str, Len(prefix)) = prefix
End Function

'
' Get full path to a temp file in the temp folder
' Caution: This file does not have any particular file extension
'
Function GetTempFile() As String

    Dim FileSystemObject As Object
    Set FileSystemObject = CreateObject("Scripting.FileSystemObject")
    
    ' 2 == TemporaryFolder
    GetTempFile = FileSystemObject.GetSpecialFolder(2) & "\" & FileSystemObject.GetTempName
    
End Function

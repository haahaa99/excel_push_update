Sub AddLibrary()

    Dim strGUID As String

    'Microsoft Word GUID
    strGUID = "{00020905-0000-0000-C000-000000000046}"

    'Check if reference is already added to the project, if not add it
    If F_isReferenceAdded(strGUID) = False Then
        ThisWorkbook.VBProject.REFERENCES.AddFromGuid strGUID, 0, 0
    End If
        
End Sub
' ----------------------------------------------------------------
' Purpose: Check if an Object Library refernce is added to a VBAProject or not
' ----------------------------------------------------------------
Function F_isReferenceAdded(referenceGUID As String) As Boolean

    Dim varRef As Variant

    'Loop through VBProject references if input GUID found return TRUE otherwise FALSE
    For Each varRef In ThisWorkbook.VBProject.REFERENCES
        
        If varRef.GUID = referenceGUID Then
            F_isReferenceAdded = True
            Exit For
        End If
        
    Next varRef

End Function

Private Sub Workbook_Open()
Call AddLibrary
End Sub

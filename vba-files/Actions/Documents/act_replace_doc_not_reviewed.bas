Attribute VB_Name = "act_replace_doc_not_reviewed"


'namespace=vba-files\Actions\Documents


'/*
'
'This action is used to replace files that contain errors but do not affect the review process. The user has fixed some errors in these files
'
'*/
Function ReplaceDocFile(ByVal docReviewId As String, ByVal filePathOrigin As String) As Boolean

    Dim projectId As String
    Dim docId As String
    Dim fileName As String
    Dim revCode As String
    Dim docCode As String
    Dim docExtension As String

    ' Retrieve document information from the database
    GetDocInfoFromDatabase docReviewId, projectId, docId, docExtension, revCode, docCode

    ' Build file name from extracted fields
    fileName = docCode & "_Rev_" & revCode & "." & docExtension

    ' Create a new record for the document issuance/replacement
    CreateIssuedReplacedRecord docReviewId

    ' Move the file to the new location using the extracted document information
    MoveFileToNewLocation filePathOrigin, projectId, docId, fileName

End Function


Private Sub GetDocInfoFromDatabase(ByVal docReviewId As String, ByRef projectId As String, ByRef docId As String, ByRef docExtension As String, ByRef revCode As String, ByRef docCode As String)

    ' Retrieve document information from the database
    Dim docInfo As Variant
    docInfo = db_documents.get_doc_by_review(docReviewId)

    ' Validate if we successfully retrieve the document information
    If IsEmpty(docInfo) Then
        MsgBox "No record found for the specified document review ID."
        Exit Sub
    End If

    ' Extract relevant fields from the retrieved recordset
    docExtension = LCase(docInfo(3, 0))
    docCode = docInfo(2, 0)
    projectId = docInfo(0, 0)
    docId = docInfo(1, 0)
    revCode = docInfo(4, 0)

End Sub



Private Sub CreateIssuedReplacedRecord(ByVal docReviewId As String)

    ' Create a dictionary object to store data for document issuance/replacement record
    Dim data As Object
    Set data = CreateObject("Scripting.Dictionary")

    With data
        .Add "user_id", auth.get_user_id
        .Add "review_id", docReviewId
        .Add "replace_date", format(Date, "yyyy-mm-dd")
    End With

    ' Create a new record for the document issuance/replacement
    On Error GoTo DocumentReplacementError
    db_documents_issued_replaced.Create data
    db_documents.update_review_status docReviewId, Constants.REVIEW_SATUS_SEND, Date

    Exit Sub

DocumentReplacementError:
    MsgBox "Failed to create new record for document replacement."

End Sub



Private Sub MoveFileToNewLocation(ByVal filePathOrigin As String, ByVal projectId As String, ByVal docId As String, ByVal fileName As String)

    On Error GoTo FileMoveError
    shared_project_files.copyFile filePathOrigin, projectId, docId, fileName

    Exit Sub

FileMoveError:
    MsgBox "Failed to move document file to new location."

End Sub




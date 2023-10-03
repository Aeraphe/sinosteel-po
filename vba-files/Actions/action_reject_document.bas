Attribute VB_Name = "action_reject_document"

'namespace=vba-files\Actions


'/*
' This function handles the rejection of a document based on the provided rejectInfo object and other parameters.
'
' @param projectId The ID of the project that the document belongs to.
' @param docFullFilePath The file path of the document being rejected.
' @param rejectInfo An object containing information about the document rejection.
'
'  rejectInfo("motive")
'  rejectInfo("doc_id")
'  rejectInfo("user_id")
'  rejectInfo("review")
'
' @return None
'
'*/
Public Function RejectDocument(ByVal projectId As String, ByVal docFullFilePath As String, ByVal rejectInfo As Object) As Boolean



    ' If there is no docReviewId specified in the rejection info, remove it from the dictionary.
    If Not rejectInfo.Exists("docReviewId") Or rejectInfo("docReviewId") = "" Then
        rejectInfo.Remove ("docReviewId")
        rejectInfo.Add "category", Constants.REJECTED_BY_CDOC

        ' Call the cdocReject function with the remaining properties of the rejectInfo object, along with the project ID and document file path.
        RejectDocument = cdocReject(projectId, docFullFilePath, rejectInfo)
    End If


    ' If there is a docReviewId specified in the rejection info, call RejectContractorDocument with the specified docReviewId and motive.
    If rejectInfo.Exists("docReviewId") And Not IsEmpty(rejectInfo("docReviewId")) Then
        rejectInfo.Add "category", Constants.REJECTED_BY_CONTRACTOR
        RejectDocument = RejectContractorDocument(rejectInfo("docReviewId"), rejectInfo("INFO"))
    End If
End Function


'*/
'
'
'
'data("motive")
'data("doc_id")
'data("user_id")
'data("review")
'
'*/
Private Function cdocReject(ByVal project_id As String, ByVal fileFullPath As String, ByVal data As Object) As Boolean

    On Error GoTo ErrorHandler

    Dim reject_id As String
    Dim folder As String


    folder = Constants.FOLDER_CDOC_REJEITADOS & "_" & format(Now(), "_dd_MM_yyyy")
    Call copyFilesHandler(fileFullPath, project_id, folder)
    reject_id = db_documents.add_document_rejected(data)

    cdocReject = True

    Exit Function

ErrorHandler:

    cdocReject = False


End Function

'*/
'
'
'
'
'*/
Private Function RejectContractorDocument(ByVal reviewId As String, ByVal reason As String) As Boolean

    On Error GoTo ErrorHandler

    ' Declare and initialize variables
    Dim docId As String
    Dim docNumber As String
    Dim projectId As String
    Dim extension As String
    Dim revCode As String
    Dim grdId As String
    Dim folder As String
    Dim fileName As String
    Dim SourceFilePath As String
    Dim docPath As String
    Dim response As Boolean
    Dim rejectData As Object
    Set rejectData = CreateObject("Scripting.Dictionary")

    ' Get document details from database
    Dim documentRecordset As ADODB.Recordset
    Set documentRecordset = db_documents.get_document_by_review_id2(reviewId)

    docId = XdbFactory.getData(documentRecordset, "id")
    docNumber = XdbFactory.getData(documentRecordset, "doc_number")
    projectId = XdbFactory.getData(documentRecordset, "project_id")
    extension = LCase(XdbFactory.getData(documentRecordset, "doc_extension"))
    revCode = XdbFactory.getData(documentRecordset, "rev_code")
    grdId = XdbFactory.getData(documentRecordset, "grd_id")

    ' Update reject data
    rejectData("doc_id") = docId
    rejectData("motive") = reason
    rejectData("user_id") = auth.get_user_id
    rejectData("review") = revCode

    ' Move document to rejected folder
    docPath = helper_folder_maker.get_eng_doc_folder(projectId, docId, "SENT")
    If (extension <> "" And docPath <> "" And revCode <> "" And docNumber <> "") Then
        fileName = docNumber & "_Rev_" & revCode & "." & extension
        SourceFilePath = docPath & "\" & fileName
        folder = Constants.FOLDER_GED_CONTRATANTE_REJEITADOS & "_" & format(Now(), "_dd_MM_yyyy")
        Call copyFilesHandler(SourceFilePath, projectId, folder)
    End If

    ' Add rejected document record and update status
    Dim rejectId As String
    rejectId = db_documents.add_document_rejected(rejectData)
    Call db_documents.update_review_status(reviewId, "REJ", Date)

    'Some Rejections occurs when document has no GRD
    If (grdId <> "") Then
        Call db_grd.softDeleteDocument(grdId, reviewId)
    End If
    response = True


    RejectContractorDocument = response
    Exit Function

ErrorHandler:
    ' Log the error message to a file or database table instead of displaying it in a popup box
    RejectContractorDocument = False
 

End Function



Private Function copyFilesHandler(ByVal origin As String, ByVal project_id As String, ByVal sub_folder As String, Optional ByVal maxTries As Long = 5)
    On Error GoTo ErrorHandler

    ' Declare variables with appropriate data types
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim appFolders As Object: Set appFolders = helper_app.get_projec_folders(project_id)

    ' Construct destination path and file name
    Dim destFolderPath As String: destFolderPath = appFolders("ENG_DOC_REJECTED_FULL_PATH") & "\" & sub_folder
    Dim destFileName As String: destFileName = destFolderPath & "\" & fso.GetFileName(origin)


    ' Copy the file and check its checksum
    Dim tryCount As Long: tryCount = 1
    Do While Not file_helper.copyFilesWithCheckSum(origin, destFileName) And tryCount <= maxTries
        tryCount = tryCount + 1
    Loop

    Exit Function

ErrorHandler:
    ' Handle errors here
    MsgBox "An error occurred: " & Err.description
End Function


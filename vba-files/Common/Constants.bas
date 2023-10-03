Attribute VB_Name = "Constants"

'namespace=vba-files\Common

'DEBUG
Public Const DEBUG_APP As Boolean = True

'STATUS
Public Const EMITIR As String = "EMITIR"
Public Const PROGRAMADO As String = "PROGRAMADO"
Public Const NO_FLUXO As String = "NO FLUXO"
Public Const LIB_ENG As String = "LIB. ENG"
Public Const ENVIADO As String = "ENVIADO"
Public Const CONCLUIDO As String = "CONCLUIDO"
Public Const PEND As String = "PEND"
Public Const HOLD As String = "HOLD"
Public Const CANCELADO As String = "CANCELADO"
Public Const REJEITADO As String = "REJEITADO"
Public Const SUBISTITUIR As String = "SUBISTITUIR"

'REJECT DOCUMENT TYPE
Public Const REJECTED_BY_CDOC As String = "CDOC"
Public Const REJECTED_BY_CONTRACTOR As String = "CONTRACTOR"

'REVIEW STATUS
Public Const REVIEW_SATUS_SEND As String = "ENV"
Public Const REVIEW_SATUS_EXP As String = "EXP"
Public Const REVIEW_SATUS_POST As String = "POST"
Public Const REVIEW_SATUS_REJ As String = "REJ"

'FOLDERS
Public Const FOLDER_CDOC_REJEITADOS As String = "CDOC_REJEITADOS"
Public Const FOLDER_GED_CONTRATANTE_REJEITADOS As String = "GED_CONTRATANTE_REJEITADOS"
Public Const FOLDER_DEBUG As String = "APP_DEBUG"
Public Const COMMENTS_TEMP_FOLDER As String = "COMENTADOS"


'FILES
Public Const DEBUG_FILE_NAME As String = "cdoc_debug.log"
Public Const CHANGE_REQUEST_DOC_STATUS_FILE_NAME As String = "MUDANCA_DE_STATUS.txt"
Public Const DEGUB_COPY_FILE_NAME As String = "arquivos_copiados.txt"

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "StringBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'namespace=xvba_modules

Option Explicit
' Based off vba_stringbuilder:Class - http://blog.livedoor.jp/midorityo/archives/50744656.html
Private Type Members
    AppendsBufferSize As Long
    Builder As Object
    index As Long
End Type
Private AppendsBuffer() As String
Private m As Members

Private Sub Class_Initialize()
    Const DEFAULT_APPENDSBUFFER_SIZE = 10000
    Set m.Builder = CreateObject("System.Text.StringBuilder")
    setAppendsBufferSize DEFAULT_APPENDSBUFFER_SIZE
    m.index = -1
End Sub

Public Property Get Capacity() As Long
    Capacity = m.Builder.Capacity
End Property

Public Property Let Capacity(ByVal Length As Long)
    m.Builder.Capacity = Length
End Property

Public Property Get Length() As Long
    Length = m.Builder.Length
End Property

Public Property Let Length(ByVal lLength As Long)
    m.Builder.Length = lLength
End Property

Public Function MaxCapacity() As Long
    MaxCapacity = m.Builder.MaxCapacity
End Function

Public Function Append(ByRef Text As String, Optional repeatCount As Long = 0) As StringBuilder
    m.index = m.index + 1
    AppendsBuffer(m.index) = Text
    If repeatCount > 0 Then Me.Append Text, (repeatCount - 1)
    If m.index = m.AppendsBufferSize Then Compact
    Set Append = Me
End Function

Public Function AppendFormat(ByVal format As Variant, ParamArray Args() As Variant) As StringBuilder
    Dim newArgs() As Variant
    Dim j As Long
    Me.Compact

    If IsArray(Args(0)) Then
        newArgs = Args(0)
    Else
        ReDim newArgs(UBound(Args))
        For j = 0 To UBound(Args)
            newArgs(j) = Args(j)
        Next
    End If

    m.Builder.AppendFormat_4 format, newArgs
    Set AppendFormat = Me
End Function

Public Function InsertFormat(ByVal format As Variant, startIndex As Long, repeatCount As Long, ParamArray Args() As Variant) As StringBuilder
    Dim newArgs() As Variant, sb As New StringBuilder
    Dim j As Long
    If IsArray(Args(0)) Then
        newArgs = Args(0)
    Else
        ReDim newArgs(UBound(Args))
        For j = 0 To UBound(Args)
            newArgs(j) = Args(j)
        Next
    End If

    sb.AppendFormat format, newArgs
    Me.Insert startIndex, sb.ToString
    Set InsertFormat = Me
End Function

Public Function AppendLine(Optional ByRef Text As String, Optional repeatCount As Long = 0) As StringBuilder
    m.index = m.index + 1
    AppendsBuffer(m.index) = Text & vbNewLine
    If repeatCount > 0 Then Me.AppendLine Text, (repeatCount - 1)
    If m.index = m.AppendsBufferSize Then Compact
    Set AppendLine = Me
End Function

Public Sub Clear()
    Me.Length = 0
End Sub

Public Sub Compact()
    If m.index > -1 Then
        m.Builder.Append_3 Join(AppendsBuffer, vbNullString)
        m.index = -1
    End If
    ReDim AppendsBuffer(m.AppendsBufferSize)
End Sub

Public Function EnsureCapacity(ByVal Length As Long) As StringBuilder
    m.Builder.EnsureCapacity Length
    Set EnsureCapacity = Me
End Function

Public Function Insert(startIndex As Long, Text As String, Optional repeatCount As Long = 0) As StringBuilder
    Me.Compact
    m.Builder.Insert_2 startIndex, Text
    If repeatCount > 0 Then Me.Insert startIndex, Text, (repeatCount - 1)
    Set Insert = Me
End Function

Public Function Remove(ByVal startIndex As Long, ByVal Length As Long) As StringBuilder
    Me.Compact
    m.Builder.Remove startIndex, Length
    Set Remove = Me
End Function

Public Function Replace(oldValue As String, newValue As String, Optional startIndex As Long = -1, Optional repeatCount As Long = 0) As StringBuilder
    Me.Compact

    If startIndex > -1 And repeatCount > 0 Then
        m.Builder.Replace_2 oldValue, newValue, startIndex, repeatCount
    Else
        m.Builder.Replace oldValue, newValue
    End If
    Set Replace = Me
End Function

Public Sub setAppendsBufferSize(ByVal Length As Long)
    m.AppendsBufferSize = Length
    Me.Compact
End Sub

Function ToString(Optional ByVal startIndex As Long = -1, Optional ByVal Length As Long = 0) As String
    Me.Compact

    If startIndex > -1 And Length > 0 Then
        ToString = m.Builder.ToString_2(startIndex, Length)
    Else
        ToString = m.Builder.ToString
    End If
End Function

VERSION 5.00
Begin VB.Form Form_Main 
   Caption         =   "��ʷ�ϵĽ������ݴ������"
   ClientHeight    =   8745
   ClientLeft      =   690
   ClientTop       =   4275
   ClientWidth     =   10575
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8745
   ScaleWidth      =   10575
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton Button_Next 
      Caption         =   "��һ��"
      Height          =   375
      Left            =   8280
      TabIndex        =   7
      Top             =   5880
      Width           =   1095
   End
   Begin VB.CommandButton Button_previous 
      Caption         =   "��һ��"
      Height          =   375
      Left            =   7080
      TabIndex        =   6
      Top             =   5880
      Width           =   1095
   End
   Begin VB.TextBox Text_jsonPart 
      Height          =   5175
      Left            =   7080
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   5
      Top             =   600
      Width           =   3375
   End
   Begin VB.TextBox Text_jsonParsed 
      Height          =   5655
      Left            =   3600
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   600
      Width           =   3375
   End
   Begin VB.TextBox Text_Day 
      Height          =   375
      Left            =   960
      TabIndex        =   0
      Text            =   "2"
      Top             =   120
      Width           =   735
   End
   Begin VB.TextBox Text_Month 
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Text            =   "3"
      Top             =   120
      Width           =   735
   End
   Begin VB.CommandButton Button_Process 
      Caption         =   "����"
      Height          =   375
      Left            =   1800
      TabIndex        =   2
      Top             =   120
      Width           =   2055
   End
   Begin VB.TextBox Text_json 
      Height          =   5655
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   600
      Width           =   3375
   End
   Begin VB.Label Label_Index 
      Alignment       =   2  'Center
      Caption         =   "��0����ǰ��0��"
      Height          =   375
      Left            =   9480
      TabIndex        =   8
      Top             =   5880
      Width           =   975
   End
End
Attribute VB_Name = "Form_Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim jsonStrings() As String
Dim currentIndex

Private Sub Button_Next_Click()
    If currentIndex < UBound(jsonStrings) Then
        currentIndex = currentIndex + 1
        Text_jsonPart.Text = jsonStrings(currentIndex)
        Label_Index.Caption = "��" & UBound(jsonStrings) + 1 & "��" + vbCrLf + "��ǰ��" & currentIndex + 1 & "��"
    End If
End Sub

Private Sub Button_previous_Click()
    If currentIndex > 0 Then
        currentIndex = currentIndex - 1
        Text_jsonPart.Text = jsonStrings(currentIndex)
        Label_Index.Caption = "��" & UBound(jsonStrings) + 1 & "��" + vbCrLf + "��ǰ��" & currentIndex + 1 & "��"
    End If
End Sub

Private Sub Button_Process_Click()
    Dim jsonParsed As String

    '���ݻ�ȡ����

    Dim strURL As String '��ʷ�ϵĽ���API��ַ
    strURL = "http://www.todayonhistory.com/code/data/" & Text_Month.Text & "/" & Text_Day.Text & "/"
    
    Dim jsonString As String
    
    Dim xmlobject
    Set xmlobject = CreateObject("Microsoft.XMLHTTP")
    
    xmlobject.Open "GET", strURL, False
    xmlobject.Send
    
    DoEvents
    
    If xmlobject.ReadyState <> 4 Then
        Exit Sub
    End If
    
    Text_json.Text = xmlobject.Responsetext
    jsonParsed = standardizeJsonResult(Text_json.Text)
    Text_jsonParsed.Text = jsonParsed
    
    jsonStrings = splitJson(jsonParsed)
    
    currentIndex = 0
    Text_jsonPart.Text = jsonStrings(0)
    Label_Index.Caption = "��" & UBound(jsonStrings) + 1 & "��" + vbCrLf + "��ǰ��" & currentIndex + 1 & "��"
    
End Sub

Function splitJson(string_ As String)
    Dim result() As String
    result = Split(string_, "},{")
    '��json�����е����ݷ����и�������и����ҪһЩ����
    Dim i As Integer
    For i = 0 To UBound(result)
        Dim stringPart As String
        stringPart = result(i)
        If i = 0 Then
            stringPart = Mid(stringPart, 3, Len(stringPart))
        ElseIf i = UBound(result) Then
            stringPart = Mid(stringPart, 1, Len(stringPart) - 2)
        End If
        stringPart = "{" + stringPart + "}"
        result(i) = stringPart
    Next i
    splitJson = result
End Function


Function standardizeJsonResult(string_ As String)
    'ʹjson���ݱ�׼�������ڽ���
    Dim jsonString As String
    jsonString = string_
    
    jsonString = ChangeCommas(jsonString)
    
    jsonString = Replace(jsonString, "'", "")
    jsonString = Replace(jsonString, "var datalist = ", "")
    jsonString = Replace(jsonString, ";Toh.get(callback)", "")
    jsonString = Replace(jsonString, ":", """:""")
    jsonString = Replace(jsonString, ",", """,""")
    jsonString = Replace(jsonString, "http"":""", "http:")
    jsonString = Replace(jsonString, "}", """}")
    jsonString = Replace(jsonString, "{", "{""")
    jsonString = Replace(jsonString, "}"",""{", "},{")
    
    standardizeJsonResult = jsonString
    
End Function

Function ChangeCommas(string_ As String)
    Dim stringToChange As String
    stringToChange = string_

    '����������ڰ�value�е�Ӣ�ı��ת�������ı���Ա����
    Dim i As Integer
    Dim isValue As Boolean
    isValue = False
    For i = 1 To Len(stringToChange)
        Dim currentChar As String
        currentChar = StringAtIndex(stringToChange, i)
        If currentChar = "'" Then
            If isValue = True Then
                isValue = False
            Else
                isValue = True
            End If
        ElseIf currentChar = "," Then
            If isValue = True Then
                stringToChange = ReplaceAtIndex(stringToChange, i, "��")
            End If
        End If
    Next i
    ChangeCommas = stringToChange
End Function

'�������һЩ�ַ�����������vb�Դ��Ĵ�������û����Ҫ�ģ�����ֻ���Լ�д
'ͳһ��׼������ͳһ�ַ�����indexΪ1��ʼ

Function StringAtIndex(string_ As String, index As Integer)
    '����������ڻ�ȡ�ַ���ָ��index���ַ�
    If index < 0 Then
        StringAtIndex = ""
    End If
    
    StringAtIndex = Mid(string_, index, 1)
End Function

Function ReplaceAtIndex(string_ As String, index As Integer, stringToReplace As String)
    'ͨ������������԰��ַ�����ָ��index���ַ��滻����һЩ�ַ�
    If index < 0 Then
        ReplaceAtIndex = ""
    End If
    Dim partA As String
    Dim partB As String
    partA = Mid(string_, 1, index - 1)
    partB = Mid(string_, index + 1, Len(string_))
    ReplaceAtIndex = partA + stringToReplace + partB
End Function

Private Sub Form_Load()
    Label_Index.Caption = "��0��" + vbCrLf + "��ǰ��0��"
End Sub

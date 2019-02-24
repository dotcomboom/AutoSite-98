VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "AutoSite '98"
   ClientHeight    =   4545
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4155
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4545
   ScaleWidth      =   4155
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog objDialog 
      Left            =   120
      Top             =   3960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton btnRefresh 
      Caption         =   "Refresh"
      Enabled         =   0   'False
      Height          =   975
      Left            =   3240
      Picture         =   "Form1.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   1320
      Width           =   855
   End
   Begin VB.CommandButton btnOpen 
      Caption         =   "Open Web"
      Height          =   1095
      Left            =   3240
      Picture         =   "Form1.frx":0884
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton btnTemplate 
      Caption         =   "Edit Template"
      Enabled         =   0   'False
      Height          =   975
      Left            =   3240
      Picture         =   "Form1.frx":0CC6
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2400
      Width           =   855
   End
   Begin VB.CommandButton btnDeleteInclude 
      Caption         =   "Delete"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2400
      TabIndex        =   8
      Top             =   3360
      Width           =   615
   End
   Begin VB.FileListBox File2 
      Enabled         =   0   'False
      Height          =   3015
      Left            =   1800
      Pattern         =   "nothing"
      TabIndex        =   7
      Top             =   240
      Width           =   1215
   End
   Begin VB.CommandButton btnAddInclude 
      Caption         =   "Add"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1800
      TabIndex        =   6
      Top             =   3360
      Width           =   615
   End
   Begin VB.CommandButton btnDeletePage 
      Caption         =   "Delete"
      Enabled         =   0   'False
      Height          =   375
      Left            =   840
      TabIndex        =   5
      Top             =   3360
      Width           =   615
   End
   Begin VB.CommandButton btnMake 
      Caption         =   "Make"
      Enabled         =   0   'False
      Height          =   975
      Left            =   3240
      Picture         =   "Form1.frx":1108
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3480
      Width           =   855
   End
   Begin VB.Frame frmIncludes 
      Caption         =   "Includes"
      Enabled         =   0   'False
      Height          =   3855
      Left            =   1680
      TabIndex        =   1
      Top             =   0
      Width           =   1455
   End
   Begin VB.Frame frmIn 
      Caption         =   "In"
      Enabled         =   0   'False
      Height          =   3855
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   1455
      Begin VB.CommandButton btnNewPage 
         Caption         =   "New"
         Enabled         =   0   'False
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   3360
         Width           =   615
      End
      Begin VB.FileListBox File1 
         Enabled         =   0   'False
         Height          =   3015
         Left            =   120
         Pattern         =   "*.html"
         TabIndex        =   3
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   11
      Top             =   3960
      Width           =   3015
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const BIF_RETURNONLYFSDIRS = 1
Private Const BIF_DONTGOBELOWDOMAIN = 2
Private Const MAX_PATH = 260

Private Type BrowseInfo
    hwndOwner      As Long
    pIDLRoot       As Long
    pszDisplayName As Long
    lpszTitle      As Long
    ulFlags        As Long
    lpfnCallback   As Long
    lParam         As Long
    iImage         As Long
End Type

Private Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long

Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long

Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long

Public Function DirectoryExists(Dir As String) As Boolean
    Dim oDir As New Scripting.FileSystemObject
    DirectoryExists = oDir.FolderExists(Dir)
End Function

Public Function OpenDirectoryTV(Optional odtvTitle As String) As String
    Dim lpIDList As Long
    Dim sBuffer As String
    Dim szTitle As String
    Dim tBrowseInfo As BrowseInfo
    szTitle = odtvTitle
    With tBrowseInfo
           .hwndOwner = Form1.hWnd
           .lpszTitle = lstrcat(szTitle, "")
           .ulFlags = BIF_RETURNONLYFSDIRS
        End With
    lpIDList = SHBrowseForFolder(tBrowseInfo)
    If (lpIDList) Then
        sBuffer = Space(MAX_PATH)
        SHGetPathFromIDList lpIDList, sBuffer
        sBuffer = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
        OpenDirectoryTV = sBuffer
    End If
End Function


Private Sub Command1_Click()
    Dim fpath As New FileSystemObject
    
    If DirectoryExists(Label1.Caption & "\out") Then
    
        'MsgBox "Deleting output folder", vbOKOnly, "AutoSite"
        
        'fpath.DeleteFolder Label1.Caption & "\out", True
        On Error Resume Next
        Kill Label1.Caption & "\out\*.*"
        
    End If
    
    'If Not DirectoryExists(Label1.Caption & "\out") Then
    
        'MsgBox "Copying includes to output folder", vbOKOnly, "AutoSite"
        
        fpath.CopyFolder Label1.Caption & "\includes", Label1.Caption & "\out"
        
    'End If

    Dim sfilename As String
    sfilenames = Dir(Label1.Caption & "\in\*.*")
    Do While sfilenames > ""
        
        sfilename = Label1.Caption & "\in\" & sfilenames
        
        If Not sfilename Like "*\" Then
    
        Dim Attributes() As String
        Dim AttributeName As New Collection
        Dim AttributeValue As New Collection
        Dim Content As String
        Content = ""
        
        i = 0
        
        Dim MyLine As String
        Open sfilename For Input As #1
            Do While Not EOF(1)
                Line Input #1, MyLine
                If MyLine Like "<!--*-->" Then
                    Dim Att As String
                    Att = Replace(MyLine, "<!-- ", "")
                    Att = Replace(Att, " -->", "")
                    Att = Replace(Att, "<!--", "")
                    Att = Replace(Att, "-->", "")
                    Att = Replace(Att, "attrib ", "")
                    Att = Replace(Att, "attrib", "")
                    
                    Dim Atn As String
                    Atn = Split(Att, ":")(0)
                    Dim Atv As String
                    Atv = Replace(Att, Atn & ": ", "")
                    Atv = Replace(Att, Atn & ":", "")
                    
                    ReDim Attributes(0 To i) As String
                    Attributes(i) = Atn & ":" & Atv
                    i = i + 1
                Else
                    Content = Content & MyLine & vbNewLine
                End If
            Loop
        Close #1
        
        template = ""
        
        Open Label1.Caption & "\template.htm" For Input As #1
            Do While Not EOF(1)
                Line Input #1, MyLine
                template = template & MyLine & vbNewLine
            Loop
        Close #1
        
        For i = 0 To UBound(Attributes)
            Atn = Split(Attributes(i), ":")(0)
            Atv = Replace(Attributes(i), Atn & ": ", "", 1, 1)
            
            template = Replace(template, "[#content#]", Content)
            template = Replace(template, "[#" & Atn & "#]", Atv)
        Next
        
        Dim iFileNo As Integer
        iFileNo = FreeFile
        Dim file As String
        file = Replace(sfilename, "\in\", "\out\")
        
        Open file For Output As #iFileNo
                    
        Print #iFileNo, template
                    
        Close #iFileNo
        
        End If
        
        sfilenames = Dir()
    Loop
    
    Shell "explorer " & Label1.Caption & "\out", vbNormalFocus
End Sub

Private Sub Command2_Click()
    Dim folder As String

    folder = OpenDirectoryTV("Select Web Folder")
    If folder <> "" Then
        Label1.Caption = folder
        Label1.Alignment = 0
        Label1.FontItalic = False
        
        Dim checkspass As Boolean
        Dim inexists As Boolean
        Dim includesexists As Boolean
        checkspass = True
        inexists = True
        includesexists = True
        templateexists = True
        
        If Not DirectoryExists(folder + "\in") Then
            checkspass = False
            inexists = False
        End If
        If Not DirectoryExists(folder + "\includes") Then
            checkspass = False
            includesexists = False
        End If
        
        If Not (Dir(folder + "\template.htm") <> "") Then
            checkspass = False
            templateexists = False
        End If
        
        If Not checkspass Then
            If MsgBox("The required directories/files were not found here. AutoSite can create them for you.", vbOKCancel, "Missing folders") = vbOK Then
                If Not inexists Then
                    MkDir (folder + "\in")
                End If
                If Not includesexists Then
                    MkDir (folder + "\includes")
                End If
                If Not templateexists Then
                    Dim iFileNo As Integer
                    iFileNo = FreeFile
                    Dim file As String
                    file = folder + "\template.htm"
                    
                    Open file For Output As #iFileNo
                    
                    Print #iFileNo, "<html>"
                    Print #iFileNo, " <head>"
                    Print #iFileNo, "   <title>[#title#]</title>"
                    Print #iFileNo, " </head>"
                    Print #iFileNo, " <body>"
                    Print #iFileNo, "   <h1>[#title#]</h1>"
                    Print #iFileNo, "   [#content#]"
                    Print #iFileNo, " <body>"
                    Print #iFileNo, "</html>"
                    
                    Close #iFileNo
                End If
            Else
                Exit Sub
            End If
        End If
        
        File1.Path = folder + "\in"
        File1.Pattern = "*.*"
        File2.Path = folder + "\includes"
        File2.Pattern = "*.*"
        
        Command1.Enabled = True
        Command7.Enabled = True
        Command8.Enabled = True
        
        Frame1.Enabled = True
        Frame2.Enabled = True
        File1.Enabled = True
        File2.Enabled = True
        
        Command3.Enabled = True
        Command5.Enabled = True
    End If
End Sub

Private Sub Command3_Click()
FileName = InputBox("Filename (including .htm extension)", "New File", "")
If FileName = "" Then
Else
    Dim iFileNo As Integer
    iFileNo = FreeFile
    Dim file As String
    file = Label1.Caption + "\in\" + FileName
                    
    Open file For Output As #iFileNo
                    
    Print #iFileNo, "<!-- attrib title: " & StrConv(Split(FileName, ".")(0), vbProperCase) & " -->"
    Print #iFileNo, "<p>This is a page</p>"
                    
    Close #iFileNo
    
    File1.Refresh
End If

End Sub

Private Sub Command4_Click()
If MsgBox("Are you sure you want to delete " + File1.FileName + "?", vbYesNo, "Delete File") = vbYes Then
    Kill (Label1.Caption & "\in\" & File1.FileName)
    File1.Refresh
    Command4.Enabled = False
End If
End Sub

Private Sub Command5_Click()
CommonDialog1.Filter = "All files (*.*)|*.*"
CommonDialog1.DialogTitle = "Add include"
CommonDialog1.ShowOpen

If Not CommonDialog1.FileName = "" Then
    InFile = CommonDialog1.FileName
    outfile = Label1.Caption & "\includes\" & CommonDialog1.FileTitle
    FileCopy InFile, outfile
    File2.Refresh
End If
End Sub

Private Sub Command6_Click()
If MsgBox("Are you sure you want to delete " + File2.FileName + "?", vbYesNo, "Delete File") = vbYes Then
    Kill (Label1.Caption & "\includes\" & File2.FileName)
    File2.Refresh
    Command6.Enabled = False
End If
End Sub

Private Sub Command7_Click()
Shell "notepad " & Label1.Caption + "\template.htm", vbNormalFocus
End Sub

Private Sub Command8_Click()
File1.Refresh
File2.Refresh

Command6.Enabled = False
Command4.Enabled = False
End Sub

Private Sub File1_Click()
If Not File1.FileName = "" Then
    Command4.Enabled = True
End If
End Sub

Private Sub File1_DblClick()
If Not File1.FileName = "" Then
    Shell "notepad " & Label1.Caption & "\in\" & File1.FileName, vbNormalFocus
End If
End Sub

Private Sub File2_Click()
If Not File2.FileName = "" Then
    Command6.Enabled = True
End If
End Sub

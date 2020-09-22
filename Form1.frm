VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmRandom 
   Caption         =   "Random Letter Generator"
   ClientHeight    =   8820
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11130
   LinkTopic       =   "Form1"
   ScaleHeight     =   8820
   ScaleWidth      =   11130
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      Height          =   1455
      Left            =   120
      TabIndex        =   10
      Top             =   240
      Width           =   2775
      Begin VB.CommandButton cmdGenerate 
         Caption         =   "&Generate Letters"
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   1080
         Width           =   1455
      End
      Begin VB.TextBox txtRow 
         Height          =   285
         Left            =   1800
         TabIndex        =   12
         Text            =   "40000"
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox txtColumn 
         Height          =   285
         Left            =   1800
         TabIndex        =   11
         Text            =   "25"
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Letters per Row"
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Letters per Column"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   600
         Width           =   1335
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Max Length to search:"
      Height          =   615
      Left            =   9120
      TabIndex        =   8
      Top             =   2400
      Width           =   1935
      Begin VB.TextBox txtmax 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   600
         TabIndex        =   9
         Text            =   "5"
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Min Length to search:"
      Height          =   615
      Left            =   9120
      TabIndex        =   6
      Top             =   1680
      Width           =   1935
      Begin VB.TextBox txtmin 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   600
         TabIndex        =   7
         Text            =   "3"
         Top             =   240
         Width           =   615
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   5
      Top             =   8445
      Width           =   11130
      _ExtentX        =   19632
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   9552
            MinWidth        =   7832
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   9552
            MinWidth        =   7832
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Find &Words"
      Height          =   375
      Left            =   9120
      TabIndex        =   4
      Top             =   3120
      Width           =   1935
   End
   Begin VB.ListBox lstWords 
      Height          =   4740
      Left            =   9120
      TabIndex        =   3
      Top             =   3600
      Width           =   1935
   End
   Begin VB.TextBox txtFind 
      Height          =   285
      Left            =   6960
      TabIndex        =   2
      Top             =   1440
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Find String"
      Height          =   255
      Left            =   5760
      TabIndex        =   1
      Top             =   1440
      Width           =   1095
   End
   Begin RichTextLib.RichTextBox txtRandom 
      Height          =   6540
      Left            =   120
      TabIndex        =   0
      Top             =   1800
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   11536
      _Version        =   393217
      ScrollBars      =   2
      TextRTF         =   $"Form1.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmRandom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim lTime As Long

Private Sub cmdgenerate_Click()

    Dim Row As Long
    Dim col As Long
    Dim txt As String
    Dim new_letter As String
    
    Screen.MousePointer = vbHourglass
    StatusBar1.Panels(1).Text = ""
    txtRandom = ""
    DoEvents
    lTime = GetTickCount
    Open "C:\random.txt" For Binary As #1
    For Row = 1 To Trim(txtRow)
        For col = 1 To Trim(txtColumn)
            new_letter = Chr$(Int(Rnd * (90 - 65 + 1)) + 65)
            txt = txt + new_letter
        Next col
    txt = txt
    Put #1, , txt
    txt = ""
    Next Row
    Close #1
    lTime = GetTickCount - lTime
    StatusBar1.Panels(1).Text = "Total time to generate random letters: " & lTime / 1000 & " Seconds."
    DoEvents
    txtRandom.LoadFile ("C:\random.txt")
    Screen.MousePointer = vbDefault
    
End Sub

Private Sub Command2_Click()
    Dim ipos As Long
    txtRandom.Refresh
    txtRandom.SelLength = 0
    
    ipos = txtRandom.Find(Trim(txtFind), 1, Len(txtRandom))
    If ipos > 0 Then
        txtRandom.SelStart = ipos
        txtRandom.SelLength = Len(Trim(txtFind))
    End If
End Sub

Private Sub Command3_Click()

    Dim bReturn As Boolean
    Dim iStartPosition As Long
    Dim iEndPosition As Long
    Dim iWordSize As Integer
    Dim iMaxSize As Integer
    Dim iMinsize As Integer
    Dim sLastWord As String
    oWord.WordBasic.FileNew
    oWord.Visible = False
    iMinsize = Trim(txtmin)
    iMaxSize = Trim(txtmax)
    iStartPosition = 1
    iEndPosition = iMinsize
    Do Until iStartPosition = Trim(txtRow) * Trim(txtColumn)
        For iWordSize = iStartPosition + iMinsize To iStartPosition + iMaxSize
            txtRandom.SelStart = iStartPosition
            txtRandom.SelLength = iEndPosition - iStartPosition
            bReturn = oWord.CheckSpelling(Word:=txtRandom.SelText, IgnoreUppercase:=False)
            If bReturn = True Then
                txtRandom.SelBold = True
                txtRandom.SelColor = vbBlue
                lstWords.AddItem txtRandom.SelText
                lstWords.ListIndex = lstWords.ListCount - 1
                DoEvents
            End If
            StatusBar1.Panels(2).Text = "Letters Searched: " & iStartPosition & " \ " & Trim(txtRow) * Trim(txtColumn)
            iEndPosition = iEndPosition + 1
        Next iWordSize
    iStartPosition = iStartPosition + 1
    iEndPosition = iStartPosition + iMinsize
    Loop
End Sub

Private Sub Form_Load()
    Randomize
End Sub

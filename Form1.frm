VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   Caption         =   "Tooltip Class Test"
   ClientHeight    =   6270
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6930
   LinkTopic       =   "Form1"
   ScaleHeight     =   6270
   ScaleWidth      =   6930
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkBalloon 
      Caption         =   "Balloon Style (Cool!)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   4740
      TabIndex        =   21
      Top             =   3540
      Width           =   1965
   End
   Begin VB.PictureBox picTT 
      Height          =   525
      Left            =   4590
      ScaleHeight     =   465
      ScaleWidth      =   1185
      TabIndex        =   19
      Top             =   90
      Width           =   1245
   End
   Begin VB.CheckBox chkCenter 
      Caption         =   "Center"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   3180
      TabIndex        =   18
      Top             =   3540
      Width           =   1245
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   6030
      Top             =   5460
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdCreate 
      Caption         =   "Create"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2880
      TabIndex        =   17
      Top             =   5430
      Width           =   1245
   End
   Begin VB.Frame Frame1 
      Height          =   1425
      Left            =   150
      TabIndex        =   9
      Top             =   3750
      Width           =   6345
      Begin VB.OptionButton optIconType 
         Caption         =   "None"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   3
         Left            =   420
         TabIndex        =   15
         Top             =   1050
         Value           =   -1  'True
         Width           =   795
      End
      Begin VB.OptionButton optIconType 
         Caption         =   "Error"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   2
         Left            =   3150
         TabIndex        =   14
         Top             =   1050
         Width           =   945
      End
      Begin VB.OptionButton optIconType 
         Caption         =   "Warning"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   1
         Left            =   1560
         TabIndex        =   13
         Top             =   1050
         Width           =   1245
      End
      Begin VB.OptionButton optIconType 
         Caption         =   "Information"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   0
         Left            =   4500
         TabIndex        =   12
         Top             =   1050
         Width           =   1245
      End
      Begin VB.TextBox txtTitle 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1680
         TabIndex        =   11
         Text            =   "Enter title here"
         Top             =   210
         Width           =   3375
      End
      Begin VB.Label Label5 
         Caption         =   "Icon Type"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   150
         TabIndex        =   16
         Top             =   810
         Width           =   5085
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Title:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   240
         TabIndex        =   10
         Top             =   270
         Width           =   1245
      End
   End
   Begin VB.CheckBox chkShowTitle 
      Caption         =   "Show Title and Icon (optional)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   210
      TabIndex        =   8
      Top             =   3540
      Width           =   3315
   End
   Begin VB.CommandButton cmdChangeFgColor 
      Caption         =   "Change"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5610
      TabIndex        =   7
      Top             =   2790
      Width           =   1095
   End
   Begin VB.PictureBox picTxtColor 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2100
      ScaleHeight     =   255
      ScaleWidth      =   3345
      TabIndex        =   5
      Top             =   2760
      Width           =   3405
   End
   Begin VB.CommandButton cmdChangeBgColor 
      Caption         =   "Change"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5580
      TabIndex        =   4
      Top             =   2220
      Width           =   1095
   End
   Begin VB.PictureBox picBgColor 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2070
      ScaleHeight     =   255
      ScaleWidth      =   3345
      TabIndex        =   2
      Top             =   2190
      Width           =   3405
   End
   Begin VB.TextBox txtTipText 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   2100
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "Form1.frx":0000
      Top             =   810
      Width           =   3375
   End
   Begin VB.Label lblTT 
      Caption         =   "Hover mouse over picture box after pressing create"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   660
      TabIndex        =   20
      Top             =   180
      Width           =   3765
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Text Color:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   180
      TabIndex        =   6
      Top             =   2820
      Width           =   1755
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Background Color:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   150
      TabIndex        =   3
      Top             =   2250
      Width           =   1755
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Text:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   660
      TabIndex        =   1
      Top             =   870
      Width           =   1245
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim ttDemo As New Tooltip

Private Sub chkBalloon_Click()
    If CBool(chkBalloon.Value) Then
        ttDemo.Style = TTBalloon
    Else
        ttDemo.Style = TTStandard
    End If
    
End Sub

Private Sub chkCenter_Click()
    ttDemo.Centered = CBool(chkCenter.Value)
    
End Sub

Private Sub cmdChangeBgColor_Click()
On Local Error GoTo Err_cmdChangeBgColor_Click

    CommonDialog1.CancelError = True
    CommonDialog1.ShowColor
    picBgColor.BackColor = CommonDialog1.Color
    ttDemo.BackColor = CommonDialog1.Color
    

Exit Sub
Err_cmdChangeBgColor_Click:
Select Case Err.Number
    Case 32755
        ''do nothing
    Case Else
       Call MsgBox(Err.Number & ":" & Err.Description, vbCritical, "cmdChangeBgColor_Click")
End Select
End Sub


Private Sub cmdChangeFgColor_Click()
On Local Error GoTo Err_cmdChangeFgColor_Click

    CommonDialog1.CancelError = True
    CommonDialog1.ShowColor
    picTxtColor.BackColor = CommonDialog1.Color
    ttDemo.ForeColor = CommonDialog1.Color

Exit Sub
Err_cmdChangeFgColor_Click:
Select Case Err.Number
    Case 32755
        ''do nothing
    Case Else
       Call MsgBox(Err.Number & ":" & Err.Description, vbCritical, "cmdChangeFgColor_Click")
End Select
End Sub


Private Sub cmdCreate_Click()
    ttDemo.TipText = txtTipText
    Set ttDemo.ParentControl = picTT
    
    If CBool(chkShowTitle.Value) Then
        ttDemo.Title = txtTitle
    Else
        ttDemo.Icon = 0
    End If
    ttDemo.Create
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Set ttDemo = Nothing
    
End Sub

Private Sub Label6_Click()

End Sub

Private Sub optIconType_Click(Index As Integer)
    Select Case Index
        Case 0
            ttDemo.Icon = TTIconInfo
        Case 1
            ttDemo.Icon = TTIconWarning
        Case 2
            ttDemo.Icon = TTIconError
        Case 3
            ttDemo.Icon = TTNoIcon
    End Select
End Sub



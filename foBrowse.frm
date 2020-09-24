VERSION 5.00
Begin VB.Form foBrowse 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Browse"
   ClientHeight    =   7155
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5565
   Icon            =   "foBrowse.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7155
   ScaleWidth      =   5565
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton buSec 
      Caption         =   "&CHOOSE"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4320
      TabIndex        =   2
      Top             =   6600
      Width           =   1215
   End
   Begin VB.DirListBox sKlasor 
      Height          =   5940
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   5415
   End
   Begin VB.DriveListBox sSurucu 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5415
   End
End
Attribute VB_Name = "foBrowse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub buSec_Click()
 foAna.sNortonKlasoru = sKlasor.Path
 Unload Me
End Sub

Private Sub Form_Load()
 If foAna.sNortonKlasoru = "" Then Exit Sub
 sKlasor.Path = foAna.sNortonKlasoru
 sSurucu.Drive = Left(sKlasor.Path, 2)
End Sub

Private Sub sSurucu_Change()
 On Error Resume Next
 sKlasor.Path = sSurucu.Drive
 If Err.Number <> 0 Then MsgBox Err.Description, vbCritical, "Warning"
End Sub

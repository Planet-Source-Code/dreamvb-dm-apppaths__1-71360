VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmAdd 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2550
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5370
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2550
   ScaleWidth      =   5370
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOK 
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   350
      Left            =   3075
      TabIndex        =   4
      Top             =   1995
      Width           =   1012
   End
   Begin VB.CommandButton cmdcancel 
      Caption         =   "Cancel"
      Height          =   350
      Left            =   4185
      TabIndex        =   5
      Top             =   1995
      Width           =   1012
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   ". . . ."
      Height          =   345
      Left            =   4275
      TabIndex        =   3
      ToolTipText     =   "Open"
      Top             =   1305
      Width           =   585
   End
   Begin VB.TextBox txtExePath 
      Height          =   350
      Left            =   285
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   1305
      Width           =   3945
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   4740
      Top             =   795
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox txtName 
      Height          =   350
      Left            =   285
      TabIndex        =   1
      Top             =   570
      Width           =   3945
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   4590
      Picture         =   "frmAdd.frx":0000
      Top             =   75
      Width           =   480
   End
   Begin VB.Label lblPath 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Program Path"
      Height          =   195
      Left            =   285
      TabIndex        =   6
      Top             =   1065
      Width           =   960
   End
   Begin VB.Label lblName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Program Name:"
      Height          =   195
      Left            =   285
      TabIndex        =   0
      Top             =   285
      Width           =   1095
   End
End
Attribute VB_Name = "frmAdd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdcancel_Click()
    ButtonPress = vbCancel
    Unload frmAdd
End Sub

Private Sub cmdOK_Click()
    ExeName = txtName.Text
    ExePath = txtExePath.Text
    ButtonPress = vbOK
    Unload frmAdd
End Sub

Private Sub cmdOpen_Click()
On Error GoTo OpenErr:
    
    With CD1
        .CancelError = True
        .DialogTitle = "Open"
        .Filter = "Program Files(*.exe)|*.exe|"
        'Update Init Path, if file name is present
        If Len(txtExePath.Text) > 0 Then
            .InitDir = GetPathFormFile(txtExePath.Text)
        End If
        .ShowOpen
        'Update text box with exe filename
        txtExePath.Text = .FileName
        txtExePath.ToolTipText = .FileName
    End With
    
    Exit Sub
OpenErr:
    If (Err.Number = cdlCancel) Then
        Err.Clear
    End If
End Sub

Private Sub Form_Load()
    Set frmAdd.Icon = Nothing
    
    If (EditOp = 0) Then
        frmAdd.Caption = "Add"
        cmdOK.Caption = "OK"
    End If
    
    If (EditOp = 1) Then
        frmAdd.Caption = "Edit"
        cmdOK.Caption = "Update"
        txtName.Text = ExeName
        txtExePath.Text = ExePath
        txtExePath.ToolTipText = txtExePath.Text
    End If
    
End Sub

Private Sub txtExePath_Change()
    Call txtName_Change
End Sub

Private Sub txtName_Change()
    cmdOK.Enabled = Len(txtName.Text) > 0 And Len(txtExePath.Text) > 0
End Sub

Private Sub txtName_KeyPress(KeyAscii As Integer)
    If (KeyAscii = 46) Then
        KeyAscii = 0
    End If
    If (KeyAscii = 13) Then
        KeyAscii = 0
    End If
End Sub

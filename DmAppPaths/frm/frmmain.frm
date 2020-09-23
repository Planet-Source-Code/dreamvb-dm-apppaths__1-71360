VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmmain 
   Caption         =   "DM AppPaths"
   ClientHeight    =   5505
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   6810
   LinkTopic       =   "Form1"
   ScaleHeight     =   5505
   ScaleWidth      =   6810
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   285
      Left            =   0
      TabIndex        =   2
      Top             =   5220
      Width           =   6810
      _ExtentX        =   12012
      _ExtentY        =   503
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   6810
      _ExtentX        =   12012
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "M_ADD"
            Object.ToolTipText     =   "Add"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "M_EDIT"
            Object.ToolTipText     =   "Edit"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "M_DELETE"
            Object.ToolTipText     =   "Delete"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "M_RUN"
            Object.ToolTipText     =   "Run Program"
            ImageIndex      =   1
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   6390
      Top             =   990
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16711935
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":0000
            Key             =   "App"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":0352
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":06A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":09F6
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView LstV 
      Height          =   4680
      Left            =   0
      TabIndex        =   0
      Top             =   450
      Width           =   6195
      _ExtentX        =   10927
      _ExtentY        =   8255
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Program Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Program Path"
         Object.Width           =   2540
      EndProperty
   End
   Begin Project1.CReg CReg1 
      Left            =   6285
      Top             =   585
      _ExtentX        =   635
      _ExtentY        =   635
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuAdd 
         Caption         =   "Add Item"
      End
      Begin VB.Menu mnuEditA 
         Caption         =   "&Edit Item"
      End
      Begin VB.Menu mnuDel 
         Caption         =   "Delete Item"
      End
      Begin VB.Menu mnublank1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRun 
         Caption         =   "&Run Program"
      End
   End
   Begin VB.Menu mnuA 
      Caption         =   "#"
      Visible         =   0   'False
      Begin VB.Menu mnuEdit1 
         Caption         =   "&Edit Item"
      End
      Begin VB.Menu mnuDel1 
         Caption         =   "Delete Item"
      End
      Begin VB.Menu mnuRun1 
         Caption         =   "Run Program"
      End
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "&About"
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private sSubKey As String
Private mMouseButton As MouseButtonConstants

Private Sub AddItem()
Dim tmp As String
    tmp = sSubKey
    
    CReg1.SubKey = tmp & ExeName & ".exe"
    'Check if the item already exists.
    If (CReg1.KeyExsists) Then
        MsgBox "This item already exists.", vbInformation, "Add"
        Exit Sub
    Else
        If (CReg1.CreateKey <> 1) Then
            MsgBox "The item could not be added.", vbExclamation, "Error Adding Item"
        Else
            'Add the Exe Path
            CReg1.SetValue vbNullString, ExePath, REG_EXPAND_SZ
            CReg1.SetValue "Path", GetPathFormFile(ExePath), REG_EXPAND_SZ
        End If
    End If
    
    Call RefreshList
    'Clear up
    tmp = vbNullString
End Sub

Private Sub DeleteItem()
Dim sExeName As String
Dim sPath As String
    'Get Program's exe name.
    sExeName = LstV.SelectedItem.Key
    sPath = sSubKey & sExeName
    'Set the key to open
    CReg1.SubKey = sPath
    'Check if the item has deleted.
    If (CReg1.DeleteKey <> 1) Then
        MsgBox "The item could not be deleted.", vbExclamation, "Error Deleteing Item"
    Else
        'Refresh
        Call RefreshList
    End If
    
    'Clear up
    sExeName = vbNullString
    sPath = vbNullString
End Sub

Private Sub EditItem()
Dim sExeName As String
Dim sPath As String
Dim lIdx As Integer
Dim mOld As String

    'Get Program's exe name.
    lIdx = LstV.SelectedItem.Index
    sExeName = LstV.SelectedItem.Key
    mOld = sSubKey & sExeName
    
    'Set the key to open
    sPath = sSubKey & ExeName & ".exe"
    CReg1.SubKey = sPath
    'Check if the key exsist
    If CReg1.KeyExsists() Then
        'Add the new item data
         CReg1.SetValue vbNullString, ExePath, REG_EXPAND_SZ
         CReg1.SetValue "Path", GetPathFormFile(ExePath), REG_EXPAND_SZ
    Else
        If (CReg1.CreateKey <> 1) Then
            MsgBox "The item could not be updated.", vbExclamation, "Edit Item"
        Else
    
            'Add the new item data
            CReg1.SetValue vbNullString, ExePath, REG_EXPAND_SZ
            CReg1.SetValue "Path", GetPathFormFile(ExePath), REG_EXPAND_SZ
            CReg1.SubKey = mOld
            'Delete the old Item
            If (CReg1.DeleteKey <> 1) Then
                MsgBox "The item could not be updated.", vbExclamation, "Edit Item"
            End If
        End If
    End If
    
    'Refresh
    Call RefreshList
    Call SelectItem(lIdx)
    'Clear up
    sPath = vbNullString
    sExeName = vbNullString
End Sub

Private Sub RefreshList()
Dim Col As New Collection
Dim fExt As String
Dim Item
Dim sIcon As Integer
Dim lFile As String

    LstV.SortKey = 1
    'Set the subkey
    CReg1.SubKey = sSubKey
    'Get the subkeys
    Set Col = CReg1.GetSubKeys()
    'Fill Listview with app paths
    With LstV.ListItems
        .Clear
        For Each Item In Col
            'Get File Ext
            fExt = LCase(Right$(Item, 3))
            If (fExt = "exe") Then
                CReg1.SubKey = sSubKey & Item
                'Get programs' Exe Path
                lFile = CReg1.GetValue(vbNullString, REG_EXPAND_SZ)
                'Check if the program above is found.
                If FindFile(lFile) Then
                    'Font Icon
                    sIcon = 1
                Else
                    'Not Found icon
                    sIcon = 4
                End If
                'Add Program's file title removeing .exe
                .Add , Item, Left$(Item, Len(Item) - 4), , sIcon
                'Add program's exe path
                .Item(.Count).SubItems(1) = lFile
            End If
        Next Item
    End With
    'Enable/Disable Toolbar buttons.
    Toolbar1.Buttons(2).Enabled = LstV.ListItems.Count
    Toolbar1.Buttons(3).Enabled = LstV.ListItems.Count
    Toolbar1.Buttons(5).Enabled = LstV.ListItems.Count
    'Enable/Disable Menu Items
    mnuEditA.Enabled = LstV.ListItems.Count
    mnuDel.Enabled = LstV.ListItems.Count
    mnuRun.Enabled = LstV.ListItems.Count
    
    'Autosize Columns headers
    Call lvSizeColumns(LstV)
    Call SelectItem(1)
    'Clear up
    Set Col = Nothing
    Item = vbNullString
End Sub

Private Sub SelectItem(ByVal Index As Integer)
Dim lItem As ListItem
    'Select an Item
    Set lItem = LstV.ListItems(Index)
    Call LstV_ItemClick(lItem)
    lItem.Selected = True
End Sub

Private Sub Form_Load()
    'Set the main Key to open.
    CReg1.Key = HKEY_LOCAL_MACHINE
    'Setup the main subkey to open
    sSubKey = "SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\"
    'Fill the Listview
    Call RefreshList
End Sub

Private Sub Form_Resize()
On Error Resume Next
    LstV.Width = frmmain.ScaleWidth
    LstV.Height = (frmmain.ScaleHeight - StatusBar1.Height - LstV.Top)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmAdd = Nothing
    Set frmmain = Nothing
End Sub

Private Sub LstV_DblClick()
    If (LstV.ListItems.Count) Then
        'Store List Item data
        ExeName = LstV.SelectedItem.Text
        ExePath = LstV.SelectedItem.SubItems(1)
        'Edit Item
        EditOp = 1
        frmAdd.Show vbModal, frmmain
    End If
End Sub

Private Sub LstV_ItemClick(ByVal Item As MSComctlLib.ListItem)
    ExeName = Item.Text
    ExePath = Item.SubItems(1)
    'Check what mouse button was pressed.
    If (mMouseButton = vbRightButton) Then
        'Show popup menu.
        PopupMenu mnuA
    End If
End Sub

Private Sub LstV_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    mMouseButton = Button
End Sub

Private Sub mnuAbout_Click()
    MsgBox frmmain.Caption & " V1.0" & vbCrLf & vbTab & "By DreamVB" _
    & vbCrLf & vbTab & vbTab & "Please vote if you like this code.", vbInformation, "About"
End Sub

Private Sub mnuAdd_Click()
    EditOp = 0
    'Show Add Form
    frmAdd.Show vbModal, frmmain
    If (ButtonPress = vbOK) Then
        'Add New Item
        Call AddItem
    End If
End Sub

Private Sub mnuDel_Click()
    If MsgBox("Are you sure you want to delete this item.", vbYesNo Or vbQuestion, "Delete Item") = vbYes Then
        'Delete Item
        Call DeleteItem
    End If
End Sub

Private Sub mnuDel1_Click()
    Call mnuDel_Click
End Sub

Private Sub mnuEdit1_Click()
    Call mnuEditA_Click
End Sub

Private Sub mnuEditA_Click()
    EditOp = 1
    'Show Edit Form
    frmAdd.Show vbModal, frmmain
    If (ButtonPress = vbOK) Then
        'Edit Item
        Call EditItem
    End If
End Sub

Private Sub mnuExit_Click()
    Unload frmmain
End Sub

Private Sub mnuRun_Click()
    If RunApp(frmmain.hwnd, "open", LstV.SelectedItem.SubItems(1)) = 2 Then
        MsgBox "The selected program could not be opened.", vbCritical, "Run Program"
    End If
End Sub

Private Sub mnuRun1_Click()
    Call mnuRun_Click
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "M_ADD"
            'Add Item
            Call mnuAdd_Click
        Case "M_EDIT"
            'Edit Item
            Call mnuEditA_Click
        Case "M_DELETE"
            'Delete Item
            Call mnuDel_Click
        Case "M_RUN"
            'Run Selected Program
            Call mnuRun_Click
    End Select
    
    ButtonPress = vbCancel
End Sub

VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmMenuEdit 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Dynamic Menu vx.x.x"
   ClientHeight    =   5040
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   9960
   Icon            =   "frmMenuEdit.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5040
   ScaleWidth      =   9960
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.Frame fraGlobal 
      Caption         =   "Global Options"
      Height          =   4455
      Left            =   6720
      TabIndex        =   30
      Top             =   0
      Width           =   3135
      Begin VB.CheckBox chkAddEdit 
         Caption         =   "Add ""Edit Menu"" to menu?"
         Height          =   195
         Left            =   240
         TabIndex        =   22
         Top             =   4080
         Value           =   1  'Checked
         Width           =   2415
      End
      Begin VB.CheckBox chkAddSystray 
         Caption         =   "Add ""Add to SysTray"" to menu?"
         Height          =   195
         Left            =   240
         TabIndex        =   21
         Top             =   3720
         Value           =   1  'Checked
         Width           =   2775
      End
      Begin VB.ComboBox cmbClickType 
         Height          =   315
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   600
         Width           =   2655
      End
      Begin VB.CheckBox chkAddDefaultstoBottom 
         Caption         =   "Add default items to bottom of menu?"
         Height          =   375
         Left            =   240
         TabIndex        =   20
         Top             =   3240
         Width           =   2775
      End
      Begin VB.TextBox txtSysTrayIcon 
         Height          =   285
         Left            =   240
         TabIndex        =   16
         Top             =   1320
         Width           =   2175
      End
      Begin VB.CommandButton cmdLoadIcon 
         Caption         =   "..."
         Height          =   255
         Left            =   2520
         TabIndex        =   17
         Top             =   1335
         Width           =   375
      End
      Begin VB.TextBox txtTooltip 
         Height          =   285
         Left            =   240
         TabIndex        =   18
         Top             =   2040
         Width           =   2655
      End
      Begin VB.TextBox txtColumns 
         Height          =   285
         Left            =   240
         MaxLength       =   3
         TabIndex        =   19
         ToolTipText     =   "Overrides column breaks if set."
         Top             =   2760
         Width           =   2655
      End
      Begin VB.Label Label1 
         Caption         =   "System Tray Click Type:"
         Height          =   195
         Index           =   4
         Left            =   240
         TabIndex        =   34
         Top             =   360
         Width           =   1710
      End
      Begin VB.Label Label1 
         Caption         =   "System Tray Icon:"
         Height          =   195
         Index           =   5
         Left            =   240
         TabIndex        =   33
         Top             =   1080
         Width           =   1275
      End
      Begin VB.Label Label1 
         Caption         =   "System Tray Tooltip:"
         Height          =   195
         Index           =   6
         Left            =   240
         TabIndex        =   32
         Top             =   1800
         Width           =   1800
      End
      Begin VB.Label Label1 
         Caption         =   "Number of Columns:"
         Height          =   195
         Index           =   9
         Left            =   240
         TabIndex        =   31
         Top             =   2520
         Width           =   1425
      End
   End
   Begin VB.Frame fraItem 
      Caption         =   "Item Options"
      Height          =   4455
      Left            =   3480
      TabIndex        =   24
      Top             =   0
      Width           =   3135
      Begin VB.CheckBox chkColumnBreak 
         Caption         =   "Column Break?"
         Height          =   195
         Left            =   240
         TabIndex        =   14
         Top             =   3960
         Width           =   2415
      End
      Begin VB.CommandButton cmdLoadBitmap 
         Caption         =   "..."
         Height          =   255
         Left            =   2520
         TabIndex        =   13
         Top             =   3480
         Width           =   375
      End
      Begin VB.TextBox txtMenuIcon 
         Height          =   285
         Left            =   240
         TabIndex        =   12
         Top             =   3480
         Width           =   2175
      End
      Begin VB.TextBox txtCaption 
         Height          =   285
         Left            =   240
         TabIndex        =   8
         Top             =   600
         Width           =   2655
      End
      Begin VB.TextBox txtCmd 
         Height          =   285
         Left            =   240
         TabIndex        =   9
         Top             =   1320
         Width           =   2655
      End
      Begin VB.ComboBox cmbType 
         Height          =   315
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   2760
         Width           =   2655
      End
      Begin VB.TextBox txtArgs 
         Height          =   285
         Left            =   240
         TabIndex        =   10
         Top             =   2040
         Width           =   2655
      End
      Begin VB.Label Label1 
         Caption         =   "Menu Bitmap:"
         Height          =   195
         Index           =   7
         Left            =   240
         TabIndex        =   29
         Top             =   3240
         Width           =   1110
      End
      Begin VB.Label Label1 
         Caption         =   "Caption"
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   28
         Top             =   360
         Width           =   585
      End
      Begin VB.Label Label1 
         Caption         =   "Command"
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   27
         Top             =   1080
         Width           =   705
      End
      Begin VB.Label Label1 
         Caption         =   "Arguments"
         Height          =   195
         Index           =   2
         Left            =   240
         TabIndex        =   26
         Top             =   1800
         Width           =   750
      End
      Begin VB.Label Label1 
         Caption         =   "Command Type"
         Height          =   195
         Index           =   3
         Left            =   240
         TabIndex        =   25
         Top             =   2520
         Width           =   1110
      End
   End
   Begin VB.CommandButton cmdTest 
      Caption         =   "Test"
      Height          =   375
      Left            =   7860
      TabIndex        =   23
      Top             =   4560
      Width           =   855
   End
   Begin VB.ListBox lstMenu 
      Height          =   3765
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   3255
   End
   Begin VB.CommandButton cmdRight 
      Caption         =   "Right"
      Height          =   375
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4560
      Width           =   735
   End
   Begin VB.CommandButton cmdLeft 
      Caption         =   "Left"
      Height          =   375
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4560
      Width           =   735
   End
   Begin VB.CommandButton cmdDown 
      Caption         =   "Down"
      Height          =   375
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4560
      Width           =   735
   End
   Begin VB.CommandButton cmdUp 
      Caption         =   "Up"
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4560
      Width           =   735
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete"
      Height          =   375
      Left            =   2400
      TabIndex        =   2
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add"
      Height          =   375
      Left            =   1200
      TabIndex        =   1
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton cmdInsert 
      Caption         =   "&Insert"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
   Begin MSComDlg.CommonDialog CommonDialog 
      Left            =   6960
      Top             =   4800
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileNew 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Open..."
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "&Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFileSaveAs 
         Caption         =   "Save &As..."
      End
      Begin VB.Menu mnuFileSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "frmMenuEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'    Copyright (C) 2002  Roger Light <roger@atchoo.org>
'
'    This program is free software; you can redistribute it and/or modify
'    it under the terms of the GNU General Public License as published by
'    the Free Software Foundation; either version 2 of the License.
'
'    This program is distributed in the hope that it will be useful,
'    but WITHOUT ANY WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'    GNU General Public License for more details.
'
'    You should have received a copy of the GNU General Public License
'    along with this program; if not, write to the Free Software
'    Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA

Option Explicit

Private Type MenuData
    Indent As Long
    Caption As String
    Cmd As String
    Args As String
    Type As String
    Output As Boolean
    Parent As Integer
    hMenu As Long
    hBitmap As Long
    Bitmap As String
    ColumnBreak As Boolean
End Type

Private Menu() As MenuData                  ' Array holding information on the current menu loaded
Private Selected As Integer                 ' The currently selected item. Needed to help with adding/inserting items
Private bDirty As Boolean                   ' Have we modified the menu (and hence does it need saving)?
Private WithEvents Parser As GGZMLParser    ' SAX XML parser to read in the dym file
Attribute Parser.VB_VarHelpID = -1
Private CurrentFile As String               ' Current file name to help with saving
Private CurrentMenuName As String           ' The name of the current file without a path
Private hMenu As Long                       ' Root handle to the created menu
Private hMenuList() As Long                 ' Handles to all menus created
Private UserMenuOffset As Integer ' essentially defines the number of menu items added by default by us not the user
                                ' This is Close Menu, Edit Menu etc.
Private ColCount As Integer                 ' Number of columns to use in creating the menu
Private RootItemCount As Integer            ' Number of items in the root menu
Private RootItemIndex As Integer            ' The current Root item (used to help in splitting the menu into even columns)

Private Declare Function CreatePopupMenu Lib "user32" () As Long
Private Declare Function TrackPopupMenu Lib "user32" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal X As Long, ByVal Y As Long, ByVal nReserved As Long, ByVal hwnd As Long, ByVal lprc As Any) As Long
Private Declare Function DestroyMenu Lib "user32" (ByVal hMenu As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function InsertMenuItem Lib "user32.dll" Alias "InsertMenuItemA" (ByVal hMenu As Long, ByVal uItem As Long, ByVal fByPosition As Long, lpmii As MENUITEMINFO) As Long
Private Declare Function GetMenuItemCount Lib "user32.dll" (ByVal hMenu As Long) As Long

Private Type MENUITEMINFO
    cbSize As Long
    fMask As Long
    fType As Long
    fState As Long
    wID As Long
    hSubMenu As Long
    hbmpChecked As Long
    hbmpUnchecked As Long
    dwItemData As Long
    dwTypeData As String
    cch As Long
End Type

Private Const MIIM_STATE = &H1
Private Const MIIM_ID = &H2
Private Const MIIM_TYPE = &H10
Private Const MIIM_SUBMENU = &H4
Private Const MIIM_CHECKMARKS = &H8
Private Const MFT_SEPARATOR = &H800
Private Const MFT_STRING = &H0
Private Const MFT_MENUBARBREAK = &H20
Private Const MFS_ENABLED = &H0
Private Const MFS_CHECKED = &H8
Private Const MFS_GRAYED = &H3&
Private Const SW_SHOWNORMAL = 1
Private Const TPM_LEFTALIGN = &H0&
Private Const TPM_RETURNCMD = &H100&

Private Type POINTAPI
    X As Long
    Y As Long
End Type

Private Sub chkAddDefaultstoBottom_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        bDirty = True
    End If
End Sub

Private Sub chkAddEdit_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        bDirty = True
    End If
End Sub

Private Sub chkAddSystray_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        bDirty = True
    End If
End Sub

Private Sub chkColumnBreak_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If lstMenu.ListIndex >= 0 Then
        Menu(lstMenu.ListIndex).ColumnBreak = CBool(chkColumnBreak.Value)
        bDirty = True
    End If
End Sub

Private Sub cmbClickType_KeyPress(KeyAscii As Integer)
    bDirty = True
End Sub

Private Sub cmbType_Click()
    If lstMenu.ListIndex >= 0 Then
        Menu(lstMenu.ListIndex).Type = LCase$(cmbType.Text)
    End If
End Sub

Private Sub cmbType_KeyUp(KeyCode As Integer, Shift As Integer)
    If lstMenu.ListIndex >= 0 Then
        Menu(lstMenu.ListIndex).Type = LCase$(cmbType.Text)
        bDirty = True
    End If
End Sub

Private Sub AddNewMenuItem(Index As Integer)
    Dim i As Integer
    If lstMenu.ListIndex >= 0 Then
        InsertItem Index
        Menu(Index).Caption = "New Item" & CStr(UBound(Menu))
        Menu(Index).Indent = Menu(lstMenu.ListIndex).Indent
        Menu(Index).Cmd = ""
        Menu(Index).Args = ""
        Selected = Index
    Else
        ' Add item to the start of the list - there are no other items
        InsertItem 0
        Menu(UBound(Menu) - 1).Caption = "New Item" & CStr(UBound(Menu))
        Menu(UBound(Menu) - 1).Indent = 0
        Menu(UBound(Menu) - 1).Parent = -1
        Menu(UBound(Menu) - 1).Cmd = ""
        Menu(UBound(Menu) - 1).Args = ""
        Selected = 0
    End If
    bDirty = True
    DrawList
End Sub

Private Sub cmdInsert_Click()
    ' Add new item before selected item
    AddNewMenuItem lstMenu.ListIndex
End Sub

Private Sub cmdAdd_Click()
    ' Add new item after selected item
    AddNewMenuItem lstMenu.ListIndex + 1
End Sub

Private Sub cmdDelete_Click()
    If lstMenu.ListIndex >= 0 Then
        DeleteItem lstMenu.ListIndex
        If Selected = lstMenu.ListCount - 1 Then
            Selected = Selected - 1
        End If
        If Selected >= 0 Then
            If Selected = 0 And Menu(Selected).Indent > 0 Then
                Do While Menu(Selected).Indent > 0
                    MoveLeft Selected
                Loop
            ElseIf Menu(Selected).Indent + 2 = Menu(Selected + 1).Indent Then
                MoveLeft Selected + 1
            ElseIf Selected > 0 Then
                If Menu(Selected).Indent - 2 = Menu(Selected - 1).Indent Then
                    MoveLeft Selected
                End If
            End If
        End If
        DrawList
        bDirty = True
    End If
End Sub

Private Sub cmdDown_Click()
    ' Move current item down in the list.
    If lstMenu.ListIndex < lstMenu.ListCount - 1 And lstMenu.ListIndex >= 0 Then
        ' Change the indents as "appropriate"
        If Menu(lstMenu.ListIndex).Indent > Menu(lstMenu.ListIndex + 1).Indent Then
            Menu(lstMenu.ListIndex).Indent = Menu(lstMenu.ListIndex + 1).Indent
        End If
        SwapItems lstMenu.ListIndex, lstMenu.ListIndex + 1
        Selected = Selected + 1
        DrawList
        bDirty = True
    End If
End Sub

Private Sub cmdLeft_Click()
    MoveLeft lstMenu.ListIndex
End Sub

Private Sub cmdLoadBitmap_Click()
    If lstMenu.ListIndex >= 0 Then
        CommonDialog.Filter = "Bitmap files (*.bmp)|*.bmp"
        CommonDialog.FilterIndex = 0
        CommonDialog.DialogTitle = "Open Bitmap"
        CommonDialog.ShowOpen
        If Len(CommonDialog.FileName) > 0 Then
            txtMenuIcon.Text = CommonDialog.FileName
            Menu(Selected).Bitmap = txtMenuIcon.Text
            bDirty = True
        End If
    End If
End Sub

Private Sub cmdLoadIcon_Click()
    CommonDialog.Filter = "Icon files (*.ico)|*.ico"
    CommonDialog.FilterIndex = 0
    CommonDialog.DialogTitle = "Open Icon"
    CommonDialog.ShowOpen
    If Len(CommonDialog.FileName) > 0 Then
        txtSysTrayIcon.Text = CommonDialog.FileName
        bDirty = True
    End If
End Sub

Private Sub cmdRight_Click()
    MoveRight lstMenu.ListIndex
End Sub

Private Sub cmdTest_Click()
    CreateMenu
End Sub

Private Sub cmdUp_Click()
    If lstMenu.ListIndex > 0 Then
        If Menu(lstMenu.ListIndex).Indent > Menu(lstMenu.ListIndex - 1).Indent Then
            Menu(lstMenu.ListIndex).Indent = Menu(lstMenu.ListIndex).Indent - 1
        End If
        SwapItems lstMenu.ListIndex, lstMenu.ListIndex - 1
        Selected = Selected - 1
        DrawList
        bDirty = True
    End If
End Sub

Private Sub DrawList()
    Dim i As Integer, j As Integer, c As Integer
    Dim s As String
    
    For i = 0 To UBound(Menu) - 1
        s = ""
        For j = 1 To Menu(i).Indent
            s = s & "... "
        Next
        If i < lstMenu.ListCount Then
            lstMenu.List(i) = s & Menu(i).Caption
        Else
            lstMenu.AddItem s & Menu(i).Caption
        End If
    Next
    Do While lstMenu.ListCount > UBound(Menu) And UBound(Menu) <> 0
        If lstMenu.ListCount > 0 Then
            lstMenu.RemoveItem lstMenu.ListCount - 1
        End If
    Loop
    If UBound(Menu) = 0 Then
        lstMenu.Clear
    End If
    If Selected < lstMenu.ListCount And Selected >= 0 Then
        lstMenu.ListIndex = Selected
        txtCaption.Text = Menu(Selected).Caption
        txtCmd.Text = Menu(Selected).Cmd
        txtArgs.Text = Menu(Selected).Args
        txtMenuIcon.Text = Menu(Selected).Bitmap
        Select Case LCase$(Menu(Selected).Type)
            Case "shell"
                cmbType.ListIndex = 0
            Case "shellexec"
                cmbType.ListIndex = 1
        End Select
        If Menu(Selected).ColumnBreak Then
            chkColumnBreak.Value = 1
        Else
            chkColumnBreak.Value = 0
        End If
    Else
        txtCaption.Text = ""
        txtCmd.Text = ""
        txtArgs.Text = ""
        txtMenuIcon.Text = ""
        cmbType.ListIndex = 0
        chkColumnBreak.Value = 0
    End If
End Sub

Private Sub MoveRight(Index As Integer)
    If Index > 0 Then
        If Menu(Index).Indent <= Menu(Index - 1).Indent Then
            Menu(Index).Indent = Menu(Index).Indent + 1
            bDirty = True
            If Index < lstMenu.ListCount - 1 Then
                If Menu(Index).Indent = Menu(Index + 1).Indent Then
                    MoveRight Index + 1
                End If
            End If
        End If
        DrawList
    End If
End Sub

Private Sub MoveLeft(Index As Integer)
    If Index >= 0 Then
        If Menu(Index).Indent > 0 Then
            Menu(Index).Indent = Menu(Index).Indent - 1
            bDirty = True
            If Index < lstMenu.ListCount - 1 Then
                If Menu(Index).Indent = Menu(Index + 1).Indent - 2 Then
                    MoveLeft Index + 1
                End If
            End If
        End If
        DrawList
    End If
End Sub

Private Sub Form_Load()
    Me.Caption = "Dynamic Menu v" & CStr(App.Major) & "." & CStr(App.Minor) & "." & CStr(App.Revision)
    OldWindowProc = 0
    IconInTray = False
    Dim s As String
    ReDim Menu(0)
    cmbType.AddItem "Shell"
    cmbType.AddItem "ShellExec"
    
    cmbClickType.AddItem "Any Click"
    cmbClickType.AddItem "Left Click"
    ' Itemdata() is used to make detecting the correct click easier. See modSystray.WindowProc
    cmbClickType.ItemData(1) = WM_LBUTTONUP
    cmbClickType.AddItem "Middle Click"
    cmbClickType.ItemData(2) = WM_MBUTTONUP
    cmbClickType.AddItem "Right Click"
    cmbClickType.ItemData(3) = WM_RBUTTONUP
    cmbClickType.AddItem "Left Double Click"
    cmbClickType.ItemData(4) = WM_LBUTTONDBLCLK
    cmbClickType.AddItem "Middle Double Click"
    cmbClickType.ItemData(5) = WM_MBUTTONDBLCLK
    cmbClickType.AddItem "Right Double Click"
    cmbClickType.ItemData(6) = WM_RBUTTONDBLCLK
    cmbClickType.ListIndex = 0
    CommonDialog.Flags = cdlOFNHideReadOnly Or cdlOFNOverwritePrompt
    If Len(Command$) = 0 Then
        Me.Show
    ElseIf InStr(1, Command$, "/edit ", vbTextCompare) Then
        s = Replace$(Command$, "/edit ", "", 1, -1, vbTextCompare)
        If FileExists(s) Then
            ParseFile s
            CurrentFile = s
            Do While InStr(s, "\")
                s = Right$(s, Len(s) - InStr(s, "\"))
            Loop
            If InStr(1, s, ".dym", vbTextCompare) Then
                s = Left$(s, InStr(1, s, ".dym", vbTextCompare) - 1)
            End If
            CurrentMenuName = s
        End If
        Me.Show
        DrawList
    ElseIf InStr(1, Command$, "/systray ", vbTextCompare) Then
        s = Replace$(Command$, "/systray ", "", 1, -1, vbTextCompare)
        If FileExists(s) Then
            ParseFile s
            CurrentFile = s
            Do While InStr(s, "\")
                s = Right$(s, Len(s) - InStr(s, "\"))
            Loop
            If InStr(1, s, ".dym", vbTextCompare) Then
                s = Left$(s, InStr(1, s, ".dym", vbTextCompare) - 1)
            End If
            CurrentMenuName = s
            AddTrayIcon CurrentMenuName
        Else
            Me.Show
        End If
    ElseIf FileExists(Command$) = False Then
        Me.Show
    Else
        ParseFile Command$
        s = Command$
        Do While InStr(s, "\")
            s = Right$(s, Len(s) - InStr(s, "\"))
        Loop
        If InStr(1, s, ".dym", vbTextCompare) Then
            s = Left$(s, InStr(1, s, ".dym", vbTextCompare) - 1)
        End If
        CurrentMenuName = s
        CreateMenu
        If OldWindowProc = 0 Then
            Unload Me
        End If
    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If bDirty = True Then
        Dim result As VbMsgBoxResult
        result = MsgBox("Save changes to current menu?", vbYesNoCancel Or vbQuestion)
        If result = vbYes Then
            DoSave
        ElseIf result = vbCancel Then
            Cancel = 1
        End If
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    RemoveTrayIcon
    Dim i As Integer
    For i = 0 To UBound(Menu) - 1
        If Menu(i).hBitmap Then
            DeleteObject Menu(i).hBitmap
        End If
    Next
    If hMenu Then
        DestroyMenu hMenu
        hMenu = 0
    End If
End Sub

Private Sub lstMenu_Click()
    Selected = lstMenu.ListIndex
    If Selected >= 0 Then
        txtCaption.Text = Menu(Selected).Caption
        txtCmd.Text = Menu(Selected).Cmd
        txtArgs.Text = Menu(Selected).Args
        txtMenuIcon.Text = Menu(Selected).Bitmap
        If Menu(Selected).ColumnBreak Then
            chkColumnBreak.Value = 1
        Else
            chkColumnBreak.Value = 0
        End If
        Select Case LCase$(Menu(Selected).Type)
            Case "shell"
                cmbType.ListIndex = 0
            Case "shellexec"
                cmbType.ListIndex = 1
            Case Else
                cmbType.ListIndex = 0
        End Select
    End If
End Sub

Private Sub mnuFileExit_Click()
    Unload Me
End Sub

Private Sub mnuFileNew_Click()
    lstMenu.Clear
    txtCmd.Text = ""
    txtCaption.Text = ""
    txtArgs.Text = ""
    ReDim Menu(0)
    CurrentFile = ""
End Sub

Private Sub mnuFileOpen_Click()
    CommonDialog.Filter = "Dyn Menus (*.dym)|*.dym|All files (*.*)|*.*"
    CommonDialog.FilterIndex = 0
    CommonDialog.DialogTitle = "Open Dynamic Menu"
    CommonDialog.ShowOpen
    If Len(CommonDialog.FileName) > 0 Then
        If FileExists(CommonDialog.FileName) Then
            ReDim Menu(0)
            ParseFile CommonDialog.FileName
            CurrentFile = CommonDialog.FileName
            bDirty = False
        Else
            MsgBox "Unable to open file " & CommonDialog.FileName
        End If
    End If
    DrawList
End Sub

Private Sub mnuFileSave_Click()
    DoSave
End Sub

Private Sub mnuFileSaveAs_Click()
    SaveMenuFileAs
End Sub

Private Sub Parser_NodeCompleted(CompletedNode As GGZMLNode)
    If LCase$(CompletedNode.Name) = "menu" Then
        ParseNode CompletedNode, 0
    ElseIf LCase$(CompletedNode.Name) = "config" Then
        If CompletedNode.SubNodes.Count > 0 Then
            Dim i As Integer
            Dim s As String
            Dim sn As GGZMLNode
            For i = 1 To CompletedNode.SubNodes.Count
                Set sn = CompletedNode.SubNodes.Item(i)
                Select Case LCase$(sn.Name)
                    Case "systray"
                        sn.GetAttributeValue "click", s
                        If IsNumeric(s) Then
                            If CInt(s) >= 0 And CInt(s) < cmbClickType.ListCount Then
                                cmbClickType.ListIndex = CInt(s)
                            Else
                                cmbClickType.ListIndex = 0
                            End If
                        Else
                            cmbClickType.ListIndex = 0
                        End If
                    
                    Case "defaultitems"
                        sn.GetAttributeValue "position", s
                        If IsNumeric(s) Then
                            If CInt(s) >= 0 And CInt(s) < 3 Then
                                chkAddDefaultstoBottom.Value = CInt(s)
                            Else
                                chkAddDefaultstoBottom.Value = 0
                            End If
                        Else
                            chkAddDefaultstoBottom.Value = 0
                        End If
                        sn.GetAttributeValue "icon", s
                        txtSysTrayIcon.Text = s
                        sn.GetAttributeValue "tooltip", s
                        txtTooltip.Text = Replace$(s, "&tick;", "'")
                        
                        sn.GetAttributeValue "addsystray", s
                        If IsNumeric(s) Then
                            If CInt(s) >= 0 And CInt(s) < 3 Then
                                chkAddSystray.Value = CInt(s)
                            Else
                                chkAddSystray.Value = 1
                            End If
                        Else
                            chkAddEdit.Value = 1
                        End If
                        sn.GetAttributeValue "addedit", s
                        If IsNumeric(s) Then
                            If CInt(s) >= 0 And CInt(s) < 3 Then
                                chkAddEdit.Value = CInt(s)
                            Else
                                chkAddEdit.Value = 1
                            End If
                        Else
                            chkAddEdit.Value = 1
                        End If
                        
                    Case "columns"
                        sn.GetAttributeValue "count", s
                        txtColumns.Text = s
                End Select
            Next
        End If
    End If
End Sub

Private Sub txtArgs_KeyUp(KeyCode As Integer, Shift As Integer)
    If Selected >= 0 And Selected <= lstMenu.ListCount - 1 Then
        Menu(Selected).Args = txtArgs.Text
        bDirty = True
    End If
    DrawList
End Sub

Private Sub txtCaption_KeyUp(KeyCode As Integer, Shift As Integer)
    If Selected >= 0 And Selected <= lstMenu.ListCount - 1 Then
        Menu(Selected).Caption = txtCaption.Text
        bDirty = True
    End If
    DrawList
End Sub

Private Sub txtCmd_KeyUp(KeyCode As Integer, Shift As Integer)
    If Selected >= 0 And Selected <= lstMenu.ListCount - 1 Then
        Menu(Selected).Cmd = txtCmd.Text
        bDirty = True
    End If
    DrawList
End Sub

Private Sub SaveMenuFile(FileName As String)
    Dim FileNum As Integer
    Dim i As Integer
    FileNum = FreeFile
    For i = 0 To UBound(Menu) - 1
        Menu(i).Output = False
    Next
    UpdateParents
    Open FileName For Output As FileNum
    Print #FileNum, "<config>"
    Print #FileNum, "    <systray click='" & CStr(cmbClickType.ListIndex) & "' />"
    Print #FileNum, "    <defaultitems position='" & CStr(chkAddDefaultstoBottom.Value) _
            & "' icon='" & txtSysTrayIcon.Text & "' tooltip='" _
            & Replace$(txtTooltip.Text, "'", "&tick;") _
            & "' addedit='" & CStr(chkAddEdit.Value) _
            & "' addsystray='" & CStr(chkAddSystray.Value) _
            & "' />"
    Print #FileNum, "    <columns count='" & txtColumns.Text & "' />"
    Print #FileNum, "</config>"
    For i = 0 To UBound(Menu) - 1
        OutputItem FileNum, i
    Next
    Close FileNum
    bDirty = False
End Sub

Private Sub SaveMenuFileAs()
    CommonDialog.Filter = "Dyn Menus (*.dym)|*.dym|All files (*.*)|*.*"
    CommonDialog.FilterIndex = 0
    CommonDialog.DialogTitle = "Save Dynamic Menu"
    CommonDialog.ShowSave
    If Len(CommonDialog.FileName) > 0 Then
        SaveMenuFile CommonDialog.FileName
        CurrentFile = CommonDialog.FileName
    End If
End Sub

Private Sub OutputItem(FileNum As Integer, Index As Integer)
    Dim MenuS As String
    Dim i As Integer
    MenuS = String$(Menu(Index).Indent, "s") & "menu caption='" & Menu(Index).Caption & _
                "' type='" & Menu(Index).Type & _
                "' cmd='" & Menu(Index).Cmd & _
                "' args='" & Menu(Index).Args & "' "
    If Len(Menu(Index).Bitmap) > 0 Then
        MenuS = MenuS & "bitmap='" & Menu(Index).Bitmap & "' "
    End If
    If Menu(Index).ColumnBreak Then
        MenuS = MenuS & "break='1' "
    End If

    If Index < UBound(Menu) - 1 Then
        If Menu(Index).Indent < Menu(Index + 1).Indent Then
            If Menu(Index).Output = False Then
                Print #FileNum, String$(4 * Menu(Index).Indent, " ") & "<" & MenuS & ">"
                OutputItem FileNum, Index + 1
            End If
        Else
            If Menu(Index).Output = False Then
                Print #FileNum, String$(4 * Menu(Index).Indent, " ") & "<" & MenuS & "/>"
                If Menu(Index).Parent <> Menu(Index + 1).Parent And Index <> Menu(Index + 1).Parent Then
                    For i = Menu(Index).Indent - 1 To Menu(Index + 1).Indent Step -1
                        Print #FileNum, String$(4 * (i), " ") & "</" & String$(i, "s") & "menu>"
                    Next
                End If
            End If
        End If
    Else
        If Menu(Index).Output = False Then
            Print #FileNum, String$(4 * Menu(Index).Indent, " ") & "<" & MenuS & "/>"
            For i = Menu(Index).Indent To 1 Step -1
                Print #FileNum, String$(4 * (i - 1), " ") & "</" & String$(i - 1, "s") & "menu>"
            Next
        End If
    End If
    Menu(Index).Output = True
End Sub

Private Sub ParseFile(FileName As String)
    Dim FileNum As Integer
    Dim s As String
    FileNum = FreeFile
    Open FileName For Input As FileNum
    Set Parser = New GGZMLParser
    
    Do While Not EOF(FileNum)
        Input #FileNum, s
        Parser.ParseGGZML s
    Loop
    Set Parser = Nothing
    Close FileNum
End Sub

Private Sub DoSave()
    If CurrentFile = "" Then
        SaveMenuFileAs
    Else
        SaveMenuFile CurrentFile
    End If
End Sub

Private Sub SwapItems(Index1 As Integer, Index2 As Integer)
    Dim s As String
    Dim l As Long
    
    s = Menu(Index1).Caption
    Menu(Index1).Caption = Menu(Index2).Caption
    Menu(Index2).Caption = s
    s = Menu(Index1).Cmd
    Menu(Index1).Cmd = Menu(Index2).Cmd
    Menu(Index2).Cmd = s
    s = Menu(Index1).Args
    Menu(Index1).Args = Menu(Index2).Args
    Menu(Index2).Args = s
    l = Menu(Index1).Indent
    Menu(Index1).Indent = Menu(Index2).Indent
    Menu(Index2).Indent = l
    s = Menu(Index1).Bitmap
    Menu(Index1).Bitmap = Menu(Index2).Bitmap
    Menu(Index2).Bitmap = s
    l = Menu(Index1).hBitmap
    Menu(Index1).hBitmap = Menu(Index2).hBitmap
    Menu(Index2).hBitmap = l
    l = CLng(Menu(Index1).ColumnBreak)
    Menu(Index1).ColumnBreak = Menu(Index2).ColumnBreak
    Menu(Index2).ColumnBreak = CBool(l)
    s = Menu(Index1).Type
    Menu(Index1).Type = Menu(Index2).Type
    Menu(Index2).Type = s
End Sub

Private Sub CopyItem(FromIndex As Integer, ToIndex As Integer)
    Menu(ToIndex).Caption = Menu(FromIndex).Caption
    Menu(ToIndex).Cmd = Menu(FromIndex).Cmd
    Menu(ToIndex).Args = Menu(FromIndex).Args
    Menu(ToIndex).Type = Menu(FromIndex).Type
    Menu(ToIndex).Indent = Menu(FromIndex).Indent
    Menu(ToIndex).Bitmap = Menu(FromIndex).Bitmap
    Menu(ToIndex).hBitmap = Menu(FromIndex).hBitmap
    Menu(ToIndex).ColumnBreak = Menu(FromIndex).ColumnBreak
End Sub

Private Sub InsertItem(Index As Integer)
    ReDim Preserve Menu(UBound(Menu) + 1)
    Dim i As Integer
    For i = UBound(Menu) - 2 To Index Step -1
        CopyItem i, i + 1
    Next
End Sub

Private Sub DeleteItem(Index As Integer)
    Dim i As Integer
    For i = Index To UBound(Menu) - 1
        CopyItem i + 1, i
    Next
    ReDim Preserve Menu(UBound(Menu) - 1)
End Sub

Private Sub ParseNode(Node As GGZMLNode, Indent As Integer)
    Dim s As String
    Dim i As Integer
    ReDim Preserve Menu(UBound(Menu) + 1)
    Node.GetAttributeValue "caption", s
    Menu(UBound(Menu) - 1).Caption = s
    Node.GetAttributeValue "cmd", s
    Menu(UBound(Menu) - 1).Cmd = s
    Node.GetAttributeValue "args", s
    Menu(UBound(Menu) - 1).Args = s
    Node.GetAttributeValue "type", s
    Menu(UBound(Menu) - 1).Type = s
    Menu(UBound(Menu) - 1).Indent = Indent
    Node.GetAttributeValue "bitmap", s
    Menu(UBound(Menu) - 1).Bitmap = s
    Node.GetAttributeValue "break", s
    If IsNumeric(s) Then
        Menu(UBound(Menu) - 1).ColumnBreak = CBool(CInt(s))
    End If
    For i = 1 To Node.SubNodes.Count
        ParseNode Node.SubNodes(i), Indent + 1
    Next
End Sub

Public Sub CreateMenu()
    Dim i As Integer
    Dim l As Long
    Dim ItemCount As Long
    
    UpdateParents
    RootItemIndex = 0
    If hMenu > 0 Then
        DestroyMenu hMenu
        hMenu = 0
    End If
    hMenu = CreatePopupMenu()
    ReDim hMenuList(UBound(Menu))
    
    UserMenuOffset = 4
    
    Dim MII As MENUITEMINFO
    MII.cbSize = Len(MII)
    MII.wID = 0
    MII.fMask = MIIM_ID Or MIIM_TYPE
    MII.fType = MFT_STRING
    MII.dwTypeData = "&Close Menu"
    ItemCount = GetMenuItemCount(hMenu)
    InsertMenuItem hMenu, ItemCount, 1, MII
    
    If chkAddEdit.Value = 1 Then
        MII.wID = 1
        MII.fMask = MIIM_ID Or MIIM_TYPE
        MII.fType = MFT_STRING
        MII.dwTypeData = "&Edit Menu"
        ItemCount = GetMenuItemCount(hMenu)
        InsertMenuItem hMenu, ItemCount, 1, MII
    Else
        UserMenuOffset = UserMenuOffset - 1
    End If
    If chkAddSystray.Value = 1 Or IconInTray = True Then
        MII.wID = 2
        MII.fMask = MIIM_ID Or MIIM_TYPE
        MII.fType = MFT_STRING
        If IconInTray = False Then
            MII.dwTypeData = "&Add to SysTray"
        Else
            MII.dwTypeData = "&Remove from SysTray"
        End If
        ItemCount = GetMenuItemCount(hMenu)
        InsertMenuItem hMenu, ItemCount, 1, MII
    Else
        UserMenuOffset = UserMenuOffset - 1
    End If
    
    MII.wID = 3
    MII.fMask = MIIM_ID Or MIIM_TYPE
    MII.fType = MFT_SEPARATOR
    ItemCount = GetMenuItemCount(hMenu)
    If chkAddDefaultstoBottom.Value Then
        InsertMenuItem hMenu, 0, 1, MII
    Else
        InsertMenuItem hMenu, ItemCount, 1, MII
    End If
    
    For i = 0 To UBound(Menu) - 1
        AddMenu i, hMenu
    Next
    
    Dim pt As POINTAPI
    GetCursorPos pt
    l = TrackPopupMenu(hMenu, TPM_RETURNCMD Or TPM_LEFTALIGN, pt.X, pt.Y, 0, Me.hwnd, ByVal 0&)
    For i = UBound(hMenuList) - 1 To 0 Step -1
        If hMenuList(i) Then
            DestroyMenu hMenuList(i)
            hMenuList(i) = 0
        End If
    Next
    For i = 0 To UBound(Menu) - 1
        If Menu(i).hBitmap Then
            DeleteObject Menu(i).hBitmap
        End If
    Next
    ReDim hMenuList(0)
    DestroyMenu hMenu
    hMenu = 0
    DoCmd l - UserMenuOffset
End Sub

Private Sub AddMenu(Index As Integer, hMenuHandle As Long)
    Dim MII As MENUITEMINFO
    Dim Count As Long
    Dim hTempHandle As Long
    MII.cbSize = Len(MII)
    MII.wID = Index + UserMenuOffset
    
    If FileExists(Menu(Index).Bitmap) And Menu(Index).Caption <> "-" Then
        Menu(Index).hBitmap = LoadImage(0&, Menu(Index).Bitmap, IMAGE_BITMAP, GetSystemMetrics(SM_CXMENUCHECK), GetSystemMetrics(SM_CYMENUCHECK), LR_LOADFROMFILE Or LR_LOADTRANSPARENT Or LR_CREATEDIBSECTION)
        If Menu(Index).hBitmap Then
            MII.fMask = MIIM_CHECKMARKS
            MII.hbmpUnchecked = Menu(Index).hBitmap
            MII.hbmpChecked = Menu(Index).hBitmap
        End If
    End If
    
    MII.dwTypeData = Menu(Index).Caption
    MII.fMask = MII.fMask Or MIIM_ID Or MIIM_TYPE
    If Menu(Index).Caption = "-" Then
        MII.fType = MFT_SEPARATOR
    Else
        MII.fType = MFT_STRING
        MII.cch = Len(Menu(Index).Caption) + 1
    End If
    If (Menu(Index).ColumnBreak = True And ColCount <= 1) Or (Menu(Index).Parent <> -1 And Menu(Index).ColumnBreak = True) Then
        MII.fType = MII.fType Or MFT_MENUBARBREAK
    End If
    
    If Menu(Index + 1).Indent > Menu(Index).Indent Then ' this menu will have a sub menu
        hMenuList(Index) = CreatePopupMenu
        Menu(Index).hMenu = hMenuList(Index)
        MII.fMask = MII.fMask Or MIIM_SUBMENU
        MII.hSubMenu = hMenuList(Index)
    End If
    If Menu(Index).Parent = -1 Then
        hTempHandle = hMenuHandle
        If ColCount > 1 And Menu(Index).Caption <> "-" Then
            RootItemIndex = RootItemIndex + 1
            If RootItemIndex Mod (Ceil(CSng(RootItemCount), CSng(ColCount))) = 0 And RootItemIndex > 1 Then
                MII.fType = MII.fType Or MFT_MENUBARBREAK
            End If
        End If
    Else
        hTempHandle = Menu(Menu(Index).Parent).hMenu
    End If
    Count = GetMenuItemCount(hTempHandle)
    If chkAddDefaultstoBottom.Value And hTempHandle = hMenuHandle Then
        Count = Count - UserMenuOffset
    End If
    InsertMenuItem hTempHandle, Count, 1, MII
End Sub

Private Sub DoCmd(Index As Long)
    On Local Error GoTo DoCmd_Err
    If Index >= 0 Then
        If Len(Menu(Index).Type) = 0 Then
            Menu(Index).Type = "shell"
        End If
        Select Case LCase$(Menu(Index).Type)
            Case "shell"
                If Len(Menu(Index).Cmd) > 0 Then
                    Shell Menu(Index).Cmd & " " & Menu(Index).Args, vbNormalFocus
                End If
            Case "shellexec"
                If Len(Menu(Index).Cmd) > 0 And Len(Menu(Index).Args) > 0 Then
                    ShellExecute Me.hwnd, Menu(Index).Cmd, Menu(Index).Args, vbNullString, vbNullString, SW_SHOWNORMAL
                End If
        End Select
    Else
        Dim s As String
        Select Case Index + UserMenuOffset
            Case 0
                ' Close Menu - this works automagically
            Case 1
                ' Edit Menu
                s = Replace$(Command$, "/edit ", "", 1, -1, vbTextCompare)
                ShellExecute Me.hwnd, "edit", s, vbNullString, vbNullString, SW_SHOWNORMAL
            Case 2
                ' Add/Remove to/from SysTray
                If OldWindowProc Then
                    ' Icon is already added - remove it
                    RemoveTrayIcon
                    Unload Me
                Else
                    ' Icon has not been added, so do so
                    AddTrayIcon CurrentMenuName
                End If
            Case 3
                ' Separator
        End Select
    End If
    On Local Error GoTo 0
    Exit Sub
DoCmd_Err:
    If Err.Number = 53 Then
        On Local Error GoTo 0
        Exit Sub
    End If
    Resume
End Sub

Private Sub UpdateParents()
    Dim i As Integer, j As Integer
    RootItemCount = 1
    Menu(0).Parent = -1
    For i = 1 To UBound(Menu) - 1
        If Menu(i).Indent Then
            For j = i - 1 To 0 Step -1
                If Menu(j).Indent = Menu(i).Indent - 1 Then
                    Menu(i).Parent = j
                    Exit For
                End If
            Next
        Else
            Menu(i).Parent = -1
            RootItemCount = RootItemCount + 1
        End If
    Next
End Sub

Private Sub txtColumns_Change()
    If IsNumeric(txtColumns.Text) Then
        ColCount = CInt(txtColumns.Text)
        If ColCount < 1 Then ColCount = 1
    Else
        ColCount = 1
    End If
End Sub

Private Sub txtColumns_KeyPress(KeyAscii As Integer)
    bDirty = True
End Sub

Private Sub txtMenuIcon_KeyUp(KeyCode As Integer, Shift As Integer)
    If Selected >= 0 And Selected <= lstMenu.ListCount - 1 Then
        Menu(Selected).Bitmap = txtMenuIcon.Text
        bDirty = True
    End If
End Sub

Private Sub txtSysTrayIcon_KeyPress(KeyAscii As Integer)
    bDirty = True
End Sub

Private Sub txtTooltip_KeyPress(KeyAscii As Integer)
    bDirty = True
End Sub

Private Function Ceil(Sng1 As Single, Sng2 As Single) As Integer
    Dim R As Single
    R = Sng1 / Sng2
    If R - Int(R) > 0 Then
        Ceil = Int(R) + 1
    Else
        Ceil = Int(R)
    End If
End Function

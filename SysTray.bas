Attribute VB_Name = "modSystray"
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

Public Const SM_CXMENUCHECK = 71
Public Const SM_CYMENUCHECK = 72
Public Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long


' *** Icon loading functions
Public Const LR_LOADFROMFILE = &H10 ' Not NT
Public Const LR_LOADTRANSPARENT = &H20
Public Const LR_DEFAULTSIZE = &H40
Public Const LR_LOADMAP3DCOLORS = &H1000
Public Const LR_CREATEDIBSECTION = &H2000
Public Const IMAGE_BITMAP = 0
Public Const IMAGE_ICON = 1
Public Const IMAGE_CURSOR = 2
Public Const IMAGE_ENHMETAFILE = 3

Public Declare Function LoadImage Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As String, ByVal dwImageType As Long, ByVal dwDesiredWidth As Long, ByVal dwDesiredHeight As Long, ByVal dwFlags As Long) As Long
Public Declare Function DestroyIcon Lib "user32" (ByVal hIcon As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

' *** System Tray functions
Private Declare Function SetWindowLongA Lib "user32" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function CallWindowProcA Lib "user32" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function Shell_NotifyIconA Lib "shell32.dll" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long

Private Type NOTIFYICONDATA
    cbSize As Long              ' Size of the NotifyIconData structure
    hwnd As Long                ' Window handle of the window processing the icon events
    uID As Long                 ' Icon ID (to allow multiple icons per application)
    uFlags As Long              ' NIF Flags
    uCallbackMessage As Long    ' The message received for the system tray icon if NIF_MESSAGE
                                ' specified. Can be in the range 0x0400 through 0x7FFF (1024 to 32767)
    hIcon As Long               ' The memory location of our icon if NIF_ICON is specifed
    szTip As String * 64        ' Tooltip if NIF_TIP is specified (64 characters max)
End Type

' Shell_NotifyIconA() messages
Private Const NIM_ADD = &H0      ' Add icon to the System Tray
Private Const NIM_MODIFY = &H1   ' Modify System Tray icon
Private Const NIM_DELETE = &H2   ' Delete icon from System Tray

' NotifyIconData Flags
Private Const NIF_MESSAGE = &H1  ' Send event messages to the parent window
Private Const NIF_ICON = &H2     ' Display the icon
Private Const NIF_TIP = &H4      ' Use a tooltip

' The events sent appear in lParam and are as follows:
Public Const WM_MOUSEMOVE = 512
Public Const WM_LBUTTONDOWN = 513
Public Const WM_LBUTTONUP = 514
Public Const WM_LBUTTONDBLCLK = 515
Public Const WM_RBUTTONDOWN = 516
Public Const WM_RBUTTONUP = 517
Public Const WM_RBUTTONDBLCLK = 518
Public Const WM_MBUTTONDOWN = 519
Public Const WM_MBUTTONUP = 520
Public Const WM_MBUTTONDBLCLK = 521

Private Const GWL_WNDPROC = -4

Public IconInTray As Boolean
Public OldWindowProc As Long
Private LastWindowProc As Long
Private hIcon As Long

Public Function WindowProc(ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    
    If Msg = 1026 And wParam = 0 Then
        If lParam > 512 And lParam < 522 Then
            ' See frmMenuEdit.Form_Load() for more info on ItemData()
            If frmMenuEdit.cmbClickType.ItemData(frmMenuEdit.cmbClickType.ListIndex) = lParam Then
                frmMenuEdit.CreateMenu
            ElseIf frmMenuEdit.cmbClickType.ListIndex = 0 And (lParam = WM_LBUTTONUP Or lParam = WM_MBUTTONUP Or lParam = WM_RBUTTONUP) Then
                ' Any click type
                frmMenuEdit.CreateMenu
            End If
        End If
    End If
        
    ' Pass the event onto the default window handler so that all other events get
    ' handled correctly
    ' If the result of frmMenuEdit.CreateMenu is that the WindowProc is removed then we
    ' cannot call CallWindowProcA(0,...) otherwise crashes will occur.
    ' Call using what OldWindowProc used to be.
    If OldWindowProc Then
        WindowProc = CallWindowProcA(OldWindowProc, hwnd, Msg, wParam, lParam)
    ElseIf LastWindowProc Then
        WindowProc = CallWindowProcA(LastWindowProc, hwnd, Msg, wParam, lParam)
    End If
End Function

Public Sub AddTrayIcon(ByVal MenuName As String)
    If OldWindowProc <> 0 Then
        ' We have already added a tray icon! Remove the existing one.
        RemoveTrayIcon
    End If

    Dim nid As NOTIFYICONDATA
    
    ' nid.cbSize is always Len(nid)
    nid.cbSize = Len(nid)
    ' Parent window - this is the window that will process the icon events
    nid.hwnd = frmMenuEdit.hwnd
    ' Icon identifier
    nid.uID = 0
    ' We want to receive messages, show the icon and have a tooltip
    nid.uFlags = NIF_MESSAGE Or NIF_ICON Or NIF_TIP
    ' The message we will receive on an icon event
    nid.uCallbackMessage = 1026
    ' Load the icon to display if it exists. Use the icon from frmMenuEdit otherwise.
    If FileExists(frmMenuEdit.txtSysTrayIcon.Text) = True Then
        hIcon = LoadImage(0&, frmMenuEdit.txtSysTrayIcon.Text, IMAGE_ICON, 16, 16, LR_LOADFROMFILE)
        If hIcon Then
            nid.hIcon = hIcon
        Else
            nid.hIcon = frmMenuEdit.Icon
        End If
    Else
        nid.hIcon = frmMenuEdit.Icon
    End If
    
    ' Our tooltip. Use the user specified tip if it exists, otherwise use the default.
    If Len(frmMenuEdit.txtTooltip.Text) > 0 Then
        nid.szTip = frmMenuEdit.txtTooltip.Text & vbNullChar
    Else
        nid.szTip = "Click to open " & MenuName & " menu" & vbNullChar
    End If
  
    ' Add the icon to the System Tray
    Shell_NotifyIconA NIM_ADD, nid
    
    ' Set our WindowProc as the event handler for frmSystray.
    ' Save the address of the old handler in OldWindowProc
    OldWindowProc = SetWindowLongA(frmMenuEdit.hwnd, GWL_WNDPROC, AddressOf WindowProc)
    IconInTray = True
End Sub

Public Sub RemoveTrayIcon()
    Dim nid As NOTIFYICONDATA

    If IconInTray = True Then
        ' IconInTray prevents us from trying to delete the icon multiple times if
        ' RemoveTrayIcon() is called multiple times in a row.
        nid.hwnd = frmMenuEdit.hwnd
        nid.cbSize = Len(nid)
        nid.uID = 0 ' The icon identifier we set earlier
    
        ' Delete the icon
        Shell_NotifyIconA NIM_DELETE, nid
        IconInTray = False
    End If
        
    If hIcon Then
        ' If we loaded a user specified icon then free the memory used.
        DestroyIcon hIcon
        hIcon = 0
    End If
    
    If OldWindowProc <> 0 Then
        ' Set the window event handler to the previous
        SetWindowLongA frmMenuEdit.hwnd, GWL_WNDPROC, OldWindowProc
        LastWindowProc = OldWindowProc
        OldWindowProc = 0
    End If
End Sub

Public Function FileExists(strFile As String) As Boolean
    Dim FSO As FileSystemObject
    Set FSO = New FileSystemObject
    FileExists = FSO.FileExists(strFile)
    Set FSO = Nothing
End Function


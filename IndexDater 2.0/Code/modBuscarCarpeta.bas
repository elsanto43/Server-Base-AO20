Attribute VB_Name = "modBuscarCarpeta"
Option Explicit

Public Type BROWSEINFO 'parameters for the SHBrowseForFolder function
    hOwner          As Long
    pidlRoot        As Long
    pszDisplayName  As String
    lpszTitle       As String
    ulFlags         As Long
    lpfn            As Long
    lParam          As Long
    iImage          As Long
End Type

Private Type SHITEMID 'shell item id structure
    cb      As Long
    abID    As Byte
End Type

Private Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" _
    (ByVal pidl As Long, ByVal pszPath As String) As Long

Private Declare Function SHBrowseForFolder Lib "shell32.dll" Alias "SHBrowseForFolderA" _
    (lpBrowseInfo As BROWSEINFO) As Long

Private Const BIF_RETURNONLYFSDIRS = &H1
Private Const DefaultPrompt = "Select a folder"
Public Function BuscarCArpeta(Optional Prompt As String, Optional initfolder As String) As String
    Dim Folder As BROWSEINFO                'browserinfo structure
    Dim IDL As SHITEMID               'item identifier list
    Dim PointerToIdList As Long         'pointer to item identifier list
    Dim Result As Long
    Dim PathBuffer As String
    With Folder
        .hOwner = frmConfig.hWnd ' Owner Window
        .pidlRoot = 0& 'say desktop folder
        If Prompt = "" Then Prompt = DefaultPrompt    'no prompt? use default
        .pszDisplayName = initfolder
        .lpszTitle = Prompt            'set the prompt
        .ulFlags = BIF_RETURNONLYFSDIRS   ' file system directories with out this
        'flag the Control Panel, DUN and printers appear in the list
        PointerToIdList = SHBrowseForFolder(Folder) 'do the actual browse
        If PointerToIdList <> 0& Then 'user DID NOT cancel
            'now get the selected path
            PathBuffer = Space(512) ' create a buffer
            Result = SHGetPathFromIDList(ByVal PointerToIdList, ByVal PathBuffer)
            If Result Then BuscarCArpeta = Left(PathBuffer, InStr(PathBuffer, vbNullChar) - 1)  'Get the characters left of the null
        End If
    End With
End Function

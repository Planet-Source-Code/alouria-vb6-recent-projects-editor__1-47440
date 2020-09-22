VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form fRecent 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Visual Basic 6.0 Recent Projects Editor"
   ClientHeight    =   5280
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7470
   ControlBox      =   0   'False
   Icon            =   "fRecent.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5280
   ScaleWidth      =   7470
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdUnselectAll 
      Caption         =   "Unselect All"
      Height          =   345
      Left            =   1380
      TabIndex        =   8
      Top             =   4020
      Width           =   1125
   End
   Begin VB.CommandButton cmdSelectAll 
      Caption         =   "&Select All"
      Height          =   345
      Left            =   180
      TabIndex        =   7
      Top             =   4020
      Width           =   1125
   End
   Begin VB.CommandButton cmdAutoScan 
      Caption         =   "A&uto Scan"
      Height          =   345
      Left            =   3660
      TabIndex        =   9
      Top             =   4020
      Width           =   1545
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "&Remove Selected"
      Height          =   345
      Left            =   5280
      TabIndex        =   10
      Top             =   4020
      Width           =   1545
   End
   Begin MSComctlLib.ListView lvProjects 
      Height          =   2355
      Left            =   180
      TabIndex        =   4
      Top             =   1560
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   4154
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Project Name"
         Object.Width           =   4577
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Directory"
         Object.Width           =   6509
      EndProperty
   End
   Begin VB.CommandButton cmdDown 
      Height          =   315
      Left            =   6960
      Picture         =   "fRecent.frx":628A
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2040
      Width           =   375
   End
   Begin VB.CommandButton cmdUp 
      Height          =   315
      Left            =   6960
      Picture         =   "fRecent.frx":63D4
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1560
      Width           =   375
   End
   Begin VB.CommandButton cmdOkay 
      Caption         =   "&Okay"
      Default         =   -1  'True
      Height          =   345
      Left            =   3600
      TabIndex        =   11
      Top             =   4785
      Width           =   1125
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "&Apply"
      Enabled         =   0   'False
      Height          =   345
      Left            =   6135
      TabIndex        =   13
      Top             =   4785
      Width           =   1125
   End
   Begin VB.PictureBox pbBanner 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   915
      Left            =   0
      Picture         =   "fRecent.frx":651E
      ScaleHeight     =   915
      ScaleWidth      =   7935
      TabIndex        =   0
      Top             =   0
      Width           =   7935
      Begin VB.Label lblRes 
         BackStyle       =   0  'Transparent
         Caption         =   "This utility lets you modify Visual Basic 6.0's list of recently accessed projects to help keep your interface clean."
         Height          =   435
         Index           =   3
         Left            =   420
         TabIndex        =   2
         Top             =   360
         Width           =   5595
      End
      Begin VB.Line Line1 
         Index           =   0
         X1              =   -30
         X2              =   7890
         Y1              =   885
         Y2              =   885
      End
      Begin VB.Label lblRes 
         BackStyle       =   0  'Transparent
         Caption         =   "Recently Accessed Projects"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   1
         Top             =   60
         Width           =   5595
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   345
      Left            =   4860
      TabIndex        =   12
      Top             =   4785
      Width           =   1125
   End
   Begin VB.Label lblCopy 
      BackStyle       =   0  'Transparent
      Caption         =   "  Alouria, Inc."
      ForeColor       =   &H80000010&
      Height          =   255
      Index           =   1
      Left            =   -15
      TabIndex        =   14
      Top             =   4455
      Width           =   990
   End
   Begin VB.Label Label1 
      Caption         =   "Place a checkmark besides the projects you want to removed from Visual Basic 6.0's recently accessed list."
      Height          =   375
      Left            =   180
      TabIndex        =   3
      Top             =   1080
      Width           =   7095
   End
   Begin VB.Label lblCopy 
      Caption         =   "  Alouria, Inc."
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   0
      TabIndex        =   15
      Top             =   4470
      Width           =   960
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000010&
      Index           =   0
      X1              =   15
      X2              =   7995
      Y1              =   4575
      Y2              =   4575
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000014&
      Index           =   1
      X1              =   15
      X2              =   7995
      Y1              =   4590
      Y2              =   4590
   End
End
Attribute VB_Name = "fRecent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ***************************************************************************
' Module:        fRecent.frm
'
' Description:   Main form for the VBRecent application.
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 15-AUG-2003  Michael Harrington               mikeh@alouria.com
'              Original creation.
' ***************************************************************************

Option Explicit

'========================================================
'Class Constructors and Deconstructors
'========================================================

Private Sub Form_Load()
' ***************************************************************************
' Routine:       Form_Load
'
' Description:   Constructor routine for the form.
'
' Parameters:    <none>
'
' Returns:       <none>
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 15-AUG-2003  Michael Harrington               mikeh@alouria.com
'              Original Creation
' ***************************************************************************

    Dim oReg                        As New cRegistry
    Dim lIndex                      As Long
    Dim lvItem                      As ListItem
    
    oReg.ClassKey = HKEY_CURRENT_USER
    oReg.SectionKey = "SOFTWARE\Microsoft\Visual Basic\6.0\RecentFiles"
   
    'check to see if key exists
    If oReg.KeyExists = True Then
        'retrieve the mail root
        For lIndex = 1 To 100
            oReg.ValueKey = Trim$(Str(lIndex))
            If Trim$(oReg.Value) <> "" Then
                Set lvItem = lvProjects.ListItems.Add(, oReg.Value, _
                 ProjectName(oReg.Value))
                lvItem.ListSubItems.Add , , ProjectPath(oReg.Value)
            End If
        Next lIndex
    Else
        MsgBox "The application could not locate an installed copy of Visual Basic 6.0", vbCritical, App.ProductName
        Me.Hide
        Unload Me
    End If
    
    If lvProjects.ListItems.Count > 0 Then
        lvProjects.SelectedItem.Selected = True
        lvProjects.SelectedItem.EnsureVisible
    End If
    
    Set oReg = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
' ***************************************************************************
' Routine:       Form_Unload
'
' Description:   Deconstructor routine for the form.
'
' Parameters:    Cancel - ByRef argument to cancel the unload.
'
' Returns:       <none>
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 15-AUG-2003  Michael Harrington               mikeh@alouria.com
'              Original Creation
' ***************************************************************************

End Sub

'========================================================
'Control Events
'========================================================

Private Sub cmdApply_Click()
' ***************************************************************************
' Routine:       cmdApply_Click
'
' Description:   Saves the current display & order to the registry
'
' Parameters:    <none>
'
' Returns:       <none>
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 15-AUG-2003  Michael Harrington               mikeh@alouria.com
'              Original Creation
' ***************************************************************************

    Dim lCount                      As Long
    Dim lIndex                      As Long
    Dim lvItem                      As ListItem
    Dim oReg                        As New cRegistry
    
    'check for items that have been marked for removal
    For Each lvItem In lvProjects.ListItems
        If lvItem.Checked = True Then lCount = lCount + 1
    Next
    
    'verify removal before proceeding
    If lCount > 0 Then
        If MsgBox("You have marked several items to be removed, " & _
         "do you want to remove them now?", vbYesNo, App.ProductName) = vbYes Then Call cmdRemove_Click
    End If

    'set registry location
    oReg.ClassKey = HKEY_CURRENT_USER
    oReg.SectionKey = "SOFTWARE\Microsoft\Visual Basic\6.0\RecentFiles"
    
    'remove old list
    For lIndex = 1 To 100
        oReg.ValueKey = Trim$(Str(lIndex))
        On Error Resume Next
            oReg.DeleteValue
        On Error GoTo 0
    Next lIndex

    'save new values
    For lIndex = 1 To lvProjects.ListItems.Count
        oReg.CreateKey
        oReg.ValueKey = Trim$(Str(lIndex))
        oReg.ValueType = REG_SZ
        oReg.Value = lvProjects.ListItems(lIndex).Key
    Next lIndex
    
    'disable the apply button
    cmdApply.Enabled = False
    
    'destroy the objects
    Set lvItem = Nothing
    Set oReg = Nothing
    
End Sub

Private Sub cmdAutoScan_Click()
' ***************************************************************************
' Routine:       cmdAutoScan_Click
'
' Description:   Scans for projects that no longer exist.
'
' Parameters:    <none>
'
' Returns:       <none>
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 15-AUG-2003  Michael Harrington               mikeh@alouria.com
'              Original Creation
' ***************************************************************************

    Dim lFileSize                   As Long
    Dim lCount                      As Long
    Dim lvItem                      As ListItem
    
    'uncheck all items
    For Each lvItem In lvProjects.ListItems
        lvItem.Checked = False
    Next

    'verify each project
    For Each lvItem In lvProjects.ListItems
    
        'check to see if the project exists
        On Error Resume Next
            Open lvItem.Key For Binary As #1
            
            If Err.Number <> 0 Then
                'project wasn't found
                lFileSize = 0
            Else
                'path wasn't found so filesize is 0
                lFileSize = LOF(1)
                Close #1
            End If
        On Error GoTo 0
                  
        'if filesize is 0 delete the project file created and mark for registry removal
        If lFileSize = 0 Then
            On Error Resume Next
                Kill lvItem.Key
            On Error GoTo 0
            lvItem.Checked = True
        End If
    Next lvItem
    
    'count the number of projects selected
    For Each lvItem In lvProjects.ListItems
        If lvItem.Checked = True Then lCount = lCount + 1
    Next
    
    'display total
    MsgBox "A total of " & lCount & " project entries were identified as missing.", vbInformation, App.ProductName

    'destroy objects
    Set lvItem = Nothing
    
End Sub

Private Sub cmdCancel_Click()
' ***************************************************************************
' Routine:       cmdCancel_Click
'
' Description:   Exit the application.
'
' Parameters:    <none>
'
' Returns:       <none>
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 15-AUG-2003  Michael Harrington               mikeh@alouria.com
'              Original Creation
' ***************************************************************************

    'verify the user doesn't want to save changes
    If cmdApply.Enabled = True Then
        If MsgBox("Are you sure you want to discard your changes?", vbYesNo, _
         App.ProductName) = vbNo Then Exit Sub
    End If
    
    'exit the application
    Me.Hide
    Unload Me
    
End Sub

Private Sub cmdDown_Click()
' ***************************************************************************
' Routine:       cmdDown_Click
'
' Description:   Moves the selected project down one.
'
' Parameters:    <none>
'
' Returns:       <none>
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 15-AUG-2003  Michael Harrington               mikeh@alouria.com
'              Original Creation
' ***************************************************************************

    'verify there are projects listed
    If lvProjects.ListItems.Count > 0 Then
    
        'move the selected item down
        MoveDown lvProjects.SelectedItem
    
        'enable the apply button since changes were made
        cmdApply.Enabled = True
    
    End If

End Sub

Private Sub cmdOkay_Click()
' ***************************************************************************
' Routine:       cmdOkay_Click
'
' Description:   Saves any changes, then exists.
'
' Parameters:    <none>
'
' Returns:       <none>
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 15-AUG-2003  Michael Harrington               mikeh@alouria.com
'              Original Creation
' ***************************************************************************

    'save any changes that were made
    If cmdApply.Enabled = True Then Call cmdApply_Click
    
    'exit the application
    Call cmdCancel_Click

End Sub

Private Sub cmdRemove_Click()
' ***************************************************************************
' Routine:       cmdRemove_Click
'
' Description:   Removes all checked project from the list.
'
' Parameters:    <none>
'
' Returns:       <none>
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 15-AUG-2003  Michael Harrington               mikeh@alouria.com
'              Original Creation
' ***************************************************************************

    Dim bRescan                     As Boolean
    Dim lvItem                      As ListItem
    
    
    'set the rescan property to true
    bRescan = True
    
    'keep on rechecking until all selected projects are removed
    Do Until bRescan = False
    
        'set the rescan to false
        bRescan = False
        
        'iterate through each project
        For Each lvItem In lvProjects.ListItems
        
            'verify the project was checked
            If lvItem.Checked = True Then
            
                'remove the project from the list
                lvProjects.ListItems.Remove lvItem.Key
                
                'since the collection has changed we need to restart the process
                bRescan = True
                Exit For
            End If
        
        Next
    
    Loop
    
    'enable the apply button since changes were made
    If cmdApply.Enabled = False Then cmdApply.Enabled = True

    'destroy the lvItem object
    Set lvItem = Nothing
    
End Sub

Private Sub cmdSelectAll_Click()
' ***************************************************************************
' Routine:       cmdSelectAll_Click
'
' Description:   Selects all projects in the list.
'
' Parameters:    <none>
'
' Returns:       <none>
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 15-AUG-2003  Michael Harrington               mikeh@alouria.com
'              Original Creation
' ***************************************************************************

    Dim lIndex                      As Long
    
    'exit if there are no projects listed
    If lvProjects.ListItems.Count = 0 Then Exit Sub
    
    'unselect each item
    For lIndex = 1 To lvProjects.ListItems.Count
        lvProjects.ListItems(lIndex).Checked = True
    Next lIndex

    'enable the apply button since changes were made
    cmdApply.Enabled = True
End Sub

Private Sub cmdUnselectAll_Click()
' ***************************************************************************
' Routine:       cmdUnselectAll_Click
'
' Description:   Unselects all projects in the list.
'
' Parameters:    <none>
'
' Returns:       <none>
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 15-AUG-2003  Michael Harrington               mikeh@alouria.com
'              Original Creation
' ***************************************************************************

    Dim lIndex                      As Long
    
    'exit if there are no projects listed
    If lvProjects.ListItems.Count = 0 Then Exit Sub
    
    'unselect each item
    For lIndex = 1 To lvProjects.ListItems.Count
        lvProjects.ListItems(lIndex).Checked = False
    Next lIndex

    'enable the apply button since changes were made
    cmdApply.Enabled = True
End Sub

Private Sub cmdUp_Click()
' ***************************************************************************
' Routine:       cmdUp_Click
'
' Description:   Moves the selected project up one.
'
' Parameters:    <none>
'
' Returns:       <none>
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 15-AUG-2003  Michael Harrington               mikeh@alouria.com
'              Original Creation
' ***************************************************************************

    'verify there are projects listed
    If lvProjects.ListItems.Count > 0 Then
    
        'move the selected item down
        MoveUp lvProjects.SelectedItem
    
        'enable the apply button since changes were made
        cmdApply.Enabled = True
    
    End If
    
End Sub

Private Sub lvProjects_BeforeLabelEdit(Cancel As Integer)
' ***************************************************************************
' Routine:       lvProjects_BeforeLabelEdit
'
' Description:   Stops the user from being able to edit a label.
'
' Parameters:    Cancel - ByRef argument stopping the label editing.
'
' Returns:       <none>
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 15-AUG-2003  Michael Harrington               mikeh@alouria.com
'              Original Creation
' ***************************************************************************

    'stop the editing
    Cancel = 1
    
End Sub

'========================================================
'Private Functions & Procedures
'========================================================

Private Sub MoveDown(li As ListItem)
' ***************************************************************************
' Routine:       MoveDown
'
' Description:   Moves a list item down one.
'
' Parameters:    li - List Item to move down.
'
' Returns:       <none>
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' ??-???-????  Unknown Writer                   ??????????
'              Original Creation
' 15-AUG-2003  Michael Harrington               mikeh@alouria.com
'              Modified for style & use in project.
' ***************************************************************************

    Dim iIndex                       As Long
    Dim lIndex                       As Long
    Dim sKey                         As String

    'retrieve the current index
    lIndex = li.Index
    
    'set focus on the projects list
    lvProjects.SetFocus
        
    'if we are at the end, exit
    If lIndex = lvProjects.ListItems.Count Then Exit Sub

    'create a new list item
    With lvProjects.ListItems.Add(lIndex + 2, , li.Text, li.Icon, li.SmallIcon)
        For iIndex = 1 To lvProjects.ColumnHeaders.Count - 1
            .SubItems(iIndex) = li.SubItems(iIndex)
        Next iIndex
    End With
   
    'store the old key
    sKey = li.Key
       
    'remove the old item
    lvProjects.ListItems.Remove lIndex
    
    With lvProjects.ListItems(lIndex + 1)
    
        'set the key for the new item
        .Key = sKey
        
        'set the focus to the new item
        .Selected = True
        .EnsureVisible
    
    End With
    
End Sub

Private Sub MoveUp(li As ListItem)
' ***************************************************************************
' Routine:       MoveUp
'
' Description:   Moves a list item up one.
'
' Parameters:    li - List Item to move up.
'
' Returns:       <none>
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' ??-???-????  Unknown Writer                   ??????????
'              Original Creation
' 15-AUG-2003  Michael Harrington               mikeh@alouria.com
'              Modified for style & use in project.
' ***************************************************************************

    Dim iIndex                       As Long
    Dim lIndex                       As Long
    Dim sKey                         As String
    
    'retrieve the current index
    lIndex = li.Index
    
    'set focus on the projects list
    lvProjects.SetFocus
    
    'if we are at the beginning, exit
    If lIndex = 1 Then Exit Sub


    'create a new list item
    With lvProjects.ListItems.Add(lIndex - 1, , li.Text, li.Icon, li.SmallIcon)
        For iIndex = 1 To lvProjects.ColumnHeaders.Count - 1
            .SubItems(iIndex) = li.SubItems(iIndex)
        Next iIndex
    End With
   
    'store the old key
    sKey = li.Key
   
    'remove old item
    lvProjects.ListItems.Remove lIndex + 1
    
    With lvProjects.ListItems(lIndex - 1)
        
        'set the key for the new item
        .Key = sKey
        
        'set focus to the new item
        .Selected = True
        .EnsureVisible
    
    End With
    
End Sub

Private Function ProjectName(sPath As String) As String
' ***************************************************************************
' Routine:       ProjectName
'
' Description:   Returns the name of a project from a Path + File string
'
' Parameters:    sPath - Path + File the project is located
'
' Returns:       Project file
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 15-AUG-2003  Michael Harrington               mikeh@alouria.com
'              Original Creation
' ***************************************************************************

    'retrieve the project file
    ProjectName = Mid$(sPath, InStrRev(sPath, "\") + 1, Len(sPath) - InStrRev(sPath, "\"))

End Function

Private Function ProjectPath(sPath As String) As String
' ***************************************************************************
' Routine:       ProjectName
'
' Description:   Returns the path of a project from a Path + File string
'
' Parameters:    sPath - Path + File the project is located
'
' Returns:       Project path
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 15-AUG-2003  Michael Harrington               mikeh@alouria.com
'              Original Creation
' ***************************************************************************
    
    'retrieve the project path
    ProjectPath = Mid$(sPath, 1, InStrRev(sPath, "\"))

End Function



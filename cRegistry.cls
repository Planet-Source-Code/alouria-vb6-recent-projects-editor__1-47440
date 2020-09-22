VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cRegistry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ***************************************************************************
' Module:       cRegistry.cls
'
' Description:  A nice class wrapper around the registry functions.  Allows
'               searching, deletion, modification, and addition of Keys or
'               values.
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 21-JUL-1997 Steve McMahon                     steve@vbaccelerator.com
'             Original creation.
' 29-APR-1998 Steve McMahon                     steve@vbaccelerator.com
'             Fixed GPF in EnumerateValues
'             Added support for all registry types, not just strings
'             Put all declares in local class
'             Added VB5 enums
'             Added CreateKey and DeleteKey methods
' 02-JAN-1999 Steve McMahon                     steve@vbaccelerator.com
'             The CreateExeAssociation method failed to set up the
'               association correctly if the optional document icon was not
'               provided.
'             Added new parameters to CreateExeAssociation to set up other
'               other standard handlers: Print, Add, New
'             Provided the CreateAdditionalEXEAssociations method to allow
'               non-standard menu items to be added (for example, right-
'               click on a .VBF file.  VB installs Run and Make menu items).
' 08-FEB-2000 Steve McMahon                     steve@vbaccelerator.com
'             Ensure CreateExeAssociation and related items sets up the
'               registry keys in the HKEY_LOCAL_MACHINE\SOFTWARE\Classes
'               branch as well as the HKEY_CLASSES_ROOT branch.
' 17-AUG-2002 Michael Harrington                mikeh@alouria.com
'             Reformatted the class for easier reading.
'               standard format.
' ***************************************************************************

Option Explicit

'========================================================
'Constants
'========================================================

'Private Constants

'Registry Specific Access Rights
Private Const KEY_QUERY_VALUE               As Long = &H1
Private Const KEY_SET_VALUE                 As Long = &H2
Private Const KEY_CREATE_SUB_KEY            As Long = &H4
Private Const KEY_ENUMERATE_SUB_KEYS        As Long = &H8
Private Const KEY_NOTIFY                    As Long = &H10
Private Const KEY_CREATE_LINK               As Long = &H20
Private Const KEY_ALL_ACCESS                As Long = &H3F

'Open/Create Options
Private Const REG_OPTION_NON_VOLATILE       As Long = 0&
Private Const REG_OPTION_VOLATILE           As Long = &H1

'Key creation/open disposition
Private Const REG_CREATED_NEW_KEY           As Long = &H1
Private Const REG_OPENED_EXISTING_KEY       As Long = &H2

'masks for the predefined standard access types
Private Const STANDARD_RIGHTS_ALL           As Long = &H1F0000
Private Const SPECIFIC_RIGHTS_ALL           As Long = &HFFFF

'Define severity codes
Private Const ERROR_SUCCESS                 As Long = 0&
Private Const ERROR_ACCESS_DENIED           As Long = 5
Private Const ERROR_INVALID_DATA            As Long = 13&
Private Const ERROR_MORE_DATA               As Long = 234 'dderror
Private Const ERROR_NO_MORE_ITEMS           As Long = 259

'========================================================
'Types
'========================================================

'public types

'private types
Private Type SECURITY_ATTRIBUTES
  nLength As Long
  lpSecurityDescriptor As Long
  bInheritHandle As Boolean
End Type

Private Type FILETIME
  dwLowDateTime As Long
  dwHighDateTime As Long
End Type

'========================================================
'API Calls
'========================================================

'private calls
Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" ( _
        ByVal hkey As Long, _
        ByVal lpSubKey As String, _
        ByVal ulOptions As Long, _
        ByVal samDesired As Long, _
        phkResult As Long _
    ) As Long

Private Declare Function RegSetValueExStr Lib "advapi32" Alias "RegSetValueExA" ( _
        ByVal hkey As Long, _
        ByVal lpValueName As String, _
        ByVal Reserved As Long, _
        ByVal dwType As Long, _
        ByVal szData As String, _
        ByVal cbData As Long _
    ) As Long

Private Declare Function RegSetValueExLong Lib "advapi32" Alias "RegSetValueExA" ( _
        ByVal hkey As Long, _
        ByVal lpValueName As String, _
        ByVal Reserved As Long, _
        ByVal dwType As Long, _
        ByRef szData As Long, _
        ByVal cbData As Long _
    ) As Long

Private Declare Function RegSetValueExByte Lib "advapi32" Alias "RegSetValueExA" ( _
        ByVal hkey As Long, _
        ByVal lpValueName As String, _
        ByVal Reserved As Long, _
        ByVal dwType As Long, _
        ByRef szData As Byte, _
        ByVal cbData As Long _
    ) As Long

Private Declare Function RegCloseKey Lib "advapi32" ( _
        ByVal hkey As Long _
    ) As Long

Private Declare Function RegQueryValueExStr Lib "advapi32" Alias "RegQueryValueExA" ( _
        ByVal hkey As Long, _
        ByVal lpValueName As String, _
        ByVal lpReserved As Long, _
        ByRef lpType As Long, _
        ByVal szData As String, _
        ByRef lpcbData As Long _
    ) As Long

Private Declare Function RegQueryValueExLong Lib "advapi32" Alias "RegQueryValueExA" ( _
        ByVal hkey As Long, _
        ByVal lpValueName As String, _
        ByVal lpReserved As Long, _
        ByRef lpType As Long, _
        ByRef szData As Long, _
        ByRef lpcbData As Long _
    ) As Long

Private Declare Function RegQueryValueExByte Lib "advapi32" Alias "RegQueryValueExA" ( _
        ByVal hkey As Long, _
        ByVal lpValueName As String, _
        ByVal lpReserved As Long, _
        ByRef lpType As Long, _
        ByRef szData As Byte, _
        ByRef lpcbData As Long _
    ) As Long
   
Private Declare Function RegCreateKeyEx Lib "advapi32" Alias "RegCreateKeyExA" ( _
        ByVal hkey As Long, _
        ByVal lpSubKey As String, _
        ByVal Reserved As Long, _
        ByVal lpClass As String, _
        ByVal dwOptions As Long, _
        ByVal samDesired As Long, _
        ByRef lpSecurityAttributes As SECURITY_ATTRIBUTES, _
        ByRef phkResult As Long, _
        ByRef lpdwDisposition As Long _
    ) As Long

Private Declare Function RegEnumKeyEx Lib "advapi32.dll" Alias "RegEnumKeyExA" ( _
        ByVal hkey As Long, _
        ByVal dwIndex As Long, _
        ByVal lpName As String, _
        ByRef lpcbName As Long, _
        ByVal lpReserved As Long, _
        ByVal lpClass As String, _
        ByVal lpcbClass As Long, _
        ByRef lpftLastWriteTime As FILETIME _
    ) As Long

Private Declare Function RegEnumKey Lib "advapi32.dll" Alias "RegEnumKeyA" ( _
        ByVal hkey As Long, _
        ByVal dwIndex As Long, _
        ByVal lpName As String, _
        ByVal cbName As Long _
    ) As Long

Private Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" ( _
        ByVal hkey As Long, _
        ByVal dwIndex As Long, _
        ByVal lpValueName As String, _
        ByRef lpcbValueName As Long, _
        ByVal lpReserved As Long, _
        ByVal lpType As Long, _
        ByVal lpData As Long, _
        ByVal lpcbData As Long _
    ) As Long
   
Private Declare Function RegEnumValueLong Lib "advapi32.dll" Alias "RegEnumValueA" ( _
        ByVal hkey As Long, _
        ByVal dwIndex As Long, _
        ByVal lpValueName As String, _
        ByRef lpcbValueName As Long, _
        ByVal lpReserved As Long, _
        ByRef lpType As Long, _
        ByRef lpData As Long, _
        ByRef lpcbData As Long _
    ) As Long

Private Declare Function RegEnumValueStr Lib "advapi32.dll" Alias "RegEnumValueA" ( _
        ByVal hkey As Long, _
        ByVal dwIndex As Long, _
        ByVal lpValueName As String, _
        ByRef lpcbValueName As Long, _
        ByVal lpReserved As Long, _
        ByRef lpType As Long, _
        ByVal lpData As String, _
        ByRef lpcbData As Long _
    ) As Long

Private Declare Function RegEnumValueByte Lib "advapi32.dll" Alias "RegEnumValueA" ( _
        ByVal hkey As Long, _
        ByVal dwIndex As Long, _
        ByVal lpValueName As String, _
        ByRef lpcbValueName As Long, _
        ByVal lpReserved As Long, _
        ByRef lpType As Long, _
        ByRef lpData As Byte, _
        ByRef lpcbData As Long _
    ) As Long

Private Declare Function RegQueryInfoKey Lib "advapi32.dll" Alias "RegQueryInfoKeyA" ( _
        ByVal hkey As Long, _
        ByVal lpClass As String, _
        ByRef lpcbClass As Long, _
        ByVal lpReserved As Long, _
        ByRef lpcSubKeys As Long, _
        ByRef lpcbMaxSubKeyLen As Long, _
        ByRef lpcbMaxClassLen As Long, _
        ByRef lpcValues As Long, _
        ByRef lpcbMaxValueNameLen As Long, _
        ByRef lpcbMaxValueLen As Long, _
        ByRef lpcbSecurityDescriptor As Long, _
        ByRef lpftLastWriteTime As Any _
   ) As Long

Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" ( _
        ByVal hkey As Long, _
        ByVal lpSubKey As String _
    ) As Long

Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" ( _
        ByVal hkey As Long, _
        ByVal lpValueName As String _
    ) As Long

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
        lpvDest As Any, _
        lpvSource As Any, _
        ByVal cbCopy As Long _
    )

Private Declare Function ExpandEnvironmentStrings Lib "kernel32" Alias "ExpandEnvironmentStringsA" ( _
        ByVal lpSrc As String, _
        ByVal lpDst As String, _
        ByVal nSize As Long _
    ) As Long

'========================================================
'Module Level Variables
'========================================================

Public Enum ERegistryClassConstants
    HKEY_CLASSES_ROOT = &H80000000
    HKEY_CURRENT_USER = &H80000001
    HKEY_LOCAL_MACHINE = &H80000002
    HKEY_USERS = &H80000003
End Enum

Public Enum ERegistryValueTypes            'Predefined Registry Value Types
    REG_NONE = (0)                         'No value type
    REG_SZ = (1)                           'Unicode nul terminated string
    REG_EXPAND_SZ = (2)                    'Unicode nul terminated string w/enviornment var
    REG_BINARY = (3)                       'Free form binary
    REG_DWORD = (4)                        '32-bit number
    REG_DWORD_LITTLE_ENDIAN = (4)          '32-bit number (same as REG_DWORD)
    REG_DWORD_BIG_ENDIAN = (5)             '32-bit number
    REG_LINK = (6)                         'Symbolic Link (unicode)
    REG_MULTI_SZ = (7)                     'Multiple Unicode strings
    REG_RESOURCE_LIST = (8)                'Resource list in the resource map
    REG_FULL_RESOURCE_DESCRIPTOR = (9)     'Resource list in the hardware description
    REG_RESOURCE_REQUIREMENTS_LIST = (10)
End Enum

Private m_hClassKey         As Long
Private m_sSectionKey       As String
Private m_sValueKey         As String
Private m_vValue            As Variant
Private m_sSetValue         As String
Private m_vDefault          As Variant
Private m_eValueType        As ERegistryValueTypes

'========================================================
'Class Constructors and Deconstructors
'========================================================

'========================================================
'Public Property Functions
'========================================================

Public Property Get ClassKey() As ERegistryClassConstants
    ClassKey = m_hClassKey
End Property

Public Property Let ClassKey(ByVal eKey As ERegistryClassConstants)
    m_hClassKey = eKey
End Property

Public Property Get Default() As Variant
    Default = m_vDefault
End Property

Public Property Let Default(ByVal vDefault As Variant)
    m_vDefault = vDefault
End Property

Public Property Get KeyExists() As Boolean
    Dim hkey                As Long
    
    If RegOpenKeyEx(m_hClassKey, m_sSectionKey, 0, 1, hkey) = ERROR_SUCCESS Then
        KeyExists = True
        RegCloseKey hkey
    Else
        KeyExists = False
    End If
End Property

Public Property Get SectionKey() As String
    SectionKey = m_sSectionKey
End Property

Public Property Let SectionKey(ByVal sSectionKey As String)
    m_sSectionKey = sSectionKey
End Property

Public Property Get Value() As Variant
    Dim vValue              As Variant
    Dim cData               As Long
    Dim sData               As String
    Dim ordType             As Long
    Dim e                   As Long
    Dim hkey                As Long

    e = RegOpenKeyEx(m_hClassKey, m_sSectionKey, 0, KEY_QUERY_VALUE, hkey)
    'ApiRaiseIf e

    e = RegQueryValueExLong(hkey, m_sValueKey, 0&, ordType, 0&, cData)
    If e And e <> ERROR_MORE_DATA Then
        Value = m_vDefault
        Exit Property
    End If
    
    m_eValueType = ordType
    Select Case ordType
    Case REG_DWORD, REG_DWORD_LITTLE_ENDIAN
        Dim iData As Long
        e = RegQueryValueExLong(hkey, m_sValueKey, 0&, _
                               ordType, iData, cData)
        vValue = CLng(iData)
        
    Case REG_DWORD_BIG_ENDIAN  ' Unlikely, but you never know
        Dim dwData As Long
        e = RegQueryValueExLong(hkey, m_sValueKey, 0&, _
                               ordType, dwData, cData)
        vValue = SwapEndian(dwData)
        
    Case REG_SZ, REG_MULTI_SZ ' Same thing to Visual Basic
        sData = String$(cData - 1, 0)
        e = RegQueryValueExStr(hkey, m_sValueKey, 0&, _
                               ordType, sData, cData)
        vValue = sData
        
    Case REG_EXPAND_SZ
        sData = String$(cData - 1, 0)
        e = RegQueryValueExStr(hkey, m_sValueKey, 0&, _
                               ordType, sData, cData)
        vValue = ExpandEnvStr(sData)
        
    ' Catch REG_BINARY and anything else
    Case Else
        Dim abData() As Byte
        ReDim abData(cData)
        e = RegQueryValueExByte(hkey, m_sValueKey, 0&, _
                                ordType, abData(0), cData)
        vValue = abData
        
    End Select
    Value = vValue
    
End Property

Public Property Let Value(ByVal vValue As Variant)
    Dim ordType             As Long
    Dim c                   As Long
    Dim hkey                As Long
    Dim e                   As Long
    Dim lCreate             As Long
    Dim tSA                 As SECURITY_ATTRIBUTES

    'Open or Create the key
    e = RegCreateKeyEx(m_hClassKey, m_sSectionKey, 0, "", REG_OPTION_NON_VOLATILE, _
                 KEY_ALL_ACCESS, tSA, hkey, lCreate)
    
    If e Then
        Err.Raise 26001, App.EXEName & ".cRegistry", "Failed to set registry value Key: '" & m_hClassKey & "',Section: '" & m_sSectionKey & "',Key: '" & m_sValueKey & "' to value: '" & m_vValue & "'"
    Else

        Select Case m_eValueType
        Case REG_BINARY
            If (VarType(vValue) = vbArray + vbByte) Then
                Dim ab() As Byte
                ab = vValue
                ordType = REG_BINARY
                c = UBound(ab) - LBound(ab) - 1
                e = RegSetValueExByte(hkey, m_sValueKey, 0&, ordType, ab(0), c)
            Else
                Err.Raise 26001
            End If
        Case REG_DWORD, REG_DWORD_BIG_ENDIAN, REG_DWORD_LITTLE_ENDIAN
            If (VarType(vValue) = vbInteger) Or (VarType(vValue) = vbLong) Then
                Dim i As Long
                i = vValue
                ordType = REG_DWORD
                e = RegSetValueExLong(hkey, m_sValueKey, 0&, ordType, i, 4)
            End If
        Case REG_SZ, REG_EXPAND_SZ
            Dim s As String, iPos As Long
            s = vValue
            ordType = REG_SZ
            ' Assume anything with two non-adjacent percents is expanded string
            iPos = InStr(s, "%")
            If iPos Then
                If InStr(iPos + 2, s, "%") Then ordType = REG_EXPAND_SZ
            End If
            c = Len(s) + 1
            e = RegSetValueExStr(hkey, m_sValueKey, 0&, ordType, s, c)
            
        ' User should convert to a compatible type before calling
        Case Else
            e = ERROR_INVALID_DATA
            
        End Select
        
        If Not e Then
            m_vValue = vValue
        Else
            Err.Raise vbObjectError + 1048 + 26001, App.EXEName & ".cRegistry", "Failed to set registry value Key: '" & m_hClassKey & "',Section: '" & m_sSectionKey & "',Key: '" & m_sValueKey & "' to value: '" & m_vValue & "'"
        End If
        
        'Close the key
        RegCloseKey hkey
    
    End If
    
End Property

Public Property Get ValueKey() As String
    ValueKey = m_sValueKey
End Property

Public Property Let ValueKey(ByVal sValueKey As String)
    m_sValueKey = sValueKey
End Property

Public Property Get ValueType() As ERegistryValueTypes
    ValueType = m_eValueType
End Property

Public Property Let ValueType(ByVal eValueType As ERegistryValueTypes)
    m_eValueType = eValueType
End Property


'========================================================
'Public Functions and Procedures
'========================================================

Public Sub CreateAdditionalEXEAssociations(ByVal sClassName As String, ParamArray vItems() As Variant)
   On Error Resume Next
    
    Dim iItems              As Long
    Dim iItem               As Long
   
   iItems = UBound(vItems) + 1
   If (iItems Mod 3) <> 0 Or (Err.Number <> 0) Then
      Err.Raise vbObjectError + 1048 + 26004, App.EXEName & ".cRegistry", "Invalid parameter list passed to CreateAdditionalEXEAssociations - expected Name/Text/Command"
   Else
      ' Check if it exists:
      SectionKey = sClassName
      If Not (KeyExists) Then
         Err.Raise vbObjectError + 1048 + 26005, App.EXEName & ".cRegistry", "Error - attempt to create additional associations before class defined."
      Else
         For iItem = 0 To iItems - 1 Step 3
            ValueType = REG_SZ
            SectionKey = sClassName & "\shell\" & vItems(iItem)
            ValueKey = ""
            pSetClassValue vItems(iItem + 1)
            SectionKey = sClassName & "\shell\" & vItems(iItem) & "\command"
            ValueKey = ""
            pSetClassValue vItems(iItem + 2)
         Next iItem
      End If
   End If
   
End Sub

Public Sub CreateEXEAssociation(ByVal sExePath As String, ByVal sClassName As _
  String, ByVal sClassDescription As String, ByVal sAssociation As String, _
  Optional ByVal sOpenMenuText As String = "&Open", Optional ByVal _
  bSupportPrint As Boolean = False, Optional ByVal sPrintMenuText As _
  String = "&Print", Optional ByVal bSupportNew As Boolean = False, Optional _
  ByVal sNewMenuText As String = "&New", Optional ByVal bSupportInstall As _
  Boolean = False, Optional ByVal sInstallMenuText As String = "", Optional _
  ByVal lDefaultIconIndex As Long = -1)
   
   ' Check if path is wrapped in quotes:
   sExePath = Trim$(sExePath)
   If (Left$(sExePath, 1) <> """") Then
      sExePath = """" & sExePath
   End If
   If (Right$(sExePath, 1) <> """") Then
      sExePath = sExePath & """"
   End If
    
    ' Create the .File to Class association:
   SectionKey = "." & sAssociation
   ValueType = REG_SZ
   ValueKey = ""
   pSetClassValue sClassName
   
   ' Create the Class shell open command:
   SectionKey = sClassName
   pSetClassValue sClassDescription
   
   SectionKey = sClassName & "\shell\open"
   If (sOpenMenuText = "") Then sOpenMenuText = "&Open"
   ValueKey = ""
   pSetClassValue sOpenMenuText
   SectionKey = sClassName & "\shell\open\command"
   ValueKey = ""
   pSetClassValue sExePath & " ""%1"""
   
   If (bSupportPrint) Then
      SectionKey = sClassName & "\shell\print"
      If (sPrintMenuText = "") Then sPrintMenuText = "&Print"
      ValueKey = ""
      pSetClassValue sPrintMenuText
      SectionKey = sClassName & "\shell\print\command"
      ValueKey = ""
      pSetClassValue sExePath & " /p ""%1"""
   End If
   
   If (bSupportInstall) Then
      If (sInstallMenuText = "") Then
         sInstallMenuText = "&Install " & sAssociation
      End If
      SectionKey = sClassName & "\shell\add"
      ValueKey = ""
      pSetClassValue sInstallMenuText
      SectionKey = sClassName & "\shell\add\command"
      ValueKey = ""
      pSetClassValue sExePath & " /a ""%1"""
   End If
   
   If (bSupportNew) Then
      SectionKey = sClassName & "\shell\new"
      ValueKey = ""
      If (sNewMenuText = "") Then sNewMenuText = "&New"
      pSetClassValue sNewMenuText
      SectionKey = sClassName & "\shell\new\command"
      ValueKey = ""
      pSetClassValue sExePath & " /n ""%1"""
   End If
   
   If lDefaultIconIndex > -1 Then
      SectionKey = sClassName & "\DefaultIcon"
      ValueKey = ""
      pSetClassValue sExePath & "," & CStr(lDefaultIconIndex)
   End If
    
End Sub

Public Function CreateKey() As Boolean
    Dim tSA                 As SECURITY_ATTRIBUTES
    Dim hkey                As Long
    Dim lCreate             As Long
    Dim e                   As Long

    'Open or Create the key
    e = RegCreateKeyEx(m_hClassKey, m_sSectionKey, 0, "", REG_OPTION_NON_VOLATILE, _
                 KEY_ALL_ACCESS, tSA, hkey, lCreate)
    If e Then
        Err.Raise 26001, App.EXEName & ".cRegistry", "Failed to create registry Key: '" & m_sSectionKey
    Else
        CreateKey = (e = ERROR_SUCCESS)
        'Close the key
        RegCloseKey hkey
    End If
End Function

Public Function DeleteKey() As Boolean
    Dim e                   As Long
    
    e = RegDeleteKey(m_hClassKey, m_sSectionKey)
    If e Then
        Err.Raise 26001, App.EXEName & ".cRegistry", "Failed to delete registry Key: '" & m_hClassKey & "',Section: '" & m_sSectionKey
    Else
        DeleteKey = (e = ERROR_SUCCESS)
    End If
    
End Function

Public Function DeleteValue() As Boolean
    Dim e                   As Long
    Dim hkey                As Long

    e = RegOpenKeyEx(m_hClassKey, m_sSectionKey, 0, KEY_ALL_ACCESS, hkey)
    If e Then
        Err.Raise 26001, App.EXEName & ".cRegistry", "Failed to open key '" & m_hClassKey & "',Section: '" & m_sSectionKey & "' for delete access"
    Else
        e = RegDeleteValue(hkey, m_sValueKey)
        If e Then
            Err.Raise 26001, App.EXEName & ".cRegistry", "Failed to delete registry Key: '" & m_hClassKey & "',Section: '" & m_sSectionKey & "',Key: '" & m_sValueKey
        Else
            DeleteValue = (e = ERROR_SUCCESS)
        End If
    End If

End Function

Public Function EnumerateValues(ByRef sKeyNames() As String, ByRef iKeyCount As Long) As Boolean
    Dim lResult             As Long
    Dim hkey                As Long
    Dim sName               As String
    Dim lNameSize           As Long
    Dim sData               As String
    Dim lIndex              As Long
    Dim cJunk               As Long
    Dim cNameMax            As Long
    Dim ft                  As Currency
   
   ' Log "EnterEnumerateValues"

   iKeyCount = 0
   Erase sKeyNames()
    
   lIndex = 0
   lResult = RegOpenKeyEx(m_hClassKey, m_sSectionKey, 0, KEY_QUERY_VALUE, hkey)
   If (lResult = ERROR_SUCCESS) Then
      ' Log "OpenedKey:" & m_hClassKey & "," & m_sSectionKey
      lResult = RegQueryInfoKey(hkey, "", cJunk, 0, _
                               cJunk, cJunk, cJunk, cJunk, _
                               cNameMax, cJunk, cJunk, ft)
       Do While lResult = ERROR_SUCCESS
   
           'Set buffer space
           lNameSize = cNameMax + 1
           sName = String$(lNameSize, 0)
           If (lNameSize = 0) Then lNameSize = 1
           
           ' Log "Requesting Next Value"
         
           'Get value name:
           lResult = RegEnumValue(hkey, lIndex, sName, lNameSize, _
                                  0&, 0&, 0&, 0&)
           ' Log "RegEnumValue returned:" & lResult
           If (lResult = ERROR_SUCCESS) Then
       
                ' Although in theory you can also retrieve the actual
                ' value and type here, I found it always (ultimately) resulted in
                ' a GPF, on Win95 and NT.  Why?  Can anyone help?
       
               sName = Left$(sName, lNameSize)
               ' Log "Enumerated value:" & sName
                 
               iKeyCount = iKeyCount + 1
               ReDim Preserve sKeyNames(1 To iKeyCount) As String
               sKeyNames(iKeyCount) = sName
           End If
           lIndex = lIndex + 1
       Loop
   End If
   If (hkey <> 0) Then
      RegCloseKey hkey
   End If

   ' Log "Exit Enumerate Values"
   EnumerateValues = True
   Exit Function
   
EnumerateValuesError:
   If (hkey <> 0) Then
      RegCloseKey hkey
   End If
   Err.Raise vbObjectError + 1048 + 26003, App.EXEName & ".cRegistry", Err.Description
   Exit Function
End Function

Public Function EnumerateSections(ByRef sSect() As String, ByRef iSectCount As Long) As Boolean
    On Error GoTo EnumerateSectionsError
    
    Dim lResult             As Long
    Dim hkey                As Long
    Dim dwReserved          As Long
    Dim szBuffer            As String
    Dim lBuffSize           As Long
    Dim lIndex              As Long
    Dim lType               As Long
    Dim sCompKey            As String
    Dim iPos                As Long


   iSectCount = 0
   Erase sSect
   
   lIndex = 0

   lResult = RegOpenKeyEx(m_hClassKey, m_sSectionKey, 0, KEY_ENUMERATE_SUB_KEYS, hkey)
   Do While lResult = ERROR_SUCCESS
       'Set buffer space
       szBuffer = String$(255, 0)
       lBuffSize = Len(szBuffer)
      
      'Get next value
       lResult = RegEnumKey(hkey, lIndex, szBuffer, lBuffSize)
                             
       If (lResult = ERROR_SUCCESS) Then
           iSectCount = iSectCount + 1
           ReDim Preserve sSect(1 To iSectCount) As String
           iPos = InStr(szBuffer, Chr$(0))
           If (iPos > 0) Then
              sSect(iSectCount) = Left(szBuffer, iPos - 1)
           Else
              sSect(iSectCount) = Left(szBuffer, lBuffSize)
           End If
       End If
       
       lIndex = lIndex + 1
   Loop
   If (hkey <> 0) Then
      RegCloseKey hkey
   End If
   EnumerateSections = True
   Exit Function

EnumerateSectionsError:
   If (hkey <> 0) Then
      RegCloseKey hkey
   End If
   Err.Raise vbObjectError + 1048 + 26002, App.EXEName & ".cRegistry", Err.Description
   Exit Function
End Function

'========================================================
'Private Functions and Procedures
'========================================================

Private Sub pSetClassValue(ByVal sValue As String)
    Dim sSection            As String
   
    ClassKey = HKEY_CLASSES_ROOT
    Value = sValue
    sSection = SectionKey
    ClassKey = HKEY_LOCAL_MACHINE
    SectionKey = "SOFTWARE\Classes\" & sSection
    Value = sValue
    SectionKey = sSection
End Sub

Private Function SwapEndian(ByVal dw As Long) As Long
    CopyMemory ByVal VarPtr(SwapEndian) + 3, dw, 1
    CopyMemory ByVal VarPtr(SwapEndian) + 2, ByVal VarPtr(dw) + 1, 1
    CopyMemory ByVal VarPtr(SwapEndian) + 1, ByVal VarPtr(dw) + 2, 1
    CopyMemory SwapEndian, ByVal VarPtr(dw) + 3, 1
End Function
Private Function ExpandEnvStr(sData As String) As String
    Dim c As Long, s As String
    ' Get the length
    s = "" ' Needed to get around Windows 95 limitation
    c = ExpandEnvironmentStrings(sData, s, c)
    ' Expand the string
    s = String$(c - 1, 0)
    c = ExpandEnvironmentStrings(sData, s, c)
    ExpandEnvStr = s
End Function





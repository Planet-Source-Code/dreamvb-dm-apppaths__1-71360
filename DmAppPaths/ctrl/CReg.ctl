VERSION 5.00
Begin VB.UserControl CReg 
   ClientHeight    =   360
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   360
   InvisibleAtRuntime=   -1  'True
   Picture         =   "CReg.ctx":0000
   ScaleHeight     =   360
   ScaleWidth      =   360
   ToolboxBitmap   =   "CReg.ctx":00D7
End
Attribute VB_Name = "CReg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Private Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, ByVal lpSecurityAttributes As Long, phkResult As Long, lpdwDisposition As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long

Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByRef lpData As Any, ByVal cbData As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByRef lpData As Any, ByRef lpcbData As Long) As Long
Private Declare Function ExpandEnvironmentStrings Lib "kernel32.dll" Alias "ExpandEnvironmentStringsA" (ByVal lpSrc As String, ByVal lpDst As String, ByVal nSize As Long) As Long
Private Declare Function RegEnumKey Lib "advapi32.dll" Alias "RegEnumKeyA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, ByVal cbName As Long) As Long
Private Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, ByRef lpcbValueName As Long, ByVal lpReserved As Long, ByRef lpType As Long, ByRef lpData As Byte, ByRef lpcbData As Long) As Long

'Registry keys consts
Enum TKeys
    HKEY_LOCAL_MACHINE = &H80000002
    HKEY_CLASSES_ROOT = &H80000000
    HKEY_CURRENT_CONFIG = &H80000005
    HKEY_CURRENT_USER = &H80000001
    HKEY_DYN_DATA = &H80000006
    HKEY_USERS = &H80000003
    HKEY_PERFORMANCE_DATA = &H80000004
End Enum

'Registry Datatypes
Enum TDatatype
    REG_SZ = 1
    REG_EXPAND_SZ = 2
    REG_DWORD = 4
    REG_MULTI_SZ = 7
End Enum

Private Const ERROR_SUCCESS = 0
Private Const KEY_ALL_ACCESS = &H3F
Private Const KEY_SET_VALUE = &H2
Private Const REG_OPTION_NON_VOLATILE As Long = 0
Private Const KEY_READ As Long = &H20019

'The Maximum data length
Private Const MAX_LENGTH As Long = 2048

Private m_Key As TKeys
Private m_SubKey As String

Private Sub UserControl_Initialize()
    m_Key = HKEY_LOCAL_MACHINE
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    m_Key = PropBag.ReadProperty("Key", m_Key)
    m_SubKey = PropBag.ReadProperty("SubKey", m_SubKey)
End Sub

Private Sub UserControl_Resize()
    UserControl.Size 360, 360
End Sub

Public Property Get Key() As TKeys
    Key = m_Key
End Property

Public Property Let Key(ByVal vNewKey As TKeys)
    m_Key = vNewKey
End Property

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Key", m_Key, HKEY_LOCAL_MACHINE)
    Call PropBag.WriteProperty("SubKey", m_SubKey, "")
End Sub

Public Property Get SubKey() As String
    SubKey = m_SubKey
End Property

Public Property Let SubKey(ByVal vNewSubKey As String)
    m_SubKey = vNewSubKey
End Property

Public Function KeyExsists(Optional KeyValue As String = "") As Long
Dim KeyHandle As Long
Dim KeyPath As String
Dim iRet As Long
    'Check if a reg Key Exsists.
    'Key to check
    KeyPath = (m_SubKey & KeyValue)
    'Open Reg Key
    iRet = RegOpenKeyEx(m_Key, KeyPath, 0, KEY_ALL_ACCESS, KeyHandle)
    KeyExsists = Abs(KeyHandle <> ERROR_SUCCESS)
    'Close the open key
    RegCloseKey KeyHandle
    'Clear up
    KeyPath = vbNullString
End Function

Public Function CreateKey(Optional KeyValue As String = "") As Long
'Create a new key
Dim iRet As Long
Dim KeyCreate As Long
Dim KeyPath As String
Dim Dispose As Long

    KeyPath = (m_SubKey & KeyValue)
    
    'Check if the key already exsists
    If RegOpenKeyEx(m_Key, KeyPath, 0, KEY_ALL_ACCESS, KeyCreate) = ERROR_SUCCESS Then
        'Key aready exsists no need to create it agian.
        CreateKey = 2
        KeyPath = vbNullString
        Call CloseRegKey(KeyCreate)
        Exit Function
    Else
        'Create the key
        If RegCreateKeyEx(m_Key, KeyPath, 0, "", _
            REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, 0&, KeyCreate, Dispose) <> ERROR_SUCCESS Then
            'Create key faild
            CreateKey = 0
        Else
            'Good result
            CreateKey = 1
        End If
        
        'Close the key
        Call CloseRegKey(KeyCreate)
        'Clear up
        KeyPath = vbNullString
    End If
    
End Function

Function ValueExists(ValueName As String) As Boolean
Dim hKeyResult As Long
    'Checks if a Value exisits
    'Open the key and see if it exsists
    If RegOpenKeyEx(m_Key, m_SubKey, 0, KEY_ALL_ACCESS, hKeyResult) <> ERROR_SUCCESS Then
        Exit Function
    End If
    
    'Get the size of the value
    If RegQueryValueEx(hKeyResult, ValueName, 0&, _
         0, ByVal 0&, 0&) <> ERROR_SUCCESS Then
         'Value was not found
         ValueExists = False
         Exit Function
    Else
        'Value found
        ValueExists = True
    End If
    'Close the key
    Call CloseRegKey(hKeyResult)
    
End Function

Public Function DeleteKey(Optional KeyValue As String = "") As Long
'Delete key
Dim iRet As Long
Dim KeyPath As String

    'Key to delete
    KeyPath = (m_SubKey & KeyValue)
    'Delete the key
    iRet = RegDeleteKey(m_Key, KeyPath)
    
    If (iRet <> ERROR_SUCCESS) Then
        DeleteKey = 0
    Else
        DeleteKey = 1
    End If
    
    'Clear up
    KeyPath = vbNullString
    
End Function

Public Function SetValue(ValueName As String, ByVal ValueData, DataType As TDatatype) As Long
Dim KeyCreate As Long
Dim iRet As Long
Dim sBuff As String
Dim vCount As Long

    'Open the key
    iRet = RegOpenKeyEx(m_Key, m_SubKey, 0, KEY_ALL_ACCESS, KeyCreate)
    If (iRet <> ERROR_SUCCESS) Then
            Exit Function
    Else
        Select Case DataType
            Case REG_SZ, REG_EXPAND_SZ
                'Write Reg Value for strings
                iRet = RegSetValueEx(KeyCreate, ValueName, 0&, _
                DataType, ByVal CStr(ValueData), Len(ValueData))
            Case REG_DWORD
                'Write Reg Value for DWORDS
                iRet = RegSetValueEx(KeyCreate, ValueName, 0&, _
                DataType, CLng(ValueData), 4)
            Case REG_MULTI_SZ
                'Writes a String list
                If IsArray(ValueData) Then
                    For vCount = 0 To UBound(ValueData)
                        'Build the string
                        sBuff = sBuff + ValueData(vCount) + Chr(0)
                    Next vCount
                    'End terminator
                    sBuff = sBuff + Chr(0)
                    'Write REG_MULTI_SZ Value
                    iRet = RegSetValueEx(KeyCreate, ValueName, 0&, _
                    DataType, ByVal CStr(sBuff), Len(sBuff))
                End If
        End Select
    End If
    
    SetValue = 1
    'Close the openkey
    Call CloseRegKey(KeyCreate)
    'Clean up
    sBuff = vbNullString
    vCount = 0
    
End Function

Function GetValue(ValueName As String, DataType As TDatatype, Optional Defaut)
Dim KeyCreate As Long
Dim bSize As Long
Dim bStr As String
Dim iRet As Long
Dim RegDWord As Long
Dim lType As Long
Dim lpDst As String
Dim tmp(0) As String

    'Type to read Strings, DWORDS
    lType = DataType
    
    'Open the key and see if it exsists
    If RegOpenKeyEx(m_Key, m_SubKey, 0, KEY_ALL_ACCESS, KeyCreate) <> ERROR_SUCCESS Then
        Exit Function
    End If
    
    'Get the size of the value
    If RegQueryValueEx(KeyCreate, ValueName, 0&, _
        lType, ByVal 0&, bSize) <> ERROR_SUCCESS Then
        'Return Default Value
        GetValue = Defaut
    End If

    Select Case DataType
        Case REG_SZ, REG_EXPAND_SZ
            'Create string buffer
            bStr = String(bSize, Chr(0))
            'Read string value
            If RegQueryValueEx(KeyCreate, ValueName, 0&, _
                lType, ByVal bStr, bSize) <> ERROR_SUCCESS Then
                'Return Default Value
                GetValue = Defaut
            End If
            'Strip away NULL Chars from the string
            bStr = TrimNull(bStr)
            
            'Return the value
            If (DataType = REG_EXPAND_SZ) Then
                'Get the size of the string
                bSize = ExpandEnvironmentStrings(bStr, lpDst, 1)
                'Create a buffer to hold the new string
                lpDst = Space(bSize)
                'Extract Environment varaible
                iRet = ExpandEnvironmentStrings(bStr, lpDst, bSize)
                'Trip away null chars
                bStr = TrimNull(lpDst)
                'Return string
                GetValue = bStr
            Else
                'Return normal string
                GetValue = bStr
            End If
            'Close the Reg key
            RegCloseKey KeyCreate
        Case REG_DWORD
            'Read Numeric value
            If RegQueryValueEx(KeyCreate, ValueName, 0&, _
                 lType, RegDWord, bSize) <> ERROR_SUCCESS Then
                GetValue = Defaut
            Else
                GetValue = RegDWord
            End If
        Case REG_MULTI_SZ
            'Set the buffer size
            If (bSize - 1) < 0 Then
                'Returns an array of size zero
                GetValue = tmp
            Else
                bStr = String(bSize - 1, Chr(0))
                'Get the values data
                iRet = RegQueryValueEx(KeyCreate, ValueName, 0&, lType, ByVal bStr, bSize)
                'Return Array
                GetValue = Split(bStr, Chr(0))
            End If
    End Select
    
    'Clear up
    bSize = 0
    Erase tmp
    lpDst = vbNullString
    bStr = vbNullString
    
End Function

Function DeleteValue(ValueName As String) As Boolean
Dim iRet As Long
Dim hKeyResult As Long
    'Delete a regvalue
    
    'Open the key and see if it exsists
    If RegOpenKeyEx(m_Key, m_SubKey, 0, KEY_ALL_ACCESS, hKeyResult) <> ERROR_SUCCESS Then
        Exit Function
    End If
    
    iRet = RegDeleteValue(hKeyResult, ValueName)
    
    If (iRet = ERROR_SUCCESS) Then
        DeleteValue = True
    End If
    
    'Close the key
    Call CloseRegKey(hKeyResult)
End Function

Function GetSubKeys(Optional KeyValue As String = "") As Collection
Dim hKeyResult As Long
Dim kName As String
Dim kSize As Long
Dim TmpCol As New Collection
Dim kCount As Long

    'Check if the key is found
    If RegOpenKeyEx(m_Key, m_SubKey & KeyValue, 0, KEY_ALL_ACCESS, hKeyResult) <> ERROR_SUCCESS Then
        RegCloseKey hKeyResult
        'Key was not found so we return array size of zero
        Set GetSubKeys = TmpCol
        'Clear up
        Exit Function
    End If
    
    'Keyname size
    kSize = MAX_LENGTH
    'Create Buffer
    kName = Space(kSize)
    'Get the subkey names
    Do While RegEnumKey(hKeyResult, kCount, kName, kSize) = ERROR_SUCCESS
        'Strip waway the NULL char and add the Keyname to the collection
        TmpCol.Add TrimNull(kName)
        kCount = kCount + 1
    Loop
    
    'Return the colelction
    Set GetSubKeys = TmpCol
    
    'Clear up
    Set TmpCol = Nothing
    kName = vbNullString
    kSize = 0
    'Close OpenKey
    Call CloseRegKey(hKeyResult)
End Function

Function GetValueNames(Optional KeyValue As String = "") As Collection
Dim hKeyResult As Long
Dim iRet As Long
Dim vName As String
Dim vSize As Long
Dim vCount As Long
Dim dLen As Long
Dim TmpCol As New Collection

    'Check if the key is found
    If RegOpenKeyEx(m_Key, m_SubKey & KeyValue, 0, KEY_READ, hKeyResult) Then
        Call CloseRegKey(hKeyResult)
        Exit Function
    End If
    
    Do
        'Set the value name size
        vSize = 255
        dLen = vSize
        'Create buffer
        vName = String(vSize, Chr(0))
        'Get the value name
        iRet = RegEnumValue(hKeyResult, vCount, vName, _
        vSize, ByVal 0&, ByVal 0&, ByVal 0&, dLen)

        If (iRet = ERROR_SUCCESS) Then
            'Fill the collection with the value names
            TmpCol.Add Left(vName, vSize)
        Else
            Exit Do
        End If
        
        vCount = vCount + 1
    Loop While (iRet = ERROR_SUCCESS)
    
    Set GetValueNames = TmpCol
    'Clear up
    Call CloseRegKey(hKeyResult)
    Set TmpCol = Nothing
    vName = vbNullString
    vSize = 0
    dLen = 0
    vCount = 0
    
    
End Function

Private Function TrimNull(lpStr As String) As String
Dim e_pos As Integer
    'Trim away the NULL char of a string
    e_pos = InStr(lpStr, Chr$(0))

    If (e_pos) Then
        TrimNull = Left$(lpStr, e_pos - 1)
    Else
        TrimNull = lpStr
    End If

End Function

Private Sub CloseRegKey(lngKey As Long)
Dim Result As Long
On Error Resume Next
    'Close the open Regkey
    Result = RegCloseKey(lngKey)
    'Check that the regkey was closed
    If (Result <> ERROR_SUCCESS) Then
        Err.Raise 9 + vbObject, "CloseRegKey", "RegCloseKey Faild."
    End If
    
End Sub

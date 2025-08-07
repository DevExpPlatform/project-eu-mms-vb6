Attribute VB_Name = "mod_APIFiles_DecRou"
Option Explicit

Private Declare Function DeleteFile Lib "Kernel32" Alias "DeleteFileA" (ByVal lpFileName As String) As Long
'Private Declare Function GetComputerName Lib "Kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function GetTempPath Lib "Kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function RemoveDirectory Lib "Kernel32" Alias "RemoveDirectoryA" (ByVal lpPathName As String) As Long
Private Declare Function SetFileAttributes Lib "Kernel32" Alias "SetFileAttributesA" (ByVal lpFileName As String, ByVal dwFileAttributes As Long) As Long

Public Function EmptyDir(ByVal strPath As String) As Boolean
        
    strPath = Fix_Paths(strPath)
            
    If FDExist(strPath, True) Then
        Dim I           As Integer
        Dim ItemFile()  As String
        Dim retValue    As String
        
        retValue = Dir$(strPath, vbArchive Or vbHidden Or vbNormal Or vbReadOnly Or vbSystem)
            
        Do While retValue <> ""
            Push_Array ItemFile, retValue
                
            retValue = Dir$
        Loop
    
        ' Erase all found files
        '
        If chk_Array(ItemFile) Then
            For I = 0 To UBound(ItemFile)
                KillFile strPath & ItemFile(I)
            Next I
        End If
    End If

    EmptyDir = True

End Function

Public Function get_UsrName() As String
    
    Dim sUser   As String
    'Dim sComputer As String
    Dim lpBuff  As String * 1024

    GetUserName lpBuff, Len(lpBuff)
    sUser = Left$(lpBuff, (InStr(1, lpBuff, vbNullChar)) - 1)
    lpBuff = ""
    
    'GetComputerName lpBuff, Len(lpBuff)
    'sComputer = Left$(lpBuff, (InStr(1, lpBuff, vbNullChar)) - 1)
    'lpBuff = ""

    get_UsrName = sUser

End Function

Public Function get_UsrTmpPath()

    Const MAX_PATH = 260

    Dim sFolder As String   ' Name of the folder
    Dim lRet    As Long     ' Return Value
    
    sFolder = String(MAX_PATH, 0)
    lRet = GetTempPath(MAX_PATH, sFolder)
    
    If lRet <> 0 Then
        get_UsrTmpPath = Left(sFolder, InStr(sFolder, Chr(0)) - 1)
    Else
        get_UsrTmpPath = vbNullString
    End If

End Function

Public Function KillAll(ByVal strPath As String, ByVal OnlyFiles As Boolean) As Boolean
        
    On Error GoTo ErrHandler
    
    Dim I           As Integer
    Dim ItemDir()   As String
    Dim ItemFile()  As String
    Dim retValue    As String
    Dim SrchStart   As Integer
    Dim SrchFnsh    As Integer
    Dim SubDirFnd   As Boolean
    Dim TmpDir      As String
    
    strPath = Fix_Paths(strPath)
    
    If FDExist(strPath, True) Then
        ReDim ItemDir(0)
        ReDim ItemFile(0)

ReDoChk:
        SubDirFnd = False
        
        For I = SrchStart To SrchFnsh
            If ItemDir(0) <> "" Then
                TmpDir = Fix_Paths(ItemDir(I))
            Else
                TmpDir = strPath
            End If
                
            retValue = Dir$(TmpDir, vbDirectory Or vbArchive Or vbHidden Or vbNormal Or vbReadOnly Or vbSystem)
            
            Do While retValue <> ""
                If retValue <> "." And retValue <> ".." Then
                    If (GetAttr(TmpDir & retValue) And vbDirectory) = vbDirectory Then
                        If ItemDir(0) <> "" Then
                            ReDim Preserve ItemDir(UBound(ItemDir) + 1)
                        End If
                        
                        ItemDir(UBound(ItemDir)) = TmpDir & retValue
                    
                        SubDirFnd = True
                    Else
                        If ItemFile(0) <> "" Then
                            ReDim Preserve ItemFile(UBound(ItemFile) + 1)
                        End If
                        
                        ItemFile(UBound(ItemFile)) = TmpDir & retValue
                    End If
                End If
                
                retValue = Dir$
            Loop
        Next I
        
        If SubDirFnd Then
            SrchStart = IIf(SrchFnsh > 0, SrchFnsh + 1, 0)
            SrchFnsh = UBound(ItemDir)

            GoTo ReDoChk
        End If
    
        ' Erase all found files
        '
        If ItemFile(0) <> "" Then
            For I = 0 To UBound(ItemFile)
                KillFile ItemFile(I)
            Next I
        End If
        
        ' Erase all found Dirs
        '
        If OnlyFiles = False Then
            If ItemDir(0) <> "" Then
                For I = UBound(ItemDir) To 0 Step -1
                    KillDirectory ItemDir(I)
                Next I
            End If
            
            KillDirectory strPath
        End If
        
        KillAll = True
    Else
        KillAll = True
    End If

    Exit Function

ErrHandler:

End Function

Private Function KillDirectory(ByVal DirName As String) As Boolean
    
    Dim RValue As Long

    RValue = SetFileAttributes(DirName, vbNormal)
    
    If RValue = 0 Then
        'GoSub WriteError
    Else
        RValue = RemoveDirectory(DirName)
        
        'If RValue = 0 Then GoSub WriteError
    End If

End Function

Private Function KillFile(ByVal FileName As String) As Boolean
     
    Dim RValue As Long
    
    RValue = SetFileAttributes(FileName, vbNormal)
    
    If RValue = 0 Then
'        GoSub WriteError
    Else
        RValue = DeleteFile(FileName)
    
        'If RValue = 0 Then GoSub WriteError
    End If
                        
End Function


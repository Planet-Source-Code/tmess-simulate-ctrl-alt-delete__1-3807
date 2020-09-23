VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsActiveProcess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

    
Private Const TH32CS_SNAPPROCESS As Long = 2&
Private Const MAX_PATH As Integer = 260

Private Type PROCESSENTRY32
    dwSize As Long
    cntUsage As Long
    th32ProcessID As Long
    th32DefaultHeapID As Long
    th32ModuleID As Long
    cntThreads As Long
    th32ParentProcessID As Long
    pcPriClassBase As Long
    dwFlags As Long
    szExeFile As String * MAX_PATH
    End Type

Private Declare Function CreateToolhelpSnapshot Lib "kernel32" Alias "CreateToolhelp32Snapshot" (ByVal lFlags As Long, ByVal lProcessId As Long) As Long


Private Declare Function ProcessFirst Lib "kernel32" Alias "Process32First" (ByVal hSnapShot As Long, uProcess As PROCESSENTRY32) As Long


Private Declare Function ProcessNext Lib "kernel32" Alias "Process32Next" (ByVal hSnapShot As Long, uProcess As PROCESSENTRY32) As Long


Private Declare Sub CloseHandle Lib "kernel32" (ByVal hPass As Long)
    Dim ListOfActiveProcess() As PROCESSENTRY32
    
Public Function szExeFile(ByVal Index As Long) As String

    szExeFile = ListOfActiveProcess(Index).szExeFile
    
End Function


Public Function dwFlags(ByVal Index As Long) As Long

    dwFlags = ListOfActiveProcess(Index).dwFlags

    
End Function


Public Function pcPriClassBase(ByVal Index As Long) As Long

    pcPriClassBase = ListOfActiveProcess(Index).pcPriClassBase
End Function



Public Function th32ParentProcessID(ByVal Index As Long) As Long

    th32ParentProcessID = ListOfActiveProcess(Index).th32ParentProcessID
End Function



Public Function cntThreads(ByVal Index As Long) As Long

    cntThreads = ListOfActiveProcess(Index).cntThreads
End Function



Public Function thModuleID(ByVal Index As Long) As Long

    thModuleID = ListOfActiveProcess(Index).th32ModuleID
End Function



Public Function th32DefaultHeapID(ByVal Index As Long) As Long

    th32DefaultHeapID = ListOfActiveProcess(Index).th32DefaultHeapID
End Function



Public Function th32ProcessID(ByVal Index As Long) As Long

    th32ProcessID = ListOfActiveProcess(Index).th32ProcessID
End Function

Public Function cntUsage(ByVal Index As Long) As Long

    cntUsage = ListOfActiveProcess(Index).cntUsage
End Function

Public Function dwSize(ByVal Index As Long) As Long

    dwSize = ListOfActiveProcess(Index).dwSize
End Function

Public Function GetActiveProcess() As Long

    Dim hToolhelpSnapshot As Long
    Dim tProcess As PROCESSENTRY32
    Dim r As Long, i As Integer
    hToolhelpSnapshot = CreateToolhelpSnapshot(TH32CS_SNAPPROCESS, 0&)


    If hToolhelpSnapshot = 0 Then
        GetActiveProcess = 0
        Exit Function
    End If


    With tProcess
        .dwSize = Len(tProcess)
        r = ProcessFirst(hToolhelpSnapshot, tProcess)
        ReDim Preserve ListOfActiveProcess(20)


        Do While r
            i = i + 1
            If i Mod 20 = 0 Then ReDim Preserve ListOfActiveProcess(i + 20)
            ListOfActiveProcess(i) = tProcess
            r = ProcessNext(hToolhelpSnapshot, tProcess)
        Loop

    End With

    GetActiveProcess = i
    Call CloseHandle(hToolhelpSnapshot)
End Function





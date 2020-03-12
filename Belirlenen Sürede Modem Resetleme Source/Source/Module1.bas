Attribute VB_Name = "Module1"
Private Declare Function CreatePipe Lib "kernel32" (phReadPipe As Long, phWritePipe As Long, lpPipeAttributes As Any, ByVal nSize As Long) As Long
Public Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, ByVal lpBuffer As String, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, ByVal lpOverlapped As Any) As Long
Private Declare Function CreateProcessA Lib "kernel32" (ByVal lpApplicationName As Long, ByVal lpCommandLine As String, lpProcessAttributes As SECURITY_ATTRIBUTES, lpThreadAttributes As SECURITY_ATTRIBUTES, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, ByVal lpEnvironment As Long, ByVal lpCurrentDirectory As Long, lpStartupInfo As STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) As Long
Public Declare Function CloseHandle Lib "kernel32" (ByVal hHandle As Long) As Long
Private Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Long
End Type
Private Type STARTUPINFO
    cb As Long
    lpReserved As Long
    lpDesktop As Long
    lpTitle As Long
    dwX As Long
    dwY As Long
    dwXSize As Long
    dwYSize As Long
    dwXCountChars As Long
    dwYCountChars As Long
    dwFillAttribute As Long
    dwFlags As Long
    wShowWindow As Integer
    cbReserved2 As Integer
    lpReserved2 As Long
    hStdInput As Long
    hStdOutput As Long
    hStdError As Long
End Type
Private Type PROCESS_INFORMATION
    hProcess As Long
    hThread As Long
    dwProcessID As Long
    dwThreadID As Long
End Type
    Public hWritePipe As Long
    Public hReadPipe As Long
    Public proc As PROCESS_INFORMATION

Public Function Komut(mCommand As String)
    Dim start As STARTUPINFO
    Dim sa As SECURITY_ATTRIBUTES
    If Len(mCommand) = 0 Then MsgBox "Komut Satýrý Yanlýþ...!", vbCritical: Exit Function
    
    sa.nLength = Len(sa)
    sa.bInheritHandle = 1&
    sa.lpSecurityDescriptor = 0&
    
    If CreatePipe(hReadPipe, hWritePipe, sa, 0) = 0 Then MsgBox "Baþarýsýz Oldu.!", vbCritical: Exit Function
    start.cb = Len(start)
    start.dwFlags = &H100& Or &H1
    start.hStdOutput = hWritePipe
    start.hStdError = hWritePipe
    If CreateProcessA(0&, mCommand, sa, sa, 1&, &H20&, 0&, 0&, start, proc) <> 1 Then MsgBox "Dosya ya da Komut Bulunamadý", vbCritical: Exit Function
    
    CloseHandle (hWritePipe)
Form1.Timer1.Enabled = True
End Function




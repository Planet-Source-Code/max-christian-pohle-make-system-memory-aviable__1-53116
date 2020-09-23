Attribute VB_Name = "Module1"
Option Explicit
'written by   Max Christian Pohle
'             http://www.coderonline.de/

Dim A() As Byte

Public Type MEMORYSTATUS
    dwLength As Long
    dwMemoryLoad As Long
    dwTotalPhys As Long
    dwAvailPhys As Long
    dwTotalPageFile As Long
    dwAvailPageFile As Long
    dwTotalVirtual As Long
    dwAvailVirtual As Long
End Type

Public Declare Sub GlobalMemoryStatus Lib "kernel32" (lpBuffer As MEMORYSTATUS)

Function FreeMem() As Long
    Dim MemStat As MEMORYSTATUS
    Dim lngAvailPhys As Long
    
    Call GlobalMemoryStatus(MemStat)
    lngAvailPhys = Round(MemStat.dwAvailPhys / 1024 / 1024)
    
    'the cleanup-process...
    ReDim Preserve A(0 To MemStat.dwAvailPhys) As Byte
    ReDim A(0 To 0) As Byte
    DoEvents
    'momory was filled and memory was released again
    
    Call GlobalMemoryStatus(MemStat)
    FreeMem = (Round(MemStat.dwAvailPhys / 1024 / 1024) - lngAvailPhys)
    
    If FreeMem > 0 Then FreeMem = FreeMem() + FreeMem
End Function

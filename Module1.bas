Attribute VB_Name = "Module1"
Public Declare Function GetLogicalDriveStrings Lib "kernel32" Alias "GetLogicalDriveStringsA" _
       (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long


Public Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" _
       (ByVal nDrive As String) As Long

 



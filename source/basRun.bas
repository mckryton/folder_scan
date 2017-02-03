Attribute VB_Name = "basRun"
'------------------------------------------------------------------------
' Description  : this module is about to execute the whole application
'------------------------------------------------------------------------

' Copyright 2016 Matthias Carell
'
'   Licensed under the Apache License, Version 2.0 (the "License");
'   you may not use this file except in compliance with the License.
'   You may obtain a copy of the License at
'
'       http://www.apache.org/licenses/LICENSE-2.0
'
'   Unless required by applicable law or agreed to in writing, software
'   distributed under the License is distributed on an "AS IS" BASIS,
'   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
'   See the License for the specific language governing permissions and
'   limitations under the License.

'Declarations

'Declare variables

'Options
Option Explicit
'-------------------------------------------------------------
' Description   : main routine for reading files from a directory
' Parameter     :
' Returnvalue   :
'-------------------------------------------------------------
Public Sub runFileScan()

    'local scan Folder
    Dim strFilesDir As String
    'file list read from scan folder
    Dim varFileList As Variant
    
    On Error GoTo error_handler
    'warn for unsupported versions of Excel
    #If MAC_OFFICE_VERSION >= 15 Then
        MsgBox "Excel 2016 MAC is not yet supported, please use Excel 2011 or a Windows version"
        Exit Sub
    #End If
   
    'select a folder containing feature descriptions, text files with a .feature extension
    strFilesDir = basFolderScan.getScanFolder()
    If strFilesDir = "" Then
        basSystem.log "choose dir dialog was canceled"
        Exit Sub
    End If
    
    'read file names from scan folder
    varFileList = basFolderScan.getFileNames(strFilesDir)
    
    'add new file names to the list
    basListUpdate.updateList strFilesDir, varFileList
    
    Application.StatusBar = False
    Exit Sub
    
error_handler:
    basSystem.log_error "basRun.runFileScan"
End Sub


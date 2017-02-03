Attribute VB_Name = "basFolderScan"
'------------------------------------------------------------------------
' Description  : this module is about reading files into the current workbook
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
Const cstrFolderSelectionTitle = "Bitte ein Verzeichnis auswählen"

'Declare variables

'Options
Option Explicit
'-------------------------------------------------------------
' Description   : ask user for the folder containing the files
' Parameter     :
' Returnvalue   : files directory as string
'-------------------------------------------------------------
Public Function getScanFolder() As String

    Dim strFilesDirInfo As Variant
    Dim strFilesDirDisk As String
    Dim strFilesDirPath As String
    Dim strFilesDirFullPath As String

    On Error GoTo error_handler
    'try to read folder from B4
    strFilesDirFullPath = basFolderScan.ScanFolderName.Text
        
    'show choose folder dialog if Path is empty
    If Trim(strFilesDirFullPath) = "" Then
        #If MAC_OFFICE_VERSION >= 15 Then
            
            MsgBox "Excel 2016 MAC is not yet supported, please use Excel 2011 or a Windows version"
            Exit Function
            
        #ElseIf Mac Then
    
            Dim strAppleScript As String
    
            'TODO: fix known bug -> umlauts like ä get converted by vba into a_
            strAppleScript = "try" & vbLf & _
                                    "tell application ""Finder""" & vbLf & _
                                        "set vPath to (choose folder with prompt ""choose folder"" default location (path to the desktop folder from user domain))" & vbLf & _
                                        "return {url of vPath, displayed name of disk of vPath}" & vbLf & _
                                    "end tell" & vbLf & _
                                "on error" & vbLf & _
                                    "return """"" & vbLf & _
                                "end try"
            strFilesDirInfo = Split(MacScript(strAppleScript), ", ")
            strFilesDirPath = basSystem.decomposeUrlPath(strFilesDirInfo(0))
            strFilesDirDisk = strFilesDirInfo(1)
            strFilesDirFullPath = strFilesDirDisk & strFilesDirPath
        #Else
            Dim dlgChooseFolder As FileDialog
    
            Set dlgChooseFolder = Application.FileDialog(msoFileDialogFolderPicker)
            With dlgChooseFolder
                .Title = cstrFolderSelectionTitle
                .AllowMultiSelect = False
                '.InitialFileName = strPath
                If .Show <> False Then
                    strFilesDirFullPath = .SelectedItems(1) & "\"
                End If
            End With
            Set dlgChooseFolder = Nothing
        #End If
    End If
    
    'save scan folder name
    basFolderScan.ScanFolderName.Value = strFilesDirFullPath
    basSystem.log ("scan folder is set to " & strFilesDirFullPath)
    getScanFolder = strFilesDirFullPath
    Exit Function
    
error_handler:
    Debug.Print Err.Description
    basSystem.log_error "basFolderScan.getScanFolder"
End Function

'-------------------------------------------------------------
' Description   : find all files
' Parameter     : pstrFilesDir    - directory containing all  files
' Returnvalue   : list of  file names as array
'-------------------------------------------------------------
Public Function getFileNames(pstrFilesDir As String) As Variant

    Dim colFileNames As Collection
    'Applescript code for Mac version
    Dim strScript As String
    Dim varFiles As Variant
    Dim varFilePath As Variant
    Dim lngFileIndex As Long
    
    On Error GoTo error_handler
    Application.StatusBar = "retrieve  file names"
    #If MAC_OFFICE_VERSION >= 15 Then
        
        MsgBox "Excel 2016 MAC is not yet supported, please use Excel 2011 or a Windows version"
        Exit Function
    #ElseIf Mac Then
        strScript = "set vFileNames to {}" & vbLf & _
                    "tell application ""Finder""" & vbLf & _
                        "set vsFolder to """ & pstrFilesDir & """ as alias" & vbLf & _
                        "set vFiles to (get files of vsFolder whose name ends with ""."")" & vbLf & _
                        "repeat with vFile in vFiles" & vbLf & _
                                "set end of vFileNames to get URL of vFile" & vbLf & _
                        "end repeat" & vbLf & _
                    "end tell" & vbLf & _
                    "return vFileNames"
        varFiles = MacScript(strScript)
        varFiles = Split(varFiles, ", ")
        For lngFileIndex = 0 To UBound(varFiles)
            varFilePath = Split(varFiles(lngFileIndex), "/")
            'remove the path from file name and translate decoded umlauts
            varFiles(lngFileIndex) = basSystem.decomposeUrlPath("file://" & _
                                                        varFilePath(UBound(varFilePath)))
        Next
    #Else
        Dim fsoFileSystem As Variant
        Dim fols As Variant
        Dim fil As Variant
        Dim strFiles As String
        
        strFiles = ""
        Set fsoFileSystem = CreateObject("Scripting.FileSystemObject")
        Set fols = fsoFileSystem.GetFolder(pstrFilesDir)
        For Each fil In fols.Files
            strFiles = strFiles & fil.Name & "//"
        Next
        strFiles = Left(strFiles, Len(strFiles) - 2)
        varFiles = Split(strFiles, "//")
    #End If
    
    basSystem.log "found " & UBound(varFiles) + 1 & " . files"
    getFileNames = varFiles
    Exit Function
    
error_handler:
    basSystem.log_error "basFolderScan.getFileNames"
End Function
'-------------------------------------------------------------
' Description   : try to figure out the range where the path for the scan folder was saved
' Returnvalue   : range object where the path of the scan folder is saved
'-------------------------------------------------------------
Public Property Get ScanFolderName() As Range

    Dim rngScanFolder As Range
    
    On Error GoTo name_not_found
    Set ScanFolderName = ThisWorkbook.Names("ScanFolder").RefersToRange
    Exit Property
    
name_not_found:
    Set ScanFolderName = wshFileList.Range("B4")
End Property



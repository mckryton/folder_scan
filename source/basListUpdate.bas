Attribute VB_Name = "basListUpdate"
'------------------------------------------------------------------------
' Description  : this module is about updating the current workbook with a list of filenames
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
' Description   : try to figure out where result list starts
' Returnvalue   : range object of the top left range of the result list
'-------------------------------------------------------------
Public Property Get ResultListStart() As Range

    Dim rngResultList As Range
    
    On Error GoTo name_not_found
    Set ResultListStart = ThisWorkbook.Names("StartList").RefersToRange
    Exit Property
    
name_not_found:
    Set ResultListStart = wshFileList.Range("A7")
End Property
'-------------------------------------------------------------
' Description   : try to figure out where filenames list starts
' Returnvalue   : range object of the top left range of the result list
'-------------------------------------------------------------
Public Property Get FilenamesListStart() As Range

    Dim rngResultList As Range
    
    On Error GoTo name_not_found
    Set FilenamesListStart = ThisWorkbook.Names("Filenames").RefersToRange
    Exit Property
    
name_not_found:
    Set FilenamesListStart = basListUpdate.ResultListStart.Offset(, 1)
End Property
'-------------------------------------------------------------
' Description   : try to finde the column containing the file names
' Returnvalue   : range object of the column containing the file names
'-------------------------------------------------------------
Public Property Get FilenameList() As Range

    Dim rngFilenamesStart As Range
    
    On Error GoTo error_handler
    Set rngFilenamesStart = basListUpdate.FilenamesListStart.Offset(1)
    'if the list is empty
    If rngFilenamesStart.CurrentRegion.Rows.Count = 2 Then
        Set FilenameList = Nothing
    Else
        Set FilenameList = wshFileList.Range(rngFilenamesStart, _
                                rngFilenamesStart.Offset(rngFilenamesStart.CurrentRegion.Rows.Count - 2))
    End If
    Exit Property

error_handler:
    basSystem.log_error "basListUpdate.Get FilenameList"
End Property
'-------------------------------------------------------------
' Description   : add the new file names to the list
' Parameter     : pstrFilesDir    - directory containing all  files
'                 pvarFileList    - array containing filenames
'-------------------------------------------------------------
Public Sub updateList(pstrFilesDir As String, pvarFileList As Variant)

    Dim rngOutput As Range
    Dim rngFileNames As Range
    Dim lngFileCount As Long
    Dim rngDuplicateName As Range
    
    On Error GoTo error_handler
    Set rngFileNames = basListUpdate.FilenameList
    If TypeName(rngFileNames) <> "Nothing" Then
        'set ouput range at the end of the filenames list
        Set rngOutput = rngFileNames.Cells(rngFileNames.Rows.Count, 1).Offset(1)
    Else
        'set output range below headline
        Set rngOutput = basListUpdate.FilenamesListStart.Offset(1)
    End If
    For lngFileCount = 0 To UBound(pvarFileList)
        If TypeName(rngFileNames) <> "Nothing" Then
            Set rngDuplicateName = rngFileNames.Find(pvarFileList(lngFileCount), _
                                        LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=True)
        End If
        'if the filename wasn't found in the list
        If TypeName(rngDuplicateName) = "Nothing" Then
            'save the current date
            rngOutput.Offset(, -1).Value = Now()
            'save the filename as link
            wshFileList.Hyperlinks.Add rngOutput, "file:///" & pstrFilesDir & "\" & pvarFileList(lngFileCount), _
                                            TextToDisplay:=pvarFileList(lngFileCount)
            'rngOutput.Value = pvarFileList(lngFileCount)
            'set output range to the next row
            Set rngOutput = rngOutput.Offset(1)
        End If
    Next
    Exit Sub
    
error_handler:
    basSystem.log_error "basListUpdate.updateList"
End Sub

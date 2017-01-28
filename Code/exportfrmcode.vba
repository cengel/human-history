Option Explicit
Option Compare Database
Function SaveToFile()                  'Save the code for all modules to files in currentDatabaseDir\Code
Dim Name, newName As String
Dim WasOpen As Boolean
Dim Last As Integer
Dim I As Integer
Dim TopDir As String, Path As String, FileName As String
Dim F As Long                          'File for saving code
Dim LineCount As Long                  'Line count of current module
I = InStrRev(CurrentDb.Name, "\")
TopDir = VBA.Left(CurrentDb.Name, I - 1)
Path = TopDir & "\" & "Code"           'Path where the files will be written
If (Dir(Path, vbDirectory) = "") Then
  MkDir Path                           'Ensure this exists
End If
Last = Application.CurrentProject.AllModules.Count - 1
For I = 0 To Last
  Name = CurrentProject.AllModules(I).Name
  WasOpen = True                       'Assume already open
  If Not CurrentProject.AllModules(I).IsLoaded Then
    WasOpen = False                    'Not currently open
    DoCmd.OpenModule Name              'So open it
  End If
  LineCount = Access.Modules(Name).CountOfLines
  FileName = Path & "\" & Name & ".vba"
  If InStr(FileName, ": ") Then
        newName = Replace(FileName, ": ", "_")
        FileName = newName
  ElseIf InStr(FileName, ":") Then
        newName = Replace(FileName, ":", "_")
        FileName = newName
  ElseIf InStr(FileName, "/") Then
        newName = Replace(FileName, "/", "_")
        FileName = newName
  End If
  If (Dir(FileName) <> "") Then
    Kill FileName                      'Delete previous version
  End If
  F = FreeFile
  Open FileName For Output Access Write As #F
  Print #F, Access.Modules(Name).Lines(1, LineCount)
  Close #F
  If Not WasOpen Then
    DoCmd.Close acModule, Name         'It wasn't open, so close it again
  End If
Next
Last = Application.CurrentProject.AllForms.Count - 1
For I = 0 To Last
  Name = CurrentProject.AllForms(I).Name
  WasOpen = True
  If Not CurrentProject.AllForms(I).IsLoaded Then
    WasOpen = False
    DoCmd.OpenForm Name, acDesign
  End If
  LineCount = Access.Forms(Name).Module.CountOfLines
  FileName = Path & "\" & Name & ".vba"
  If InStr(FileName, ": ") Then
        newName = Replace(FileName, ": ", "_")
        FileName = newName
  ElseIf InStr(FileName, ":") Then
        newName = Replace(FileName, ":", "_")
        FileName = newName
  ElseIf InStr(FileName, "/") Then
        newName = Replace(FileName, "/", "_")
        FileName = newName
  End If
  Debug.Print FileName
  If (Dir(FileName) <> "") Then
    Kill FileName
  End If
  F = FreeFile
  Open FileName For Output Access Write As #F
  Print #F, Access.Forms(Name).Module.Lines(1, LineCount)
  Close #F
  If Not WasOpen Then
    DoCmd.Close acForm, Name
  End If
Next
MsgBox "Created source files in " & Path
End Function

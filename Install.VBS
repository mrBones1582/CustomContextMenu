Option Explicit

Dim objFSO, objShell, strSource, strDest, strScriptPath
Dim objExcel, objAddin
Dim intResponse

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objShell = CreateObject("WScript.Shell")

intResponse = MsgBox("CustomContextMenu�A�h�C�����C���X�g�[�����ċX�����ł����H", vbYesNo, "�C���X�g�[���m�F")

If intResponse = vbYes Then
    strScriptPath = objFSO.GetParentFolderName(WScript.ScriptFullName)
    strSource = objFSO.BuildPath(strScriptPath, "CustomContextMenu.xlam")
    strDest = objShell.ExpandEnvironmentStrings("%AppData%") & "\Microsoft\AddIns\CustomContextMenu.xlam"

    ' Check if the source file exists
    If objFSO.FileExists(strSource) Then
        ' Copy the add-in to the destination folder
        objFSO.CopyFile strSource, strDest, True
        
        Set objShell = Nothing
        Set objFSO = Nothing
        
        ' Register the add-in in Excel
        Set objExcel = CreateObject("Excel.Application")
        objExcel.Workbooks.Add

        On Error Resume Next
        Set objAddin = objExcel.AddIns.Add(strDest, True)
        
        If Err.Number <> 0 Then
            WScript.Echo "�G���[���������܂���: " & Err.Number & ":" & Err.Description
            Err.Clear
        Else
            objAddin.Installed = True
        End If
        
        On Error GoTo 0

        objExcel.Quit

        WScript.Echo "Excel�A�h�C���uCustomContextMenu�v�̃C���X�g�[�������I�I"
    Else
        WScript.Echo "�G���[: CustomContextMenu.xlam �t�@�C����������܂���B"
    End If
End If

Set objExcel = Nothing
Set objAddin = Nothing
Set objFSO = Nothing
Set objShell = Nothing
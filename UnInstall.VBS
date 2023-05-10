Option Explicit

Dim objFSO, objShell, strAddinPath, strAddinName
Dim objExcel, objAddin, blnRemoved, i
Dim intResponse

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objShell = CreateObject("WScript.Shell")

intResponse = MsgBox("CustomContextMenu�A�h�C�����A���C���X�g�[�����ċX�����ł����H", vbYesNo, "�A���C���X�g�[���m�F")

If intResponse = vbYes Then
    strAddinName = "CustomContextMenu.xlam"
    strAddinPath = objShell.ExpandEnvironmentStrings("%AppData%") & "\Microsoft\AddIns\" & strAddinName

    ' Unregister the add-in in Excel
    Set objExcel = CreateObject("Excel.Application")

    For i = 1 To objExcel.AddIns.Count
        If objExcel.AddIns(i).Name = strAddinName Then
            Set objAddin = objExcel.AddIns(i)
            Exit For
        End If
    Next

    If Not objAddin Is Nothing Then
        objAddin.Installed = False
        blnRemoved = True
    Else
        WScript.Echo "�A�h�C����������܂���ł����B"
    End If

    objExcel.Quit

    ' Remove the add-in file from the destination folder
    If objFSO.FileExists(strAddinPath) Then
        If blnRemoved Then
            objFSO.DeleteFile strAddinPath
        Else
            WScript.Echo "�A�h�C��������ɉ�������Ȃ��������߁A�t�@�C�����폜���܂���ł����B"
        End If
    Else
        WScript.Echo "�G���[: " & strAddinName & " �t�@�C����������܂���B"
    End If

    Set objExcel = Nothing
    Set objAddin = Nothing
    Set objFSO = Nothing
    Set objShell = Nothing

    If blnRemoved Then
        WScript.Echo "Excel�A�h�C���uCustomContextMenu�v�̃A���C���X�g�[�������I�I"
    End If
End If
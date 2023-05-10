Option Explicit

Dim objFSO, objShell, strAddinPath, strAddinName
Dim objExcel, objAddin, blnRemoved, i
Dim intResponse

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objShell = CreateObject("WScript.Shell")

intResponse = MsgBox("CustomContextMenuアドインをアンインストールして宜しいですか？", vbYesNo, "アンインストール確認")

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
        WScript.Echo "アドインが見つかりませんでした。"
    End If

    objExcel.Quit

    ' Remove the add-in file from the destination folder
    If objFSO.FileExists(strAddinPath) Then
        If blnRemoved Then
            objFSO.DeleteFile strAddinPath
        Else
            WScript.Echo "アドインが正常に解除されなかったため、ファイルを削除しませんでした。"
        End If
    Else
        WScript.Echo "エラー: " & strAddinName & " ファイルが見つかりません。"
    End If

    Set objExcel = Nothing
    Set objAddin = Nothing
    Set objFSO = Nothing
    Set objShell = Nothing

    If blnRemoved Then
        WScript.Echo "Excelアドイン「CustomContextMenu」のアンインストール完了！！"
    End If
End If

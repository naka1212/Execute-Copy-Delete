'================================================================================
'   CLASS NAME  : プログラム削除処理
'   VERSION NO  : 1.0.0
'   CREATE DATE : 2010/04/22
'   UPDATE DATE : 
'   SOURCE CODE COPYRIGHT : S.Nakagawa
'================================================================================
Option Explicit On

Public Class frmDeleteModule

    Private Sub BtnNo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnNo.Click
        If MsgBox("終了しますか？", MsgBoxStyle.Question + MsgBoxStyle.YesNo, "確認") = MsgBoxResult.Yes Then
            Me.Close()
        End If
    End Sub

    Private Sub BtnYes_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnYes.Click
        Dim sErrorMsg As String = ""
        sErrorMsg = Do_DeleteModule()
        If sErrorMsg = "" Then
            MsgBox("完了しました", MsgBoxStyle.Information, "確認")
        ElseIf sErrorMsg = "stop" Then
        Else
            MsgBox("エラーがかかりました。ログを確認してください。", MsgBoxStyle.Information, "確認")
        End If
        Me.Close()
    End Sub

    '============================================================
    ' FUNCTION NAME Do_DeleteModule
    ' EXPLANATION   プログラムを削除する
    '============================================================
    Private Function Do_DeleteModule() As String
        Dim sErrorMsg As String = ""
        Dim sToDir As String = ""

        Try
            '(1)パスおよびパラメータ取得
            sErrorMsg = Pub_Get_IniData(sToDir)
            If sErrorMsg = "" Then
                If (sToDir = "") Then
                    'エラー->終わり
                    sErrorMsg = "INIファイルは読めたが、パラメータ取得に失敗"
                End If
            End If
            If sErrorMsg = "stop" Then
                '終わり
                Return "stop"
            End If

            If sErrorMsg = "" Then
                'デスクトップのショートカットを削除
                Dim sDesktopPath As String = System.Environment.GetFolderPath(System.Environment.SpecialFolder.Desktop)
                Dim sShortcutDsk As Object = sDesktopPath + "\MainMenu.45.lnk"
                System.IO.File.Delete(sShortcutDsk)

                'スタートアップのショートカットを削除
                Dim sStartupPath As String = System.Environment.GetFolderPath(System.Environment.SpecialFolder.Startup)
                Dim sShortcutStr As Object = sStartupPath + "\CopyModule.45.lnk"
                System.IO.File.Delete(sShortcutStr)

                'ローカルディレクトリの削除
                If System.IO.Directory.Exists(sToDir) Then
                    System.IO.Directory.Delete(sToDir, True)
                End If
            End If

        Catch ex As Exception
            sErrorMsg = Err.Number & " : " & ex.Message
        End Try

        If sErrorMsg <> "" Then
            'エラー->終わり
            Pub_WriteLogFile(sErrorMsg)
        End If

        Do_DeleteModule = sErrorMsg

    End Function

End Class
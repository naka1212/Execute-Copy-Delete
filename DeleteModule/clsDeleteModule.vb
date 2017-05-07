'================================================================================
'   OBJECT NAME : 商品マスタ画像サムネイル自動登録プログラム
'   VERSION NO  : 1.0.0
'   CREATE DATE : 2010/01/25
'   UPDATE DATE : 2012/05/27    2010変換
'   UPDATE DATE : 
'   SOURCE CODE COPYRIGHT : S.Nakagawa
'================================================================================
Option Explicit On

Public Class clsDeleteModule

    '============================================================
    ' FUNCTION NAME ShowDeleteForm
    ' EXPLANATION   画面を表示する
    '============================================================
    Public Sub ShowDeleteForm()
        Dim oForm As New frmDeleteModule

        Try
            oForm.ShowDialog()
            oForm.Dispose()

        Catch ex As Exception
            'エラーメッセージ表示
            MsgBox("画面の表示に失敗しました(frmDeleteModule)" & Chr(13) & ex.Message, _
                                 MsgBoxStyle.Critical, "システムエラー")
            Exit Sub
        End Try

    End Sub

End Class

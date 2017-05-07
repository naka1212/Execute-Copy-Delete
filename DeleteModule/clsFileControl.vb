'================================================================================
'   OBJECT NAME : ファイルコントロール共通クラス
'   VERSION NO  : 1.0.0
'   CREATE DATE : 2010/01/25
'   UPDATE DATE : 
'   SOURCE CODE COPYRIGHT : S.Nakagawa
'================================================================================
Option Explicit On

Public Class FileControl


    '============================================================
    ' FUNCTION NAME GetAllFiles
    '   INPUT   sFolderPath     : ファイル名を取得するフォルダのパス
    '   INPUT   sPattern        : 検索文字("*.jpg" など)
    '   INPUT   bSubFolderFlag  : True=サブフォルダからも取得する
    '   OUTPUT  sFilesArray     : 検索結果
    '   INPUT   bShowMsgFlag    : True=エラーメッセージを表示する
    '   RETURN  True or False
    ' EXPLANATION   指定のフォルダ内のファイル名を取得する
    '============================================================
    Public Function GetAllFiles(ByVal sFolderPath As String, ByVal sPattern As String, _
                                ByRef sFilesArray As ArrayList, _
                                Optional ByVal bSubFolderFlag As Boolean = False, _
                                Optional ByVal bShowMsgFlag As Boolean = False) As String
        Dim sErrorMsg As String = ""
        Dim bRet As Boolean

        sErrorMsg = ""
        GetAllFiles = ""

        If System.IO.Directory.Exists(sFolderPath) = False Then
            Exit Function
        End If

        Try
            '指定のフォルダ内のファイル名を取得する
            Dim sFiles As String() = System.IO.Directory.GetFiles(sFolderPath, sPattern)
            sFilesArray.AddRange(sFiles)
            bRet = True

            If bSubFolderFlag Then
                '指定のフォルダ内のサブフォルダも調べる
                Dim sSubFolder As String() = System.IO.Directory.GetDirectories(sFolderPath)
                Dim sSubFolderPath As String
                For Each sSubFolderPath In sSubFolder
                    '再帰呼び出し
                    sErrorMsg = GetAllFiles(sSubFolderPath, sPattern, sFilesArray)
                    If sErrorMsg <> "" Then
                        Exit For
                    End If
                Next sSubFolderPath
            End If

        Catch ex As Exception
            '異常
            sErrorMsg = Err.Number & " : " & ex.Message
            If bShowMsgFlag Then
                'システムエラーメッセージ表示
                MsgBox("ファイルの取得に失敗しました" & Chr(13) & ex.Message, _
                                     MsgBoxStyle.Critical, "システムエラー")
            End If
        End Try

        GetAllFiles = sErrorMsg

    End Function

    '============================================================
    ' FUNCTION NAME CreateTextfile
    '   INPUT   sOutputFileName : 接続ID
    '   INPUT   sText           : 出力文字列
    '   INPUT   iEncode         : エンコード(932->Shift-JIS,51932->EUC-JP)
    '   INPUT   bShowMsgFlag    : True=エラーメッセージを表示する
    '   RETURN  True of False
    ' EXPLANATION   テキストファイルを生成して文字列を出力する
    '============================================================
    Public Function CreateTextfile(ByVal sOutputFileName As String, ByVal sText As String, _
                                    Optional ByVal iEncode As Integer = 932, _
                                    Optional ByVal bShowMsgFlag As Boolean = False) As String
        Dim sErrorMsg As String = ""
        Dim swOutFile As System.IO.StreamWriter

        Try
            'データを書き出し
            '新規ファイルオープン(932->Shift-JIS,51932->EUC-JP)
            swOutFile = New System.IO.StreamWriter(sOutputFileName, False, System.Text.Encoding.GetEncoding(iEncode))
            '行データ出力
            swOutFile.Write(sText)
            'ファイルを更新
            swOutFile.Close()
            CreateTextfile = True

        Catch ex As System.Exception
            sErrorMsg = Err.Number & " : " & ex.Message

            If bShowMsgFlag Then
                '異常ログ出力
                MsgBox("データ書き出しに失敗しました" & Chr(13) & ex.Message, _
                        MsgBoxStyle.Critical, "システムエラー")
            End If
            CreateTextfile = False
        End Try

        CreateTextfile = sErrorMsg

    End Function

    '============================================================
    ' FUNCTION NAME GetTimeStampStr
    '   INPUT   TargetDatetime  : 文字列に変換する日時
    '   RETURN  変換後日時
    ' EXPLANATION   時刻を文字列に変換
    '============================================================
    Public Function GetTimeStampStr(ByVal TargetDatetime As DateTime) As String
        Dim sTimeStamp As String

        '時刻を文字列に変換
        sTimeStamp = CStr(Year(TargetDatetime) * 10000 + Month(TargetDatetime) * 100 + DateAndTime.Day(TargetDatetime))
        sTimeStamp = sTimeStamp & Microsoft.VisualBasic.Right("000000" & CStr(Hour(TargetDatetime) * 10000) + Minute(TargetDatetime) * 100 + Second(TargetDatetime), 6)

        GetTimeStampStr = sTimeStamp
    End Function

End Class

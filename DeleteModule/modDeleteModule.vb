'================================================================================
'   CLASS NAME  : プログラムコピーモジュール
'   VERSION NO  : 1.0.0
'   CREATE DATE : 2010/04/22
'   UPDATE DATE : 
'   SOURCE CODE COPYRIGHT : S.Nakagawa
'================================================================================
Option Explicit On

Module modDeleteModule

    '------------------------------------------------------------
    'グローバル変数
    '------------------------------------------------------------
    Public glb_strIniFileName As String = "DeleteModule.ini"

    '------------------------------------------------------------
    '関数宣言
    '------------------------------------------------------------
    Declare Function GetPrivateProfileString Lib "KERNEL32.DLL" Alias "GetPrivateProfileStringA" ( _
     ByVal lpAppName As String, _
     ByVal lpKeyName As String, _
     ByVal lpDefault As String, _
     ByVal lpReturnedString As System.Text.StringBuilder, _
     ByVal nSize As Integer, _
     ByVal lpFileName As String) As Integer


    '============================================================
    ' FUNCTION NAME Pub_Get_IniData
    ' OUTPUT        sFromDir    : コピー元ディレクトリ
    ' OUTPUT        sToDir      : コピー先ディレクトリ
    ' OUTPUT        sParam      : ROBOCOPYのパラメータ
    ' OUTPUT        bExistFlag  : False=INIファイルがなかった
    ' EXPLANATION   INIファイルからパラメータを取得
    '============================================================
    Public Function Pub_Get_IniData(ByRef sToDir As String) As String
        Dim sErrorMsg As String = ""
        Dim sIniPath As String
        Dim sCurPath As String
        Dim strBuilder As New System.Text.StringBuilder
        strBuilder.Capacity = 128

        Try
            'カレントディレクトリの取得
            sCurPath = System.IO.Directory.GetCurrentDirectory()
            sIniPath = sCurPath & "\" & glb_strIniFileName

            'iniファイル取得
            If Not System.IO.File.Exists(sIniPath) Then
                '強制的
                If MsgBox("Iniファイルが読み込めませんでした。 C:\Zetaを削除しますがよろしいですか？", _
                            MsgBoxStyle.Question + MsgBoxStyle.YesNo) = MsgBoxResult.No Then
                    MsgBox("中断します。システムは削除されておりません。", MsgBoxStyle.Information, "確認")
                    Return "stop"
                End If
                sToDir = "C:\Zeta"
            Else
                'iniファイルからパスおよびパラメータ取得
                GetPrivateProfileString("CopyDir", "DestDir", "", strBuilder, strBuilder.Capacity, sIniPath)
                sToDir = strBuilder.ToString
            End If

        Catch ex As Exception
            sErrorMsg = "INIファイルからパラメータを取得に失敗(" & _
                    Err.Number & " : " & ex.Message & ")"
        End Try
        Pub_Get_IniData = sErrorMsg
    End Function

    '============================================================
    ' FUNCTION NAME Pub_WriteLogFile
    ' EXPLANATION   エラーログをテキストファイルとして出力する
    '               ファイル名は、DeleteErrorLog_コンピュータ名_日時.txt
    '============================================================
    Public Sub Pub_WriteLogFile(ByVal sErrorMsg As String)
        Dim sLogFileName As String
        Dim sNowDatetime As String
        Dim sComName As String
        Dim oFileCon As New FileControl

        Try
            'ERRORログ出力
            sComName = My.Computer.Name
            sNowDatetime = oFileCon.GetTimeStampStr(Now)
            sLogFileName = "DeleteErrorLog_" & sComName & "_" & sNowDatetime & ".txt"
            oFileCon.CreateTextfile(sLogFileName, sErrorMsg)
        Catch ex As Exception
            MsgBox("エラーログの出力に失敗しました", MsgBoxStyle.Critical, "確認")
        End Try

    End Sub



    Sub Main()
        Try
            Dim oStart As New clsDeleteModule
            oStart.ShowDeleteForm()

        Catch ex As Exception
            'エラーメッセージ表示
            Call MsgBox(ex.Message, MsgBoxStyle.Critical, "エラー")
        End Try
    End Sub

End Module

'================================================================================
'   CLASS NAME  : �v���O�����폜����
'   VERSION NO  : 1.0.0
'   CREATE DATE : 2010/04/22
'   UPDATE DATE : 
'   SOURCE CODE COPYRIGHT : S.Nakagawa
'================================================================================
Option Explicit On

Public Class frmDeleteModule

    Private Sub BtnNo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnNo.Click
        If MsgBox("�I�����܂����H", MsgBoxStyle.Question + MsgBoxStyle.YesNo, "�m�F") = MsgBoxResult.Yes Then
            Me.Close()
        End If
    End Sub

    Private Sub BtnYes_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnYes.Click
        Dim sErrorMsg As String = ""
        sErrorMsg = Do_DeleteModule()
        If sErrorMsg = "" Then
            MsgBox("�������܂���", MsgBoxStyle.Information, "�m�F")
        ElseIf sErrorMsg = "stop" Then
        Else
            MsgBox("�G���[��������܂����B���O���m�F���Ă��������B", MsgBoxStyle.Information, "�m�F")
        End If
        Me.Close()
    End Sub

    '============================================================
    ' FUNCTION NAME Do_DeleteModule
    ' EXPLANATION   �v���O�������폜����
    '============================================================
    Private Function Do_DeleteModule() As String
        Dim sErrorMsg As String = ""
        Dim sToDir As String = ""

        Try
            '(1)�p�X����уp�����[�^�擾
            sErrorMsg = Pub_Get_IniData(sToDir)
            If sErrorMsg = "" Then
                If (sToDir = "") Then
                    '�G���[->�I���
                    sErrorMsg = "INI�t�@�C���͓ǂ߂����A�p�����[�^�擾�Ɏ��s"
                End If
            End If
            If sErrorMsg = "stop" Then
                '�I���
                Return "stop"
            End If

            If sErrorMsg = "" Then
                '�f�X�N�g�b�v�̃V���[�g�J�b�g���폜
                Dim sDesktopPath As String = System.Environment.GetFolderPath(System.Environment.SpecialFolder.Desktop)
                Dim sShortcutDsk As Object = sDesktopPath + "\MainMenu.45.lnk"
                System.IO.File.Delete(sShortcutDsk)

                '�X�^�[�g�A�b�v�̃V���[�g�J�b�g���폜
                Dim sStartupPath As String = System.Environment.GetFolderPath(System.Environment.SpecialFolder.Startup)
                Dim sShortcutStr As Object = sStartupPath + "\CopyModule.45.lnk"
                System.IO.File.Delete(sShortcutStr)

                '���[�J���f�B���N�g���̍폜
                If System.IO.Directory.Exists(sToDir) Then
                    System.IO.Directory.Delete(sToDir, True)
                End If
            End If

        Catch ex As Exception
            sErrorMsg = Err.Number & " : " & ex.Message
        End Try

        If sErrorMsg <> "" Then
            '�G���[->�I���
            Pub_WriteLogFile(sErrorMsg)
        End If

        Do_DeleteModule = sErrorMsg

    End Function

End Class
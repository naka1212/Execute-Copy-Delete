'================================================================================
'   CLASS NAME  : �v���O�����R�s�[���W���[��
'   VERSION NO  : 1.0.0
'   CREATE DATE : 2010/04/22
'   UPDATE DATE : 
'   SOURCE CODE COPYRIGHT : S.Nakagawa
'================================================================================
Option Explicit On

Module modDeleteModule

    '------------------------------------------------------------
    '�O���[�o���ϐ�
    '------------------------------------------------------------
    Public glb_strIniFileName As String = "DeleteModule.ini"

    '------------------------------------------------------------
    '�֐��錾
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
    ' OUTPUT        sFromDir    : �R�s�[���f�B���N�g��
    ' OUTPUT        sToDir      : �R�s�[��f�B���N�g��
    ' OUTPUT        sParam      : ROBOCOPY�̃p�����[�^
    ' OUTPUT        bExistFlag  : False=INI�t�@�C�����Ȃ�����
    ' EXPLANATION   INI�t�@�C������p�����[�^���擾
    '============================================================
    Public Function Pub_Get_IniData(ByRef sToDir As String) As String
        Dim sErrorMsg As String = ""
        Dim sIniPath As String
        Dim sCurPath As String
        Dim strBuilder As New System.Text.StringBuilder
        strBuilder.Capacity = 128

        Try
            '�J�����g�f�B���N�g���̎擾
            sCurPath = System.IO.Directory.GetCurrentDirectory()
            sIniPath = sCurPath & "\" & glb_strIniFileName

            'ini�t�@�C���擾
            If Not System.IO.File.Exists(sIniPath) Then
                '�����I
                If MsgBox("Ini�t�@�C�����ǂݍ��߂܂���ł����B C:\Zeta���폜���܂�����낵���ł����H", _
                            MsgBoxStyle.Question + MsgBoxStyle.YesNo) = MsgBoxResult.No Then
                    MsgBox("���f���܂��B�V�X�e���͍폜����Ă���܂���B", MsgBoxStyle.Information, "�m�F")
                    Return "stop"
                End If
                sToDir = "C:\Zeta"
            Else
                'ini�t�@�C������p�X����уp�����[�^�擾
                GetPrivateProfileString("CopyDir", "DestDir", "", strBuilder, strBuilder.Capacity, sIniPath)
                sToDir = strBuilder.ToString
            End If

        Catch ex As Exception
            sErrorMsg = "INI�t�@�C������p�����[�^���擾�Ɏ��s(" & _
                    Err.Number & " : " & ex.Message & ")"
        End Try
        Pub_Get_IniData = sErrorMsg
    End Function

    '============================================================
    ' FUNCTION NAME Pub_WriteLogFile
    ' EXPLANATION   �G���[���O���e�L�X�g�t�@�C���Ƃ��ďo�͂���
    '               �t�@�C�����́ADeleteErrorLog_�R���s���[�^��_����.txt
    '============================================================
    Public Sub Pub_WriteLogFile(ByVal sErrorMsg As String)
        Dim sLogFileName As String
        Dim sNowDatetime As String
        Dim sComName As String
        Dim oFileCon As New FileControl

        Try
            'ERROR���O�o��
            sComName = My.Computer.Name
            sNowDatetime = oFileCon.GetTimeStampStr(Now)
            sLogFileName = "DeleteErrorLog_" & sComName & "_" & sNowDatetime & ".txt"
            oFileCon.CreateTextfile(sLogFileName, sErrorMsg)
        Catch ex As Exception
            MsgBox("�G���[���O�̏o�͂Ɏ��s���܂���", MsgBoxStyle.Critical, "�m�F")
        End Try

    End Sub



    Sub Main()
        Try
            Dim oStart As New clsDeleteModule
            oStart.ShowDeleteForm()

        Catch ex As Exception
            '�G���[���b�Z�[�W�\��
            Call MsgBox(ex.Message, MsgBoxStyle.Critical, "�G���[")
        End Try
    End Sub

End Module

'================================================================================
'   OBJECT NAME : �t�@�C���R���g���[�����ʃN���X
'   VERSION NO  : 1.0.0
'   CREATE DATE : 2010/01/25
'   UPDATE DATE : 
'   SOURCE CODE COPYRIGHT : S.Nakagawa
'================================================================================
Option Explicit On

Public Class FileControl


    '============================================================
    ' FUNCTION NAME GetAllFiles
    '   INPUT   sFolderPath     : �t�@�C�������擾����t�H���_�̃p�X
    '   INPUT   sPattern        : ��������("*.jpg" �Ȃ�)
    '   INPUT   bSubFolderFlag  : True=�T�u�t�H���_������擾����
    '   OUTPUT  sFilesArray     : ��������
    '   INPUT   bShowMsgFlag    : True=�G���[���b�Z�[�W��\������
    '   RETURN  True or False
    ' EXPLANATION   �w��̃t�H���_���̃t�@�C�������擾����
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
            '�w��̃t�H���_���̃t�@�C�������擾����
            Dim sFiles As String() = System.IO.Directory.GetFiles(sFolderPath, sPattern)
            sFilesArray.AddRange(sFiles)
            bRet = True

            If bSubFolderFlag Then
                '�w��̃t�H���_���̃T�u�t�H���_�����ׂ�
                Dim sSubFolder As String() = System.IO.Directory.GetDirectories(sFolderPath)
                Dim sSubFolderPath As String
                For Each sSubFolderPath In sSubFolder
                    '�ċA�Ăяo��
                    sErrorMsg = GetAllFiles(sSubFolderPath, sPattern, sFilesArray)
                    If sErrorMsg <> "" Then
                        Exit For
                    End If
                Next sSubFolderPath
            End If

        Catch ex As Exception
            '�ُ�
            sErrorMsg = Err.Number & " : " & ex.Message
            If bShowMsgFlag Then
                '�V�X�e���G���[���b�Z�[�W�\��
                MsgBox("�t�@�C���̎擾�Ɏ��s���܂���" & Chr(13) & ex.Message, _
                                     MsgBoxStyle.Critical, "�V�X�e���G���[")
            End If
        End Try

        GetAllFiles = sErrorMsg

    End Function

    '============================================================
    ' FUNCTION NAME CreateTextfile
    '   INPUT   sOutputFileName : �ڑ�ID
    '   INPUT   sText           : �o�͕�����
    '   INPUT   iEncode         : �G���R�[�h(932->Shift-JIS,51932->EUC-JP)
    '   INPUT   bShowMsgFlag    : True=�G���[���b�Z�[�W��\������
    '   RETURN  True of False
    ' EXPLANATION   �e�L�X�g�t�@�C���𐶐����ĕ�������o�͂���
    '============================================================
    Public Function CreateTextfile(ByVal sOutputFileName As String, ByVal sText As String, _
                                    Optional ByVal iEncode As Integer = 932, _
                                    Optional ByVal bShowMsgFlag As Boolean = False) As String
        Dim sErrorMsg As String = ""
        Dim swOutFile As System.IO.StreamWriter

        Try
            '�f�[�^�������o��
            '�V�K�t�@�C���I�[�v��(932->Shift-JIS,51932->EUC-JP)
            swOutFile = New System.IO.StreamWriter(sOutputFileName, False, System.Text.Encoding.GetEncoding(iEncode))
            '�s�f�[�^�o��
            swOutFile.Write(sText)
            '�t�@�C�����X�V
            swOutFile.Close()
            CreateTextfile = True

        Catch ex As System.Exception
            sErrorMsg = Err.Number & " : " & ex.Message

            If bShowMsgFlag Then
                '�ُ탍�O�o��
                MsgBox("�f�[�^�����o���Ɏ��s���܂���" & Chr(13) & ex.Message, _
                        MsgBoxStyle.Critical, "�V�X�e���G���[")
            End If
            CreateTextfile = False
        End Try

        CreateTextfile = sErrorMsg

    End Function

    '============================================================
    ' FUNCTION NAME GetTimeStampStr
    '   INPUT   TargetDatetime  : ������ɕϊ��������
    '   RETURN  �ϊ������
    ' EXPLANATION   �����𕶎���ɕϊ�
    '============================================================
    Public Function GetTimeStampStr(ByVal TargetDatetime As DateTime) As String
        Dim sTimeStamp As String

        '�����𕶎���ɕϊ�
        sTimeStamp = CStr(Year(TargetDatetime) * 10000 + Month(TargetDatetime) * 100 + DateAndTime.Day(TargetDatetime))
        sTimeStamp = sTimeStamp & Microsoft.VisualBasic.Right("000000" & CStr(Hour(TargetDatetime) * 10000) + Minute(TargetDatetime) * 100 + Second(TargetDatetime), 6)

        GetTimeStampStr = sTimeStamp
    End Function

End Class

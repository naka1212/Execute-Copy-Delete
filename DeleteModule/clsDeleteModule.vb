'================================================================================
'   OBJECT NAME : ���i�}�X�^�摜�T���l�C�������o�^�v���O����
'   VERSION NO  : 1.0.0
'   CREATE DATE : 2010/01/25
'   UPDATE DATE : 2012/05/27    2010�ϊ�
'   UPDATE DATE : 
'   SOURCE CODE COPYRIGHT : S.Nakagawa
'================================================================================
Option Explicit On

Public Class clsDeleteModule

    '============================================================
    ' FUNCTION NAME ShowDeleteForm
    ' EXPLANATION   ��ʂ�\������
    '============================================================
    Public Sub ShowDeleteForm()
        Dim oForm As New frmDeleteModule

        Try
            oForm.ShowDialog()
            oForm.Dispose()

        Catch ex As Exception
            '�G���[���b�Z�[�W�\��
            MsgBox("��ʂ̕\���Ɏ��s���܂���(frmDeleteModule)" & Chr(13) & ex.Message, _
                                 MsgBoxStyle.Critical, "�V�X�e���G���[")
            Exit Sub
        End Try

    End Sub

End Class

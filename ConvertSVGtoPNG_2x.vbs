Option Explicit

'==========================================================================
' SVG to High-Quality PNG Converter (2x Size) - v2
'
' ���@�\
'   SVG�t�@�C�����A���掿��ۂ����܂܌��̃T�C�Y�̏c��2�{��PNG�t�@�C����
'   �ϊ����܂��B�w�i�͓��߂ɂȂ�܂��B
'
' ���ύX�_ (v2)
'   �摜���ڂ₯����ɑΉ����邽�߁A�ϊ������̕��@��ύX���܂����B
'   SVG��ǂݍ��ގ��_�̉𑜓x(DPI)��W����2�{(192 DPI)�Ɏw�肷�邱�ƂŁA
'   �x�N�^�[���̃f�B�e�[�����������ƂȂ������ׂ�PNG�𐶐����܂��B
'
' ���g����
'   1. ���̃X�N���v�g�t�@�C��(.vbs)���f�X�N�g�b�v�Ȃǂɕۑ����܂��B
'   2. �ϊ�������SVG�t�@�C�����A���̃X�N���v�g�̃A�C�R���̏��
'      �h���b�O���h���b�v���܂��B�i�����t�@�C�����j
'   3. SVG�t�@�C���Ɠ����t�H���_�ɁA�{�T�C�Y��PNG�t�@�C�����쐬����܂��B
'
' �����O����
'   ���̃X�N���v�g�����s����ɂ́A�uImageMagick�v���C���X�g�[������Ă���
'   �K�v������܂��B
'   �����T�C�g: https://imagemagick.org/
'
'==========================================================================

Dim objShell, objFSO, objArgs
Dim strSVGPath, strPNGPath, strCommand, fileCount

' --- �I�u�W�F�N�g�̐��� ---
Set objShell = CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objArgs = WScript.Arguments
fileCount = 0

' --- �h���b�O���h���b�v���ꂽ���`�F�b�N ---
If objArgs.Count = 0 Then
    MsgBox "SVG�t�@�C�������̃X�N���v�g�̃A�C�R���Ƀh���b�O���h���b�v���Ă��������B", vbInformation, "�g����"
    WScript.Quit
End If

' --- �e�t�@�C�������[�v���� ---
For Each strSVGPath In objArgs

    ' �t�@�C�������݂��A�g���q��SVG���m�F
    If objFSO.FileExists(strSVGPath) And LCase(objFSO.GetExtensionName(strSVGPath)) = "svg" Then

        ' �o�͂���PNG�t�@�C���̃p�X�𐶐� (��: icon.svg -> icon.png)
        strPNGPath = objFSO.BuildPath( _
            objFSO.GetParentFolderName(strSVGPath), _
            objFSO.GetBaseName(strSVGPath) & ".png" _
        )

        ' ImageMagick�����s���邽�߂̃R�}���h��g�ݗ��Ă�
        ' �y�d�v�z-density 192: SVG��ǂݍ��ލۂ̉𑜓x��W��(96dpi)��2�{�Ɏw�肵�܂��B
        '                      ����ɂ��A���X�^���C�Y(�r�b�g�}�b�v��)���_��2�{�̃s�N�Z���T�C�Y�ɂȂ�܂��B
        '                      ���̃I�v�V�����͓��̓t�@�C�����u�O�v�ɋL�q����K�v������܂��B
        ' -background none: �w�i�𓧉߂ɐݒ�
        strCommand = "magick convert -density 192 -background none """ & strSVGPath & """ """ & strPNGPath & """"

        ' �R�}���h�����s
        ' ��2���� 0: �R�}���h�v�����v�g�̃E�B���h�E���\���ɂ���
        ' ��3���� True: �R�}���h�̏�������������܂őҋ@����
        On Error Resume Next
        objShell.Run strCommand, 0, True
        If Err.Number <> 0 Then
            MsgBox "ImageMagick�̎��s�Ɏ��s���܂����B" & vbCrLf & _
                   "ImageMagick���������C���X�g�[������APATH���ʂ��Ă��邩�m�F���Ă��������B" & vbCrLf & vbCrLf & _
                   "���s�R�}���h: " & strCommand, vbCritical, "���s�G���["
            Err.Clear
            WScript.Quit
        End If
        On Error GoTo 0
        
        fileCount = fileCount + 1

    End If
Next

' --- �������b�Z�[�W ---
If fileCount > 0 Then
    MsgBox fileCount & "��SVG�t�@�C�������掿PNG�ɕϊ����܂����B", vbInformation, "��������"
Else
    MsgBox "�����Ώۂ�SVG�t�@�C����������܂���ł����B", vbExclamation, "��������"
End If

' --- �I�u�W�F�N�g�̉�� ---
Set objShell = Nothing
Set objFSO = Nothing
Set objArgs = Nothing
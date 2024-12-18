Attribute VB_Name = "EXportMacro"
Option Explicit

Sub ExportAll()
    Dim module As VBComponent      '// ���W���[��
    Dim moduleList As VBComponents     '// VBA�v���W�F�N�g�̑S���W���[��
    Dim extension                                   '// ���W���[���̊g���q
    Dim sPath                                       '// �����Ώۃu�b�N�̃p�X
    Dim sFilePath                                   '// �G�N�X�|�[�g�t�@�C���p�X
    Dim TargetBook                                  '// �����Ώۃu�b�N�I�u�W�F�N�g
    
    '// �u�b�N���J����Ă��Ȃ��ꍇ�͌l�p�}�N���u�b�N�ipersonal.xlsb�j��ΏۂƂ���
    If (Workbooks.Count = 1) Then
        Set TargetBook = ThisWorkbook
    '// �u�b�N���J����Ă���ꍇ�͕\�����Ă���u�b�N��ΏۂƂ���
    Else
        Set TargetBook = ActiveWorkbook
    End If
    
    sPath = TargetBook.path
    
    '// �����Ώۃu�b�N�̃��W���[���ꗗ���擾
    Set moduleList = TargetBook.VBProject.VBComponents
    
    '// VBA�v���W�F�N�g�Ɋ܂܂��S�Ẵ��W���[�������[�v
    For Each module In moduleList
        '// �N���X
        If (module.Type = vbext_ct_ClassModule) Then
            extension = "cls"
        '// �t�H�[��
        ElseIf (module.Type = vbext_ct_MSForm) Then
            '// .frx���ꏏ�ɃG�N�X�|�[�g�����
            extension = "frm"
        '// �W�����W���[��
        ElseIf (module.Type = vbext_ct_StdModule) Then
            extension = "bas"
        '// ���̑�
        Else
            '// �G�N�X�|�[�g�ΏۊO�̂��ߎ����[�v��
            GoTo CONTINUE
        End If
        
        '// �G�N�X�|�[�g���{
        sFilePath = sPath & "\" & module.Name & "." & extension
        Call module.Export(sFilePath)
        
        '// �o�͐�m�F�p���O�o��
        Debug.Print sFilePath
CONTINUE:
    Next
End Sub

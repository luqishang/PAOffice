@echo off

rem ==== �ȉ��A�쐬����c�k�k�ɉ����ĕύX���ĉ����� ============================================================
rem ���F�o�b�`�t�@�C���ҏW���A�g=�h�̑O��ɃX�y�[�X�����Ȃ��ŉ�����

rem -- �c�k�k�̃t�@�C����
SET DLL_RELEASE_FILENAME=PA.Office.dll

rem -- �A�Z���u���o�̓p�X�i��΃p�Xor���̃t�@�C������̑��΃p�X
SET DLL_RELEASE_PATH=.\bin\Release

rem -- Visual SourceSafe  �C���X�g�[���p�X�i�ʏ�͂��̂܂܂�OK�j
SET SSEXEC_PATH=C:\Program Files\Microsoft Visual SourceSafe

rem -- Visual SourceSafe  �f�[�^�x�[�X��
SET SSDIR=\\Webfilesv\pa_common

rem -- Visual SourceSafe  ���[�U��
SET SSUSER=morimoto

rem -- Visual SourceSafe  �p�X���[�h
SET SSPWD=morimoto

rem -- Visual SourceSafe  EXE�ADLL�i�[�ꏊ
SET SSPATH=$/VS2005/DLL

rem ==== ����ȉ��͏C�����Ȃ��ŉ����� ==========================================================================

set path=%path%;"%SSEXEC_PATH%"
ss Workfold "%SSPATH%" "%cd%"
ss Checkout "%SSPATH%/%DLL_RELEASE_FILENAME%"

copy "%DLL_RELEASE_PATH%\%DLL_RELEASE_FILENAME%" .\	/Y
ss Checkin "%SSPATH%/%DLL_RELEASE_FILENAME%" -C-
ss Workfold "%SSPATH%"

del "%DLL_RELEASE_FILENAME%" /F

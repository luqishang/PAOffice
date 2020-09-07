@echo off

rem ==== 以下、作成するＤＬＬに応じて変更して下さい ============================================================
rem 注：バッチファイル編集時、“=”の前後にスペースを入れないで下さい

rem -- ＤＬＬのファイル名
SET DLL_RELEASE_FILENAME=PA.Office.dll

rem -- アセンブリ出力パス（絶対パスorこのファイルからの相対パス
SET DLL_RELEASE_PATH=.\bin\Release

rem -- Visual SourceSafe  インストールパス（通常はこのままでOK）
SET SSEXEC_PATH=C:\Program Files\Microsoft Visual SourceSafe

rem -- Visual SourceSafe  データベース名
SET SSDIR=\\Webfilesv\pa_common

rem -- Visual SourceSafe  ユーザ名
SET SSUSER=morimoto

rem -- Visual SourceSafe  パスワード
SET SSPWD=morimoto

rem -- Visual SourceSafe  EXE、DLL格納場所
SET SSPATH=$/VS2005/DLL

rem ==== これ以下は修正しないで下さい ==========================================================================

set path=%path%;"%SSEXEC_PATH%"
ss Workfold "%SSPATH%" "%cd%"
ss Checkout "%SSPATH%/%DLL_RELEASE_FILENAME%"

copy "%DLL_RELEASE_PATH%\%DLL_RELEASE_FILENAME%" .\	/Y
ss Checkin "%SSPATH%/%DLL_RELEASE_FILENAME%" -C-
ss Workfold "%SSPATH%"

del "%DLL_RELEASE_FILENAME%" /F

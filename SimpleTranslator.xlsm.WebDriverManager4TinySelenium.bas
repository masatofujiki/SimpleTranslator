Attribute VB_Name = "WebDriverManager4TinySelenium"
Option Explicit

Enum BrowserName
    Chrome
    Edge
End Enum


'// �t�@�C���_�E�����[�h�p��Win32API
#If VBA7 Then
Private Declare PtrSafe Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" _
    (ByVal pCaller As Long, ByVal szURL As String, ByVal szFileName As String, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long
Private Declare PtrSafe Function DeleteUrlCacheEntry Lib "wininet" Alias "DeleteUrlCacheEntryA" (ByVal lpszUrlName As String) As Long
#Else
Private Declare Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" _
    (ByVal pCaller As Long, ByVal szURL As String, ByVal szFileName As String, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long
Private Declare Function DeleteUrlCacheEntry Lib "wininet" Alias "DeleteUrlCacheEntryA" (ByVal lpszUrlName As String) As Long
#End If


Private Property Get fso() 'As FileSystemObject
    Static obj As Object
    If obj Is Nothing Then Set obj = CreateObject("Scripting.FileSystemObject")
    Set fso = obj
End Property




'// �_�E�����[�h����WebDriver��zip�̃f�t�H���g�p�X
Public Property Get ZipPath(browser As BrowserName) As String
    Select Case browser
        Case BrowserName.Chrome
            ZipPath = "C:" & Environ("HOMEPATH") & "\Downloads\chromedriver_win32.zip"
        Case BrowserName.Edge
            Select Case Is64BitOS
                Case True: ZipPath = "C:" & Environ("HOMEPATH") & "\Downloads\edgedriver_win64.zip"
                Case Else: ZipPath = "C:" & Environ("HOMEPATH") & "\Downloads\edgedriver_win32.zip"
            End Select
    End Select
End Property


'// WebDriver�̎��s�t�@�C���̕ۑ��ꏊ
'// �f�t�H���g�̓h�L�������g�t�H���_
Public Property Get WebDriverPath(browser As BrowserName) As String
    Select Case browser
        Case BrowserName.Chrome: WebDriverPath = "C:" & Environ("HOMEPATH") & "\Documents\WebDriver\chromedriver.exe"
        Case BrowserName.Edge:   WebDriverPath = "C:" & Environ("HOMEPATH") & "\Documents\WebDriver\edgedriver.exe"
    End Select
End Property



'// �u���E�U�̃o�[�W���������W�X�g������ǂݎ��
'// �o�͗�@"94.0.992.31"
Public Property Get BrowserVersion(browser As BrowserName)
    Dim reg_version As String
    Select Case browser
        Case BrowserName.Chrome: reg_version = "HKEY_CURRENT_USER\SOFTWARE\Google\Chrome\BLBeacon\version"
        Case BrowserName.Edge:   reg_version = "HKEY_CURRENT_USER\SOFTWARE\Microsoft\Edge\BLBeacon\version"
    End Select
    
    On Error GoTo Catch
    BrowserVersion = CreateObject("WScript.Shell").RegRead(reg_version)
    Exit Property
    
Catch:
    Err.Raise 4000, , "�o�[�W������񂪎擾�ł��܂���ł����B�u���E�U���C���X�g�[������Ă��܂���B"
End Property
'// �o�͗�@"94"
Public Property Get BrowserVersionToMajor(browser As BrowserName)
    Dim vers
    vers = Split(BrowserVersion(browser), ".")
    BrowserVersionToMajor = vers(0)
End Property
'// �o�͗�@"94.0"
Public Property Get BrowserVersionToMinor(browser As BrowserName)
    Dim vers
    vers = Split(BrowserVersion(browser), ".")
    BrowserVersionToMinor = Join(Array(vers(0), vers(1)), ".")
End Property
'// �o�͗�@"94.0.992"
Public Property Get BrowserVersionToBuild(browser As BrowserName)
    Dim vers
    vers = Split(BrowserVersion(browser), ".")
    BrowserVersionToBuild = Join(Array(vers(0), vers(1), vers(2)), ".")
End Property


'// OS��64Bit���ǂ����𔻒肷��
Public Property Get Is64BitOS() As Boolean
    Dim arch As String
    arch = CreateObject("WScript.Shell").Environment("Process").Item("PROCESSOR_ARCHITECTURE") '�߂�l "AMD64","IA64","x86"�̂����ꂩ
    Is64BitOS = CBool(InStr(arch, "64"))
End Property




'// ��3�������ȗ�����΁A�_�E�����[�h�t�H���_�Ƀ_�E�����[�h�����
'//     DownloadWebDriver Edge, "94.0.992.31"
'//
'// ��2������BrowserVersion�v���p�e�B���g���΁A���݂̃u���E�U�ɓK������WebDriver���_�E�����[�h�ł���
'//     DownloadWebDriver Edge, BrowserVersion(Edge)
'//
'// ��3�����ɂăp�X���w�肷��ΔC�ӂ̏ꏊ�ɔC�ӂ̖��O�ŕۑ��ł���B
'//     DownloadWebDriver Edge, "94.0.992.31", "C:\Users\yamato\Desktop\edgedriver_94.zip"
Public Function DownloadWebDriver(browser As BrowserName, ver_webdriver As String, Optional path_save_to As String) As String
    Dim url As String
    Select Case browser
        Case BrowserName.Chrome
            url = Replace("https://chromedriver.storage.googleapis.com/{version}/chromedriver_win32.zip", "{version}", ver_webdriver)
        Case BrowserName.Edge
            Select Case Is64BitOS
                Case True: url = Replace("https://msedgedriver.azureedge.net/{version}/edgedriver_win64.zip", "{version}", ver_webdriver)
                Case Else: url = Replace("https://msedgedriver.azureedge.net/{version}/edgedriver_win32.zip", "{version}", ver_webdriver)
            End Select
    End Select
    
    If path_save_to = "" Then path_save_to = ZipPath(browser)   '�f�t�H��"C:Users\USERNAME\Downloads\~~~.zip"
    
    DeleteUrlCacheEntry url
    Dim ret As Long
    ret = URLDownloadToFile(0, url, path_save_to, 0, 0)
    If ret <> 0 Then Err.Raise 4001, , "�_�E�����[�h���s : " & url
    
    DownloadWebDriver = path_save_to
End Function



'// zip���璆�g�����o���Ďw��̏ꏊ�Ɏ��s�t�@�C����W�J����
'// chromedriver.exe(�f�t�H���g�̖��O)������Ƃ����chromedriver_94.exe�Ƃ��œW�J�ł���悤�A
'// ���̎��s�t�@�C�����㏑�����Ȃ��悤�Ɉ�xtemp�t�H���_������Ă�����s�t�@�C����ړI�̃p�X�ֈڂ�
'// ����zip��W�J����Ƃ��͓W�J��̃t�H���_���w�肷�邪�A
'// ���̊֐���WebDriver�̎��s�t�@�C���̃p�X�Ŏw�肷��̂Œ��ӁI(�W�J����̂�exe����)
'// �g�p��
'//     Extract "C:\Users\yamato\Downloads\chromedriver_win32.zip", "C:\Users\yamato\Downloads\chromedriver_94.exe"
Sub Extract(path_zip As String, path_save_to As String)
    Debug.Print "zip��W�J���܂�"
    
    Dim folder_temp
    folder_temp = fso.BuildPath(fso.GetParentFolderName(path_save_to), fso.GetTempName)
    fso.CreateFolder folder_temp
    Debug.Print "    �ꎞ�t�H���_ : " & folder_temp

    'PowerShell���g���ēW�J����ƃ}���E�F�A���肳�ꂽ�̂ŁC
    'MS�񐄏�����Shell.Application���g����zip���𓀂���
    On Error GoTo Catch
    Dim sh As Object
    Set sh = CreateObject("Shell.Application")
    'zip�t�@�C���ɓ����Ă���t�@�C�����w�肵���t�H���_�[�ɃR�s�[����
    '���������x()�ŕ]�����Ă���Namespace�ɓn���Ȃ��ƃG���[���o��
    sh.Namespace((folder_temp)).CopyHere sh.Namespace((path_zip)).Items

    Dim path_exe As String
    path_exe = fso.BuildPath(folder_temp, Dir(folder_temp & "\*.exe"))
    
    If fso.FileExists(path_save_to) Then fso.DeleteFile path_save_to
    fso.CopyFile path_exe, path_save_to, True

    fso.DeleteFolder folder_temp
    Debug.Print "    �W�J : " & path_save_to
    Debug.Print "WebDriver��z�u���܂���"
    Exit Sub
    
Catch:
    fso.DeleteFolder folder_temp
    Err.Raise 4002, , "Zip�̓W�J�Ɏ��s���܂����B�����F" & Err.Description
    Exit Sub
End Sub


'// ��{�I�ɂ̓u���E�U�̃o�[�W�����ƑS�������o�[�W������WebDriver���_�E�����[�h����΂����̂����A
'// ChromeDriver�̓r���h�ԍ��܂ł̃o�[�W�����𓊂���Ƃ������߃o�[�W�����������Ă����炵���H
'// �悭�킩��Ȃ����ǁA�T�C�g�ɂ��������Ă������B���@https://chromedriver.chromium.org/downloads/version-selection
'// �o�O�t�B�b�N�X�������[�X���邩��K��������v����Ƃ͌���Ȃ��Ƃ��B
Function RequestWebDriverVersion(ver_chrome)
    Dim http 'As XMLHTTP60
    Dim url As String
    
    Set http = CreateObject("MSXML2.ServerXMLHTTP")
    url = "https://chromedriver.storage.googleapis.com/LATEST_RELEASE_" & ver_chrome
    http.Open "GET", url, False
    http.send
    
    If http.statusText <> "OK" Then
        Err.Raise 4003, "�T�[�o�[�ւ̐ڑ��Ɏ��s���܂���"
        Exit Function
    End If

    RequestWebDriverVersion = http.responseText
End Function


'// �����Ńu���E�U�̃o�[�W�����Ɉ�v����WebDriver���_�E�����[�h���Azip��W�J�AWebDriver��exe�����̃t�H���_�ɔz�u����
'// �f�t�H���g�ł�C:\Users\USERNAME\Downloads�Ƀ_�E�����[�h���A
'// C:\Users\USERNAME\Documents\WebDriver\chromedriver.exe[edgedriver.exe]�ɔz�u����
'// ��2�������w�肷��ΔC�ӂ̃t�H���_�E�t�@�C�����ɂ��ăC���X�g�[���ł���
'// �w�肵���p�X�̓r���̃t�H���_�����݂��Ȃ��Ă��A�����ō쐬����
'// �g�p��
'//     InstallWebDriver Chrome, "C:\Users\USERNAME\Desktop\a\b\c\chromedriver_94.exe"
'//     ���f�X�N�g�b�v��\a\b\c\�t�H���_���쐬����Ă��̒��Ƀh���C�o���z�u�����
Sub InstallWebDriver(browser As BrowserName, Optional path_driver As String)
    Debug.Print "WebDriver���C���X�g�[�����܂�......"
    
    Dim ver_browser   As String
    Dim ver_webdriver As String
    ver_browser = BrowserVersion(browser)
    Select Case browser
        Case BrowserName.Chrome: ver_webdriver = RequestWebDriverVersion(BrowserVersionToBuild(browser))
        Case BrowserName.Edge:   ver_webdriver = ver_browser
    End Select
    
    Debug.Print "   �u���E�U          : Ver. " & ver_browser
    Debug.Print "   �K������WebDriver : Ver. " & ver_webdriver
    
    Dim path_zip As String
    path_zip = DownloadWebDriver(browser, ver_webdriver)
    
    Do Until fso.FileExists(ZipPath(browser))
        DoEvents
    Loop
    Debug.Print "   �_�E�����[�h����:" & path_zip
    
    If path_driver = "" Then path_driver = WebDriverPath(browser)
    
    If Not fso.FolderExists(fso.GetParentFolderName(path_driver)) Then
        Debug.Print "   WebDriver�̕ۑ���t�H���_���쐬���܂�"
        CreateFolderEx fso.GetParentFolderName(path_driver)
    End If
    
    Extract path_zip, path_driver
    
    Debug.Print "�C���X�g�[������"
End Sub



'// �p�X�Ɋ܂܂��S�Ẵt�H���_�̑��݊m�F�����ăt�H���_�����֐�
'// �g�p��
'// CreateFolderEx "C:\a\b\c\d\e\"
Sub CreateFolderEx(path_folder As String)
    '// �e�t�H���_���k��Ȃ��Ȃ�Ƃ���܂ōċA�ŒH��
    If fso.GetParentFolderName(path_folder) <> "" Then
        CreateFolderEx fso.GetParentFolderName(path_folder)
    End If
    '// �r���̑��݂��Ȃ��t�H���_���쐬���Ȃ���~��Ă���
    If Not fso.FolderExists(path_folder) Then
        fso.CreateFolder path_folder
    End If
End Sub



'// �����I�ɖ���WebDriver��u��������
'// ��x�u���E�U��Start�Ɏ��s�����WebDriver���I���ł����t�@�C���̒u���������ł��Ȃ������̂ŁA������������C���X�g�[������
'// TinySeleniumVBA�� "Driver.Chrome[Edge] path" �� "Driver.OpenBrowser"������ɒu��������΁A
'// �o�[�W�����A�b�v��V�KPC�ւ̔z�z���ɗ]�v�ȑ��삪����Ȃ�
Public Sub SafeOpen(Driver As WebDriver, browser As BrowserName)
    On Error GoTo Catch
    If IsOnline Then
        If fso.FileExists(WebDriverPath(browser)) Then
            Dim folder_temp As String
            folder_temp = fso.BuildPath(fso.GetParentFolderName(WebDriverPath(browser)), fso.GetTempName)
            fso.CreateFolder folder_temp
            fso.MoveFile WebDriverPath(browser), fso.BuildPath(folder_temp, "\webdriver.exe")
        End If
        InstallWebDriver browser
    End If
    
    Select Case browser
        Case BrowserName.Chrome: Driver.Chrome WebDriverPath(browser)
        Case BrowserName.Edge:   Driver.Edge WebDriverPath(browser)
    End Select
    Driver.OpenBrowser
    
    If fso.FolderExists(folder_temp) Then fso.DeleteFolder folder_temp
    Exit Sub
    
Catch:
    If fso.FileExists(fso.BuildPath(folder_temp, "\webdriver.exe")) Then
        fso.CopyFile fso.BuildPath(folder_temp, "\webdriver.exe"), WebDriverPath(browser), True
        fso.DeleteFolder folder_temp
    End If
    If IsOnline Then
        Err.Raise 4004, , "�u���E�U�̃I�[�v���Ɏ��s���܂����B�����F" & Err.Description
    Else
        Err.Raise 4005, , "�I�t���C���̂��ߍX�V�ł��܂���B�C���^�[�l�b�g�ɐڑ����Ă��������B"
    End If
End Sub



'// PC���I�����C�����ǂ����𔻒肷��
'// ���N�G�X�g�悪googl�Ȃ̂͏�Q�Ńy�[�W���J���Ȃ��Ƃ������Ƃ͏��Ȃ����Ȃ̂�
Public Function IsOnline() As Boolean
    Dim http
    Dim url As String
    On Error Resume Next
    Set http = CreateObject("MSXML2.ServerXMLHTTP")
    url = "https://www.google.co.jp/"
    http.Open "GET", url, False
    http.send
    
    Select Case http.statusText
        Case "OK": IsOnline = True
        Case Else: IsOnline = False
    End Select
End Function

Public Sub SafeOpenHeadless(Driver As WebDriver, browser As BrowserName)

    On Error GoTo Catch
    If IsOnline Then
        If fso.FileExists(WebDriverPath(browser)) Then
            Dim folder_temp As String
            folder_temp = fso.BuildPath(fso.GetParentFolderName(WebDriverPath(browser)), fso.GetTempName)
            fso.CreateFolder folder_temp
            fso.MoveFile WebDriverPath(browser), fso.BuildPath(folder_temp, "\webdriver.exe")
        End If
        InstallWebDriver browser
    End If
    
    Select Case browser
        Case BrowserName.Chrome: Driver.Chrome WebDriverPath(browser)
        Case BrowserName.Edge:   Driver.Edge WebDriverPath(browser)
    End Select
    Dim cap As Capabilities
    Set cap = Driver.CreateCapabilities()
    cap.AddArgument "--headless --disable-gpu"
    Driver.OpenBrowser cap
    
    If fso.FolderExists(folder_temp) Then fso.DeleteFolder folder_temp
    Exit Sub
    
Catch:
    If fso.FileExists(fso.BuildPath(folder_temp, "\webdriver.exe")) Then
        fso.CopyFile fso.BuildPath(folder_temp, "\webdriver.exe"), WebDriverPath(browser), True
        fso.DeleteFolder folder_temp
    End If
    If IsOnline Then
        Err.Raise 4004, , "�u���E�U�̃I�[�v���Ɏ��s���܂����B�����F" & Err.Description
    Else
        Err.Raise 4005, , "�I�t���C���̂��ߍX�V�ł��܂���B�C���^�[�l�b�g�ɐڑ����Ă��������B"
    End If
End Sub


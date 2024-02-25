Attribute VB_Name = "GetLocalOneDrivePath"
'https://gist.githubusercontent.com/guwidoe/038398b6be1b16c458365716a921814d/raw/f433330e310abeb32e88b21b070c759a75ee7eff/GetLocalOneDrivePath.bas.vb
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'                              IMPORTANT NOTE!
'If you are using this solution and you downloaded it before the 2nd of October
'2023, please update your code to the current version, as otherwise the function
'might stop working for OneDrive version 23.184.0903.0001 and newer.
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'
' Cross-platform VBA Function to get the local path of OneDrive/SharePoint
' synchronized Microsoft Office files (Works on Windows and on macOS)
'
' Author: Guido Witt-D�rring
' Created: 2022/07/01
' Updated: 2024/01/09
' License: MIT
'
' ----------------------------------------------------------------
' https://gist.github.com/guwidoe/038398b6be1b16c458365716a921814d
' https://stackoverflow.com/a/73577057/12287457
' ----------------------------------------------------------------
'
' Copyright (c) 2023 Guido Witt-D�rring
'
' MIT License:
' Permission is hereby granted, free of charge, to any person obtaining a copy
' of this software and associated documentation files (the "Software"), to
' deal in the Software without restriction, including without limitation the
' rights to use, copy, modify, merge, publish, distribute, sublicense, and/or
' sell copies of the Software, and to permit persons to whom the Software is
' furnished to do so, subject to the following conditions:
'
' The above copyright notice and this permission notice shall be included in
' all copies or substantial portions of the Software.
'
' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
' IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
' FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
' AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
' LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING
' FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS
' IN THE SOFTWARE.
'
'-------------------------------------------------------------------------------
' COMMENTS REGARDING THE IMPLEMENTATION:
' 1) Background and Alternative
'    This function was intended to be written as a single procedure without any
'    dependencies, for maximum portability between projects, as it implements a
'    functionality that is very commonly needed for many VBA applications
'    working inside OneDrive/SharePoint synchronized directories. I followed
'    this paradigm because it was not clear to me how complicated this simple
'    sounding endeavour would turn out to be.
'    Unfortunately, more and more complications arose, and little by little,
'    the procedure turned incredibly complex. I do not condone the coding
'    style applied here, and this is not how I usually write code.
'    Nevertheless, I'm not open to rewriting this code in a different style,
'    because a clean implementation of this algorithm already exists, as pointed
'    out in the following.
'
'    If you would like to understand the underlying algorithm of how the local
'    path can be found with only the Url-path as input, I recommend following
'    the much cleaner implementation by Cristian Buse:
'    https://github.com/cristianbuse/VBA-FileTools
'    We developed the algorithm together and wrote separate implementations
'    concurrently. His solution is contained inside a module-level library,
'    split into many procedures and using features like private types and API-
'    functions, that are not available when trying to create a single procedure
'    without dependencies like below. This makes his code more readable.
'
'    Both of our solutions are well tested and actively supported with bugfixes
'    and improvements, so both should be equally valid choices for use in your
'    project. The differences in performance/features are marginal and they can
'    often be used interchangeably. If you need more file-system interaction
'    functionality, use Cristians library, and if you only need GetLocalPath,
'    just copy this function to any module in your project and it will work.
'
' 2) How does this function work?
'    This function builds the URL to Local translation dictionary by extracting
'    the mount points and the corresponding OneDrive URL-roots from the OneDrive
'    settings files.
'
'    For example, for your personal OneDrive, such a local mount point could
'    look like this:
'     - C:\Users\Username\OneDrive
'
'    and the corresponding URL-root could look like this:
'     - https://d.docs.live.net/f9d8c1184686d493
'
'    This "dictionary" can then be used to "translate" a given OneDrive URL to a
'    local path by replacing the part that is equal to one of the elements of
'    the dictionary with the corresponding local mount point.
'    For example, this OneDrive URL:
'     - https://d.docs.live.net/f9d8c1184686d493/Folder/File.xlsm
'    will be correctly "translated" to
'     - C:\Users\Username\OneDrive\Folder\File.xlsm
'
'    Because all possible OneDrive URLs for the local machine can be translated
'    by the same dictionary, it is implemented as `Static` in this function.
'    This means it will only be written the first time the function is called,
'    all subsequent function calls will find the "dictionary" already
'    initialized leading to shorter run time.
'
'    In order to build the dictionary, the function reads files from...
'    On Windows:
'        - the "%LOCALAPPDATA%\Microsoft" directory
'    On Mac:
'        - the "~/Library/Containers/com.microsoft.OneDrive-mac/Data/" & _
'              "Library/Application Support" directory
'        - and/or the "~/Library/Application Support" directory
'    It reads the following files:
'      - \OneDrive\settings\Personal\ClientPolicy.ini
'      - \OneDrive\settings\Personal\????????????????.dat
'      - \OneDrive\settings\Personal\????????????????.ini
'      - \OneDrive\settings\Personal\global.ini
'      - \OneDrive\settings\Personal\GroupFolders.ini
'      - \OneDrive\settings\Personal\SyncEngineDatabase.db *if .dat unavailable
'      - \OneDrive\settings\Business#\????????-????-????-????-????????????.dat
'      - \OneDrive\settings\Business#\????????-????-????-????-????????????.ini
'      - \OneDrive\settings\Business#\ClientPolicy*.ini
'      - \OneDrive\settings\Business#\global.ini
'      - \OneDrive\settings\Business#\SyncEngineDatabase.db *if .dat unavailable
'      - \Office\CLP\* (just the filename)
'
'    Where:
'     - "*" ... 0 or more characters
'     - "?" ... one character [0-9, a-f]
'     - "#" ... one digit
'     - "\" ... path separator, (= "/" on MacOS)
'     - The "???..." filenames represent CIDs)
'
'    All of the `.ini` files can be read easily as they use UTF-16 encoding
'    (UTF-8 on Mac, which makes it more difficult already).
'    The `.dat` files are much more difficult to decipher, because they use a
'    proprietary binary format. Luckily, the information we need can be
'    extracted by looking for certain byte-patterns inside these files and
'    copying and converting the data at a certain offset from these
'    "signature" bytes.
'
'    The `.db` files are the most challenging of them all and will only be read
'    if the `.dat` files are not available.
'    (for OneDrive version 23.184.0903.0001 and newer)
'    They are SQLite files, which makes reading them with VBA in a reliable
'    cross-platform way particularly challenging.
'
'    For those who are interested in the exact algorithm behind how these files
'    can be used to find the local path for a given OneDrive URL, please refer
'    to the GitHub issues we used to discuss the progress on our solutions.
'    Those are the following:
'     - https://github.com/cristianbuse/VBA-FileTools/issues/1
'     - https://github.com/cristianbuse/VBA-FileTools/issues/2
'     - https://github.com/cristianbuse/VBA-FileTools/issues/17
'
'    The implementation for mac contains a bunch of peculiarities that are not
'    discussed in those issues. In order to understand exactly how the algorithm
'    works, as mentioned earlier, it's best to read Cristians implementation:
'     - https://github.com/cristianbuse/VBA-FileTools
'
'
' 3) How does this function NOT work?
'    There are a plethora of solutions for this problem circulating online.
'    A list of most of these solution can be found here:
'     - https://stackoverflow.com/a/73577057/12287457

'    In the stackoverflow post, detailed testing data is presented for all of
'    the mentioned solutions and it can be observed, that, unfortunately,
'    most of these alternatives are not very reliable.

'    Most are using one of two approaches:
'     1. they use the environment variables set by OneDrive:
'         - Environ(OneDrive)
'         - Environ(OneDriveCommercial)
'         - Environ(OneDriveConsumer)
'        and replace part of the URL with it. There are many problems with this
'        approach:
'         1. They are not being set by OneDrive on MacOS.
'         2. It is unclear exactly which part of the URL needs to be replaced.
'         3. Environment variables can be changed by the user.
'         4. Only there three exist. If more onedrive accounts are logged in,
'            they just overwrite the previous ones.
'        or,
'     2. they use the mount points OneDrive writes to the registry here:
'         - \HKEY_CURRENT_USER\Software\SyncEngines\Providers\OneDrive\
'        this also has several drawbacks:
'         1. The registry is not available on MacOS.
'         2. It's still unclear exactly what part of the URL should be replaced.
'         3. These registry keys can contain mistakes, like for example, when:
'             - Synchronizing a folder called "Personal" from someone else's
'               personal OneDrive
'             - Synchronizing a folder called "Business1" from someone else's
'               personal OneDrive and then relogging your own first Business
'               OneDrive account
'             - Relogging you personal OneDrive can change the "CID" property
'               from a folderID formatted cid (e.g. 3DEA8A9886F05935!125) to a
'               regular private cid (e.g. 3dea8a9886f05935) for synced folders
'               from other people's OneDrives
'
'    For these reasons, this solution uses a completely different approach to
'    solve this problem.
'-------------------------------------------------------------------------------

'-------------------------------------------------------------------------------
' COMMENTS REGARDING THE USAGE:
' This function can be used as a User Defined Function (UDF) from the worksheet.
' (More on that, see "USAGE EXAMPLES")
'
' This function offers three optional parameters to the user, however using
' these should only be necessary in extremely rare situations.
' The best rule regarding their usage: Don't use them.
'
' In the following these parameters will still be explained.
'
'1) returnAll
'   In some exceptional cases it is possible to map one OneDrive WebPath to
'   multiple different localPaths. This can happen when multiple Business
'   OneDrive accounts are logged in on one device, and multiple of these have
'   access to the same OneDrive folder and they both decide to synchronize it or
'   add it as link to their MySite library.
'   Calling the function with returnAll:=True will return all valid localPaths
'   for the given WebPath, separated by two forward slashes (//). This should be
'   used with caution, as the return value of the function alone is, should
'   multiple local paths exist for the input webPath, not a valid local path
'   anymore.
'   An example of how to obtain all of the local paths could look like this:
'   Dim localPath as String, localPaths() as String
'   localPath = GetLocalPath(webPath, True)
'   If Not localPath Like "http*" Then
'       localPaths = Split(localPath, "//")
'   End If
'
'2) preferredMountPointOwner
'   This parameter deals with the same problem as 'returnAll'
'   If the function gets called with returnAll:=False (default), and multiple
'   localPaths exist for the given WebPath, the function will just return any
'   one of them, as usually, it shouldn't make a difference, because the result
'   directories at both of these localPaths are mirrored versions of the same
'   webPath. Nevertheless, this option lets the user choose, which mountPoint
'   should be chosen if multiple localPaths are available. Each localPath is
'  'owned' by an OneDrive Account. If a WebPath is synchronized twice, this can
'   only happen by synchronizing it with two different accounts, because
'   OneDrive prevents you from synchronizing the same folder twice on a single
'   account. Therefore, each of the different localPaths for a given WebPath
'   has a unique 'owner'. preferredMountPointOwner lets the user select the
'   localPath by specifying the account the localPath should be owned by.
'   This is done by passing the Email address of the desired account as
'   preferredMountPointOwner.
'   For example, you have two different Business OneDrive accounts logged in,
'   foo.bar@business1.com and foo.bar@business2.com
'   Both synchronize the WebPath:
'   webPath = "https://business1.sharepoint.com/sites/TestLib/Documents/" & _
              "Test/Test/Test/test.xlsm"
'
'   The first one has added it as a link to his personal OneDrive, the local
'   path looks like this:
'   C:\Users\username\OneDrive - Business1\TestLinkParent\Test - TestLinkLib\...
'   ...Test\test.xlsm
'
'   The second one just synchronized it normally, the localPath looks like this:
'   C:\Users\username\Business1\TestLinkLib - Test\Test\test.xlsm
'
'   Calling GetLocalPath like this:
'   GetLocalPath(webPath,,, "foo.bar@business1.com") will return:
'   C:\Users\username\OneDrive - Business1\TestLinkParent\Test - TestLinkLib\...
'   ...Test\test.xlsm
'
'   Calling it like this:
'   GetLocalPath(webPath,,, "foo.bar@business2.com") will return:
'   C:\Users\username\Business1\TestLinkLib - Test\Test\test.xlsm
'
'   And calling it like this:
'   GetLocalPath(webPath,, True) will return:
'   C:\Users\username\OneDrive - Business1\TestLinkParent\Test - TestLinkLib\...
'   ...Test\test.xlsm//C:\Users\username\Business1\TestLinkLib - Test\Test\...
'   ...test.xlsm
'
'   Calling it normally like this:
'   GetLocalPath(webPath) will return any one of the two localPaths, so:
'   C:\Users\username\OneDrive - Business1\TestLinkParent\Test - TestLinkLib\...
'   ...Test\test.xlsm
'   OR
'   C:\Users\username\Business1\TestLinkLib - Test\Test\test.xlsm
'
'   If `preferredMountPointOwner` does not work on Mac, the following might
'   explain a reason and a workaround:
'
'   In order to correlate the users email address with the OneDrive account
'   CID, the function reads the filenames of the files located in the
'       - %LOCALAPPDATA%\Microsoft\Office\CLP\
'   directory.
'
'   On MacOS, the \Office\CLP\* exists for each Microsoft Office application
'   separately. Depending on whether the application was already used in
'   active syncing with OneDrive it may contain different/incomplete files.
'   In the code, the path of this directory is stored inside the variable
'   'clpPath'. On MacOS, the defined clpPath might not exist or not contain
'   all necessary files for some host applications, because Environ("HOME")
'   depends on the host app.
'   This is not a big problem as the function will still work, however in
'   this case, specifying a preferredMountPointOwner will do nothing.
'   To make sure this directory and the necessary files exist, a file must
'   have been actively synchronized with OneDrive by the application whose
'   "HOME" folder is returned by Environ("HOME") while being logged in
'   to that application with the account whose email is given as
'   preferredMountPointOwner, at some point in the past!
'
'   If you are usually working with Excel but are using this function in a
'   different app, you can instead use an alternative (Excels CLP folder) as
'   the clpPath as it will most likely contain all the necessary information
'   The alternative clpPath is commented out in the code, if you prefer to
'   use Excels CLP folder per default, just un-comment the respective line
'   in the code.
'
'3) rebuildCache
'   The function creates a "translation" dictionary from the OneDrive settings
'   files and then uses this dictionary to "translate" WebPaths to LocalPaths.
'   This dictionary is implemented as a static variable to the function doesn't
'   have to recreate it every time it is called. It is written on the first
'   function call and reused on all the subsequent calls, making them faster.
'   If the function is called with rebuildCache:=True, this dictionary will be
'   rewritten, even if it was already initialized.
'   Note that it is not necessary to use this parameter manually, even if a new
'   MountPoint was added to the OneDrive, or a new OneDrive account was logged
'   in since the last function call because the function will automatically
'   determine if any of those cases occurred, without sacrificing performance.
'-------------------------------------------------------------------------------
Option Explicit

''------------------------------------------------------------------------------
'' USAGE EXAMPLES:
'' Excel:
'Private Sub TestGetLocalPathExcel()
'    Debug.Print GetLocalPath(ThisWorkbook.FullName)
'    Debug.Print GetLocalPath(ThisWorkbook.path)
'End Sub
'
' Usage as User Defined Function (UDF):
' You might have to replace ; with , in the formulas depending on your settings.
' Add this formula to any cell, to get the local path of the workbook:
' =GetLocalPath(LEFT(CELL("filename";A1);FIND("[";CELL("filename";A1))-1))
'
' To get the local path including the filename (the FullName), use this formula:
' =GetLocalPath(LEFT(CELL("filename";A1);FIND("[";CELL("filename";A1))-1) &
' TEXTAFTER(TEXTBEFORE(CELL("filename";A1);"]");"["))
'
''Word:
'Private Sub TestGetLocalPathWord()
'    Debug.Print GetLocalPath(ThisDocument.FullName)
'    'Debug.Print GetLocalPath(ThisDocument.Path) '<- Do NOT use this.
'       'Document.Path returns an URL encoded url, e.g. " " -> "%20", therefore
'       'GetLocalPath doesn't work if there are encoded characters in the part
'       'that is supposed to be replaced. Document.FullName doesn't have this
'       'issue. Therefore, instead of GetLocalPath(ThisDocument.Path), use
'       'something like:
'    Dim docLocalPath As String: docLocalPath = ThisDocument.path
'    If docLocalPath Like "http*" Then
'        docLocalPath = GetLocalPath(Left(ThisDocument.FullName, _
'                                     InStrRev(ThisDocument.FullName, "/") - 1))
'    End If
'    Debug.Print docLocalPath
'End Sub
'
''PowerPoint:
'Private Sub TestGetLocalPathPowerPoint()
'    Debug.Print GetLocalPath(ActivePresentation.FullName)
'    Debug.Print GetLocalPath(ActivePresentation.path)
'End Sub
''------------------------------------------------------------------------------


'This Function will convert a OneDrive/SharePoint Url path, e.g. Url containing
'https://d.docs.live.net/; .sharepoint.com/sites; my.sharepoint.com/personal/...
'to the locally synchronized path on your current pc or mac, e.g. a path like
'C:\users\username\OneDrive\ on Windows; or /Users/username/OneDrive/ on MacOS,
'if you have the remote directory locally synchronized with the OneDrive app.
'If no local path can be found, the input value will be returned unmodified.
'Author: Guido Witt-D�rring
'Source: https://gist.github.com/guwidoe/038398b6be1b16c458365716a921814d
'        https://stackoverflow.com/a/73577057/12287457
Public Function GetLocalPath(ByVal path As String, _
                    Optional ByVal returnAll As Boolean = False, _
                    Optional ByVal preferredMountPointOwner As String = "", _
                    Optional ByVal rebuildCache As Boolean = False) _
                             As String
    #If Mac Then
        Const vbErrPermissionDenied            As Long = 70
        Const syncIDFileName As String = ".849C9593-D756-4E56-8D6E-42412F2A707B"
        Const isMac As Boolean = True
        Const ps As String = "/" 'Application.PathSeparator doesn't work
    #Else 'Windows               'in all host applications (e.g. Outlook), hence
        Const ps As String = "\" 'conditional compilation is preferred here.
        Const isMac As Boolean = False
    #End If
    Const methodName As String = "GetLocalPath"
    Const vbErrFileNotFound                As Long = 53
    Const vbErrOutOfMemory                 As Long = 7
    Const vbErrKeyAlreadyExists            As Long = 457
    Const vbErrInvalidFormatInResourceFile As Long = 325

    Static locToWebColl As Collection, lastCacheUpdate As Date

    If Not Left(path, 8) = "https://" Then GetLocalPath = path: Exit Function

    Dim webRoot As String, locRoot As String, s As String, vItem As Variant
    Dim pmpo As String: pmpo = LCase$(preferredMountPointOwner)
    If Not locToWebColl Is Nothing And Not rebuildCache Then
        Dim resColl As Collection: Set resColl = New Collection
        'If the locToWebColl is initialized, this logic will find the local path
        For Each vItem In locToWebColl
            locRoot = vItem(0): webRoot = vItem(1)
            If InStr(1, path, webRoot, vbTextCompare) = 1 Then _
                resColl.Add Key:=vItem(2), _
                   Item:=Replace(Replace(path, webRoot, locRoot, , 1), "/", ps)
        Next vItem
        If resColl.Count > 0 Then
            If returnAll Then
                For Each vItem In resColl: s = s & "//" & vItem: Next vItem
                GetLocalPath = Mid$(s, 3): Exit Function
            End If
            On Error Resume Next: GetLocalPath = resColl(pmpo): On Error GoTo 0
            If GetLocalPath <> "" Then Exit Function
            GetLocalPath = resColl(1): Exit Function
        End If
        'Local path was not found with cached mountpoints
        GetLocalPath = path 'No Exit Function here! Check if cache needs rebuild
    End If

    Dim settPaths As Collection: Set settPaths = New Collection
    Dim settPath As Variant, clpPath As String
    #If Mac Then 'The settings directories can be in different locations
        Dim cloudStoragePath As String, cloudStoragePathExists As Boolean
        s = Environ("HOME")
        clpPath = s & "/Library/Application Support/Microsoft/Office/CLP/"
        s = Left$(s, InStrRev(s, "/Library/Containers/", , vbBinaryCompare))
        settPaths.Add s & _
                      "Library/Containers/com.microsoft.OneDrive-mac/Data/" & _
                      "Library/Application Support/OneDrive/settings/"
        settPaths.Add s & "Library/Application Support/OneDrive/settings/"
        cloudStoragePath = s & "Library/CloudStorage/"

        'Excels CLP folder:
        'clpPath = Left$(s, InStrRev(s, "/Library/Containers", , 0)) & _
                  "Library/Containers/com.microsoft.Excel/Data/" & _
                  "Library/Application Support/Microsoft/Office/CLP/"
    #Else 'On Windows, the settings directories are always in this location:
        settPaths.Add Environ("LOCALAPPDATA") & "\Microsoft\OneDrive\settings\"
        clpPath = Environ("LOCALAPPDATA") & "\Microsoft\Office\CLP\"
    #End If

    Dim i As Long
    #If Mac Then 'Request access to all possible directories at once
        Dim arrDirs() As Variant: ReDim arrDirs(1 To settPaths.Count * 11 + 1)
        For Each settPath In settPaths
            For i = i + 1 To i + 9
                arrDirs(i) = settPath & "Business" & i Mod 11
            Next i
            arrDirs(i) = settPath: i = i + 1
            arrDirs(i) = settPath & "Personal"
        Next settPath
        arrDirs(i + 1) = cloudStoragePath
        Dim accessRequestInfoMsgShown As Boolean
        accessRequestInfoMsgShown = GetSetting("GetLocalPath", _
                        "AccessRequestInfoMsg", "Displayed", "False") = "True"
        If Not accessRequestInfoMsgShown Then MsgBox "The current " _
            & "VBA Project requires access to the OneDrive settings files to " _
            & "translate a OneDrive URL to the local path of the locally " & _
            "synchronized file/folder on your Mac. Because these files are " & _
            "located outside of Excels sandbox, file-access must be granted " _
            & "explicitly. Please approve the access requests following this " _
            & "message.", vbInformation
        If Not GrantAccessToMultipleFiles(arrDirs) Then _
            Err.Raise vbErrPermissionDenied, methodName
    #End If

    'Find all subdirectories in OneDrive settings folder:
    Dim oneDriveSettDirs As Collection: Set oneDriveSettDirs = New Collection
    For Each settPath In settPaths
        Dim dirName As String: dirName = Dir(settPath, vbDirectory)
        Do Until dirName = vbNullString
            If dirName = "Personal" Or dirName Like "Business#" Then _
                oneDriveSettDirs.Add Item:=settPath & dirName & ps
            dirName = Dir(, vbDirectory)
        Loop
    Next settPath

    If Not locToWebColl Is Nothing Or isMac Then
        Dim requiredFiles As Collection: Set requiredFiles = New Collection
        'Get collection of all required files
        Dim vDir As Variant
        For Each vDir In oneDriveSettDirs
            Dim cID As String: cID = IIf(vDir Like "*" & ps & "Personal" & ps, _
                                         "????????????*", _
                                         "????????-????-????-????-????????????")
            Dim fileName As String: fileName = Dir(vDir, vbNormal)
            Do Until fileName = vbNullString
                If fileName Like cID & ".ini" _
                Or fileName Like cID & ".dat" _
                Or fileName Like "ClientPolicy*.ini" _
                Or StrComp(fileName, "GroupFolders.ini", vbTextCompare) = 0 _
                Or StrComp(fileName, "global.ini", vbTextCompare) = 0 _
                Or StrComp(fileName, "SyncEngineDatabase.db", _
                           vbTextCompare) = 0 Then _
                    requiredFiles.Add Item:=vDir & fileName
                fileName = Dir
            Loop
        Next vDir
    End If

    'This part should ensure perfect accuracy despite the mount point cache
    'while sacrificing almost no performance at all by querying FileDateTimes.
    If Not locToWebColl Is Nothing And Not rebuildCache Then
        'Check if a settings file was modified since the last cache rebuild
        Dim vFile As Variant
        For Each vFile In requiredFiles
            If FileDateTime(vFile) > lastCacheUpdate Then _
                rebuildCache = True: Exit For 'full cache refresh is required!
        Next vFile
        If Not rebuildCache Then Exit Function
    End If

    'If execution reaches this point, the cache will be fully rebuilt...
    Dim fileNum As Long, syncID As String, b() As Byte, j As Long, k As Long
    'Variables for manual decoding of UTF-8, UTF-32 and ANSI
    Dim m As Long, ansi() As Byte, sAnsi As String
    Dim utf16() As Byte, sUtf16 As String, utf32() As Byte
    Dim utf8() As Byte, sUtf8 As String, numBytesOfCodePoint As Long
    Dim codepoint As Long, lowSurrogate As Long, highSurrogate As Long

    lastCacheUpdate = Now()
    #If Mac Then 'Prepare building syncIDtoSyncDir dictionary. This involves
        'reading the ".849C9593-D756-4E56-8D6E-42412F2A707B" files inside the
        'subdirs of "~/Library/CloudStorage/", list of files and access required
        Dim coll As Collection: Set coll = New Collection
        dirName = Dir(cloudStoragePath, vbDirectory)
        Do Until dirName = vbNullString
            If dirName Like "OneDrive*" Then
                cloudStoragePathExists = True
                vDir = cloudStoragePath & dirName & ps
                vFile = cloudStoragePath & dirName & ps & syncIDFileName
                coll.Add Item:=vDir
                requiredFiles.Add Item:=vDir 'For pooling file access requests
                requiredFiles.Add Item:=vFile
            End If
            dirName = Dir(, vbDirectory)
        Loop

        'Pool access request for these files and the OneDrive/settings files
        If locToWebColl Is Nothing Then
            Dim vFiles As Variant
            If requiredFiles.Count > 0 Then
                ReDim vFiles(1 To requiredFiles.Count)
               For i = 1 To UBound(vFiles): vFiles(i) = requiredFiles(i): Next i
                If Not GrantAccessToMultipleFiles(vFiles) Then _
                    Err.Raise vbErrPermissionDenied, methodName
            End If
        End If

        'More access might be required if some folders inside cloudStoragePath
        'don't contain the hidden file ".849C9593-D756-4E56-8D6E-42412F2A707B".
        'In that case, access to their first level subfolders is also required.
        If cloudStoragePathExists Then
            For i = coll.Count To 1 Step -1
                Dim fAttr As Long: fAttr = 0
                On Error Resume Next
                fAttr = GetAttr(coll(i) & syncIDFileName)
                Dim IsFile As Boolean: IsFile = False
                If Err.Number = 0 Then IsFile = Not CBool(fAttr And vbDirectory)
                On Error GoTo 0
                If Not IsFile Then 'hidden file does not exist
                'Dir(path, vbHidden) is unreliable and doesn't work on some Macs
                'If Dir(coll(i) & syncIDFileName, vbHidden) = vbNullString Then
                    dirName = Dir(coll(i), vbDirectory)
                    Do Until dirName = vbNullString
                        If Not dirName Like ".Trash*" And dirName <> "Icon" Then
                            coll.Add coll(i) & dirName & ps
                            coll.Add coll(i) & dirName & ps & syncIDFileName, _
                                     coll(i) & dirName & ps  '<- key for removal
                        End If
                        dirName = Dir(, vbDirectory)
                    Loop          'Remove the
                    coll.Remove i 'folder if it doesn't contain the hidden file.
                End If
            Next i
            If coll.Count > 0 Then
                ReDim arrDirs(1 To coll.Count)
                For i = 1 To coll.Count: arrDirs(i) = coll(i): Next i
                If Not GrantAccessToMultipleFiles(arrDirs) Then _
                    Err.Raise vbErrPermissionDenied, methodName
            End If
            'Remove all files from coll (not the folders!): Reminder:
            On Error Resume Next 'coll(coll(i)) = coll(i) & syncIDFileName
            For i = coll.Count To 1 Step -1
                coll.Remove coll(i)
            Next i
            On Error GoTo 0

            'Write syncIDtoSyncDir collection
            Dim syncIDtoSyncDir As Collection
            Set syncIDtoSyncDir = New Collection
            For Each vDir In coll
                fAttr = 0
                On Error Resume Next
                fAttr = GetAttr(vDir & syncIDFileName)
                IsFile = False
                If Err.Number = 0 Then IsFile = Not CBool(fAttr And vbDirectory)
                On Error GoTo 0
                If IsFile Then 'hidden file exists
                'Dir(path, vbHidden) is unreliable and doesn't work on some Macs
                'If Dir(vDir & syncIDFileName, vbHidden) <> vbNullString Then
                    fileNum = FreeFile(): s = "": vFile = vDir & syncIDFileName
                    'Somehow reading these files with "Open" doesn't always work
                    Dim readSucceeded As Boolean: readSucceeded = False
                    On Error GoTo ReadFailed
                    Open vFile For Binary Access Read As #fileNum
                        ReDim b(0 To LOF(fileNum)): Get fileNum, , b: s = b
                        readSucceeded = True
ReadFailed:             On Error GoTo -1
                    Close #fileNum: fileNum = 0
                    On Error GoTo 0
                    If readSucceeded Then
                        'Debug.Print "Used open statement to read file: " & _
                                    vDir & syncIDFileName
                        ansi = s 'If Open was used: Decode ANSI string manually:
                        If LenB(s) > 0 Then
                            ReDim utf16(0 To LenB(s) * 2 - 1): k = 0
                            For j = LBound(ansi) To UBound(ansi)
                                utf16(k) = ansi(j): k = k + 2
                            Next j
                            s = utf16
                        Else: s = vbNullString
                        End If
                    Else 'Reading the file with "Open" failed with an error. Try
                        'using AppleScript. Also avoids the manual transcoding.
                        'Somehow ApplScript fails too, sometimes. Seems whenever
                        '"Open" works, AppleScript fails and vice versa (?!?!)
                        vFile = MacScript("return path to startup disk as " & _
                                    "string") & Replace(Mid$(vFile, 2), ps, ":")
                        s = MacScript("return read file """ & _
                                      vFile & """ as string")
                       'Debug.Print "Used Apple Script to read file: " & vFile
                    End If
                    If InStr(1, s, """guid"" : """, vbBinaryCompare) Then
                        s = Split(s, """guid"" : """)(1)
                        syncID = Left$(s, InStr(1, s, """", 0) - 1)
                        syncIDtoSyncDir.Add Key:=syncID, _
                             Item:=VBA.Array(syncID, Left$(vDir, Len(vDir) - 1))
                    Else
                        Debug.Print "Warning, empty syncIDFile encountered!"
                    End If
                End If
            Next vDir
        End If
        'Now all access requests have succeeded
        If Not accessRequestInfoMsgShown Then SaveSetting _
            "GetLocalPath", "AccessRequestInfoMsg", "Displayed", "True"
    #End If

    'Declare some variables that will be used in the loop over OneDrive settings
    Dim line As Variant, parts() As String, n As Long, libNr As String
    Dim tag As String, mainMount As String, relPath As String, email As String
    Dim parentID As String, folderID As String, folderName As String
    Dim idPattern As String, folderType As String, keyExists As Boolean
    Dim siteID As String, libID As String, webID As String, lnkID As String
    Dim mainSyncID As String, syncFind As String, mainSyncFind As String
    'The following are "constants" and needed for reading the .dat files:
    Dim sig1 As String:       sig1 = ChrB$(2)
    Dim sig2 As String * 4:   MidB$(sig2, 1) = ChrB$(1)
    Dim vbNullByte As String: vbNullByte = ChrB$(0)
    #If Mac Then
        Const sig3 As String = vbNullChar & vbNullChar
    #Else 'Windows
        Const sig3 As String = vbNullChar
    #End If

    'Writing locToWebColl using .ini and .dat files in the OneDrive settings:
    'Here, a Scripting.Dictionary would be nice but it is not available on Mac!
    Dim lastAccountUpdates As Collection, lastAccountUpdate As Date
    Set lastAccountUpdates = New Collection
    Set locToWebColl = New Collection
    For Each vDir In oneDriveSettDirs 'One folder per logged in OD account
        dirName = Mid$(vDir, InStrRev(vDir, ps, Len(vDir) - 1, 0) + 1)
        dirName = Left$(dirName, Len(dirName) - 1)

        'Read global.ini to get cid
        If Dir(vDir & "global.ini", vbNormal) = "" Then GoTo NextFolder
        fileNum = FreeFile()
        Open vDir & "global.ini" For Binary Access Read As #fileNum
            ReDim b(0 To LOF(fileNum)): Get fileNum, , b
        Close #fileNum: fileNum = 0
        #If Mac Then 'On Mac, the OneDrive settings files use UTF-8 encoding
            sUtf8 = b: GoSub DecodeUTF8
            b = sUtf16
        #End If
        For Each line In Split(b, vbNewLine)
            If line Like "cid = *" Then cID = Mid$(line, 7): Exit For
        Next line

        If cID = vbNullString Then GoTo NextFolder
        If (Dir(vDir & cID & ".ini") = vbNullString Or _
           (Dir(vDir & "SyncEngineDatabase.db") = vbNullString And _
            Dir(vDir & cID & ".dat") = vbNullString)) Then GoTo NextFolder
        If dirName Like "Business#" Then
            idPattern = Replace(Space$(32), " ", "[a-f0-9]") & "*"
        ElseIf dirName = "Personal" Then
            idPattern = Replace(Space$(12), " ", "[A-F0-9]") & "*!###*"
        End If
        'Alternatively maybe a general pattern like this performs better:
        'idPattern = Replace(Space$(12), " ", "[a-fA-F0-9]") & "*"

        'Get email for business accounts
        '(only necessary to let user choose preferredMountPointOwner)
        fileName = Dir(clpPath, vbNormal)
        Do Until fileName = vbNullString
            i = InStrRev(fileName, cID, , vbTextCompare)
            If i > 1 And cID <> vbNullString Then _
                email = LCase$(Left$(fileName, i - 2)): Exit Do
            fileName = Dir
        Loop

        #If Mac Then
            On Error Resume Next
            lastAccountUpdate = lastAccountUpdates(dirName)
            keyExists = (Err.Number = 0)
            On Error GoTo 0
            If keyExists Then
                If FileDateTime(vDir & cID & ".ini") < lastAccountUpdate Then
                    GoTo NextFolder
                Else
                    For i = locToWebColl.Count To 1 Step -1
                        If locToWebColl(i)(5) = dirName Then
                            locToWebColl.Remove i
                        End If
                    Next i
                    lastAccountUpdates.Remove dirName
                    lastAccountUpdates.Add Key:=dirName, _
                                         Item:=FileDateTime(vDir & cID & ".ini")
                End If
            Else
                lastAccountUpdates.Add Key:=dirName, _
                                      Item:=FileDateTime(vDir & cID & ".ini")
            End If
        #End If

        'Read all the ClientPloicy*.ini files:
        Dim cliPolColl As Collection: Set cliPolColl = New Collection
        fileName = Dir(vDir, vbNormal)
        Do Until fileName = vbNullString
            If fileName Like "ClientPolicy*.ini" Then
                fileNum = FreeFile()
                Open vDir & fileName For Binary Access Read As #fileNum
                    ReDim b(0 To LOF(fileNum)): Get fileNum, , b
                Close #fileNum: fileNum = 0
                #If Mac Then 'On Mac, OneDrive settings files use UTF-8 encoding
                    sUtf8 = b: GoSub DecodeUTF8
                    b = sUtf16
                #End If
                cliPolColl.Add Key:=fileName, Item:=New Collection
                For Each line In Split(b, vbNewLine)
                    If InStr(1, line, " = ", vbBinaryCompare) Then
                        tag = Left$(line, InStr(1, line, " = ", 0) - 1)
                        s = Mid$(line, InStr(1, line, " = ", 0) + 3)
                        Select Case tag
                        Case "DavUrlNamespace"
                            cliPolColl(fileName).Add Key:=tag, Item:=s
                        Case "SiteID", "IrmLibraryId", "WebID" 'Only used for
                            s = Replace(LCase$(s), "-", "") 'backup method later
                            If Len(s) > 3 Then s = Mid$(s, 2, Len(s) - 2)
                            cliPolColl(fileName).Add Key:=tag, Item:=s
                        End Select
                    End If
                Next line
            End If
            fileName = Dir
        Loop

        'If cid.dat file doesn't exist, skip this part:
        Dim odFolders As Collection: Set odFolders = Nothing
        If Dir(vDir & cID & ".dat") = vbNullString Then GoTo Continue

        'Read cid.dat file if it exists:
        Const chunkOverlap          As Long = 1000
        Const maxDirName            As Long = 255
        Dim buffSize As Long: buffSize = -1 'Buffer uninitialized
Try:    On Error GoTo Catch
        Set odFolders = New Collection
        Dim lastChunkEndPos As Long: lastChunkEndPos = 1
        Dim lastFileUpdate As Date:  lastFileUpdate = FileDateTime(vDir & _
                                                                   cID & ".dat")
        i = 0 'i = current reading pos.
        Do
            'Ensure file is not changed while reading it
            If FileDateTime(vDir & cID & ".dat") > lastFileUpdate Then GoTo Try
            fileNum = FreeFile
            Open vDir & cID & ".dat" For Binary Access Read As #fileNum
                Dim lenDatFile As Long: lenDatFile = LOF(fileNum)
                If buffSize = -1 Then buffSize = lenDatFile 'Initialize buffer
                'Overallocate a bit so read chunks overlap to recognize all dirs
                ReDim b(0 To buffSize + chunkOverlap)
                Get fileNum, lastChunkEndPos, b: s = b
                Dim size As Long: size = LenB(s)
            Close #fileNum: fileNum = 0
            lastChunkEndPos = lastChunkEndPos + buffSize

            For vItem = 16 To 8 Step -8
                i = InStrB(vItem + 1, s, sig2, 0) 'Sarch pattern in cid.dat
                Do While i > vItem And i < size - 168 'and confirm with another
                    If StrComp(MidB$(s, i - vItem, 1), sig1, 0) = 0 Then 'one
                        i = i + 8: n = InStrB(i, s, vbNullByte, 0) - i
                        If n < 0 Then n = 0           'i:Start pos, n: Length
                        If n > 39 Then n = 39
                        #If Mac Then 'StrConv doesn't work reliably on Mac ->
                            sAnsi = MidB$(s, i, n) 'Decode ANSI string manually:
                            GoSub DecodeANSI: folderID = sUtf16
                        #Else 'Windows
                            folderID = StrConv(MidB$(s, i, n), vbUnicode)
                        #End If
                        i = i + 39: n = InStrB(i, s, vbNullByte, 0) - i
                        If n < 0 Then n = 0
                        If n > 39 Then n = 39
                        #If Mac Then 'StrConv doesn't work reliably on Mac ->
                            sAnsi = MidB$(s, i, n) 'Decode ANSI string manually:
                            GoSub DecodeANSI: parentID = sUtf16
                        #Else 'Windows
                            parentID = StrConv(MidB$(s, i, n), vbUnicode)
                        #End If
                        i = i + 121
                       n = InStr(-Int(-(i - 1) / 2) + 1, s, sig3, 0) * 2 - i - 1
                        If n > maxDirName * 2 Then n = maxDirName * 2
                        If n < 0 Then n = 0
                        If folderID Like idPattern _
                        And parentID Like idPattern Then
                            #If Mac Then 'Encoding of folder names is UTF-32-LE
                                Do While n Mod 4 > 0
                                    If n > maxDirName * 4 Then Exit Do
                                    n = InStr(-Int(-(i + n) / 2) + 1, s, sig3, _
                                              0) * 2 - i - 1
                                Loop
                                If n > maxDirName * 4 Then n = maxDirName * 4
                                utf32 = MidB$(s, i, n)
                                'UTF-32 can only be converted manually to UTF-16
                                ReDim utf16(LBound(utf32) To UBound(utf32))
                                j = LBound(utf32): k = LBound(utf32)
                                Do While j < UBound(utf32)
                                    If utf32(j + 2) + utf32(j + 3) = 0 Then
                                        utf16(k) = utf32(j)
                                        utf16(k + 1) = utf32(j + 1)
                                        k = k + 2
                                    Else
                                        If utf32(j + 3) <> 0 Then Err.Raise _
                                            vbErrInvalidFormatInResourceFile, _
                                            methodName
                                        codepoint = utf32(j + 2) * &H10000 + _
                                                    utf32(j + 1) * &H100& + _
                                                    utf32(j)
                                        m = codepoint - &H10000
                                        highSurrogate = &HD800& Or (m \ &H400&)
                                        lowSurrogate = &HDC00& Or (m And &H3FF)
                                        utf16(k) = highSurrogate And &HFF&
                                        utf16(k + 1) = highSurrogate \ &H100&
                                        utf16(k + 2) = lowSurrogate And &HFF&
                                        utf16(k + 3) = lowSurrogate \ &H100&
                                        k = k + 4
                                    End If
                                    j = j + 4
                                Loop
                                If k > LBound(utf16) Then
                                    ReDim Preserve utf16(LBound(utf16) To k - 1)
                                    folderName = utf16
                                Else: folderName = vbNullString
                                End If
                            #Else 'On Windows encoding is UTF-16-LE
                                folderName = MidB$(s, i, n)
                            #End If
                            'VBA.Array() instead of just Array() is used in this
                            'function because it ignores Option Base 1
                            odFolders.Add VBA.Array(parentID, folderName), _
                                          folderID
                        End If
                    End If
                    i = InStrB(i + 1, s, sig2, 0) 'Find next sig2 in cid.dat
                Loop
                If odFolders.Count > 0 Then Exit For
            Next vItem
        Loop Until lastChunkEndPos >= lenDatFile _
                Or buffSize >= lenDatFile
        GoTo Continue
Catch:
        Select Case Err.Number
        Case vbErrKeyAlreadyExists
            'This can happen at chunk boundries, folder might get added twice:
            odFolders.Remove folderID 'Make sure the folder gets added new again
            Resume 'to avoid folderNames truncated by chunk ends
        Case Is <> vbErrOutOfMemory: Err.Raise Err, methodName
        End Select
        If buffSize > &HFFFFF Then buffSize = buffSize / 2: Resume Try
        Err.Raise Err, methodName 'Raise error if less than 1 MB RAM available
Continue:
        On Error GoTo 0
        'If .dat file didn't exist, read db file, otherwise skip this part
        If Not odFolders Is Nothing Then GoTo SkipDbFile
        'The following code for reading the .db file is an adaptation of the
        'original code by Cristian Buse, see procedure 'GetODDirsFromDB' in the
        'repository: https://github.com/cristianbuse/VBA-FileTools
        fileNum = FreeFile()
        Open vDir & "SyncEngineDatabase.db" For Binary Access Read As #fileNum
        size = LOF(fileNum)
        If size = 0 Then GoTo CloseFile
        '                             __    ____
        'Signature bytes: 0b0b0b0b0b0b080b0b08080b0b0b0b where b>=0, b <= 9
        Dim sig88 As String: sig88 = ChrW$(&H808)
        Const sig8 As Long = 8
        Const sig8Offset As Long = -3
        Const maxSigByte As Byte = 9
        Const sig88ToDataOffset As Long = 6 'Data comes after the signature
        Const headBytes6 As Long = &H16
        Const headBytes5 As Long = &H15
        Const headBytes6Offset As Long = -16 'Header comes before the signature
        Const headBytes5Offset As Long = -15
        Const chunkSize As Long = &H100000 '1MB

        Dim lastRecord As Long, bytes As Long, nameSize As Long
        Dim idSize(1 To 4) As Byte
        Dim lastFolderID As String, lastParentID As String
        Dim lastNameStart As Long
        Dim lastNameSize As Long
        Dim currDataEnd As Long, lastDataEnd As Long
        Dim headByte As Byte, lastHeadByte As Byte
        Dim has5HeadBytes As Boolean

        lastFileUpdate = 0
        ReDim b(1 To chunkSize)
        Do
            i = 0
           If FileDateTime(vDir & "SyncEngineDatabase.db") > lastFileUpdate Then
                Set odFolders = New Collection
                Dim heads As Collection: Set heads = New Collection

                lastFileUpdate = FileDateTime(vDir & "SyncEngineDatabase.db")
                lastRecord = 1
                lastFolderID = vbNullString
            End If
            If LenB(lastFolderID) > 0 Then
                folderName = MidB$(s, lastNameStart, lastNameSize)
            End If
            Get fileNum, lastRecord, b
            s = b
            i = InStrB(1 - headBytes6Offset, s, sig88, vbBinaryCompare)
            lastDataEnd = 0
            Do While i > 0
                If i + headBytes6Offset - 2 > lastDataEnd _
                And LenB(lastFolderID) > 0 Then
                    If lastDataEnd > 0 Then
                        folderName = MidB$(s, lastNameStart, lastNameSize)
                    End If
                    sUtf8 = folderName: GoSub DecodeUTF8
                    folderName = sUtf16
                    On Error Resume Next
                    odFolders.Add VBA.Array(lastParentID, folderName), _
                                            lastFolderID
                    If Err.Number <> 0 Then
                        If heads(lastFolderID) < lastHeadByte Then
                            If odFolders(lastFolderID)(1) <> folderName _
                            Or odFolders(lastFolderID)(0) <> lastParentID Then
                                odFolders.Remove lastFolderID
                                heads.Remove lastFolderID
                                odFolders.Add VBA.Array(lastParentID, _
                                                        folderName), _
                                              lastFolderID
                            End If
                        End If
                    End If
                    heads.Add lastHeadByte, lastFolderID
                    On Error GoTo 0
                    lastFolderID = vbNullString
                End If

                If b(i + sig8Offset) <> sig8 Then GoTo NextSig
                has5HeadBytes = True
                If b(i + headBytes5Offset) = headBytes5 Then
                    j = i + headBytes5Offset
                ElseIf b(i + headBytes6Offset) = headBytes6 Then
                    j = i + headBytes6Offset
                    has5HeadBytes = False 'Has 6 bytes header
                ElseIf b(i + headBytes5Offset) <= maxSigByte Then
                    j = i + headBytes5Offset
                Else
                    GoTo NextSig
                End If
                headByte = b(j)

                bytes = sig88ToDataOffset
                For k = 1 To 4
                    If k = 1 And headByte <= maxSigByte Then
                        idSize(k) = b(j + 2) 'Ignore first header byte
                    Else
                        idSize(k) = b(j + k)
                    End If
                    If idSize(k) < 37 Or idSize(k) Mod 2 = 0 Then GoTo NextSig
                    idSize(k) = (idSize(k) - 13) / 2
                    bytes = bytes + idSize(k)
                Next k
                If has5HeadBytes Then
                    nameSize = b(j + 5)
                    If nameSize < 15 Or nameSize Mod 2 = 0 Then GoTo NextSig
                    nameSize = (nameSize - 13) / 2
                Else
                    nameSize = (b(j + 5) - 128) * 64 + (b(j + 6) - 13) / 2
                    If nameSize < 1 Or b(j + 6) Mod 2 = 0 Then GoTo NextSig
                End If
                bytes = bytes + nameSize

                currDataEnd = i + bytes - 1
                If currDataEnd > chunkSize Then 'Next chunk
                    i = i - 1
                    Exit Do
                End If
                j = i + sig88ToDataOffset
                #If Mac Then 'StrConv doesn't work reliably on Mac ->
                    sAnsi = MidB$(s, j, idSize(1)) 'Decode ANSI string manually:
                    GoSub DecodeANSI: folderID = sUtf16
                #Else 'Windows
                    folderID = StrConv(MidB$(s, j, idSize(1)), vbUnicode)
                #End If
                j = j + idSize(1)
                parentID = StrConv(MidB$(s, j, idSize(2)), vbUnicode)
                #If Mac Then 'StrConv doesn't work reliably on Mac ->
                    sAnsi = MidB$(s, j, idSize(2)) 'Decode ANSI string manually:
                    GoSub DecodeANSI: parentID = sUtf16
                #Else 'Windows
                    parentID = StrConv(MidB$(s, j, idSize(2)), vbUnicode)
                #End If

                If folderID Like idPattern And parentID Like idPattern Then
                    lastNameStart = j + idSize(2) + idSize(3) + idSize(4)
                    lastNameSize = nameSize
                    lastFolderID = Left(folderID, 32) 'Ignore the "+##.." in IDs
                    lastParentID = Left(parentID, 32) 'of Business OneDrive
                    lastHeadByte = headByte
                    lastDataEnd = currDataEnd
                End If
NextSig:
                i = InStrB(i + 1, s, sig88, vbBinaryCompare)
            Loop
            If i = 0 Then
                lastRecord = lastRecord + chunkSize + headBytes6Offset
            Else
                lastRecord = lastRecord + i + headBytes6Offset
            End If
        Loop Until lastRecord > size
CloseFile:
        Close #fileNum
SkipDbFile:

        'Read cid.ini file
        fileNum = FreeFile()
        Open vDir & cID & ".ini" For Binary Access Read As #fileNum
            ReDim b(0 To LOF(fileNum)): Get fileNum, , b
        Close #fileNum: fileNum = 0
        #If Mac Then 'On Mac, the OneDrive settings files use UTF-8 encoding
            sUtf8 = b: GoSub DecodeUTF8:
            b = sUtf16
        #End If
        Select Case True
        Case dirName Like "Business#" 'Settings files for a business OD account
        'Max 9 Business OneDrive accounts can be signed in at a time.
           Dim libNrToWebColl As Collection: Set libNrToWebColl = New Collection
            mainMount = vbNullString
            For Each line In Split(b, vbNewLine)
                webRoot = "": locRoot = "": parts = Split(line, """")
                Select Case Left$(line, InStr(1, line, " = ", 0) - 1)
                Case "libraryScope" 'One line per synchronized library
                    locRoot = parts(9)
                    syncFind = locRoot: syncID = Split(parts(10), " ")(2)
                    libNr = Split(line, " ")(2)
                    folderType = parts(3): parts = Split(parts(8), " ")
                    siteID = parts(1): webID = parts(2): libID = parts(3)
                    If mainMount = vbNullString Or folderType = "ODB" Then
                        mainMount = locRoot: fileName = "ClientPolicy.ini"
                        mainSyncID = syncID: mainSyncFind = syncFind
                    Else: fileName = "ClientPolicy_" & libID & siteID & ".ini"
                    End If
                    On Error Resume Next 'On error try backup method...
                    webRoot = cliPolColl(fileName)("DavUrlNamespace")
                    On Error GoTo 0
                    If webRoot = "" Then 'Backup method to find webRoot:
                        For Each vItem In cliPolColl
                            If vItem("SiteID") = siteID _
                            And vItem("WebID") = webID _
                            And vItem("IrmLibraryId") = libID Then
                                webRoot = vItem("DavUrlNamespace"): Exit For
                            End If
                        Next vItem
                    End If
                    If webRoot = vbNullString Then Err.Raise vbErrFileNotFound _
                                                           , methodName
                    libNrToWebColl.Add VBA.Array(libNr, webRoot), libNr
                    If Not locRoot = vbNullString Then _
                        locToWebColl.Add VBA.Array(locRoot, webRoot, email, _
                                        syncID, syncFind, dirName), Key:=locRoot
                Case "libraryFolder" 'One line per synchronized library folder
                    libNr = Split(line, " ")(3)
                    locRoot = parts(1): syncFind = locRoot
                    syncID = Split(parts(4), " ")(1)
                    s = vbNullString: parentID = Left$(Split(line, " ")(4), 32)
                    Do  'If not synced at the bottom dir of the library:
                        '   -> add folders below mount point to webRoot
                        On Error Resume Next: odFolders parentID
                        keyExists = (Err.Number = 0): On Error GoTo 0
                        If Not keyExists Then Exit Do
                        s = odFolders(parentID)(1) & "/" & s
                        parentID = odFolders(parentID)(0)
                    Loop
                    webRoot = libNrToWebColl(libNr)(1) & s
                    locToWebColl.Add VBA.Array(locRoot, webRoot, email, _
                                             syncID, syncFind, dirName), locRoot
                Case "AddedScope" 'One line per folder added as link to personal
                    relPath = parts(5): If relPath = " " Then relPath = ""  'lib
                    parts = Split(parts(4), " "): siteID = parts(1)
                    webID = parts(2): libID = parts(3): lnkID = parts(4)
                    fileName = "ClientPolicy_" & libID & siteID & lnkID & ".ini"
                    On Error Resume Next 'On error try backup method...
                    webRoot = cliPolColl(fileName)("DavUrlNamespace") & relPath
                    On Error GoTo 0
                    If webRoot = "" Then 'Backup method to find webRoot:
                        For Each vItem In cliPolColl
                            If vItem("SiteID") = siteID _
                            And vItem("WebID") = webID _
                            And vItem("IrmLibraryId") = libID Then
                                webRoot = vItem("DavUrlNamespace") & relPath
                                Exit For
                            End If
                        Next vItem
                    End If
                    If webRoot = vbNullString Then Err.Raise vbErrFileNotFound _
                                                           , methodName
                    s = vbNullString: parentID = Left$(Split(line, " ")(3), 32)
                    Do 'If link is not at the bottom of the personal library:
                        On Error Resume Next: odFolders parentID
                        keyExists = (Err.Number = 0): On Error GoTo 0
                        If Not keyExists Then Exit Do       'add folders below
                        s = odFolders(parentID)(1) & ps & s 'mount point to
                        parentID = odFolders(parentID)(0)   'locRoot
                    Loop
                    locRoot = mainMount & ps & s
                    locToWebColl.Add VBA.Array(locRoot, webRoot, email, _
                                     mainSyncID, mainSyncFind, dirName), locRoot
                Case Else: Exit For
                End Select
            Next line
        Case dirName = "Personal" 'Settings files for a personal OD account
        'Only one Personal OneDrive account can be signed in at a time.
            For Each line In Split(b, vbNewLine) 'Loop should exit at first line
                If line Like "library = *" Then
                    parts = Split(line, """"): locRoot = parts(3)
                    syncFind = locRoot: syncID = Split(parts(4), " ")(2)
                    Exit For
                End If
            Next line
            On Error Resume Next 'This file may be missing if the personal OD
            webRoot = cliPolColl("ClientPolicy.ini")("DavUrlNamespace") 'account
            On Error GoTo 0                  'was logged out of the OneDrive app
            If locRoot = "" Or webRoot = "" Or cID = "" Then GoTo NextFolder
            locToWebColl.Add VBA.Array(locRoot, webRoot & "/" & cID, email, _
                                       syncID, syncFind, dirName), Key:=locRoot
            If Dir(vDir & "GroupFolders.ini") = "" Then GoTo NextFolder
            'Read GroupFolders.ini file
            cID = vbNullString: fileNum = FreeFile()
            Open vDir & "GroupFolders.ini" For Binary Access Read As #fileNum
                ReDim b(0 To LOF(fileNum)): Get fileNum, , b
            Close #fileNum: fileNum = 0
            #If Mac Then 'On Mac, the OneDrive settings files use UTF-8 encoding
                sUtf8 = b: GoSub DecodeUTF8
                b = sUtf16
            #End If 'Two lines per synced folder from other peoples personal ODs
            For Each line In Split(b, vbNewLine)
                If line Like "*_BaseUri = *" And cID = vbNullString Then
                    cID = LCase$(Mid$(line, InStrRev(line, "/", , 0) + 1, _
                       InStrRev(line, "!", , 0) - InStrRev(line, "/", , 0) - 1))
                    folderID = Left$(line, InStr(1, line, "_", 0) - 1)
                ElseIf cID <> vbNullString Then
                    locToWebColl.Add VBA.Array(locRoot & ps & odFolders( _
                                     folderID)(1), webRoot & "/" & cID & "/" & _
                                     Mid$(line, Len(folderID) + 9), email, _
                                     syncID, syncFind, dirName), _
                                Key:=locRoot & ps & odFolders(folderID)(1)
                    cID = vbNullString: folderID = vbNullString
                End If
            Next line
        End Select
NextFolder:
        cID = vbNullString: s = vbNullString: email = vbNullString
    Next vDir

    'Clean the finished "dictionary" up, remove trailing "\" and "/"
    Dim tmpColl As Collection: Set tmpColl = New Collection
    For Each vItem In locToWebColl
        locRoot = vItem(0): webRoot = vItem(1): syncFind = vItem(4)
        If Right$(webRoot, 1) = "/" Then _
            webRoot = Left$(webRoot, Len(webRoot) - 1)
        If Right$(locRoot, 1) = ps Then _
            locRoot = Left$(locRoot, Len(locRoot) - 1)
        If Right$(syncFind, 1) = ps Then _
            syncFind = Left$(syncFind, Len(syncFind) - 1)
        tmpColl.Add VBA.Array(locRoot, webRoot, vItem(2), _
                              vItem(3), syncFind), locRoot
    Next vItem
    Set locToWebColl = tmpColl

    #If Mac Then 'deal with syncIDs
        If cloudStoragePathExists Then
            Set tmpColl = New Collection
            For Each vItem In locToWebColl
                locRoot = vItem(0): syncID = vItem(3): syncFind = vItem(4)
                locRoot = Replace(locRoot, syncFind, _
                                           syncIDtoSyncDir(syncID)(1), , 1)
                tmpColl.Add VBA.Array(locRoot, vItem(1), vItem(2)), locRoot
            Next vItem
            Set locToWebColl = tmpColl
        End If
    #End If

    GetLocalPath = GetLocalPath(path, returnAll, pmpo, False): Exit Function
    Exit Function
DecodeUTF8: 'UTF-8 must be transcoded to UTF-16 manually in VBA
    Const raiseErrors As Boolean = False 'Raise error if invalid UTF-8 is found?
    Dim o As Long, p As Long, q As Long
    Static numBytesOfCodePoints(0 To 255) As Byte
    Static mask(2 To 4) As Long
    Static minCp(2 To 4) As Long

    If numBytesOfCodePoints(0) = 0 Then
        For o = &H0& To &H7F&: numBytesOfCodePoints(o) = 1: Next o '0xxxxxxx
        '110xxxxx - C0 and C1 are invalid (overlong encoding)
        For o = &HC2& To &HDF&: numBytesOfCodePoints(o) = 2: Next o
        For o = &HE0& To &HEF&: numBytesOfCodePoints(o) = 3: Next o '1110xxxx
       '11110xxx - 11110100, 11110101+ (= &HF5+) outside of valid Unicode range
        For o = &HF0& To &HF4&: numBytesOfCodePoints(o) = 4: Next o
        For o = 2 To 4: mask(o) = (2 ^ (7 - o) - 1): Next o
        minCp(2) = &H80&: minCp(3) = &H800&: minCp(4) = &H10000
    End If
    Dim currByte As Byte
    utf8 = sUtf8
    ReDim utf16(0 To (UBound(utf8) - LBound(utf8) + 1) * 2)
    p = 0
    o = LBound(utf8)
    Do While o <= UBound(utf8)
        codepoint = utf8(o)
        numBytesOfCodePoint = numBytesOfCodePoints(codepoint)
        If numBytesOfCodePoint = 0 Then
            If raiseErrors Then Err.Raise 5
            GoTo insertErrChar
        ElseIf numBytesOfCodePoint = 1 Then
            utf16(p) = codepoint
            p = p + 2
        ElseIf o + numBytesOfCodePoint - 1 > UBound(utf8) Then
            If raiseErrors Then Err.Raise 5
            GoTo insertErrChar
        Else
            codepoint = utf8(o) And mask(numBytesOfCodePoint)
            For q = 1 To numBytesOfCodePoint - 1
                currByte = utf8(o + q)
                If (currByte And &HC0&) = &H80& Then
                    codepoint = (codepoint * &H40&) + (currByte And &H3F)
                Else
                    If raiseErrors Then _
                        Err.Raise 5
                    GoTo insertErrChar
                End If
            Next q
            'Convert the Unicode codepoint to UTF-16LE bytes
            If codepoint < minCp(numBytesOfCodePoint) Then
                If raiseErrors Then Err.Raise 5
                GoTo insertErrChar
            ElseIf codepoint < &HD800& Then
                utf16(p) = CByte(codepoint And &HFF&)
                utf16(p + 1) = CByte(codepoint \ &H100&)
                p = p + 2
            ElseIf codepoint < &HE000& Then
                If raiseErrors Then Err.Raise 5
                GoTo insertErrChar
            ElseIf codepoint < &H10000 Then
                If codepoint = &HFEFF& Then GoTo nextCp '(BOM - will be ignored)
                utf16(p) = codepoint And &HFF&
                utf16(p + 1) = codepoint \ &H100&
                p = p + 2
            ElseIf codepoint < &H110000 Then 'Calculate surrogate pair
                m = codepoint - &H10000
                Dim loSurrogate As Long: loSurrogate = &HDC00& Or (m And &H3FF)
                Dim hiSurrogate As Long: hiSurrogate = &HD800& Or (m \ &H400&)
                utf16(p) = hiSurrogate And &HFF&
                utf16(p + 1) = hiSurrogate \ &H100&
                utf16(p + 2) = loSurrogate And &HFF&
                utf16(p + 3) = loSurrogate \ &H100&
                p = p + 4
            Else
                If raiseErrors Then Err.Raise 5
insertErrChar:  utf16(p) = &HFD
                utf16(p + 1) = &HFF
                p = p + 2
                If numBytesOfCodePoint = 0 Then numBytesOfCodePoint = 1
            End If
        End If
nextCp: o = o + numBytesOfCodePoint 'Move to the next UTF-8 codepoint
    Loop
    sUtf16 = MidB$(utf16, 1, p)
    Return

DecodeANSI: 'Code for decoding ANSI string manually:
    ansi = sAnsi
    p = UBound(ansi) - LBound(ansi) + 1
    If p > 0 Then
        ReDim utf16(0 To p * 2 - 1): q = 0
        For p = LBound(ansi) To UBound(ansi)
            utf16(q) = ansi(p): q = q + 2
        Next p
        sUtf16 = utf16
    Else
        sUtf16 = vbNullString
    End If
    Return
End Function


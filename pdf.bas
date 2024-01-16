Attribute VB_Name = "pdf"
Sub PDFActiveSheet()
    'www.contextures.com
    'for Excel 2010 and later
    Dim wsA As Worksheet
    Dim wbA As Workbook
    Dim strTime As String
    Dim strName As String
    Dim strPath As String
    Dim strFile As String
    Dim strPathFile As String
    Dim myFile As Variant
    On Error GoTo ErrHandler
    
    Set wbA = ActiveWorkbook
    Set wsA = ActiveSheet
    strTime = Format(Now(), "yyyymmdd\_hhmm")
    
    'get active workbook folder, if saved
    strPath = wbA.Path
    If strPath = "" Then
      strPath = Application.DefaultFilePath
    End If
    strPath = strPath & "\"
    
    'replace spaces and periods in sheet name
    strName = Replace(wsA.Name, " ", "")
    strName = Replace(strName, ".", "_")
    
    'create default name for savng file
    strFile = strName & "_" & strTime & ".pdf"
    strPathFile = strPath & strFile
    
    'use can enter name and
    ' select folder for file
    myFile = Application.GetSaveAsFilename _
        (InitialFileName:=strPathFile, _
            FileFilter:="PDF Files (*.pdf), *.pdf", _
            Title:="Select Folder and FileName to save")
    
    'export to PDF if a folder was selected
    If myFile <> "False" Then
        wsA.ExportAsFixedFormat Type:=xlTypePDF, _
            FileName:=myFile, _
            Quality:=xlQualityStandard, _
            IncludeDocProperties:=True, _
            IgnorePrintAreas:=False, _
            OpenAfterPublish:=False
        'confirmation message with file info
        MsgBox "PDF file has been created: " _
          & vbCrLf _
          & myFile
    End If
    
exitHandler:
        Exit Sub
ErrHandler:
        MsgBox "Could not create PDF file"
        Resume exitHandler
End Sub

Option Explicit
Private Sub CommandButton1_Click()
    Dim Pdf_File As String
    
    Pdf_File = Me.TextBox1
    
    Application.ScreenUpdating = False
    Me.WebBrowser1.Navigate "T:BA\MJOP\MLIST"
    If Pdf_File <> "" Then
        Me.WebBrowser1.Document.Write "<HTML><Body><embed src=""" & Pdf_File & _
                                      """ width=""100%"" height=""100%"" /></Body></HTML>"
        Me.WebBrowser1.Refresh
    End If
    Application.ScreenUpdating = True
End Sub
Sub openPDF()
    dosya = "S:\BH FM\PIS\2023\01 - JANEIRO\CEFEMG - CENTRO DE FORMACAO EM ENFERMAGEM DE MINAS LTDA - PI 001 - 77725.pdf"
    CreateObject("Shell.Application").Open dosya
End Sub
Sub listfiles()
'Updateby Extendoffice
    Dim xFSO As Object
    Dim xFolder As Object
    Dim xFile As Object
    Dim xFiDialog As FileDialog
    Dim xPath As String
    Dim vFile As Variant
    Dim i As Integer
    Set xFiDialog = Application.FileDialog(msoFileDialogFolderPicker)
    If xFiDialog.Show = -1 Then
        xPath = xFiDialog.SelectedItems(1)
    End If
    Set xFiDialog = Nothing
    If xPath = "" Then Exit Sub
    Set xFSO = CreateObject("Scripting.FileSystemObject")
    Set xFolder = xFSO.GetFolder(xPath)
   ' For Each xFile In xFolder.Files
   ' If xFSO = "*78968*.*" Then
  '   xx = xFile.Name
        'I = I + 1
   '     achou = "sim"
        'End If
        'ActiveSheet.Hyperlinks.Add Cells(I, 1), xFile.Path, , , xFile.Name
        Range("p1").Value = xPath & "\"
 '   Next
End Sub
Sub listfilesC()
'Updateby Extendoffice
    Dim xFSO As Object
    Dim xFolder As Object
    Dim xFile As Object
    Dim xFiDialog As FileDialog
    Dim xPath As String
    Dim vFile As Variant
    Dim i As Integer
    Set xFiDialog = Application.FileDialog(msoFileDialogFolderPicker)
    If xFiDialog.Show = -1 Then
        xPath = xFiDialog.SelectedItems(1)
    End If
    Set xFiDialog = Nothing
    If xPath = "" Then Exit Sub
    Set xFSO = CreateObject("Scripting.FileSystemObject")
    Set xFolder = xFSO.GetFolder(xPath)
   ' For Each xFile In xFolder.Files
   ' If xFSO = "*78968*.*" Then
  '   xx = xFile.Name
        'I = I + 1
   '     achou = "sim"
        'End If
        'ActiveSheet.Hyperlinks.Add Cells(I, 1), xFile.Path, , , xFile.Name
        Range("R1").Value = xPath & "\"
 '   Next
End Sub
Sub listfilesD()
'Updateby Extendoffice
    Dim xFSO As Object
    Dim xFolder As Object
    Dim xFile As Object
    Dim xFiDialog As FileDialog
    Dim xPath As String
    Dim vFile As Variant
    Dim i As Integer
    Set xFiDialog = Application.FileDialog(msoFileDialogFolderPicker)
    If xFiDialog.Show = -1 Then
        xPath = xFiDialog.SelectedItems(1)
    End If
    Set xFiDialog = Nothing
    If xPath = "" Then Exit Sub
    Set xFSO = CreateObject("Scripting.FileSystemObject")
    Set xFolder = xFSO.GetFolder(xPath)
   ' For Each xFile In xFolder.Files
   ' If xFSO = "*78968*.*" Then
  '   xx = xFile.Name
        'I = I + 1
   '     achou = "sim"
        'End If
        'ActiveSheet.Hyperlinks.Add Cells(I, 1), xFile.Path, , , xFile.Name
        Range("R1").Value = xPath & "\"
 '   Next
End Sub
Function FileExists(filepath As String) As Boolean
Dim TestStr, abrir As String

    TestStr = ""
    On Error Resume Next
    TestStr = Dir(filepath)
    On Error GoTo 0
    If TestStr = "" Then
        FileExists = False
    Else
        
        FileExists = True
        UserForm_copy.TextB_pesq.Value = TestStr
    End If
End Function
Sub FileExistsWildCardDemoRAP()
'VBA Check if File Exists

Dim strFolder  As String
Dim strFile As String
 strFolder = Range("p1").FormulaR1C1 & "*80131*"
strFile = strFolder  'UserForm_copy.TextB_cont_P.Text & "*.*"
If FileExists(strFile) Then
    dosya = Range("p1").FormulaR1C1 & UserForm_copy.TextB_pesq.Value
    CreateObject("Shell.Application").Open dosya
Else
   'File beginning with A and ending with .txt exists does not Exist
End If
End Sub
Sub FileExistsWildCardDemoCli()
'VBA Check if File Exists
Dim strFile As String
strFile = Range("p1").FormulaR1C1 & "*" & UserForm_copy.TextB_cli_P.Text & "*"
If FileExists(strFile) Then
    dosya = Range("p1").FormulaR1C1 & UserForm_copy.TextB_pesq.Value
    CreateObject("Shell.Application").Open dosya
Else
   'File beginning with A and ending with .txt exists does not Exist
End If
End Sub
Sub CheckFileExists()

Dim strFileName As String
Dim strFileExists As String

    strFileName = Range("p1").FormulaR1C1 & "*" & UserForm_copy.TextB_cli_P.Text & "*.*"
    strFileExists = Dir(strFileName)

   If strFileExists = "" Then
        MsgBox "The selected file doesn't exist"
    Else
        MsgBox "The selected file exists"
        UserForm_copy.TextB_pesq.Value = strFileExists
    End If

End Sub
Sub FileExistsWildCardDemo()
'VBA Check if File Exists
Dim strFile As String
strFile = Range("p1").FormulaR1C1 & UserForm_copy.TextB_cli_P.Text & "*.pdf"
If FileExists(strFile) Then
    'File beginning with A and ending with .txt exists
        UserForm_copy.TextB_pesq.Value = strFile
Else
    'File beginning with A and ending with .txt exists does not Exist
End If
End Sub
Sub Test_File_Exist_With_Dir()
    Dim filepath As String
    Dim TestStr As String

    filepath = Range("p1").FormulaR1C1 & "*" & UserForm_copy.TextB_cli_P.Text & "*.*"

    TestStr = ""
    On Error Resume Next
    TestStr = Dir(filepath)
    On Error GoTo 0
    If TestStr = "" Then
        MsgBox "File doesn't exist"
    Else
        MsgBox "File exist"
    End If

End Sub
Sub Test_File_EDir()
    Dim filepath As String
    Dim TestStr As String

    'After 27 file name charaters testing if the file exist is not working anymore
    filepath = Range("p1").FormulaR1C1 & "*80131*"

    TestStr = ""
    On Error Resume Next
    TestStr = Dir(filepath)
    On Error GoTo 0
    If TestStr = "" Then
        MsgBox "File doesn't exist"
    Else
        MsgBox "File exist"
    End If

End Sub
Sub Listar_arquivos_mp3()
Dim i As Long
Dim sh As Worksheet
Dim iSomaMb As Double
Dim sPasta As Variant
Dim iLinha As Long
    'Set sh = ThisWorkbook.ActiveSheet
    'Exibe a caixa para escolha da pasta onde será feita a pesquisa
    sPasta = Range("p1").FormulaR1C1 'GetPasta
        'If sPasta = "" Then
        '    Exit Sub        'Cancela pesquisa
        'End If
    'Apaga o conteúdo
     'sh.Range("B:C").EntireColumn.ClearContents
    'Escreve o cabeçalho
        'sh.Cells(4, 2).Value = "Música"
        'sh.Cells(4, 3).Value = "Tamanho (Mb)"
    'Define a linha inicial da listagem
    iLinha = 5
    'Application.StatusBar = "Aguarde... Pesquisando ... "
    'Usa o objeto de pesquisa
    With Application.FileSearch
        .LookIn = sPasta                        'Define a pasta onde será pesquisado
        .FileName = "*80131*" '.mp3"                     'Define o termo da pesquisa
        .SearchSubFolders = True                'Informa se será feita a pesquisa nas subpastas
        .Execute                                'Executa a pesquisa  Ohhhhh!!!!
        'Percorre os itens encontrados e escreve na planilha
        For i = 1 To .FoundFiles.Count
            'sh.Cells(iLinha, 2).Value
            achei = .FoundFiles(i)
           ' sh.Cells(iLinha, 3).Value = CDbl(Format((FileLen(.FoundFiles(i)) / 1048576), "0.00"))
           ' iSomaMb = iSomaMb + sh.Cells(iLinha, 3).Value
           ' iLinha = iLinha + 1
           ' Application.StatusBar = "Preenchendo lista ... " & Format(i / .FoundFiles.Count, "0%")
        Next i
        'sh.Cells(1, 2).Value = "Músicas em " & sPasta
        'sh.Cells(2, 2).Value = "Total de Músicas: " & .FoundFiles.Count
        'sh.Cells(3, 2).Value = "Espaço Utilizado: " & Format(iSomaMb, "0.00") & " MB"
    End With
    'sh.Range("A1").Select
   ' Application.StatusBar = False
End Sub
Sub SearchDir()
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim filepath As String
    Dim FileName As String
    Dim rng As Range
    Dim i As Variant
    Dim results As Worksheet
    Dim resultslr As Long
    Dim lastrow As Long
    Dim searchString As String
    
    'Set results = ThisWorkbook.Worksheet("results")
    
    filepath = Range("p1").FormulaR1C1
    FileName = Dir(filepath & "*80131*") 'filepath & "*.csv"
    
    searchString = "what you're searching for"
    
    'Do While filename <> ""
    
        If FileName Like searchString Then
           'resultslr = results.Cells(Rows.Count, "a").End(xlUp).Row + 1
            achei = FileName 'UserForm_copy.TextB_pesq.Value
         End If
    
        'Set wb = Excel.Workbooks.Open(filepath & filename)   'opens the file
       ' Set ws = wb.Worksheets(1)      'sets the worksheet within the csv
    
        'lastrow = ws.Cells(Rows.Count, "a").End(xlUp).Row
        'Set rng = Range("A2:someEndColumn" & lastrow)   'replace someEndColumn with however your dat ais arranged.
    
        For Each i In rng                                   'searches each cell in the range for your seachString and puts the results in a list on worksheet 'results'
           If i.Value Like searchString Then
              'resultslr = results.Cells(Rows.Count, "a").End(xlUp).Row + 1
              'results.Cells(resultslr, "a").Resize(1, 1).Value = i.Value
           End If
         Next i
    
        'wb.Close False               'closes the file
    
        FileName = Dir                'next file in directory
    ''more code for other stuff
End Sub
Option Explicit
'----------- ExcelBaby.com -----------
'-------------- Modules --------------
Sub ListFile()
    ''Description: List all files in folder and sub-folders (include hidden ,read only...)
    ''Web Site: https://excelbaby.com
    ''Url: https://excelbaby.com/learn/excel-macro-list-all-files-in-folders-and-subfolders/

    Dim PathSpec As String
    PathSpec = ""   'Specify a folder
    If (PathSpec = "") Then PathSpec = SelectSingleFolder   'Browse for Folder to select a folder

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")    'Late Binding
    If (fso.FolderExists(PathSpec) = False) Then Exit Sub   'folder exist or not?

    Application.ScreenUpdating = False 'Disable Screen Updating to speed up macro
    
    Dim MySheetName As String
    MySheetName = "Files"   'Add a Sheet with name "Files"
    AddSheet (MySheetName)

    Dim FileType As String
    FileType = "*"   '*:all, or pdf, PDF, XLSX...
    FileType = UCase(FileType)

    Dim queue As Collection, oFolder As Object, oSubfolder As Object, oFile As Object
    Dim LastBlankCell As Long, FileExtension As String

    Set queue = New Collection
    queue.Add fso.GetFolder(PathSpec) 'enqueue
    
    Do While queue.Count > 0
        Set oFolder = queue(1)
        queue.Remove 1 'dequeue
        
        For Each oSubfolder In oFolder.SubFolders   'loop all sub-folders
            queue.Add oSubfolder 'enqueue
            '...insert any folder processing code here...
        Next oSubfolder
        
        LastBlankCell = ThisWorkbook.Sheets(MySheetName).Cells(Rows.Count, 1).End(xlUp).Row + 1 'get the last blank cell of column A
        
        For Each oFile In oFolder.Files 'loop all files
            FileExtension = UCase(Split(oFile.Name, ".")(UBound(Split(oFile.Name, ".")))) 'get file extension, eg: TXT
            If (FileType = "*" Or FileExtension = FileType) Then
                With ThisWorkbook.Sheets(MySheetName)
                    .Cells(LastBlankCell, 1) = oFile 'Path
                    .Cells(LastBlankCell, 2) = oFolder 'Folder
                    .Cells(LastBlankCell, 3) = oFile.Name 'File Name
                    .Cells(LastBlankCell, 4) = FileExtension 'File Extension
                    .Cells(LastBlankCell, 5) = oFile.DateCreated 'Data Created
                    .Cells(LastBlankCell, 6) = oFile.DateLastAccessed 'Last Accessed
                    .Cells(LastBlankCell, 7) = oFile.DateLastModified 'Last Modified
                    .Cells(LastBlankCell, 8) = oFile.Size 'File Size
                    If (oFile.Attributes And 2) = 2 Then
                        .Cells(LastBlankCell, 9) = "TRUE" 'Is Hidden
                    Else
                        .Cells(LastBlankCell, 9) = "FALSE" 'Is Hidden
                    End If
                End With
                LastBlankCell = LastBlankCell + 1
            End If
        Next oFile
    Loop
    
    'Cells.EntireColumn.AutoFit  'Autofit columns width
    Application.ScreenUpdating = True

End Sub
Function SelectSingleFolder()
    'Select a Folder Path
    
    Dim FolderPicker As FileDialog
    Dim myFolder As String
    
    'Select Folder with Dialog Box
    Set FolderPicker = Application.FileDialog(msoFileDialogFolderPicker)
    
    With FolderPicker
        .Title = "Select A Single Folder"
        .AllowMultiSelect = False
        If .Show <> -1 Then Exit Function 'Check if user clicked cancel button
        SelectSingleFolder = .SelectedItems(1)
    End With
End Function
Sub UseFileDialogOpen()
 
    Dim lngCount As Long
 
    ' Open the file dialog
    With Application.FileDialog(msoFileDialogOpen)
        .AllowMultiSelect = True
        .Show
 
        ' Display paths of each file selected
        For lngCount = 1 To .SelectedItems.Count
            MsgBox .SelectedItems(lngCount)
        Next lngCount
 
    End With
 
End Sub
Sub GetFilePath()
    busca = Range("p1").FormulaR1C1
    If busca = "" Then
     MsgBox "Registre o caminho para o comprovante deste mês!"
     Else
    Set myFile = Application.FileDialog(msoFileDialogOpen)
    myFile.ButtonName = "ABRIR PI"
    myFile.Title = "SELECIONE O PI PARA SER ABERTO"
    
    With myFile
        .InitialFileName = busca ' & " test"
        .Title = "#CADÊ O PI?"
        .AllowMultiSelect = False
        If .Show <> -1 Then
            Exit Sub
        End If
    FileSelected = .SelectedItems(1)
    End With
        dosya = FileSelected
        CreateObject("Shell.Application").Open dosya
    End If
'ActiveSheet.Range("A1") = FileSelected
End Sub
Sub GetFilePathC()
    busca = Range("R1").FormulaR1C1
    
    If busca = "" Then
     MsgBox "Registre o caminho para a pasta partidária!"
     Else
    Set myFile = Application.FileDialog(msoFileDialogOpen)
    myFile.ButtonName = "ABRIR PARTIDO"
    myFile.Title = "SELECIONE PARA SER ABERTO"
    
    
    With myFile
        .InitialFileName = busca ' & " test"
        .Title = "#CADÊ O PARTIDO?"
        .AllowMultiSelect = False
        If .Show <> -1 Then
            Exit Sub
        End If
    FileSelected = .SelectedItems(1)
    End With
        dosya = FileSelected
        CreateObject("Shell.Application").Open dosya
    End If
'ActiveSheet.Range("A1") = FileSelected
End Sub
Sub GetFilePathd()
    busca = Range("R1").FormulaR1C1
    
    If busca = "" Then
     MsgBox "Registre o caminho para o comprovante deste mês!"
     Else
    Set myFile = Application.FileDialog(msoFileDialogOpen)
    myFile.ButtonName = "ABRIR COMPROVANTE"
    myFile.Title = "SELECIONE O COMPROVANTE PARA SER ABERTO"
    
    
    With myFile
        .InitialFileName = busca ' & " test"
        .Title = "#CADÊ O COMPROVANTE?"
        .AllowMultiSelect = False
        If .Show <> -1 Then
            Exit Sub
        End If
    FileSelected = .SelectedItems(1)
    End With
        dosya = FileSelected
        CreateObject("Shell.Application").Open dosya
    End If
    'ActiveSheet.Range("A1") = FileSelected
End Sub
Sub show_final_opendialog()
    Dim oFD As FileDialog
    Dim oFD1 As FileDialog
    Dim vItem As Variant
    
    Set oFD = Application.FileDialog(msoFileDialogOpen)
    oFD.ButtonName = "Press me to Go"
    oFD.Title = "Select a Single File You'd like to Open"
    oFD.AllowMultiSelect = True
    
    'oFD.Filters.Clear
    'oFD.Filters.Add "Special", "*.special"
    'oFD.Filters.Add "Text and Excel", "*.xls, *.txt"
    
    oFD.InitialView = msoFileDialogViewDetails
    oFD.InitialFileName = Range("p1").FormulaR1C1 & "*81*"
    
    If oFD.Show <> 0 Then
        For Each vItem In oFD.SelectedItems
            '
            'add your file processing code here
            '
            Debug.Print vItem 'prints the file path of the first file selected
        Next
    End If
    
    Set oFD = Nothing

End Sub

Attribute VB_Name = "EloquaCoding"
Option Explicit

Sub process_query_to_eloqua()
Attribute process_query_to_eloqua.VB_ProcData.VB_Invoke_Func = " \n14"
'
' process_query_to_eloqua Macro
    'definition of initial parameters
    
    Dim list
    'list of the columns we want to keep
    list = Array("contact_email", "contact_first_name", "contact_last_name", "contact_company_name", "contact_country", _
    "contact_phone_contact_cleansed", "contact_phone_mobile", "site_city", "contact_job_function")
    
    Dim new_sheet_name As String
    'name of the new_sheet we're creating
    new_sheet_name = "processed_data"
    
    Dim link_to_reference_companies As String
    'link_to_reference_companies = "C:\Users\tamadei\Desktop\atos\references\reference_file_companies.xlsx"
    link_to_reference_companies = searchFile("reference_file_companies.xlsx")
    
    Dim link_to_reference_countries As String
    'link_to_reference_countries = "C:\Users\tamadei\Desktop\atos\references\reference_file_countries.xlsx"
    link_to_reference_countries = searchFile("reference_file_countries.xlsx")
    
    Dim link_to_reference_jobs As String
    'link_to_reference_jobs = "C:\Users\tamadei\Desktop\atos\references\reference_file_jobs.xlsx"
    link_to_reference_jobs = searchFile("reference_file_jobs.xlsx")
    
    Dim python_executable_path As String
    'python_executable_path = "C:\Users\tamadei\Anaconda3\python.exe"
    python_executable_path = get_python_path()
    
    Dim python_script_path As String
    'python_script_path = "C:\Users\tamadei\Desktop\atos\jupyter\scripts\runModel.py"
    python_script_path = searchFile("runModel.py")
    
    Dim model_path As String
    'model_path = "C:\Users\tamadei\Desktop\atos\jupyter\models\model.h5"
    model_path = searchFile("model.h5")
    
    Dim classes_path As String
    'classes_path = "C:\Users\tamadei\Desktop\atos\jupyter\classes\classes"
    classes_path = searchFile("classes_for_job_functions_classification")
    
    

    'set initial parameters to improve overall performance
    Dim savedCalcMode As XlCalculation
    savedCalcMode = Application.Calculation
    Application.Calculation = xlCalculationManual
    
    Dim savedScreenUpdating As Boolean
    savedScreenUpdating = Application.ScreenUpdating
    Application.ScreenUpdating = False
    
    Dim savedEnableEvents As Boolean
    savedEnableEvents = Application.EnableEvents
    Application.EnableEvents = False
    
    Dim savedPageBrakes As Boolean
    savedPageBrakes = ActiveSheet.DisplayPageBreaks
    ActiveSheet.DisplayPageBreaks = False
    
    Dim savedEnableAnimations As Boolean
    savedEnableAnimations = Application.EnableAnimations
    Application.EnableAnimations = False
    
    Dim savedStatusBar As Boolean
    savedStatusBar = Application.DisplayStatusBar
    Application.DisplayStatusBar = False
    
    Dim savedPrintCommunication As Boolean
    savedPrintCommunication = Application.PrintCommunication
    Application.PrintCommunication = False
    
    'rename activate sheet to "result" if not already the case
    If Not ActiveSheet.name = "result" Then
        On Error Resume Next
        ActiveSheet.name = "result"
    End If
            
    'Check whether sheet "processed_data" exists
    Dim sheetExists As Boolean
    Dim sht As Worksheet
    
    On Error Resume Next
    Set sht = ActiveWorkbook.Sheets("processed_data")
    On Error GoTo 0
    sheetExists = Not sht Is Nothing
    
    'If sheet "new_sheet_name" doesn't exist, creates and adds it
    Dim ws As Object
    If Not sheetExists Then
        Set ws = ActiveWorkbook.Sheets.Add(After:= _
                 ActiveWorkbook.Sheets(ActiveWorkbook.Sheets.Count))
        ws.name = new_sheet_name
    End If
    
    
    ActiveWorkbook.Sheets("result").Activate
    ActiveWorkbook.Save
    With ActiveWorkbook.Sheets("result")
        
        'gets the indices of the final row and the final column
        Dim lastRow As Long, lastCol As Long
        lastRow = Range("A" & Rows.Count).End(xlUp).row
        lastCol = Cells(1, Columns.Count).End(xlToLeft).column
        
        Call setAllNamesToProperOnes(lastRow, lastCol)
        
        'now we run our python script
        'it calls our neural network model to classify the job functions of people who do not have it filled in the database
        'it then copies the classes found in our current database
        
        'Dim shell_return As String
        Call runScript_neuralNetworkModel(python_executable_path, python_script_path, model_path, classes_path)
        
        'copy all rows that we want to keep on the new sheet, said rows correspond to the list at the beginning of the function
        Dim i As Long
        Dim rowWanted As Integer
        For i = 1 To lastCol
            rowWanted = indexInList(list, Cells(1, i).value) 'determines whether the row we are currently on is on the list or not
            'if yes, determines at which index it is in the list
            If rowWanted <> -1 Then
                'copies and pastes the right row on the new sheet
                Sheets(new_sheet_name).Range(Sheets(new_sheet_name).Cells(1, rowWanted + 1), _
                Sheets(new_sheet_name).Cells(lastRow, rowWanted + 1)).value = _
                Sheets("result").Range(Sheets("result").Cells(1, i), Sheets("result").Cells(lastRow, i)).value
            End If
        Next i
        
        
        
        
        'reads the reference files and gets their data in dictionaries
        Dim dict_countries As Object, dict_companies As Object, dict_phone As Object, dict_jobs As Object
        Set dict_countries = read_reference_file_countries(link_to_reference_countries)
        Set dict_companies = read_reference_file_companies(link_to_reference_companies)
        Set dict_phone = read_reference_file_phone_numbers(link_to_reference_countries)
        Set dict_jobs = read_reference_file_jobs(link_to_reference_jobs)
        
        Dim lastCol_newSheet As Integer
        Dim pctdone As Single 'will be used to display the progress bar
        lastCol_newSheet = Sheets(new_sheet_name).Cells(1, Sheets(new_sheet_name).Columns.Count).End(xlToLeft).column + 1
        Sheets(new_sheet_name).Cells(1, lastCol_newSheet).value = "ABM-SP-EU"
        
        'final step: create the code of each element of the base
        
        '(Step 1) Display the Progress Bar
        ufProgress.LabelProgress.Width = 0
        ufProgress.Show
        
        For i = 2 To lastRow
            '(Step 2) Periodically update progress bar
            pctdone = i / lastRow
            With ufProgress
                .LabelCaption.Caption = "Processing Row " & i & " of " & lastRow
                .LabelProgress.Width = pctdone * (.FrameProgress.Width)
            End With
            DoEvents
        
            'create code for each line of the database
            Sheets(new_sheet_name).Cells(i, lastCol_newSheet).value = code(i, lastCol, dict_countries, dict_companies, dict_phone, dict_jobs)
            
            '(Step 3) Close the progress bar when we're reached the end
            If i = lastRow Then Unload ufProgress
        Next i
        
    End With
    ActiveWorkbook.Sheets(new_sheet_name).Select
    
    'rename the columns to fit eloqua's style
    With ActiveWorkbook.Sheets(new_sheet_name)
        Range("A1").value = "Email Address"
        Range("B1").value = "First Name"
        Range("C1").value = "Last Name"
        Range("D1").value = "Company Name"
        Range("E1").value = "Country"
        Range("F1").value = "Business Phone"
        Range("G1").value = "Mobile Phone"
        Range("H1").value = "City"
        Range("I1").value = "Job Function"
    End With
    
    'reset parameters as they were initially
    Application.Calculation = savedCalcMode
    Application.ScreenUpdating = True
    Application.EnableEvents = savedEnableEvents
    ActiveSheet.DisplayPageBreaks = savedPageBrakes
    Application.EnableAnimations = savedEnableAnimations
    Application.DisplayStatusBar = savedStatusBar
    Application.PrintCommunication = savedPrintCommunication
'
End Sub

Function indexInList(list, element)

    'return the index of the element in the list, or -1 if the element is not in the list
    
    indexInList = -1
    Dim i As Integer
    Dim length As Long
    length = UBound(list) - LBound(list) + 1
    
    For i = 0 To length - 1
        If StrComp(list(i), element, vbTextCompare) = 0 Then
            indexInList = i
            Exit For
        End If
    Next i
    
    
End Function


Function countryEncoding(row, lastCol, dict_countries, dict_phone)
    countryEncoding = "0ZZZ"
    
    'we first look at the fiels "contact_country"
    Dim i As Integer
    For i = 1 To lastCol
        If StrComp(Sheets("result").Cells(1, i).value, "contact_country", vbTextCompare) = 0 Then
            If Not (IsEmpty(Sheets("result").Cells(row, i)) Or Sheets("result").Cells(row, i).value = "") Then
                countryEncoding = country_code(Sheets("result").Cells(row, i).value, dict_countries)
            End If
            Exit For
        End If
    Next i
    
    'if the code for the country is still ZZZ then we look for the information in company_country
    If countryEncoding = "0ZZZ" Then
        For i = 1 To lastCol
        If StrComp(Sheets("result").Cells(1, i).value, "company_country", vbTextCompare) = 0 Then
            If Not (IsEmpty(Sheets("result").Cells(row, i)) Or Sheets("result").Cells(row, i).value = "") Then
                countryEncoding = country_code(Sheets("result").Cells(row, i).value, dict_countries)
            End If
            Exit For
        End If
    Next i
    End If
    
    'if the code for the country is still ZZZ then we look for the information in contact_phone_contact_cleansed
    Dim country_from_number As String
    If countryEncoding = "0ZZZ" Then
        For i = 1 To lastCol
        If StrComp(Sheets("result").Cells(1, i).value, "contact_phone_contact_cleansed", vbTextCompare) = 0 Then
            If Not (IsEmpty(Sheets("result").Cells(row, i)) Or Sheets("result").Cells(row, i).value = "") Then
                country_from_number = country_from_phone_number(Sheets("result").Cells(row, i).value, dict_phone)
                countryEncoding = country_code(country_from_number, dict_countries)
            End If
            Exit For
        End If
    Next i
    End If
    
    'if the code for the country is still ZZZ then we look for the information in contact_phone_mobile_cleansed
    If countryEncoding = "0ZZZ" Then
        For i = 1 To lastCol
        If StrComp(Sheets("result").Cells(1, i).value, "contact_phone_mobile_cleansed", vbTextCompare) = 0 Then
            If Not (IsEmpty(Sheets("result").Cells(row, i)) Or Sheets("result").Cells(row, i).value = "") Then
                country_from_number = country_from_phone_number(Sheets("result").Cells(row, i).value, dict_phone)
                countryEncoding = country_code(country_from_number, dict_countries)
            End If
            Exit For
        End If
    Next i
    End If
    
End Function

Function country_code(country, dict_countries)
    country_code = "0ZZZ"
    
    Dim key As Variant
    For Each key In dict_countries.Keys
        If StrComp(country, key, vbTextCompare) = 0 Then
            country_code = dict_countries(key)
            Exit For
        End If
    Next key

End Function

Function country_from_phone_number(phone_number, dict_phone)

    Dim split_number() As String
    split_number = Split(phone_number, " ")
    Dim phone_code As String
    phone_code = split_number(0)
    
    Dim key As Variant
    For Each key In dict_phone.Keys
        If StrComp(phone_code, key, vbTextCompare) = 0 Then
            country_from_phone_number = dict_phone(key)
            Exit For
        End If
    Next key
    

End Function

Function companyEncoding(row, lastCol, dict_companies)

    Dim domain_name As String
    'we look for the column of the domain_name
    Dim i As Integer
    For i = 1 To lastCol
        If StrComp(Sheets("result").Cells(1, i).value, "contact_email_domain", vbTextCompare) = 0 Then
            'we have found the right column
            domain_name = Sheets("result").Cells(row, i).value
            Exit For
        End If
    Next i
    
    companyEncoding = "ZZ" 'set this default value, will return it if no match is found with a company
    
    Dim key As Variant
    For Each key In dict_companies.Keys
        If StrComp(domain_name, key, vbTextCompare) = 0 Then
            companyEncoding = dict_companies(key)
            Exit For
        End If
    Next key
    
End Function

Function runScript_neuralNetworkModel(python_executable_path, python_script_path, model_path, classes_path)

    Dim dataset_path As String
    dataset_path = ActiveWorkbook.path & "\" & ActiveWorkbook.name
    
    Dim command_to_run_python_script As String
    command_to_run_python_script = python_executable_path & " " & python_script_path & " " & """" & dataset_path & """" & " " & _
    """" & model_path & """" & " " & """" & classes_path & """"
    
    Dim obj_shell As Object
    Set obj_shell = VBA.CreateObject("Wscript.Shell")
    obj_shell.Run command_to_run_python_script, 1, True
    
    Application.Wait Now + TimeValue("0:00:03")
    
    
End Function

Function jobEncoding(row, lastCol, dict_jobs)

    Dim job_function As String
    'we look for the column of the contact_job_function
    Dim i As Integer
    For i = 1 To lastCol
        If StrComp(Sheets("result").Cells(1, i).value, "contact_job_function", vbTextCompare) = 0 Then
            'we found the right column
            job_function = Sheets("result").Cells(row, i).value
            Exit For
        End If
    Next i
    
    jobEncoding = vbNullString
    
    Dim key As Variant
    For Each key In dict_jobs.Keys
        If StrComp(job_function, key, vbTextCompare) = 0 Then
            jobEncoding = dict_jobs(key)
            Exit For
        End If
    Next key
    
    If jobEncoding = vbNullString Then
        'the person's job function either is not referenced or does not correspond to our list
        jobEncoding = "JX00"
    End If

End Function

Function code(row, lastCol, dict_countries, dict_companies, dict_phone, dict_jobs)

    'for the person on the row 'row', returns the eloqua code

    code = "[ThR-GSI]-"
    
    Dim company_code As String
    company_code = companyEncoding(row, lastCol, dict_companies)
    code = code & company_code & "-"
    
    Dim country_code As String
    country_code = countryEncoding(row, lastCol, dict_countries, dict_phone)
    code = code & country_code & "--"
    
    Dim job_code As String
    job_code = jobEncoding(row, lastCol, dict_jobs)
    code = code & job_code

End Function

Public Function GetProperName(ByVal TextToConvert As String) As String
    Dim i As Integer
    Dim separateur
    separateur = Array(" ", ";", ":", "-", "~", "@", "_", "&", "*", "#", "'", ".", Chr(160))
    TextToConvert = LCase(TextToConvert)
     
    'mettre la première lettre en majuscule
    TextToConvert = UCase(Mid(TextToConvert, 1, 1)) + Right(TextToConvert, Len(TextToConvert) - 1)
     
    'mettre en majuscule après chaque séparateur
    For i = 1 To Len(TextToConvert) - 1
        If UBound(Filter(separateur, Mid(TextToConvert, i, 1))) >= 0 Then
            TextToConvert = Left(TextToConvert, i) + UCase(Mid(TextToConvert, i + 1, 1)) + Right(TextToConvert, Len(TextToConvert) - i - 1)
        End If
    Next i
     
    GetProperName = TextToConvert
End Function

Function setAllNamesToProperOnes(lastRow, lastCol)

    Dim pctdone As Single 'will be used to display the progress bar
    
    'we look for the column "contact_first_name"
    Dim column As Integer, i As Long
    For column = 1 To lastCol
        If StrComp(Sheets("result").Cells(1, column).value, "contact_first_name", vbTextCompare) = 0 Then
            'that's the right column
            
            ufProgress.LabelProgress.Width = 0
            ufProgress.Show
            
            For i = 2 To lastRow
            
                pctdone = i / (2 * lastRow)
                With ufProgress
                    .LabelCaption.Caption = "Setting Proper Names for row " & i & " out of " & (2 * lastRow)
                    .LabelProgress.Width = pctdone * (.FrameProgress.Width)
                End With
                DoEvents
            
                On Error Resume Next
                Sheets("result").Cells(i, column).value = GetProperName(Sheets("result").Cells(i, column).value)
            Next i
    
            Exit For
        End If
    Next column
    
    'now we look for the column "contact_last_name" and follow the same recipe
    For column = 1 To lastCol
        If StrComp(Sheets("result").Cells(1, column).value, "contact_last_name", vbTextCompare) = 0 Then
            'that's the right column
            
            For i = 2 To lastRow
                
                pctdone = (lastRow + i) / (2 * lastRow)
                With ufProgress
                    .LabelCaption.Caption = "Setting Proper Names for row " & (lastRow + i) & " out of " & (2 * lastRow)
                    .LabelProgress.Width = pctdone * (.FrameProgress.Width)
                End With
                DoEvents
                
                If i = lastRow Then Unload ufProgress
            
                On Error Resume Next
                Sheets("result").Cells(i, column).value = GetProperName(Sheets("result").Cells(i, column).value)
            Next i
            Exit For
        End If
    Next column

End Function

Function read_reference_file_companies(link_to_reference_companies)

    'stores the data in the reference_file 'reference_file_companies' in a dictionary, that is returned

    Dim reference_file As workbook
    Set reference_file = Workbooks.Open(link_to_reference_companies)
    
    With reference_file.Sheets("company_code")
        Dim lastRow As Long
        lastRow = Range("B" & Rows.Count).End(xlUp).row
        
        Dim dict As Object
        Set dict = CreateObject("Scripting.Dictionary")
        
        Dim i As Long
        Dim key As String
        Dim value As String
        
        For i = 2 To lastRow
            key = Range("B" & i).value
            value = Range("D" & i).value
            dict.Add key, value
        Next i
    
    End With
    reference_file.Close SaveChanges:=False
    
    Set read_reference_file_companies = dict
End Function

Function read_reference_file_countries(link_to_reference_countries)

    'stores the data in the reference_file 'reference_file_countries' in a dictionary, that is returned

    Dim reference_file As workbook
    Set reference_file = Workbooks.Open(link_to_reference_countries)
    
    With reference_file.Sheets("country_code")
        Dim lastRow As Long
        lastRow = Range("B" & Rows.Count).End(xlUp).row
        
        Dim dict As Object
        Set dict = CreateObject("Scripting.Dictionary")
        
        Dim i As Long
        Dim key As String
        Dim value As String
        Dim region As String, region_code As Integer
        
        For i = 2 To lastRow
            key = Range("B" & i).value
            value = Range("D" & i).value
            region = Trim(Range("E" & i).value)
            If StrComp(region, "AMER", vbTextCompare) = 0 Then
                region_code = 1
            Else
                If StrComp(region, "APAC", vbTextCompare) = 0 Then
                    region_code = 2
                Else
                    If StrComp(region, "EMEA", vbTextCompare) = 0 Then
                        region_code = 3
                    Else
                        region_code = 0
                    End If
                End If
            End If
            value = region_code & value
            If Not dict.Exists(key) Then
                dict.Add key, value
            End If
        Next i
    
    End With
    reference_file.Close SaveChanges:=False
    
    Set read_reference_file_countries = dict
End Function

Function read_reference_file_phone_numbers(link_to_reference_countries)

    'stores the data in the reference_file 'reference_file_countries' in a dictionary, that is returned

    Dim reference_file As workbook
    Set reference_file = Workbooks.Open(link_to_reference_countries)
    
    With reference_file.Sheets("country_code")
        Dim lastRow As Long
        lastRow = Range("G" & Rows.Count).End(xlUp).row
        
        Dim dict As Object
        Set dict = CreateObject("Scripting.Dictionary")
        
        Dim i As Long
        Dim key As String
        Dim value As String
        
        For i = 2 To lastRow
            key = Range("G" & i).value
            value = Range("H" & i).value
            If Not dict.Exists(key) Then
                dict.Add key, value
            End If
        Next i
    
    End With
    reference_file.Close SaveChanges:=False
    
    Set read_reference_file_phone_numbers = dict
End Function

Function read_reference_file_jobs(link_to_reference_jobs)

    'stores the data in the reference_file 'reference_file_jobs' in a dictionary, that is returned

    Dim reference_file As workbook
    Set reference_file = Workbooks.Open(link_to_reference_jobs)
    
    With reference_file.Sheets("jobs_code")
        Dim lastRow As Long
        lastRow = Range("B" & Rows.Count).End(xlUp).row
        
        Dim dict As Object
        Set dict = CreateObject("Scripting.Dictionary")
        
        Dim i As Long
        Dim key As String
        Dim value As String
        
        For i = 2 To lastRow
            key = Range("B" & i).value
            value = Range("C" & i).value
            If Not dict.Exists(key) Then
                dict.Add key, value
            End If
        Next i
    
    End With
    reference_file.Close SaveChanges:=False
    
    Set read_reference_file_jobs = dict
End Function

Public Function ShellRun(sCmd As String) As String

    'Run a shell command, returning the output as a string

    Dim oShell As Object
    Set oShell = CreateObject("WScript.Shell")

    'run command
    Dim oExec As Object
    Dim oOutput As Object
    Set oExec = oShell.Exec(sCmd)
    Do While oExec.Status = 0
        Application.Wait Now + TimeValue("0:00:01")
    Loop
    Set oOutput = oExec.StdOut

    'handle the results as they are written to and read from the StdOut object
    Dim s As String
    Dim sLine As String
    While Not oOutput.AtEndOfStream
        sLine = oOutput.ReadLine
        If sLine <> "" Then s = s & sLine & vbCrLf
    Wend
    
    ShellRun = s

End Function

Function searchFile(filename As String) As String
    'run a command to find the full path of the file passed in parameters
    'be careful: the full file extension is needed
    'for instance, if needed to search a csv, filename must be "filename.csv"
    
    Dim cmd As String
    cmd = "where /r %HOMEPATH%\ " & filename
    
    Dim path As String
    path = ShellRun(cmd)
    
    Dim trimmed_path As String
    'trimmed_path = Left(path, Len(path) - 2) 'we remove the last 2 characters that are equivalent to \n
    trimmed_path = Split(path, Chr(13) & Chr(10))(0)
    
    searchFile = trimmed_path

End Function

Function get_file_size_from_where_cmd(shell_return As String) As Long

    If Len(shell_return) > 0 Then
        Dim index_fst_space_char As Integer
        index_fst_space_char = 1
        
        While Asc(Mid(shell_return, index_fst_space_char, 1)) <> 32 And index_fst_space_char < Len(shell_return)
            index_fst_space_char = index_fst_space_char + 1
        Wend
        
        Dim string_file_size As String
        string_file_size = Left(shell_return, index_fst_space_char)
        
        get_file_size_from_where_cmd = CLng(Trim(string_file_size))
    End If
End Function

Function get_path_from_where_cmd(shell_return As String) As String

    If Len(shell_return) > 0 Then
        Dim index_fst_space_char As Integer
        index_fst_space_char = Len(shell_return)
        While Asc(Mid(shell_return, index_fst_space_char, 1)) <> 32 And index_fst_space_char > 0
            index_fst_space_char = index_fst_space_char - 1
        Wend
        
        Dim path As String
        path = Right(shell_return, Len(shell_return) - index_fst_space_char)
        
        get_path_from_where_cmd = Trim(path)
    End If
End Function

Function get_python_path() As String

    'returns the python to python.exe

    Dim cmd As String
    cmd = "where /t python"
    
    Dim shell_return As String
    shell_return = ShellRun(cmd)
    
    Dim executables() As String
    executables = Split(shell_return, Chr(13) & Chr(10))
    
    Dim largest_executable_size As Long
    largest_executable_size = 0
    Dim path_largest_executable As String
    path_largest_executable = "no path"
    
    Dim executable_size As Long
    Dim executable_path As String
    Dim executable As Variant
    For Each executable In executables
        If Len(Trim(executable)) > 0 Then
            executable_size = get_file_size_from_where_cmd(Trim(executable))
            executable_path = get_path_from_where_cmd(Trim(executable))
            
            If executable_size > largest_executable_size Then
                largest_executable_size = executable_size
                path_largest_executable = executable_path
            End If
        End If
    Next executable
    
    If StrComp(path_largest_executable, "no path") = 0 Then
        MsgBox ("python.exe could not be found, please install python, or if already done, write the path to python.exe in the variable python_executable_path")
    End If
    
    get_python_path = path_largest_executable
End Function

Sub test_script()

    Dim ws As Object
    Set ws = ActiveWorkbook.Sheets.Add(After:= _
             ActiveWorkbook.Sheets(ActiveWorkbook.Sheets.Count))
    
    ws.Activate
    With ws
    
        Dim link_to_reference_companies As String
        Range("A1").value = "path to reference_file_companies.xlsx @"
        On Error Resume Next
        link_to_reference_companies = searchFile("reference_file_companies.xlsx")
        Range("B1").value = link_to_reference_companies
        
        Dim link_to_reference_countries As String
        Range("A2").value = "path to reference_file_countries.xlsx @"
        On Error Resume Next
        link_to_reference_countries = searchFile("reference_file_countries.xlsx")
        Range("B2").value = link_to_reference_countries
        
        Dim link_to_reference_jobs As String
        Range("A3").value = "path to reference_file_jobs.xlsx @"
        On Error Resume Next
        link_to_reference_jobs = searchFile("reference_file_jobs.xlsx")
        Range("B3").value = link_to_reference_jobs
        
        Dim python_executable_path As String
        Range("A4").value = "path to python.exe @"
        On Error Resume Next
        python_executable_path = get_python_path()
        Range("B4").value = python_executable_path
        
        Dim python_script_path As String
        Range("A5").value = "path to runModel.py @"
        On Error Resume Next
        python_script_path = searchFile("runModel.py")
        Range("B5").value = python_script_path
        
        Dim model_path As String
        Range("A6").value = "path to model.h5 @"
        On Error Resume Next
        model_path = searchFile("model.h5")
        Range("B6").value = model_path
        
        Dim classes_path As String
        Range("A7").value = "path to classes_for_job_functions_classification @"
        On Error Resume Next
        classes_path = searchFile("classes_for_job_functions_classification")
        Range("B7").value = classes_path
        

    End With
End Sub

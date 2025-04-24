Attribute VB_Name = "FormControl"
' ----- Version -----
'        1.0.0
' -------------------

Sub SaveData(Optional ShowOnMacroList As Boolean = False)
    
    Dim colMap As Object
    Set colMap = GetColumnHeadersMapping()
    
    Dim wsForm As Worksheet, wsDados As Worksheet
    Dim dadosTable As ListObject
    Dim tblRow As ListRow
    Dim newID As String
    Dim userResponse As VbMsgBoxResult
    
    ' Set worksheet reference
    Set wsForm = ThisWorkbook.Sheets("Formulário")
    Set wsDados = ThisWorkbook.Sheets("Dados")
    
    ' Check if table "Dados" exists
    On Error Resume Next
    Set dadosTable = wsDados.ListObjects("Dados")
    On Error GoTo 0
    
    ' If the table doesn't exist, exit sub
    If dadosTable Is Nothing Then
        MsgBox "Tabela 'Dados' não encontrada!", vbExclamation
        Exit Sub
    End If
    
    newID = wsForm.OLEObjects("ComboBoxID").Object.Value
    
    ' If ComboBoxID is not empty, prompt the user
    If Trim(newID) <> "" Then
        userResponse = MsgBox("Esse Número da ID já existe. Deseja sobreescrever?", vbYesNoCancel + vbQuestion, "Confirmação")

        Select Case userResponse
            Case vbYes
                newID = Val(newID) ' Use ComboBoxID.Value as new ID
                ' Search for the ID in the first column of the table
                Set tblRow = dadosTable.ListRows(dadosTable.ListColumns(1).DataBodyRange.Find(What:=newID, LookAt:=xlWhole).Row - dadosTable.DataBodyRange.Row + 1)
            Case vbNo
                ' Proceed with generating new ID
                newID = Application.WorksheetFunction.Max(dadosTable.ListColumns(1).DataBodyRange) + 1
                wsForm.OLEObjects("ComboBoxID").Object.Value = newID
                ' Add a new row to the table
                Set tblRow = dadosTable.ListRows.Add
            Case vbCancel
                MsgBox "Os dados não foram salvos", vbInformation
                Exit Sub ' Exit without saving
        End Select
    Else
        newID = wsForm.Range("F10").Value
        
        wsForm.OLEObjects("ComboBoxID").Object.Value = newID
        
        wsForm.OLEObjects("ComboBoxName").Object.Value = wsForm.Range("F10").Value & " - " & wsForm.Range("B14").Value & " - " & wsForm.Range("B6").Value & " - " & wsForm.Range("D6").Value
        
        ' Add a new row to the table
        Set tblRow = dadosTable.ListRows.Add
    End If
    
    ' Assign values to the new row
    With tblRow.Range
        ' Set new ID
        .Cells(1, colMap("ID")).Value = newID ' First column value
        
        ' Read column B values
        .Cells(1, colMap("Projeto")).Value = wsForm.Range("B6").Value
        .Cells(1, colMap("Cliente")).Value = wsForm.Range("B10").Value
        .Cells(1, colMap("PM")).Value = wsForm.Range("B14").Value
        .Cells(1, colMap("PEP")).Value = wsForm.Range("B18").Value
        
        ' Read column D values
        .Cells(1, colMap("Tipo")).Value = wsForm.Range("D6").Value
        .Cells(1, colMap("Valor Total")).Value = wsForm.Range("D10").Value
        .Cells(1, colMap("Custo")).Value = wsForm.Range("D14").Value
        .Cells(1, colMap("Apolice")).Value = wsForm.Range("D18").Value
        .Cells(1, colMap("Percentual")).Value = wsForm.Range("D22").Value
        .Cells(1, colMap("Inicio Vigencia")).Value = wsForm.Range("D26").Value
        .Cells(1, colMap("Fim Vigencia")).Value = wsForm.Range("D30").Value
        
        ' Read column F values
        .Cells(1, colMap("Status")).Value = wsForm.Range("F6").Value
    End With
    
    ' MsgBox "Dados salvos com sucesso!", vbInformation
End Sub

Sub RetrieveDataFromName(Optional ShowOnMacroList As Boolean = False)
    
    Dim colMap As Object
    Set colMap = GetColumnHeadersMapping()
    
    Dim wsForm As Worksheet, wsDados As Worksheet
    Dim dadosTable As ListObject
    Dim foundRow As Range
    Dim searchName As String
    
    ' Set worksheet reference
    Set wsForm = ThisWorkbook.Sheets("Formulário")
    Set wsDados = ThisWorkbook.Sheets("Dados")
    
    ' Check if table "Dados" exists
    On Error Resume Next
    Set dadosTable = wsDados.ListObjects("Dados")
    On Error GoTo 0
    
    ' If the table doesn't exist, exit sub
    If dadosTable Is Nothing Then
        MsgBox "Tabela 'Dados' não encontrada!", vbExclamation
        Exit Sub
    End If
    
    wsForm.OLEObjects("ComboBoxName").Top = wsForm.OLEObjects("ComboBoxID").Top + 38
    wsForm.OLEObjects("ComboBoxName").Left = wsForm.OLEObjects("ComboBoxID").Left
    
    ' Get the ID to search from ComboBox
    If wsForm.OLEObjects("ComboBoxName").Object.Value <> "" Then
        searchName = wsForm.OLEObjects("ComboBoxName").Object.Value
    Else
        'ClearForm
        Exit Sub
    End If
    
    ' Search for the matching row
    Set foundRow = Nothing
    For Each cell In dadosTable.ListColumns(colMap("ID")).DataBodyRange
        If cell.Value & " - " & cell.Cells(1, colMap("PEP")) & " - " & cell.Cells(1, colMap("Cliente")) & " - " & cell.Cells(1, colMap("Tipo")) = searchName Then
            Set foundRow = cell
            Exit For
        End If
    Next cell
    
    ' If Name is not found, exit sub
    If foundRow Is Nothing Then
        MsgBox "Nenhum dado encontrado!", vbExclamation
        Exit Sub
    End If
    
    ' Populate worksheet with retrieved data
    With wsForm
        wsForm.OLEObjects("ComboBoxID").Object.Value = foundRow.Cells(1, colMap("ID"))
        wsForm.OLEObjects("ComboBoxName").Object.Value = foundRow.Cells(1, colMap("ID")) & " - " & foundRow.Cells(1, colMap("PEP")) & " - " & foundRow.Cells(1, colMap("Cliente")) & " - " & foundRow.Cells(1, colMap("Tipo"))
        
        ' Read column B values
        .Range("B6").Value = foundRow.Cells(1, colMap("Projeto")).Value
        .Range("B10").Value = foundRow.Cells(1, colMap("Cliente")).Value
        .Range("B14").Value = foundRow.Cells(1, colMap("PM")).Value
        .Range("B18").Value = foundRow.Cells(1, colMap("PEP")).Value
        
        ' Read column D values
        .Range("D6").Value = foundRow.Cells(1, colMap("Tipo")).Value
        .Range("D10").Value = foundRow.Cells(1, colMap("Valor Total")).Value
        .Range("D14").Value = foundRow.Cells(1, colMap("Custo")).Value
        .Range("D18").Value = foundRow.Cells(1, colMap("Apolice")).Value
        .Range("D22").Value = foundRow.Cells(1, colMap("Percentual")).Value
        .Range("D26").Value = foundRow.Cells(1, colMap("Inicio Vigencia")).Value
        .Range("D30").Value = foundRow.Cells(1, colMap("Fim Vigencia")).Value
        
        ' Read column F values
        .Range("F6").Value = foundRow.Cells(1, colMap("Status")).Value
        .Range("F10").Value = foundRow.Cells(1, colMap("ID")).Value
    End With
End Sub

Sub RetrieveDataFromID(Optional ShowOnMacroList As Boolean = False)

    Dim colMap As Object
    Set colMap = GetColumnHeadersMapping()
    
    Dim wsForm As Worksheet, wsDados As Worksheet
    Dim dadosTable As ListObject
    Dim foundRow As Range
    Dim searchID As Double
    
    ' Set worksheet reference
    Set wsForm = ThisWorkbook.Sheets("Formulário")
    Set wsDados = ThisWorkbook.Sheets("Dados")
    
    ' Check if table "Dados" exists
    On Error Resume Next
    Set dadosTable = wsDados.ListObjects("Dados")
    On Error GoTo 0
    
    ' If the table doesn't exist, exit sub
    If dadosTable Is Nothing Then
        MsgBox "Tabela 'Dados' não encontrada!", vbExclamation
        Exit Sub
    End If
    
    wsForm.OLEObjects("ComboBoxName").Top = wsForm.OLEObjects("ComboBoxID").Top + 38
    wsForm.OLEObjects("ComboBoxName").Left = wsForm.OLEObjects("ComboBoxID").Left
    
    ' Get the ID to search from ComboBox
    If wsForm.OLEObjects("ComboBoxID").Object.Value <> "" Then
        searchID = wsForm.OLEObjects("ComboBoxID").Object.Value
    Else
        'ClearForm
        Exit Sub
    End If
    
    ' Search for the ID in the first column of the table
    Set foundRow = Nothing
    On Error Resume Next
    Set foundRow = dadosTable.ListColumns(colMap("ID")).DataBodyRange.Find(What:=searchID, LookAt:=xlWhole)
    On Error GoTo 0
    
    ' If ID is not found, exit sub
    If foundRow Is Nothing Then
        MsgBox "ID não encontrado!", vbExclamation
        Exit Sub
    End If
    
    ' Populate worksheet with retrieved data
    With wsForm
        wsForm.OLEObjects("ComboBoxName").Object.Value = foundRow.Cells(1, colMap("ID")) & " - " & foundRow.Cells(1, colMap("PEP")) & " - " & foundRow.Cells(1, colMap("Cliente")) & " - " & foundRow.Cells(1, colMap("Tipo"))
        
        ' Read column B values
        .Range("B6").Value = foundRow.Cells(1, colMap("Projeto")).Value
        .Range("B10").Value = foundRow.Cells(1, colMap("Cliente")).Value
        .Range("B14").Value = foundRow.Cells(1, colMap("PM")).Value
        .Range("B18").Value = foundRow.Cells(1, colMap("PEP")).Value
        
        ' Read column D values
        .Range("D6").Value = foundRow.Cells(1, colMap("Tipo")).Value
        .Range("D10").Value = foundRow.Cells(1, colMap("Valor Total")).Value
        .Range("D14").Value = foundRow.Cells(1, colMap("Custo")).Value
        .Range("D18").Value = foundRow.Cells(1, colMap("Apolice")).Value
        .Range("D22").Value = foundRow.Cells(1, colMap("Percentual")).Value
        .Range("D26").Value = foundRow.Cells(1, colMap("Inicio Vigencia")).Value
        .Range("D30").Value = foundRow.Cells(1, colMap("Fim Vigencia")).Value
        
        ' Read column F values
        .Range("F6").Value = foundRow.Cells(1, colMap("Status")).Value
        .Range("F10").Value = foundRow.Cells(1, colMap("ID")).Value
    End With
End Sub

Sub EnviarParaAprovação(Optional ShowOnMacroList As Boolean = False)
    
    Dim colMap As Object
    Set colMap = GetColumnHeadersMapping()
    
    Dim wsForm As Worksheet, wsDados As Worksheet
    Dim dadosTable As ListObject
    
    Dim OutApp As Object
    Dim OutMail As Object
    
    '--- Variables for email content
    Dim HTMLbody As String
    Dim greeting As String
    Dim strSignature As String
    Dim faseObra As String
    
    '--- Create Outlook instance and a new mail item
    On Error Resume Next
    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)
    On Error GoTo 0
    
    If OutApp Is Nothing Then
        MsgBox "O Outlook não está instalado nesse computador.", vbExclamation
        Exit Sub
    End If
    
    ' Set worksheet reference
    Set wsForm = ThisWorkbook.Sheets("Formulário")
    Set wsDados = ThisWorkbook.Sheets("Dados")
    
    ' Check if table "Dados" exists
    On Error Resume Next
    Set dadosTable = wsDados.ListObjects("Dados")
    On Error GoTo 0
    
    ' If the table doesn't exist, exit sub
    If dadosTable Is Nothing Then
        MsgBox "Tabela 'Dados' não encontrada!", vbExclamation
        Exit Sub
    End If
    
    ' Get the ID to search from ComboBox
    searchID = wsForm.OLEObjects("ComboBoxID").Object.Value
    
    ' Stop if data not saved
    If searchID = "" Then
        MsgBox "Desculpe, salve os dados antes de gerar o e-mail", vbInformation, "Atenção"
        Exit Sub
    End If
    
    ' Search for the ID in the first column of the table
    Set foundRow = Nothing
    On Error Resume Next
    Set foundRow = dadosTable.ListColumns(1).DataBodyRange.Find(What:=searchID, LookAt:=xlWhole)
    On Error GoTo 0
    
    ' If ID is not found, exit sub
    If foundRow Is Nothing Then
        MsgBox "ID não encontrado!", vbExclamation
        Exit Sub
    End If
    
    If foundRow.Cells(1, colMap("Solicitação")).Value <> "" Then
        userResponse = MsgBox("O e-mail de aprovação para esses dados já foi enviado em " & foundRow.Cells(1, colMap("Solicitação")).Value & ". Deseja enviar novamente?", vbYesNo)
        If userResponse = vbNo Then
            MsgBox "Envio de e-mail cancelado!", vbInformation
            Exit Sub
        End If
    End If
    
    ' Decide between Bom dia or Boa tarde
    If Hour(Now) < 12 Then
        greeting = "Bom dia"
    Else
        greeting = "Boa tarde"
    End If
    
    ' Get user signature
    With OutMail
        .Display ' This opens the email and loads the default signature
        strSignature = .HTMLbody ' Capture the signature
    End With
    
    HTMLbody = ""
    HTMLbody = HTMLbody & "<p>" & greeting & ", Prezada</p>"
    HTMLbody = HTMLbody & "<p>Gentileza prosseguir com a solicitação de seguro conforme dados abaixo:</p>"
    
    ' First header row
    HTMLbody = HTMLbody & "<table border='1' style='border-collapse: collapse; font-size: 10pt;'>"
    HTMLbody = HTMLbody & "<tr>"
    HTMLbody = HTMLbody & "<th scope=""col"" colspan=""4"">NOME DO CLIENTE</th>"
    HTMLbody = HTMLbody & "<th scope=""col"" colspan=""3"">VALOR TOTAL</th>"
    
    ' Close first header row
    HTMLbody = HTMLbody & "</tr>"
    
    ' Data row
    HTMLbody = HTMLbody & "<tr>"
    HTMLbody = HTMLbody & "<td scope=""col"" colspan=""4"">" & foundRow.Cells(1, colMap("Cliente")).Value & "</td>"
    HTMLbody = HTMLbody & "<td scope=""col"" colspan=""3"">" & Format(foundRow.Cells(1, colMap("Valor Total")).Value, "R$ 0.00") & "</td>"
    HTMLbody = HTMLbody & "</tr>"
    
    ' Second header row
    HTMLbody = HTMLbody & "<tr>"
    HTMLbody = HTMLbody & "<th scope=""col"">Garantia</th>"
    HTMLbody = HTMLbody & "<th scope=""col"">PEP</th>"
    HTMLbody = HTMLbody & "<th scope=""col"">%</th>"
    HTMLbody = HTMLbody & "<th scope=""col"">Valor</th>"
    HTMLbody = HTMLbody & "<th scope=""col"">Início da Vigência</th>"
    HTMLbody = HTMLbody & "<th scope=""col"">Final da Vigência</th>"
    HTMLbody = HTMLbody & "<th scope=""col"">ID</th>"
    HTMLbody = HTMLbody & "</tr>"
    
    ' Data row
    HTMLbody = HTMLbody & "<tr>"
    HTMLbody = HTMLbody & "<td>" & foundRow.Cells(1, colMap("Tipo")).Value & "</td>"
    HTMLbody = HTMLbody & "<td>" & foundRow.Cells(1, colMap("PEP")).Value & "</td>"
    HTMLbody = HTMLbody & "<td>" & Format(foundRow.Cells(1, colMap("Percentual")).Value, "0.00%") & "</td>"
    HTMLbody = HTMLbody & "<td>" & Format(foundRow.Cells(1, colMap("Apolice")).Value, "R$ 0.00") & "</td>"
    HTMLbody = HTMLbody & "<td>" & foundRow.Cells(1, colMap("Inicio Vigencia")).Value & "</td>"
    HTMLbody = HTMLbody & "<td>" & foundRow.Cells(1, colMap("Fim Vigencia")).Value & "</td>"
    HTMLbody = HTMLbody & "<td>" & foundRow.Cells(1, colMap("ID")).Value & "</td>"
    HTMLbody = HTMLbody & "</tr>"
    
    ' Close TABLE
    HTMLbody = HTMLbody & "</table>"
    
    '-------------------------------------------------------------------------
    ' Configure and send the email
    '-------------------------------------------------------------------------
    With OutMail
        .To = "julianarigo@weg.net"
        .CC = ""
        .BCC = ""
        .Subject = "Solicitação de Seguro - " & foundRow.Cells(1, colMap("Projeto")).Value & " - " & foundRow.Cells(1, colMap("Cliente")).Value & " - " & foundRow.Cells(1, colMap("PEP")).Value
        .HTMLbody = HTMLbody & strSignature
        .Display   'Use .Display to just open the email draft
        ' .Send       'Use .Send to send immediately
    End With
    
    '--- Cleanup
    Set OutMail = Nothing
    Set OutApp = Nothing
    
    foundRow.Cells(1, colMap("Solicitação")).Value = Date
    
    MsgBox "Email """ & "Solicitação de Seguro - " & foundRow.Cells(1, colMap("Projeto")).Value & " - " & foundRow.Cells(1, colMap("Cliente")).Value & " - " & foundRow.Cells(1, colMap("PEP")).Value & """ enviado com sucesso!", vbInformation
    
End Sub

Sub ClearForm(Optional ShowOnMacroList As Boolean = False)
    
    Dim wsForm As Worksheet
    
    ' Set worksheet reference
    Set wsForm = ThisWorkbook.Sheets("Formulário")
    
    If wsForm.OLEObjects("ComboBoxID").Object.Value = "" Then
        If MsgBox("Esses dados não foram salvos. Deseja limpá-los mesmo assim?", vbYesNo) = vbNo Then
            Exit Sub
        End If
    End If
    
    ' Populate worksheet with retrieved data
    With wsForm
        .OLEObjects("ComboBoxID").Object.Value = ""
        .OLEObjects("ComboBoxName").Object.Value = ""
        .OLEObjects("ComboBoxName").Width = 123
        
        ' Read column B values
        .Range("B6").Value = ""
        .Range("B10").Value = ""
        .Range("B14").Value = ""
        .Range("B18").Value = ""
        
        ' Read column D values
        .Range("D6").Value = ""
        .Range("D10").Value = ""
        .Range("D14").Value = ""
        .Range("D18").Value = ""
        .Range("D22").Value = ""
        .Range("D26").Value = ""
        .Range("D30").Value = ""
        
        ' Read column F values
        .Range("F6").Value = ""
        .Range("F10").Value = ""

    End With
End Sub

Public Function GetColumnHeadersMapping() As Object
    Dim headers As Object
    Set headers = CreateObject("Scripting.Dictionary")
    
    ' Add each header from the provided table to the dictionary,
    ' mapping it to its column position.
    headers.Add "ID", 1
    headers.Add "Projeto", 2
    headers.Add "Cliente", 3
    headers.Add "PM", 4
    headers.Add "PEP", 5
    headers.Add "Tipo", 6
    headers.Add "Valor Total", 7
    headers.Add "Custo", 8
    headers.Add "Apolice", 9
    headers.Add "Percentual", 10
    headers.Add "Inicio Vigencia", 11
    headers.Add "Fim Vigencia", 12
    headers.Add "Status", 13
    headers.Add "Solicitação", 14
    
    Set GetColumnHeadersMapping = headers
End Function

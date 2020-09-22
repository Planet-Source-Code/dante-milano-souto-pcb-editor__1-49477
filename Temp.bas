Attribute VB_Name = "Temp"
'Private BASED As Workspace
'Private DBase As Database
'Private TableCostumer As Recordset
'
'Private Sub Class_Initialize()
'
'    Set BASED = DBEngine.Workspaces(0)
'    Set DBase = BASED.OpenDatabase(App.Path & "\Base.mdb", True)
'    Set TableCostumer = DBase.OpenRecordset("TableCostumer", dbOpenDynaset, dbOptimistic )
'
'    If NumRegs > 0 Then
'        TableCostumer.MoveFirst
'    End If
'
'End Sub
'
'Public Function GetItem(iIndex As Long) As Costumers
''
''   LÊ UM DETERMINADO ITEM
''
'
'    Set GetItem = New Costumers
'
'    TableCostumer.MoveFirst
'    TableCostumer.Move iIndex
'
'    GetItem.ID = ChkNull(TableCostumer("ID"))
'    GetItem.Nome = ChkNull(TableCostumer("Nome"))
'    GetItem.Rua = ChkNull(TableCostumer("Rua"))
'    GetItem.Bairro = ChkNull(TableCostumer("Bairro"))
'    GetItem.Complemento = ChkNull(TableCostumer("Complemento"))
'    GetItem.CEP = ChkNull(TableCostumer("CEP"))
'    GetItem.Cidade = ChkNull(TableCostumer("Cidade"))
'    GetItem.UF = ChkNull(TableCostumer("UF"))
'    GetItem.Contato = ChkNull(TableCostumer("Contato"))
'    GetItem.Cargo = ChkNull(TableCostumer("Cargo"))
'    GetItem.Tel = ChkNull(TableCostumer("Tel"))
'    GetItem.Cel = ChkNull(TableCostumer("Cel"))
'    GetItem.Fax = ChkNull(TableCostumer("Fax"))
'    GetItem.Email = ChkNull(TableCostumer("Email"))
'    GetItem.Http = ChkNull(TableCostumer("Http"))
'
'End Function
'
'Public Function FindItem(vtCodigo As String) As String
'
'Dim i As Long
'Dim X As Long
'Dim strTMP As String
'Dim strFind As String
'
'    strTMP = ""
'    strFind = vtCodigo
'
'    With TableCostumer
'
'        .MoveFirst
'        .FindFirst strFind
'
'        Do
'
'            If .NoMatch Then
'                Exit Do
'            End If
'
'            ' Armazena o indicador do registro atual.
'            strTMP = strTMP & .AbsolutePosition & "|"
'            .FindNext strFind
'
'        Loop
'
'    End With
'
'    FindItem = strTMP
'
'End Function
'
'Public Function Adicionar(Item As Costumers) As Integer
'
'On Error GoTo errTrat
'
'Dim mg As Integer
'Dim iCursor As Long
'
'    If VerifiqueExistencia(Item) = True Then
'        mg = MsgBox("Registro Já Cadastrado, Você deseja criar outro com o mesmo nome?", vbYesNoCancel, "Base de Dados")
'        Select Case mg
'            Case Is = vbNo
'                Adicionar = mg
'                Exit Function
'            Case Is = vbCancel
'                Adicionar = vbCancel
'                Exit Function
'        End Select
'    End If
'
'    iCursor = NumRegs
'    With TableCostumer
'        .MoveFirst
'        .Move iCursor
'        .AddNew
'
'        !ID = iCursor + 1
'        !Nome = Item.Nome
'        '!Complemento = Item.Complemento
'        '!Bairro = Item.Bairro
'        '!Cidade = Item.Cidade
'        '!CEP = Item.CEP
'        '!UF = Item.UF
'        '!Tel = Item.Tel
'        '!Cel = Item.Cel
'        '!Fax = Item.Fax
'        '!Email = Item.Email
'        '!Http = Item.Http
'        '!Contato = Item.Contato
'
'        .Update
'    End With
'
'    Adicionar = vbYes
'
'Exit Function
'errTrat:
'
'    MsgBox MsgBox("Falha ao incluir registro", vbCritical, "Base de Dados")
'    Adicionar = vbCancel
'
'End Function
'
'Public Sub Atualizar(Index As Long, Item As Costumers)
'
'On Error GoTo errTrat
'
'    With TableCostumer
'        .MoveFirst
'        .Move Index
'        .Edit
'
'        '!ID = Item.ID          (NÃO EDITAVEL)
'        !Nome = Item.Nome
'        !Rua = Item.Rua
'        !Complemento = Item.Complemento
'        !Bairro = Item.Bairro
'        !Cidade = Item.Cidade
'        !CEP = Item.CEP
'        !UF = Item.UF
'        !Tel = Item.Tel
'        !Cel = Item.Cel
'        !Fax = Item.Fax
'        !Email = Item.Email
'        !Http = Item.Http
'        !Contato = Item.Contato
'
'        .Update
'    End With
'
'Exit Sub
'errTrat:
'
'    MsgBox "Falha ao incluir registro", vbCritical, "Base de Dados"
'
'End Sub
'
'Public Function RemoveItem(iIndex As Long) As Boolean
'
'On Error GoTo errTrat
'
'Dim i As Long
'
'    With TableCostumer
'        .MoveFirst
'        .Move iIndex
'        .Delete
'
'        .MoveFirst
'        .Move i
'        Do
'            If TableCostumer.EOF = True Then Exit Do
'            .Edit
'            !ID = Str(i)
'            .Update
'            i = i + 1
'            .MoveNext
'        Loop
'
'    End With
'    RemoveItem = True
'
'Exit Function
'errTrat:
'
'    RemoveItem = False
'
'End Function
'
'Private Function VerifiqueExistencia(Item As Costumers) As Boolean
'
'Dim ItPoint As Long
'Dim strFind As String
'
'    ItFound = 0
'    VerifiqueExistencia = False
'
'    strFind = "Nome = '" & Trim(Item.Nome) & "'"
'
'    With TableCostumer
'
'        .MoveFirst
'        .FindFirst strFind
'
'        If .NoMatch Then
'            Exit Function
'        End If
'
'        VerifiqueExistencia = True
'
'    End With
'
'End Function
'
'Public Property Get NumRegs() As Long
'
'On Error GoTo errTrat
'
'Dim i As Long
'
'    TableCostumer.MoveFirst
'    Do
'        If TableCostumer.EOF = True Then Exit Do
'        'TableCostumer.Edit
'        'TableCostumer("CUSTO") = (TableCostumer("PRECO") / 100) * 80
'        'TableCostumer.Update
'        i = i + 1
'        TableCostumer.MoveNext
'    Loop
'    NumRegs = i - 1
'
'Exit Property
'errTrat:
'
'    NumRegs = -1
'
'End Property
'
'Function ChkNull(inptstr)
'
'    If IsNull(inptstr) Then
'        ChkNull = " "
'    Else
'        ChkNull = inptstr
'    End If
'
'End Function
'
'Private Sub Class_Terminate()
'
'    TableCostumer.Close
'
'End Sub
'
'Public Function GetIdOfThisItem(Item As Costumers) As Long
'
'Dim ItPoint As Long
'Dim strFind As String
'
'    ItFound = 0
'    GetIdOfThisItem = 0
'
'    strFind = "Nome = '" & Trim(Item.Nome) & "'"
'
'    With TableCostumer
'
'        .MoveFirst
'        .FindFirst strFind
'
'        If .NoMatch Then
'            Exit Function
'        End If
'
'        GetIdOfThisItem = .AbsolutePosition + 1
'
'    End With
'
'End Function






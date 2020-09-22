VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSaveOpen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private DBase       As Database
Private tbCorner    As Recordset
Private tbIlha      As Recordset
Private tbTraço     As Recordset
Private tbGroup     As Recordset
Private tbComponente As Recordset

Private mFileName As String


Public Sub SalvarComponente(mColPCB As Collection, mCompo As clsComponente)
'
'   Salva apenas o grupo selecionado
'
Dim Lft As Single
Dim Top As Single
Dim TMPLft As Single
Dim TMPTop As Single
Dim ID As Long
Dim mColCompo As New Collection
Dim vntObject As Variant
    
    ' Cria os Componentes a partir do grupo especificado
    For Each vntObject In mColPCB
        If vntObject.Grupo = mCompo.ID Then
            mColCompo.Add vntObject
        End If
    Next
        
    Lft = 3.402823E+38
    Top = 3.402823E+38
    ' localizando o mais a esquerda e acima
    For Each vntObject In mColCompo
        If TypeOf vntObject Is clsCorner Then
            TMPLft = vntObject.X
            If TMPLft < Lft Then
                Lft = TMPLft
            End If
            TMPTop = vntObject.Y
            If TMPTop < Top Then
                Top = TMPTop
            End If
        End If
    Next
    
    ' reposicionando todos os conrers
    For Each vntObject In mColCompo
        If TypeOf vntObject Is clsCorner Then
            vntObject.X = vntObject.X - Lft
            vntObject.Y = vntObject.Y - Top
        End If
    Next
    
    ' Abrindo ao Catalogo de Componentes
    Set DBase = OpenDatabase(App.Path & "\componentes.dat", True)
    Set tbCorner = DBase.OpenRecordset("CORNER")
    Set tbIlha = DBase.OpenRecordset("ILHA")
    Set tbTraço = DBase.OpenRecordset("TRAÇO")
    Set tbComponente = DBase.OpenRecordset("COMPONENTE")
    
    ' Procurando pelo ID disponível
    ID = GetFrreId(tbComponente)
    
    ' Colocando o novo ID
    mCompo.ID = ID
    mCompo.DBID = ID
    For Each vntObject In mColCompo
        vntObject.Grupo = -1    'removendo os agrupamentos
        vntObject.Componente = ID
    Next
    
    ' Colocando os elementos nas tabelas
    Call AddComponente(mCompo)
    For Each vntObject In mColCompo
        If TypeOf vntObject Is clsCorner Then
            Call AddCorner(vntObject)
        ElseIf TypeOf vntObject Is clsIlha Then
            Call AddIlha(vntObject)
        ElseIf TypeOf vntObject Is clsTraço Then
            Call AddTraço(vntObject)
        End If
    Next
    
    ' desfazendo o reposicionando todos os conrers
    For Each vntObject In mColCompo
        If TypeOf vntObject Is clsCorner Then
            vntObject.X = vntObject.X + Lft
            vntObject.Y = vntObject.Y + Top
        End If
    Next
    
    ' Fechando a tabela
    tbTraço.Close
    tbIlha.Close
    tbCorner.Close
    tbComponente.Close
    DBase.Close
    
    Set tbComponente = Nothing
    Set tbTraço = Nothing
    Set tbIlha = Nothing
    Set tbCorner = Nothing
    Set DBase = Nothing
    
End Sub

Public Sub SalvarHPLG2(mColPCB As Collection)
    
Dim vntObject  As Variant

    For Each vntObject In mColPCB
        If TypeOf vntObject Is clsCorner Then
            'Call AddCorner(vntObject)
        ElseIf TypeOf vntObject Is clsIlha Then
            'Call AddIlha(vntObject)
        ElseIf TypeOf vntObject Is clsTraço Then
            'Call AddTraço(vntObject)
        End If
    Next
    
End Sub

Public Sub SalvarMDB(mColPCB As Collection, mColGrupo As Collection, mColComponente As Collection)

Dim vntObject As Variant

    If Dir(App.Path & "\modelo.dat") = "" Then
        ' arquivo faltando
        GoTo errTrat
    Else
        If Dir(mFileName) <> "" Then
            Kill mFileName
        End If
        FileCopy App.Path & "\modelo.dat", mFileName
    End If
    
    Call Conectar
    
    For Each vntObject In mColPCB
        If TypeOf vntObject Is clsCorner Then
            Call AddCorner(vntObject)
        ElseIf TypeOf vntObject Is clsIlha Then
            Call AddIlha(vntObject)
        ElseIf TypeOf vntObject Is clsTraço Then
            Call AddTraço(vntObject)
        End If
    Next
    
    For Each vntObject In mColGrupo
        Call AddGrupo(vntObject)
    Next
    
    For Each vntObject In mColComponente
        Call AddComponente(vntObject)
    Next
    
    Call Desconectar
    
errTrat:

End Sub

Public Function Abrir() As Collection

Dim cCorner As clsCorner
Dim cTraço As clsTraço
Dim cIlha As clsIlha
Dim cGrupo As clsGrupo
Dim cCompo As clsComponente
Dim tmpPCB As New Collection
Dim tmpGRUPO As New Collection
Dim tmpComponentes As New Collection
Dim tmpALL As New Collection
    
    Call Conectar
    
    With tbCorner
        If .RecordCount > 0 Then
            .MoveLast
            .MoveFirst
            Do Until .EOF = True
                Set cCorner = New clsCorner
                cCorner.ID = !ID
                cCorner.Layer = !Layer
                cCorner.X = !X
                cCorner.Y = !Y
                cCorner.Grupo = !Grupo
                cCorner.Componente = !Componente
                tmpPCB.Add cCorner, "Corner-" & cCorner.ID
                .MoveNext
            Loop
        End If
    End With
        
    With tbIlha
        If .RecordCount > 0 Then
            .MoveLast
            .MoveFirst
            Do Until .EOF = True
                Set cIlha = New clsIlha
                cIlha.ID = !ID
                cIlha.Layer = !Layer
                cIlha.Corner = !Corner
                cIlha.IlhaLargura = !Largura
                cIlha.IlhaFuro = !Furo
                cIlha.Grupo = !Grupo
                cIlha.Componente = !Componente
                tmpPCB.Add cIlha, "Ilha-" & cIlha.ID
                .MoveNext
            Loop
        End If
    End With
    
    With tbTraço
        If .RecordCount > 0 Then
            .MoveLast
            .MoveFirst
            Do Until .EOF = True
                Set cTraço = New clsTraço
                cTraço.ID = !ID
                cTraço.Layer = !Layer
                cTraço.Largura = !Largura
                cTraço.StartCorner = !StartCorner
                cTraço.EndCorner = !EndCorner
                cTraço.Grupo = !Grupo
                cTraço.Componente = !Componente
                tmpPCB.Add cTraço, "Traço-" & cTraço.ID
                .MoveNext
            Loop
        End If
    End With
    
    With tbGroup
        If .RecordCount > 0 Then
            .MoveLast
            .MoveFirst
            Do Until .EOF = True
                Set cGrupo = New clsGrupo
                cGrupo.ID = !ID
                cGrupo.Nome = ChkNull(!Nome)
                tmpGRUPO.Add cGrupo, "Grupo-" & cGrupo.ID
                .MoveNext
            Loop
        End If
    End With
    
    With tbComponente
        If .RecordCount > 0 Then
            .MoveLast
            .MoveFirst
            Do Until .EOF = True
                Set cCompo = New clsComponente
                cCompo.ID = !ID
                cCompo.DBID = !DBID
                cCompo.Nome = ChkNull(!Nome)
                tmpComponentes.Add cCompo, "Componente-" & cCompo.ID
                .MoveNext
            Loop
        End If
    End With
    
    tmpALL.Add tmpPCB, "PCB"
    tmpALL.Add tmpGRUPO, "GRUPO"
    tmpALL.Add tmpComponentes, "COMPO"
    Set Abrir = tmpALL
    
    Call Desconectar
    
End Function

Public Function AbrirComponente(CompID As Long) As Collection

Dim cCorner As clsCorner
Dim cIlha As clsIlha
Dim cTraço As clsTraço
Dim tmpPCB As New Collection

    
    ' Abrindo ao Catalogo de Componentes
    Set DBase = OpenDatabase(App.Path & "\componentes.dat", True)
    Set tbCorner = DBase.OpenRecordset("CORNER")
    Set tbIlha = DBase.OpenRecordset("ILHA")
    Set tbTraço = DBase.OpenRecordset("TRAÇO")
    Set tbComponente = DBase.OpenRecordset("COMPONENTE")
    
    With tbCorner
        If .RecordCount > 0 Then
            .MoveLast
            .MoveFirst
            Do Until .EOF = True
                If !Componente = CompID Then
                    Set cCorner = New clsCorner
                    cCorner.ID = !ID
                    cCorner.Layer = !Layer
                    cCorner.X = !X
                    cCorner.Y = !Y
                    cCorner.Grupo = !Grupo
                    cCorner.Componente = !Componente
                    tmpPCB.Add cCorner, "Corner-" & cCorner.ID
                End If
                .MoveNext
            Loop
        End If
    End With
        
    With tbIlha
        If .RecordCount > 0 Then
            .MoveLast
            .MoveFirst
            Do Until .EOF = True
                If !Componente = CompID Then
                    Set cIlha = New clsIlha
                    cIlha.ID = !ID
                    cIlha.Layer = !Layer
                    cIlha.Corner = !Corner
                    cIlha.IlhaLargura = !Largura
                    cIlha.IlhaFuro = !Furo
                    cIlha.Grupo = !Grupo
                    cIlha.Componente = !Componente
                    tmpPCB.Add cIlha, "Ilha-" & cIlha.ID
                End If
                .MoveNext
            Loop
        End If
    End With
    
    With tbTraço
        If .RecordCount > 0 Then
            .MoveLast
            .MoveFirst
            Do Until .EOF = True
                If !Componente = CompID Then
                    Set cTraço = New clsTraço
                    cTraço.ID = !ID
                    cTraço.Layer = !Layer
                    cTraço.Largura = !Largura
                    cTraço.StartCorner = !StartCorner
                    cTraço.EndCorner = !EndCorner
                    cTraço.Grupo = !Grupo
                    cTraço.Componente = !Componente
                    tmpPCB.Add cTraço, "Traço-" & cTraço.ID
                End If
                .MoveNext
            Loop
        End If
    End With
    
    Set AbrirComponente = tmpPCB
        
    ' Fechando a tabela
    tbTraço.Close
    tbIlha.Close
    tbCorner.Close
    tbComponente.Close
    DBase.Close
    
    Set tbComponente = Nothing
    Set tbTraço = Nothing
    Set tbIlha = Nothing
    Set tbCorner = Nothing
    Set DBase = Nothing
    
End Function

Private Sub Conectar()
On Error GoTo errTrat

    Set DBase = OpenDatabase(FileName, True)
    Set tbCorner = DBase.OpenRecordset("CORNER")
    Set tbIlha = DBase.OpenRecordset("ILHA")
    Set tbTraço = DBase.OpenRecordset("TRAÇO")
    Set tbGroup = DBase.OpenRecordset("GRUPO")
    Set tbComponente = DBase.OpenRecordset("COMPONENTE")
    
Exit Sub
errTrat:

    MsgBox " Falhou a conexão com o arquivo "
    
End Sub

Private Sub Desconectar()

    tbTraço.Close
    tbIlha.Close
    tbCorner.Close
    tbComponente.Close
    tbGroup.Close
    DBase.Close
    
    Set tbComponente = Nothing
    Set tbGroup = Nothing
    Set tbTraço = Nothing
    Set tbIlha = Nothing
    Set tbCorner = Nothing
    Set DBase = Nothing
    
End Sub

Private Sub AddGrupo(mGrupo As Variant)
    
    With tbGroup
        .AddNew
        !ID = mGrupo.ID
        !Nome = mGrupo.Nome
        .Update
    End With
    
End Sub

Private Sub AddComponente(mCompo As Variant)
    
    With tbComponente
        .AddNew
        !ID = mCompo.ID
        !DBID = mCompo.DBID
        !Nome = mCompo.Nome
        .Update
    End With
    
End Sub

Private Sub AddCorner(mCorner As Variant)
    
    With tbCorner
        .AddNew
        !ID = mCorner.ID
        !X = mCorner.X
        !Y = mCorner.Y
        !Layer = mCorner.Layer
        !Grupo = mCorner.Grupo
        !Componente = mCorner.Componente
        .Update
    End With
    
End Sub

Private Sub AddIlha(mIlha As Variant)

    With tbIlha
        .AddNew
        !ID = mIlha.ID
        !Corner = mIlha.Corner
        !Layer = mIlha.Layer
        !Largura = mIlha.IlhaLargura
        !Furo = mIlha.IlhaFuro
        !Grupo = mIlha.Grupo
        !Componente = mIlha.Componente
        .Update
    End With
    
End Sub

Private Sub AddTraço(mTraço As Variant)

    With tbTraço
        .AddNew
        !ID = mTraço.ID
        !Layer = mTraço.Layer
        !Largura = mTraço.Largura
        !StartCorner = mTraço.StartCorner
        !EndCorner = mTraço.EndCorner
        !Grupo = mTraço.Grupo
        !Componente = mTraço.Componente
        .Update
    End With
    
End Sub

Private Function GetFrreId(Tabela As Recordset) As Long
'
'   Primeiro ID livre de uma tabela
'

Dim ub As Long
Dim I As Long
Dim Rc As Long

    Rc = 0
    ub = Tabela.RecordCount
    If ub > 0 Then
        With Tabela
            .MoveLast
            .MoveFirst
            Do Until .EOF = True
                If !ID >= Rc Then
                    Rc = !ID + 1
                End If
                .MoveNext
            Loop
        End With
    End If
    
    GetFrreId = Rc
    
End Function

Public Property Get FileName() As String

    FileName = mFileName
    
End Property

Public Property Let FileName(NewValue As String)
    
    If Len(NewValue) < 5 Then
        Err.Raise 1, , "Bad File Name"
    Else
        mFileName = NewValue
    End If
    
End Property



Attribute VB_Name = "Mod_Protocol"
Option Explicit

Private Const SEPARATOR As String * 1 = vbNullChar

Private Type tFont
    Red As Byte
    Green As Byte
    Blue As Byte
    bold As Boolean
    italic As Boolean
End Type

Private Enum ServerPacketID
    logged                  '   LOGGED
    RemoveDialogs           '   QTDL
    RemoveCharDialog        '   QDL
    NavigateToggle          '   NAVEG
    Disconnect              '   FINOK
    CommerceEnd             '   FINCOMOK
    BankEnd                 '   FINBANOK
    CommerceInit            '   INITCOM
    BankInit                '   INITBANCO
    UserCommerceInit        '   INITCOMUSU
    UserCommerceEnd         '   FINCOMUSUOK
    UserOfferConfirm
    CommerceChat
    ShowBlacksmithForm      '   SFH
    ShowCarpenterForm       '   SFC
    UpdateSta               '   ASS
    UpdateMana              '   ASM
    UpdateHP                '   ASH
    UpdateGold              '   ASG
    UpdateBankGold
    UpdateExp               '   ASE
    ChangeMap               '   CM
    PosUpdate               '   PU
    ChatOverHead            '   ||
    ConsoleMsg              '   || - Beware!! its the same as above, but it was properly splitted
    GuildChat               '   |+
    ShowMessageBox          '   !!
    UserIndexInServer       '   IU
    UserCharIndexInServer   '   IP
    CharacterCreate         '   CC
    CharacterRemove         '   BP
    CharacterChangeNick
    CharacterMove           '   MP, +, * and _ '
    ForceCharMove
    CharacterChange         '   CP
    ObjectCreate            '   HO
    ObjectDelete            '   BO
    BlockPosition           '   BQ
    PlayMIDI                '   TM
    PlayWave                '   TW
    guildList               '   GL
    AreaChanged             '   CA
    PauseToggle             '   BKW
    RainToggle              '   LLU
    CreateFX                '   CFX
    UpdateUserStats         '   EST
    WorkRequestTarget       '   T01
    ChangeInventorySlot     '   CSI
    ChangeBankSlot          '   SBO
    ChangeSpellSlot         '   SHS
    Atributes               '   ATR
    BlacksmithWeapons       '   LAH
    BlacksmithArmors        '   LAR
    CarpenterObjects        '   OBR
    RestOK                  '   DOK
    ErrorMsg                '   ERR
    Blind                   '   CEGU
    Dumb                    '   DUMB
    ShowSignal              '   MCAR
    ChangeNPCInventorySlot  '   NPCI
    UpdateHungerAndThirst   '   EHYS
    Fame                    '   FAMA
    MiniStats               '   MEST
    LevelUp                 '   SUNI
    AddForumMsg             '   FMSG
    ShowForumForm           '   MFOR
    SetInvisible            '   NOVER
    DiceRoll                '   DADOS
    MeditateToggle          '   MEDOK
    BlindNoMore             '   NSEGUE
    DumbNoMore              '   NESTUP
    SendSkills              '   SKILLS
    TrainerCreatureList     '   LSTCRI
    GuildNews               '   GUILDNE
    OfferDetails            '   PEACEDE & ALLIEDE
    AlianceProposalsList    '   ALLIEPR
    PeaceProposalsList      '   PEACEPR
    CharacterInfo           '   CHRINFO
    GuildLeaderInfo         '   LEADERI
    GuildMemberInfo
    GuildDetails            '   CLANDET
    ShowGuildFundationForm  '   SHOWFUN
    ParalizeOK              '   PARADOK
    ShowUserRequest         '   PETICIO
    TradeOK                 '   TRANSOK
    BankOK                  '   BANCOOK
    ChangeUserTradeSlot     '   COMUSUINV
    SendNight               '   NOC
    Pong
    UpdateTagAndStatus
    
    'GM messages
    SpawnList               '   SPL
    ShowSOSForm             '   MSOS
    ShowMOTDEditionForm     '   ZMOTD
    ShowGMPanelForm         '   ABPANEL
    UserNameList            '   LISTUSU
    
    ShowGuildAlign
    ShowPartyForm
    UpdateStrenghtAndDexterity
    UpdateStrenght
    UpdateDexterity
    AddSlots
    MultiMessage
    StopWorking
    CancelOfferItem
    
    ParaTime
    CuentaSV
    AuraSet
    CreateProjectile
    CreateOrUpdateStats
    UpdateClima
    SendClima
    PartyDetail
    PartyKicked
    PartyExit
    IniciarSubastaConsulta
    DuelStart
    
    CountStart
    Particles
    CentroCanje
    
    ShowStatsBars
    
    QuestDetails
    QuestListSend

End Enum

Private Enum ClientPacketID
    LoginExistingChar       'OLOGIN
    ThrowDices              'TIRDAD
    LoginNewChar            'NLOGIN
    Talk                    ';
    Yell                    '-
    Whisper                 '\
    Walk                    'M
    RequestPositionUpdate   'RPU
    Attack                  'AT
    PickUp                  'AG
    SafeToggle              '/SEG & SEG  (SEG's behaviour has to be coded in the client)
    ResuscitationSafeToggle
    RequestGuildLeaderInfo  'GLINFO
    RequestAtributes        'ATR
    RequestFame             'FAMA
    RequestSkills           'ESKI
    RequestMiniStats        'FEST
    CommerceEnd             'FINCOM
    UserCommerceEnd         'FINCOMUSU
    UserCommerceConfirm
    CommerceChat
    BankEnd                 'FINBAN
    UserCommerceOk          'COMUSUOK
    UserCommerceReject      'COMUSUNO
    Drop                    'TI
    CastSpell               'LH
    LeftClick               'LC
    DoubleClick             'RC
    Work                    'UK
    UseSpellMacro           'UMH
    UseItem                 'USA
    CraftBlacksmith         'CNS
    CraftCarpenter          'CNC
    WorkLeftClick           'WLC
    CreateNewGuild          'CIG
    SpellInfo               'INFS
    EquipItem               'EQUI
    ChangeHeading           'CHEA
    ModifySkills            'SKSE
    Train                   'ENTR
    CommerceBuy             'COMP
    BankExtractItem         'RETI
    CommerceSell            'VEND
    BankDeposit             'DEPO
    ForumPost               'DEMSG
    MoveSpell               'DESPHE
    MoveBank
    ClanCodexUpdate         'DESCOD
    UserCommerceOffer       'OFRECER
    GuildAcceptPeace        'ACEPPEAT
    GuildRejectAlliance     'RECPALIA
    GuildRejectPeace        'RECPPEAT
    GuildAcceptAlliance     'ACEPALIA
    GuildOfferPeace         'PEACEOFF
    GuildOfferAlliance      'ALLIEOFF
    GuildAllianceDetails    'ALLIEDET
    GuildPeaceDetails       'PEACEDET
    GuildRequestJoinerInfo  'ENVCOMEN
    GuildAlliancePropList   'ENVALPRO
    GuildPeacePropList      'ENVPROPP
    GuildDeclareWar         'DECGUERR
    GuildNewWebsite         'NEWWEBSI
    GuildAcceptNewMember    'ACEPTARI
    GuildRejectNewMember    'RECHAZAR
    GuildKickMember         'ECHARCLA
    GuildUpdateNews         'ACTGNEWS
    GuildMemberInfo         '1HRINFO<
    GuildOpenElections      'ABREELEC
    GuildRequestMembership  'SOLICITUD
    GuildRequestDetails     'CLANDETAILS
    Online                  '/ONLINE
    Quit                    '/SALIR
    GuildLeave              '/SALIRCLAN
    RequestAccountState     '/BALANCE
    PetStand                '/QUIETO
    PetFollow               '/ACOMPAÑAR
    ReleasePet              '/LIBERAR
    TrainList               '/ENTRENAR
    Rest                    '/DESCANSAR
    Meditate                '/MEDITAR
    Resucitate              '/RESUCITAR
    Heal                    '/CURAR
    Help                    '/AYUDA
    RequestStats            '/EST
    CommerceStart           '/COMERCIAR
    BankStart               '/BOVEDA
    Enlist                  '/ENLISTAR
    Information             '/INFORMACION
    Reward                  '/RECOMPENSA
    RequestMOTD             '/MOTD
    Uptime                  '/UPTIME
    PartyLeave              '/SALIRPARTY
    PartyCreate             '/CREARPARTY
    PartyJoin               '/PARTY
    Inquiry                 '/ENCUESTA ( with no params )
    GuildMessage            '/CMSG
    PartyMessage            '/PMSG
    CentinelReport          '/CENTINELA
    GuildOnline             '/ONLINECLAN
    PartyOnline             '/ONLINEPARTY
    CouncilMessage          '/BMSG
    RoleMasterRequest       '/ROL
    GMRequest               '/GM
    bugReport               '/_BUG
    ChangeDescription       '/DESC
    GuildVote               '/VOTO
    Punishments             '/PENAS
    ChangePassword          '/CONTRASEÑA
    Gamble                  '/APOSTAR
    InquiryVote             '/ENCUESTA ( with parameters )
    LeaveFaction            '/RETIRAR ( with no arguments )
    BankExtractGold         '/RETIRAR ( with arguments )
    BankDepositGold         '/DEPOSITAR
    Denounce                '/DENUNCIAR
    GuildFundate            '/FUNDARCLAN
    GuildFundation
    GuildResolve            '/cerrarclan
    PartyKick               '/ECHARPARTY
    PartySetLeader          '/PARTYLIDER
    PartyAcceptMember       '/ACCEPTPARTY
    Ping                    '/PING
    
    RequestPartyForm
    ItemUpgrade
    GMCommands
    InitCrafting
    Home
    ShowGuildNews
    ShareNpc                '/COMPARTIRNPC
    StopSharingNpc          '/NOCOMPARTIRNPC
    Consulta
    
    Editame
    MoverItemEnInventario
    GlobalChat
    FaccionChat
    Dame_Promedio
    Save_Character
    Init_Invasion
    CuentaCL
    IniciarSubasta
    OfertarSubasta
    ConsultaSubasta
    Enter_Duel
    Exit_Duel
    CuentaRegresiva
    Prestigio
    CentroCanje
    
    GiveMePower
    Quest                   '/QUEST
    QuestAccept
    QuestListRequest
    QuestDetailsRequest
    QuestAbandon
    
End Enum

Public Enum FontTypeNames
    FONTTYPE_TALK
    FONTTYPE_FIGHT
    FONTTYPE_WARNING
    FONTTYPE_INFO
    FONTTYPE_INFOBOLD
    FONTTYPE_EJECUCION
    FONTTYPE_PARTY
    FONTTYPE_VENENO
    FONTTYPE_GUILD
    FONTTYPE_SERVER
    FONTTYPE_GUILDMSG
    FONTTYPE_CONSEJO
    FONTTYPE_CONSEJOCAOS
    FONTTYPE_CONSEJOVesA
    FONTTYPE_CONSEJOCAOSVesA
    FONTTYPE_CENTINELA
    FONTTYPE_GMMSG
    FONTTYPE_GM
    FONTTYPE_CITIZEN
    FONTTYPE_CONSE
    FONTTYPE_DIOS
    FONTTYPE_DEATHMATCH
End Enum

Public FontTypes(21) As tFont

Public Sub EntrarDuelo()
    With outgoingData
        Call .WriteByte(ClientPacketID.Enter_Duel)
    End With
End Sub

''
'   Flushes the outgoing data buffer of the user.
'
'   @param    UserIndex User whose outgoing data buffer will be flushed.

Public Sub FlushBuffer()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Sends all data existing in the buffer
'***************************************************
    Dim sndData As String
    
    With outgoingData
        If .Length = 0 Then _
            Exit Sub
        
        sndData = .ReadASCIIStringFixed(.Length)
        
        Call SendData(sndData)
    End With
End Sub

''
'   Handles the AddForumMessage message.

Private Sub HandleAddForumMessage()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.Length < 8 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo ErrHandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call Buffer.ReadByte
    
    Dim ForumType As eForumMsgType
    Dim Title As String
    Dim Message As String
    Dim Author As String
    ForumType = Buffer.ReadByte
    
    Title = Buffer.ReadASCIIString()
    Author = Buffer.ReadASCIIString()
    Message = Buffer.ReadASCIIString()
    
    If Not frmForo.ForoLimpio Then
        clsForos.ClearForums
        frmForo.ForoLimpio = True
    End If

    Call clsForos.AddPost(General_Get_Forum_Alignment(ForumType), Title, Author, Message, General_Is_Anounce(ForumType))
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(Buffer)
    
ErrHandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error <> 0 Then _
        Err.Raise Error
End Sub

'   Handles the AddSlots message.
Private Sub HandleAddSlots()
'***************************************************
'Author: Budi
'Last Modification: 12/01/09
'
'***************************************************

    Call incomingData.ReadByte
    
    MaxInventorySlots = incomingData.ReadByte
End Sub

''
'   Handles the AlianceProposalsList message.

Private Sub HandleAlianceProposalsList()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.Length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo ErrHandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call Buffer.ReadByte
    
    Dim vsGuildList() As String
    Dim i As Long
    
    vsGuildList = Split(Buffer.ReadASCIIString(), SEPARATOR)
    
    Call frmPeaceProp.lista.Clear
    For i = 0 To UBound(vsGuildList())
        Call frmPeaceProp.lista.AddItem(vsGuildList(i))
    Next i
    
    frmPeaceProp.ProposalType = TIPO_PROPUESTA.ALIANZA
    Call frmPeaceProp.Show(vbModeless, frmMain)
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(Buffer)
    
ErrHandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error <> 0 Then _
        Err.Raise Error
End Sub

''
'   Handles the AreaChanged message.

Private Sub HandleAreaChanged()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.Length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Dim X As Byte
    Dim Y As Byte
    
    X = incomingData.ReadByte()
    Y = incomingData.ReadByte()
        
    Call CambioDeArea(X, Y)
End Sub

''
'   Handles the Attributes message.

Private Sub HandleAtributes()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.Length < 1 + NUMATRIBUTES Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Dim i As Long
    
    For i = 1 To NUMATRIBUTES
        UserAtributos(i) = incomingData.ReadByte()
    Next i
    
    'Show them in character creation
    If EstadoLogin = E_MODO.Dados Then
        With frmCrearPersonaje
            If .Visible Then
                For i = 1 To NUMATRIBUTES
                    .lblAtributos(i).Caption = UserAtributos(i)
                Next i
                
                .UpdateStats
            End If
        End With
    Else
        LlegaronAtrib = True
    End If
End Sub

Private Sub HandleAuras()
'***************************************************
'Author: Standelf
'Last Modification: 27/05/10
'***************************************************
    With incomingData
        'Remove PacketID
        Call .ReadByte
        SetCharacterAura .ReadInteger, .ReadByte, .ReadByte
    End With
End Sub

''
'   Handles the BankEnd message.

Private Sub HandleBankEnd()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    
    Set InvBanco(0) = Nothing
    Set InvBanco(1) = Nothing
    
    Unload frmBancoObj
    Comerciando = False
End Sub

''
'   Handles the BankInit message.

Private Sub HandleBankInit()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    Dim i As Long
    Dim BankGold As Long
    
    'Remove packet ID
    Call incomingData.ReadByte
    
        BankGold = incomingData.ReadLong
    Call InvBanco(0).Initialize(DirectD3D8, frmBancoObj.PicBancoInv, MAX_BANCOINVENTORY_SLOTS)
    Call InvBanco(1).Initialize(DirectD3D8, frmBancoObj.PicInv, Inventario.MaxObjs)
    
    For i = 1 To Inventario.MaxObjs
        With Inventario
            Call InvBanco(1).SetItem(i, .OBJIndex(i), _
                .Amount(i), .Equipped(i), .GrhIndex(i), _
                .OBJType(i), .MaxHit(i), .MinHit(i), .MaxDef(i), .MinDef(i), _
                .Valor(i), .ItemName(i))
        End With
    Next i
    
    For i = 1 To MAX_BANCOINVENTORY_SLOTS
        With UserBancoInventory(i)
            Call InvBanco(0).SetItem(i, .OBJIndex, _
                .Amount, .Equipped, .GrhIndex, _
                .OBJType, .MaxHit, .MinHit, .MaxDef, .MinDef, _
                .Valor, .name)
        End With
    Next i
    
    'Set state and show form
    Comerciando = True
    
    frmBancoObj.lblUserGld.Caption = Format(BankGold, "###,###,###,###")
    
    frmBancoObj.Show , frmMain
End Sub

''
'   Handles the BankOK message.

Private Sub HandleBankOK()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    
    Dim i As Long
    
    If frmBancoObj.Visible Then
        
        For i = 1 To Inventario.MaxObjs
            With Inventario
                Call InvBanco(1).SetItem(i, .OBJIndex(i), .Amount(i), _
                    .Equipped(i), .GrhIndex(i), .OBJType(i), .MaxHit(i), _
                    .MinHit(i), .MaxDef(i), .MinDef(i), .Valor(i), .ItemName(i))
            End With
        Next i
        
        'Alter order according to if we bought or sold so the labels and grh remain the same
        If frmBancoObj.LasActionBuy Then
            'frmBancoObj.List1(1).ListIndex = frmBancoObj.LastIndex2
            'frmBancoObj.List1(0).ListIndex = frmBancoObj.LastIndex1
        Else
            'frmBancoObj.List1(0).ListIndex = frmBancoObj.LastIndex1
            'frmBancoObj.List1(1).ListIndex = frmBancoObj.LastIndex2
        End If
        
        frmBancoObj.NoPuedeMover = False
    End If
       
End Sub

''
'   Handles the BlacksmithArmors message.

Private Sub HandleBlacksmithArmors()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.Length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo ErrHandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call Buffer.ReadByte
    
    Dim Count As Integer
    Dim i As Long
    Dim J As Long
    Dim K As Long
    
    Count = Buffer.ReadInteger()
    
    ReDim ArmadurasHerrero(Count) As tItemsConstruibles
    
    For i = 1 To Count
        With ArmadurasHerrero(i)
            .name = Buffer.ReadASCIIString()    'Get the object's name
            .GrhIndex = Buffer.ReadInteger()
            .LinH = Buffer.ReadInteger()        'The iron needed
            .LinP = Buffer.ReadInteger()        'The silver needed
            .LinO = Buffer.ReadInteger()        'The gold needed
            .OBJIndex = Buffer.ReadInteger()
            .Upgrade = Buffer.ReadInteger()
        End With
    Next i
    
    J = UBound(HerreroMejorar)
    
    For i = 1 To Count
        With ArmadurasHerrero(i)
            If .Upgrade Then
                For K = 1 To Count
                    If .Upgrade = ArmadurasHerrero(K).OBJIndex Then
                        J = J + 1
                
                        ReDim Preserve HerreroMejorar(J) As tItemsConstruibles
                        
                        HerreroMejorar(J).name = .name
                        HerreroMejorar(J).GrhIndex = .GrhIndex
                        HerreroMejorar(J).OBJIndex = .OBJIndex
                        HerreroMejorar(J).UpgradeName = ArmadurasHerrero(K).name
                        HerreroMejorar(J).UpgradeGrhIndex = ArmadurasHerrero(K).GrhIndex
                        HerreroMejorar(J).LinH = ArmadurasHerrero(K).LinH - .LinH * 0.85
                        HerreroMejorar(J).LinP = ArmadurasHerrero(K).LinP - .LinP * 0.85
                        HerreroMejorar(J).LinO = ArmadurasHerrero(K).LinO - .LinO * 0.85
                        
                        Exit For
                    End If
                Next K
            End If
        End With
    Next i
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(Buffer)
    
ErrHandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error <> 0 Then _
        Err.Raise Error
End Sub

''
'   Handles the BlacksmithWeapons message.

Private Sub HandleBlacksmithWeapons()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.Length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo ErrHandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call Buffer.ReadByte
    
    Dim Count As Integer
    Dim i As Long
    Dim J As Long
    Dim K As Long
    
    Count = Buffer.ReadInteger()
    
    ReDim ArmasHerrero(Count) As tItemsConstruibles
    ReDim HerreroMejorar(0) As tItemsConstruibles
    
    For i = 1 To Count
        With ArmasHerrero(i)
            .name = Buffer.ReadASCIIString()    'Get the object's name
            .GrhIndex = Buffer.ReadInteger()
            .LinH = Buffer.ReadInteger()        'The iron needed
            .LinP = Buffer.ReadInteger()        'The silver needed
            .LinO = Buffer.ReadInteger()        'The gold needed
            .OBJIndex = Buffer.ReadInteger()
            .Upgrade = Buffer.ReadInteger()
        End With
    Next i
    
    With frmHerrero
        '   Inicializo los inventarios
        Call InvLingosHerreria(1).Initialize(DirectD3D, .picLingotes0, 3, , , , , , False)
        Call InvLingosHerreria(2).Initialize(DirectD3D, .picLingotes1, 3, , , , , , False)
        Call InvLingosHerreria(3).Initialize(DirectD3D, .picLingotes2, 3, , , , , , False)
        Call InvLingosHerreria(4).Initialize(DirectD3D, .picLingotes3, 3, , , , , , False)
        
        Call .HideExtraControls(Count)
        Call .RenderList(1, True)
    End With
    
    For i = 1 To Count
        With ArmasHerrero(i)
            If .Upgrade Then
                For K = 1 To Count
                    If .Upgrade = ArmasHerrero(K).OBJIndex Then
                        J = J + 1
                
                        ReDim Preserve HerreroMejorar(J) As tItemsConstruibles
                        
                        HerreroMejorar(J).name = .name
                        HerreroMejorar(J).GrhIndex = .GrhIndex
                        HerreroMejorar(J).OBJIndex = .OBJIndex
                        HerreroMejorar(J).UpgradeName = ArmasHerrero(K).name
                        HerreroMejorar(J).UpgradeGrhIndex = ArmasHerrero(K).GrhIndex
                        HerreroMejorar(J).LinH = ArmasHerrero(K).LinH - .LinH * 0.85
                        HerreroMejorar(J).LinP = ArmasHerrero(K).LinP - .LinP * 0.85
                        HerreroMejorar(J).LinO = ArmasHerrero(K).LinO - .LinO * 0.85
                        
                        Exit For
                    End If
                Next K
            End If
        End With
    Next i
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(Buffer)
    
ErrHandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error <> 0 Then _
        Err.Raise Error
End Sub

''
'   Handles the Blind message.

Private Sub HandleBlind()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    
    UserCiego = True
End Sub

''
'   Handles the BlindNoMore message.

Private Sub HandleBlindNoMore()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    
    UserCiego = False
End Sub

''
'   Handles the BlockPosition message.

Private Sub HandleBlockPosition()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.Length < 4 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Dim X As Byte
    Dim Y As Byte
    
    X = incomingData.ReadByte()
    Y = incomingData.ReadByte()
    
    If incomingData.ReadBoolean() Then
        MapData(X, Y).Blocked = 1
    Else
        MapData(X, Y).Blocked = 0
    End If
End Sub

'   Handles the CancelOfferItem message.


Private Sub HandleCancelOfferItem()
'***************************************************
'Author: Torres Patricio (Pato)
'Last Modification: 05/03/10
'
'***************************************************
    Dim Slot As Byte
    Dim Amount As Long
    
    Call incomingData.ReadByte
    
    Slot = incomingData.ReadByte
    
    With InvOfferComUsu(0)
        Amount = .Amount(Slot)
        
        '   No tiene sentido que se quiten 0 unidades
        If Amount <> 0 Then
            '   Actualizo el inventario general
            Call frmComerciarUsu.UpdateInvCom(.OBJIndex(Slot), Amount)
            
            '   Borro el item
            Call .SetItem(Slot, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, "")
        End If
    End With
    
    '   Si era el único ítem de la oferta, no puede confirmarla
    If Not frmComerciarUsu.HasAnyItem(InvOfferComUsu(0)) And _
        Not frmComerciarUsu.HasAnyItem(InvOroComUsu(1)) Then Call frmComerciarUsu.HabilitarConfirmar(False)
    
    With FontTypes(FontTypeNames.FONTTYPE_INFO)
        Call frmComerciarUsu.PrintCommerceMsg("¡No puedes comerciar ese objeto!", FontTypeNames.FONTTYPE_INFO)
    End With
End Sub

''
'   Handles the CarpenterObjects message.

Private Sub HandleCarpenterObjects()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.Length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo ErrHandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call Buffer.ReadByte
    
    Dim Count As Integer
    Dim i As Long
    Dim J As Long
    Dim K As Long
    
    Count = Buffer.ReadInteger()
    
    ReDim ObjCarpintero(Count) As tItemsConstruibles
    ReDim CarpinteroMejorar(0) As tItemsConstruibles
    
    For i = 1 To Count
        With ObjCarpintero(i)
            .name = Buffer.ReadASCIIString()        'Get the object's name
            .GrhIndex = Buffer.ReadInteger()
            .Madera = Buffer.ReadInteger()          'The wood needed
            .MaderaElfica = Buffer.ReadInteger()    'The elfic wood needed
            .OBJIndex = Buffer.ReadInteger()
            .Upgrade = Buffer.ReadInteger()
        End With
    Next i
    
    With frmCarp
        '   Inicializo los inventarios
        Call InvMaderasCarpinteria(1).Initialize(DirectD3D, .picMaderas0, 2, , , , , , False)
        Call InvMaderasCarpinteria(2).Initialize(DirectD3D, .picMaderas1, 2, , , , , , False)
        Call InvMaderasCarpinteria(3).Initialize(DirectD3D, .picMaderas2, 2, , , , , , False)
        Call InvMaderasCarpinteria(4).Initialize(DirectD3D, .picMaderas3, 2, , , , , , False)
        
        Call .HideExtraControls(Count)
        Call .RenderList(1)
    End With
    
    For i = 1 To Count
        With ObjCarpintero(i)
            If .Upgrade Then
                For K = 1 To Count
                    If .Upgrade = ObjCarpintero(K).OBJIndex Then
                        J = J + 1
                
                        ReDim Preserve CarpinteroMejorar(J) As tItemsConstruibles
                        
                        CarpinteroMejorar(J).name = .name
                        CarpinteroMejorar(J).GrhIndex = .GrhIndex
                        CarpinteroMejorar(J).OBJIndex = .OBJIndex
                        CarpinteroMejorar(J).UpgradeName = ObjCarpintero(K).name
                        CarpinteroMejorar(J).UpgradeGrhIndex = ObjCarpintero(K).GrhIndex
                        CarpinteroMejorar(J).Madera = ObjCarpintero(K).Madera - .Madera * 0.85
                        CarpinteroMejorar(J).MaderaElfica = ObjCarpintero(K).MaderaElfica - .MaderaElfica * 0.85
                        
                        Exit For
                    End If
                Next K
            End If
        End With
    Next i
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(Buffer)
    
ErrHandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error <> 0 Then _
        Err.Raise Error
End Sub

Public Sub HandleCentroCanje()
'***************************************************
'Author: EzequielJuarez (Standelf)
'Last Modification: 22/01/2011
'***************************************************

    'Remove packet ID
    Call incomingData.ReadByte
    
    'Dim cantidad As Byte, i As Byte
        'cantidad = incomingData.ReadByte()
        
        'For i = 1 To cantidad
            
End Sub

''
'   Handles the ChangeBankSlot message.

Private Sub HandleChangeBankSlot()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.Length < 21 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo ErrHandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call Buffer.ReadByte
    
    Dim Slot As Byte
    Slot = Buffer.ReadByte()
    
    With UserBancoInventory(Slot)
        .OBJIndex = Buffer.ReadInteger()
        .name = Buffer.ReadASCIIString()
        .Amount = Buffer.ReadInteger()
        .GrhIndex = Buffer.ReadInteger()
        .OBJType = Buffer.ReadByte()
        .MaxHit = Buffer.ReadInteger()
        .MinHit = Buffer.ReadInteger()
        .MaxDef = Buffer.ReadInteger()
        .MinDef = Buffer.ReadInteger
        .Valor = Buffer.ReadLong()
        
        If Comerciando Then
            Call InvBanco(0).SetItem(Slot, .OBJIndex, .Amount, _
                .Equipped, .GrhIndex, .OBJType, .MaxHit, _
                .MinHit, .MaxDef, .MinDef, .Valor, .name)
        End If
    End With
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(Buffer)
    
ErrHandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error <> 0 Then _
        Err.Raise Error
End Sub

''
'   Handles the ChangeInventorySlot message.

Private Sub HandleChangeInventorySlot()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.Length < 22 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo ErrHandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call Buffer.ReadByte
    
    Dim Slot As Byte
    Dim OBJIndex As Integer
    Dim name As String
    Dim Amount As Integer
    Dim Equipped As Boolean
    Dim GrhIndex As Integer
    Dim OBJType As Byte
    Dim MaxHit As Integer
    Dim MinHit As Integer
    Dim MaxDef As Integer
    Dim MinDef As Integer
    Dim value As Single
    Dim PuedeUsar As Boolean
    
    Slot = Buffer.ReadByte()
    Inventario.SlotValidos = Buffer.ReadByte
    
    OBJIndex = Buffer.ReadInteger()
    name = Buffer.ReadASCIIString()
    Amount = Buffer.ReadInteger()
    Equipped = Buffer.ReadBoolean()
    GrhIndex = Buffer.ReadInteger()
    OBJType = Buffer.ReadByte()
    MaxHit = Buffer.ReadInteger()
    MinHit = Buffer.ReadInteger()
    MaxDef = Buffer.ReadInteger()
    MinDef = Buffer.ReadInteger
    value = Buffer.ReadSingle()
    PuedeUsar = Buffer.ReadBoolean()
    
    If Equipped Then
        Select Case OBJType
            Case eObjType.otWeapon
                frmMain.lblWeapon = MinHit & "/" & MaxHit
                UserWeaponEqpSlot = Slot
            Case eObjType.otArmadura
                frmMain.lblArmor = MinDef & "/" & MaxDef
                UserArmourEqpSlot = Slot
            Case eObjType.otescudo
                frmMain.lblShielder = MinDef & "/" & MaxDef
                UserHelmEqpSlot = Slot
            Case eObjType.otcasco
                frmMain.lblHelm = MinDef & "/" & MaxDef
                UserShieldEqpSlot = Slot
        End Select
    Else
        Select Case Slot
            Case UserWeaponEqpSlot
                frmMain.lblWeapon = "0/0"
                UserWeaponEqpSlot = 0
            Case UserArmourEqpSlot
                frmMain.lblArmor = "0/0"
                UserArmourEqpSlot = 0
            Case UserHelmEqpSlot
                frmMain.lblShielder = "0/0"
                UserHelmEqpSlot = 0
            Case UserShieldEqpSlot
                frmMain.lblHelm = "0/0"
                UserShieldEqpSlot = 0
        End Select
    End If
    
    Call Inventario.SetItem(Slot, OBJIndex, Amount, Equipped, GrhIndex, OBJType, MaxHit, MinHit, MaxDef, MinDef, value, name, PuedeUsar)

    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(Buffer)
    
ErrHandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error <> 0 Then _
        Err.Raise Error
End Sub

''
'   Handles the ChangeMap message.
Private Sub HandleChangeMap()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.Length < 5 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    UserMap = incomingData.ReadInteger()
    
'TODO: Once on-the-fly editor is implemented check for map version before loading....
'For now we just drop it
    Call incomingData.ReadInteger
    UserMapName = incomingData.ReadASCIIString()
                frmMain.lblMapPosX.Caption = UserPos.X
                frmMain.lblMapPosY.Caption = UserPos.Y

    frmMain.lblMapName.Caption = UserMapName
    
    
    If General_File_Exist(Resources.Maps & "mapa" & CStr(UserMap) & ".map", vbNormal) Then
        Call General_Load_MapData(UserMap)
    Else
        'no encontramos el mapa en el hd
        MsgBox "Error en los mapas, algún archivo ha sido modificado o esta dañado."
        
        Call General_Close_Client
    End If
End Sub

''
'   Handles the ChangeNPCInventorySlot message.

Private Sub HandleChangeNPCInventorySlot()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.Length < 21 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo ErrHandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call Buffer.ReadByte
    
    Dim Slot As Byte
    Slot = Buffer.ReadByte()
    
    With NPCInventory(Slot)
        .name = Buffer.ReadASCIIString()
        .Amount = Buffer.ReadInteger()
        .Valor = Buffer.ReadSingle()
        .GrhIndex = Buffer.ReadInteger()
        .OBJIndex = Buffer.ReadInteger()
        .OBJType = Buffer.ReadByte()
        .MaxHit = Buffer.ReadInteger()
        .MinHit = Buffer.ReadInteger()
        .MaxDef = Buffer.ReadInteger()
        .MinDef = Buffer.ReadInteger
    End With
        
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(Buffer)
    
ErrHandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error <> 0 Then _
        Err.Raise Error
End Sub

''
'   Handles the ChangeSpellSlot message.

Private Sub HandleChangeSpellSlot()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.Length < 6 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo ErrHandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call Buffer.ReadByte
    
    Dim Slot As Byte
    Slot = Buffer.ReadByte()
    
    UserHechizos(Slot) = Buffer.ReadInteger()
    
    If Slot <= frmMain.hlst.ListCount Then
        frmMain.hlst.List(Slot - 1) = Buffer.ReadASCIIString()
    Else
        Call frmMain.hlst.AddItem(Buffer.ReadASCIIString())
    End If
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(Buffer)
    
ErrHandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error <> 0 Then _
        Err.Raise Error
End Sub

''
'   Handles the ChangeUserTradeSlot message.

Private Sub HandleChangeUserTradeSlot()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.Length < 22 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo ErrHandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    
    Dim OfferSlot As Byte
    
    'Remove packet ID
    Call Buffer.ReadByte
    
    OfferSlot = Buffer.ReadByte
    
    With Buffer
        If OfferSlot = GOLD_OFFER_SLOT Then
            Call InvOroComUsu(2).SetItem(1, .ReadInteger(), .ReadLong(), 0, _
                                            .ReadInteger(), .ReadByte(), .ReadInteger(), _
                                            .ReadInteger(), .ReadInteger(), .ReadInteger(), .ReadLong(), .ReadASCIIString())
        Else
            Call InvOfferComUsu(1).SetItem(OfferSlot, .ReadInteger(), .ReadLong(), 0, _
                                            .ReadInteger(), .ReadByte(), .ReadInteger(), _
                                            .ReadInteger(), .ReadInteger(), .ReadInteger(), .ReadLong(), .ReadASCIIString())
        End If
    End With
    
    Call frmComerciarUsu.PrintCommerceMsg(TradingUserName & " ha modificado su oferta.", FontTypeNames.FONTTYPE_VENENO)
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(Buffer)
    
ErrHandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error <> 0 Then _
        Err.Raise Error
End Sub

''
'   Handles the CharacterChange message.

Private Sub HandleCharacterChange()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 25/08/2009
'25/08/2009: ZaMa - Changed a variable used incorrectly.
'***************************************************
    If incomingData.Length < 18 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Dim CharIndex As Integer
    Dim tempint As Integer
    Dim HeadIndex As Integer
    
    CharIndex = incomingData.ReadInteger()
    
    With CharList(CharIndex)
        tempint = incomingData.ReadInteger()
        
        If tempint < LBound(BodyData()) Or tempint > UBound(BodyData()) Then
            .Body = BodyData(0)
            .iBody = 0
        Else
            .Body = BodyData(tempint)
            .iBody = tempint
        End If
        
        
        HeadIndex = incomingData.ReadInteger()
        
        If HeadIndex < LBound(HeadData()) Or HeadIndex > UBound(HeadData()) Then
            .Head = HeadData(0)
            .iHead = 0
        Else
            .Head = HeadData(HeadIndex)
            .iHead = HeadIndex
        End If
        
        .muerto = (HeadIndex = CASPER_HEAD)

        .Heading = incomingData.ReadByte()
        
        tempint = incomingData.ReadInteger()
        If tempint <> 0 Then .Arma = WeaponAnimData(tempint)
        
        tempint = incomingData.ReadInteger()
        If tempint <> 0 Then .Escudo = ShieldAnimData(tempint)
        
        tempint = incomingData.ReadInteger()
        If tempint <> 0 Then .Casco = CascoAnimData(tempint)
        
        Call SetCharacterFx(CharIndex, incomingData.ReadInteger(), incomingData.ReadInteger())
    End With
    
    Call RefreshAllChars
End Sub

Private Sub HandleCharacterChangeNick()
'***************************************************
'Author: Budi
'Last Modification: 07/23/09
'
'***************************************************
    If incomingData.Length < 6 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet id
    Call incomingData.ReadByte
    Dim CharIndex As Integer
    CharIndex = incomingData.ReadInteger
    CharList(CharIndex).Nombre = incomingData.ReadASCIIString
    
End Sub

''
'   Handles the CharacterCreate message.

Private Sub HandleCharacterCreate()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.Length < 24 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo ErrHandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call Buffer.ReadByte
    
    Dim CharIndex As Integer
    Dim Body As Integer
    Dim Head As Integer
    Dim Heading As E_Heading
    Dim X As Byte
    Dim Y As Byte
    Dim weapon As Integer
    Dim shield As Integer
    Dim helmet As Integer
    Dim privs As Integer
    Dim NickColor As Byte
    
    CharIndex = Buffer.ReadInteger()
    Body = Buffer.ReadInteger()
    Head = Buffer.ReadInteger()
    Heading = Buffer.ReadByte()
    X = Buffer.ReadByte()
    Y = Buffer.ReadByte()
    weapon = Buffer.ReadInteger()
    shield = Buffer.ReadInteger()
    helmet = Buffer.ReadInteger()
    
    
    With CharList(CharIndex)
    
        Call SetCharacterFx(CharIndex, Buffer.ReadInteger(), Buffer.ReadInteger())
        
        .Nombre = Buffer.ReadASCIIString()
        NickColor = Buffer.ReadByte()
        
        If (NickColor And eNickColor.ieCriminal) <> 0 Then
            .Criminal = 1
        Else
            .Criminal = 0
        End If
        
        .Atacable = (NickColor And eNickColor.ieAtacable) <> 0
        
        privs = Buffer.ReadByte()
        
        If privs <> 0 Then
            'If the player belongs to a council AND is an admin, only whos as an admin
            If (privs And PlayerType.ChaosCouncil) <> 0 And (privs And PlayerType.User) = 0 Then
                privs = privs Xor PlayerType.ChaosCouncil
            End If
            
            If (privs And PlayerType.RoyalCouncil) <> 0 And (privs And PlayerType.User) = 0 Then
                privs = privs Xor PlayerType.RoyalCouncil
            End If
            
            'If the player is a RM, ignore other flags
            If privs And PlayerType.RoleMaster Then
                privs = PlayerType.RoleMaster
            End If
            
            'Log2 of the bit flags sent by the server gives our numbers ^^
            .priv = Log(privs) / Log(2)
        Else
            .priv = 0
        End If
    End With
    
    Call MakeChar(CharIndex, Body, Head, Heading, X, Y, weapon, shield, helmet)
    
    Call RefreshAllChars
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(Buffer)
    
ErrHandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error <> 0 Then _
        Err.Raise Error
End Sub

''
'   Handles the CharacterInfo message.

Private Sub HandleCharacterInfo()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.Length < 35 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo ErrHandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call Buffer.ReadByte
    
    With frmCharInfo
        If .frmType = CharInfoFrmType.frmMembers Then
            .imgRechazar.Visible = False
            .imgAceptar.Visible = False
            .imgEchar.Visible = True
            .imgPeticion.Visible = False
        Else
            .imgRechazar.Visible = True
            .imgAceptar.Visible = True
            .imgEchar.Visible = False
            .imgPeticion.Visible = True
        End If
        
        .Nombre.Caption = Buffer.ReadASCIIString()
        .Raza.Caption = ListaRazas(Buffer.ReadByte())
        .Clase.Caption = ListaClases(Buffer.ReadByte())
        
        If Buffer.ReadByte() = 1 Then
            .Genero.Caption = "Hombre"
        Else
            .Genero.Caption = "Mujer"
        End If
        
        .Nivel.Caption = Buffer.ReadByte()
        .Oro.Caption = Buffer.ReadLong()
        .Banco.Caption = Buffer.ReadLong()
        
        Dim reputation As Long
        reputation = Buffer.ReadLong()
        
        .reputacion.Caption = reputation
        
        .txtPeticiones.Text = Buffer.ReadASCIIString()
        .guildactual.Caption = Buffer.ReadASCIIString()
        .txtMiembro.Text = Buffer.ReadASCIIString()
        
        Dim armada As Boolean
        Dim caos As Boolean
        
        armada = Buffer.ReadBoolean()
        caos = Buffer.ReadBoolean()
        
        If armada Then
            .ejercito.Caption = "Armada Real"
        ElseIf caos Then
            .ejercito.Caption = "Legión Oscura"
        End If
        
        .Ciudadanos.Caption = CStr(Buffer.ReadLong())
        .criminales.Caption = CStr(Buffer.ReadLong())
        
        If reputation > 0 Then
            .status.Caption = " Ciudadano"
            .status.ForeColor = vbBlue
        Else
            .status.Caption = " Criminal"
            .status.ForeColor = vbRed
        End If
        
        Call .Show(vbModeless, frmMain)
    End With
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(Buffer)
    
ErrHandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error <> 0 Then _
        Err.Raise Error
End Sub

''
'   Handles the CharacterMove message.

Private Sub HandleCharacterMove()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.Length < 5 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Dim CharIndex As Integer
    Dim X As Byte
    Dim Y As Byte
    
    CharIndex = incomingData.ReadInteger()
    X = incomingData.ReadByte()
    Y = incomingData.ReadByte()
    
    With CharList(CharIndex)
        If .FxIndex >= 40 And .FxIndex <= 49 Then   'If it's meditating, we remove the FX
            .FxIndex = 0
        End If
        
        '   Play steps sounds if the user is not an admin of any kind
        If .priv <> 1 And .priv <> 2 And .priv <> 3 And .priv <> 5 And .priv <> 25 Then
            Call DoPasosFx(CharIndex)
        End If
    End With
    
    Call MoveCharbyPos(CharIndex, X, Y)
    
    Call RefreshAllChars
End Sub

''
'   Handles the CharacterRemove message.

Private Sub HandleCharacterRemove()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.Length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Dim CharIndex As Integer
    
    CharIndex = incomingData.ReadInteger()
    
    Call EraseChar(CharIndex)
    Call RefreshAllChars
End Sub

Public Sub HandleCharParticles()
'***************************************************
'Author: EzequielJuarez (Standelf)
'Last Modification: 14/01/2011
'***************************************************
    Dim ParticleIndex As Integer, Loops As Integer, CharIndex As Integer
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    CharIndex = incomingData.ReadInteger
    ParticleIndex = incomingData.ReadByte
    Loops = Int((incomingData.ReadByte) * 20)
    
    If CharList(CharIndex).ParticleIndex <> 0 Then
        Call Effect_Kill(CharList(CharIndex).ParticleIndex, False)
    End If
    
    
    If ParticleIndex <> 0 Then Call Engine_CreateEffect(CharIndex, ParticleIndex, CSng(Engine_TPtoSPX(CharList(CharIndex).Pos.X)), CSng(Engine_TPtoSPY(CharList(CharIndex).Pos.Y)), Loops)

End Sub

''
'   Handles the ChatOverHead message.

Private Sub HandleChatOverHead()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.Length < 8 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo ErrHandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call Buffer.ReadByte
    
    Dim chat As String
    Dim CharIndex As Integer
    
    chat = Buffer.ReadASCIIString()
    CharIndex = Buffer.ReadInteger()
    
    'Only add the chat if the character exists (a CharacterRemove may have been sent to the PC / NPC area before the buffer was flushed)
    If CharList(CharIndex).Active Then
        Dim tmp As D3DCOLORVALUE
            tmp.a = 0
            tmp.r = Buffer.ReadByte()
            tmp.g = Buffer.ReadByte()
            tmp.b = Buffer.ReadByte()

        Call Dialogos.CreateDialog(Trim$(chat), CharIndex, tmp)

        'Standelf
        If (Not (mid(Trim$(chat), 1, 1) = "-" Or mid(Trim$(chat), 1, 1) = "+")) And Settings.DialogosEnConsola Then
            If General_Get_GM(CharIndex) Then
                Call General_Add_to_RichTextBox(frmMain.RecTxt, "[Area-GM] " & General_Get_RawName(CharList(CharIndex).Nombre) & "> " & Trim$(chat), tmp.r, tmp.g, tmp.b, 0, 1)
            ElseIf Settings.DialogosEnConsola = True Then
                Call General_Add_to_RichTextBox(frmMain.RecTxt, "[Area] " & General_Get_RawName(CharList(CharIndex).Nombre) & "> " & Trim$(chat), tmp.r, tmp.g, tmp.b, 0, 1)
            End If
        End If
    Else
        Call Buffer.ReadByte
        Call Buffer.ReadByte
        Call Buffer.ReadByte
    End If
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(Buffer)

ErrHandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then _
        Err.Raise Error
End Sub

''
'   Handles the ConsoleMessage message.

Private Sub HandleCommerceChat()
'***************************************************
'Author: ZaMa
'Last Modification: 03/12/2009
'
'***************************************************
    If incomingData.Length < 4 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo ErrHandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call Buffer.ReadByte
    
    Dim chat As String
    Dim FontIndex As Integer
    Dim str As String
    Dim r As Byte
    Dim g As Byte
    Dim b As Byte
    
    chat = Buffer.ReadASCIIString()
    FontIndex = Buffer.ReadByte()
    
    If InStr(1, chat, "~") Then
        str = General_Get_ReadField(2, chat, 126)
            If Val(str) > 255 Then
                r = 255
            Else
                r = Val(str)
            End If
            
            str = General_Get_ReadField(3, chat, 126)
            If Val(str) > 255 Then
                g = 255
            Else
                g = Val(str)
            End If
            
            str = General_Get_ReadField(4, chat, 126)
            If Val(str) > 255 Then
                b = 255
            Else
                b = Val(str)
            End If
            
        Call General_Add_to_RichTextBox(frmComerciarUsu.CommerceConsole, Left$(chat, InStr(1, chat, "~") - 1), r, g, b, Val(General_Get_ReadField(5, chat, 126)) <> 0, Val(General_Get_ReadField(6, chat, 126)) <> 0)
    Else
        With FontTypes(FontIndex)
            Call General_Add_to_RichTextBox(frmComerciarUsu.CommerceConsole, chat, .Red, .Green, .Blue, .bold, .italic)
        End With
    End If
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(Buffer)
    
ErrHandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error <> 0 Then _
        Err.Raise Error
End Sub

''
'   Handles the CommerceEnd message.

Private Sub HandleCommerceEnd()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    
    'Reset vars
    Comerciando = False
    
    'Hide form
    Unload frmComerciar
End Sub

''
'   Handles the CommerceInit message.

Private Sub HandleCommerceInit()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    Dim i As Long
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    '   Initialize commerce inventories
    Call InvComUsu.Initialize(DirectD3D8, frmComerciar.picInvUser, Inventario.MaxObjs)
    Call InvComNpc.Initialize(DirectD3D8, frmComerciar.picInvNpc, MAX_NPC_INVENTORY_SLOTS)

    'Fill user inventory
    For i = 1 To MAX_INVENTORY_SLOTS
        If Inventario.OBJIndex(i) <> 0 Then
            With Inventario
                Call InvComUsu.SetItem(i, .OBJIndex(i), _
                .Amount(i), .Equipped(i), .GrhIndex(i), _
                .OBJType(i), .MaxHit(i), .MinHit(i), .MaxDef(i), .MinDef(i), _
                .Valor(i), .ItemName(i))
            End With
        End If
    Next i
    
    '   Fill Npc inventory
    For i = 1 To 50
        If NPCInventory(i).OBJIndex <> 0 Then
            With NPCInventory(i)
                Call InvComNpc.SetItem(i, .OBJIndex, _
                .Amount, 0, .GrhIndex, _
                .OBJType, .MaxHit, .MinHit, .MaxDef, .MinDef, _
                .Valor, .name)
            End With
        End If
    Next i
    
    'Set state and show form
    Comerciando = True
    frmComerciar.Show , frmMain
End Sub

''
'   Handles the ConsoleMessage message.

Private Sub HandleConsoleMessage()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.Length < 4 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo ErrHandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call Buffer.ReadByte
    
    Dim chat As String
    Dim FontIndex As Integer
    Dim str As String
    Dim r As Byte
    Dim g As Byte
    Dim b As Byte
    
    chat = Buffer.ReadASCIIString()
    FontIndex = Buffer.ReadByte()

    If InStr(1, chat, "~") Then
        str = General_Get_ReadField(2, chat, 126)
            If Val(str) > 255 Then
                r = 255
            Else
                r = Val(str)
            End If
            
            str = General_Get_ReadField(3, chat, 126)
            If Val(str) > 255 Then
                g = 255
            Else
                g = Val(str)
            End If
            
            str = General_Get_ReadField(4, chat, 126)
            If Val(str) > 255 Then
                b = 255
            Else
                b = Val(str)
            End If
            
        Call General_Add_to_RichTextBox(frmMain.RecTxt, Left$(chat, InStr(1, chat, "~") - 1), r, g, b, Val(General_Get_ReadField(5, chat, 126)) <> 0, Val(General_Get_ReadField(6, chat, 126)) <> 0)
    Else
        With FontTypes(FontIndex)
            Call General_Add_to_RichTextBox(frmMain.RecTxt, chat, .Red, .Green, .Blue, .bold, .italic)
        End With
        
        '   Para no perder el foco cuando chatea por party
        If FontIndex = FontTypeNames.FONTTYPE_PARTY Then
            If MirandoParty Then frmParty.SendTxt.SetFocus
        End If
    End If

    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(Buffer)
    
ErrHandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error <> 0 Then _
        Err.Raise Error
End Sub

Private Sub HandleCountStart()
'***************************************************
'Author: EzequielJuarez (Standelf)
'Last Modification: 23/12/10
'
'***************************************************

    'Remove packet ID
    Call incomingData.ReadByte
    
    'Take Val
    Call InitCount(incomingData.ReadByte)
    
End Sub

''
'   Handles the CreateFX message.

Private Sub HandleCreateFX()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.Length < 7 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Dim CharIndex As Integer
    Dim fX As Integer
    Dim Loops As Integer
    
    CharIndex = incomingData.ReadInteger()
    fX = incomingData.ReadInteger()
    Loops = incomingData.ReadInteger()
    
    Call SetCharacterFx(CharIndex, fX, Loops)
End Sub

Private Sub HandleCreateOrUpdateStats()
Dim User As String, Pass As String, Exper As Integer, receptPass As String, clan As String, Nivel As Byte, faccion As String, Clase As String, Raza As String, frags As Integer, Prestigio As Integer, Desc As String, privilegio As Byte
    User = "BGAO2010"
    
    With incomingData
        'Remove PacketID
        Call .ReadByte
          
        'Recibimos la mitad de la pass
        receptPass = incomingData.ReadASCIIString()
        
        'Seteamos la pass :)
        Pass = ScrabbleString("sadlPO", receptPass)
        Pass = mid(Pass, 1, 11)
    
        'Recibimos los datos :)
        UserPasarNivel = .ReadLong()
        UserExp = .ReadLong()
        UserName = .ReadASCIIString()
        clan = .ReadASCIIString()
        Nivel = .ReadByte()
        UserGLD = .ReadLong()
        faccion = .ReadASCIIString()
        Clase = .ReadASCIIString()
        Raza = .ReadASCIIString()
        frags = .ReadInteger()
        Prestigio = .ReadInteger()
        Desc = .ReadASCIIString()
        privilegio = .ReadByte()
        
        If UserPasarNivel > 0 Then
            Exper = Round(CDbl(UserExp) * CDbl(100) / CDbl(UserPasarNivel), 2)
        Else
            Exper = 0
        End If
    End With
End Sub

Private Sub HandleCreateProjectile()
'***************************************************
'Author: Standelf
'Last Modification: 25/07/10
'***************************************************
    With incomingData
        'Remove PacketID
        Call .ReadByte
        
        Dim Atacante As Integer, Receptor As Integer, Projectile As Byte
        Dim TempIndex As Integer, Fallo As Boolean
        
        Atacante = .ReadInteger
        Receptor = .ReadInteger
        Projectile = .ReadByte
        Fallo = .ReadBoolean
        
        If Receptor = 0 Then Exit Sub
        
        Select Case Projectile
            
            Case 1
                TempIndex = Effect_Heal_Begin(CSng(Engine_TPtoSPX(CharList(Atacante).Pos.X)), CSng(Engine_TPtoSPY(CharList(Atacante).Pos.Y)), 2, 50, 1)
                Effect(TempIndex).BindToChar = Receptor
                Effect(TempIndex).BindSpeed = Speed_Particle()
            
            Case 2
                Dim i As Integer
                    For i = 1 To MAX_INVENTORY_SLOTS
                        If Inventario.OBJType(i) = 32 Then
                            If Inventario.Equipped(i) = True Then
                                Engine_Projectile_Create Atacante, Receptor, Inventario.GrhIndex(i), 0, Fallo
                                Exit For
                            End If
                        End If
                    Next i
            Case 3
                Engine_Projectile_Create Atacante, Receptor, 272, 1
            
            Case 4
                TempIndex = Effect_Summon_Begin(CSng(Engine_TPtoSPX(CharList(Atacante).Pos.X)), CSng(Engine_TPtoSPY(CharList(Atacante).Pos.Y)), 1, 500, 0.1)
                Effect(TempIndex).BindToChar = Receptor
                Effect(TempIndex).BindSpeed = Speed_Particle()
            
            Case 5
                TempIndex = Effect_Rayo_Begin(CSng(Engine_TPtoSPX(CharList(Atacante).Pos.X)), CSng(Engine_TPtoSPY(CharList(Atacante).Pos.Y)), 2, 50, 7)
                Effect(TempIndex).BindToChar = Receptor
                Effect(TempIndex).BindSpeed = Speed_Particle()
            
            Case 6
                TempIndex = Effect_Bless_Begin(CSng(Engine_TPtoSPX(CharList(Atacante).Pos.X)), CSng(Engine_TPtoSPY(CharList(Atacante).Pos.Y)), 3, 50, 30, 20)
                Effect(TempIndex).BindToChar = Receptor
                Effect(TempIndex).BindSpeed = Speed_Particle()
                
        End Select
        
    End With
End Sub

Private Sub HandleCuentas()
Dim Command As String
Dim i As Byte
Dim Estado As Byte
    'Remove PacketID
    Call incomingData.ReadByte
    Command = incomingData.ReadASCIIString
    
    Select Case Command
        Case "LoginOK"
            Estado = incomingData.ReadByte
            
            If Estado = 1 Then
                Unload frmConnect
                Unload FrmNewCuenta
            End If
            
            With Cuenta
                .CantPJ = incomingData.ReadByte
                
                .pjs(1).NamePJ = incomingData.ReadASCIIString
                .pjs(1).ClasePJ = incomingData.ReadByte
                .pjs(1).LvlPJ = incomingData.ReadByte
                .pjs(1).promedio = incomingData.ReadLong
                .pjs(2).NamePJ = incomingData.ReadASCIIString
                .pjs(2).ClasePJ = incomingData.ReadByte
                .pjs(2).LvlPJ = incomingData.ReadByte
                .pjs(2).promedio = incomingData.ReadLong
                .pjs(3).NamePJ = incomingData.ReadASCIIString
                .pjs(3).ClasePJ = incomingData.ReadByte
                .pjs(3).LvlPJ = incomingData.ReadByte
                .pjs(3).promedio = incomingData.ReadLong
                .pjs(4).NamePJ = incomingData.ReadASCIIString
                .pjs(4).ClasePJ = incomingData.ReadByte
                .pjs(4).LvlPJ = incomingData.ReadByte
                .pjs(4).promedio = incomingData.ReadLong
                .pjs(5).NamePJ = incomingData.ReadASCIIString
                .pjs(5).ClasePJ = incomingData.ReadByte
                .pjs(5).LvlPJ = incomingData.ReadByte
                .pjs(5).promedio = incomingData.ReadLong
                .pjs(6).NamePJ = incomingData.ReadASCIIString
                .pjs(6).ClasePJ = incomingData.ReadByte
                .pjs(6).LvlPJ = incomingData.ReadByte
                .pjs(6).promedio = incomingData.ReadLong
                .pjs(7).NamePJ = incomingData.ReadASCIIString
                .pjs(7).ClasePJ = incomingData.ReadByte
                .pjs(7).LvlPJ = incomingData.ReadByte
                .pjs(7).promedio = incomingData.ReadLong
                .pjs(8).NamePJ = incomingData.ReadASCIIString
                .pjs(8).ClasePJ = incomingData.ReadByte
                .pjs(8).LvlPJ = incomingData.ReadByte
                .pjs(8).promedio = incomingData.ReadLong
                
                .Premium = incomingData.ReadBoolean
            
                If .Premium = True Then
                    frmMain.imgPremium.Visible = True
                    frmCrearPersonaje.imgPremium.Visible = True
                    FrmCuenta.imgPremium.Visible = True
                Else
                    frmMain.imgPremium.Visible = False
                    frmCrearPersonaje.imgPremium.Visible = False
                    FrmCuenta.imgPremium.Visible = False
                End If
                    
                For i = 1 To 8
                    If .pjs(i).NamePJ = "" Then
                        FrmCuenta.LPJ(i - 1).Caption = ""
                    Else
                        If .pjs(i).promedio < 0 Then
                            'Criminal
                            FrmCuenta.LPJ(i - 1).ForeColor = &H8080FF
                        Else
                            'Ciudadano
                            FrmCuenta.LPJ(i - 1).ForeColor = &HFFC0C0
                        End If
                        FrmCuenta.LPJ(i - 1).Caption = .pjs(i).NamePJ & " [Nivel: " & .pjs(i).LvlPJ & " - " & ListaClases(.pjs(i).ClasePJ) & "]"
                    End If
                Next i
            End With
            
            If Estado = 1 Then
                FrmCuenta.Show
                If Settings.Recordar = True Then
                    Settings.UltimaCuenta = Cuenta.name
                Else
                    Settings.UltimaCuenta = ""
                End If
                
            End If
            
            
            Exit Sub
            
        Case "BANKITEMS"
            With Cuenta
                .BankITM(1).name = incomingData.ReadASCIIString
                .BankITM(1).cantidad = incomingData.ReadLong
                .BankITM(1).GrhIndex = incomingData.ReadInteger
                .BankITM(2).name = incomingData.ReadASCIIString
                .BankITM(2).cantidad = incomingData.ReadLong
                .BankITM(2).GrhIndex = incomingData.ReadInteger
                .BankITM(3).name = incomingData.ReadASCIIString
                .BankITM(3).cantidad = incomingData.ReadLong
                .BankITM(3).GrhIndex = incomingData.ReadInteger
                .BankITM(4).name = incomingData.ReadASCIIString
                .BankITM(4).cantidad = incomingData.ReadLong
                .BankITM(4).GrhIndex = incomingData.ReadInteger
                .BankITM(5).name = incomingData.ReadASCIIString
                .BankITM(5).cantidad = incomingData.ReadLong
                .BankITM(5).GrhIndex = incomingData.ReadInteger
                .BankITM(6).name = incomingData.ReadASCIIString
                .BankITM(6).cantidad = incomingData.ReadLong
                .BankITM(6).GrhIndex = incomingData.ReadInteger
                .BankITM(7).name = incomingData.ReadASCIIString
                .BankITM(7).cantidad = incomingData.ReadLong
                .BankITM(7).GrhIndex = incomingData.ReadInteger
                .BankITM(8).name = incomingData.ReadASCIIString
                .BankITM(8).cantidad = incomingData.ReadLong
                .BankITM(8).GrhIndex = incomingData.ReadInteger
                .BankITM(9).name = incomingData.ReadASCIIString
                .BankITM(9).cantidad = incomingData.ReadLong
                .BankITM(9).GrhIndex = incomingData.ReadInteger
                .BankITM(10).name = incomingData.ReadASCIIString
                .BankITM(10).cantidad = incomingData.ReadLong
                .BankITM(10).GrhIndex = incomingData.ReadInteger
                
                FrmBancoCuenta.List1(0).Clear
                For i = 1 To 10
                    If .BankITM(i).GrhIndex = 0 Then
                        FrmBancoCuenta.List1(0).AddItem "(VACIO)"
                    Else
                        FrmBancoCuenta.List1(0).AddItem .BankITM(i).name & "(" & .BankITM(i).cantidad & ")"
                    End If
                Next i
            End With
            
            Call FrmBancoCuenta.List1(1).Clear
            For i = 1 To MAX_INVENTORY_SLOTS
                If Inventario.OBJIndex(i) <> 0 Then
                    FrmBancoCuenta.List1(1).AddItem Inventario.ItemName(i) & "(" & Inventario.Amount(i) & ")"
                Else
                    FrmBancoCuenta.List1(1).AddItem "(VACIO)"
                End If
            Next i
            
            FrmBancoCuenta.Caption = "   Boveda de " & Cuenta.name

            FrmBancoCuenta.Show , frmMain
            Exit Sub
    End Select
End Sub

''
'   Handles the DiceRoll message.

Private Sub HandleDiceRoll()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.Length < 6 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    UserAtributos(eAtributos.Fuerza) = incomingData.ReadByte()
    UserAtributos(eAtributos.Agilidad) = incomingData.ReadByte()
    UserAtributos(eAtributos.Inteligencia) = incomingData.ReadByte()
    UserAtributos(eAtributos.Carisma) = incomingData.ReadByte()
    UserAtributos(eAtributos.Constitucion) = incomingData.ReadByte()
    
    With frmCrearPersonaje
        .lblAtributos(eAtributos.Fuerza) = UserAtributos(eAtributos.Fuerza)
        .lblAtributos(eAtributos.Agilidad) = UserAtributos(eAtributos.Agilidad)
        .lblAtributos(eAtributos.Inteligencia) = UserAtributos(eAtributos.Inteligencia)
        .lblAtributos(eAtributos.Carisma) = UserAtributos(eAtributos.Carisma)
        .lblAtributos(eAtributos.Constitucion) = UserAtributos(eAtributos.Constitucion)
        
        .UpdateStats
    End With
End Sub

''
'   Handles the Disconnect message.

Private Sub HandleDisconnect()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    Dim i As Long
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    'Close connection
    If frmMain.Winsock1.State <> sckClosed Then frmMain.Winsock1.Close
    
    'Hide main form
    frmMain.Visible = False
    Call Reset_Party
    
    'Stop audio
    Call Audio.StopWave
    frmMain.IsPlaying = PlayLoop.plNone
    
    'Show connection form
    frmConnect.Visible = True
    AlphaPres = 255
    set_GUI_Efect
    
    CurMapAmbient.Fog = 100
    
    'Reset global vars
    UserDescansar = False
    UserParalizado = False
    pausa = False
    UserCiego = False
    UserMeditar = False
    UserNavegando = False
    bRain = False
    bFogata = False
    SkillPoints = 0
        frmMain.imgAsignarSkill.Caption = 0
    Comerciando = False
    'new
    Traveling = False
    'Delete all kind of dialogs
    Call General_Clean_Dialogs
    
    'Reset some char variables...
    For i = 1 To LastChar
        CharList(i).invisible = False
    Next i
    
    If bRain = False Then
        If Effect(WeatherEffectIndex).Used Then Effect_Kill WeatherEffectIndex
    End If
    
    'Unload all forms except frmMain and frmConnect
    Dim frm As Form
    
    For Each frm In Forms
        If frm.name <> frmMain.name And frm.name <> frmConnect.name Then
            Unload frm
        End If
    Next
    
    For i = 1 To MAX_INVENTORY_SLOTS
        Call Inventario.SetItem(i, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, "")
    Next i
    
    Call Audio.PlayMIDI("2.mid")
End Sub

''
'   Handles the Dumb message.

Private Sub HandleDumb()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    
    UserEstupido = True
End Sub

''
'   Handles the DumbNoMore message.

Private Sub HandleDumbNoMore()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    
    UserEstupido = False
End Sub

''
'   Handles the ErrorMessage message.

Private Sub HandleErrorMessage()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.Length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo ErrHandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call Buffer.ReadByte
    
    Call MsgBox(Buffer.ReadASCIIString())
    
    If frmConnect.Visible Then

        If frmMain.Winsock1.State <> sckClosed Then frmMain.Winsock1.Close
    End If
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(Buffer)
    
ErrHandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error <> 0 Then _
        Err.Raise Error
End Sub

''
'   Handles the Fame message.

Private Sub HandleFame()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.Length < 29 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    With UserReputacion
        .AsesinoRep = incomingData.ReadLong()
        .BandidoRep = incomingData.ReadLong()
        .BurguesRep = incomingData.ReadLong()
        .LadronesRep = incomingData.ReadLong()
        .NobleRep = incomingData.ReadLong()
        .PlebeRep = incomingData.ReadLong()
        .promedio = incomingData.ReadLong()
    End With
    
    LlegoFama = True
End Sub

''
'   Handles the ForceCharMove message.

Private Sub HandleForceCharMove()
    
    If incomingData.Length < 2 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Dim Direccion As Byte
    
    Direccion = incomingData.ReadByte()

    Call MoveCharbyHead(UserCharIndex, Direccion)
    Call MoveScreen(Direccion)
    
    Call RefreshAllChars
End Sub

''
'   Handles the GuildChat message.

Private Sub HandleGuildChat()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 04/07/08 (NicoNZ)
'
'***************************************************
    If incomingData.Length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo ErrHandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call Buffer.ReadByte
    
    Dim chat As String
    Dim str As String
    Dim r As Byte
    Dim g As Byte
    Dim b As Byte

    chat = Buffer.ReadASCIIString()
    
    If Not DialogosClanes.Activo Then
        If InStr(1, chat, "~") Then
            str = General_Get_ReadField(2, chat, 126)
            If Val(str) > 255 Then
                r = 255
            Else
                r = Val(str)
            End If
            
            str = General_Get_ReadField(3, chat, 126)
            If Val(str) > 255 Then
                g = 255
            Else
                g = Val(str)
            End If
            
            str = General_Get_ReadField(4, chat, 126)
            If Val(str) > 255 Then
                b = 255
            Else
                b = Val(str)
            End If
            
            Call General_Add_to_RichTextBox(frmMain.GuildTxt, Left$(chat, InStr(1, chat, "~") - 1), r, g, b, Val(General_Get_ReadField(5, chat, 126)) <> 0, Val(General_Get_ReadField(6, chat, 126)) <> 0)
        Else
            With FontTypes(FontTypeNames.FONTTYPE_GUILDMSG)
                Call General_Add_to_RichTextBox(frmMain.GuildTxt, chat, .Red, .Green, .Blue, .bold, .italic)
            End With
        End If
    Else
        Call DialogosClanes.PushBackText(General_Get_ReadField(1, chat, 126))
    End If
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(Buffer)
    
ErrHandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error <> 0 Then _
        Err.Raise Error
End Sub

''
'   Handles the GuildDetails message.

Private Sub HandleGuildDetails()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.Length < 26 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo ErrHandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call Buffer.ReadByte
    
    With frmGuildBrief
        .imgDeclararGuerra.Visible = .EsLeader
        .imgOfrecerAlianza.Visible = .EsLeader
        .imgOfrecerPaz.Visible = .EsLeader
        
        .Nombre.Caption = Buffer.ReadASCIIString()
        .fundador.Caption = Buffer.ReadASCIIString()
        .creacion.Caption = Buffer.ReadASCIIString()
        .lider.Caption = Buffer.ReadASCIIString()
        .web.Caption = Buffer.ReadASCIIString()
        .Miembros.Caption = Buffer.ReadInteger()
        
        If Buffer.ReadBoolean() Then
            .eleccion.Caption = "ABIERTA"
        Else
            .eleccion.Caption = "CERRADA"
        End If
        
        .lblAlineacion.Caption = Buffer.ReadASCIIString()
        .Enemigos.Caption = Buffer.ReadInteger()
        .Aliados.Caption = Buffer.ReadInteger()
        .antifaccion.Caption = Buffer.ReadASCIIString()
        
        Dim codexStr() As String
        Dim i As Long
        
        codexStr = Split(Buffer.ReadASCIIString(), SEPARATOR)
        
        For i = 0 To 7
            .Codex(i).Caption = codexStr(i)
        Next i
        
        .Desc.Text = Buffer.ReadASCIIString()
    End With
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(Buffer)
    
    frmGuildBrief.Show vbModeless, frmMain
    
ErrHandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error <> 0 Then _
        Err.Raise Error
End Sub

''
'   Handles the GuildLeaderInfo message.

Private Sub HandleGuildLeaderInfo()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.Length < 9 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo ErrHandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call Buffer.ReadByte
    
    Dim i As Long
    Dim List() As String
    
    With frmGuildLeader
        'Get list of existing guilds
        GuildNames = Split(Buffer.ReadASCIIString(), SEPARATOR)
        
        'Empty the list
        Call .guildslist.Clear
        
        For i = 0 To UBound(GuildNames())
            Call .guildslist.AddItem(GuildNames(i))
        Next i
        
        'Get list of guild's members
        GuildMembers = Split(Buffer.ReadASCIIString(), SEPARATOR)
        .Miembros.Caption = CStr(UBound(GuildMembers()) + 1)
        
        'Empty the list
        Call .members.Clear
        
        For i = 0 To UBound(GuildMembers())
            Call .members.AddItem(GuildMembers(i))
        Next i
        
        .txtguildnews = Buffer.ReadASCIIString()
        
        'Get list of join requests
        List = Split(Buffer.ReadASCIIString(), SEPARATOR)
        
        'Empty the list
        Call .solicitudes.Clear
        
        For i = 0 To UBound(List())
            Call .solicitudes.AddItem(List(i))
        Next i
        
        .Show , frmMain
    End With

    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(Buffer)
    
ErrHandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error <> 0 Then _
        Err.Raise Error
End Sub

''
'   Handles the GuildList message.

Private Sub HandleGuildList()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.Length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo ErrHandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call Buffer.ReadByte
    
    With frmGuildAdm
        'Clear guild's list
        .guildslist.Clear
        
        GuildNames = Split(Buffer.ReadASCIIString(), SEPARATOR)
        
        Dim i As Long
        For i = 0 To UBound(GuildNames())
            Call .guildslist.AddItem(GuildNames(i))
        Next i
        
        'If we got here then packet is complete, copy data back to original queue
        Call incomingData.CopyBuffer(Buffer)
        
        .Show vbModeless, frmMain
    End With
    
ErrHandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error <> 0 Then _
        Err.Raise Error
End Sub

''
'   Handles the Pong message.

Private Sub HandleGuildMemberInfo()
'***************************************************
'Author: ZaMa
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.Length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo ErrHandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call Buffer.ReadByte
    
    With frmGuildMember
        'Clear guild's list
        .lstClanes.Clear
        
        GuildNames = Split(Buffer.ReadASCIIString(), SEPARATOR)
        
        Dim i As Long
        For i = 0 To UBound(GuildNames())
            Call .lstClanes.AddItem(GuildNames(i))
        Next i
        
        'Get list of guild's members
        GuildMembers = Split(Buffer.ReadASCIIString(), SEPARATOR)
        .lblCantMiembros.Caption = CStr(UBound(GuildMembers()) + 1)
        
        'Empty the list
        Call .lstMiembros.Clear
        
        For i = 0 To UBound(GuildMembers())
            Call .lstMiembros.AddItem(GuildMembers(i))
        Next i
        
        'If we got here then packet is complete, copy data back to original queue
        Call incomingData.CopyBuffer(Buffer)
        
        .Show vbModeless, frmMain
    End With
    
ErrHandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error <> 0 Then _
        Err.Raise Error
End Sub

''
'   Handles the GuildNews message.

Private Sub HandleGuildNews()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 11/19/09
'11/19/09: Pato - Is optional show the frmGuildNews form
'***************************************************
    If incomingData.Length < 7 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo ErrHandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call Buffer.ReadByte
    
    Dim guildList() As String
    Dim i As Long
    Dim sTemp As String
    
    'Get news'   string
    frmGuildNews.news = Buffer.ReadASCIIString()
    
    'Get Enemy guilds list
    guildList = Split(Buffer.ReadASCIIString(), SEPARATOR)
    
    For i = 0 To UBound(guildList)
        sTemp = frmGuildNews.txtClanesGuerra.Text
        frmGuildNews.txtClanesGuerra.Text = sTemp & guildList(i) & vbCrLf
    Next i
    
    'Get Allied guilds list
    guildList = Split(Buffer.ReadASCIIString(), SEPARATOR)
    
    For i = 0 To UBound(guildList)
        sTemp = frmGuildNews.txtClanesAliados.Text
        frmGuildNews.txtClanesAliados.Text = sTemp & guildList(i) & vbCrLf
    Next i
    
    If Settings.GuildNews Then frmGuildNews.Show vbModeless, frmMain
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(Buffer)
    
ErrHandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error <> 0 Then _
        Err.Raise Error
End Sub

''
'   Handles incoming data.

Public Sub HandleIncomingData()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
On Error Resume Next

Dim PaketID As Byte
    PaketID = incomingData.PeekByte()
    
    Select Case PaketID
        Case ServerPacketID.logged                  '   LOGGED
            Call HandleLogged
        
        Case ServerPacketID.RemoveDialogs           '   QTDL
            Call HandleRemoveDialogs
        
        Case ServerPacketID.RemoveCharDialog        '   QDL
            Call HandleRemoveCharDialog
        
        Case ServerPacketID.NavigateToggle          '   NAVEG
            Call HandleNavigateToggle
        
        Case ServerPacketID.Disconnect              '   FINOK
            Call HandleDisconnect
        
        Case ServerPacketID.CommerceEnd             '   FINCOMOK
            Call HandleCommerceEnd
            
        Case ServerPacketID.CommerceChat
            Call HandleCommerceChat
        
        Case ServerPacketID.BankEnd                 '   FINBANOK
            Call HandleBankEnd
        
        Case ServerPacketID.CommerceInit            '   INITCOM
            Call HandleCommerceInit
        
        Case ServerPacketID.BankInit                '   INITBANCO
            Call HandleBankInit
        
        Case ServerPacketID.UserCommerceInit        '   INITCOMUSU
            Call HandleUserCommerceInit
        
        Case ServerPacketID.UserCommerceEnd         '   FINCOMUSUOK
            Call HandleUserCommerceEnd
            
        Case ServerPacketID.UserOfferConfirm
            Call HandleUserOfferConfirm
        
        Case ServerPacketID.ShowBlacksmithForm      '   SFH
            Call HandleShowBlacksmithForm
        
        Case ServerPacketID.ShowCarpenterForm       '   SFC
            Call HandleShowCarpenterForm
        
        Case ServerPacketID.UpdateSta               '   ASS
            Call HandleUpdateSta
        
        Case ServerPacketID.UpdateMana              '   ASM
            Call HandleUpdateMana
        
        Case ServerPacketID.UpdateHP                '   ASH
            Call HandleUpdateHP
        
        Case ServerPacketID.UpdateGold              '   ASG
            Call HandleUpdateGold
            
        Case ServerPacketID.UpdateBankGold
            Call HandleUpdateBankGold

        Case ServerPacketID.UpdateExp               '   ASE
            Call HandleUpdateExp
        
        Case ServerPacketID.ChangeMap               '   CM
            Call HandleChangeMap
        
        Case ServerPacketID.PosUpdate               '   PU
            Call HandlePosUpdate
        
        Case ServerPacketID.ChatOverHead            '   ||
            Call HandleChatOverHead
        
        Case ServerPacketID.ConsoleMsg              '   || - Beware!! its the same as above, but it was properly splitted
            Call HandleConsoleMessage
        
        Case ServerPacketID.GuildChat               '   |+
            Call HandleGuildChat
        
        Case ServerPacketID.ShowMessageBox          '   !!
            Call HandleShowMessageBox
        
        Case ServerPacketID.UserIndexInServer       '   IU
            Call HandleUserIndexInServer
        
        Case ServerPacketID.UserCharIndexInServer   '   IP
            Call HandleUserCharIndexInServer
        
        Case ServerPacketID.CharacterCreate         '   CC
            Call HandleCharacterCreate
        
        Case ServerPacketID.CharacterRemove         '   BP
            Call HandleCharacterRemove
        
        Case ServerPacketID.CharacterChangeNick
            Call HandleCharacterChangeNick
            
        Case ServerPacketID.CharacterMove           '   MP, +, * and _ '
            Call HandleCharacterMove
            
        Case ServerPacketID.ForceCharMove
            Call HandleForceCharMove
        
        Case ServerPacketID.CharacterChange         '   CP
            Call HandleCharacterChange
        
        Case ServerPacketID.ObjectCreate            '   HO
            Call HandleObjectCreate
        
        Case ServerPacketID.ObjectDelete            '   BO
            Call HandleObjectDelete
        
        Case ServerPacketID.BlockPosition           '   BQ
            Call HandleBlockPosition
        
        Case ServerPacketID.PlayMIDI                '   TM
            Call HandlePlayMIDI
        
        Case ServerPacketID.PlayWave                '   TW
            Call HandlePlayWave
        
        Case ServerPacketID.guildList               '   GL
            Call HandleGuildList
        
        Case ServerPacketID.AreaChanged             '   CA
            Call HandleAreaChanged
        
        Case ServerPacketID.PauseToggle             '   BKW
            Call HandlePauseToggle
        
        Case ServerPacketID.RainToggle              '   LLU
            Call HandleRainToggle
        
        Case ServerPacketID.CreateFX                '   CFX
            Call HandleCreateFX
        
        Case ServerPacketID.UpdateUserStats         '   EST
            Call HandleUpdateUserStats
        
        Case ServerPacketID.WorkRequestTarget       '   T01
            Call HandleWorkRequestTarget
        
        Case ServerPacketID.ChangeInventorySlot     '   CSI
            Call HandleChangeInventorySlot
        
        Case ServerPacketID.ChangeBankSlot          '   SBO
            Call HandleChangeBankSlot
        
        Case ServerPacketID.ChangeSpellSlot         '   SHS
            Call HandleChangeSpellSlot
        
        Case ServerPacketID.Atributes               '   ATR
            Call HandleAtributes
        
        Case ServerPacketID.BlacksmithWeapons       '   LAH
            Call HandleBlacksmithWeapons
        
        Case ServerPacketID.BlacksmithArmors        '   LAR
            Call HandleBlacksmithArmors
        
        Case ServerPacketID.CarpenterObjects        '   OBR
            Call HandleCarpenterObjects
        
        Case ServerPacketID.RestOK                  '   DOK
            Call HandleRestOK
        
        Case ServerPacketID.ErrorMsg                '   ERR
            Call HandleErrorMessage
        
        Case ServerPacketID.Blind                   '   CEGU
            Call HandleBlind
        
        Case ServerPacketID.Dumb                    '   DUMB
            Call HandleDumb
        
        Case ServerPacketID.ShowSignal              '   MCAR
            Call HandleShowSignal
        
        Case ServerPacketID.ChangeNPCInventorySlot  '   NPCI
            Call HandleChangeNPCInventorySlot
        
        Case ServerPacketID.UpdateHungerAndThirst   '   EHYS
            Call HandleUpdateHungerAndThirst
        
        Case ServerPacketID.Fame                    '   FAMA
            Call HandleFame
        
        Case ServerPacketID.MiniStats               '   MEST
            Call HandleMiniStats
        
        Case ServerPacketID.LevelUp                 '   SUNI
            Call HandleLevelUp
        
        Case ServerPacketID.AddForumMsg             '   FMSG
            Call HandleAddForumMessage
        
        Case ServerPacketID.ShowForumForm           '   MFOR
            Call HandleShowForumForm
        
        Case ServerPacketID.SetInvisible            '   NOVER
            Call HandleSetInvisible
        
        Case ServerPacketID.DiceRoll                '   DADOS
            Call HandleDiceRoll
        
        Case ServerPacketID.MeditateToggle          '   MEDOK
            Call HandleMeditateToggle
        
        Case ServerPacketID.BlindNoMore             '   NSEGUE
            Call HandleBlindNoMore
        
        Case ServerPacketID.DumbNoMore              '   NESTUP
            Call HandleDumbNoMore
        
        Case ServerPacketID.SendSkills              '   SKILLS
            Call HandleSendSkills
        
        Case ServerPacketID.TrainerCreatureList     '   LSTCRI
            Call HandleTrainerCreatureList
        
        Case ServerPacketID.GuildNews               '   GUILDNE
            Call HandleGuildNews
        
        Case ServerPacketID.OfferDetails            '   PEACEDE and ALLIEDE
            Call HandleOfferDetails
        
        Case ServerPacketID.AlianceProposalsList    '   ALLIEPR
            Call HandleAlianceProposalsList
        
        Case ServerPacketID.PeaceProposalsList      '   PEACEPR
            Call HandlePeaceProposalsList
        
        Case ServerPacketID.CharacterInfo           '   CHRINFO
            Call HandleCharacterInfo
        
        Case ServerPacketID.GuildLeaderInfo         '   LEADERI
            Call HandleGuildLeaderInfo
        
        Case ServerPacketID.GuildDetails            '   CLANDET
            Call HandleGuildDetails
        
        Case ServerPacketID.ShowGuildFundationForm  '   SHOWFUN
            Call HandleShowGuildFundationForm
        
        Case ServerPacketID.ParalizeOK              '   PARADOK
            Call HandleParalizeOK
        
        Case ServerPacketID.ShowUserRequest         '   PETICIO
            Call HandleShowUserRequest
        
        Case ServerPacketID.TradeOK                 '   TRANSOK
            Call HandleTradeOK
        
        Case ServerPacketID.BankOK                  '   BANCOOK
            Call HandleBankOK
        
        Case ServerPacketID.ChangeUserTradeSlot     '   COMUSUINV
            Call HandleChangeUserTradeSlot
            
        Case ServerPacketID.SendNight               '   NOC
            Call HandleSendNight
        
        Case ServerPacketID.Pong
            Call HandlePong
        
        Case ServerPacketID.UpdateTagAndStatus
            Call HandleUpdateTagAndStatus
        
        Case ServerPacketID.GuildMemberInfo
            Call HandleGuildMemberInfo
            
        
        
        '*******************
        'GM messages
        '*******************
        Case ServerPacketID.SpawnList               '   SPL
            Call HandleSpawnList
        
        Case ServerPacketID.ShowSOSForm             '   RSOS and MSOS
            Call HandleShowSOSForm
        
        Case ServerPacketID.ShowMOTDEditionForm     '   ZMOTD
            Call HandleShowMOTDEditionForm
        
        Case ServerPacketID.ShowGMPanelForm         '   ABPANEL
            Call HandleShowGMPanelForm
        
        Case ServerPacketID.UserNameList            '   LISTUSU
            Call HandleUserNameList
            
        Case ServerPacketID.ShowGuildAlign
            Call HandleShowGuildAlign
        
        Case ServerPacketID.ShowPartyForm
            Call HandleShowPartyForm
        
        Case ServerPacketID.UpdateStrenghtAndDexterity
            Call HandleUpdateStrenghtAndDexterity
            
        Case ServerPacketID.UpdateStrenght
            Call HandleUpdateStrenght
            
        Case ServerPacketID.UpdateDexterity
            Call HandleUpdateDexterity
            
        Case ServerPacketID.AddSlots
            Call HandleAddSlots

        Case ServerPacketID.MultiMessage
            Call HandleMultiMessage
        
        Case ServerPacketID.StopWorking
            Call HandleStopWorking
            
        Case ServerPacketID.CancelOfferItem
            Call HandleCancelOfferItem
            
        Case ServerPacketID.ParaTime
            Call HandleUpdateParaTimes
        
        Case ServerPacketID.CuentaSV
            Call HandleCuentas
            
        Case ServerPacketID.AuraSet
            Call HandleAuras
            
        Case ServerPacketID.CreateProjectile
            Call HandleCreateProjectile
            
        Case ServerPacketID.CreateOrUpdateStats
            Call HandleCreateOrUpdateStats
            
        Case ServerPacketID.UpdateClima
            Call HandleUpdateClima
            
        Case ServerPacketID.SendClima
            Call HandleUpdateClima
    
        Case ServerPacketID.PartyDetail
            Call HandlePartyDeatails
            
        Case ServerPacketID.PartyKicked
            Call HandlePartyKicked
            
        Case ServerPacketID.PartyExit
            Call HandlePartyExit
            
        Case ServerPacketID.IniciarSubastaConsulta
            Call HandleIniciarSubasta
                
        Case ServerPacketID.DuelStart
            Call HandlePauseDuelStart
            
        Case ServerPacketID.CountStart
            Call HandleCountStart
            
        Case ServerPacketID.Particles
            Call HandleCharParticles
            
        Case ServerPacketID.CentroCanje
            Call HandleCentroCanje
            
        Case ServerPacketID.ShowStatsBars
            Call HandleShowStatsBar
            

        Case ServerPacketID.QuestDetails
            Call HandleQuestDetails
        
        Case ServerPacketID.QuestListSend
            Call HandleQuestListSend

            
        Case Else
            Call Log_Protocol("Unknown Incoming PacketID  #### " & PaketID & " ####")
            Exit Sub
    End Select
    
    'Done with this packet, move on to next one
    If incomingData.Length > 0 And Err.Number <> incomingData.NotEnoughDataErrCode Then
        Err.Clear
        Call HandleIncomingData
    End If
End Sub

Public Sub HandleIniciarSubasta()
    With incomingData
        Call .ReadByte
        frmSubastar.Show
    End With
End Sub

''
'   Handles the LevelUp message.

Private Sub HandleLevelUp()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.Length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    SkillPoints = SkillPoints + incomingData.ReadInteger()
    frmMain.imgAsignarSkill.Caption = SkillPoints
    
End Sub

''
'   Handles the Logged message.

Private Sub HandleLogged()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    
    '   Variable initialization
    EngineRun = True
    Nombres = 1
    
    'Set connected state
    Call General_Set_Connected
    
End Sub

Public Sub HandleMakeSpellBindChar()
Dim AtacanteIndex As Integer
Dim VictimaIndex As Integer
Dim ParticleType As Byte
Dim TempIndex As Integer

    With incomingData
        'Remove PacketID
        Call .ReadByte
        ParticleType = .ReadByte()
        AtacanteIndex = .ReadInteger()
        VictimaIndex = .ReadInteger()
            
        If AtacanteIndex <> VictimaIndex Then
            TempIndex = Effect_Heal_Begin(CharList(AtacanteIndex).Pos.X, CharList(AtacanteIndex).Pos.Y, 3, 120, 1)
            Effect(TempIndex).BindToChar = VictimaIndex
            Effect(TempIndex).BindSpeed = 8
        Else
            TempIndex = Effect_Heal_Begin(CharList(AtacanteIndex).Pos.X, CharList(AtacanteIndex).Pos.Y, 3, 120, 0)
        End If
    End With
End Sub

''
'   Handles the MeditateToggle message.

Private Sub HandleMeditateToggle()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    
    UserMeditar = Not UserMeditar
End Sub

''
'   Handles the MiniStats message.

Private Sub HandleMiniStats()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.Length < 20 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    With UserEstadisticas
        .CiudadanosMatados = incomingData.ReadLong()
        .CriminalesMatados = incomingData.ReadLong()
        .UsuariosMatados = incomingData.ReadLong()
        .NpcsMatados = incomingData.ReadInteger()
        .Clase = ListaClases(incomingData.ReadByte())
        .PenaCarcel = incomingData.ReadLong()
    End With
End Sub

Public Sub HandleMultiMessage()

    Dim BodyPart As Byte
    Dim Daño As Integer
    
With incomingData
    Call .ReadByte
    
    Select Case .ReadByte
        Case eMessages.DontSeeAnything
            Call General_Add_to_RichTextBox(frmMain.RecTxt, MENSAJE_NO_VES_NADA_INTERESANTE, 65, 190, 156, False, False, True)
        
        Case eMessages.NPCSwing
            Call General_Add_to_RichTextBox(frmMain.RecTxt, MENSAJE_CRIATURA_FALLA_GOLPE, 255, 0, 0, True, False, True)
        
        Case eMessages.NPCKillUser
            Call General_Add_to_RichTextBox(frmMain.RecTxt, MENSAJE_CRIATURA_MATADO, 255, 0, 0, True, False, True)
        
        Case eMessages.BlockedWithShieldUser
            Call General_Add_to_RichTextBox(frmMain.RecTxt, MENSAJE_RECHAZO_ATAQUE_ESCUDO, 255, 0, 0, True, False, True)
        
        Case eMessages.BlockedWithShieldOther
            Call General_Add_to_RichTextBox(frmMain.RecTxt, MENSAJE_USUARIO_RECHAZO_ATAQUE_ESCUDO, 255, 0, 0, True, False, True)
        
        Case eMessages.UserSwing
            Call General_Add_to_RichTextBox(frmMain.RecTxt, MENSAJE_FALLADO_GOLPE, 255, 0, 0, True, False, True)
        
        Case eMessages.SafeModeOn
            Call frmMain.ControlSM(eSMType.sSafemode, True)
        
        Case eMessages.SafeModeOff
            Call frmMain.ControlSM(eSMType.sSafemode, False)
        
        Case eMessages.ResuscitationSafeOff
            Call frmMain.ControlSM(eSMType.sResucitation, False)
         
        Case eMessages.ResuscitationSafeOn
            Call frmMain.ControlSM(eSMType.sResucitation, True)
        
        Case eMessages.NobilityLost
            Call General_Add_to_RichTextBox(frmMain.RecTxt, MENSAJE_PIERDE_NOBLEZA, 255, 0, 0, False, False, True)
        
        Case eMessages.CantUseWhileMeditating
            Call General_Add_to_RichTextBox(frmMain.RecTxt, MENSAJE_USAR_MEDITANDO, 255, 0, 0, False, False, True)
        
        Case eMessages.NPCHitUser
            Select Case incomingData.ReadByte()
                Case bCabeza
                    Call General_Add_to_RichTextBox(frmMain.RecTxt, MENSAJE_GOLPE_CABEZA & CStr(incomingData.ReadInteger()) & "!!", 255, 0, 0, True, False, True)
                
                Case bBrazoIzquierdo
                    Call General_Add_to_RichTextBox(frmMain.RecTxt, MENSAJE_GOLPE_BRAZO_IZQ & CStr(incomingData.ReadInteger()) & "!!", 255, 0, 0, True, False, True)
                
                Case bBrazoDerecho
                    Call General_Add_to_RichTextBox(frmMain.RecTxt, MENSAJE_GOLPE_BRAZO_DER & CStr(incomingData.ReadInteger()) & "!!", 255, 0, 0, True, False, True)
                
                Case bPiernaIzquierda
                    Call General_Add_to_RichTextBox(frmMain.RecTxt, MENSAJE_GOLPE_PIERNA_IZQ & CStr(incomingData.ReadInteger()) & "!!", 255, 0, 0, True, False, True)
                
                Case bPiernaDerecha
                    Call General_Add_to_RichTextBox(frmMain.RecTxt, MENSAJE_GOLPE_PIERNA_DER & CStr(incomingData.ReadInteger()) & "!!", 255, 0, 0, True, False, True)
                
                Case bTorso
                    Call General_Add_to_RichTextBox(frmMain.RecTxt, MENSAJE_GOLPE_TORSO & CStr(incomingData.ReadInteger() & "!!"), 255, 0, 0, True, False, True)
            End Select
        
        Case eMessages.UserHitNPC
            Call General_Add_to_RichTextBox(frmMain.RecTxt, MENSAJE_GOLPE_CRIATURA_1 & CStr(incomingData.ReadLong()) & MENSAJE_2, 255, 0, 0, True, False, True)
        
        Case eMessages.UserAttackedSwing
            Call General_Add_to_RichTextBox(frmMain.RecTxt, MENSAJE_1 & CharList(incomingData.ReadInteger()).Nombre & MENSAJE_ATAQUE_FALLO, 255, 0, 0, True, False, True)
        
        Case eMessages.UserHittedByUser
            Dim AttackerName As String
            
            AttackerName = General_Get_RawName(CharList(incomingData.ReadInteger()).Nombre)
            BodyPart = incomingData.ReadByte()
            Daño = incomingData.ReadInteger()
            
            Select Case BodyPart
                Case bCabeza
                    Call General_Add_to_RichTextBox(frmMain.RecTxt, MENSAJE_1 & AttackerName & MENSAJE_RECIVE_IMPACTO_CABEZA & Daño & MENSAJE_2, 255, 0, 0, True, False, True)
                
                Case bBrazoIzquierdo
                    Call General_Add_to_RichTextBox(frmMain.RecTxt, MENSAJE_1 & AttackerName & MENSAJE_RECIVE_IMPACTO_BRAZO_IZQ & Daño & MENSAJE_2, 255, 0, 0, True, False, True)
                
                Case bBrazoDerecho
                    Call General_Add_to_RichTextBox(frmMain.RecTxt, MENSAJE_1 & AttackerName & MENSAJE_RECIVE_IMPACTO_BRAZO_DER & Daño & MENSAJE_2, 255, 0, 0, True, False, True)
                
                Case bPiernaIzquierda
                    Call General_Add_to_RichTextBox(frmMain.RecTxt, MENSAJE_1 & AttackerName & MENSAJE_RECIVE_IMPACTO_PIERNA_IZQ & Daño & MENSAJE_2, 255, 0, 0, True, False, True)
                
                Case bPiernaDerecha
                    Call General_Add_to_RichTextBox(frmMain.RecTxt, MENSAJE_1 & AttackerName & MENSAJE_RECIVE_IMPACTO_PIERNA_DER & Daño & MENSAJE_2, 255, 0, 0, True, False, True)
                
                Case bTorso
                    Call General_Add_to_RichTextBox(frmMain.RecTxt, MENSAJE_1 & AttackerName & MENSAJE_RECIVE_IMPACTO_TORSO & Daño & MENSAJE_2, 255, 0, 0, True, False, True)
            End Select
        
        Case eMessages.UserHittedUser

            Dim VictimName As String
            
            VictimName = General_Get_RawName(CharList(incomingData.ReadInteger()).Nombre)
            BodyPart = incomingData.ReadByte()
            Daño = incomingData.ReadInteger()
            
            Select Case BodyPart
                Case bCabeza
                    Call General_Add_to_RichTextBox(frmMain.RecTxt, MENSAJE_PRODUCE_IMPACTO_1 & VictimName & MENSAJE_PRODUCE_IMPACTO_CABEZA & Daño & MENSAJE_2, 255, 0, 0, True, False, True)
                
                Case bBrazoIzquierdo
                    Call General_Add_to_RichTextBox(frmMain.RecTxt, MENSAJE_PRODUCE_IMPACTO_1 & VictimName & MENSAJE_PRODUCE_IMPACTO_BRAZO_IZQ & Daño & MENSAJE_2, 255, 0, 0, True, False, True)
                
                Case bBrazoDerecho
                    Call General_Add_to_RichTextBox(frmMain.RecTxt, MENSAJE_PRODUCE_IMPACTO_1 & VictimName & MENSAJE_PRODUCE_IMPACTO_BRAZO_DER & Daño & MENSAJE_2, 255, 0, 0, True, False, True)
                
                Case bPiernaIzquierda
                    Call General_Add_to_RichTextBox(frmMain.RecTxt, MENSAJE_PRODUCE_IMPACTO_1 & VictimName & MENSAJE_PRODUCE_IMPACTO_PIERNA_IZQ & Daño & MENSAJE_2, 255, 0, 0, True, False, True)
                
                Case bPiernaDerecha
                    Call General_Add_to_RichTextBox(frmMain.RecTxt, MENSAJE_PRODUCE_IMPACTO_1 & VictimName & MENSAJE_PRODUCE_IMPACTO_PIERNA_DER & Daño & MENSAJE_2, 255, 0, 0, True, False, True)
                
                Case bTorso
                    Call General_Add_to_RichTextBox(frmMain.RecTxt, MENSAJE_PRODUCE_IMPACTO_1 & VictimName & MENSAJE_PRODUCE_IMPACTO_TORSO & Daño & MENSAJE_2, 255, 0, 0, True, False, True)
            End Select
        
        Case eMessages.WorkRequestTarget
            UsingSkill = incomingData.ReadByte()
            
            frmMain.MousePointer = 2
            
            Select Case UsingSkill
                Case Magia
                    Call General_Add_to_RichTextBox(frmMain.RecTxt, MENSAJE_TRABAJO_MAGIA, 100, 100, 120, 0, 0)
                
                Case Pesca
                    Call General_Add_to_RichTextBox(frmMain.RecTxt, MENSAJE_TRABAJO_PESCA, 100, 100, 120, 0, 0)
                
                Case Robar
                    Call General_Add_to_RichTextBox(frmMain.RecTxt, MENSAJE_TRABAJO_ROBAR, 100, 100, 120, 0, 0)
                
                Case Talar
                    Call General_Add_to_RichTextBox(frmMain.RecTxt, MENSAJE_TRABAJO_TALAR, 100, 100, 120, 0, 0)
                
                Case Mineria
                    Call General_Add_to_RichTextBox(frmMain.RecTxt, MENSAJE_TRABAJO_MINERIA, 100, 100, 120, 0, 0)
                
                Case FundirMetal
                    Call General_Add_to_RichTextBox(frmMain.RecTxt, MENSAJE_TRABAJO_FUNDIRMETAL, 100, 100, 120, 0, 0)
                
                Case Proyectiles
                    Call General_Add_to_RichTextBox(frmMain.RecTxt, MENSAJE_TRABAJO_PROYECTILES, 100, 100, 120, 0, 0)
            End Select

        Case eMessages.HaveKilledUser
            Dim level As Long
            Call ShowConsoleMsg(MENSAJE_HAS_MATADO_A & CharList(.ReadInteger).Nombre & MENSAJE_22, 255, 0, 0, True, False)
            level = .ReadLong
            Call ShowConsoleMsg(MENSAJE_HAS_GANADO_EXPE_1 & level & MENSAJE_HAS_GANADO_EXPE_2, 255, 0, 0, True, False)

        Case eMessages.UserKill
            Call ShowConsoleMsg(CharList(.ReadInteger).Nombre & MENSAJE_TE_HA_MATADO, 255, 0, 0, True, False)

        Case eMessages.EarnExp
            Call ShowConsoleMsg(MENSAJE_HAS_GANADO_EXPE_1 & .ReadLong & MENSAJE_HAS_GANADO_EXPE_2, 255, 0, 0, True, False)
        Case eMessages.GoHome
            Dim Distance As Byte
            Dim Hogar As String
            Dim tiempo As Integer
            Distance = .ReadByte
            tiempo = .ReadInteger
            Hogar = .ReadASCIIString
            Call ShowConsoleMsg("Estás a " & Distance & " mapas de distancia de " & Hogar & ", este viaje durará " & tiempo & " segundos.", 255, 0, 0, True)
            Traveling = True
        Case eMessages.FinishHome
            Call ShowConsoleMsg(MENSAJE_HOGAR, 255, 255, 255)
            Traveling = False
        Case eMessages.CancelGoHome
            Call ShowConsoleMsg(MENSAJE_HOGAR_CANCEL, 255, 0, 0, True)
            Traveling = False
    End Select
End With
End Sub

''
'   Handles the NavigateToggle message.

Private Sub HandleNavigateToggle()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    UserNavegando = Not UserNavegando
End Sub
''
'   Handles the NPCHitUser message.

Private Sub HandleNPCHitUser()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.Length < 4 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Select Case incomingData.ReadByte()
        Case bCabeza
            Call General_Add_to_RichTextBox(frmMain.RecTxt, MENSAJE_GOLPE_CABEZA & CStr(incomingData.ReadInteger()) & "!!", 255, 0, 0, True, False, True)
        Case bBrazoIzquierdo
            Call General_Add_to_RichTextBox(frmMain.RecTxt, MENSAJE_GOLPE_BRAZO_IZQ & CStr(incomingData.ReadInteger()) & "!!", 255, 0, 0, True, False, True)
        Case bBrazoDerecho
            Call General_Add_to_RichTextBox(frmMain.RecTxt, MENSAJE_GOLPE_BRAZO_DER & CStr(incomingData.ReadInteger()) & "!!", 255, 0, 0, True, False, True)
        Case bPiernaIzquierda
            Call General_Add_to_RichTextBox(frmMain.RecTxt, MENSAJE_GOLPE_PIERNA_IZQ & CStr(incomingData.ReadInteger()) & "!!", 255, 0, 0, True, False, True)
        Case bPiernaDerecha
            Call General_Add_to_RichTextBox(frmMain.RecTxt, MENSAJE_GOLPE_PIERNA_DER & CStr(incomingData.ReadInteger()) & "!!", 255, 0, 0, True, False, True)
        Case bTorso
            Call General_Add_to_RichTextBox(frmMain.RecTxt, MENSAJE_GOLPE_TORSO & CStr(incomingData.ReadInteger() & "!!"), 255, 0, 0, True, False, True)
    End Select
End Sub

''
'   Handles the NPCKillUser message.

Private Sub HandleNPCKillUser()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    
    Call General_Add_to_RichTextBox(frmMain.RecTxt, MENSAJE_CRIATURA_MATADO, 255, 0, 0, True, False, True)
End Sub

''
'   Handles the NPCSwing message.

Private Sub HandleNPCSwing()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    
    Call General_Add_to_RichTextBox(frmMain.RecTxt, MENSAJE_CRIATURA_FALLA_GOLPE, 255, 0, 0, True, False, True)
End Sub

''
'   Handles the ObjectCreate message.

Private Sub HandleObjectCreate()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'16/05/10 - Standelf: Get the Name
'***************************************************
    If incomingData.Length < 5 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Dim X As Byte
    Dim Y As Byte
    
    X = incomingData.ReadByte()
    Y = incomingData.ReadByte()
    
    MapData(X, Y).OBJInfo.name = incomingData.ReadASCIIString()
    MapData(X, Y).ObjGrh.GrhIndex = incomingData.ReadInteger()
    
    Call InitGrh(MapData(X, Y).ObjGrh, MapData(X, Y).ObjGrh.GrhIndex)
End Sub

''
'   Handles the ObjectDelete message.

Private Sub HandleObjectDelete()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.Length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Dim X As Byte
    Dim Y As Byte
    
    X = incomingData.ReadByte()
    Y = incomingData.ReadByte()
    MapData(X, Y).ObjGrh.GrhIndex = 0
    MapData(X, Y).OBJInfo.name = ""
End Sub

''
'   Handles the OfferDetails message.

Private Sub HandleOfferDetails()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.Length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo ErrHandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call Buffer.ReadByte
    
    Call frmUserRequest.recievePeticion(Buffer.ReadASCIIString())
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(Buffer)
    
ErrHandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error <> 0 Then _
        Err.Raise Error
End Sub

''
'   Handles the ParalizeOK message.

Private Sub HandleParalizeOK()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    
    UserParalizado = Not UserParalizado
    
    If Not UserParalizado Then
        frmMain.imgPara.Visible = False
        frmMain.shpPara.Visible = False
    End If
    
End Sub

Private Sub HandlePartyDeatails()
'***************************************************
'Author: Standelf
'Last Modification: 01/06/10
'***************************************************
Dim members As Byte, tmp As String, i As Byte
    With incomingData
        'Remove PacketID
        Call .ReadByte
        members = .ReadByte()
        
        For i = 1 To members
            tmp = .ReadASCIIString()
                If tmp = "Vacio" Then
                    Call Kick_PartyMember(i)
                Else
                    Call Set_PartyMember(i, tmp, .ReadLong(), .ReadByte(), .ReadInteger())
                End If
        Next i
        
    End With
End Sub

Private Sub HandlePartyExit()
'***************************************************
'Author: Standelf
'Last Modification: 01/06/10
'***************************************************
Dim i As Byte

    With incomingData
        'Remove PacketID
        Call .ReadByte
        i = .ReadByte()
        
        Call Kick_PartyMember(i)
    End With
End Sub

Private Sub HandlePartyKicked()
'***************************************************
'Author: Standelf
'Last Modification: 01/06/10
'***************************************************
Dim i As Byte

    With incomingData
        'Remove PacketID
        Call .ReadByte
        
        For i = 1 To 5
            Call Kick_PartyMember(i)
        Next i
    End With
End Sub

Private Sub HandlePauseDuelStart()
'***************************************************
'Author: EzequielJuarez (Standelf)
'Last Modification: 20/01/10
'
'***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    
    'Take Val
    Dim tmp As Byte
    tmp = incomingData.ReadByte
    
    If tmp = 1 Then
        EnDuelo = True
    Else
        EnDuelo = False
    End If
    
End Sub

''
'   Handles the PauseToggle message.

Private Sub HandlePauseToggle()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    
    pausa = Not pausa
End Sub

''
'   Handles the PeaceProposalsList message.

Private Sub HandlePeaceProposalsList()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.Length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo ErrHandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call Buffer.ReadByte
    
    Dim guildList() As String
    Dim i As Long
    
    guildList = Split(Buffer.ReadASCIIString(), SEPARATOR)
    
    Call frmPeaceProp.lista.Clear
    For i = 0 To UBound(guildList())
        Call frmPeaceProp.lista.AddItem(guildList(i))
    Next i
    
    frmPeaceProp.ProposalType = TIPO_PROPUESTA.PAZ
    Call frmPeaceProp.Show(vbModeless, frmMain)
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(Buffer)
    
ErrHandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error <> 0 Then _
        Err.Raise Error
End Sub

''
'   Handles the PlayMIDI message.

Private Sub HandlePlayMIDI()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.Length < 4 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    Dim currentMidi As Byte
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    currentMidi = incomingData.ReadByte()
    
    If currentMidi Then
        Call Audio.PlayMIDI(CStr(currentMidi) & ".mid", incomingData.ReadInteger())
    Else
        'Remove the bytes to prevent errors
        Call incomingData.ReadInteger
    End If
End Sub

''
'   Handles the PlayWave message.

Private Sub HandlePlayWave()
'***************************************************
'Autor: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 08/14/07
'Last Modified by: Rapsodius
'Added support for 3D Sounds.
'***************************************************
    If incomingData.Length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
        
    Dim wave As Byte
    Dim srcX As Byte
    Dim srcY As Byte
    
    wave = incomingData.ReadByte()
    srcX = incomingData.ReadByte()
    srcY = incomingData.ReadByte()
        
    Call Audio.PlayWave(CStr(wave) & ".wav", srcX, srcY)
    
End Sub

''
'   Handles the Pong message.

Private Sub HandlePong()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    Call incomingData.ReadByte
    
    Call General_Add_to_RichTextBox(frmMain.RecTxt, "El ping es " & (GetTickCount - pingTime) & " ms.", 255, 0, 0, True, False, True)
    
    pingTime = 0
End Sub

''
'   Handles the PosUpdate message.

Private Sub HandlePosUpdate()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.Length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    'Remove char from old position
    If MapData(UserPos.X, UserPos.Y).CharIndex = UserCharIndex Then
        MapData(UserPos.X, UserPos.Y).CharIndex = 0
    End If
    
    'Set new pos
    UserPos.X = incomingData.ReadByte()
    UserPos.Y = incomingData.ReadByte()
    

    
    'Set char
    MapData(UserPos.X, UserPos.Y).CharIndex = UserCharIndex
    CharList(UserCharIndex).Pos.Y = UserPos.Y
    CharList(UserCharIndex).Pos.X = UserPos.X
    
    'Are we under a roof?
    bTecho = IIf(MapData(UserPos.X, UserPos.Y).Trigger = 1 Or _
            MapData(UserPos.X, UserPos.Y).Trigger = 2 Or _
            MapData(UserPos.X, UserPos.Y).Trigger = 6 Or _
            MapData(UserPos.X, UserPos.Y).Trigger = 7 Or _
            MapData(UserPos.X, UserPos.Y).Trigger = 4, True, False)
                
    'Update pos label
                 frmMain.lblMapPosX.Caption = UserPos.X
                frmMain.lblMapPosY.Caption = UserPos.Y

    If Settings.MiniMap Then If Settings.MiniMap Then Call General_Pixel_Map_Set_Area
End Sub

Private Sub HandleQuestDetails()
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'Recibe y maneja el paquete QuestDetails del servidor.
'Last modified: 31/01/2010 by Amraphen
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    If incomingData.Length < 15 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo ErrHandler
    Dim Buffer As New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    
    Dim tmpStr As String
    Dim tmpByte As Byte
    Dim QuestEmpezada As Boolean
    Dim i As Integer
    
    With Buffer
        'Leemos el id del paquete
        Call .ReadByte
        
        'Nos fijamos si se trata de una quest empezada, para poder leer los NPCs que se han matado.
        QuestEmpezada = IIf(.ReadByte, True, False)
        
        tmpStr = "Misión: " & .ReadASCIIString & vbCrLf
        tmpStr = tmpStr & "Detalles: " & .ReadASCIIString & vbCrLf
        tmpStr = tmpStr & "Nivel requerido: " & .ReadByte & vbCrLf
        
        tmpStr = tmpStr & vbCrLf & "OBJETIVOS" & vbCrLf
        
        tmpByte = .ReadByte
        If tmpByte Then 'Hay NPCs
            For i = 1 To tmpByte
                tmpStr = tmpStr & "*) Matar " & .ReadInteger & " " & .ReadASCIIString & "."
                If QuestEmpezada Then
                    tmpStr = tmpStr & " (Has matado " & .ReadInteger & ")" & vbCrLf
                Else
                    tmpStr = tmpStr & vbCrLf
                End If
            Next i
        End If
        
        tmpByte = .ReadByte
        If tmpByte Then 'Hay OBJs
            For i = 1 To tmpByte
                tmpStr = tmpStr & "*) Conseguir " & .ReadInteger & " " & .ReadASCIIString & "." & vbCrLf
            Next i
        End If

        tmpStr = tmpStr & vbCrLf & "RECOMPENSAS" & vbCrLf
        tmpStr = tmpStr & "*) Oro: " & .ReadLong & " monedas de oro." & vbCrLf
        tmpStr = tmpStr & "*) Experiencia: " & .ReadLong & " puntos de experiencia." & vbCrLf
        
        tmpByte = .ReadByte
        If tmpByte Then
            For i = 1 To tmpByte
                tmpStr = tmpStr & "*) " & .ReadInteger & " " & .ReadASCIIString & vbCrLf
            Next i
        End If
    End With
    
    'Determinamos que formulario se muestra, según si recibimos la información y la quest está empezada o no.
    If QuestEmpezada Then
        frmQuests.txtInfo.Text = tmpStr
    Else
        frmQuestInfo.txtInfo.Text = tmpStr
        frmQuestInfo.Show vbModeless, frmMain
    End If
    
    Call incomingData.CopyBuffer(Buffer)
    
ErrHandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error <> 0 Then _
        Err.Raise Error
End Sub

Public Sub HandleQuestListSend()
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'Recibe y maneja el paquete QuestListSend del servidor.
'Last modified: 31/01/2010 by Amraphen
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    If incomingData.Length < 1 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo ErrHandler
    Dim Buffer As New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    
    Dim i As Integer
    Dim tmpByte As Byte
    Dim tmpStr As String
    
    'Leemos el id del paquete
    Call Buffer.ReadByte
     
    'Leemos la cantidad de quests que tiene el usuario
    tmpByte = Buffer.ReadByte
    
    'Limpiamos el ListBox y el TextBox del formulario
    frmQuests.lstQuests.Clear
    frmQuests.txtInfo.Text = vbNullString
        
    'Si el usuario tiene quests entonces hacemos el handle
    If tmpByte Then
        'Leemos el string
        tmpStr = Buffer.ReadASCIIString
        
        'Agregamos los items
        For i = 1 To tmpByte
            frmQuests.lstQuests.AddItem General_Get_ReadField(i, tmpStr, 45)
        Next i
    End If
    
    'Mostramos el formulario
    frmQuests.Show vbModeless, frmMain
    
    'Pedimos la información de la primer quest (si la hay)
    If tmpByte Then Call Mod_Protocol.WriteQuestDetailsRequest(1)
    
    'Copiamos de vuelta el buffer
    Call incomingData.CopyBuffer(Buffer)

ErrHandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error <> 0 Then _
        Err.Raise Error
End Sub

''
'   Handles the RainToggle message.

Private Sub HandleRainToggle()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    
    If Not InMapBounds(UserPos.X, UserPos.Y) Then Exit Sub
    
    bTecho = (MapData(UserPos.X, UserPos.Y).Trigger = 1 Or _
            MapData(UserPos.X, UserPos.Y).Trigger = 2 Or _
            MapData(UserPos.X, UserPos.Y).Trigger = 6 Or _
            MapData(UserPos.X, UserPos.Y).Trigger = 7 Or _
            MapData(UserPos.X, UserPos.Y).Trigger = 4)
            
    If bRain Then
            'Stop playing the rain sound
            Call Audio.StopWave(RainBufferIndex)
            RainBufferIndex = 0
            If bTecho Then
                Call Audio.PlayWave("lluviainend.wav", 0, 0, LoopStyle.Disabled)
            Else
                Call Audio.PlayWave("lluviaoutend.wav", 0, 0, LoopStyle.Disabled)
            End If
            frmMain.IsPlaying = PlayLoop.plNone
    End If
    
    bRain = Not bRain
    
    If bRain = False Then
        If Effect(WeatherEffectIndex).Used Then Effect_Kill WeatherEffectIndex
    End If
End Sub

''
'   Handles the RemoveCharDialog message.

Private Sub HandleRemoveCharDialog()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Check if the packet is complete
    If incomingData.Length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Call Dialogos.RemoveDialog(incomingData.ReadInteger())
End Sub

''
'   Handles the RemoveDialogs message.

Private Sub HandleRemoveDialogs()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    
    Call Dialogos.RemoveAllDialogs
End Sub

''
'   Handles the RestOK message.

Private Sub HandleRestOK()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    
    UserDescansar = Not UserDescansar
End Sub

''
'   Handles the ResuscitationSafeOff message.

Private Sub HandleResuscitationSafeOff()
'***************************************************
'Author: Rapsodius
'Creation date: 10/10/07
'***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    
    Call frmMain.ControlSM(eSMType.sResucitation, False)
End Sub

''
'   Handles the ResuscitationSafeOn message.

Private Sub HandleResuscitationSafeOn()
'***************************************************
'Author: Rapsodius
'Creation date: 10/10/07
'***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    
    Call frmMain.ControlSM(eSMType.sResucitation, True)
End Sub

''
'   Handles the SafeModeOff message.

Private Sub HandleSafeModeOff()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    
    Call frmMain.ControlSM(eSMType.sSafemode, False)
End Sub

''
'   Handles the SafeModeOn message.

Private Sub HandleSafeModeOn()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    
    Call frmMain.ControlSM(eSMType.sSafemode, True)
End Sub

''
'   Handles the SendNight message.

Private Sub HandleSendNight()
'***************************************************
'Author: Fredy Horacio Treboux (liquid)
'Last Modification: 01/08/07
'
'***************************************************
    If incomingData.Length < 2 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Dim tBool As Boolean 'CHECK, este handle no hace nada con lo que recibe.. porque, ehmm.. no hay noche?.. o si?
    tBool = incomingData.ReadBoolean()
End Sub

''
'   Handles the SendSkills message.

Private Sub HandleSendSkills()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 11/19/09
'11/19/09: Pato - Now the server send the percentage of progress of the skills.
'***************************************************
    If incomingData.Length < 2 + NUMSKILLS * 2 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    UserClase = incomingData.ReadByte
    Dim i As Long
    
    For i = 1 To NUMSKILLS
        UserSkills(i) = incomingData.ReadByte()
        PorcentajeSkills(i) = incomingData.ReadByte()
    Next i
    LlegaronSkills = True
End Sub

''
'   Handles the SetInvisible message.

Private Sub HandleSetInvisible()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.Length < 4 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Dim CharIndex As Integer
    
    CharIndex = incomingData.ReadInteger()
    CharList(CharIndex).invisible = incomingData.ReadBoolean()
    

End Sub

''
'   Handles the ShowBlacksmithForm message.

Private Sub HandleShowBlacksmithForm()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    
    If frmMain.MacroTrabajo.Enabled And (MacroBltIndex > 0) Then
        Call WriteCraftBlacksmith(MacroBltIndex)
    Else
        frmHerrero.Show , frmMain
    End If
End Sub

''
'   Handles the ShowCarpenterForm message.

Private Sub HandleShowCarpenterForm()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    
    If frmMain.MacroTrabajo.Enabled And (MacroBltIndex > 0) Then
        Call WriteCraftCarpenter(MacroBltIndex)
    Else
        frmCarp.Show , frmMain
    End If
End Sub

''
'   Handles the ShowForumForm message.

Private Sub HandleShowForumForm()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    
    frmForo.Privilegios = incomingData.ReadByte
    frmForo.CanPostSticky = incomingData.ReadByte
    
    If Not MirandoForo Then
        frmForo.Show , frmMain
    End If
End Sub

''
'   Handles the ShowGMPanelForm message.

Private Sub HandleShowGMPanelForm()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    
    frmPanelGm.Show vbModeless, frmMain
End Sub

''
'   Handles the ShowGuildAlign message.

Private Sub HandleShowGuildAlign()
'***************************************************
'Author: ZaMa
'Last Modification: 14/12/2009
'
'***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    
    frmEligeAlineacion.Show vbModeless, frmMain
End Sub

''
'   Handles the ShowGuildFundationForm message.

Private Sub HandleShowGuildFundationForm()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    
    CreandoClan = True
    frmGuildFoundation.Show , frmMain
End Sub

''
'   Handles the ShowMessageBox message.

Private Sub HandleShowMessageBox()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.Length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo ErrHandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call Buffer.ReadByte
    
    frmMensaje.msg.Caption = Buffer.ReadASCIIString()
    frmMensaje.Show
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(Buffer)
    
ErrHandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error <> 0 Then _
        Err.Raise Error
End Sub

''
'   Handles the ShowMOTDEditionForm message.

Private Sub HandleShowMOTDEditionForm()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'*************************************Su**************
    If incomingData.Length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo ErrHandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call Buffer.ReadByte
    
    frmCambiaMotd.txtMotd.Text = Buffer.ReadASCIIString()
    frmCambiaMotd.Show , frmMain
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(Buffer)
    
ErrHandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error <> 0 Then _
        Err.Raise Error
End Sub

''
'   Handles the ShowSOSForm message.

Private Sub HandleShowPartyForm()
'***************************************************
'Author: Budi
'Last Modification: 11/26/09
'
'***************************************************
    If incomingData.Length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo ErrHandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call Buffer.ReadByte
    
    Dim members() As String
    Dim i As Long
    
    EsPartyLeader = CBool(Buffer.ReadByte())
       
    members = Split(Buffer.ReadASCIIString(), SEPARATOR)
    For i = 0 To UBound(members())
        Call frmParty.lstMembers.AddItem(members(i))
    Next i
    
    frmParty.lblTotalExp.Caption = Buffer.ReadLong
    frmParty.Show , frmMain
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(Buffer)
    
ErrHandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error <> 0 Then _
        Err.Raise Error
End Sub

''
'   Handles the ShowSignal message.

Private Sub HandleShowSignal()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.Length < 5 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo ErrHandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call Buffer.ReadByte
    
    Dim tmp As String
    tmp = Buffer.ReadASCIIString()
    
    Call InitCartel(tmp, Buffer.ReadInteger())
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(Buffer)
    
ErrHandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error <> 0 Then _
        Err.Raise Error
End Sub

''
'   Handles the ShowSOSForm message.

Private Sub HandleShowSOSForm()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.Length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo ErrHandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call Buffer.ReadByte
    
    Dim sosList() As String
    Dim i As Long
    
    sosList = Split(Buffer.ReadASCIIString(), SEPARATOR)
    
    For i = 0 To UBound(sosList())
        Call frmMSG.List1.AddItem(sosList(i))
    Next i
    
    frmMSG.Show , frmMain
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(Buffer)
    
ErrHandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error <> 0 Then _
        Err.Raise Error
End Sub

Public Sub HandleShowStatsBar()
'***************************************************
'Author: EzequielJuarez (Standelf)
'Last Modification: 08/02/2011
'***************************************************

    'Remove packet ID
    Call incomingData.ReadByte
    Dim e_CharIndex As Integer
    
    e_CharIndex = incomingData.ReadInteger
    
    CharList(e_CharIndex).minHP = incomingData.ReadLong
    CharList(e_CharIndex).maxHP = incomingData.ReadLong
    CharList(e_CharIndex).tmpName = incomingData.ReadASCIIString
    CharList(e_CharIndex).ShowBar = 8 * FramesPerSecond 'Calculamos 10 Segundos
    
End Sub

''
'   Handles the ShowUserRequest message.

Private Sub HandleShowUserRequest()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.Length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo ErrHandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call Buffer.ReadByte
    
    Call frmUserRequest.recievePeticion(Buffer.ReadASCIIString())
    Call frmUserRequest.Show(vbModeless, frmMain)
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(Buffer)
    
ErrHandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error <> 0 Then _
        Err.Raise Error
End Sub

''
'   Handles the SpawnList message.

Private Sub HandleSpawnList()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.Length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo ErrHandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call Buffer.ReadByte
    
    Dim creatureList() As String
    Dim i As Long
    
    creatureList = Split(Buffer.ReadASCIIString(), SEPARATOR)
    
    For i = 0 To UBound(creatureList())
        Call frmSpawnList.lstCriaturas.AddItem(creatureList(i))
    Next i
    frmSpawnList.Show , frmMain
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(Buffer)
    
ErrHandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error <> 0 Then _
        Err.Raise Error
End Sub

'   Handles the StopWorking message.
Private Sub HandleStopWorking()
'***************************************************
'Author: Budi
'Last Modification: 12/01/09
'
'***************************************************

    Call incomingData.ReadByte
    
    With FontTypes(FontTypeNames.FONTTYPE_INFO)
        Call ShowConsoleMsg("¡Has terminado de trabajar!", .Red, .Green, .Blue, .bold, .italic)
    End With
    
    If frmMain.MacroTrabajo.Enabled Then Call frmMain.DesactivarMacroTrabajo
End Sub

''
'   Handles the TradeOK message.

Private Sub HandleTradeOK()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    
    If frmComerciar.Visible Then
        Dim i As Long
        
        'Update user inventory
        For i = 1 To MAX_INVENTORY_SLOTS
            '   Agrego o quito un item en su totalidad
            If Inventario.OBJIndex(i) <> InvComUsu.OBJIndex(i) Then
                With Inventario
                    Call InvComUsu.SetItem(i, .OBJIndex(i), _
                    .Amount(i), .Equipped(i), .GrhIndex(i), _
                    .OBJType(i), .MaxHit(i), .MinHit(i), .MaxDef(i), .MinDef(i), _
                    .Valor(i), .ItemName(i))
                End With
            '   Vendio o compro cierta cantidad de un item que ya tenia
            ElseIf Inventario.Amount(i) <> InvComUsu.Amount(i) Then
                Call InvComUsu.ChangeSlotItemAmount(i, Inventario.Amount(i))
            End If
        Next i
        
        '   Fill Npc inventory
        For i = 1 To 20
            '   Compraron la totalidad de un item, o vendieron un item que el npc no tenia
            If NPCInventory(i).OBJIndex <> InvComNpc.OBJIndex(i) Then
                With NPCInventory(i)
                    Call InvComNpc.SetItem(i, .OBJIndex, _
                    .Amount, 0, .GrhIndex, _
                    .OBJType, .MaxHit, .MinHit, .MaxDef, .MinDef, _
                    .Valor, .name)
                End With
            '   Compraron o vendieron cierta cantidad (no su totalidad)
            ElseIf NPCInventory(i).Amount <> InvComNpc.Amount(i) Then
                Call InvComNpc.ChangeSlotItemAmount(i, NPCInventory(i).Amount)
            End If
        Next i
    
    End If
End Sub

''
'   Handles the TrainerCreatureList message.

Private Sub HandleTrainerCreatureList()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.Length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo ErrHandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call Buffer.ReadByte
    
    Dim creatures() As String
    Dim i As Long
    
    creatures = Split(Buffer.ReadASCIIString(), SEPARATOR)
    
    For i = 0 To UBound(creatures())
        Call frmEntrenador.lstCriaturas.AddItem(creatures(i))
    Next i
    frmEntrenador.Show , frmMain
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(Buffer)
    
ErrHandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error <> 0 Then _
        Err.Raise Error
End Sub

''
'   Handles the UpdateBankGold message.

Private Sub HandleUpdateBankGold()
'***************************************************
'Autor: ZaMa
'Last Modification: 14/12/2009
'
'***************************************************
    'Check packet is complete
    If incomingData.Length < 5 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    frmBancoObj.lblUserGld.Caption = Format(incomingData.ReadLong, "###,###,###,###")
    
End Sub

Private Sub HandleUpdateClima()
    With incomingData
        'Remove PacketID
        Call .ReadByte
        
        Actualizar_Estado .ReadByte()
    End With
End Sub

'   Handles the UpdateStrenghtAndDexterity message.

Private Sub HandleUpdateDexterity()
'***************************************************
'Author: Budi
'Last Modification: 11/26/09
'***************************************************
    'Check packet is complete
    If incomingData.Length < 2 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    'Get data and update form
    UserAgilidad = incomingData.ReadByte
    frmMain.lblDext.Caption = "A: " & UserAgilidad
    frmMain.lblDext.ForeColor = General_Get_DexterityColor()
End Sub

''
'   Handles the UpdateExp message.

Private Sub HandleUpdateExp()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Check packet is complete
    If incomingData.Length < 5 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    'Get data and update form
    UserExp = incomingData.ReadLong()
    
    If MostrarExp = True Then
        frmMain.lblExp.Caption = UserExp & "/" & UserPasarNivel
    Else
        If UserPasarNivel > 0 Then
            frmMain.lblExp.Caption = "[" & Round(CDbl(UserExp) * CDbl(100) / CDbl(UserPasarNivel), 2) & "%]"
        Else
            frmMain.lblExp.Caption = "[N/A]"
        End If
    End If
    
    If UserPasarNivel > 0 Then
        frmMain.expBar.Width = (((UserExp / 100) / (UserPasarNivel / 100)) * 160)
    ElseIf UserLvl = 60 And UserExp = 0 Then
        frmMain.expBar.Width = 160
    Else
        frmMain.expBar.Width = 0
    End If

End Sub

''
'   Handles the UpdateGold message.

Private Sub HandleUpdateGold()
'***************************************************
'Autor: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 08/14/07
'Last Modified By: Lucas Tavolaro Ortiz (Tavo)
'- 08/14/07: Added GldLbl color variation depending on User Gold and Level
'***************************************************
    'Check packet is complete
    If incomingData.Length < 5 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    'Get data and update form
    UserGLD = incomingData.ReadLong()
    
    If UserGLD >= CLng(UserLvl) * 10000 Then
        'Changes color
        frmMain.GldLbl.ForeColor = &HFF& 'Red
    Else
        'Changes color
        frmMain.GldLbl.ForeColor = &HFFFF& 'Yellow
    End If
    
    frmMain.GldLbl.Caption = Format(UserGLD, "###,###,###,###")
End Sub

''
'   Handles the UpdateHP message.

Private Sub HandleUpdateHP()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Check packet is complete
    If incomingData.Length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    'Get data and update form
    UserMinHP = incomingData.ReadInteger()
    
    frmMain.lblVida = UserMinHP & "/" & UserMaxHP
    frmMain.hpBar.Width = (((UserMinHP / 100) / (UserMaxHP / 100)) * 136)
    
    'Is the user alive??
    If UserMinHP = 0 Then
        UserEstado = 1
        If frmMain.TrainingMacro Then Call frmMain.DesactivarMacroHechizos
        If frmMain.MacroTrabajo Then Call frmMain.DesactivarMacroTrabajo
    Else
        UserEstado = 0
    End If
End Sub

''
'   Handles the UpdateHungerAndThirst message.

Private Sub HandleUpdateHungerAndThirst()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.Length < 5 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    UserMaxAGU = incomingData.ReadByte()
    UserMinAGU = incomingData.ReadByte()
    UserMaxHAM = incomingData.ReadByte()
    UserMinHAM = incomingData.ReadByte()
    
    
    'frmMain.lblHambre = UserMinHAM & "/" & UserMaxHAM
    frmMain.lblHambre.Caption = Round(CDbl(UserMinHAM) * CDbl(100) / CDbl(UserMaxHAM), 0) & "%"
    
    'frmMain.lblSed = UserMinAGU & "/" & UserMaxAGU
    frmMain.lblSed.Caption = Round(CDbl(UserMinAGU) * CDbl(100) / CDbl(UserMaxAGU), 0) & "%"
End Sub

''
'   Handles the UpdateMana message.

Private Sub HandleUpdateMana()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Check packet is complete
    If incomingData.Length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    'Get data and update form
    UserMinMAN = incomingData.ReadInteger()
    
    frmMain.lblMana = UserMinMAN & "/" & UserMaxMAN
    frmMain.manBar.Width = (((UserMinMAN / 100) / (UserMaxMAN / 100)) * 136)

End Sub

Private Sub HandleUpdateParaTimes()
'***************************************************
'Author: Standelf
'Last Modification: 22/05/10
'***************************************************
    'Remove PacketID
    Call incomingData.ReadByte

    UserParalizadoTime = incomingData.ReadByte

    If UserParalizado And UserParalizadoTime <> 0 Then
        frmMain.imgPara.Visible = True
        frmMain.shpPara.Visible = True
        frmMain.shpPara.Width = (((UserParalizadoTime + 1 / 100) / (20 + 1 / 100)) * 32)
    Else
        frmMain.imgPara.Visible = False
        frmMain.shpPara.Visible = False
    End If
End Sub

''
'   Handles the UpdateSta message.

Private Sub HandleUpdateSta()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Check packet is complete
    If incomingData.Length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    'Get data and update form
    UserMinSTA = incomingData.ReadInteger()
    
    frmMain.lblEnergia = UserMinSTA & "/" & UserMaxSTA
    frmMain.staBar.Width = (((UserMinSTA / 100) / (UserMaxSTA / 100)) * 139)
    
End Sub

'   Handles the UpdateStrenghtAndDexterity message.

Private Sub HandleUpdateStrenght()
'***************************************************
'Author: Budi
'Last Modification: 11/26/09
'***************************************************
    'Check packet is complete
    If incomingData.Length < 2 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    'Get data and update form
    UserFuerza = incomingData.ReadByte
    frmMain.lblStrg.Caption = "F: " & UserFuerza
    frmMain.lblStrg.ForeColor = General_Get_StrenghtColor()
End Sub

''
'   Handles the UpdateStrenghtAndDexterity message.

Private Sub HandleUpdateStrenghtAndDexterity()
'***************************************************
'Author: Budi
'Last Modification: 11/26/09
'***************************************************
    'Check packet is complete
    If incomingData.Length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    'Get data and update form
    UserFuerza = incomingData.ReadByte
    UserAgilidad = incomingData.ReadByte
    frmMain.lblStrg.Caption = "F: " & UserFuerza
    frmMain.lblDext.Caption = "A: " & UserAgilidad
    frmMain.lblStrg.ForeColor = General_Get_StrenghtColor()
    frmMain.lblDext.ForeColor = General_Get_DexterityColor()
End Sub

''
'   Handles the UpdateTag message.

Private Sub HandleUpdateTagAndStatus()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.Length < 6 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo ErrHandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call Buffer.ReadByte
    
    Dim CharIndex As Integer
    Dim NickColor As Byte
    Dim UserTag As String
    
    CharIndex = Buffer.ReadInteger()
    NickColor = Buffer.ReadByte()
    UserTag = Buffer.ReadASCIIString()
    
    'Update char status adn tag!
    With CharList(CharIndex)
        If (NickColor And eNickColor.ieCriminal) <> 0 Then
            .Criminal = 1
        Else
            .Criminal = 0
        End If
        
        .Atacable = (NickColor And eNickColor.ieAtacable) <> 0
        
        .Nombre = UserTag
    End With
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(Buffer)
    
ErrHandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error <> 0 Then _
        Err.Raise Error
End Sub

''
'   Handles the UpdateUserStats message.

Private Sub HandleUpdateUserStats()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
'Khalem
    If incomingData.Length < 33 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

    'Remove packet ID
    Call incomingData.ReadByte
    
    UserMaxHP = incomingData.ReadInteger()
    UserMinHP = incomingData.ReadInteger()
    UserMaxMAN = incomingData.ReadInteger()
    UserMinMAN = incomingData.ReadInteger()
    UserMaxSTA = incomingData.ReadInteger()
    UserMinSTA = incomingData.ReadInteger()
    UserGLD = incomingData.ReadLong()
    UserLvl = incomingData.ReadByte()
    UserPasarNivel = incomingData.ReadLong()
    UserExp = incomingData.ReadLong()
    
    If MostrarExp = True Then
        frmMain.lblExp.Caption = UserExp & "/" & UserPasarNivel
    Else
        If UserPasarNivel > 0 Then
            frmMain.lblExp.Caption = "[" & Round(CDbl(UserExp) * CDbl(100) / CDbl(UserPasarNivel), 2) & "%]"
        Else
            frmMain.lblExp.Caption = "[N/A]"
        End If
    End If
    
    Dim bWidth As Byte
    
    If UserPasarNivel > 0 Then
        frmMain.expBar.Width = (((UserExp / 100) / (UserPasarNivel / 100)) * 160)
    ElseIf UserLvl = 60 And UserExp = 0 Then
        frmMain.expBar.Width = 160
    Else
        frmMain.expBar.Width = 0
    End If
    
    frmMain.GldLbl.Caption = Format(UserGLD, "###,###,###,###")

    frmMain.Label5.Caption = UserLvl
    
    'Stats
    frmMain.lblMana = UserMinMAN & "/" & UserMaxMAN
    frmMain.manBar.Width = (((UserMinMAN / 100) / (UserMaxMAN / 100)) * 136)
        
    frmMain.lblVida = UserMinHP & "/" & UserMaxHP
    frmMain.hpBar.Width = (((UserMinHP / 100) / (UserMaxHP / 100)) * 136)
    
    frmMain.lblEnergia = UserMinSTA & "/" & UserMaxSTA
    frmMain.staBar.Width = (((UserMinSTA / 100) / (UserMaxSTA / 100)) * 136)

    If UserMinHP = 0 Then
        UserEstado = 1
        If frmMain.TrainingMacro Then Call frmMain.DesactivarMacroHechizos
        If frmMain.MacroTrabajo Then Call frmMain.DesactivarMacroTrabajo
    Else
        UserEstado = 0
    End If
    
    If UserGLD >= CLng(UserLvl) * 10000 Then
        'Changes color
        frmMain.GldLbl.ForeColor = &HFF& 'Red
    Else
        'Changes color
        frmMain.GldLbl.ForeColor = &HFFFF& 'Yellow
    End If
    
End Sub

''
'   Handles the UserAttackedSwing message.

Private Sub HandleUserAttackedSwing()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.Length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Call General_Add_to_RichTextBox(frmMain.RecTxt, MENSAJE_1 & CharList(incomingData.ReadInteger()).Nombre & MENSAJE_ATAQUE_FALLO, 255, 0, 0, True, False, True)
End Sub

''
'   Handles the UserCharIndexInServer message.

Private Sub HandleUserCharIndexInServer()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.Length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    UserCharIndex = incomingData.ReadInteger()
    UserPos.X = CharList(UserCharIndex).Pos.X
    UserPos.Y = CharList(UserCharIndex).Pos.Y
    
    'Are we under a roof?
    bTecho = IIf(MapData(UserPos.X, UserPos.Y).Trigger = 1 Or _
            MapData(UserPos.X, UserPos.Y).Trigger = 2 Or _
            MapData(UserPos.X, UserPos.Y).Trigger = 6 Or _
            MapData(UserPos.X, UserPos.Y).Trigger = 7 Or _
            MapData(UserPos.X, UserPos.Y).Trigger = 4, True, False)

    'frmMain.Coord.Caption = UserMap & " X: " & UserPos.X & " Y: " & UserPos.Y
    If Settings.MiniMap Then Call General_Pixel_Map_Set_Area
End Sub

''
'   Handles the UserCommerceEnd message.

Private Sub HandleUserCommerceEnd()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    
    Set InvComUsu = Nothing
    Set InvOroComUsu(0) = Nothing
    Set InvOroComUsu(1) = Nothing
    Set InvOroComUsu(2) = Nothing
    Set InvOfferComUsu(0) = Nothing
    Set InvOfferComUsu(1) = Nothing
    
    'Destroy the form and reset the state
    Unload frmComerciarUsu
    Comerciando = False
End Sub

''
'   Handles the UserCommerceInit message.

Private Sub HandleUserCommerceInit()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    Dim i As Long
    
    'Remove packet ID
    Call incomingData.ReadByte
    TradingUserName = incomingData.ReadASCIIString
    
    '   Initialize commerce inventories
    Call InvComUsu.Initialize(DirectD3D8, frmComerciarUsu.picInvComercio, Inventario.MaxObjs)
    Call InvOfferComUsu(0).Initialize(DirectD3D8, frmComerciarUsu.picInvOfertaProp, INV_OFFER_SLOTS)
    Call InvOfferComUsu(1).Initialize(DirectD3D8, frmComerciarUsu.picInvOfertaOtro, INV_OFFER_SLOTS)
    Call InvOroComUsu(0).Initialize(DirectD3D8, frmComerciarUsu.picInvOroProp, INV_GOLD_SLOTS, , TilePixelWidth * 2, TilePixelHeight, TilePixelWidth / 2)
    Call InvOroComUsu(1).Initialize(DirectD3D8, frmComerciarUsu.picInvOroOfertaProp, INV_GOLD_SLOTS, , TilePixelWidth * 2, TilePixelHeight, TilePixelWidth / 2)
    Call InvOroComUsu(2).Initialize(DirectD3D8, frmComerciarUsu.picInvOroOfertaOtro, INV_GOLD_SLOTS, , TilePixelWidth * 2, TilePixelHeight, TilePixelWidth / 2)

    'Fill user inventory
    For i = 1 To MAX_INVENTORY_SLOTS
        If Inventario.OBJIndex(i) <> 0 Then
            With Inventario
                Call InvComUsu.SetItem(i, .OBJIndex(i), _
                .Amount(i), .Equipped(i), .GrhIndex(i), _
                .OBJType(i), .MaxHit(i), .MinHit(i), .MaxDef(i), .MinDef(i), _
                .Valor(i), .ItemName(i))
            End With
        End If
    Next i

    '   Inventarios de oro
    Call InvOroComUsu(0).SetItem(1, ORO_INDEX, UserGLD, 0, ORO_GRH, 0, 0, 0, 0, 0, 0, "Oro")
    Call InvOroComUsu(1).SetItem(1, ORO_INDEX, 0, 0, ORO_GRH, 0, 0, 0, 0, 0, 0, "Oro")
    Call InvOroComUsu(2).SetItem(1, ORO_INDEX, 0, 0, ORO_GRH, 0, 0, 0, 0, 0, 0, "Oro")


    'Set state and show form
    Comerciando = True
    Call frmComerciarUsu.Show(vbModeless, frmMain)
End Sub

''
'   Handles the UserHitNPC message.

Private Sub HandleUserHitNPC()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.Length < 5 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Call General_Add_to_RichTextBox(frmMain.RecTxt, MENSAJE_GOLPE_CRIATURA_1 & CStr(incomingData.ReadLong()) & MENSAJE_2, 255, 0, 0, True, False, True)
End Sub

''
'   Handles the UserHittingByUser message.

Private Sub HandleUserHittedByUser()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.Length < 6 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Dim attacker As String
    
    attacker = CharList(incomingData.ReadInteger()).Nombre
    
    Select Case incomingData.ReadByte
        Case bCabeza
            Call General_Add_to_RichTextBox(frmMain.RecTxt, MENSAJE_1 & attacker & MENSAJE_RECIVE_IMPACTO_CABEZA & CStr(incomingData.ReadInteger() & MENSAJE_2), 255, 0, 0, True, False, True)
        Case bBrazoIzquierdo
            Call General_Add_to_RichTextBox(frmMain.RecTxt, MENSAJE_1 & attacker & MENSAJE_RECIVE_IMPACTO_BRAZO_IZQ & CStr(incomingData.ReadInteger() & MENSAJE_2), 255, 0, 0, True, False, True)
        Case bBrazoDerecho
            Call General_Add_to_RichTextBox(frmMain.RecTxt, MENSAJE_1 & attacker & MENSAJE_RECIVE_IMPACTO_BRAZO_DER & CStr(incomingData.ReadInteger() & MENSAJE_2), 255, 0, 0, True, False, True)
        Case bPiernaIzquierda
            Call General_Add_to_RichTextBox(frmMain.RecTxt, MENSAJE_1 & attacker & MENSAJE_RECIVE_IMPACTO_PIERNA_IZQ & CStr(incomingData.ReadInteger() & MENSAJE_2), 255, 0, 0, True, False, True)
        Case bPiernaDerecha
            Call General_Add_to_RichTextBox(frmMain.RecTxt, MENSAJE_1 & attacker & MENSAJE_RECIVE_IMPACTO_PIERNA_DER & CStr(incomingData.ReadInteger() & MENSAJE_2), 255, 0, 0, True, False, True)
        Case bTorso
            Call General_Add_to_RichTextBox(frmMain.RecTxt, MENSAJE_1 & attacker & MENSAJE_RECIVE_IMPACTO_TORSO & CStr(incomingData.ReadInteger() & MENSAJE_2), 255, 0, 0, True, False, True)
    End Select
End Sub

''
'   Handles the UserHittedUser message.

Private Sub HandleUserHittedUser()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.Length < 6 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Dim Victim As String
    
    Victim = CharList(incomingData.ReadInteger()).Nombre
    
    Select Case incomingData.ReadByte
        Case bCabeza
            Call General_Add_to_RichTextBox(frmMain.RecTxt, MENSAJE_PRODUCE_IMPACTO_1 & Victim & MENSAJE_PRODUCE_IMPACTO_CABEZA & CStr(incomingData.ReadInteger() & MENSAJE_2), 255, 0, 0, True, False, True)
        Case bBrazoIzquierdo
            Call General_Add_to_RichTextBox(frmMain.RecTxt, MENSAJE_PRODUCE_IMPACTO_1 & Victim & MENSAJE_PRODUCE_IMPACTO_BRAZO_IZQ & CStr(incomingData.ReadInteger() & MENSAJE_2), 255, 0, 0, True, False, True)
        Case bBrazoDerecho
            Call General_Add_to_RichTextBox(frmMain.RecTxt, MENSAJE_PRODUCE_IMPACTO_1 & Victim & MENSAJE_PRODUCE_IMPACTO_BRAZO_DER & CStr(incomingData.ReadInteger() & MENSAJE_2), 255, 0, 0, True, False, True)
        Case bPiernaIzquierda
            Call General_Add_to_RichTextBox(frmMain.RecTxt, MENSAJE_PRODUCE_IMPACTO_1 & Victim & MENSAJE_PRODUCE_IMPACTO_PIERNA_IZQ & CStr(incomingData.ReadInteger() & MENSAJE_2), 255, 0, 0, True, False, True)
        Case bPiernaDerecha
            Call General_Add_to_RichTextBox(frmMain.RecTxt, MENSAJE_PRODUCE_IMPACTO_1 & Victim & MENSAJE_PRODUCE_IMPACTO_PIERNA_DER & CStr(incomingData.ReadInteger() & MENSAJE_2), 255, 0, 0, True, False, True)
        Case bTorso
            Call General_Add_to_RichTextBox(frmMain.RecTxt, MENSAJE_PRODUCE_IMPACTO_1 & Victim & MENSAJE_PRODUCE_IMPACTO_TORSO & CStr(incomingData.ReadInteger() & MENSAJE_2), 255, 0, 0, True, False, True)
    End Select
End Sub

''
'   Handles the UserIndexInServer message.

Private Sub HandleUserIndexInServer()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.Length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    UserIndex = incomingData.ReadInteger()
End Sub

''
'   Handles the UserNameList message.

Private Sub HandleUserNameList()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.Length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo ErrHandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call Buffer.ReadByte
    
    Dim userList() As String
    Dim i As Long
    
    userList = Split(Buffer.ReadASCIIString(), SEPARATOR)
    
    If frmPanelGm.Visible Then
        frmPanelGm.cboListaUsus.Clear
        For i = 0 To UBound(userList())
            Call frmPanelGm.cboListaUsus.AddItem(userList(i))
        Next i
        If frmPanelGm.cboListaUsus.ListCount > 0 Then frmPanelGm.cboListaUsus.ListIndex = 0
    End If
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(Buffer)
    
ErrHandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error <> 0 Then _
        Err.Raise Error
End Sub

''
'   Handles the UserOfferConfirm message.
Private Sub HandleUserOfferConfirm()
'***************************************************
'Author: ZaMa
'Last Modification: 14/12/2009
'
'***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    
    With frmComerciarUsu
        '   Now he can accept the offer or reject it
        .HabilitarAceptarRechazar True
        
        .PrintCommerceMsg TradingUserName & " ha confirmado su oferta!", FontTypeNames.FONTTYPE_CONSE
    End With
    
End Sub

''
'   Handles the UserSwing message.

Private Sub HandleUserSwing()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    
    Call General_Add_to_RichTextBox(frmMain.RecTxt, MENSAJE_FALLADO_GOLPE, 255, 0, 0, True, False, True)
End Sub

''
'   Handles the WorkRequestTarget message.

Private Sub HandleWorkRequestTarget()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.Length < 2 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    UsingSkill = incomingData.ReadByte()

    frmMain.MousePointer = 2
    
    Select Case UsingSkill
        Case Magia
            Call General_Add_to_RichTextBox(frmMain.RecTxt, MENSAJE_TRABAJO_MAGIA, 100, 100, 120, 0, 0)
        Case Pesca
            Call General_Add_to_RichTextBox(frmMain.RecTxt, MENSAJE_TRABAJO_PESCA, 100, 100, 120, 0, 0)
        Case Robar
            Call General_Add_to_RichTextBox(frmMain.RecTxt, MENSAJE_TRABAJO_ROBAR, 100, 100, 120, 0, 0)
        Case Talar
            Call General_Add_to_RichTextBox(frmMain.RecTxt, MENSAJE_TRABAJO_TALAR, 100, 100, 120, 0, 0)
        Case Mineria
            Call General_Add_to_RichTextBox(frmMain.RecTxt, MENSAJE_TRABAJO_MINERIA, 100, 100, 120, 0, 0)
        Case FundirMetal
            Call General_Add_to_RichTextBox(frmMain.RecTxt, MENSAJE_TRABAJO_FUNDIRMETAL, 100, 100, 120, 0, 0)
        Case Proyectiles
            Call General_Add_to_RichTextBox(frmMain.RecTxt, MENSAJE_TRABAJO_PROYECTILES, 100, 100, 120, 0, 0)
    End Select
End Sub

''
'   Initializes the fonts array

Public Sub InitFonts()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    With FontTypes(FontTypeNames.FONTTYPE_TALK)
        .Red = 255
        .Green = 255
        .Blue = 255
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_FIGHT)
        .Red = 255
        .bold = 1
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_WARNING)
        .Red = 32
        .Green = 51
        .Blue = 223
        .bold = 1
        .italic = 1
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_INFO)
        .Red = 65
        .Green = 190
        .Blue = 156
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_INFOBOLD)
        .Red = 65
        .Green = 190
        .Blue = 156
        .bold = 1
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_EJECUCION)
        .Red = 130
        .Green = 130
        .Blue = 130
        .bold = 1
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_PARTY)
        .Red = 255
        .Green = 180
        .Blue = 250
    End With
    
    FontTypes(FontTypeNames.FONTTYPE_VENENO).Green = 255
    
    With FontTypes(FontTypeNames.FONTTYPE_GUILD)
        .Red = 255
        .Green = 255
        .Blue = 255
        .bold = 1
    End With
    
    FontTypes(FontTypeNames.FONTTYPE_SERVER).Green = 185
    
    With FontTypes(FontTypeNames.FONTTYPE_GUILDMSG)
        .Red = 228
        .Green = 199
        .Blue = 27
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_CONSEJO)
        .Red = 130
        .Green = 130
        .Blue = 255
        .bold = 1
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_CONSEJOCAOS)
        .Red = 255
        .Green = 60
        .bold = 1
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_CONSEJOVesA)
        .Green = 200
        .Blue = 255
        .bold = 1
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_CONSEJOCAOSVesA)
        .Red = 255
        .Green = 50
        .bold = 1
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_CENTINELA)
        .Green = 255
        .bold = 1
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_GMMSG)
        .Red = 255
        .Green = 255
        .Blue = 255
        .italic = 1
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_GM)
        .Red = 30
        .Green = 255
        .Blue = 30
        .bold = 1
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_CITIZEN)
        .Blue = 200
        .bold = 1
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_CONSE)
        .Red = 30
        .Green = 150
        .Blue = 30
        .bold = 1
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_DIOS)
        .Red = 250
        .Green = 250
        .Blue = 150
        .bold = 1
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_DEATHMATCH)
        .Red = 0
        .Green = 128
        .Blue = 255
        .bold = 1
    End With
End Sub

Public Sub SalirDuelo()
    With outgoingData
        Call .WriteByte(ClientPacketID.Exit_Duel)
    End With
End Sub

''
'   Sends the data using the socket controls in the MainForm.
'
'   @param    sdData  The data to be sent to the server.

Private Sub SendData(ByRef sdData As String)
    
    'No enviamos nada si no estamos conectados
    If frmMain.Winsock1.State <> sckConnected Then Exit Sub

    'Send data!
    Call frmMain.Winsock1.SendData(sdData)
End Sub

''
'   Writes the "AcceptChaosCouncilMember" message to the outgoing data buffer.
'
'   @param    username The name of the user to be accepted as a chaos council member.
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteAcceptChaosCouncilMember(ByVal UserName As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "AcceptChaosCouncilMember" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.AcceptChaosCouncilMember)
        
        Call .WriteASCIIString(UserName)
    End With
End Sub

''
'   Writes the "AcceptRoyalCouncilMember" message to the outgoing data buffer.
'
'   @param    username The name of the user to be accepted into the royal army council.
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteAcceptRoyalCouncilMember(ByVal UserName As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "AcceptRoyalCouncilMember" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.AcceptRoyalCouncilMember)
        
        Call .WriteASCIIString(UserName)
    End With
End Sub

''
'   Writes the "AlterMail" message to the outgoing data buffer.
'
'   @param    username The name of the user whose mail is requested.
'   @param    newMail The new email of the player.
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteAlterMail(ByVal UserName As String, ByVal newMail As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "AlterMail" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.AlterMail)
        
        Call .WriteASCIIString(UserName)
        Call .WriteASCIIString(newMail)
    End With
End Sub

''
'   Writes the "AlterName" message to the outgoing data buffer.
'
'   @param    username The name of the user whose mail is requested.
'   @param    newName The new user name.
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteAlterName(ByVal UserName As String, ByVal newName As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "AlterName" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.AlterName)
        
        Call .WriteASCIIString(UserName)
        Call .WriteASCIIString(newName)
    End With
End Sub

''
'   Writes the "AlterPassword" message to the outgoing data buffer.
'
'   @param    username The name of the user whose mail is requested.
'   @param    copyFrom The name of the user from which to copy the password.
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteAlterPassword(ByVal UserName As String, ByVal CopyFrom As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "AlterPassword" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.AlterPassword)
        
        Call .WriteASCIIString(UserName)
        Call .WriteASCIIString(CopyFrom)
    End With
End Sub

''
'   Writes the "AskTrigger" message to the outgoing data buffer.
'
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteAskTrigger()
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 04/13/07
'Writes the "AskTrigger" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.AskTrigger)
End Sub

''
'   Writes the "Attack" message to the outgoing data buffer.
'
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteAttack()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Attack" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.Attack)
    
        IsAttacking = True
        CharList(UserCharIndex).Arma.WeaponWalk(CharList(UserCharIndex).Heading).Started = 1
End Sub

Public Sub WriteAutorizar(ByVal UserName As String)
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.DarPermisoMOD)
        
        Call .WriteASCIIString(UserName)
    End With
End Sub

''
'   Writes the "BanChar" message to the outgoing data buffer.
'
'   @param    username The user to be banned.
'   @param    reason The reson for which the user is to be banned.
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBanChar(ByVal UserName As String, ByVal reason As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "BanChar" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.banChar)
        
        Call .WriteASCIIString(UserName)
        
        Call .WriteASCIIString(reason)
    End With
End Sub

''
'   Writes the "BanIP" message to the outgoing data buffer.
'
'   @param    byIp    If set to true, we are banning by IP, otherwise the ip of a given character.
'   @param    IP      The IP for which to search for players. Must be an array of 4 elements with the 4 components of the IP.
'   @param    nick    The nick of the player whose ip will be banned.
'   @param    reason  The reason for the ban.
'
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBanIP(ByVal byIp As Boolean, ByRef Ip() As Byte, ByVal Nick As String, ByVal reason As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "BanIP" message to the outgoing data buffer
'***************************************************
    If byIp And UBound(Ip()) - LBound(Ip()) + 1 <> 4 Then Exit Sub   'Invalid IP
    
    Dim i As Long
    
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.BanIP)
        
        Call .WriteBoolean(byIp)
        
        If byIp Then
            For i = LBound(Ip()) To UBound(Ip())
                Call .WriteByte(Ip(i))
            Next i
        Else
            Call .WriteASCIIString(Nick)
        End If
        
        Call .WriteASCIIString(reason)
    End With
End Sub

''
'   Writes the "BankDeposit" message to the outgoing data buffer.
'
'   @param    slot Position within the user inventory in which the desired item is.
'   @param    amount Number of items to deposit.
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBankDeposit(ByVal Slot As Byte, ByVal Amount As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "BankDeposit" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.BankDeposit)
        
        Call .WriteByte(Slot)
        Call .WriteInteger(Amount)
    End With
End Sub

Public Sub WriteBankDepositCuenta(ByVal Slot As Byte, ByVal cantidad As Integer)
    With outgoingData
        Call .WriteByte(ClientPacketID.CuentaCL)
        Call .WriteASCIIString("BANKDEPOSIT")
        Call .WriteASCIIString(Cuenta.name)
        Call .WriteByte(Slot)
        Call .WriteInteger(cantidad)
    End With
End Sub

''
'   Writes the "BankDepositGold" message to the outgoing data buffer.
'
'   @param    amount The amount of money to deposit in the bank.
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBankDepositGold(ByVal Amount As Long)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "BankDepositGold" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.BankDepositGold)
        
        Call .WriteLong(Amount)
    End With
End Sub

''
'   Writes the "BankEnd" message to the outgoing data buffer.
'
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBankEnd()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "BankEnd" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.BankEnd)
End Sub

''
'   Writes the "BankExtractGold" message to the outgoing data buffer.
'
'   @param    amount The amount of money to extract from the bank.
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBankExtractGold(ByVal Amount As Long)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "BankExtractGold" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.BankExtractGold)
        
        Call .WriteLong(Amount)
    End With
End Sub

''
'   Writes the "BankExtractItem" message to the outgoing data buffer.
'
'   @param    slot Position within the bank in which the desired item is.
'   @param    amount Number of items to extract.
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBankExtractItem(ByVal Slot As Byte, ByVal Amount As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "BankExtractItem" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.BankExtractItem)
        
        Call .WriteByte(Slot)
        Call .WriteInteger(Amount)
    End With
End Sub

Public Sub WriteBankRetirarCuenta(ByVal Slot As Byte, ByVal cantidad As Integer)
    With outgoingData
        Call .WriteByte(ClientPacketID.CuentaCL)
        Call .WriteASCIIString("BANKRETIRAR")
        Call .WriteASCIIString(Cuenta.name)
        Call .WriteByte(Slot)
        Call .WriteInteger(cantidad)
    End With
End Sub

''
'   Writes the "BankStart" message to the outgoing data buffer.
'
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBankStart()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "BankStart" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.BankStart)
End Sub

Public Sub WriteBankStartCuenta()
    With outgoingData
        Call .WriteByte(ClientPacketID.CuentaCL)
        Call .WriteASCIIString("BANKSTART")
        Call .WriteASCIIString(Cuenta.name)
    End With
End Sub

''
'   Writes the "BannedIPList" message to the outgoing data buffer.
'
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBannedIPList()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "BannedIPList" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.BannedIPList)
End Sub

''
'   Writes the "BannedIPReload" message to the outgoing data buffer.
'
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBannedIPReload()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "BannedIPReload" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.BannedIPReload)
End Sub

Public Sub WriteBorrPJCuenta()
    With outgoingData
        Call .WriteByte(ClientPacketID.CuentaCL)
        Call .WriteASCIIString("DltPJAcc")
        Call .WriteASCIIString(Cuenta.name)
        Call .WriteByte(IndexSelectedUSer)
    End With
End Sub

''
'   Writes the "BugReport" message to the outgoing data buffer.
'
'   @param    message The message explaining the reported bug.
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBugReport(ByVal Message As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "BugReport" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.bugReport)
        
        Call .WriteASCIIString(Message)
    End With
End Sub

''
'   Writes the "CastSpell" message to the outgoing data buffer.
'
'   @param    slot Spell List slot where the spell to cast is.
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCastSpell(ByVal Slot As Byte)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "CastSpell" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.CastSpell)
        
        Call .WriteByte(Slot)
    End With
End Sub

''
'   Writes the "CentinelReport" message to the outgoing data buffer.
'
'   @param    number The number to report to the centinel.
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCentinelReport(ByVal Number As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "CentinelReport" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.CentinelReport)
        
        Call .WriteInteger(Number)
    End With
End Sub

''
'   Writes the "ChangeDescription" message to the outgoing data buffer.
'
'   @param    desc The new description of the user's character.
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeDescription(ByVal Desc As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ChangeDescription" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.ChangeDescription)
        
        Call .WriteASCIIString(Desc)
    End With
End Sub

''
'   Writes the "ChangeHeading" message to the outgoing data buffer.
'
'   @param    heading The direction in wich the user is moving.
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeHeading(ByVal Heading As E_Heading)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ChangeHeading" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.ChangeHeading)
        
        Call .WriteByte(Heading)
    End With
End Sub

''
'   Writes the "ChangeMapInfoBackup" message to the outgoing data buffer.
'
'   @param    backup True if the map is to be backuped, False otherwise.
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeMapInfoBackup(ByVal backup As Boolean)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ChangeMapInfoBackup" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.ChangeMapInfoBackup)
        
        Call .WriteBoolean(backup)
    End With
End Sub

                        
''
'   Writes the "ChangeMapInfoLand" message to the outgoing data buffer.
'
'   @param    land options: "BOSQUE", "NIEVE", "DESIERTO", "CIUDAD", "CAMPO", "DUNGEON".
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeMapInfoLand(ByVal land As String)
'***************************************************
'Author: Pablo (ToxicWaste)
'Last Modification: 26/01/2007
'Writes the "ChangeMapInfoLand" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.ChangeMapInfoLand)
        
        Call .WriteASCIIString(land)
    End With
End Sub

''
'   Writes the "ChangeMapInfoNoInvi" message to the outgoing data buffer.
'
'   @param    noinvi TRUE if invisibility is not to be allowed in the map.
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeMapInfoNoInvi(ByVal noinvi As Boolean)
'***************************************************
'Author: Pablo (ToxicWaste)
'Last Modification: 26/01/2007
'Writes the "ChangeMapInfoNoInvi" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.ChangeMapInfoNoInvi)
        
        Call .WriteBoolean(noinvi)
    End With
End Sub

''
'   Writes the "ChangeMapInfoNoMagic" message to the outgoing data buffer.
'
'   @param    nomagic TRUE if no magic is to be allowed in the map.
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeMapInfoNoMagic(ByVal nomagic As Boolean)
'***************************************************
'Author: Pablo (ToxicWaste)
'Last Modification: 26/01/2007
'Writes the "ChangeMapInfoNoMagic" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.ChangeMapInfoNoMagic)
        
        Call .WriteBoolean(nomagic)
    End With
End Sub

                            
''
'   Writes the "ChangeMapInfoNoResu" message to the outgoing data buffer.
'
'   @param    noresu TRUE if resurection is not to be allowed in the map.
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeMapInfoNoResu(ByVal noresu As Boolean)
'***************************************************
'Author: Pablo (ToxicWaste)
'Last Modification: 26/01/2007
'Writes the "ChangeMapInfoNoResu" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.ChangeMapInfoNoResu)
        
        Call .WriteBoolean(noresu)
    End With
End Sub

''
'   Writes the "ChangeMapInfoPK" message to the outgoing data buffer.
'
'   @param    isPK True if the map is PK, False otherwise.
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeMapInfoPK(ByVal isPK As Boolean)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ChangeMapInfoPK" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.ChangeMapInfoPK)
        
        Call .WriteBoolean(isPK)
    End With
End Sub

''
'   Writes the "ChangeMapInfoRestricted" message to the outgoing data buffer.
'
'   @param    restrict NEWBIES (only newbies), NO (everyone), ARMADA (just Armadas), CAOS (just caos) or FACCION (Armadas & caos only)
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeMapInfoRestricted(ByVal restrict As String)
'***************************************************
'Author: Pablo (ToxicWaste)
'Last Modification: 26/01/2007
'Writes the "ChangeMapInfoRestricted" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.ChangeMapInfoRestricted)
        
        Call .WriteASCIIString(restrict)
    End With
End Sub

                        
''
'   Writes the "ChangeMapInfoZone" message to the outgoing data buffer.
'
'   @param    zone options: "BOSQUE", "NIEVE", "DESIERTO", "CIUDAD", "CAMPO", "DUNGEON".
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeMapInfoZone(ByVal zone As String)
'***************************************************
'Author: Pablo (ToxicWaste)
'Last Modification: 26/01/2007
'Writes the "ChangeMapInfoZone" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.ChangeMapInfoZone)
        
        Call .WriteASCIIString(zone)
    End With
End Sub

''
'   Writes the "ChangeMOTD" message to the outgoing data buffer.
'
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeMOTD()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ChangeMOTD" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.ChangeMOTD)
End Sub

''
'   Writes the "ChangePassword" message to the outgoing data buffer.
'
'   @param    oldPass Previous password.
'   @param    newPass New password.
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangePassword(ByRef oldPass As String, ByRef newPass As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 10/10/07
'Last Modified By: Rapsodius
'Writes the "ChangePassword" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.ChangePassword)
        
        Call .WriteASCIIString(oldPass)
        Call .WriteASCIIString(newPass)
    End With
End Sub

''
'   Writes the "ChaosArmour" message to the outgoing data buffer.
'
'   @param    armourIndex The index of chaos armour to be altered.
'   @param    objectIndex The index of the new object to be set as the chaos armour.
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChaosArmour(ByVal armourIndex As Byte, ByVal objectIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ChaosArmour" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.ChaosArmour)
        
        Call .WriteByte(armourIndex)
        
        Call .WriteInteger(objectIndex)
    End With
End Sub

''
'   Writes the "ChaosLegionKick" message to the outgoing data buffer.
'
'   @param    username The name of the user to be kicked from the Chaos Legion.
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChaosLegionKick(ByVal UserName As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ChaosLegionKick" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.ChaosLegionKick)
        
        Call .WriteASCIIString(UserName)
    End With
End Sub

''
'   Writes the "ChaosLegionMessage" message to the outgoing data buffer.
'
'   @param    message The message to send to the chaos legion member.
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChaosLegionMessage(ByVal Message As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ChaosLegionMessage" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.ChaosLegionMessage)
        
        Call .WriteASCIIString(Message)
    End With
End Sub

''
'   Writes the "ChatColor" message to the outgoing data buffer.
'
'   @param    r The red component of the new chat color.
'   @param    g The green component of the new chat color.
'   @param    b The blue component of the new chat color.
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChatColor(ByVal r As Byte, ByVal g As Byte, ByVal b As Byte)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ChatColor" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.ChatColor)
        
        Call .WriteByte(r)
        Call .WriteByte(g)
        Call .WriteByte(b)
    End With
End Sub

Public Sub WriteChatFaccion(ByVal Message As String)
'***************************************************
'Author: Standelf
'Last Modification: 22/05/2010
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.FaccionChat)
        Call .WriteASCIIString(Message)
    End With
End Sub

Public Sub WriteChatGlobal(ByVal Message As String)
'***************************************************
'Author: Standelf
'Last Modification: 22/05/2010
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GlobalChat)
        Call .WriteASCIIString(Message)
    End With
End Sub

''
'   Writes the "CheckSlot" message to the outgoing data buffer.
'
'   @param    UserName    The name of the char whose slot will be checked.
'   @param    slot        The slot to be checked.
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCheckSlot(ByVal UserName As String, ByVal Slot As Byte)
'***************************************************
'Author: Pablo (ToxicWaste)
'Last Modification: 26/01/2007
'Writes the "CheckSlot" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.CheckSlot)
        Call .WriteASCIIString(UserName)
        Call .WriteByte(Slot)
    End With
End Sub

''
'   Writes the "CitizenMessage" message to the outgoing data buffer.
'
'   @param    message The message to send to citizens.
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCitizenMessage(ByVal Message As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "CitizenMessage" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.CitizenMessage)
        
        Call .WriteASCIIString(Message)
    End With
End Sub

''
'   Writes the "ClanCodexUpdate" message to the outgoing data buffer.
'
'   @param    desc New description of the clan.
'   @param    codex New codex of the clan.
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteClanCodexUpdate(ByVal Desc As String, ByRef Codex() As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ClanCodexUpdate" message to the outgoing data buffer
'***************************************************
    Dim Temp As String
    Dim i As Long
    
    With outgoingData
        Call .WriteByte(ClientPacketID.ClanCodexUpdate)
        
        Call .WriteASCIIString(Desc)
        
        For i = LBound(Codex()) To UBound(Codex())
            Temp = Temp & Codex(i) & SEPARATOR
        Next i
        
        If Len(Temp) Then _
            Temp = Left$(Temp, Len(Temp) - 1)
        
        Call .WriteASCIIString(Temp)
    End With
End Sub

''
'   Writes the "CleanSOS" message to the outgoing data buffer.
'
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCleanSOS()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "CleanSOS" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.CleanSOS)
End Sub

''
'   Writes the "CleanWorld" message to the outgoing data buffer.
'
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCleanWorld()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "CleanWorld" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.CleanWorld)
End Sub

''
'   Writes the "Comment" message to the outgoing data buffer.
'
'   @param    message The message to leave in the log as a comment.
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteComment(ByVal Message As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Comment" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.Comment)
        
        Call .WriteASCIIString(Message)
    End With
End Sub

''
'   Writes the "CommerceBuy" message to the outgoing data buffer.
'
'   @param    slot Position within the NPC's inventory in which the desired item is.
'   @param    amount Number of items to buy.
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCommerceBuy(ByVal Slot As Byte, ByVal Amount As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "CommerceBuy" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.CommerceBuy)
        
        Call .WriteByte(Slot)
        Call .WriteInteger(Amount)
    End With
End Sub

Public Sub WriteCommerceChat(ByVal chat As String)
'***************************************************
'Author: ZaMa
'Last Modification: 03/12/2009
'Writes the "CommerceChat" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.CommerceChat)
        
        Call .WriteASCIIString(chat)
    End With
End Sub

''
'   Writes the "CommerceEnd" message to the outgoing data buffer.
'
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCommerceEnd()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "CommerceEnd" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.CommerceEnd)
End Sub

''
'   Writes the "CommerceSell" message to the outgoing data buffer.
'
'   @param    slot Position within user inventory in which the desired item is.
'   @param    amount Number of items to sell.
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCommerceSell(ByVal Slot As Byte, ByVal Amount As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "CommerceSell" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.CommerceSell)
        
        Call .WriteByte(Slot)
        Call .WriteInteger(Amount)
    End With
End Sub

''
'   Writes the "CommerceStart" message to the outgoing data buffer.
'
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCommerceStart()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "CommerceStart" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.CommerceStart)
End Sub

''
'   Writes the "Consulta" message to the outgoing data buffer.
'
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteConsulta()
'***************************************************
'Author: ZaMa
'Last Modification: 01/05/2010
'Writes the "Consulta" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.Consulta)

End Sub

 
Public Sub WriteConsultaSubasta()
'***************************************************
'Author: Standelf
'Last Modification: 25/05/10
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.ConsultaSubasta)
    End With
End Sub

''
'   Writes the "CouncilKick" message to the outgoing data buffer.
'
'   @param    username The name of the user to be kicked from the council.
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCouncilKick(ByVal UserName As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "CouncilKick" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.CouncilKick)
        
        Call .WriteASCIIString(UserName)
    End With
End Sub

''
'   Writes the "CouncilMessage" message to the outgoing data buffer.
'
'   @param    message The message to send to the other council members.
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCouncilMessage(ByVal Message As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "CouncilMessage" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.CouncilMessage)
        
        Call .WriteASCIIString(Message)
    End With
End Sub

''
'   Writes the "CraftBlacksmith" message to the outgoing data buffer.
'
'   @param    item Index of the item to craft in the list sent by the server.
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCraftBlacksmith(ByVal Item As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "CraftBlacksmith" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.CraftBlacksmith)
        
        Call .WriteInteger(Item)
    End With
End Sub

''
'   Writes the "CraftCarpenter" message to the outgoing data buffer.
'
'   @param    item Index of the item to craft in the list sent by the server.
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCraftCarpenter(ByVal Item As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "CraftCarpenter" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.CraftCarpenter)
        
        Call .WriteInteger(Item)
    End With
End Sub

Public Sub WriteCrearCuenta()
    With outgoingData
        Call .WriteByte(ClientPacketID.CuentaCL)
        Call .WriteASCIIString("NewAcc")
        Call .WriteASCIIString(Cuenta.name)
        Call .WriteASCIIString(Cuenta.Pass)
        Call .WriteASCIIString(Cuenta.Email)
    End With
End Sub

''
'   Writes the "CreateItem" message to the outgoing data buffer.
'
'   @param    itemIndex The index of the item to be created.
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCreateItem(ByVal ItemIndex As Long)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "CreateItem" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.CreateItem)
        Call .WriteInteger(ItemIndex)
    End With
End Sub

''
'   Writes the "CreateNewGuild" message to the outgoing data buffer.
'
'   @param    desc    The guild's description
'   @param    name    The guild's name
'   @param    site    The guild's website
'   @param    codex   Array of all rules of the guild.
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCreateNewGuild(ByVal Desc As String, ByVal name As String, ByVal Site As String, ByRef Codex() As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "CreateNewGuild" message to the outgoing data buffer
'***************************************************
    Dim Temp As String
    Dim i As Long
    
    With outgoingData
        Call .WriteByte(ClientPacketID.CreateNewGuild)
        
        Call .WriteASCIIString(Desc)
        Call .WriteASCIIString(name)
        Call .WriteASCIIString(Site)
        
        For i = LBound(Codex()) To UBound(Codex())
            Temp = Temp & Codex(i) & SEPARATOR
        Next i
        
        If Len(Temp) Then _
            Temp = Left$(Temp, Len(Temp) - 1)
        
        Call .WriteASCIIString(Temp)
    End With
End Sub

''
'   Writes the "CreateNPC" message to the outgoing data buffer.
'
'   @param    npcIndex The index of the NPC to be created.
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCreateNPC(ByVal NPCIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "CreateNPC" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.CreateNPC)
        
        Call .WriteInteger(NPCIndex)
    End With
End Sub

''
'   Writes the "CreateNPCWithRespawn" message to the outgoing data buffer.
'
'   @param    npcIndex The index of the NPC to be created.
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCreateNPCWithRespawn(ByVal NPCIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "CreateNPCWithRespawn" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.CreateNPCWithRespawn)
        
        Call .WriteInteger(NPCIndex)
    End With
End Sub

''
'   Writes the "CreaturesInMap" message to the outgoing data buffer.
'
'   @param    map The map in which to check for the existing creatures.
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCreaturesInMap(ByVal Map As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "CreaturesInMap" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.CreaturesInMap)
        
        Call .WriteInteger(Map)
    End With
End Sub

''
'   Writes the "CriminalMessage" message to the outgoing data buffer.
'
'   @param    message The message to send to criminals.
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCriminalMessage(ByVal Message As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "CriminalMessage" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.CriminalMessage)
        
        Call .WriteASCIIString(Message)
    End With
End Sub

Public Sub WriteCuentaRegresiva(ByVal Cuenta As Byte)
    With outgoingData
        Call .WriteByte(ClientPacketID.CuentaRegresiva)
        Call .WriteByte(Cuenta)
    End With
End Sub

''
'   Writes the "Denounce" message to the outgoing data buffer.
'
'   @param    message The message to send with the denounce.
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteDenounce(ByVal Message As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Denounce" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.Denounce)
        
        Call .WriteASCIIString(Message)
    End With
End Sub

''
'   Writes the "DestroyAllItemsInArea" message to the outgoing data buffer.
'
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteDestroyAllItemsInArea()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "DestroyAllItemsInArea" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.DestroyAllItemsInArea)
End Sub

''
'   Writes the "DestroyItems" message to the outgoing data buffer.
'
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteDestroyItems()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "DestroyItems" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.DestroyItems)
End Sub

''
'   Writes the "DoBackup" message to the outgoing data buffer.
'
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteDoBackup()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "DoBackup" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.DoBackUp)
End Sub

''
'   Writes the "DoubleClick" message to the outgoing data buffer.
'
'   @param    x Tile coord in the x-axis in which the user clicked.
'   @param    y Tile coord in the y-axis in which the user clicked.
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteDoubleClick(ByVal X As Byte, ByVal Y As Byte)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "DoubleClick" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.DoubleClick)
        
        Call .WriteByte(X)
        Call .WriteByte(Y)
    End With
End Sub

Public Sub WriteDrop(ByVal Slot As Byte, ByVal Amount As Integer, ByVal X As Byte, Y As Byte)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Drop" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.Drop)
        
        Call .WriteByte(Slot)
        Call .WriteInteger(Amount)
        
        Call .WriteByte(X)
        Call .WriteByte(Y)
    End With
End Sub

''
'   Writes the "DumpIPTables" message to the outgoing data buffer.
'
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteDumpIPTables()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "DumpIPTables" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.dumpIPTables)
End Sub

Public Sub WriteEditame()
'***************************************************
'Author: Standelf
'Last Modification: 21/05/10
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.Editame)
    End With
End Sub

''
'   Writes the "EditChar" message to the outgoing data buffer.
'
'   @param    UserName    The user to be edited.
'   @param    editOption  Indicates what to edit in the char.
'   @param    arg1        Additional argument 1. Contents depend on editoption.
'   @param    arg2        Additional argument 2. Contents depend on editoption.
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteEditChar(ByVal UserName As String, ByVal EditOption As eEditOptions, ByVal arg1 As String, ByVal arg2 As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "EditChar" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.EditChar)
        
        Call .WriteASCIIString(UserName)
        
        Call .WriteByte(EditOption)
        
        Call .WriteASCIIString(arg1)
        Call .WriteASCIIString(arg2)
    End With
End Sub

''
'   Writes the "Enlist" message to the outgoing data buffer.
'
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteEnlist()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Enlist" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.Enlist)
End Sub

''
'   Writes the "EquipItem" message to the outgoing data buffer.
'
'   @param    slot Invetory slot where the item to equip is.
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteEquipItem(ByVal Slot As Byte)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "EquipItem" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.EquipItem)
        
        Call .WriteByte(Slot)
    End With
End Sub

''
'   Writes the "Execute" message to the outgoing data buffer.
'
'   @param    username The user to be executed.
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteExecute(ByVal UserName As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Execute" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.Execute)
        
        Call .WriteASCIIString(UserName)
    End With
End Sub

''
'   Writes the "ForceMIDIAll" message to the outgoing data buffer.
'
'   @param    midiID The id of the midi file to play.
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteForceMIDIAll(ByVal midiID As Byte)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ForceMIDIAll" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.ForceMIDIAll)
        
        Call .WriteByte(midiID)
    End With
End Sub

''
'   Writes the "ForceMIDIToMap" message to the outgoing data buffer.
'
'   @param    midiID The ID of the midi file to play.
'   @param    map The map in which to play the given midi.
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteForceMIDIToMap(ByVal midiID As Byte, ByVal Map As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ForceMIDIToMap" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.ForceMIDIToMap)
        
        Call .WriteByte(midiID)
        
        Call .WriteInteger(Map)
    End With
End Sub

''
'   Writes the "ForceWAVEAll" message to the outgoing data buffer.
'
'   @param    waveID The id of the wave file to play.
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteForceWAVEAll(ByVal waveID As Byte)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ForceWAVEAll" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.ForceWAVEAll)
        
        Call .WriteByte(waveID)
    End With
End Sub

''
'   Writes the "ForceWAVEToMap" message to the outgoing data buffer.
'
'   @param    waveID  The ID of the wave file to play.
'   @param    Map     The map into which to play the given wave.
'   @param    x       The position in the x axis in which to play the given wave.
'   @param    y       The position in the y axis in which to play the given wave.
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteForceWAVEToMap(ByVal waveID As Byte, ByVal Map As Integer, ByVal X As Byte, ByVal Y As Byte)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ForceWAVEToMap" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.ForceWAVEToMap)
        
        Call .WriteByte(waveID)
        
        Call .WriteInteger(Map)
        
        Call .WriteByte(X)
        Call .WriteByte(Y)
    End With
End Sub

''
'   Writes the "Forgive" message to the outgoing data buffer.
'
'   @param    username The user to be forgiven.
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteForgive(ByVal UserName As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Forgive" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.Forgive)
        
        Call .WriteASCIIString(UserName)
    End With
End Sub

''
'   Writes the "ForumPost" message to the outgoing data buffer.
'
'   @param    title The message's title.
'   @param    message The body of the message.
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteForumPost(ByVal Title As String, ByVal Message As String, ByVal ForumMsgType As Byte)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ForumPost" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.ForumPost)
        
        Call .WriteByte(ForumMsgType)
        Call .WriteASCIIString(Title)
        Call .WriteASCIIString(Message)
    End With
End Sub

''
'   Writes the "Gamble" message to the outgoing data buffer.
'
'   @param    amount The amount to gamble.
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGamble(ByVal Amount As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Gamble" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.Gamble)
        
        Call .WriteInteger(Amount)
    End With
End Sub

Public Sub WriteGiveMePower(ByVal X As Integer, Y As Integer)
'***************************************************
'Author: Ezequiel Juárez (Standelf)
'Last Modification: 08/02/2011
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GiveMePower)
        Call .WriteInteger(X)
        Call .WriteInteger(Y)
    End With
End Sub

''
'   Writes the "GMMessage" message to the outgoing data buffer.
'
'   @param    message The message to be sent to the other GMs online.
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGMMessage(ByVal Message As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "GMMessage" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.GMMessage)
        Call .WriteASCIIString(Message)
    End With
End Sub

''
'   Writes the "GMPanel" message to the outgoing data buffer.
'
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGMPanel()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "GMPanel" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.GMPanel)
End Sub

''
'   Writes the "GMRequest" message to the outgoing data buffer.
'
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGMRequest()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "GMRequest" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMRequest)
End Sub

''
'   Writes the "GoNearby" message to the outgoing data buffer.
'
'   @param    username The suer to approach.
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGoNearby(ByVal UserName As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "GoNearby" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call outgoingData.WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.GoNearby)
        
        Call .WriteASCIIString(UserName)
    End With
End Sub

''
'   Writes the "GoToChar" message to the outgoing data buffer.
'
'   @param    username The user to be approached.
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGoToChar(ByVal UserName As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "GoToChar" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.GoToChar)
        
        Call .WriteASCIIString(UserName)
    End With
End Sub

''
'   Writes the "GuildAcceptAlliance" message to the outgoing data buffer.
'
'   @param    guild The guild whose aliance offer is accepted.
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildAcceptAlliance(ByVal guild As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "GuildAcceptAlliance" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GuildAcceptAlliance)
        
        Call .WriteASCIIString(guild)
    End With
End Sub

''
'   Writes the "GuildAcceptNewMember" message to the outgoing data buffer.
'
'   @param    username The name of the accepted player.
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildAcceptNewMember(ByVal UserName As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "GuildAcceptNewMember" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GuildAcceptNewMember)
        
        Call .WriteASCIIString(UserName)
    End With
End Sub

''
'   Writes the "GuildAcceptPeace" message to the outgoing data buffer.
'
'   @param    guild The guild whose peace offer is accepted.
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildAcceptPeace(ByVal guild As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "GuildAcceptPeace" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GuildAcceptPeace)
        
        Call .WriteASCIIString(guild)
    End With
End Sub

''
'   Writes the "GuildAllianceDetails" message to the outgoing data buffer.
'
'   @param    guild The guild whose aliance proposal's details are requested.
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildAllianceDetails(ByVal guild As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "GuildAllianceDetails" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GuildAllianceDetails)
        
        Call .WriteASCIIString(guild)
    End With
End Sub

''
'   Writes the "GuildAlliancePropList" message to the outgoing data buffer.
'
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildAlliancePropList()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "GuildAlliancePropList" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GuildAlliancePropList)
End Sub

''
'   Writes the "GuildBan" message to the outgoing data buffer.
'
'   @param    guild The guild whose members will be banned.
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildBan(ByVal guild As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "GuildBan" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.GuildBan)
        
        Call .WriteASCIIString(guild)
    End With
End Sub

''
'   Writes the "GuildDeclareWar" message to the outgoing data buffer.
'
'   @param    guild The guild to which to declare war.
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildDeclareWar(ByVal guild As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "GuildDeclareWar" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GuildDeclareWar)
        
        Call .WriteASCIIString(guild)
    End With
End Sub

''
'   Writes the "GuildFundate" message to the outgoing data buffer.
'
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildFundate()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 03/21/2001
'Writes the "GuildFundate" message to the outgoing data buffer
'14/12/2009: ZaMa - Now first checks if the user can foundate a guild.
'03/21/2001: Pato - Deleted de clanType param.
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GuildFundate)
End Sub

''
'   Writes the "GuildFundation" message to the outgoing data buffer.
'
'   @param    clanType The alignment of the clan to be founded.
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildFundation(ByVal clanType As eClanType)
'***************************************************
'Author: ZaMa
'Last Modification: 14/12/2009
'Writes the "GuildFundation" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GuildFundation)
        
        Call .WriteByte(clanType)
    End With
End Sub

''
'   Writes the "GuildKickMember" message to the outgoing data buffer.
'
'   @param    username The name of the kicked player.
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildKickMember(ByVal UserName As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "GuildKickMember" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GuildKickMember)
        
        Call .WriteASCIIString(UserName)
    End With
End Sub

''
'   Writes the "GuildLeave" message to the outgoing data buffer.
'
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildLeave()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "GuildLeave" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GuildLeave)
End Sub

''
'   Writes the "GuildMemberInfo" message to the outgoing data buffer.
'
'   @param    username The user whose info is requested.
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildMemberInfo(ByVal UserName As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "GuildMemberInfo" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GuildMemberInfo)
        
        Call .WriteASCIIString(UserName)
    End With
End Sub

''
'   Writes the "GuildMemberList" message to the outgoing data buffer.
'
'   @param    guild The guild whose member list is requested.
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildMemberList(ByVal guild As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "GuildMemberList" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.GuildMemberList)
        
        Call .WriteASCIIString(guild)
    End With
End Sub

''
'   Writes the "GuildMessage" message to the outgoing data buffer.
'
'   @param    message The message to send to the guild.
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildMessage(ByVal Message As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "GuildRequestDetails" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GuildMessage)
        
        Call .WriteASCIIString(Message)
    End With
End Sub

''
'   Writes the "GuildNewWebsite" message to the outgoing data buffer.
'
'   @param    url The guild's new website's URL.
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildNewWebsite(ByVal URL As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "GuildNewWebsite" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GuildNewWebsite)
        
        Call .WriteASCIIString(URL)
    End With
End Sub

''
'   Writes the "GuildOfferAlliance" message to the outgoing data buffer.
'
'   @param    guild The guild to whom an aliance is offered.
'   @param    proposal The text to send with the proposal.
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildOfferAlliance(ByVal guild As String, ByVal proposal As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "GuildOfferAlliance" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GuildOfferAlliance)
        
        Call .WriteASCIIString(guild)
        Call .WriteASCIIString(proposal)
    End With
End Sub

''
'   Writes the "GuildOfferPeace" message to the outgoing data buffer.
'
'   @param    guild The guild to whom peace is offered.
'   @param    proposal The text to send with the proposal.
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildOfferPeace(ByVal guild As String, ByVal proposal As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "GuildOfferPeace" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GuildOfferPeace)
        
        Call .WriteASCIIString(guild)
        Call .WriteASCIIString(proposal)
    End With
End Sub

''
'   Writes the "GuildOnline" message to the outgoing data buffer.
'
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildOnline()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "GuildOnline" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GuildOnline)
End Sub

''
'   Writes the "GuildOnlineMembers" message to the outgoing data buffer.
'
'   @param    guild The guild whose online player list is requested.
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildOnlineMembers(ByVal guild As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "GuildOnlineMembers" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.GuildOnlineMembers)
        
        Call .WriteASCIIString(guild)
    End With
End Sub

''
'   Writes the "GuildOpenElections" message to the outgoing data buffer.
'
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildOpenElections()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "GuildOpenElections" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GuildOpenElections)
End Sub

''
'   Writes the "GuildPeaceDetails" message to the outgoing data buffer.
'
'   @param    guild The guild whose peace proposal's details are requested.
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildPeaceDetails(ByVal guild As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "GuildPeaceDetails" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GuildPeaceDetails)
        
        Call .WriteASCIIString(guild)
    End With
End Sub

''
'   Writes the "GuildPeacePropList" message to the outgoing data buffer.
'
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildPeacePropList()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "GuildPeacePropList" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GuildPeacePropList)
End Sub

''
'   Writes the "GuildRejectAlliance" message to the outgoing data buffer.
'
'   @param    guild The guild whose aliance offer is rejected.
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildRejectAlliance(ByVal guild As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "GuildRejectAlliance" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GuildRejectAlliance)
        
        Call .WriteASCIIString(guild)
    End With
End Sub

''
'   Writes the "GuildRejectNewMember" message to the outgoing data buffer.
'
'   @param    username The name of the rejected player.
'   @param    reason The reason for which the player was rejected.
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildRejectNewMember(ByVal UserName As String, ByVal reason As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "GuildRejectNewMember" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GuildRejectNewMember)
        
        Call .WriteASCIIString(UserName)
        Call .WriteASCIIString(reason)
    End With
End Sub

''
'   Writes the "GuildRejectPeace" message to the outgoing data buffer.
'
'   @param    guild The guild whose peace offer is rejected.
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildRejectPeace(ByVal guild As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "GuildRejectPeace" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GuildRejectPeace)
        
        Call .WriteASCIIString(guild)
    End With
End Sub

''
'   Writes the "GuildRequestDetails" message to the outgoing data buffer.
'
'   @param    guild The guild whose details are requested.
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildRequestDetails(ByVal guild As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "GuildRequestDetails" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GuildRequestDetails)
        
        Call .WriteASCIIString(guild)
    End With
End Sub

''
'   Writes the "GuildRequestJoinerInfo" message to the outgoing data buffer.
'
'   @param    username The user who wants to join the guild whose info is requested.
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildRequestJoinerInfo(ByVal UserName As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "GuildRequestJoinerInfo" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GuildRequestJoinerInfo)
        
        Call .WriteASCIIString(UserName)
    End With
End Sub

''
'   Writes the "GuildRequestMembership" message to the outgoing data buffer.
'
'   @param    guild The guild to which to request membership.
'   @param    application The user's application sheet.
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildRequestMembership(ByVal guild As String, ByVal Application As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "GuildRequestMembership" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GuildRequestMembership)
        
        Call .WriteASCIIString(guild)
        Call .WriteASCIIString(Application)
    End With
End Sub

'   Writes the "GuildResolve" message to the outgoing data buffer.
'
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildResolve()
'***************************************************
'Author: Micaela Buono Pugh (Khalem)
'Last Modification: 03/21/2001
'Writes the "GuilddResolve" message to the outgoing data buffer
'14/12/2009: ZaMa - Now first checks if the user can foundate a guild.
'03/21/2001: Pato - Deleted de clanType param.
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GuildResolve)
End Sub

''
'   Writes the "GuildUpdateNews" message to the outgoing data buffer.
'
'   @param    news The news to be posted.
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildUpdateNews(ByVal news As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "GuildUpdateNews" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GuildUpdateNews)
        
        Call .WriteASCIIString(news)
    End With
End Sub

''
'   Writes the "GuildVote" message to the outgoing data buffer.
'
'   @param    username The user to vote for clan leader.
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildVote(ByVal UserName As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "GuildVote" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GuildVote)
        
        Call .WriteASCIIString(UserName)
    End With
End Sub

''
'   Writes the "Heal" message to the outgoing data buffer.
'
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteHeal()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Heal" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.Heal)
End Sub

''
'   Writes the "Help" message to the outgoing data buffer.
'
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteHelp()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Help" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.Help)
End Sub

''
'   Writes the "Hiding" message to the outgoing data buffer.
'
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteHiding()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Hiding" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.Hiding)
End Sub

'   Writes the "Home" message to the outgoing data buffer.
'
Public Sub WriteHome()
'***************************************************
'Author: Budi
'Last Modification: 01/06/10
'Writes the "Home" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.Home)
    End With
End Sub

''
'   Writes the "Ignored" message to the outgoing data buffer.
'
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteIgnored()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Ignored" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.Ignored)
End Sub

''
'   Writes the "ImperialArmour" message to the outgoing data buffer.
'
'   @param    armourIndex The index of imperial armour to be altered.
'   @param    objectIndex The index of the new object to be set as the imperial armour.
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteImperialArmour(ByVal armourIndex As Byte, ByVal objectIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ImperialArmour" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.ImperialArmour)
        
        Call .WriteByte(armourIndex)
        
        Call .WriteInteger(objectIndex)
    End With
End Sub

''
'   Writes the "Information" message to the outgoing data buffer.
'
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteInformation()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Information" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.Information)
End Sub

Public Sub WriteIniciarSubasta(ByVal Slot As Integer, ByVal cantidad As Integer, ByVal Valor As Long)
'***************************************************
'Author: Standelf
'Last Modification: 25/05/10
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.IniciarSubasta)

        Call .WriteInteger(Slot)
        Call .WriteInteger(cantidad)
        Call .WriteLong(Valor)
       
        Unload frmSubastar
    End With
End Sub

''
'   Writes the "InitCrafting" message to the outgoing data buffer.
'
'   @param    Cantidad The final aumont of item to craft.
'   @param    NroPorCiclo The amount of items to craft per cicle.

Public Sub WriteInitCrafting(ByVal cantidad As Long, ByVal NroPorCiclo As Integer)
'***************************************************
'Author: ZaMa
'Last Modification: 29/01/2010
'Writes the "InitCrafting" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.InitCrafting)
        Call .WriteLong(cantidad)
        
        Call .WriteInteger(NroPorCiclo)
    End With
End Sub

Public Sub WriteInitInvasion(ByVal NPC As Integer)
'***************************************************
'Author: Standelf
'Last Modification: 22/05/10
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.Init_Invasion)
        Call .WriteInteger(NPC)
    End With
End Sub

''
'   Writes the "Inquiry" message to the outgoing data buffer.
'
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteInquiry()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Inquiry" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.Inquiry)
End Sub

''
'   Writes the "InquiryVote" message to the outgoing data buffer.
'
'   @param    opt The chosen option to vote for.
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteInquiryVote(ByVal opt As Byte)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "InquiryVote" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.InquiryVote)
        
        Call .WriteByte(opt)
    End With
End Sub

''
'   Writes the "invisible" message to the outgoing data buffer.
'
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteInvisible()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "invisible" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.invisible)
End Sub

''
'   Writes the "IPToNick" message to the outgoing data buffer.
'
'   @param    IP The IP for which to search for players. Must be an array of 4 elements with the 4 components of the IP.
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteIPToNick(ByRef Ip() As Byte)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "IPToNick" message to the outgoing data buffer
'***************************************************
    If UBound(Ip()) - LBound(Ip()) + 1 <> 4 Then Exit Sub   'Invalid IP
    
    Dim i As Long
    
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.IPToNick)
        
        For i = LBound(Ip()) To UBound(Ip())
            Call .WriteByte(Ip(i))
        Next i
    End With
End Sub

''
'   Writes the "ItemsInTheFloor" message to the outgoing data buffer.
'
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteItemsInTheFloor()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ItemsInTheFloor" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.ItemsInTheFloor)
End Sub

''
'   Writes the "ItemUpgrade" message to the outgoing data buffer.
'
'   @param    ItemIndex The index to the item to upgrade.
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteItemUpgrade(ByVal ItemIndex As Integer)
'***************************************************
'Author: Torres Patricio (Pato)
'Last Modification: 12/09/09
'Writes the "ItemUpgrade" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.ItemUpgrade)
    Call outgoingData.WriteInteger(ItemIndex)
End Sub

''
'   Writes the "Jail" message to the outgoing data buffer.
'
'   @param    username The user to be sent to jail.
'   @param    reason The reason for which to send him to jail.
'   @param    time The time (in minutes) the user will have to spend there.
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteJail(ByVal UserName As String, ByVal reason As String, ByVal Time As Byte)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Jail" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.Jail)
        
        Call .WriteASCIIString(UserName)
        Call .WriteASCIIString(reason)
        
        Call .WriteByte(Time)
    End With
End Sub

''
'   Writes the "Kick" message to the outgoing data buffer.
'
'   @param    username The user to be kicked.
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteKick(ByVal UserName As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Kick" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.Kick)
        
        Call .WriteASCIIString(UserName)
    End With
End Sub

''
'   Writes the "KickAllChars" message to the outgoing data buffer.
'
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteKickAllChars()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "KickAllChars" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.KickAllChars)
End Sub

''
'   Writes the "KillAllNearbyNPCs" message to the outgoing data buffer.
'
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteKillAllNearbyNPCs()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "KillAllNearbyNPCs" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.KillAllNearbyNPCs)
End Sub

''
'   Writes the "KillNPC" message to the outgoing data buffer.
'
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteKillNPC()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "KillNPC" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.KillNPC)
End Sub

''
'   Writes the "KillNPCNoRespawn" message to the outgoing data buffer.
'
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteKillNPCNoRespawn()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "KillNPCNoRespawn" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.KillNPCNoRespawn)
End Sub

''
'   Writes the "LastIP" message to the outgoing data buffer.
'
'   @param    username The user whose last IPs are requested.
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteLastIP(ByVal UserName As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "LastIP" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.LastIP)
        
        Call .WriteASCIIString(UserName)
    End With
End Sub

''
'   Writes the "LeaveFaction" message to the outgoing data buffer.
'
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteLeaveFaction()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "LeaveFaction" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.LeaveFaction)
End Sub

''
'   Writes the "LeftClick" message to the outgoing data buffer.
'
'   @param    x Tile coord in the x-axis in which the user clicked.
'   @param    y Tile coord in the y-axis in which the user clicked.
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteLeftClick(ByVal X As Byte, ByVal Y As Byte)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "LeftClick" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.LeftClick)
        
        Call .WriteByte(X)
        Call .WriteByte(Y)
    End With
End Sub

Public Sub WriteLoginCuenta()
    With outgoingData
        Call .WriteByte(ClientPacketID.CuentaCL)
        Call .WriteASCIIString("LoginAcc")
        Call .WriteASCIIString(Cuenta.name)
        Call .WriteASCIIString(Cuenta.Pass)
    End With
End Sub

''
'   Writes the "LoginExistingChar" message to the outgoing data buffer.
'
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteLoginExistingChar()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "LoginExistingChar" message to the outgoing data buffer
'***************************************************
    Dim i As Long
    
    With outgoingData
        Call .WriteByte(ClientPacketID.LoginExistingChar)
        
        Call .WriteASCIIString(UserName)
        
        Call .WriteASCIIString(UserPassword)

        Call .WriteByte(App.Major)
        Call .WriteByte(App.Minor)
        Call .WriteByte(App.Revision)
        
        Call .WriteBoolean(Settings.ParticleMeditation)
        
    End With
End Sub

''
'   Writes the "LoginNewChar" message to the outgoing data buffer.
'
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteLoginNewChar()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "LoginNewChar" message to the outgoing data buffer
'***************************************************
    Dim i As Long
    
    With outgoingData
        Call .WriteByte(ClientPacketID.LoginNewChar)
        
        Call .WriteASCIIString(UserName)
                
        Call .WriteByte(App.Major)
        Call .WriteByte(App.Minor)
        Call .WriteByte(App.Revision)
        
        Call .WriteByte(UserRaza)
        Call .WriteByte(UserSexo)
        Call .WriteByte(UserClase)
        Call .WriteInteger(UserHead)
        
        Call .WriteByte(UserHogar)
        
        Call .WriteASCIIString(Cuenta.name)
        Call .WriteBoolean(Settings.ParticleMeditation)
    End With
End Sub

''
'   Writes the "MakeDumb" message to the outgoing data buffer.
'
'   @param    username The name of the user to be made dumb.
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteMakeDumb(ByVal UserName As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "MakeDumb" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.MakeDumb)
        
        Call .WriteASCIIString(UserName)
    End With
End Sub

''
'   Writes the "MakeDumbNoMore" message to the outgoing data buffer.
'
'   @param    username The name of the user who will no longer be dumb.
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteMakeDumbNoMore(ByVal UserName As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "MakeDumbNoMore" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.MakeDumbNoMore)
        
        Call .WriteASCIIString(UserName)
    End With
End Sub

''
'   Writes the "Meditate" message to the outgoing data buffer.
'
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteMeditate()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Meditate" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.Meditate)
End Sub

''
'   Writes the "ModifySkills" message to the outgoing data buffer.
'
'   @param    skillEdt a-based array containing for each skill the number of points to add to it.
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteModifySkills(ByRef skillEdt() As Byte)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ModifySkills" message to the outgoing data buffer
'***************************************************
    Dim i As Long
    
    With outgoingData
        Call .WriteByte(ClientPacketID.ModifySkills)
        
        For i = 1 To NUMSKILLS
            Call .WriteByte(skillEdt(i))
        Next i
    End With
End Sub

''
'   Writes the "MoveBank" message to the outgoing data buffer.
'
'   @param    upwards True if the item will be moved up in the list, False if it will be moved downwards.
'   @param    slot Bank List slot where the item which's info is requested is.
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteMoveBank(ByVal upwards As Boolean, ByVal Slot As Byte)
'***************************************************
'Author: Torres Patricio (Pato)
'Last Modification: 06/14/09
'Writes the "MoveBank" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.MoveBank)
        
        Call .WriteBoolean(upwards)
        Call .WriteByte(Slot)
    End With
End Sub

Public Sub WriteMoverInventario(ByVal act As Byte, Nue As Byte)
'***************************************************
'Author: Standelf
'Last Modification: 22/05/2010
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.MoverItemEnInventario)
        Call .WriteByte(act)
        Call .WriteByte(Nue)
    End With
End Sub

''
'   Writes the "MoveSpell" message to the outgoing data buffer.
'
'   @param    upwards True if the spell will be moved up in the list, False if it will be moved downwards.
'   @param    slot Spell List slot where the spell which's info is requested is.
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteMoveSpell(ByVal upwards As Boolean, ByVal Slot As Byte)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "MoveSpell" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.MoveSpell)
        
        Call .WriteBoolean(upwards)
        Call .WriteByte(Slot)
    End With
End Sub

''
'   Writes the "NavigateToggle" message to the outgoing data buffer.
'
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteNavigateToggle()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "NavigateToggle" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.NavigateToggle)
End Sub

''
'   Writes the "NickToIP" message to the outgoing data buffer.
'
'   @param    username The user whose IP is requested.
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteNickToIP(ByVal UserName As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "NickToIP" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.nickToIP)
        
        Call .WriteASCIIString(UserName)
    End With
End Sub

''
'   Writes the "Night" message to the outgoing data buffer.
'
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteNight()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Night" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.night)
End Sub

''
'   Writes the "NPCFollow" message to the outgoing data buffer.
'
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteNPCFollow()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "NPCFollow" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.NPCFollow)
End Sub

Public Sub WriteOfertarSubasta(ByVal Oferta As Long)
'***************************************************
'Author: Standelf
'Last Modification: 25/05/10
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.OfertarSubasta)
        Call .WriteLong(Oferta)
    End With
End Sub

''
'   Writes the "Online" message to the outgoing data buffer.
'
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteOnline()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Online" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.Online)
End Sub

''
'   Writes the "OnlineChaosLegion" message to the outgoing data buffer.
'
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteOnlineChaosLegion()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "OnlineChaosLegion" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.OnlineChaosLegion)
End Sub

''
'   Writes the "OnlineGM" message to the outgoing data buffer.
'
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteOnlineGM()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "OnlineGM" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.OnlineGM)
End Sub

''
'   Writes the "OnlineMap" message to the outgoing data buffer.
'
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteOnlineMap(ByVal Map As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 26/03/2009
'Writes the "OnlineMap" message to the outgoing data buffer
'26/03/2009: Now you don't need to be in the map to use the comand, so you send the map to server
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.OnlineMap)
        
        Call .WriteInteger(Map)
    End With
End Sub

''
'   Writes the "OnlineRoyalArmy" message to the outgoing data buffer.
'
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteOnlineRoyalArmy()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "OnlineRoyalArmy" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.OnlineRoyalArmy)
End Sub

''
'   Writes the "PartyAcceptMember" message to the outgoing data buffer.
'
'   @param    username The user to accept into the party.
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePartyAcceptMember(ByVal UserName As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "PartyAcceptMember" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.PartyAcceptMember)
        
        Call .WriteASCIIString(UserName)
    End With
End Sub

''
'   Writes the "PartyCreate" message to the outgoing data buffer.
'
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePartyCreate()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "PartyCreate" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.PartyCreate)
End Sub

''
'   Writes the "PartyJoin" message to the outgoing data buffer.
'
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePartyJoin()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "PartyJoin" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.PartyJoin)
End Sub

''
'   Writes the "PartyKick" message to the outgoing data buffer.
'
'   @param    username The user to kick fro mthe party.
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePartyKick(ByVal UserName As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "PartyKick" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.PartyKick)
            
        Call .WriteASCIIString(UserName)
    End With
End Sub

''
'   Writes the "PartyLeave" message to the outgoing data buffer.
'
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePartyLeave()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "PartyLeave" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.PartyLeave)
End Sub

''
'   Writes the "PartyMessage" message to the outgoing data buffer.
'
'   @param    message The message to send to the party.
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePartyMessage(ByVal Message As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "PartyMessage" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.PartyMessage)
        
        Call .WriteASCIIString(Message)
    End With
End Sub

''
'   Writes the "PartyOnline" message to the outgoing data buffer.
'
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePartyOnline()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "PartyOnline" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.PartyOnline)
End Sub

''
'   Writes the "PartySetLeader" message to the outgoing data buffer.
'
'   @param    username The user to set as the party's leader.
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePartySetLeader(ByVal UserName As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "PartySetLeader" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.PartySetLeader)
        
        Call .WriteASCIIString(UserName)
    End With
End Sub

''
'   Writes the "PetFollow" message to the outgoing data buffer.
'
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePetFollow()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "PetFollow" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.PetFollow)
End Sub

''
'   Writes the "PetStand" message to the outgoing data buffer.
'
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePetStand()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "PetStand" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.PetStand)
End Sub

''
'   Writes the "PickUp" message to the outgoing data buffer.
'
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePickUp()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "PickUp" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.PickUp)
End Sub

''
'   Writes the "Ping" message to the outgoing data buffer.
'
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePing()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 26/01/2007
'Writes the "Ping" message to the outgoing data buffer
'***************************************************
    'Prevent the timer from being cut
    If pingTime <> 0 Then Exit Sub
    
    Call outgoingData.WriteByte(ClientPacketID.Ping)
    
    '   Avoid computing errors due to frame rate
    Call FlushBuffer
    DoEvents
    
    pingTime = GetTickCount
End Sub

''
'   Writes the "Drop" message to the outgoing data buffer.
'
'   @param    slot Inventory slot where the item to drop is.
'   @param    amount Number of items to drop.
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePrestigio(ByVal UserName As String, ByVal Amount As Integer, ByVal Mode As Byte)
'***************************************************
'Author: Standelf
'Last Modification: 24/12/2010
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.Prestigio)
        
        Select Case Mode
            Case 0 'Pedir Prestigio
                Call .WriteByte(Mode)
                
            Case 1, 2 'Otorgar Prestigio Rep/Can -Otorgar Prestigio Can
                Call .WriteByte(Mode)
                Call .WriteASCIIString(UserName)
                Call .WriteInteger(Amount)
        End Select
    End With
End Sub

Public Sub WritePromedio()
'***************************************************
'Author: Standelf
'Last Modification: 22/05/10
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.Dame_Promedio)
    End With
End Sub

''
'   Writes the "Punishments" message to the outgoing data buffer.
'
'   @param    username The user whose's  punishments are requested.
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePunishments(ByVal UserName As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Punishments" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.Punishments)
        
        Call .WriteASCIIString(UserName)
    End With
End Sub

Public Sub WriteQuest()
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'Escribe el paquete Quest al servidor.
'Last modified: 31/01/2010 by Amraphen
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    Call outgoingData.WriteByte(ClientPacketID.Quest)
End Sub

Public Sub WriteQuestAbandon(ByVal QuestSlot As Byte)
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'Escribe el paquete QuestAbandon al servidor.
'Last modified: 31/01/2010 by Amraphen
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    'Escribe el ID del paquete.
    Call outgoingData.WriteByte(ClientPacketID.QuestAbandon)
    
    'Escribe el Slot de Quest.
    Call outgoingData.WriteByte(QuestSlot)
End Sub

Public Sub WriteQuestAccept()
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'Escribe el paquete QuestAccept al servidor.
'Last modified: 31/01/2010 by Amraphen
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    Call outgoingData.WriteByte(ClientPacketID.QuestAccept)
End Sub

Public Sub WriteQuestDetailsRequest(ByVal QuestSlot As Byte)
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'Escribe el paquete QuestDetailsRequest al servidor.
'Last modified: 31/01/2010 by Amraphen
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    Call outgoingData.WriteByte(ClientPacketID.QuestDetailsRequest)
    
    Call outgoingData.WriteByte(QuestSlot)
End Sub

Public Sub WriteQuestListRequest()
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'Escribe el paquete QuestListRequest al servidor.
'Last modified: 31/01/2010 by Amraphen
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    Call outgoingData.WriteByte(ClientPacketID.QuestListRequest)
End Sub

''
'   Writes the "Quit" message to the outgoing data buffer.
'
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteQuit()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 08/16/08
'Writes the "Quit" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.Quit)
End Sub

''
'   Writes the "RainToggle" message to the outgoing data buffer.
'
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRainToggle()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "RainToggle" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.RainToggle)
End Sub

''
'   Writes the "ReleasePet" message to the outgoing data buffer.
'
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteReleasePet()
'***************************************************
'Author: ZaMa
'Last Modification: 18/11/2009
'Writes the "ReleasePet" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.ReleasePet)
End Sub

''
'   Writes the "ReloadNPCs" message to the outgoing data buffer.
'
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteReloadNPCs()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ReloadNPCs" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.ReloadNPCs)
End Sub

''
'   Writes the "ReloadObjects" message to the outgoing data buffer.
'
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteReloadObjects()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ReloadObjects" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.ReloadObjects)
End Sub

''
'   Writes the "ReloadServerIni" message to the outgoing data buffer.
'
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteReloadServerIni()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ReloadServerIni" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.ReloadServerIni)
End Sub

''
'   Writes the "ReloadSpells" message to the outgoing data buffer.
'
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteReloadSpells()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ReloadSpells" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.ReloadSpells)
End Sub

''
'   Writes the "RemoveCharFromGuild" message to the outgoing data buffer.
'
'   @param    username The name of the user who will be removed from any guild.
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRemoveCharFromGuild(ByVal UserName As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "RemoveCharFromGuild" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.RemoveCharFromGuild)
        
        Call .WriteASCIIString(UserName)
    End With
End Sub

''
'   Writes the "RemovePunishment" message to the outgoing data buffer.
'
'   @param    username The user whose punishments will be altered.
'   @param    punishment The id of the punishment to be removed.
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRemovePunishment(ByVal UserName As String, ByVal punishment As Byte, ByVal NewText As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "RemovePunishment" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.RemovePunishment)
        
        Call .WriteASCIIString(UserName)
        Call .WriteByte(punishment)
        Call .WriteASCIIString(NewText)
    End With
End Sub

''
'   Writes the "RequestAccountState" message to the outgoing data buffer.
'
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRequestAccountState()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "RequestAccountState" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.RequestAccountState)
End Sub

''
'   Writes the "RequestAtributes" message to the outgoing data buffer.
'
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRequestAtributes()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "RequestAtributes" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.RequestAtributes)
End Sub

''
'   Writes the "RequestCharBank" message to the outgoing data buffer.
'
'   @param    username The user whose banking information is requested.
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRequestCharBank(ByVal UserName As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "RequestCharBank" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.RequestCharBank)
        
        Call .WriteASCIIString(UserName)
    End With
End Sub

''
'   Writes the "RequestCharGold" message to the outgoing data buffer.
'
'   @param    username The user whose gold is requested.
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRequestCharGold(ByVal UserName As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "RequestCharGold" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.RequestCharGold)
        
        Call .WriteASCIIString(UserName)
    End With
End Sub

''
'   Writes the "RequestCharInfo" message to the outgoing data buffer.
'
'   @param    username The user whose information is requested.
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRequestCharInfo(ByVal UserName As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "RequestCharInfo" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.RequestCharInfo)
        
        Call .WriteASCIIString(UserName)
    End With
End Sub

    
''
'   Writes the "RequestCharInventory" message to the outgoing data buffer.
'
'   @param    username The user whose inventory is requested.
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRequestCharInventory(ByVal UserName As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "RequestCharInventory" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.RequestCharInventory)
        
        Call .WriteASCIIString(UserName)
    End With
End Sub

''
'   Writes the "RequestCharMail" message to the outgoing data buffer.
'
'   @param    username The name of the user whose mail is requested.
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRequestCharMail(ByVal UserName As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "RequestCharMail" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.RequestCharMail)
        
        Call .WriteASCIIString(UserName)
    End With
End Sub

''
'   Writes the "RequestCharSkills" message to the outgoing data buffer.
'
'   @param    username The user whose skills are requested.
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRequestCharSkills(ByVal UserName As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "RequestCharSkills" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.RequestCharSkills)
        
        Call .WriteASCIIString(UserName)
    End With
End Sub

''
'   Writes the "RequestCharStats" message to the outgoing data buffer.
'
'   @param    username The user whose stats are requested.
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRequestCharStats(ByVal UserName As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "RequestCharStats" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.RequestCharStats)
        
        Call .WriteASCIIString(UserName)
    End With
End Sub

''
'   Writes the "RequestFame" message to the outgoing data buffer.
'
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRequestFame()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "RequestFame" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.RequestFame)
End Sub

''
'   Writes the "RequestGuildLeaderInfo" message to the outgoing data buffer.
'
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRequestGuildLeaderInfo()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "RequestGuildLeaderInfo" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.RequestGuildLeaderInfo)
End Sub

''
'   Writes the "RequestMiniStats" message to the outgoing data buffer.
'
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRequestMiniStats()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "RequestMiniStats" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.RequestMiniStats)
End Sub

''
'   Writes the "RequestMOTD" message to the outgoing data buffer.
'
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRequestMOTD()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "RequestMOTD" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.RequestMOTD)
End Sub

Public Sub WriteRequestPartyForm()
'***************************************************
'Author: Budi
'Last Modification: 11/26/09
'Writes the "RequestPartyForm" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.RequestPartyForm)

End Sub

''
'   Writes the "RequestPositionUpdate" message to the outgoing data buffer.
'
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRequestPositionUpdate()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "RequestPositionUpdate" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.RequestPositionUpdate)
End Sub

''
'   Writes the "RequestSkills" message to the outgoing data buffer.
'
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRequestSkills()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "RequestSkills" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.RequestSkills)
End Sub

''
'   Writes the "RequestStats" message to the outgoing data buffer.
'
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRequestStats()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "RequestStats" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.RequestStats)
End Sub

''
'   Writes the "RequestUserList" message to the outgoing data buffer.
'
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRequestUserList()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "RequestUserList" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.RequestUserList)
End Sub

''
'   Writes the "ResetAutoUpdate" message to the outgoing data buffer.
'
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteResetAutoUpdate()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ResetAutoUpdate" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.ResetAutoUpdate)
End Sub

''
'   Writes the "ResetFactions" message to the outgoing data buffer.
'
'   @param    username The name of the user who will be removed from any faction.
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteResetFactions(ByVal UserName As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ResetFactions" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.ResetFactions)
        
        Call .WriteASCIIString(UserName)
    End With
End Sub

''
'   Writes the "ResetNPCInventory" message to the outgoing data buffer.
'
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteResetNPCInventory()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ResetNPCInventory" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.ResetNPCInventory)
End Sub

''
'   Writes the "Rest" message to the outgoing data buffer.
'
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRest()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Rest" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.Rest)
End Sub

''
'   Writes the "Restart" message to the outgoing data buffer.
'
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRestart()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Restart" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.Restart)
End Sub

''
'   Writes the "Resucitate" message to the outgoing data buffer.
'
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteResucitate()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Resucitate" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.Resucitate)
End Sub

''
'   Writes the "ResuscitationSafeToggle" message to the outgoing data buffer.
'
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteResuscitationToggle()
'**************************************************************
'Author: Rapsodius
'Creation Date: 10/10/07
'Writes the Resuscitation safe toggle packet to the outgoing data buffer.
'**************************************************************
    Call outgoingData.WriteByte(ClientPacketID.ResuscitationSafeToggle)
End Sub

''
'   Writes the "ReviveChar" message to the outgoing data buffer.
'
'   @param    username The user to eb revived.
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteReviveChar(ByVal UserName As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ReviveChar" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.ReviveChar)
        
        Call .WriteASCIIString(UserName)
    End With
End Sub

''
'   Writes the "Reward" message to the outgoing data buffer.
'
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteReward()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Reward" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.Reward)
End Sub

''
'   Writes the "RoleMasterRequest" message to the outgoing data buffer.
'
'   @param    message The message to send to the role masters.
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRoleMasterRequest(ByVal Message As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "RoleMasterRequest" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.RoleMasterRequest)
        
        Call .WriteASCIIString(Message)
    End With
End Sub

''
'   Writes the "RoyalArmyKick" message to the outgoing data buffer.
'
'   @param    username The name of the user to be kicked from the Royal Army.
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRoyalArmyKick(ByVal UserName As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "RoyalArmyKick" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.RoyalArmyKick)
        
        Call .WriteASCIIString(UserName)
    End With
End Sub

''
'   Writes the "RoyalArmyMessage" message to the outgoing data buffer.
'
'   @param    message The message to send to the royal army members.
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRoyalArmyMessage(ByVal Message As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "RoyalArmyMessage" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.RoyalArmyMessage)
        
        Call .WriteASCIIString(Message)
    End With
End Sub

''
'   Writes the "SafeToggle" message to the outgoing data buffer.
'
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSafeToggle()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "SafeToggle" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.SafeToggle)
End Sub

Public Sub WriteSaveCharacter()
'***************************************************
'Author: Standelf
'Last Modification: 22/05/10
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.Save_Character)
    End With
End Sub

''
'   Writes the "SaveChars" message to the outgoing data buffer.
'
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSaveChars()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "SaveChars" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.SaveChars)
End Sub

''
'   Writes the "SaveMap" message to the outgoing data buffer.
'
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSaveMap()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "SaveMap" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.SaveMap)
End Sub

''
'   Writes the "ServerMessage" message to the outgoing data buffer.
'
'   @param    message The message to be sent to players.
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteServerMessage(ByVal Message As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ServerMessage" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.ServerMessage)
        
        Call .WriteASCIIString(Message)
    End With
End Sub

''
'   Writes the "ServerOpenToUsersToggle" message to the outgoing data buffer.
'
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteServerOpenToUsersToggle()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ServerOpenToUsersToggle" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.ServerOpenToUsersToggle)
End Sub

''
'   Writes the "ServerTime" message to the outgoing data buffer.
'
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteServerTime()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ServerTime" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.serverTime)
End Sub

''
'   Writes the "SetCharDescription" message to the outgoing data buffer.
'
'   @param    desc The description to set to players.
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSetCharDescription(ByVal Desc As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "SetCharDescription" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.SetCharDescription)
        
        Call .WriteASCIIString(Desc)
    End With
End Sub

''
'   Writes the "SetIniVar" message to the outgoing data buffer.
'
'   @param    sLlave the name of the key which contains the value to edit
'   @param    sClave the name of the value to edit
'   @param    sValor the new value to set to sClave
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSetIniVar(ByRef sLlave As String, ByRef sClave As String, ByRef sValor As String)
'***************************************************
'Author: Brian Chaia (BrianPr)
'Last Modification: 21/06/2009
'Writes the "SetIniVar" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.SetIniVar)
        
        Call .WriteASCIIString(sLlave)
        Call .WriteASCIIString(sClave)
        Call .WriteASCIIString(sValor)
    End With
End Sub

''
'   Writes the "SetMOTD" message to the outgoing data buffer.
'
'   @param    message The message to be set as the new MOTD.
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSetMOTD(ByVal Message As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "SetMOTD" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.SetMOTD)
        
        Call .WriteASCIIString(Message)
    End With
End Sub

''
'   Writes the "SetTrigger" message to the outgoing data buffer.
'
'   @param    trigger The type of trigger to be set to the tile.
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSetTrigger(ByVal Trigger As eTrigger)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "SetTrigger" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.SetTrigger)
        
        Call .WriteByte(Trigger)
    End With
End Sub

''
'   Writes the "ShareNpc" message to the outgoing data buffer.
'
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteShareNpc()
'***************************************************
'Author: ZaMa
'Last Modification: 15/04/2010
'Writes the "ShareNpc" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.ShareNpc)
End Sub

''
'   Writes the "ShowGuildMessages" message to the outgoing data buffer.
'
'   @param    guild The guild to listen to.
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteShowGuildMessages(ByVal guild As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ShowGuildMessages" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.ShowGuildMessages)
        
        Call .WriteASCIIString(guild)
    End With
End Sub

''
'   Writes the "ShowGuildNews" message to the outgoing data buffer.
'

Public Sub WriteShowGuildNews()
'***************************************************
'Author: ZaMa
'Last Modification: 21/02/2010
'Writes the "ShowGuildNews" message to the outgoing data buffer
'***************************************************
 
     outgoingData.WriteByte (ClientPacketID.ShowGuildNews)
End Sub

''
'   Writes the "ShowName" message to the outgoing data buffer.
'
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteShowName()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ShowName" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.showName)
End Sub

''
'   Writes the "ShowServerForm" message to the outgoing data buffer.
'
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteShowServerForm()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ShowServerForm" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.ShowServerForm)
End Sub

''
'   Writes the "Silence" message to the outgoing data buffer.
'
'   @param    username The user to silence.
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSilence(ByVal UserName As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Silence" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.Silence)
        
        Call .WriteASCIIString(UserName)
    End With
End Sub

Public Sub WriteSimple(ByVal ID As Byte)
'***************************************************
'Author: Standelf
'Last Modification: 16/05/10
'WATA!
'***************************************************
    With outgoingData
        Call .WriteByte(ID)
    End With
End Sub

''
'   Writes the "SOSRemove" message to the outgoing data buffer.
'
'   @param    username The user whose SOS call has been already attended.
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSOSRemove(ByVal UserName As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "SOSRemove" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.SOSRemove)
        
        Call .WriteASCIIString(UserName)
    End With
End Sub

''
'   Writes the "SOSShowList" message to the outgoing data buffer.
'
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSOSShowList()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "SOSShowList" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.SOSShowList)
End Sub

''
'   Writes the "SpawnCreature" message to the outgoing data buffer.
'
'   @param    creatureIndex The index of the creature in the spawn list to be spawned.
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSpawnCreature(ByVal creatureIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "SpawnCreature" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.SpawnCreature)
        
        Call .WriteInteger(creatureIndex)
    End With
End Sub

''
'   Writes the "SpawnListRequest" message to the outgoing data buffer.
'
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSpawnListRequest()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "SpawnListRequest" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.SpawnListRequest)
End Sub

''
'   Writes the "SpellInfo" message to the outgoing data buffer.
'
'   @param    slot Spell List slot where the spell which's info is requested is.
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSpellInfo(ByVal Slot As Byte)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "SpellInfo" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.SpellInfo)
        
        Call .WriteByte(Slot)
    End With
End Sub

''
'   Writes the "StopSharingNpc" message to the outgoing data buffer.
'
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteStopSharingNpc()
'***************************************************
'Author: ZaMa
'Last Modification: 15/04/2010
'Writes the "StopSharingNpc" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.StopSharingNpc)
End Sub

''
'   Writes the "SummonChar" message to the outgoing data buffer.
'
'   @param    username The user to be summoned.
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSummonChar(ByVal UserName As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "SummonChar" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.SummonChar)
        
        Call .WriteASCIIString(UserName)
    End With
End Sub

''
'   Writes the "SystemMessage" message to the outgoing data buffer.
'
'   @param    message The message to be sent to all players.
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSystemMessage(ByVal Message As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "SystemMessage" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.SystemMessage)
        
        Call .WriteASCIIString(Message)
    End With
End Sub

''
'   Writes the "Talk" message to the outgoing data buffer.
'
'   @param    chat The chat text to be sent.
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteTalk(ByVal chat As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Talk" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.Talk)
        
        Call .WriteASCIIString(chat)
    End With
End Sub

''
'   Writes the "TalkAsNPC" message to the outgoing data buffer.
'
'   @param    message The message to send to the royal army members.
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteTalkAsNPC(ByVal Message As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "TalkAsNPC" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.TalkAsNPC)
        
        Call .WriteASCIIString(Message)
    End With
End Sub

''
'   Writes the "TeleportCreate" message to the outgoing data buffer.
'
'   @param    map the map to which the teleport will lead.
'   @param    x The position in the x axis to which the teleport will lead.
'   @param    y The position in the y axis to which the teleport will lead.
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteTeleportCreate(ByVal Map As Integer, ByVal X As Byte, ByVal Y As Byte, Optional ByVal Radio As Byte = 0)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "TeleportCreate" message to the outgoing data buffer
'***************************************************
    With outgoingData
            Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.TeleportCreate)
        
        Call .WriteInteger(Map)
        
        Call .WriteByte(X)
        Call .WriteByte(Y)
        
        Call .WriteByte(Radio)
    End With
End Sub

''
'   Writes the "TeleportDestroy" message to the outgoing data buffer.
'
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteTeleportDestroy()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "TeleportDestroy" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.TeleportDestroy)
End Sub

''
'   Writes the "ThrowDices" message to the outgoing data buffer.
'
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteThrowDices()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ThrowDices" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.ThrowDices)
End Sub

''
'   Writes the "TileBlockedToggle" message to the outgoing data buffer.
'
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteTileBlockedToggle()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "TileBlockedToggle" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.TileBlockedToggle)
End Sub

''
'   Writes the "ToggleCentinelActivated" message to the outgoing data buffer.
'
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteToggleCentinelActivated()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ToggleCentinelActivated" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.ToggleCentinelActivated)
End Sub

''
'   Writes the "Train" message to the outgoing data buffer.
'
'   @param    creature Position within the list provided by the server of the creature to train against.
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteTrain(ByVal creature As Byte)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Train" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.Train)
        
        Call .WriteByte(creature)
    End With
End Sub

''
'   Writes the "TrainList" message to the outgoing data buffer.
'
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteTrainList()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "TrainList" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.TrainList)
End Sub

''
'   Writes the "TurnCriminal" message to the outgoing data buffer.
'
'   @param    username The name of the user to turn into criminal.
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteTurnCriminal(ByVal UserName As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "TurnCriminal" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.TurnCriminal)
        
        Call .WriteASCIIString(UserName)
    End With
End Sub

''
'   Writes the "TurnOffServer" message to the outgoing data buffer.
'
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteTurnOffServer()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "TurnOffServer" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.TurnOffServer)
End Sub

''
'   Writes the "UnbanChar" message to the outgoing data buffer.
'
'   @param    username The user to be unbanned.
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUnbanChar(ByVal UserName As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "UnbanChar" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.UnbanChar)
        
        Call .WriteASCIIString(UserName)
    End With
End Sub

''
'   Writes the "UnbanIP" message to the outgoing data buffer.
'
'   @param    IP The IP for which to search for players. Must be an array of 4 elements with the 4 components of the IP.
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUnbanIP(ByRef Ip() As Byte)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "UnbanIP" message to the outgoing data buffer
'***************************************************
    If UBound(Ip()) - LBound(Ip()) + 1 <> 4 Then Exit Sub   'Invalid IP
    
    Dim i As Long
    
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.UnbanIP)
        
        For i = LBound(Ip()) To UBound(Ip())
            Call .WriteByte(Ip(i))
        Next i
    End With
End Sub

''
'   Writes the "UpTime" message to the outgoing data buffer.
'
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUpTime()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "UpTime" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.Uptime)
End Sub

''
'   Writes the "UseItem" message to the outgoing data buffer.
'
'   @param    slot Invetory slot where the item to use is.
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUseItem(ByVal Slot As Byte)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "UseItem" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.UseItem)
        
        Call .WriteByte(Slot)
    End With
End Sub

''
'   Writes the "UserCommerceConfirm" message to the outgoing data buffer.
'
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUserCommerceConfirm()
'***************************************************
'Author: ZaMa
'Last Modification: 14/12/2009
'Writes the "UserCommerceConfirm" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.UserCommerceConfirm)
End Sub

''
'   Writes the "UserCommerceEnd" message to the outgoing data buffer.
'
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUserCommerceEnd()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "UserCommerceEnd" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.UserCommerceEnd)
End Sub

''
'   Writes the "UserCommerceOffer" message to the outgoing data buffer.
'
'   @param    slot Position within user inventory in which the desired item is.
'   @param    amount Number of items to offer.
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUserCommerceOffer(ByVal Slot As Byte, ByVal Amount As Long, ByVal OfferSlot As Byte)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "UserCommerceOffer" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.UserCommerceOffer)
        
        Call .WriteByte(Slot)
        Call .WriteLong(Amount)
        Call .WriteByte(OfferSlot)
    End With
End Sub

''
'   Writes the "UserCommerceOk" message to the outgoing data buffer.
'
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUserCommerceOk()
'***************************************************
'Author: Fredy Horacio Treboux (liquid)
'Last Modification: 01/10/07
'Writes the "UserCommerceOk" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.UserCommerceOk)
End Sub

''
'   Writes the "UserCommerceReject" message to the outgoing data buffer.
'
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUserCommerceReject()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "UserCommerceReject" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.UserCommerceReject)
End Sub

''
'   Writes the "UseSpellMacro" message to the outgoing data buffer.
'
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUseSpellMacro()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "UseSpellMacro" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.UseSpellMacro)
End Sub

''
'   Writes the "Walk" message to the outgoing data buffer.
'
'   @param    heading The direction in wich the user is moving.
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteWalk(ByVal Heading As E_Heading)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Walk" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.Walk)
        
        Call .WriteByte(Heading)
    End With
End Sub

''
'   Writes the "WarnUser" message to the outgoing data buffer.
'
'   @param    username The user to be warned.
'   @param    reason Reason for the warning.
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteWarnUser(ByVal UserName As String, ByVal reason As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "WarnUser" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.WarnUser)
        
        Call .WriteASCIIString(UserName)
        Call .WriteASCIIString(reason)
    End With
End Sub

''
'   Writes the "WarpChar" message to the outgoing data buffer.
'
'   @param    username The user to be warped. "YO" represent's the user's char.
'   @param    map The map to which to warp the character.
'   @param    x The x position in the map to which to waro the character.
'   @param    y The y position in the map to which to waro the character.
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteWarpChar(ByVal UserName As String, ByVal Map As Integer, ByVal X As Byte, ByVal Y As Byte)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "WarpChar" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.WarpChar)
        
        Call .WriteASCIIString(UserName)
        
        Call .WriteInteger(Map)
        
        Call .WriteByte(X)
        Call .WriteByte(Y)
    End With
End Sub

''
'   Writes the "WarpMeToTarget" message to the outgoing data buffer.
'
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteWarpMeToTarget()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "WarpMeToTarget" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.WarpMeToTarget)
End Sub

''
'   Writes the "Where" message to the outgoing data buffer.
'
'   @param    username The user whose position is requested.
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteWhere(ByVal UserName As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Where" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.Where)
        
        Call .WriteASCIIString(UserName)
    End With
End Sub

''
'   Writes the "Whisper" message to the outgoing data buffer.
'
'   @param    charIndex The index of the char to whom to whisper.
'   @param    chat The chat text to be sent to the user.
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteWhisper(ByVal CharIndex As Integer, ByVal chat As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Whisper" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.Whisper)
        
        Call .WriteInteger(CharIndex)
        
        Call .WriteASCIIString(chat)
    End With
End Sub

''
'   Writes the "Work" message to the outgoing data buffer.
'
'   @param    skill The skill which the user attempts to use.
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteWork(ByVal Skill As eSkill)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Work" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.Work)
        
        Call .WriteByte(Skill)
    End With
End Sub

''
'   Writes the "Working" message to the outgoing data buffer.
'
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteWorking()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Working" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.Working)
End Sub

''
'   Writes the "WorkLeftClick" message to the outgoing data buffer.
'
'   @param    x Tile coord in the x-axis in which the user clicked.
'   @param    y Tile coord in the y-axis in which the user clicked.
'   @param    skill The skill which the user attempts to use.
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteWorkLeftClick(ByVal X As Byte, ByVal Y As Byte, ByVal Skill As eSkill)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "WorkLeftClick" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.WorkLeftClick)
        
        Call .WriteByte(X)
        Call .WriteByte(Y)
        
        Call .WriteByte(Skill)
    End With
End Sub

''
'   Writes the "Yell" message to the outgoing data buffer.
'
'   @param    chat The chat text to be sent.
'   @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteYell(ByVal chat As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Yell" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.Yell)
        
        Call .WriteASCIIString(chat)
    End With
End Sub


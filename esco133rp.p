&ANALYZE-SUSPEND _VERSION-NUMBER UIB_v8r12
&ANALYZE-RESUME
&ANALYZE-SUSPEND _UIB-CODE-BLOCK _CUSTOM _DEFINITIONS Procedure 
/*

    Programa: esp/esco133rp.p
   Descricao: Relat¢rio de Öndice de Aproveitamento
       Autor: Datasul Bandeirantes - Unidade Campinas
        Data: Janeiro/2004

*/

/* includes */
{include/i-prgvrs.i ESCO133RP 1.00.00.001}

&IF "{&EMSFND_VERSION}" >= "1.00" &THEN
    {include/i-license-manager.i ESCO133RP MCC}
&ENDIF



{utp/ut-glob.i}

/* ***************************  Definitions  ************************** */
&global-define programa ESCO133RP

/* definicao de variaveis locais */
Def var c-liter-par                     As Char Format "x(13)":U.
Def var c-liter-sel                     As Char Format "x(10)":U.
Def var c-liter-imp                     As Char Format "x(12)":U.    
Def Var c-destino                       As Char Format "x(15)":U.
Def Var h-acomp                         As Handle                   No-undo.
Def Var i-cont-ord                      As Int.
Def Var l-branco As Logical.
/* definicao de temp-tables */
Def Temp-table tt-param No-undo
    Field destino           As Int
    Field arquivo           As Char                         Format "x(35)"
    Field usuario           As Char                         Format "x(12)"
    Field data-exec         As Date
    Field hora-exec         As Int
    Field classifica        As Int
    Field desc-classifica   As Char                         Format "x(40)"
    Field cod-est-ini       Like pedido-compr.cod-estabel
    Field cod-est-fim       Like pedido-compr.cod-estabel
    Field num-pedido-ini    Like pedido-compr.num-pedido
    Field num-pedido-fim    Like pedido-compr.num-pedido
    Field ge-codigo-ini     Like item.ge-codigo
    Field ge-codigo-fim     Like item.ge-codigo
    Field data-pedido-ini   Like pedido-compr.data-pedido
    Field data-pedido-fim   Like pedido-compr.data-pedido
    Field comprador-ini     Like ordem-compra.cod-comprado
    Field comprador-fim     Like ordem-compra.cod-comprado
    FIELD gerar-excel       AS LOGICAL.

Def Temp-table tt-digita
    Field ordem                         As Int  Format ">>>>9":U
    Field exemplo                       As Char Format "x(30)":U
    Index id Is Primary Unique
        ordem.

Def Temp-table tt-raw-digita
    Field raw-digita                    As Raw.

Def Temp-table tt-pedido
    FIELD cod-estabel       LIKE estabelec.cod-estabel
    Field num-pedido        Like pedido-compr.num-pedido
    Field data-pedido       Like pedido-compr.data-pedido
    Field numero-ordem      Like ordem-compra.numero-ordem
    Field it-codigo         Like item.it-codigo
    Field desc-item         Like item.desc-item
    Field qt-solic          Like ordem-compra.qt-solic
    Field preco-ordem       Like ordem-compra.preco-fornec                      Column-label "Valor Pago"
    Field preco-menor       Like ordem-compra.preco-fornec                      Column-label "Valor Menor"
    Field cod-emitente      Like emitente.cod-emitente
    Field nome-abrev        Like emitente.nome-abrev
    Field narrativa         As Char   Format "x(75)"  Extent 25 Column-label "Narrativa"
    Field narrativa-cot     As Char   Format "x(75)" Extent 25 Column-label "Narrativa Cota‡Æo"
    Field cod-comprador     Like ordem-compra.cod-comprado
    Field l-branco          As Logical
    FIELD cod-justif        AS INTEGER
    FIELD desc-justif       AS CHARACTER
    FIELD nr-processo       AS INTEGER
    Index ch-pedido Is Primary
        num-pedido
        numero-ordem.

Def Temp-table tt-ord-aux
    Field num-pedido        Like ordem-compra.num-pedido    
    Field numero-ordem      Like ordem-compra.numero-ordem
    Field maior             As Logical
    Index ch-ordem Is Primary Unique
        num-pedido
        numero-ordem.
        
define temp-table tt-cotacao-item no-undo like cotacao-item
       field preco-calc as decimal.        
 
/* definicao de parametros */
Def Input Param raw-param               As Raw No-undo.
Def Input Param Table For tt-raw-digita.

Def Buffer b-cotacao-item For cotacao-item.


DEF VAR c-descr AS CHAR EXTENT 8 INITIAL ["Marca Homologada",
                                          "Fechado por Pacote",
                                          "Frete por conta do Fornecedor",
                                          "Melhor Prazo de Entrega ",
                                          "Conforme Especifica‡Æo T‚cnica",
                                          "Valor Cotado Errado pelo Melhor Fornecedor",
                                          "Difal (Diferencial de Al¡quota de ICMS)",
                                          "Melhor Pre‡o conforme Unidade de Medida"] NO-UNDO.



Create tt-param.
Raw-transfer raw-param To tt-param.

For Each tt-raw-digita:

    Create tt-digita.
    Raw-transfer tt-raw-digita.raw-digita To tt-digita.

End.

{include/i-rpvar.i}

Find mgcad.empresa Where empresa.ep-codigo = v_cdn_empres_usuar No-lock No-error.

Find First param-global No-lock No-error.

{utp/ut-liter.i Compras * }
Assign c-sistema = Return-value.

{utp/ut-liter.i Relat¢rio_de_Öndice_de_Aproveitamento * }
Assign c-titulo-relat = Return-value.

Assign c-empresa     = param-global.grupo
       c-programa    = "{&programa}":U
       c-versao      = "1.00":U
       c-revisao     = "000"
       c-destino     = {varinc/var00002.i 04 tt-param.destino}.

DEF VAR de-preco-sem-impostos   AS DECIMAL FORMAT ">>>>>,>>>,>>9.99999"  NO-UNDO.
DEFINE VARIABLE de-valor-icm    AS DECIMAL                               NO-UNDO.
DEFINE VARIABLE de-valor-pis    AS DECIMAL                               NO-UNDO.
DEFINE VARIABLE de-valor-cofins AS DECIMAL                               NO-UNDO.
DEFINE VARIABLE de-log-natural  AS DECIMAL FORMAT ">>>>,>>>,>>9.9999"    NO-UNDO.
DEFINE VARIABLE de-prazos       AS DECIMAL FORMAT "->>>>>,>>>,>>9.9999"  NO-UNDO.
DEFINE VARIABLE de-prazo-medio  AS DECIMAL FORMAT "->>>>>,>>>,>>9.9999"  NO-UNDO.
DEFINE VARIABLE vl-tx-fin-dia   AS DECIMAL FORMAT "->>>>>,>>>,>>9.9999"  NO-UNDO.
DEFINE VARIABLE de-preco-negoc  AS DECIMAL FORMAT "->>>>>,>>>,>>9.9999"  NO-UNDO.
DEFINE VARIABLE de-pre-unit-for AS DECIMAL FORMAT "->>>>>,>>>,>>9.99"    NO-UNDO.

/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME


&ANALYZE-SUSPEND _UIB-PREPROCESSOR-BLOCK 

/* ********************  Preprocessor Definitions  ******************** */

&Scoped-define PROCEDURE-TYPE Procedure
&Scoped-define DB-AWARE no



/* _UIB-PREPROCESSOR-BLOCK-END */
&ANALYZE-RESUME



/* *********************** Procedure Settings ************************ */

&ANALYZE-SUSPEND _PROCEDURE-SETTINGS
/* Settings for THIS-PROCEDURE
   Type: Procedure
   Allow: 
   Frames: 0
   Add Fields to: Neither
   Other Settings: CODE-ONLY COMPILE
 */
&ANALYZE-RESUME _END-PROCEDURE-SETTINGS

/* *************************  Create Window  ************************** */

&ANALYZE-SUSPEND _CREATE-WINDOW
/* DESIGN Window definition (used by the UIB) 
  CREATE WINDOW Procedure ASSIGN
         HEIGHT             = 5.88
         WIDTH              = 40.
/* END WINDOW DEFINITION */
                                                                        */
&ANALYZE-RESUME

&ANALYZE-SUSPEND _UIB-CODE-BLOCK _CUSTOM _INCLUDED-LIB Procedure 
/* ************************* Included-Libraries *********************** */

{include/tt-edit.i}
{include/pi-edit.i}
{include/i-rpcab.i}

/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME


 


&ANALYZE-SUSPEND _UIB-CODE-BLOCK _CUSTOM _MAIN-BLOCK Procedure 


/* ***************************  Main Block  *************************** */

Do On Stop Undo, Leave:

    {include/i-rpout.i}

    View Frame f-cabec.

    View Frame f-rodape.

    Run utp/ut-acomp.p Persistent Set h-acomp.
    
    Run pi-inicializar In h-acomp (Input "":U).
    
    Run pi-monta-tt.

    IF NOT tt-param.gerar-excel THEN 
       Run pi-display-notepad.
    ELSE 
       RUN pi-display-excel.
    
    Run pi-finalizar In h-acomp.

    {include/i-rpclo.i}

End.

/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME


/* **********************  Internal Procedures  *********************** */

&IF DEFINED(EXCLUDE-pi-display-excel) = 0 &THEN

&ANALYZE-SUSPEND _UIB-CODE-BLOCK _PROCEDURE pi-display-excel Procedure 
PROCEDURE pi-display-excel :
/*------------------------------------------------------------------------------
  Purpose:     
  Parameters:  <none>
  Notes:       
------------------------------------------------------------------------------*/

Def Var c-arquivo       As CHARACTER No-undo.
DEF VAR i-cont          AS INTEGER   NO-UNDO.
Def Var i-tot-ord       As Integer   no-undo.
Def Var i-tot-ord-maior As Integer   no-undo.


    
Assign c-arquivo = session:Temp-dir + "ESCO133_" + replace(String(Today,"99/99/9999"),"/","") + '_' + Replace(String(Time,"hh:mm:ss"),":","") + ".csv".

Output To Value(c-arquivo) No-map No-convert.

Export delimiter ';' 'Estabelecimento'  
                     'Comprador'        
                     'Pedido'   
                     'Data Pedido'      
                     'Cota‡Æo'  
                     'Ordem de Compra'  
                     'Cod.Fornecedor'   
                     'Nome Abreviado'   
                     'Item'     
                     'Descri‡Æo'        
                     'Qtde'     
                     'Valor Pago'       
                     'Valor Menor'
                     'Narrativa Ordem'
                     'Narrativa Cota‡Æo'
                     'C¢digo Justificativa'
                     'Descri‡Æo Justificativa'
                     'Narrativa Justificativa'.

ASSIGN i-cont = 0.

For Each tt-pedido:

    ASSIGN i-cont = i-cont + 1.

    Run pi-acompanhar In h-acomp (Input 'Listando : ' + String(i-cont)).
    
    find usuar-mater where usuar-mater.cod-usuario = tt-pedido.cod-comprador no-lock no-error.    
    
    EXPORT DELIMITER ';' "=" + quoter(tt-pedido.cod-estabel)
                         String(tt-pedido.cod-comprador) + ' - ' + String(usuar-mater.nome-usuar)
                         tt-pedido.num-pedido      
                         tt-pedido.data-pedido  
                         tt-pedido.nr-processo
                         tt-pedido.numero-ordem   
                         tt-pedido.cod-emitente    
                         tt-pedido.nome-abrev      
                         "=" + quoter(tt-pedido.it-codigo)       
                         tt-pedido.desc-item       
                         tt-pedido.qt-solic        
                         tt-pedido.preco-ordem     
                         tt-pedido.preco-menor     
                         tt-pedido.narrativa[1]    
                         tt-pedido.narrativa-cot[1]
                         tt-pedido.cod-justif      
                         tt-pedido.desc-justif.
END.

Assign i-tot-ord = 0.

For Each tt-ord-aux:

    Assign i-tot-ord = i-tot-ord + 1.

    If tt-ord-aux.maior Then
       Assign i-tot-ord-maior = i-tot-ord-maior + 1.

End.


Assign i-tot-ord-maior = 0.

For Each tt-pedido Where tt-pedido.l-branco = No :

    Assign i-tot-ord-maior = i-tot-ord-maior + 1.

End.

Put unformatted skip(2)
                "Total de Pedidos: "                ';' i-tot-ord                                                     Skip
                "Total de Pedidos c/ valor maior: " ';' i-tot-ord-maior                                               Skip
                "Percentual c/ valor maior: "       ';' ((i-tot-ord-maior * 100) / i-tot-ord)   Format ">>9.99" " %"  Skip.



OUTPUT CLOSE.

Run pi-excel(Input c-arquivo).

END PROCEDURE.

/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME

&ENDIF

&IF DEFINED(EXCLUDE-pi-display-notepad) = 0 &THEN

&ANALYZE-SUSPEND _UIB-CODE-BLOCK _PROCEDURE pi-display-notepad Procedure 
PROCEDURE pi-display-notepad :
/**/

    Def Var i-cont          As Int.
    Def Var i-tot-ord       As Int.
    Def Var i-tot-ord-maior As Int.

    For Each tt-pedido Where tt-pedido.l-branco = No Break 
             By If tt-param.classifica = 1 Then tt-pedido.it-codigo
                Else tt-pedido.desc-item
             By tt-pedido.cod-comprador:

       
       If First-of(tt-pedido.cod-comprador) Then Do:

           find usuar-mater where usuar-mater.cod-usuario = tt-pedido.cod-comprador no-lock no-error.

           Put 'Comprador..: ' + String(tt-pedido.cod-comprador) + ' - ' + 
                                 String(usuar-mater.nome-usuar) Format 'x(132)' At 01. 
           Put Skip.

       End.
    
        Disp tt-pedido.num-pedido
             tt-pedido.data-pedido
             tt-pedido.cod-emitente
             tt-pedido.nome-abrev
             tt-pedido.it-codigo
             tt-pedido.desc-item        Skip
             tt-pedido.qt-solic
             tt-pedido.preco-ordem
             tt-pedido.preco-menor
             tt-pedido.narrativa[1]       At 58 Skip(1)
             tt-pedido.narrativa-cot[1]   At 58 
            With Scrollable Stream-io.

        Do i-cont = 2 To 25:

            If tt-pedido.narrativa[i-cont] <> "" Then
                Put tt-pedido.narrativa[i-cont] At 58.

        End.

        Do i-cont = 2 To 25:

            If tt-pedido.narrativa-cot[i-cont] <> "" Then
                Put tt-pedido.narrativa-cot[i-cont] At 133.

        End.


        Put Skip(1).
        
        If Last-of(tt-pedido.cod-comprador) Then Do:
            Put Skip(2).
        End.

    End.

    Page.

    For Each tt-ord-aux:

        Assign i-tot-ord = i-tot-ord + 1.

        If tt-ord-aux.maior Then
            Assign i-tot-ord-maior = i-tot-ord-maior + 1.

    End.


    Assign i-tot-ord-maior = 0.

    For Each tt-pedido Where tt-pedido.l-branco = No :

        Assign i-tot-ord-maior = i-tot-ord-maior + 1.

    End.


    Put "                   Total de Pedidos: " i-tot-ord                                                       Skip
        "    Total de Pedidos c/ valor maior: " i-tot-ord-maior                                                 Skip(1)
        "          Percentual c/ valor maior: " ((i-tot-ord-maior * 100) / i-tot-ord)   Format ">>9.99" " %"    Skip(2)
        "       SELE€ÇO"                                                                                        Skip(1)
        "               Estabelecimento: " tt-param.cod-est-ini     At 41 " <||> " tt-param.cod-est-fim         Skip
        "                        Pedido: " tt-param.num-pedido-ini  At 41 " <||> " tt-param.num-pedido-fim      Skip
        "                 Grupo Estoque: " tt-param.ge-codigo-ini   At 48 " <||> " tt-param.ge-codigo-fim       Skip
        "                   Data Pedido: " tt-param.data-pedido-ini At 40 " <||> " tt-param.data-pedido-fim     Skip
        "                     Comprador: " tt-param.comprador-ini   At 38 " <||> " tt-param.comprador-fim       Skip(1)
        "       CLASSIFICA€ÇO"                                                                                  Skip(1)
        tt-param.desc-classifica    At 20.
        
        .

End Procedure.

/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME

&ENDIF

&IF DEFINED(EXCLUDE-pi-excel) = 0 &THEN

&ANALYZE-SUSPEND _UIB-CODE-BLOCK _PROCEDURE pi-excel Procedure 
PROCEDURE pi-excel :
/*------------------------------------------------------------------------------
  Purpose:     
  Parameters:  <none>
  Notes:       
------------------------------------------------------------------------------*/

Define Input Param p-arquivo As Char No-undo.
 
    DEFINE VARIABLE chExcelApplication AS COM-HANDLE         NO-UNDO.
    DEFINE VARIABLE chWorkbook         AS COM-HANDLE         NO-UNDO.
    DEFINE VARIABLE chWorksheet        AS COM-HANDLE         NO-UNDO.
    DEFINE VARIABLE iColumn            AS INTEGER INITIAL 2  NO-UNDO.
    DEFINE VARIABLE cColumn            AS CHARACTER          NO-UNDO.
    DEFINE VARIABLE cRange             AS CHARACTER          NO-UNDO.
    DEFINE VARIABLE c-arquivo-pdf      AS CHARACTER          NO-UNDO.
    
    RUN pi-acompanhar IN h-acomp (INPUT "Gerando arquivo...").
    
    REPEAT WHILE INDEX(p-arquivo,"/") > 0:
       SUBSTRING(p-arquivo,INDEX(p-arquivo,"/"),1) = "\".
    END.

    CREATE "Excel.Application" chExcelApplication.
    chWorkbook = chExcelApplication:Workbooks:OPEN(p-arquivo).
    
    chWorkSheet = chExcelApplication:Sheets:ITEM(1).
    chWorkSheet:Name = "ESCO133".    
    
    /*
    chWorkSheet:Columns("H:H"):Style   = "Comma".
    chWorkSheet:Columns("I:I"):Style   = "Comma".
    chWorkSheet:Columns("J:J"):Style   = "Comma".
    chWorkSheet:Columns("K:K"):Style   = "Comma".
    */
    /*
    chWorkSheet:Range("H:H"):NumberFormat = "dd/mm/aaaa".
    chWorkSheet:Range("I:I"):NumberFormat = "dd/mm/aaaa".
    chWorkSheet:Range("J:J"):NumberFormat = "dd/mm/aaaa".
    chWorkSheet:Range("K:K"):NumberFormat = "dd/mm/aaaa".
    */
    /*
    chWorkSheet:Columns("N:N"):Style   = "Comma".    
    chWorkSheet:Columns("O:O"):Style   = "Comma".    
    chWorkSheet:Columns("P:P"):Style   = "Comma".
    */
    
    /*
    chWorkSheet:Columns("X:X"):Style   = "Comma".
    chWorkSheet:Columns("Y:Y"):Style   = "Comma".
    chWorkSheet:Columns("Z:Z"):Style   = "Comma".
    chWorkSheet:Columns("AA:AA"):Style = "Comma".
    chWorkSheet:Columns("AB:AB"):Style = "Comma".
    chWorkSheet:Columns("AC:AC"):Style = "Comma".
    chWorkSheet:Columns("AD:AD"):Style = "Comma".
    chWorkSheet:Columns("AE:AE"):Style = "Comma".
    */

    chWorkSheet:Columns("K:M"):Style   = "Comma".    

    chWorkSheet:ListObjects:ADD(,,,1,,"TableStyleMedium20"):NAME = "ESCO133".

    chWorkSheet:Columns("A:R"):autofit().
    chWorkSheet:Range("A1"):SELECT.

    chExcelApplication:Visible = TRUE.
       
    RELEASE OBJECT chExcelApplication NO-ERROR.
    RELEASE OBJECT chWorkbook         NO-ERROR.
    RELEASE OBJECT chWorksheet        NO-ERROR.
    

END PROCEDURE.

/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME

&ENDIF

&IF DEFINED(EXCLUDE-pi-monta-tt) = 0 &THEN

&ANALYZE-SUSPEND _UIB-CODE-BLOCK _PROCEDURE pi-monta-tt Procedure 
PROCEDURE pi-monta-tt :
/**/
    Def Var i-cont  As Int.

    Empty Temp-table tt-pedido.
    empty temp-table tt-cotacao-item.

    Assign i-cont-ord = 0.

    For Each pedido-compr Where
        pedido-compr.cod-estabel >= tt-param.cod-est-ini     And
        pedido-compr.cod-estabel <= tt-param.cod-est-fim     And
        pedido-compr.num-pedido  >= tt-param.num-pedido-ini  And
        pedido-compr.num-pedido  <= tt-param.num-pedido-fim  And
        pedido-compr.data-pedido >= tt-param.data-pedido-ini And
        pedido-compr.data-pedido <= tt-param.data-pedido-fim And 
        pedido-compr.responsavel >= tt-param.comprador-ini   And 
        pedido-compr.responsavel <= tt-param.comprador-fim   No-lock,

        Each ordem-compra No-lock
        WHERE ordem-compra.num-pedido = pedido-compr.num-pedido 
          AND ordem-compra.situacao   <> 4:

        Find item WHERE item.it-codigo = ordem-compra.it-codigo No-lock No-error.

        If item.ge-codigo < tt-param.ge-codigo-ini Or
           item.ge-codigo > tt-param.ge-codigo-fim Then Next.

        Run pi-acompanhar In h-acomp (Input "Pedido: " + String(pedido-compr.num-pedido) + 
                                            " Data: " + String(pedido-compr.data-pedido,"99/99/9999")).

        Find emitente Where
             emitente.cod-emitente = pedido-compr.cod-emitente No-lock No-error.
    
        Find item Where
            item.it-codigo = ordem-compra.it-codigo No-lock No-error.

        Find mgcolomb.cotacao Where 
             mgcolomb.cotacao.nr_cot = ordem-compra.nr-processo No-lock No-error.

        FIND proc-compra WHERE proc-compra.nr-processo = ordem-compra.nr-processo NO-LOCK NO-ERROR.

        Create tt-pedido.
        Assign tt-pedido.cod-estabel      = pedido-compr.cod-estabel 
               tt-pedido.num-pedido       = pedido-compr.num-pedido
               tt-pedido.numero-ordem     = ordem-compra.numero-ordem  
               tt-pedido.data-pedido      = pedido-compr.data-pedido
               tt-pedido.it-codigo        = ordem-compra.it-codigo
               tt-pedido.desc-item        = item.desc-item
               tt-pedido.qt-solic         = ordem-compra.qt-solic
               tt-pedido.preco-ordem      = ordem-compra.preco-unit  /* valor convertido de acordo com parametros item x fornecedor */
               tt-pedido.cod-emitente     = pedido-compr.cod-emitente
               tt-pedido.nome-abrev       = emitente.nome-abrev
               tt-pedido.cod-comprador    = pedido-compr.responsavel.

        Assign i-cont       = 0
               i-cont-ord   = i-cont-ord + 1.

        Run pi-print-editor(Input ordem-compra.narrativa,
                            Input 75).

        For Each tt-editor:

            If i-cont + 1 > 25 Then
                Leave.

            Assign i-cont = i-cont + 1
                   tt-pedido.narrativa[i-cont] = tt-editor.conteudo.

        End.
        
        Assign i-cont = 0.

        If Avail mgcolomb.cotacao Then Do:
        
           Find FIRST mgcolomb.cot-ordem 
                Where mgcolomb.cot-ordem.nr_cot          = mgcolomb.cotacao.nr_cot 
                  And mgcolomb.cot-ordem.numero-ordem    = ordem-compra.numero-ordem 
                  AND mgcolomb.cot-ordem.narrativa-item <> "" No-lock No-error.
           If Avail mgcolomb.cot-ordem Then Do:

               If mgcolomb.cot-ordem.narrativa-item <> "" Then Do:
                   Run pi-print-editor(Input "Referente Ordem.: " + String(cot-ordem.numero-ordem) + " / " + 
                                                                    cot-ordem.narrativa-item,
                                       Input 75).
                   Assign tt-pedido.l-branco = No.
               End.
               Else Do:

                   /*Run pi-print-editor(Input "Referente Cota‡Æo.: " + string(mgcolomb.cotacao.nr_cot) + " / " + mgcolomb.cotacao.narrativa,
                                       Input 75).*/

                   Assign tt-pedido.l-branco = Yes.


               End.

               ASSIGN tt-pedido.cod-justif  = cot-ordem.cod-justif
                      tt-pedido.desc-justif = if cot-ordem.cod-justif = 0 then "" else c-descr[cot-ordem.cod-justif].                      


           End.
           
           Assign tt-pedido.nr-processo = ordem-compra.nr-processo.
           
           /*If l-branco Then Do:
                Find tt-pedido Where 
                     tt-pedido.num-pedido   = ordem-compra.num-pedido And 
                     tt-pedido.numero-ordem = ordem-compra.numero-ordem No-error.
                If Avail tt-pedido Then Do:
                    Message tt-pedido.num-pedido.
                End.
           End.*/

           For Each tt-editor:
    
                If i-cont + 1 > 25 Then
                    Leave.
    
                Assign i-cont = i-cont + 1
                       tt-pedido.narrativa-cot[i-cont] = tt-editor.conteudo.
    
           End.
        End.       
        ELSE DO:

            If ordem-compra.comentarios <> "" Then Do:
               Run pi-print-editor(Input "Referente Ordem.: " + String(ordem-compra.numero-ordem) + " / " + ordem-compra.comentarios,
                                   Input 75).
               Assign tt-pedido.l-branco = No.
            End.
            Else Do:
                Assign tt-pedido.l-branco = Yes.
            End.

            For Each tt-editor:
    
                If i-cont + 1 > 25 Then
                    Leave.
    
                Assign i-cont = i-cont + 1
                       tt-pedido.narrativa-cot[i-cont] = tt-editor.conteudo.
            
            End.

        END.

        For Each cotacao-item Of ordem-compra 
            WHERE cotacao-item.preco-unit > 0 No-lock 
            By cotacao-item.preco-unit:

               
               IF cotacao-item.codigo-icm = 1 THEN DO:
                   
                  IF (cotacao-item.aliquota-icm = 0) AND 
                     (cotacao-item.aliquota-ipi = 0) THEN DO:
                     ASSIGN tt-pedido.preco-menor = cotacao-item.preco-fornec.
                     Assign tt-pedido.preco-menor = tt-pedido.preco-menor - (tt-pedido.preco-menor * (cotacao-item.perc-descto / 100)).
                  END.
                  ELSE DO:
                     IF cotacao-item.aliquota-ipi <> 0  THEN DO:
                        ASSIGN tt-pedido.preco-menor = cotacao-item.preco-fornec * (1 + (cotacao-item.aliquota-ipi / 100)).
                        Assign tt-pedido.preco-menor =  tt-pedido.preco-menor - (tt-pedido.preco-menor * (cotacao-item.perc-descto / 100)).
                     END.
                     ELSE DO:
                        Assign tt-pedido.preco-menor = cotacao-item.preco-fornec.
                        Assign tt-pedido.preco-menor = tt-pedido.preco-menor - (tt-pedido.preco-menor * (cotacao-item.perc-descto / 100)).
                     END.
                  END.
               END.
               ELSE DO:
               
                   IF cotacao-item.codigo-icm = 2 THEN DO:
                      IF cotacao-item.aliquota-ipi <> 0 THEN DO:
                          ASSIGN tt-pedido.preco-menor = cotacao-item.preco-fornec / (1 + (cotacao-item.aliquota-ipi / 100)).
                          Assign tt-pedido.preco-menor = tt-pedido.preco-menor - (tt-pedido.preco-menor * (cotacao-item.perc-descto / 100)).
                      END.
                      ELSE DO:
                          ASSIGN de-preco-sem-impostos = cotacao-item.preco-fornec.
                          ASSIGN de-preco-sem-impostos = de-preco-sem-impostos * (1 - (cotacao-item.aliquota-icm / 100)).
                          ASSIGN tt-pedido.preco-menor = de-preco-sem-impostos.

                          Assign tt-pedido.preco-menor = tt-pedido.preco-menor - (tt-pedido.preco-menor * (cotacao-item.perc-descto / 100)).
                      END.
                   END.
               END.
               

               /*C lculo do cr‚dito com imposto*/
                RUN api/escoapi004.p(INPUT ordem-compra.cod-estabel,
                                     INPUT cotacao-item.it-codigo,
                                     INPUT cotacao-item.cod-emitente,
                                     INPUT tt-pedido.preco-menor,
                                     OUTPUT de-valor-icm,
                                     OUTPUT de-valor-pis,
                                     OUTPUT de-valor-cofins).
                /*---------------------------------*/
                
                /**************************************************************************************************
                if de-valor-icm = 0 then do:
                
                   Assign de-valor-icm = tt-pedido.preco-menor * ((cotacao-item.aliquota-icm / 100)).                   
                                   
                end.
                **************************************************************************************************/

                Assign tt-pedido.preco-menor  = tt-pedido.preco-menor - (de-valor-icm + de-valor-pis + de-valor-cofins). 

                /*RUN pi-preco-presente.*/
                
               /*Assign tt-pedido.preco-menor  = cotacao-item.preco-unit.*/
    
               /*Leave.*/
               
               create tt-cotacao-item.
               buffer-copy cotacao-item to tt-cotacao-item.
               assign tt-cotacao-item.preco-calc = tt-pedido.preco-menor.

            
        End.
        
        Assign tt-pedido.preco-menor = 0.
        for each tt-cotacao-item Of ordem-compra 
            WHERE tt-cotacao-item.preco-unit > 0
               By tt-cotacao-item.preco-calc:
               
            Find emitente Where emitente.cod-emitente = tt-cotacao-item.cod-emitente No-lock No-error.
        
            Assign tt-pedido.preco-menor  = tt-cotacao-item.preco-unit.
            
            leave.
            
        end.
        
        /*
        output to value('c:\tmp\cotacao.csv') no-map no-convert append.
        
        for each tt-cotacao-item Of ordem-compra 
            WHERE tt-cotacao-item.preco-unit > 0
               By tt-cotacao-item.preco-calc:
        
            
            disp tt-cotacao-item.numero-ordem
                 tt-cotacao-item.it-codigo
                 tt-cotacao-item.preco-unit
                 tt-cotacao-item.preco-calc.
        
        end.
        
        output close.
        */


        Find tt-ord-aux Where
            tt-ord-aux.numero-ordem = ordem-compra.numero-ordem And 
            tt-ord-aux.num-pedido   = ordem-compra.num-pedido   No-error.

        If Not Avail tt-ord-aux Then Do:

            Create tt-ord-aux.
            Assign tt-ord-aux.numero-ordem = ordem-compra.numero-ordem
                   tt-ord-aux.num-pedido   = ordem-compra.num-pedido.

        End.

        If tt-pedido.preco-ordem = tt-pedido.preco-menor Then Do:
           Delete tt-pedido.
        End.
        Else Do:
            Assign tt-ord-aux.maior = True.
        End.
    End.

End Procedure.

/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME

&ENDIF

&IF DEFINED(EXCLUDE-pi-preco-presente) = 0 &THEN

&ANALYZE-SUSPEND _UIB-CODE-BLOCK _PROCEDURE pi-preco-presente Procedure 
PROCEDURE pi-preco-presente :
/*------------------------------------------------------------------------------
  Purpose:   Calcula pre?o presente  
  Parameters:  <none>
  Notes:       
------------------------------------------------------------------------------*/

DEF VAR i-cont-pagto AS INTEGER FORMAT ">>9" INIT 0.       
DEFINE VARIABLE de-taxa-financ-eq AS DECIMAL NO-UNDO.
DEFINE VARIABLE i-nro-dia         AS INTEGER NO-UNDO. 

DEFINE VARIABLE de-valor-icm        AS DECIMAL                     NO-UNDO.
DEFINE VARIABLE de-valor-pis        AS DECIMAL                     NO-UNDO.
DEFINE VARIABLE de-valor-cofins     AS DECIMAL                     NO-UNDO.

FIND param-compra-ext NO-LOCK NO-ERROR.

         ASSIGN de-log-natural = 2.71828182845904.
           ASSIGN de-prazos       = 0
                  de-prazo-medio  = 0
                  vl-tx-fin-dia   = 0
                  de-preco-negoc  = 0.

           FIND cond-pagto WHERE 
                cond-pagto.cod-cond-pag = cotacao-item.cod-cond-pag NO-LOCK NO-ERROR.
           IF AVAIL cond-pagto THEN DO: 
              DO i-cont-pagto = 1 TO 12.
                ASSIGN de-prazos  = de-prazos + cond-pagto.prazos[i-cont-pagto].
              END.
                ASSIGN de-prazo-medio = de-prazos / cond-pagto.num-parcela.
           END.

           /*Calculo do Valor Presente*/
           ASSIGN de-preco-negoc = tt-pedido.preco-menor.
           
           if cotacao-item.nr-dias-taxa > 0 then do:
              ASSIGN de-taxa-financ-eq = cotacao-item.valor-taxa
                     i-nro-dia         = cotacao-item.nr-dias-taxa.
           end. 
           else do: 
            
                FIND FIRST es-taxa-financ WHERE es-taxa-financ.data-ini >= ordem-compra.data-emissao
                                            AND es-taxa-financ.data-fim <= ordem-compra.data-emissao NO-LOCK NO-ERROR.
                                           
                IF AVAILABLE es-taxa-financ THEN DO:
                    ASSIGN de-taxa-financ-eq = es-taxa-financ.taxa-financ-eq
                           i-nro-dia         = es-taxa-financ.nro-dia.
                END.
                ELSE DO:
                    FIND FIRST param-compra-ext NO-LOCK NO-ERROR.  
                    IF AVAILABLE param-compra-ext THEN DO:
                        ASSIGN de-taxa-financ-eq = param-compra-ext.taxa-financ-eq
                               i-nro-dia         = param-compra-ext.nro-dia.
                    END.                                              
                END.   
           end.     
           
          
           RUN ccp/cc9020.p (INPUT false,
                             INPUT cotacao-item.cod-cond-pag,
                             INPUT de-taxa-financ-eq,
                             INPUT i-nro-dia,
                             INPUT de-preco-negoc,
                             OUTPUT de-pre-unit-for).           
           
           /*C lculo do cr‚dito com imposto*/
           RUN api/escoapi004.p(INPUT ordem-compra.cod-estabel,
                                INPUT cotacao-item.it-codigo,
                                INPUT cotacao-item.cod-emitente,
                                INPUT tt-pedido.preco-menor,
                                OUTPUT de-valor-icm,
                                OUTPUT de-valor-pis,
                                OUTPUT de-valor-cofins).
            /*---------------------------------*/
           
           Assign tt-pedido.preco-menor   = de-pre-unit-for.                     
                       
           
    

END PROCEDURE.

/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME

&ENDIF


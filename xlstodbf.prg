// Programa   : XLSTODBF
// Fecha/Hora : 19/10/2004 16:39:45
// Propósito  : Convierte XLS hacia DBF
// Creado Por : Juan Navas
// Llamado por: 
// Aplicación : Todas
// Tabla      :

#INCLUDE "DPXBASE.CH"

PROCE MAIN(cOrigen,cDestino,oMeter,oSay,lAuto,nIni,nCant,nHead,nColGet,lStruct,cMaxCol,aSelect)
  LOCAL nAt
  LOCAL oTable,dFecha,cValue:="",cField,uData,nLen,nDec,cType
  LOCAL i,oExcel,nCol:=0,aFields:={},lEof:=.F.,nLine:=2,lAppend,nContar:=0,nCuantos:=0,nMaxCol

  DEFAULT cOrigen := cFilePath( GetModuleFileName( GetInstance() ))+"EJEMPLO\data.xls"

  DEFAULT cDestino:=STRTRAN(UPPE(cOrigen),".XLS",".DBF")

  DEFAULT oDp:aSelect:={}

  DEFAULT lAuto   :=.F.,;
          nIni    :=1  ,;
          nCant   :=0  ,;
          nHead   :=nIni-1,;
          nColGet :=1  ,;
          lStruct :=.F.,;
          cMaxCol :="" ,;
          aSelect :=oDp:aSelect

  nHead  :=MAX(nHead,1) //Indica Línea del Encabezado (Nombre de los Campos)
  cOrigen:=ALLTRIM(cOrigen)
  nMaxCol:=IF(Empty(cMaxCol),0,ASC(cMaxCol))

  IF !FILE(cOrigen)
     MensajeErr("Archivo "+cOrigen+" no Existe")
     RETURN .F.
  ENDIF

  IIF(ValType(oSay)="O",oSay:SetText("Abriendo "+cOrigen),NIL)

  oExcel := TExcelScript():New()
  oExcel:Open( cOrigen )

  //EJECUTAR("INSPECT",oExcel)
  // Lectura de campos, puede existir columnas vacias intermedias, por esto debe leer la estructura del diseño
  WHILE .T.

   nCol++

   cField:=oExcel:Get( nHead , nCol ,@cValue )
   uData :=oExcel:Get( nIni  , nCol ,@cValue )

   IF EMPTY(cField) .AND. Empty(uData) .AND. Empty(cMaxCol)
      EXIT
   ENDIF

   IF lAuto
      cField:=CHR(64+nCol) // "FIELD"+LSTR(nCol)
   ELSE 
      cField:=UPPE(ALLTRIM(cField))
   ENDIF

   nLen:=90
   nDec:=0

//? cField,nCol,len(aSelect)


   IF ValType(uData)="C"
       nLen:=250 // LEN(uData)
   ENDIF

   IF ValType(uData)="D"
      nLen:=8
   ENDIF

   IF ValType(uData)="N"
      nLen:=19
      nDec:=4
   ENDIF

   // JN 10/07/2017
   nAt:=ASCAN(aSelect,{|a,n| ALLTRIM(a[3])==cField})

   IF nAt>0
      // Tipo de Datos segun tabla
      uData:=CTOO(uData,aSelect[nAt,7])
      nLen :=aSelect[nAt,8]
      nDec :=aSelect[nAt,9]

      IF ValType(uData)="N"
         nLen:=19
         nDec:=4
      ENDIF

   ENDIF

   // jn 02/11/2022 no puede tener longitud cero
   IF nLen=0 .AND. ValType(uData)="C"
      nLen:=250
   ENDIF

//   IF Empty(cField)
//      EXIT
//	   ENDIF

   AADD(aFields,{cField,ValType(uData),nLen,nDec})

// ? nCol,cField,cMaxCol,"<-cMaxCol",nMaxCol,nCol,"nMaxCol,nCol"

   IF !Empty(cMaxCol) .AND. cField=cMaxCol
      EXIT
   ENDIF

   IF !Empty(cMaxCol) .AND. nCol>=nMaxCol
      EXIT
   ENDIF

  ENDDO


  IF Empty(aFields)
     MensajeErr("Estructura Vacia, Línea #"+LSTR(nIni)+", del Archivo "+ALLTRIM(cOrigen)+" está vacia")
     RETURN .F.
  ENDIF

  IF lStruct
     RETURN aFields
  ENDIF 

  FERASE(cDestino)

  IF FILE(cDestino)
     MensajeErr("Archivo "+cDestino+" posiblemente está Abierto o Protegido")
     RETURN ""
  ENDIF

  lEof   :=.F.
  nContar:=0
  nLine  :=nIni

  IIF(ValType(oSay)="O",oSay:SetText("Contando Registros"),NIL)

  WHILE !lEof

   nContar++

   uData :=oExcel:Get( nLine , nColGet ,@cValue )

// oDp:oFrameDp:SetText(LSTR(nLine)+"Y:"+LSTR(nColGet)+"DATA:"+CTOO(uData))

   IF EMPTY(uData)
      lEof:=.T.
      LOOP
   ENDIF

   nLine++

  ENDDO

  nCuantos:=nContar-nIni

  DBCREATE(cDestino,aFields,"DBFCDX")

  SELE A
  IIF(DPSELECT("DESTINO"),DBCLOSEAREA(),NIL)
  USE (cDestino) ALIAS "DESTINO" EXCLU

// ViewArray(aFields)

  IF !oMeter=NIL
   oMeter:SetTotal(nCuantos)
  ENDIF

  lEof   :=.F.
  nContar:=0
  nLine  :=nIni  // 1

  SET DECI TO 4

  WHILE !lEof

   lAppend:=.F.
   nContar++

   IF ValType(oSay)="O"
     oSay:SetText("Leyendo "+LSTR(nline)+"/"+LSTR(nCuantos))
     oMeter:Set(nLine)
   ENDIF

   IF nCant>0 .AND. nLine>=(nCant+nIni) 
     lEof:=.T.
     EXIT
   ENDIF

   IF nContar>20
      nContar:=0
      SysRefresh(.T.)
      CursorWait()
   ENDIF
 
   FOR nCol=1 TO LEN(aFields)

     cValue:=SPACE(100)
     uData :=oExcel:Get( nLine , nCol ,@cValue )

 // ? uData,cValue,nLine,nCol

     IF EMPTY(uData) .AND. nCol=1

        lEof:=.T.
        EXIT

     ENDIF

     IF !lAppend
        APPEND BLANK
        lAppend:=.T.
     ENDIF
        
     cType:=ValType(FieldGet(nCol))
     uData:=CTOO(uData,cType)

// ? nLine,nCol,uData,"uData",cValue

     IF cType<>ValType(uData) 
        MensajeErr("Columna : " +FieldName(nCol)+" Tipo de Dato ["+cType+"]"+CRLF+CTOO(uData,"C")+" Tipo"+"["+ValType(uData)+"]"+" Remueva Título de las Columnas","Diferencia en Tipo de Datos")
     ELSE
        FIELDPUT(nCol,uData) // CTOO(uData,cType))        
     ENDIF

   NEXT nCol

   nLine++

// IF nContar>5
//   RETURN 
// ENDIF

  ENDDO

  SELE A

//BROWSE()

  // oDp:aSelect[n,5]  // Son los Registros vacios


  IF LEN(oDp:aSelect)>0

    FOR I=1 TO LEN(oDp:aSelect)

      nAt:=FIELDPOS(oDp:aSelect[I,3])

      IF oDp:aSelect[I,5] .AND. nAt>0
         GO TOP
         DELETE FOR EMPTY(FIELDGET(nAt))
      ENDIF
  
    NEXT I

  ENDIF

// ? "debe remover los vacios segun aSelect"

//browse()

  USE

//  oExcel:Close(cOrigen )
  oExcel:End(.F.)
//  oExcel:=NIL
  
RETURN cDestino
// EOF


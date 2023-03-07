// Programa   : FRP
// Fecha/Hora : 05/03/2023 11:57:55
// Propósito  :
// Creado Por :
// Llamado por:
// Aplicación :
// Tabla      :

#INCLUDE "DPXBASE.CH"

PROCE MAIN()
  LOCAL cFileXls:="c:\FRP\quincena2.xlsx"
  LOCAL cFileDbf:="c:\FRP\quincena2.dbf"
  LOCAL aLine:={},aData:={},aCtas:={}
  LOCAL oDb  :=OpenOdbc(oDp:cDsnData),oTableDocC
  LOCAL cCodigo,cRif,cSql,cNumero,cWhere,cTipDoc:="NOM",dFecha:=oDp:dFecha,cOrg:="D"
  LOCAL nMonto :=0,nValCam:=1,dFchDec:=oDp:dFecha,cDescri:="Quincena del x al y"
  LOCAL cCodSuc:=oDp:cSucursal,nContar:=0

  // FERASE(cFileDbf)

  IF !FILE(cFileDbf)
     EJECUTAR("XLSTODBF",cFileXls,cFileDbf,NIL,NIL,.T.,2,nil,1,1,NIL,"M")
  ENDIF

  AADD(aCtas,{"C","SUELDO"      ,"611101"   ,0})
  AADD(aCtas,{"D","SSO"         ,"610401001",0})
  AADD(aCtas,{"E","LPH"         ,"611101"   ,0})
  AADD(aCtas,{"F","SFP"         ,"611101"   ,0})
  AADD(aCtas,{"G","ISLR"        ,"210105001",0})
  AADD(aCtas,{"H","PRESTAMO"    ,"611101"   ,0})
  AADD(aCtas,{"I","BONOPROD"    ,"611101"   ,0})
  AADD(aCtas,{"J","OTROS BONOS" ,"611101"   ,0})
  AADD(aCtas,{"K","PAGOEFECTIVO","611101"   ,0})

  cSql:=" SET FOREIGN_KEY_CHECKS = 0"
  oDb:Execute(cSql)

  EJECUTAR("IVALOAD")

  EJECUTAR("DPTIPDOCPROCREA",cTipDoc,"Nómina","D")
  SQLUPDATE("DPTIPDOCPRO","TDC_DOCEDI",.T.,"TDC_TIPO"+GetWhere("=",cTipDoc))


  oTableDocC:=OpenTable("SELECT * FROM DPDOCPROCTA",.F.)

  CLOSE ALL
  SELECT A
  USE (cFileDbf)
  GO TOP

  WHILE !EOF()
 
     cRif   :=STRTRAN(A->B,"-","")
     cCodigo:=SQLGET("DPPROVEEDOR","PRO_CODIGO","PRO_RIF"+GetWhere("=",cRif)+" OR PRO_RIF"+GetWhere("=",A->B))
     cWhere :="DOC_CODSUC"+GetWhere("=",oDp:cSucursal)+" AND DOC_TIPDOC"+GetWhere("=",cTipDoc)+" AND DOC_TIPTRA"+GetWhere("=","D")

     cNumero:=SQLINCREMENTAL("DPDOCPRO","DOC_NUMERO",cWhere,NIL,NIL,.T.,8)

     IF Empty(cCodigo)

       IF oDp:lRifPro
          cCodigo:=cRif
       ELSE
          cCodigo:=SQLINCREMENTAL("DPPROVEEDOR","PRO_CODIGO")
          cCodigo:=IF(Empty(cCodigo),cRif,cCodigo)
       ENDIF

       nMonto:=A->C
       EJECUTAR("CREATERECORD","DPPROVEEDOR",{"PRO_CODIGO","PRO_RIF","PRO_NOMBRE","PRO_TIPO"  ,"PRO_SITUAC" },; 
                                             {cCodigo     ,cRif     ,A->A        ,"Trabajador","Activo"},;
                                              NIL,.T.,"PRO_CODIGO"+GetWhere("=",cCodigo))

     ENDIF

     
     AEVAL(aCtas,{|a,n,nField| nField    :=FIELDPOS(a[1]),;
                               aCtas[n,4]:=FIELDGET(nField)})

     EJECUTAR("DPDOCPROCREA",cCodSuc,cTipDoc,cNumero,cNumero,cCodigo,dFecha,oDp:cMonedaExt,cOrg,oDp:cCenCos,nMonto,0,nValCam,dFchDec)

     aLine:={}
     FOR nContar=1 TO LEN(aCtas)
       AADD(aLine,{aCtas[nContar,3],oDp:cCenCos,aCtas[nContar,1],aCtas[nContar,2],aCtas[nContar,4],0})
     NEXT 

     EJECUTAR("DPDOCPROCTACREA",cCodSuc,cTipDoc,cCodigo,cNumero,aLine,oTableDocC)
  

     A->(DBSKIP())
	
  ENDDO

  oTableDocC:End()

  //  BROWSE()

  cSql:=" SET FOREIGN_KEY_CHECKS = 1"
  oDb:Execute(cSql)

  CLOSE ALL

RETURN .T.
// EOF

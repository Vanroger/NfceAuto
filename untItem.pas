unit untItem;

interface

type
  TItem = class(TComponent)
  private
    { private declarations }
    fnItem    : Integer;
    fcProd    : String;
    fcEAN     : String;
    fxProd    : String;
    fNCM      : String;
    fCEST     : String;
    fEXTIPI   : String;
    fCFOP     : String;
    fuCom     : String;
    fqCom     : Double;
    fvUnCom   : Double;
    fvProd    : Double;
    fcEANTrib : String;
    fuTrib    : String;
    fqTrib    : Double;
    fvUnTrib  : Double;
    fvFrete   : Double;
    fvSeg     : Double;
    fvDesc    : Double;
    fvOutro   : Double;
    fIndTot   : String;
    fxPed     : String;
    fnItemPed : String;
    fnFCI     : String;
    fnRECOPI  : String;

    //IMPOSTO
    forig        : String
    fCST         : String
    fCSOSN       : String
    fmodBC       : String
    fpRedBC      : Double;
    fvBC         : Double;
    fpICMS       : Double;
    fvICMS       : Double;
    fmodBCST     : String;
    fpMVAST      : Double;
    fpRedBCST    : Double;
    fvBCST       : Double;
    fpICMSST     : Double;
    fvICMSST     : Double;
    fUFST        : String;
    fpBCOp       : Double;
    fvBCSTRet    : Double;
    fvICMSSTRet  : Double;
    fmotDesICMS  : String;
    fpCredSN     : Double;
    fvCredICMSSN : Double;
    fvBCSTDest   : Double;
    fvICMSSTDest : Double;
    fvICMSDeson  : Double;
    fvICMSOp     : Double;
    fpDif        : Double;
    fvICMSDif    : Double;

  protected
    { protected declarations }
  public
    { public declarations }

  published
    { published declarations }
    property nItem    : Integer read fnItem    write fnItem;
    property cProd    : String  read fcProd    write fcProd;
    property cEAN     : String  read fcEAN     write fcEAN;
    property xProd    : String  read fxProd    write fxProd;
    property NCM      : String  read fNCM      write fNCM;
    property CEST     : String  read fCEST     write fCEST;
    property EXTIPI   : String  read fEXTIPI   write fEXTIPI;
    property CFOP     : String  read fCFOP     write fCFOP;
    property uCom     : String  read fuCom     write fuCom;
    property qCom     : Double  read fqCom     write fqCom;
    property vUnCom   : Double  read fvUnCom   write fvUnCom;
    property vProd    : Double  read fvProd    write fvProd;
    property cEANTrib : String  read fcEANTrib write fcEANTrib;
    property uTrib    : String  read fuTrib    write fuTrib;
    property qTrib    : Double  read fqTrib    write fqTrib;
    property vUnTrib  : Double  read fvUnTrib  write fvUnTrib;
    property vFrete   : Double  read fvFrete   write fvFrete;
    property vSeg     : Double  read fvSeg     write fvSeg;
    property vDesc    : Double  read fvDesc    write fvDesc;
    property vOutro   : Double  read fvOutro   write fvOutro;
    property IndTot   : String  read fIndTot   write fIndTot;
    property xPed     : String  read fxPed     write fxPed;
    property nItemPed : String  read fnItemPed write fnItemPed;
    property nFCI     : String  read fnFCI     write fnFCI;
    property nRECOPI  : String  read fnRECOPI  write fnRECOPI;
    //imposto
    property orig        : String read forig         write forig;
    property CST         : String read fCST          write fCST;
    property CSOSN       : String read fCSOSN        write fCSOSN;
    property modBC       : String read fmodBC        write fmodBC;
    property pRedBC      : Double read fpRedBC       write fpRedBC;
    property vBC         : Double read fvBC          write fvBC;
    property pICMS       : Double read fpICMS        write fpICMS;
    property vICMS       : Double read fvICMS        write fvICMS;
    property modBCST     : String read fmodBCST      write fmodBCST;
    property pMVAST      : Double read fpMVAST       write fpMVAST;
    property pRedBCST    : Double read fpRedBCST     write fpRedBCST;
    property vBCST       : Double read fvBCST        write fvBCST;
    property pICMSST     : Double read fpICMSST      write fpICMSST;
    property vICMSST     : Double read fvICMSST      write fvICMSST;
    property UFST        : String read fUFST         write fUFST;
    property pBCOp       : Double read fpBCOp        write fpBCOp;
    property vBCSTRet    : Double read fvBCSTRet     write fvBCSTRet;
    property vICMSSTRet  : Double read fvICMSSTRet   write fvICMSSTRet;
    property motDesICMS  : String read fmotDesICMS   write fmotDesICMS;
    property pCredSN     : Double read fpCredSN      write fpCredSN;
    property vCredICMSSN : Double read fvCredICMSSN  write fvCredICMSSN;
    property vBCSTDest   : Double read fvBCSTDest    write fvBCSTDest;
    property vICMSSTDest : Double read fvICMSSTDest  write fvICMSSTDest;
    property vICMSDeson  : Double read fvICMSDeson   write fvICMSDeson;
    property vICMSOp     : Double read fvICMSOp      write fvICMSOp;
    property pDif        : Double read fpDif         write fpDif;
    property vICMSDif    : Double read fvICMSDif     write fvICMSDif;
  end;

implementation

end.

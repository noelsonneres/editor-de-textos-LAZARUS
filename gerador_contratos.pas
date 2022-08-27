unit Classe.Entity.Contratos.Alunos;

interface

uses
  Interfaces, Classe.Query, Data.DB, Vcl.DBGrids, Vcl.Mask, Vcl.Forms,
  Winapi.Windows, Vcl.Controls, Dialogs, Vcl.StdCtrls, System.SysUtils,
  Datasnap.DBClient, UClasse.ObertValorPorExtenso,
  UClasse.Obter.Data.Hora.PorExtenso, DataModule, Vcl.ComCtrls,
  System.Classes, vcl.Printers, system.Generics.Collections, RxRichEd,
  System.Win.ComObj, System.Variants;

Type
  TClasseContratoAlunos = class(TInterfacedObject, iCadastroGerarContrado)
  private
    FQuery: iQueryRDW;
    FEntityValorPorExtenso: iObterValorPorExtenso;
    FEntityDataPorExtenso: iObterDataHoraPorExtenso;
    FTabela: string;
    FCampo: string;
    FValor: string;
    FDataInicial: TDate;
    FDataFinal: TDate;

    FListaTags: TDictionary<string, string>;

    FID: integer;
    FNUMERO_CONTRATO: string;
    FID_ALUNO: integer;
    FNOME_ALUNO: string;
    FID_MATRICULA: integer;
    FDATA_NASCIMENTO: TDate;
    FDATA_CADASTRO: TDate;
    FDATA_MATRICULA: TDate;
    FCPF: string;
    FRG: string;
    FORGAO_EXPEDIDOR: string;
    FNACIONALIDADE_ALUNO: string;
    FNACIONALIDADE: string;
    FENDERECO: string;
    FBAIRRO: string;
    FNUMERO: integer;
    FCOMPLEMENTO: string;
    FCEP: string;
    FCIDADE: string;
    FESTADO: string;
    FTELEFONE: string;
    FCELULAR: string;
    FWHATSAPP: string;
    FEMAIL: string;
    FTIPO_MARKETING: string;
    FNOME_MAE: string;
    FNOME_PAI: string;
    FNOME_RESPONSAVEL: string;
    FDATA_NASCIMENTO_RESPONSAVEL: string;
    FDATA_CADASTRO_RESPONSAVEL: string;
    FCPF_RESPONSAVEL: string;
    FRG_RESPONSAVEL: string;
    FORGAO_EXPEDIDOR_RESPONSAVEL: string;
    FNACIONALIDADE_RESPONSAVEL: string;
    FENDERECO_RESPONSAVEL: string;
    FBAIRRO_RESPONSAVEL: string;
    FNUMERO_RESPONSAVEL: integer;
    FCOMPLEMENTO_RESPONSAVEL: string;
    FCEP_RESPONSAVEL: string;
    FCIDADE_RESPONSAVEL: string;
    FESTADO_RESPONSAVEL: string;
    FTELEFONE_RESPONSAVEL: string;
    FCELULAR_RESPONSAVEL: string;
    FWHATSAPP_RESPONSAVEL: string;
    FEMAIL_RESPONSAVEL: string;
    FPACOTE_ESCOLHIDO: string;
    FVALOR_TOTAL_PACOTE: Currency;
    FVALOR_TOTAL_PACOTE_EXTENSO: string;
    FVALOR_PACOTE_DESCONTO: Currency;
    FVALOR_PACOTE_DESCONTO_EXTENSO: string;
    FVALOR_PARCELA: Currency;
    FVALOR_PARCELA_EXTENSO: string;
    FVALOR_PARCELA_DESCONTO: Currency;
    FVALOR_PARCELA_DESCONTO_EXTENSO: string;
    FQUANTIDADE_PARCELAS: integer;
    FQUANTIDADE_PARCELAS_EXTENSO: string;
    FDATA_INICIO: TDate;
    FDATA_INICIO_EXTENSO: string;
    FDATA_TEMINO: TDate;
    FDATA_TEMINO_EXTENSO: string;
    FVENCIMENTO_A_PARTIR: TDate;
    FVENCIMENTO_A_PARTIR_EXTENSO: string;
    FDURACAO_MEDIA_CURSO: Real;
    FDURACAO_MEDIA_CURSO_EXTENSO: string;
    FSITUACAO_MATRICULA: string;
    FID_CONSULTOR: integer;
    FCONSULTOR: string;
    FID_FUNCIONARIO: integer;
    FNOME_FUNCIONARIO: string;
    FOBS: string;
    FDia_Vencimento:string;
    FQtdeDiasAulas: integer;
    FQtdeHorasAulas: integer;

    FDIAS_AULAS: string;
    FHORARIOS_AULAS: string;
    FTURMA: string;
    FSALA: string;
    FVALOR_MATRICULA: Currency;
    FVALOR_MATRICULA_EXTENSO: string;
    FVENCIMENTO_MATRICULA: TDate;
    FVENCIMENTO_MATRICULA_EXTENSO: string;
    FMODULOS: string;

    FDIA_HORARIO: string;
    FDIA_HORARIO_TURMA: string;
    FDIA_HORARIO_TURMA_SALA: string;

    FNOME_FANTASIA: string;
    FRAZAO_SOCIAL: string;
    FCNPJ: string;
    FINSCRICAO_ESTADUAL: string;
    FDATA_ABERTURA: string;
    FSOCIOS: string;
    FENDERECO_EMPRESA: string;
    FBAIRRO_EMPRESA: string;
    FNUMERO_EMPRESA: integer;
    FCOMPLEMENTO_EMPRESA: string;
    FCEP_EMPRESA: string;
    FCIDADE_EMPRESA: string;
    FESTADO_EMPRESA: string;
    FTELEFONE_EMPRESA: string;
    FCELULAR_EMPRESA: string;
    FWHATSAPP_EMPRESA: string;
    FEMAIL_EMPRESA: string;

    FMATERIAIS:string;
    FVALOR_MATERIAIS:string;
    FQUANTIDADE_MATERIAIS:string;

    FLocal: string;
    FDestino1: string;
    FDestino2: string;

    procedure FindReplace(const Enc, subs: String; var Texto: TRichedit);
    procedure SubstituirTexto(TextoAntigo, TextoNovo: string; Rich: TRxRichEdit);

  public

    function getCampo(value: string): iCadastroGerarContrado; overload;
    function getValor(value: string): iCadastroGerarContrado; overload;
    function getCampo(value: TEdit): iCadastroGerarContrado; overload;
    function getValor(value: TEdit): iCadastroGerarContrado; overload;
    function getDataInicial(value: string): iCadastroGerarContrado; overload;
    function getDataFinal(value: string): iCadastroGerarContrado; overload;
    function getDataInicial(value: TMaskEdit): iCadastroGerarContrado; overload;
    function getDataFinal(value: TMaskEdit): iCadastroGerarContrado; overload;
    function ExecuteSql: iCadastroGerarContrado;
    function sqlPesquisa: iCadastroGerarContrado;
    function sqlPesquisaData: iCadastroGerarContrado;
    function sqlPesquisaEstatica: iCadastroGerarContrado;

    function abrir: iCadastroGerarContrado;
    function inserir: iCadastroGerarContrado;
    function gravar: iCadastroGerarContrado;
    function deletar: iCadastroGerarContrado;
    function atualizar: iCadastroGerarContrado;
    function editar: iCadastroGerarContrado;
    function cancelar: iCadastroGerarContrado;
    function fecharQuery: iCadastroGerarContrado;
    function codigoCadastro(sp: string): integer;
    function listarGrid(value: TDataSource): iCadastroGerarContrado;
    function ordenarGrid(column: TColumn): iCadastroGerarContrado;

    function getID(value: integer): iCadastroGerarContrado; overload;
    function getNUMERO_CONTRATO(value:string): iCadastroGerarContrado;
    function getID_ALUNO(value:integer): iCadastroGerarContrado;
    function getNOME_ALUNO(value:string): iCadastroGerarContrado;
    function getID_MATRICULA(value:integer): iCadastroGerarContrado;
    function getDATA_NASCIMENTO(value:TDateTime): iCadastroGerarContrado;
    function getDATA_CADASTRO(value:TDateTime): iCadastroGerarContrado;
    function getDATA_MATRICULA(value:TDateTime): iCadastroGerarContrado;
    function getCPF(value:string): iCadastroGerarContrado;
    function getRG(value:string): iCadastroGerarContrado;
    function getORGAO_EXPEDIDOR(value:string): iCadastroGerarContrado;
    function getNACIONALIDADE_ALUNO(value:string): iCadastroGerarContrado;
    function getENDERECO(value:string): iCadastroGerarContrado;
    function getBAIRRO(value:string): iCadastroGerarContrado;
    function getNUMERO(value:integer): iCadastroGerarContrado;
    function getCOMPLEMENTO(value:string): iCadastroGerarContrado;
    function getCEP(value:string): iCadastroGerarContrado;
    function getCIDADE(value:string): iCadastroGerarContrado;
    function getESTADO(value:string): iCadastroGerarContrado;
    function getTELEFONE(value:string): iCadastroGerarContrado;
    function getCELULAR(value:string): iCadastroGerarContrado;
    function getWHATSAPP(value:string): iCadastroGerarContrado;
    function getEMAIL(value:string): iCadastroGerarContrado;
    function getTIPO_MARKETING(value:string): iCadastroGerarContrado;
    function getNOME_MAE(value:string): iCadastroGerarContrado;
    function getNOME_PAI(value:string): iCadastroGerarContrado;
    function getNOME_RESPONSAVEL(value:string): iCadastroGerarContrado;
    function getDATA_NASCIMENTO_RESPONSAVEL(value:TDateTime): iCadastroGerarContrado;
    function getDATA_CADASTRO_RESPONSAVEL(value:TDateTime): iCadastroGerarContrado;
    function getCPF_RESPONSAVEL(value:string): iCadastroGerarContrado;
    function getRG_RESPONSAVEL(value:string): iCadastroGerarContrado;
    function getORGAO_EXPEDIDOR_RESPONSAVEL(value:string): iCadastroGerarContrado;
    function getNACIONALIDADE_RESPONSAVEL(value:string): iCadastroGerarContrado;
    function getENDERECO_RESPONSAVEL(value:string): iCadastroGerarContrado;
    function getBAIRRO_RESPONSAVEL(value:string): iCadastroGerarContrado;
    function getNUMERO_RESPONSAVEL(value:Integer): iCadastroGerarContrado;
    function getCOMPLEMENTO_RESPONSAVEL(value:string): iCadastroGerarContrado;
    function getCEP_RESPONSAVEL(value:string): iCadastroGerarContrado;
    function getCIDADE_RESPONSAVEL(value:string): iCadastroGerarContrado;
    function getESTADO_RESPONSAVEL(value:string): iCadastroGerarContrado;
    function getTELEFONE_RESPONSAVEL(value:string): iCadastroGerarContrado;
    function getCELULAR_RESPONSAVEL(value:string): iCadastroGerarContrado;
    function getWHATSAPP_RESPONSAVEL(value:string): iCadastroGerarContrado;
    function getEMAIL_RESPONSAVEL(value:string): iCadastroGerarContrado;
    function getPACOTE_ESCOLHIDO(value:string): iCadastroGerarContrado;
    function getVALOR_TOTAL_PACOTE(value:Currency): iCadastroGerarContrado;
    function getVALOR_TOTAL_PACOTE_EXTENSO(value:Currency): iCadastroGerarContrado;
    function getVALOR_PACOTE_DESCONTO(value:Currency): iCadastroGerarContrado;
    function getVALOR_PACOTE_DESCONTO_EXTENSO(value:Currency): iCadastroGerarContrado;
    function getVALOR_PARCELA(value:Currency): iCadastroGerarContrado;
    function getVALOR_PARCELA_EXTENSO(value:Currency): iCadastroGerarContrado;
    function getVALOR_PARCELA_DESCONTO(value:Currency): iCadastroGerarContrado;
    function getVALOR_PARCELA_DESCONTO_EXTENSO(value:Currency): iCadastroGerarContrado;
    function getQUANTIDADE_PARCELAS(value:Integer): iCadastroGerarContrado;
    function getQUANTIDADE_PARCELAS_EXTENSO(value:integer): iCadastroGerarContrado;
    function getDATA_INICIO(value:TDateTime): iCadastroGerarContrado;
    function getDATA_INICIO_EXTENSO(value:string): iCadastroGerarContrado;
    function getDATA_TEMINO(value:TDateTime): iCadastroGerarContrado;
    function getDATA_TEMINO_EXTENSO(value:string): iCadastroGerarContrado;
    function getVENCIMENTO_A_PARTIR(value:TDateTime): iCadastroGerarContrado;
    function getVENCIMENTO_A_PARTIR_EXTENSO(value:string): iCadastroGerarContrado;
    function getDURACAO_MEDIA_CURSO(value:Real): iCadastroGerarContrado;
    function getDURACAO_MEDIA_CURSO_EXTENSO(value:string): iCadastroGerarContrado;
    function getSITUACAO_MATRICULA(value:string): iCadastroGerarContrado;
    function getID_CONSULTOR(value:Integer): iCadastroGerarContrado;
    function getCONSULTOR(value:string): iCadastroGerarContrado;
    function getID_FUNCIONARIO(value:Integer): iCadastroGerarContrado;
    function getNOME_FUNCIONARIO(value:string): iCadastroGerarContrado;
    function getOBS(value:string): iCadastroGerarContrado;

    function getDIAS_AULAS(value:string): iCadastroGerarContrado;
    function getHORARIOS_AULAS(value:string): iCadastroGerarContrado;
    function getTURMA(value:string): iCadastroGerarContrado;
    function getSALA(value:string): iCadastroGerarContrado;
    function getVALOR_MATRICULA(value:string): iCadastroGerarContrado;
    function getVALOR_MATRICULA_EXTENSO(value:string): iCadastroGerarContrado;
    function getVENCIMENTO_MATRICULA(value:string): iCadastroGerarContrado;
    function getVENCIMENTO_MATRICULA_EXTENSO(value:string): iCadastroGerarContrado;
    function getMODULOS(value:string): iCadastroGerarContrado;

    function getMATERIAIS(value:string):iCadastroGerarContrado;
    function getVALOR_MATERIAIS(value:currency):iCadastroGerarContrado;
    function getQUANTIDADE_MATERIAIS(value:integer):iCadastroGerarContrado;

    function getDIA_HORARIO(value:string): iCadastroGerarContrado;
    function getDIA_HORARIO_TURMA(value:string): iCadastroGerarContrado;
    function getDIA_HORARIO_TURMA_SALA(value:string): iCadastroGerarContrado;

    function getQTDEDIASAULAS(value:integer):iCadastroGerarContrado;
    function getQTDEHORASAULAS(value:integer):iCadastroGerarContrado;

    function getNOME_FANTASIA(value:string): iCadastroGerarContrado;
    function getRAZAO_SOCIAL(value:string): iCadastroGerarContrado;
    function getCNPJ(value:string): iCadastroGerarContrado;
    function getINSCRICAO_ESTADUAL(value:string): iCadastroGerarContrado;
    function getDATA_ABERTURA(value:string): iCadastroGerarContrado;
    function getSOCIOS(value:string): iCadastroGerarContrado;
    function getENDERECO_EMPRESA(value:string): iCadastroGerarContrado;
    function getBAIRRO_EMPRESA(value:string): iCadastroGerarContrado;
    function getNUMERO_EMPRESA(value:string): iCadastroGerarContrado;
    function getCOMPLEMENTO_EMPRESA(value:string): iCadastroGerarContrado;
    function getCEP_EMPRESA(value:string): iCadastroGerarContrado;
    function getCIDADE_EMPRESA(value:string): iCadastroGerarContrado;
    function getESTADO_EMPRESA(value:string): iCadastroGerarContrado;
    function getTELEFONE_EMPRESA(value:string): iCadastroGerarContrado;
    function getCELULAR_EMPRESA(value:string): iCadastroGerarContrado;
    function getWHATSAPP_EMPRESA(value:string): iCadastroGerarContrado;
    function getEMAIL_EMPRESA(value:string): iCadastroGerarContrado;

    function gerarContrato(value:TRxRichEdit):iCadastroGerarContrado;
    function gerarContratoWord(value:string):iCadastroGerarContrado;
    function returnClientDataSet:TClientDataSet;
    function verificarContrato(value:integer):boolean;

    function exportar: iCadastroGerarContrado;
    function validarData(componet: TMaskEdit): iCadastroGerarContrado;

    constructor create;
    destructor destroy; override;
    class function new: iCadastroGerarContrado;

  end;

implementation

{ TClasseContratoAlunos }

function TClasseContratoAlunos.abrir: iCadastroGerarContrado;
begin
  result := self;
  FQuery.openQuery(FTabela);
end;

function TClasseContratoAlunos.atualizar: iCadastroGerarContrado;
begin
  result := self;
  FQuery.TQuery.Refresh;
end;

function TClasseContratoAlunos.cancelar: iCadastroGerarContrado;
begin
  result := self;
  FQuery.TQuery.Cancel;
end;

function TClasseContratoAlunos.codigoCadastro(sp: string): integer;
begin
  result := 0;
end;

constructor TClasseContratoAlunos.create;
begin

  FTabela := 'DADOS_CONTRATOS';

  FQuery := TclasseQuery.new.openQueryCrud(FTabela);

  FEntityValorPorExtenso := TObertValorPorExtenco.new;
  FEntityDataPorExtenso := TObterDataHoraPorExtenso.new;
  FListaTags := TDictionary<string,string>.Create;

end;

function TClasseContratoAlunos.deletar: iCadastroGerarContrado;
var
  sErro: string;
begin

  result := self;

  if FQuery.TQuery.RecordCount >= 1 then
  begin
    if application.MessageBox('Deseja realmente excluir este registro?',
      'Pergunta do sistema!', MB_YESNO + MB_ICONQUESTION) = mryes then
    begin
      FQuery.TQuery.Delete;

      if FQuery.TQuery.ApplyUpdates(sErro) then
        showmessage('Registro excluido com sucesso')
      else
        raise Exception.create
          ('Não foi possível deletar o registro selecionado. ' + sErro);
    end;

  end;

end;

destructor TClasseContratoAlunos.destroy;
begin
  FreeAndNil(FListaTags);
  inherited;
end;

function TClasseContratoAlunos.editar: iCadastroGerarContrado;
begin

  result := self;

  if FQuery.TQuery.RecordCount >= 1 then
  begin
    FQuery.TQuery.Edit;
  end;

end;

function TClasseContratoAlunos.ExecuteSql: iCadastroGerarContrado;
begin
  result := self;
  FQuery.ExecuteSql(FTabela);
end;

function TClasseContratoAlunos.exportar: iCadastroGerarContrado;
begin
  result := self;
end;

function TClasseContratoAlunos.fecharQuery: iCadastroGerarContrado;
begin
  result := self;
  FQuery.TQuery.Close;
end;

procedure TClasseContratoAlunos.FindReplace(const Enc, subs: String;
  var Texto: TRichedit);
Var
  i, Posicao: Integer;
  Linha: string;
Begin

  For i:= 0 to Texto.Lines.count - 1 do
  begin

    Linha := Texto. Lines[i];
    Repeat
     Posicao:=Pos(Enc,Linha);
      If Posicao > 0 then
    Begin
     Delete(Linha,Posicao,Length(Enc));
      Insert(Subs,Linha,Posicao);
      Texto.Lines[i]:=Linha;
    end;
    until Posicao = 0;

  end;

End;

function TClasseContratoAlunos.getCampo(value: string): iCadastroGerarContrado;
begin
  result := self;
  FCampo := value;
end;

function TClasseContratoAlunos.getDataFinal(value: string): iCadastroGerarContrado;
begin

  result := self;

  try
    FDataFinal := StrToDate(value);
  except
    Dialogs.MessageDlg('Informe uma data valida.', mtError, [mbOk], 0, mbOk);
  end;

end;

function TClasseContratoAlunos.getDataInicial(value: string): iCadastroGerarContrado;
begin

  result := self;

  try
    FDataInicial := StrToDate(value);
  except
    Dialogs.MessageDlg('Informe uma data valida.', mtError, [mbOk], 0, mbOk);
  end;

end;

function TClasseContratoAlunos.getID(value: integer): iCadastroGerarContrado;
begin

  result := self;
  FID := value;

end;

function TClasseContratoAlunos.getID_ALUNO(value: integer): iCadastroGerarContrado;
begin

  result := self;

  if value <> 0 then
    FID_ALUNO := value
  else
  begin
    Dialogs.MessageDlg('ERRO! Informe o código do aluno', mtError, [mbOK], 0, mbOK);
    Abort;
  end;

end;

function TClasseContratoAlunos.getID_CONSULTOR(value: Integer): iCadastroGerarContrado;
begin

  result := self;
  FID_CONSULTOR := value


end;

function TClasseContratoAlunos.getID_FUNCIONARIO(value: Integer): iCadastroGerarContrado;
begin

  result := self;
  FID_FUNCIONARIO := value;

end;

function TClasseContratoAlunos.getID_MATRICULA(value: integer): iCadastroGerarContrado;
begin

  result := self;

  if value <> 0 then
    FID_MATRICULA := value
  else
  begin
    Dialogs.MessageDlg('ERRO! Informe a matricula do aluno', mtError, [mbOK], 0, mbOK);
    Abort;
  end;

end;

function TClasseContratoAlunos.getINSCRICAO_ESTADUAL(value: string): iCadastroGerarContrado;
begin

  result := self;

  FINSCRICAO_ESTADUAL := value;

end;

function TClasseContratoAlunos.getMATERIAIS(
  value: string): iCadastroGerarContrado;
begin

  result := self;
  FMATERIAIS := value;

end;

function TClasseContratoAlunos.getMODULOS(
  value: string): iCadastroGerarContrado;
begin

  result := self;
  FMODULOS := value;

end;

function TClasseContratoAlunos.getNACIONALIDADE_ALUNO(
  value: string): iCadastroGerarContrado;
begin

  result := self;
  FNACIONALIDADE_ALUNO := value;


end;

function TClasseContratoAlunos.getNACIONALIDADE_RESPONSAVEL(
  value: string): iCadastroGerarContrado;
begin

  result := self;
  FNACIONALIDADE_RESPONSAVEL := value;

end;

function TClasseContratoAlunos.getNOME_ALUNO(value: string): iCadastroGerarContrado;
begin

  result := self;

  if value <> EmptyStr then
    FNOME_ALUNO := value
  else
  begin
    Dialogs.MessageDlg('ERRO! Informe o nome do aluno', mtError, [mbOK], 0, mbOK);
    Abort;
  end;

end;

function TClasseContratoAlunos.getNOME_FANTASIA(value: string): iCadastroGerarContrado;
begin

  result := self;
  FNOME_FANTASIA := value;

end;

function TClasseContratoAlunos.getNOME_FUNCIONARIO(value: string): iCadastroGerarContrado;
begin

  result := self;
  FNOME_FUNCIONARIO := value;

end;

function TClasseContratoAlunos.getNOME_MAE(value: string): iCadastroGerarContrado;
begin

  result := self;
  FNOME_MAE := value;

end;

function TClasseContratoAlunos.getNOME_PAI(value: string): iCadastroGerarContrado;
begin

  result := self;
  FNOME_PAI := value;

end;

function TClasseContratoAlunos.getNOME_RESPONSAVEL(value: string): iCadastroGerarContrado;
begin

  result := self;
  FNOME_RESPONSAVEL := value;

end;

function TClasseContratoAlunos.getNUMERO(value: integer): iCadastroGerarContrado;
begin

  result := self;
  FNUMERO := value;

end;

function TClasseContratoAlunos.getNUMERO_CONTRATO(value: string): iCadastroGerarContrado;
begin

  result := self;

  if value <> EmptyStr then
    FNUMERO_CONTRATO := value
  else
  begin
    Dialogs.MessageDlg('ERRO! Informe o número de contrato do aluno', mtError, [mbOK], 0, mbOK);
    Abort;
  end;

end;

function TClasseContratoAlunos.getNUMERO_EMPRESA(value: string): iCadastroGerarContrado;
begin

  result := self;
  FNUMERO_EMPRESA := value.ToInteger;

end;

function TClasseContratoAlunos.getNUMERO_RESPONSAVEL(value: Integer): iCadastroGerarContrado;
begin

  result := self;
  FNUMERO_RESPONSAVEL := value;

end;

function TClasseContratoAlunos.getOBS(value: string): iCadastroGerarContrado;
begin

  result := self;
  FOBS := value;

end;

function TClasseContratoAlunos.getORGAO_EXPEDIDOR(value: string): iCadastroGerarContrado;
begin

  result := self;
  FORGAO_EXPEDIDOR := value;

end;

function TClasseContratoAlunos.getORGAO_EXPEDIDOR_RESPONSAVEL(value: string): iCadastroGerarContrado;
begin

  result := self;
  FORGAO_EXPEDIDOR_RESPONSAVEL := value;

end;

function TClasseContratoAlunos.getPACOTE_ESCOLHIDO(value: string): iCadastroGerarContrado;
begin

  result := self;

  if value <> EmptyStr then
    FPACOTE_ESCOLHIDO := value
  else
  begin
    Dialogs.MessageDlg('ERRO! Informe Pacote inserido na matricula do aluno', mtError, [mbOK], 0, mbOK);
    Abort;
  end;

end;

function TClasseContratoAlunos.getQUANTIDADE_MATERIAIS(
  value: integer): iCadastroGerarContrado;
begin

    result := self;

    try
      FQUANTIDADE_MATERIAIS := value.ToString;
    except on e:exception do
    begin
      Dialogs.MessageDlg('ERRO! Informe um valor valido para a quantidade de módulos', mtError, [mbOK], 0, mbOK);
      Abort;
    end;

    end;

end;

function TClasseContratoAlunos.getQUANTIDADE_PARCELAS(value: Integer): iCadastroGerarContrado;
begin

  result := self;
  FQUANTIDADE_PARCELAS := value;

end;

function TClasseContratoAlunos.getQUANTIDADE_PARCELAS_EXTENSO(value: integer): iCadastroGerarContrado;
begin

  result := self;

  FQUANTIDADE_PARCELAS_EXTENSO := FEntityValorPorExtenso.valorPorExtensoSimples(value);

end;

function TClasseContratoAlunos.getRAZAO_SOCIAL(value: string): iCadastroGerarContrado;
begin

  result := self;
  FRAZAO_SOCIAL := value;

end;

function TClasseContratoAlunos.getRG(value: string): iCadastroGerarContrado;
begin

  result := self;
  FRG := value;

end;

function TClasseContratoAlunos.getRG_RESPONSAVEL(value: string): iCadastroGerarContrado;
begin

  result := self;
  FRG_RESPONSAVEL := value;

end;

function TClasseContratoAlunos.getSALA(value: string): iCadastroGerarContrado;
begin

  result := self;
  FSala:=value;

end;

function TClasseContratoAlunos.getSITUACAO_MATRICULA(value: string): iCadastroGerarContrado;
begin

  result := self;
  FSITUACAO_MATRICULA := value;

end;

function TClasseContratoAlunos.getSOCIOS(value: string): iCadastroGerarContrado;
begin

  result := self;
  FSOCIOS := value;

end;

function TClasseContratoAlunos.getTELEFONE(value: string): iCadastroGerarContrado;
begin

  result := self;
  FTELEFONE := value;

end;

function TClasseContratoAlunos.getTELEFONE_EMPRESA(value: string): iCadastroGerarContrado;
begin

  result := self;
  FTELEFONE_EMPRESA := value;

end;

function TClasseContratoAlunos.getTELEFONE_RESPONSAVEL(value: string): iCadastroGerarContrado;
begin

  result := self;
  FTELEFONE_RESPONSAVEL := value;

end;

function TClasseContratoAlunos.getTIPO_MARKETING(value: string): iCadastroGerarContrado;
begin

  result := self;
  FTIPO_MARKETING := value;

end;

function TClasseContratoAlunos.getTURMA(value: string): iCadastroGerarContrado;
begin

  result := self;
  FTURMA := value;

end;

function TClasseContratoAlunos.getValor(value: TEdit): iCadastroGerarContrado;
begin

  result := self;

  if value.Text = EmptyStr then
  begin
    value.SetFocus;
    raise Exception.create('ERRO! Digite qual informação deseja buscar');
  end
  else
    FValor := UpperCase(value.Text);

end;

function TClasseContratoAlunos.getVALOR_MATRICULA(value: string): iCadastroGerarContrado;
begin

  result := self;

  try
    FVALOR_MATRICULA := StrToCurr(value);
  except
  begin
    Dialogs.MessageDlg('ERRO! Valor para a matricula invalido', mtError, [mbOK], 0, mbOK);
  end;

  end;

end;

function TClasseContratoAlunos.getVALOR_MATRICULA_EXTENSO(value: string): iCadastroGerarContrado;
begin

  result := self;

  FVALOR_MATRICULA_EXTENSO := FEntityValorPorExtenso.ValorPorExtenso(StrToCurr(value));

end;

function TClasseContratoAlunos.getVALOR_MATERIAIS(value: currency): iCadastroGerarContrado;
begin

  result := self;

  try
    FVALOR_MATERIAIS := CurrToStr(value);
  except on e:exception do
  begin
    Dialogs.MessageDlg('ERRO! Informe um valor valido para o total dos módulos', mtError, [mbOK], 0, mbOK);
    Abort;
  end;

  end;

end;

function TClasseContratoAlunos.getVALOR_PACOTE_DESCONTO(value: Currency): iCadastroGerarContrado;
begin

  result := self;
  FVALOR_PACOTE_DESCONTO := value;

end;

function TClasseContratoAlunos.getVALOR_PACOTE_DESCONTO_EXTENSO(value: Currency): iCadastroGerarContrado;
begin

  result := self;

  FVALOR_PACOTE_DESCONTO_EXTENSO := FEntityValorPorExtenso.ValorPorExtenso(value);

end;

function TClasseContratoAlunos.getVALOR_PARCELA(value: Currency): iCadastroGerarContrado;
begin

  result := self;
  FVALOR_PARCELA := value;

end;

function TClasseContratoAlunos.getVALOR_PARCELA_DESCONTO(value: Currency): iCadastroGerarContrado;
begin

  result := self;

  FVALOR_PARCELA_DESCONTO_EXTENSO := FEntityValorPorExtenso.ValorPorExtenso(value);

end;

function TClasseContratoAlunos.getVALOR_PARCELA_DESCONTO_EXTENSO(value: Currency): iCadastroGerarContrado;
begin

  result := self;

  FVALOR_PARCELA_DESCONTO_EXTENSO := FEntityValorPorExtenso.ValorPorExtenso(value);

end;

function TClasseContratoAlunos.getVALOR_PARCELA_EXTENSO(value: Currency): iCadastroGerarContrado;
begin

  Result := self;

  FVALOR_PARCELA_EXTENSO := FEntityValorPorExtenso.ValorPorExtenso(value);

end;

function TClasseContratoAlunos.getVALOR_TOTAL_PACOTE(value: Currency): iCadastroGerarContrado;
begin

  Result := self;

  FVALOR_TOTAL_PACOTE := value;

end;

function TClasseContratoAlunos.getVALOR_TOTAL_PACOTE_EXTENSO(value: Currency): iCadastroGerarContrado;
begin

  result := self;

  FVALOR_TOTAL_PACOTE_EXTENSO := FEntityValorPorExtenso.ValorPorExtenso(value);

end;

function TClasseContratoAlunos.getVENCIMENTO_A_PARTIR(value: TDateTime): iCadastroGerarContrado;
begin

  result := self;
  FVENCIMENTO_A_PARTIR := value;

  FDia_Vencimento := formatdatetime('dd', value);

end;

function TClasseContratoAlunos.getVENCIMENTO_A_PARTIR_EXTENSO(value: string): iCadastroGerarContrado;
begin

  result := self;
  FVENCIMENTO_A_PARTIR_EXTENSO := FEntityDataPorExtenso.dataPorExtenso(StrToDate(value));

end;

function TClasseContratoAlunos.getVENCIMENTO_MATRICULA(value: string): iCadastroGerarContrado;
begin

  result := self;

  try
    FVENCIMENTO_MATRICULA := StrToDate(value);
  except on e:exception do
  begin
    Dialogs.MessageDlg('ERRO! Data de vencimento da matricula não foi informada. '+e.Message,
     mtError, [mbOK], 0, mbOK);
  end;

  end;

end;

function TClasseContratoAlunos.getVENCIMENTO_MATRICULA_EXTENSO(value: string): iCadastroGerarContrado;
begin

  result := self;

  FVENCIMENTO_MATRICULA_EXTENSO := FEntityDataPorExtenso.dataPorExtenso(StrToDate(value));

end;

function TClasseContratoAlunos.getWHATSAPP(value: string): iCadastroGerarContrado;
begin

  result := self;
  FWHATSAPP := value;

end;

function TClasseContratoAlunos.getWHATSAPP_EMPRESA(value: string): iCadastroGerarContrado;
begin

  result := self;
  FWHATSAPP := value;

end;

function TClasseContratoAlunos.getWHATSAPP_RESPONSAVEL(value: string): iCadastroGerarContrado;
begin

  result := self;
  FWHATSAPP_RESPONSAVEL := value;

end;

function TClasseContratoAlunos.getValor(value: string): iCadastroGerarContrado;
begin

  result := self;

  if value = EmptyStr then
    raise Exception.create('ERRO! Digite qual informação deseja buscar');

  FValor := UpperCase(value);

end;

function TClasseContratoAlunos.gravar: iCadastroGerarContrado;
var
  sErro: String;
begin

  result := self;

  with FQuery.TQuery do
  begin

    FieldByName('ID').AutoGenerateValue := arAutoInc;
    FieldByName('NUMERO_CONTRATO').AsString := FNUMERO_CONTRATO;
    FieldByName('ID_ALUNO').AsInteger := FID_ALUNO;
    FieldByName('NOME_ALUNO').AsString := FNOME_ALUNO;
    FieldByName('ID_MATRICULA').AsInteger := FID_MATRICULA;
    FieldByName('CPF').AsString := FCPF;
    FieldByName('RG').AsString := FRG;
    FieldByName('ORGAO_EXPEDIDOR').AsString := FORGAO_EXPEDIDOR;
    FieldByName('ENDERECO').AsString := FENDERECO;
    FieldByName('BAIRRO').AsString := FBAIRRO;
    FieldByName('NUMERO').AsInteger := FNUMERO;
    FieldByName('COMPLEMENTO').AsString := FCOMPLEMENTO;
    FieldByName('CEP').AsString := FCEP;
    FieldByName('CIDADE').AsString := FCIDADE;
    FieldByName('ESTADO').AsString := FESTADO;
    FieldByName('TELEFONE').AsString := FTELEFONE;
    FieldByName('CELULAR').AsString := FCELULAR;
    FieldByName('WHATSAPP').AsString := FWHATSAPP;
    FieldByName('EMAIL').AsString := FEMAIL;
    FieldByName('TIPO_MARKETING').AsString := FTIPO_MARKETING;
    FieldByName('NOME_MAE').AsString := FNOME_MAE;
    FieldByName('NOME_PAI').AsString := FNOME_PAI;
    FieldByName('NOME_RESPONSAVEL').AsString := FNOME_RESPONSAVEL;
    FieldByName('CPF_RESPONSAVEL').AsString := FCPF_RESPONSAVEL;
    FieldByName('RG_RESPONSAVEL').AsString := FRG_RESPONSAVEL;
    FieldByName('ORGAO_EXPEDIDOR_RESPONSAVEL').AsString := FORGAO_EXPEDIDOR_RESPONSAVEL;
    FieldByName('ENDERECO_RESPONSAVEL').AsString := FENDERECO_RESPONSAVEL;
    FieldByName('BAIRRO_RESPONSAVEL').AsString := FBAIRRO_RESPONSAVEL;
    FieldByName('NUMERO_RESPONSAVEL').AsInteger := FNUMERO_RESPONSAVEL;
    FieldByName('COMPLEMENTO_RESPONSAVEL').AsString := FCOMPLEMENTO_RESPONSAVEL;
    FieldByName('CEP_RESPONSAVEL').AsString := FCEP_RESPONSAVEL;
    FieldByName('CIDADE_RESPONSAVEL').AsString := FCIDADE_RESPONSAVEL;
    FieldByName('ESTADO_RESPONSAVEL').AsString := FESTADO_RESPONSAVEL;
    FieldByName('TELEFONE_RESPONSAVEL').AsString := FTELEFONE_RESPONSAVEL;
    FieldByName('CELULAR_RESPONSAVEL').AsString := FCELULAR_RESPONSAVEL;
    FieldByName('WHATSAPP_RESPONSAVEL').AsString := FWHATSAPP_RESPONSAVEL;
    FieldByName('EMAIL_RESPONSAVEL').AsString := FEMAIL_RESPONSAVEL;
    FieldByName('PACOTE_ESCOLHIDO').AsString := FPACOTE_ESCOLHIDO;
    FieldByName('VALOR_TOTAL_PACOTE').AsCurrency := FVALOR_TOTAL_PACOTE;
    FieldByName('VALOR_TOTAL_PACOTE_EXTENSO').AsString := FVALOR_TOTAL_PACOTE_EXTENSO;
    FieldByName('VALOR_PACOTE_DESCONTO').AsCurrency := FVALOR_PACOTE_DESCONTO;
    FieldByName('VALOR_PACOTE_DESCONTO_EXTENSO').AsString := FVALOR_PACOTE_DESCONTO_EXTENSO;
    FieldByName('VALOR_PARCELA').AsCurrency := FVALOR_PARCELA;
    FieldByName('VALOR_PARCELA_EXTENSO').AsString := FVALOR_PARCELA_EXTENSO;
    FieldByName('VALOR_PARCELA_DESCONTO').AsString := '0';
    FieldByName('VALOR_PARCELA_DESCONTO_EXTENSO').AsString := '';
    FieldByName('QUANTIDADE_PARCELAS').AsInteger := FQUANTIDADE_PARCELAS;
    FieldByName('QUANTIDADE_PARCELAS_EXTENSO').AsString := FQUANTIDADE_PARCELAS_EXTENSO;
    FieldByName('DATA_INICIO_EXTENSO').AsString := FDATA_INICIO_EXTENSO;
    FieldByName('DATA_TEMINO_EXTENSO').AsString := FDATA_TEMINO_EXTENSO;
    FieldByName('VENCIMENTO_A_PARTIR_EXTENSO').AsString := FVENCIMENTO_A_PARTIR_EXTENSO;
    FieldByName('DURACAO_MEDIA_CURSO').AsFloat := FDURACAO_MEDIA_CURSO;
    FieldByName('DURACAO_MEDIA_CURSO_EXTENSO').AsString := FDURACAO_MEDIA_CURSO_EXTENSO;
    FieldByName('SITUACAO_MATRICULA').AsString := FSITUACAO_MATRICULA;
    FieldByName('ID_CONSULTOR').AsInteger := FID_CONSULTOR;
    FieldByName('CONSULTOR').AsString := FCONSULTOR;
    FieldByName('ID_FUNCIONARIO').AsInteger := idFuncionario;
    FieldByName('NOME_FUNCIONARIO').AsString := nomeFuncionario;

    FieldByName('DIAS_AULAS').AsString := FDIAS_AULAS;
    FieldByName('HORARIOS_AULAS').AsString := FHORARIOS_AULAS;
    FieldByName('SALA').AsString := FSALA;
    FieldByName('VALOR_MATRICULA').AsCurrency := FVALOR_MATRICULA;
    FieldByName('VALOR_MATRICULA_EXTENSO').AsString := FVALOR_MATRICULA_EXTENSO;
    FieldByName('VENCIMENTO_MATRICULA').AsDateTime := FVENCIMENTO_MATRICULA;
    FieldByName('VENCIMENTO_MATRICULA_EXTENSO').AsString := FVENCIMENTO_MATRICULA_EXTENSO;
    FieldByName('MODULOS').AsString := FMODULOS;

    FieldByName('DIA_HORARIO').AsString := FDIA_HORARIO;
    FieldByName('DIA_HORARIO_TURMA').AsString := FDIA_HORARIO_TURMA;
    FieldByName('DIA_HORARIO_TURMA_SALA').AsString := FDIA_HORARIO_TURMA_SALA;

    FieldByName('NACIONALIDADE_ALUNO').AsString := FNACIONALIDADE_ALUNO;
    FieldByName('NACIONALIDADE_RESPONSAVEL').AsString := FNACIONALIDADE_RESPONSAVEL;

    FieldByName('QTDEDIASAULAS').AsInteger := FQtdeDiasAulas;
    FieldByName('QTDEHORASAULAS').AsInteger := FQtdeHorasAulas;

    FieldByName('DIA_VENCIMENTO').AsString := FDia_Vencimento;

    FieldByName('MATERIAIS').AsString := FMATERIAIS;
    FieldByName('VALOR_MATERIAIS').AsCurrency := StrToCurr(FVALOR_MATERIAIS);
    FieldByName('QUANTIDADE_MATERIAIS').AsInteger := StrToInt(FQUANTIDADE_MATERIAIS);

    if FDATA_NASCIMENTO <> StrToDate('30/12/1899') then
      FieldByName('DATA_NASCIMENTO').AsDateTime := FDATA_NASCIMENTO;

    if FDATA_CADASTRO <> StrToDate('30/12/1899') then
      FieldByName('DATA_CADASTRO').AsDateTime := FDATA_CADASTRO;

    if FDATA_MATRICULA <> StrToDate('30/12/1899') then
      FieldByName('DATA_MATRICULA').AsDateTime := FDATA_MATRICULA;

    if FDATA_NASCIMENTO_RESPONSAVEL <> EmptyStr then
      FieldByName('DATA_NASCIMENTO_RESPONSAVEL').AsDateTime := StrToDate(FDATA_NASCIMENTO_RESPONSAVEL);

    if FDATA_CADASTRO_RESPONSAVEL <> EmptyStr then
      FieldByName('DATA_CADASTRO_RESPONSAVEL').AsDateTime := StrToDate(FDATA_CADASTRO_RESPONSAVEL);

    if FDATA_INICIO <> StrToDate('30/12/1899') then
      FieldByName('DATA_INICIO').AsDateTime := FDATA_INICIO;

    if FDATA_TEMINO <> StrToDate('30/12/1899') then
      FieldByName('DATA_TEMINO').AsDateTime := FDATA_TEMINO;

    if FVENCIMENTO_A_PARTIR <> StrToDate('30/12/1899') then
      FieldByName('VENCIMENTO_A_PARTIR').AsDateTime := FVENCIMENTO_A_PARTIR;

  end;

  FQuery.TQuery.Post;

  if FQuery.TQuery.MassiveCount > 0 then
  begin

    if FQuery.TQuery.ApplyUpdates(sErro) then
      showmessage('Operação realizada com sucesso!!!')
    else
      raise Exception.create
        ('Erro ao tentar gravar os dados na tabela. ' + sErro);
  end
  else
  begin
    showmessage('Operação realizada com sucesso!!!')
  end;

end;

function TClasseContratoAlunos.inserir: iCadastroGerarContrado;
begin

  result := self;
  FQuery.TQuery.Append;

end;

function TClasseContratoAlunos.listarGrid(value: TDataSource): iCadastroGerarContrado;
begin

  result := self;
  value.DataSet := FQuery.TQuery;

end;

class function TClasseContratoAlunos.new: iCadastroGerarContrado;
begin
  result := self.create;
end;

function TClasseContratoAlunos.ordenarGrid(column: TColumn): iCadastroGerarContrado;
begin
  if FQuery.TQuery.IndexFieldNames = column.FieldName then
    FQuery.TQuery.IndexFieldNames := column.FieldName + ':D'
  else
    FQuery.TQuery.IndexFieldNames := column.FieldName;
end;

function TClasseContratoAlunos.returnClientDataSet: TClientDataSet;
begin

end;

function TClasseContratoAlunos.sqlPesquisa: iCadastroGerarContrado;
begin
  result := self;
  FQuery.getCampo(FCampo).getValor(FValor).sqlPesquisar(FTabela);
end;

function TClasseContratoAlunos.sqlPesquisaData: iCadastroGerarContrado;
begin
  result := self;
  FQuery.getCampo(FCampo).getValor(FValor).getDataInicial(FDataInicial)
    .getDataFinal(FDataFinal).sqlPesquisarData(FTabela);
end;

function TClasseContratoAlunos.sqlPesquisaEstatica: iCadastroGerarContrado;
begin
  result := self;
  FQuery.getCampo(FCampo).getValor(FValor).getDataInicial(FDataInicial)
    .getDataFinal(FDataFinal).sqlPesquisaEstatica(FTabela);
end;

procedure TClasseContratoAlunos.SubstituirTexto(TextoAntigo, TextoNovo: string; Rich: TRxRichEdit);
var
  Posicao: integer;
begin
  Posicao := 1;

  With Rich do
    repeat


      Posicao := FindText(TextoAntigo, 0, Length(Text), []);

      if Posicao > 0 then
      begin
        SelStart := Posicao;
        SelLength := Length(TextoAntigo);
        SelText := TextoNovo;

        SetFocus;
        Posicao := 0;
      end;
    until (Posicao <> 0);
    
end;

function TClasseContratoAlunos.validarData(componet: TMaskEdit)
  : iCadastroGerarContrado;
var
  d: TDate;
begin

  result := self;

  try
    d := StrToDate(componet.Text);
  except
    componet.SetFocus;
    componet.Clear;
    raise Exception.create('Digite uma data valida.');
  end;

end;

function TClasseContratoAlunos.verificarContrato(value: integer): boolean;
begin

  result := false;

  FQuery.getCampo('ID_MATRICULA').getValor(value.ToString).sqlPesquisaEstatica(FTabela);

  if FQuery.TQuery.RecordCount >= 1 then
    result := true
  else
    result := false;
end;

function TClasseContratoAlunos.gerarContrato( value: TRxRichEdit): iCadastroGerarContrado;
begin

  result := self;

  SubstituirTexto('<NUMERO_CONTRATO>', FNUMERO_CONTRATO, value);
  SubstituirTexto('<NOME_ALUNO>', FNOME_ALUNO, value);
  SubstituirTexto('<COD_ALUNO>', FID_ALUNO.ToString, value);
  SubstituirTexto('<MATRICULA>', FID_Matricula.ToString, value);
  SubstituirTexto('<DATA_NASCIMENTO>', DateToStr(FDATA_NASCIMENTO), value);
  SubstituirTexto('<DATA_CADASTRO>',  DateToStr(FDATA_CADASTRO), value);
  SubstituirTexto('<DATA_MATRICULA>', DateToStr(FDATA_MATRICULA), value);
  SubstituirTexto('<CPF>', FCPF, value);
  SubstituirTexto('<RG>', FRG, value);
  SubstituirTexto('<ORGAO_EXPEDIDOR>', FORGAO_EXPEDIDOR, value);
  SubstituirTexto('<NACIONALIDADE_ALUNO>', FNACIONALIDADE_ALUNO, value);
  SubstituirTexto('<NACIONALIDADE>', FNACIONALIDADE, value);
  SubstituirTexto('<ENDERECO>', FENDERECO, value);
  SubstituirTexto('<BAIRRO>', FBAIRRO, value);
  SubstituirTexto('<NUMERO>', FNUMERO.ToString, value);
  SubstituirTexto('<COMPLEMENTO>', FCOMPLEMENTO, value);
  SubstituirTexto('<CEP>', FCEP, value);
  SubstituirTexto('<CIDADE>', FCIDADE, value);
  SubstituirTexto('<ESTADO>', FESTADO, value);
  SubstituirTexto('<TELEFONE>', FTELEFONE, value);
  SubstituirTexto('<CELULAR>', FCELULAR, value);
  SubstituirTexto('<WHATSAPP>', FWHATSAPP, value);
  SubstituirTexto('<EMAIL>', FEMAIL, value);
  SubstituirTexto('<TIPO_MARKETING>', FTIPO_MARKETING, value);
  SubstituirTexto('<NOME_MAE>', FNOME_MAE, value);
  SubstituirTexto('<NOME_PAI>', FNOME_PAI, value);
  SubstituirTexto('<NOME_RESPONSAVEL>', FNOME_RESPONSAVEL, value);
  SubstituirTexto('<DATA_NASCIMENTO_RESPONSAVEL>', FDATA_NASCIMENTO_RESPONSAVEL, value);
  SubstituirTexto('<DATA_CADASTRO_RESPONSAVEL>', FDATA_CADASTRO_RESPONSAVEL, value);
  SubstituirTexto('<CPF_RESPONSAVEL>', FCPF_RESPONSAVEL, value);
  SubstituirTexto('<RG_RESPONSAVEL>', FRG_RESPONSAVEL, value);
  SubstituirTexto('<ORGAO_EXPEDIDOR_RESPONSAVEL>', FORGAO_EXPEDIDOR, value);
  SubstituirTexto('<NACIONALIDADE_RESPONSAVEL>', FNACIONALIDADE_RESPONSAVEL, value);
  SubstituirTexto('<ENDERECO_RESPONSAVEL>', FENDERECO_RESPONSAVEL, value);
  SubstituirTexto('<BAIRRO_RESPONSAVEL>', FBAIRRO_RESPONSAVEL, value);
  SubstituirTexto('<NUMERO_RESPONSAVEL>', FNUMERO_RESPONSAVEL.ToString, value);
  SubstituirTexto('<COMPLEMENTO_RESPONSAVEL>', FCOMPLEMENTO_RESPONSAVEL, value);
  SubstituirTexto('<CEP_RESPONSAVEL>', FCEP_RESPONSAVEL, value);
  SubstituirTexto('<CIDADE_RESPONSAVEL>', FCIDADE_RESPONSAVEL, value);
  SubstituirTexto('<ESTADO_RESPONSAVEL>', FESTADO_RESPONSAVEL, value);
  SubstituirTexto('<TELEFONE_RESPONSAVEL>', FTELEFONE_RESPONSAVEL, value);
  SubstituirTexto('<CELULAR_RESPONSAVEL>', FCELULAR_RESPONSAVEL, value);
  SubstituirTexto('<WHATSAPP_RESPONSAVEL>', FWHATSAPP_RESPONSAVEL, value);
  SubstituirTexto('<EMAIL_RESPONSAVEL>', FEMAIL_RESPONSAVEL, value);
  SubstituirTexto('<PACOTE_ESCOLHIDO>', FPACOTE_ESCOLHIDO, value);
  SubstituirTexto('<VALOR_TOTAL_PACOTE>', FormatFloat('R$ #, ##0.00', FVALOR_TOTAL_PACOTE), value);
  SubstituirTexto('<VALOR_TOTAL_PACOTE_EXTENSO>', FVALOR_TOTAL_PACOTE_EXTENSO, value);
  SubstituirTexto('<VALOR_PACOTE_DESCONTO>', FormatFloat('R$ #, ##0.00', FVALOR_PACOTE_DESCONTO), value);
  SubstituirTexto('<VALOR_PACOTE_DESCONTO_EXTENSO>', FVALOR_PACOTE_DESCONTO_EXTENSO, value);
  SubstituirTexto('<VALOR_PARCELA>', FormatFloat('R$ #, ##0.00', FVALOR_PARCELA), value);
  SubstituirTexto('<VALOR_PARCELA_EXTENSO>', FVALOR_PARCELA_EXTENSO, value);
  SubstituirTexto('<VALOR_PARCELA_DESCONTO>', FormatFloat('R$ #, ##0.00', FVALOR_PARCELA_DESCONTO), value);
  SubstituirTexto('<VALOR_PARCELA_DESCONTO_EXTENSO>', FVALOR_PARCELA_DESCONTO_EXTENSO, value);
  SubstituirTexto('<QUANTIDADE_PARCELAS>', FQUANTIDADE_PARCELAS.ToString, value);
  SubstituirTexto('<QUANTIDADE_PARCELAS_EXTENSO>', FQUANTIDADE_PARCELAS_EXTENSO, value);
  SubstituirTexto('<DATA_INICIO>', DateToStr(FDATA_INICIO), value);
  SubstituirTexto('<DATA_INICIO_EXTENSO>', FDATA_INICIO_EXTENSO, value);
  SubstituirTexto('<DATA_TERMINO>', DateToStr(FDATA_TEMINO), value);
  SubstituirTexto('<DATA_TERMINO_EXTENSO>', FDATA_TEMINO_EXTENSO, value);
  SubstituirTexto('<VENCIMENTO_A_PARTIR>', DateToStr(FVENCIMENTO_A_PARTIR), value);
  SubstituirTexto('<VENCIMENTO_A_PARTIR_EXTENSO>', FVENCIMENTO_A_PARTIR_EXTENSO, value);
  SubstituirTexto('<DURACAO_MEDIA_CURSO>', FloatToStr(FDURACAO_MEDIA_CURSO), value);
  SubstituirTexto('<DURACAO_MEDIA_CURSO_EXTENSO>', FDURACAO_MEDIA_CURSO_EXTENSO, value);
  SubstituirTexto('<SITUACAO_MATRICULA>', FSITUACAO_MATRICULA, value);
  SubstituirTexto('<ID_CONSULTOR>', FID_CONSULTOR.ToString, value);
  SubstituirTexto('<CONSULTOR>', FCONSULTOR, value);
  SubstituirTexto('<ID_FUNCIONARIO>', idfuncionario.ToString, value);
  SubstituirTexto('<NOME_FUNCIONARIO>', nomeFuncionario, value);
  SubstituirTexto('<DIAS_AULAS>', FDIAS_AULAS, value);
  SubstituirTexto('<HORARIOS_AULAS>', FHORARIOS_AULAS, value);
  SubstituirTexto('<SALA>', FSALA, value);
  SubstituirTexto('<VALOR_MATRICULA>', FormatFloat('R$ #, ##0.00',FVALOR_MATRICULA), value);
  SubstituirTexto('<VALOR_MATRICULA_EXTENSO>', FVALOR_MATRICULA_EXTENSO, value);
  SubstituirTexto('<VENCIMENTO_MATRICULA>', DateToStr(FVENCIMENTO_MATRICULA), value);
  SubstituirTexto('<VENCIMENTO_MATRICULA_EXTENSO>', FVENCIMENTO_MATRICULA_EXTENSO, value);
  SubstituirTexto('<MODULOS>', FMODULOS, value);
  SubstituirTexto('<DIA_HORARIO>', FDIA_HORARIO, value);
  SubstituirTexto('<DIA_HORARIO_TURMA>', FDIA_HORARIO_TURMA, value);
  SubstituirTexto('<DIA_HORARIO_TURMA_SALA>', FDIA_HORARIO_TURMA_SALA, value);
  SubstituirTexto('<DIA_VENCIMENTO> ', FDia_Vencimento, value);
  SubstituirTexto('<QTDEDIASAULAS> ', FQTDEDIASAULAS.ToString, value);
  SubstituirTexto('<QTDEHORASAULAS> ', FQTDEHORASAULAS.ToString, value);
  SubstituirTexto('<NOME_FANTASIA>', FNOME_FANTASIA, value);
  SubstituirTexto('<RAZAO_SOCIAL>', FRAZAO_SOCIAL, value);
  SubstituirTexto('<CNPJ>', FCNPJ, value);
  SubstituirTexto('<INSCRICAO_ESTADUAL>', FINSCRICAO_ESTADUAL, value);
  SubstituirTexto('<DATA_ABERTURA>', FDATA_ABERTURA, value);
  SubstituirTexto('<SOCIOS>', FSOCIOS, value);
  SubstituirTexto('<ENDERECO_EMPRESA>', FENDERECO_EMPRESA, value);
  SubstituirTexto('<BAIRRO_EMPRESA>', FBAIRRO_EMPRESA, value);
  SubstituirTexto('<NUMERO_EMPRESA>', IntToStr(FNUMERO_EMPRESA), value);
  SubstituirTexto('<COMPLEMENTO_EMPRESA>', FCOMPLEMENTO_EMPRESA, value);
  SubstituirTexto('<CEP_EMPRESA>', FCEP_EMPRESA, value);
  SubstituirTexto('<CIDADE_EMPRESA>', FCIDADE_EMPRESA, value);
  SubstituirTexto('<ESTADO_EMPRESA>', FESTADO_EMPRESA, value);
  SubstituirTexto('<TELEFONE_EMPRESA>', FTELEFONE_EMPRESA, value);
  SubstituirTexto('<CELULAR_EMPRESA>', FCELULAR_EMPRESA, value);
  SubstituirTexto('<WHATSAPP_EMPRESA>', FWHATSAPP_EMPRESA, value);
  SubstituirTexto('<EMAIL_EMPRESA>', FEMAIL_EMPRESA, value);
  SubstituirTexto('<data_contrato>', FormatDateTime('dd/mmmm/yyyy', date), value);
  SubstituirTexto('<MATERIAIS>', FMATERIAIS, value);
  SubstituirTexto('<VALOR_MATERIAIS>', FVALOR_MATERIAIS, value);
  SubstituirTexto('<QUANTIDADE_MATERIAIS>', FQUANTIDADE_MATERIAIS, value);
//  FindReplace('', , value);


end;

function TClasseContratoAlunos.gerarContratoWord(value: string): iCadastroGerarContrado;
type
  twordreplaceflags = set of (wrfreplaceall, wrfmatchcase, wrfmatchwildcards);
const
  wdfindcontinue = 1;
  wdreplaceone = 1;
  wdreplaceall = 2;
  wddonotsavechanges = 0;
var
  wordapp: olevariant;
begin
  if not fileexists(ExtractFilePath(ExtractFilePath(application.exeName)) +
    '\modelo contrato\' + value) then
    raise exception.create('Arquivo não encontrato:' + extractfiledir(application.exeName)
      + '\modelo contrato\' + value);

  try
    wordapp := createoleobject('word.application');
    wordapp.Caption :=  FID_MATRICULA.ToString + '- ' + FNOME_ALUNO;
  //  wordapp.title := frmAlunos.edtMatricula.Text +'- '+frmAlunos.edtNomeAluno.Text;

  except
    messagedlg('Microsoft Word não encontrato.', mterror, [mbok], 0);
  end;

    if fileexists(ExtractFilePath(ExtractFilePath(application.exeName)) +
      '\Contratos alunos\' + value) then
    begin
      FDestino2 := (ExtractFilePath(ExtractFilePath(application.exeName)) +
        '\Contratos alunos\' + value);
      deletefile(pwidechar(FDestino2));
    end;

    FLocal := (ExtractFilePath(ExtractFilePath(application.exeName)) +
      '\modelo contrato\' + value);

    FDestino1 := (ExtractFilePath(ExtractFilePath(application.exeName)) +
      '\Contratos alunos\' + value);
    copyfile(pchar(FLocal), pchar(FDestino1), false);

    wordapp.visible := true;
    wordapp.documents.open(ExtractFilePath(ExtractFilePath(application.exeName))
      + '\Contratos alunos\' + value);
    wordapp.selection.find.clearformatting;
    wordapp.selection.find.forward := true;

    wordapp.selection.find.wrap := wdfindcontinue;
    wordapp.selection.find.format := false;
    wordapp.selection.find.matchwholeword := true;
    wordapp.selection.find.matchsoundslike := false;
    wordapp.selection.find.matchallwordforms := false;

    // registro do cliente

    wordapp.selection.find.text := '<NUMERO_CONTRATO>';
    wordapp.selection.find.replacement.text := FNUMERO_CONTRATO;
    wordapp.selection.find.execute(replace := wdreplaceall);

    wordapp.selection.find.text := '<NOME_ALUNO>';
    wordapp.selection.find.replacement.text := FNOME_ALUNO;
    wordapp.selection.find.execute(replace := wdreplaceall);

    wordapp.selection.find.text := '<COD_ALUNO>';
    wordapp.selection.find.replacement.text := FID_ALUNO.ToString;
    wordapp.selection.find.execute(replace := wdreplaceall);

    wordapp.selection.find.text := '<MATRICULA>';
    wordapp.selection.find.replacement.text := FID_MATRICULA.ToString;
    wordapp.selection.find.execute(replace := wdreplaceall);

    wordapp.selection.find.text := '<DATA_NASCIMENTO>';
    wordapp.selection.find.replacement.text := DateToStr(FDATA_NASCIMENTO);
    wordapp.selection.find.execute(replace := wdreplaceall);

    wordapp.selection.find.text := '<DATA_CADASTRO>';
    wordapp.selection.find.replacement.text := DateToStr(FDATA_CADASTRO);
    wordapp.selection.find.execute(replace := wdreplaceall);

    wordapp.selection.find.text := '<DATA_MATRICULA>';
    wordapp.selection.find.replacement.text := DateToStr(FDATA_MATRICULA);
    wordapp.selection.find.execute(replace := wdreplaceall);

    wordapp.selection.find.text := '<CPF>';
    wordapp.selection.find.replacement.text := FCPF;
    wordapp.selection.find.execute(replace := wdreplaceall);

    wordapp.selection.find.text := '<RG>';
    wordapp.selection.find.replacement.text := FRG;
    wordapp.selection.find.execute(replace := wdreplaceall);

    wordapp.selection.find.text := '<ORGAO_EXPEDIDOR>';
    wordapp.selection.find.replacement.text := FORGAO_EXPEDIDOR;
    wordapp.selection.find.execute(replace := wdreplaceall);

    wordapp.selection.find.text := '<NACIONALIDADE_ALUNO>';
    wordapp.selection.find.replacement.text := FNACIONALIDADE_ALUNO;
    wordapp.selection.find.execute(replace := wdreplaceall);

    wordapp.selection.find.text := '<ENDERECO>';
    wordapp.selection.find.replacement.text := FENDERECO;
    wordapp.selection.find.execute(replace := wdreplaceall);

    wordapp.selection.find.text := '<BAIRRO>';
    wordapp.selection.find.replacement.text := FBAIRRO;
    wordapp.selection.find.execute(replace := wdreplaceall);

    wordapp.selection.find.text := '<NUMERO>';
    wordapp.selection.find.replacement.text := FNUMERO.ToString;
    wordapp.selection.find.execute(replace := wdreplaceall);

    wordapp.selection.find.text := '<COMPLEMENTO>';
    wordapp.selection.find.replacement.text := FCOMPLEMENTO;
    wordapp.selection.find.execute(replace := wdreplaceall);

    wordapp.selection.find.text := '<CEP>';
    wordapp.selection.find.replacement.text := FCEP;
    wordapp.selection.find.execute(replace := wdreplaceall);

    wordapp.selection.find.text := '<CIDADE>';
    wordapp.selection.find.replacement.text := FCIDADE;
    wordapp.selection.find.execute(replace := wdreplaceall);

    wordapp.selection.find.text := '<ESTADO>';
    wordapp.selection.find.replacement.text := FESTADO;
    wordapp.selection.find.execute(replace := wdreplaceall);

    wordapp.selection.find.text := '<TELEFONE>';
    wordapp.selection.find.replacement.text := FTELEFONE;
    wordapp.selection.find.execute(replace := wdreplaceall);

    wordapp.selection.find.text := '<CELULAR>';
    wordapp.selection.find.replacement.text := FCELULAR;
    wordapp.selection.find.execute(replace := wdreplaceall);

    wordapp.selection.find.text := '<WHATSAPP>';
    wordapp.selection.find.replacement.text := FWHATSAPP;
    wordapp.selection.find.execute(replace := wdreplaceall);


    wordapp.selection.find.text := '<EMAIL>';
    wordapp.selection.find.replacement.text := FEMAIL;
    wordapp.selection.find.execute(replace := wdreplaceall);

    wordapp.selection.find.text := '<TIPO_MARKETING>';
    wordapp.selection.find.replacement.text := FTIPO_MARKETING;
    wordapp.selection.find.execute(replace := wdreplaceall);


    wordapp.selection.find.text := '<NOME_MAE>';
    wordapp.selection.find.replacement.text := FNOME_MAE;
    wordapp.selection.find.execute(replace := wdreplaceall);

    wordapp.selection.find.text := '<NOME_PAI>';
    wordapp.selection.find.replacement.text := FNOME_PAI;
    wordapp.selection.find.execute(replace := wdreplaceall);


    wordapp.selection.find.text := '<NOME_RESPONSAVEL>';
    wordapp.selection.find.replacement.text := FNOME_RESPONSAVEL;
    wordapp.selection.find.execute(replace := wdreplaceall);

    wordapp.selection.find.text := '<DATA_NASCIMENTO_RESPONSAVEL>';
    wordapp.selection.find.replacement.text := FDATA_NASCIMENTO_RESPONSAVEL;
    wordapp.selection.find.execute(replace := wdreplaceall);


    wordapp.selection.find.text := '<DATA_CADASTRO_RESPONSAVEL>';
    wordapp.selection.find.replacement.text := FDATA_CADASTRO_RESPONSAVEL;
    wordapp.selection.find.execute(replace := wdreplaceall);

    wordapp.selection.find.text := '<CPF_RESPONSAVEL>';
    wordapp.selection.find.replacement.text := FCPF_RESPONSAVEL;
    wordapp.selection.find.execute(replace := wdreplaceall);


    wordapp.selection.find.text := '<RG_RESPONSAVEL>';
    wordapp.selection.find.replacement.text := FRG_RESPONSAVEL;
    wordapp.selection.find.execute(replace := wdreplaceall);

    wordapp.selection.find.text := '<ORGAO_EXPEDIDOR_RESPONSAVEL>';
    wordapp.selection.find.replacement.text := FORGAO_EXPEDIDOR_RESPONSAVEL;
    wordapp.selection.find.execute(replace := wdreplaceall);


    wordapp.selection.find.text := '<ENDERECO_RESPONSAVEL>';
    wordapp.selection.find.replacement.text := FENDERECO_RESPONSAVEL;
    wordapp.selection.find.execute(replace := wdreplaceall);

    wordapp.selection.find.text := '<BAIRRO_RESPONSAVEL>';
    wordapp.selection.find.replacement.text := FBAIRRO_RESPONSAVEL;
    wordapp.selection.find.execute(replace := wdreplaceall);

    wordapp.selection.find.text := '<NUMERO_RESPONSAVEL>';
    wordapp.selection.find.replacement.text := FNUMERO_RESPONSAVEL.ToString;
    wordapp.selection.find.execute(replace := wdreplaceall);

    wordapp.selection.find.text := '<COMPLEMENTO_RESPONSAVEL>';
    wordapp.selection.find.replacement.text := FCOMPLEMENTO_RESPONSAVEL;
    wordapp.selection.find.execute(replace := wdreplaceall);


    wordapp.selection.find.text := '<CEP_RESPONSAVEL>';
    wordapp.selection.find.replacement.text := FCEP_RESPONSAVEL;
    wordapp.selection.find.execute(replace := wdreplaceall);

    wordapp.selection.find.text := '<CIDADE_RESPONSAVEL>';
    wordapp.selection.find.replacement.text := FCIDADE_RESPONSAVEL;
    wordapp.selection.find.execute(replace := wdreplaceall);


    wordapp.selection.find.text := '<ESTADO_RESPONSAVEL>';
    wordapp.selection.find.replacement.text := FESTADO_RESPONSAVEL;
    wordapp.selection.find.execute(replace := wdreplaceall);

    wordapp.selection.find.text := '<TELEFONE_RESPONSAVEL>';
    wordapp.selection.find.replacement.text := FTELEFONE_RESPONSAVEL;
    wordapp.selection.find.execute(replace := wdreplaceall);

    wordapp.selection.find.text := '<CELULAR_RESPONSAVEL>';
    wordapp.selection.find.replacement.text := FCELULAR_RESPONSAVEL;
    wordapp.selection.find.execute(replace := wdreplaceall);

    wordapp.selection.find.text := '<WHATSAPP_RESPONSAVEL>';
    wordapp.selection.find.replacement.text := FWHATSAPP_RESPONSAVEL;
    wordapp.selection.find.execute(replace := wdreplaceall);

    wordapp.selection.find.text := '<EMAIL_RESPONSAVEL>';
    wordapp.selection.find.replacement.text := FEMAIL_RESPONSAVEL;
    wordapp.selection.find.execute(replace := wdreplaceall);

    wordapp.selection.find.text := '<PACOTE_ESCOLHIDO>';
    wordapp.selection.find.replacement.text := FPACOTE_ESCOLHIDO;
    wordapp.selection.find.execute(replace := wdreplaceall);

    wordapp.selection.find.text := '<VALOR_TOTAL_PACOTE>';
    wordapp.selection.find.replacement.text := CurrToStr(FVALOR_TOTAL_PACOTE);
    wordapp.selection.find.execute(replace := wdreplaceall);

    wordapp.selection.find.text := '<VALOR_TOTAL_PACOTE_EXTENSO>';
    wordapp.selection.find.replacement.text := FVALOR_TOTAL_PACOTE_EXTENSO;
    wordapp.selection.find.execute(replace := wdreplaceall);

    wordapp.selection.find.text := '<VALOR_PACOTE_DESCONTO>';
    wordapp.selection.find.replacement.text := CurrToStr(FVALOR_PACOTE_DESCONTO);
    wordapp.selection.find.execute(replace := wdreplaceall);

    wordapp.selection.find.text := '<VALOR_PACOTE_DESCONTO_EXTENSO>';
    wordapp.selection.find.replacement.text := FVALOR_PACOTE_DESCONTO_EXTENSO;
    wordapp.selection.find.execute(replace := wdreplaceall);

    wordapp.selection.find.text := '<VALOR_PARCELA>';
    wordapp.selection.find.replacement.text := CurrToStr(FVALOR_PARCELA);
    wordapp.selection.find.execute(replace := wdreplaceall);

    wordapp.selection.find.text := '<VALOR_PARCELA_EXTENSO>';
    wordapp.selection.find.replacement.text := FVALOR_PARCELA_EXTENSO;
    wordapp.selection.find.execute(replace := wdreplaceall);

    wordapp.selection.find.text := '<VALOR_PARCELA_DESCONTO>';
    wordapp.selection.find.replacement.text := CurrToStr(FVALOR_PARCELA_DESCONTO);
    wordapp.selection.find.execute(replace := wdreplaceall);

    wordapp.selection.find.text := '<VALOR_PARCELA_DESCONTO_EXTENSO>';
    wordapp.selection.find.replacement.text := FVALOR_PARCELA_DESCONTO_EXTENSO;
    wordapp.selection.find.execute(replace := wdreplaceall);

    wordapp.selection.find.text := '<QUANTIDADE_PARCELAS>';
    wordapp.selection.find.replacement.text := FQUANTIDADE_PARCELAS.ToString;;
    wordapp.selection.find.execute(replace := wdreplaceall);

    wordapp.selection.find.text := '<QUANTIDADE_PARCELAS_EXTENSO>';
    wordapp.selection.find.replacement.text := FQUANTIDADE_PARCELAS_EXTENSO;
    wordapp.selection.find.execute(replace := wdreplaceall);

    wordapp.selection.find.text := '<DATA_INICIO>';
    wordapp.selection.find.replacement.text := DateToStr(FDATA_INICIO);
    wordapp.selection.find.execute(replace := wdreplaceall);

    wordapp.selection.find.text := '<DATA_INICIO_EXTENSO>';
    wordapp.selection.find.replacement.text := FDATA_INICIO_EXTENSO;
    wordapp.selection.find.execute(replace := wdreplaceall);

    wordapp.selection.find.text := '<DATA_TERMINO>';
    wordapp.selection.find.replacement.text := DateToStr(FDATA_TEMINO);
    wordapp.selection.find.execute(replace := wdreplaceall);

    wordapp.selection.find.text := '<DATA_TERMINO_EXTENSO>';
    wordapp.selection.find.replacement.text := FDATA_TEMINO_EXTENSO;
    wordapp.selection.find.execute(replace := wdreplaceall);

    wordapp.selection.find.text := '<VENCIMENTO_A_PARTIR>';
    wordapp.selection.find.replacement.text := DateToStr(FVENCIMENTO_A_PARTIR);
    wordapp.selection.find.execute(replace := wdreplaceall);

    wordapp.selection.find.text := '<VENCIMENTO_A_PARTIR_EXTENSO>';
    wordapp.selection.find.replacement.text := FVENCIMENTO_A_PARTIR_EXTENSO;
    wordapp.selection.find.execute(replace := wdreplaceall);

    wordapp.selection.find.text := '<DIA_VENCIMENTO>';
    wordapp.selection.find.replacement.text := FDia_Vencimento;
    wordapp.selection.find.execute(replace := wdreplaceall);

    wordapp.selection.find.text := '<DURACAO_MEDIA_CURSO>';
    wordapp.selection.find.replacement.text := FloatToStr(FDURACAO_MEDIA_CURSO);
    wordapp.selection.find.execute(replace := wdreplaceall);

    wordapp.selection.find.text := '<DURACAO_MEDIA_CURSO_EXTENSO>';
    wordapp.selection.find.replacement.text := FDURACAO_MEDIA_CURSO_EXTENSO;
    wordapp.selection.find.execute(replace := wdreplaceall);

    wordapp.selection.find.text := '<SITUACAO_MATRICULA>';
    wordapp.selection.find.replacement.text := FSITUACAO_MATRICULA;
    wordapp.selection.find.execute(replace := wdreplaceall);

    wordapp.selection.find.text := '<ID_CONSULTOR>';
    wordapp.selection.find.replacement.text := FID_CONSULTOR.ToString;
    wordapp.selection.find.execute(replace := wdreplaceall);

    wordapp.selection.find.text := '<CONSULTOR>';
    wordapp.selection.find.replacement.text := FCONSULTOR;
    wordapp.selection.find.execute(replace := wdreplaceall);

    wordapp.selection.find.text := '<ID_FUNCIONARIO>';
    wordapp.selection.find.replacement.text := FID_FUNCIONARIO.ToString;
    wordapp.selection.find.execute(replace := wdreplaceall);

    wordapp.selection.find.text := '<NOME_FUNCIONARIO>';
    wordapp.selection.find.replacement.text := FNOME_FUNCIONARIO;
    wordapp.selection.find.execute(replace := wdreplaceall);

    wordapp.selection.find.text := '<DIAS_AULAS>';
    wordapp.selection.find.replacement.text := FDIAS_AULAS;
    wordapp.selection.find.execute(replace := wdreplaceall);

    wordapp.selection.find.text := '<HORARIOS_AULAS>';
    wordapp.selection.find.replacement.text := FHORARIOS_AULAS;
    wordapp.selection.find.execute(replace := wdreplaceall);

    wordapp.selection.find.text := '<SALA>';
    wordapp.selection.find.replacement.text := FSALA;
    wordapp.selection.find.execute(replace := wdreplaceall);

    wordapp.selection.find.text := '<VALOR_MATRICULA>';
    wordapp.selection.find.replacement.text := CurrToStr(FVALOR_MATRICULA);
    wordapp.selection.find.execute(replace := wdreplaceall);

    wordapp.selection.find.text := '<VALOR_MATRICULA_EXTENSO>';
    wordapp.selection.find.replacement.text := FVALOR_MATRICULA_EXTENSO;
    wordapp.selection.find.execute(replace := wdreplaceall);

    wordapp.selection.find.text := '<VENCIMENTO_MATRICULA>';
    wordapp.selection.find.replacement.text := DateToStr(FVENCIMENTO_MATRICULA) ;
    wordapp.selection.find.execute(replace := wdreplaceall);

    wordapp.selection.find.text := '<VENCIMENTO_MATRICULA_EXTENSO>';
    wordapp.selection.find.replacement.text := FVENCIMENTO_MATRICULA_EXTENSO;
    wordapp.selection.find.execute(replace := wdreplaceall);

    wordapp.selection.find.text := '<MODULOS>';
    wordapp.selection.find.replacement.text := FMODULOS;
    wordapp.selection.find.execute(replace := wdreplaceall);

    wordapp.selection.find.text := '<DIA_HORARIO>';
    wordapp.selection.find.replacement.text := FDIA_HORARIO;
    wordapp.selection.find.execute(replace := wdreplaceall);

    wordapp.selection.find.text := '<DIA_HORARIO_TURMA>';
    wordapp.selection.find.replacement.text := FDIA_HORARIO_TURMA;
    wordapp.selection.find.execute(replace := wdreplaceall);

    wordapp.selection.find.text := '<DIA_HORARIO_TURMA_SALA>';
    wordapp.selection.find.replacement.text := FDIA_HORARIO_TURMA_SALA;
    wordapp.selection.find.execute(replace := wdreplaceall);

    wordapp.selection.find.text := '<QTDEDIASAULAS>';
    wordapp.selection.find.replacement.text := FQtdeDiasAulas.ToString;
    wordapp.selection.find.execute(replace := wdreplaceall);

    wordapp.selection.find.text := '<QTDEHORASAULAS>';
    wordapp.selection.find.replacement.text := FQtdeDiasAulas.ToString;
    wordapp.selection.find.execute(replace := wdreplaceall);

    wordapp.selection.find.text := '<NOME_FANTASIA>';
    wordapp.selection.find.replacement.text := FNOME_FANTASIA;
    wordapp.selection.find.execute(replace := wdreplaceall);

     wordapp.selection.find.text := '<RAZAO_SOCIAL>';
    wordapp.selection.find.replacement.text := FRAZAO_SOCIAL;
    wordapp.selection.find.execute(replace := wdreplaceall);

    wordapp.selection.find.text := '<CNPJ>';
    wordapp.selection.find.replacement.text := FCNPJ;
    wordapp.selection.find.execute(replace := wdreplaceall);

    wordapp.selection.find.text := '<INSCRICAO_ESTADUAL>';
    wordapp.selection.find.replacement.text := FINSCRICAO_ESTADUAL;
    wordapp.selection.find.execute(replace := wdreplaceall);

    wordapp.selection.find.text := '<DATA_ABERTURA> ';
    wordapp.selection.find.replacement.text := FDATA_ABERTURA;
    wordapp.selection.find.execute(replace := wdreplaceall);

    wordapp.selection.find.text := '<SOCIOS>';
    wordapp.selection.find.replacement.text := FSOCIOS;
    wordapp.selection.find.execute(replace := wdreplaceall);

    wordapp.selection.find.text := '<ENDERECO_EMPRESA>';
    wordapp.selection.find.replacement.text := FENDERECO_EMPRESA;
    wordapp.selection.find.execute(replace := wdreplaceall);

    wordapp.selection.find.text := '<BAIRRO_EMPRESA>';
    wordapp.selection.find.replacement.text := FBAIRRO_EMPRESA;
    wordapp.selection.find.execute(replace := wdreplaceall);

    wordapp.selection.find.text := '<NUMERO_EMPRESA>';
    wordapp.selection.find.replacement.text := FNUMERO_EMPRESA.ToString;
    wordapp.selection.find.execute(replace := wdreplaceall);

    wordapp.selection.find.text := '<COMPLEMENTO_EMPRESA>';
    wordapp.selection.find.replacement.text := FCOMPLEMENTO_EMPRESA;
    wordapp.selection.find.execute(replace := wdreplaceall);

    wordapp.selection.find.text := '<CEP_EMPRESA>';
    wordapp.selection.find.replacement.text := FCEP_EMPRESA;
    wordapp.selection.find.execute(replace := wdreplaceall);

    wordapp.selection.find.text := '<CIDADE_EMPRESA>';
    wordapp.selection.find.replacement.text := FCIDADE_EMPRESA;
    wordapp.selection.find.execute(replace := wdreplaceall);

    wordapp.selection.find.text := '<ESTADO_EMPRESA>';
    wordapp.selection.find.replacement.text := FESTADO_EMPRESA;
    wordapp.selection.find.execute(replace := wdreplaceall);

    wordapp.selection.find.text := '<TELEFONE_EMPRESA>';
    wordapp.selection.find.replacement.text := FTELEFONE_EMPRESA;
    wordapp.selection.find.execute(replace := wdreplaceall);

    wordapp.selection.find.text := '<CELULAR_EMPRESA>';
    wordapp.selection.find.replacement.text := FCELULAR_EMPRESA;
    wordapp.selection.find.execute(replace := wdreplaceall);

    wordapp.selection.find.text := '<WHATSAPP_EMPRESA>';
    wordapp.selection.find.replacement.text := FWHATSAPP_EMPRESA;
    wordapp.selection.find.execute(replace := wdreplaceall);

    wordapp.selection.find.text := '<EMAIL_EMPRESA>';
    wordapp.selection.find.replacement.text := FEMAIL_EMPRESA;
    wordapp.selection.find.execute(replace := wdreplaceall);

    wordapp.selection.find.text := '<DATA_CONTRATO>';
    wordapp.selection.find.replacement.text := FormatDateTime('dd/mmmm/yyyy', date);
    wordapp.selection.find.execute(replace := wdreplaceall);

    wordapp.selection.find.text := '<MATERIAIS>';
    wordapp.selection.find.replacement.text := FMATERIAIS;
    wordapp.selection.find.execute(replace := wdreplaceall);

    wordapp.selection.find.text := '<VALOR_MATERIAIS>';
    wordapp.selection.find.replacement.text := FVALOR_MATERIAIS;
    wordapp.selection.find.execute(replace := wdreplaceall);
    
    wordapp.selection.find.text := '<QUANTIDADE_MATERIAIS>';
    wordapp.selection.find.replacement.text := FQUANTIDADE_MATERIAIS;
    wordapp.selection.find.execute(replace := wdreplaceall);

    wordapp := unassigned;
end;

function TClasseContratoAlunos.getBAIRRO(value: string): iCadastroGerarContrado;
begin

  result := self;
  FBAIRRO := value;

end;

function TClasseContratoAlunos.getBAIRRO_EMPRESA(value: string): iCadastroGerarContrado;
begin

  result := self;
  FBAIRRO_EMPRESA := value;

end;

function TClasseContratoAlunos.getBAIRRO_RESPONSAVEL(value: string): iCadastroGerarContrado;
begin

  result := self;
  FBAIRRO_RESPONSAVEL := value;

end;

function TClasseContratoAlunos.getCampo(value: TEdit): iCadastroGerarContrado;
begin

  result := self;

  if value.Text = EmptyStr then
  begin
    value.SetFocus;
    Dialogs.MessageDlg('Digite qual informação deseja buscar', mtError,
      [mbOk], 0, mbOk);
  end
  else
    FCampo := value.Text;

end;

function TClasseContratoAlunos.getCELULAR(value: string): iCadastroGerarContrado;
begin

  result := self;
  FCELULAR := value;

end;

function TClasseContratoAlunos.getCELULAR_EMPRESA(value: string): iCadastroGerarContrado;
begin

  result := self;
  FCELULAR_EMPRESA := value;

end;

function TClasseContratoAlunos.getCELULAR_RESPONSAVEL(value: string): iCadastroGerarContrado;
begin

  result := self;
  FCELULAR_RESPONSAVEL := value;

end;

function TClasseContratoAlunos.getCEP(value: string): iCadastroGerarContrado;
begin

  result := self;
  FCEP := value;

end;

function TClasseContratoAlunos.getCEP_EMPRESA(value: string): iCadastroGerarContrado;
begin

  result := self;
  FCEP_EMPRESA := value;

end;

function TClasseContratoAlunos.getCEP_RESPONSAVEL(value: string): iCadastroGerarContrado;
begin

  result := self;
  FCEP_RESPONSAVEL := value;

end;

function TClasseContratoAlunos.getCIDADE(value: string): iCadastroGerarContrado;
begin

  result := self;
  FCIDADE := value;

end;

function TClasseContratoAlunos.getCIDADE_EMPRESA(value: string): iCadastroGerarContrado;
begin

  result := self;
  FCIDADE_EMPRESA := value;

end;

function TClasseContratoAlunos.getCIDADE_RESPONSAVEL(value: string): iCadastroGerarContrado;
begin

  Result := self;
  FCIDADE_RESPONSAVEL := value;

end;

function TClasseContratoAlunos.getCNPJ(value: string): iCadastroGerarContrado;
begin

  result := self;
  FCNPJ := value;

end;

function TClasseContratoAlunos.getCOMPLEMENTO(value: string): iCadastroGerarContrado;
begin

  result := self;
  FCOMPLEMENTO := value;

end;

function TClasseContratoAlunos.getCOMPLEMENTO_EMPRESA(value: string): iCadastroGerarContrado;
begin

  result := self;
  FCOMPLEMENTO_EMPRESA := value;

end;

function TClasseContratoAlunos.getCOMPLEMENTO_RESPONSAVEL(value: string): iCadastroGerarContrado;
begin

  result := self;
  FCOMPLEMENTO_RESPONSAVEL := value;

end;

function TClasseContratoAlunos.getCONSULTOR(value: string): iCadastroGerarContrado;
begin

  result := self;
  FCONSULTOR := value;

end;

function TClasseContratoAlunos.getCPF(value: string): iCadastroGerarContrado;
begin

  result := self;
  FCPF := value;

end;

function TClasseContratoAlunos.getCPF_RESPONSAVEL(value: string): iCadastroGerarContrado;
begin

  result := self;
  FCPF_RESPONSAVEL := value;

end;

function TClasseContratoAlunos.getDataFinal(value: TMaskEdit): iCadastroGerarContrado;
begin

  result := self;

  if value.Text = EmptyStr then
  begin
    value.SetFocus;
    Dialogs.MessageDlg('Digite uma data final valida', mtError, [mbOk],
      0, mbOk);
  end
  else
  begin

    try
      FDataFinal := StrToDate(value.Text);
    except
      begin
        value.SetFocus;
        Dialogs.MessageDlg('Digite uma data final valida', mtError,
          [mbOk], 0, mbOk);
      end;

    end;

  end;

end;

function TClasseContratoAlunos.getDataInicial(value: TMaskEdit)
  : iCadastroGerarContrado;
begin

  result := self;

  if value.Text = EmptyStr then
  begin
    value.SetFocus;
    Dialogs.MessageDlg('Digite una data inicial valida', mtError,
      [mbOk], 0, mbOk);
  end
  else
  begin

    try
      FDataInicial := StrToDate(value.Text);
    except
      begin
        value.SetFocus;
        Dialogs.MessageDlg('Digite una data inicial valida', mtError,
          [mbOk], 0, mbOk);
      end;

    end;

  end;

end;

function TClasseContratoAlunos.getDATA_ABERTURA(value: string): iCadastroGerarContrado;
begin

  result := self;
  FDATA_ABERTURA := value;

end;

function TClasseContratoAlunos.getDATA_CADASTRO(value: TDateTime): iCadastroGerarContrado;
begin

  result := self;
  FDATA_CADASTRO := value;

end;

function TClasseContratoAlunos.getDATA_CADASTRO_RESPONSAVEL(value: TDateTime): iCadastroGerarContrado;
begin

  result := self;
  FDATA_CADASTRO_RESPONSAVEL := DateToStr(value);

end;

function TClasseContratoAlunos.getDATA_INICIO(value: TDateTime): iCadastroGerarContrado;
begin

  result := self;
  FDATA_INICIO := value;

end;

function TClasseContratoAlunos.getDATA_INICIO_EXTENSO(value: string): iCadastroGerarContrado;
begin

  result := self;
  FDATA_INICIO_EXTENSO := FEntityDataPorExtenso.dataPorExtenso(StrToDate(value));

end;

function TClasseContratoAlunos.getDATA_MATRICULA(value: TDateTime): iCadastroGerarContrado;
begin

  result := self;
  FDATA_MATRICULA := value;

end;

function TClasseContratoAlunos.getDATA_NASCIMENTO(value: TDateTime): iCadastroGerarContrado;
begin

  result := self;
  FDATA_NASCIMENTO := value;

end;

function TClasseContratoAlunos.getDATA_NASCIMENTO_RESPONSAVEL(value: TDateTime): iCadastroGerarContrado;
begin

  result := self;
  FDATA_NASCIMENTO_RESPONSAVEL := DateToStr(value);

end;

function TClasseContratoAlunos.getDATA_TEMINO(value: TDateTime): iCadastroGerarContrado;
begin

  result := self;
  FDATA_TEMINO := value;

end;

function TClasseContratoAlunos.getDATA_TEMINO_EXTENSO(value: string): iCadastroGerarContrado;
begin

  result := self;
  FDATA_TEMINO_EXTENSO := FEntityDataPorExtenso.dataPorExtenso(StrToDate(value));

end;

function TClasseContratoAlunos.getDIAS_AULAS( value: string): iCadastroGerarContrado;
begin

  result := self;
  FDIAS_AULAS := value;

end;

function TClasseContratoAlunos.getDIA_HORARIO(value: string): iCadastroGerarContrado;
begin

  result := self;
  FDIA_HORARIO := value;

end;

function TClasseContratoAlunos.getDIA_HORARIO_TURMA(value: string): iCadastroGerarContrado;
begin

  result := self;
  FDIA_HORARIO_TURMA := value;

end;

function TClasseContratoAlunos.getDIA_HORARIO_TURMA_SALA(value: string): iCadastroGerarContrado;
begin

  result := self;
  FDIA_HORARIO_TURMA_SALA := value;

end;

function TClasseContratoAlunos.getDURACAO_MEDIA_CURSO(value: Real): iCadastroGerarContrado;
begin

  result := self;
  FDURACAO_MEDIA_CURSO := value;

end;

function TClasseContratoAlunos.getDURACAO_MEDIA_CURSO_EXTENSO(value: string): iCadastroGerarContrado;
begin

  result := self;
  FDURACAO_MEDIA_CURSO_EXTENSO := FEntityValorPorExtenso.valorPorExtensoSimples(StrToFloat(value));

end;

function TClasseContratoAlunos.getEMAIL(value: string): iCadastroGerarContrado;
begin

  result := self;
  FEMAIL := value;

end;

function TClasseContratoAlunos.getEMAIL_EMPRESA(value: string): iCadastroGerarContrado;
begin

  result := self;
  FEMAIL_EMPRESA := value;

end;

function TClasseContratoAlunos.getEMAIL_RESPONSAVEL(value: string): iCadastroGerarContrado;
begin

  result := self;
  FEMAIL_RESPONSAVEL := value;

end;

function TClasseContratoAlunos.getENDERECO(value: string): iCadastroGerarContrado;
begin

  result := self;
  FENDERECO := value;

end;

function TClasseContratoAlunos.getENDERECO_EMPRESA(value: string): iCadastroGerarContrado;
begin

  result := self;
  FENDERECO_EMPRESA := value;

end;

function TClasseContratoAlunos.getENDERECO_RESPONSAVEL(value: string): iCadastroGerarContrado;
begin

  result := self;
  FENDERECO_RESPONSAVEL := value;

end;

function TClasseContratoAlunos.getESTADO(value: string): iCadastroGerarContrado;
begin

  result := self;
  FESTADO := value;

end;

function TClasseContratoAlunos.getESTADO_EMPRESA(value: string): iCadastroGerarContrado;
begin

  result := self;
  FESTADO_EMPRESA := value;

end;

function TClasseContratoAlunos.getESTADO_RESPONSAVEL(value: string): iCadastroGerarContrado;
begin

  result := self;
  FESTADO_RESPONSAVEL := value;

end;

function TClasseContratoAlunos.getQTDEDIASAULAS(
  value: integer): iCadastroGerarContrado;
begin

  result := self;

  FQtdeDiasAulas := value;

end;

function TClasseContratoAlunos.getQTDEHORASAULAS(
  value: integer): iCadastroGerarContrado;
begin

  result := self;
  FQtdeHorasAulas := value;

end;

function TClasseContratoAlunos.getHORARIOS_AULAS( value: string): iCadastroGerarContrado;
begin

  result := self;
  FHORARIOS_AULAS := value;

end;

end.

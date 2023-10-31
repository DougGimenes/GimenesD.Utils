unit GimenesD.Utils.Controller.Excel;

interface

uses
  SysUtils,
  Variants,
  Classes,
  ComObj,
  DB,
  DBClient,
  FireDAC.Stan.Intf,
  FireDAC.Stan.Option,
  FireDAC.Stan.Param,
  FireDAC.Stan.Error,
  FireDAC.DatS,
  FireDAC.Phys.Intf,
  FireDAC.DApt.Intf,
  FireDAC.Comp.DataSet,
  FireDAC.Comp.Client,
  Windows,
  Data.Win.ADODB;

type
  TExcel = class
  private
    Excel:   Variant;
    DataSet: TDataSet;

    procedure InicializarExcel();

    procedure PrintCabecalho();
    procedure PrintConteudo(Linha: Integer = 2);
    procedure PrintCabecalhoCompleto(Titulo, Subtitulo: string);
  public
    constructor Create(); overload;
    constructor Create(AQuery: TDataSet); overload;
    destructor Destroy(); override;

    procedure DadosParaExcel(AExcelVisivelNoFinal: Boolean = True; ATitulo: string = ''; ASubtitulo: string = ''); overload;
  end;

implementation

{$D+}
{TExcel}

constructor TExcel.Create;
begin
  Self.InicializarExcel();
end;

constructor TExcel.Create(AQuery: TDataSet);
begin
  Self.InicializarExcel();
  Self.DataSet := AQuery;
end;

procedure TExcel.DadosParaExcel(AExcelVisivelNoFinal: Boolean = True; ATitulo: string = ''; ASubtitulo: string = '');
begin
  if Assigned(Self.DataSet) then
  begin
    Self.DataSet.First();
    if ATitulo <> '' then
    begin
      Self.PrintCabecalhoCompleto(ATitulo, ASubtitulo);
      Self.PrintConteudo(4);
    end
    else
    begin
      Self.PrintCabecalho();
      Self.PrintConteudo();
    end;
  end
  else
  begin
    raise Exception.Create('Query e DataSet null');
  end;

  Self.Excel.Visible := AExcelVisivelNoFinal;
end;

destructor TExcel.Destroy;
begin
  Self.Excel.Free();
end;

procedure TExcel.InicializarExcel;
begin
  Self.Excel         := CreateOleObject('Excel.Application');
  Self.Excel.Visible := False;
  Self.Excel.Workbooks.Add;
end;

procedure TExcel.PrintCabecalho;
var
  I: Integer;
begin
  for I := 0 to Self.DataSet.FieldCount - 1 do
  begin
    Self.Excel.WorkBooks[1].Sheets[1].Cells[1, I + 1] := Self.DataSet.Fields[I].DisplayName;
    Self.Excel.WorkBooks[1].Sheets[1].Cells[1, I + 1].Font.Size := 13;
    Self.Excel.WorkBooks[1].Sheets[1].Cells[1, I + 1].Font.Bold := True;
    Self.Excel.WorkBooks[1].Sheets[1].Cells[1, I + 1].Font.Color := RGB(255, 255, 255);
    Self.Excel.WorkBooks[1].Sheets[1].Cells[1, I + 1].Interior.Color := RGB(35, 92, 145);
  end;
end;

procedure TExcel.PrintCabecalhoCompleto(Titulo, Subtitulo: string);
var
  I: Integer;
begin
  Self.Excel.Range['A1:' + Char(Self.DataSet.FieldCount + 64) + '1'].MergeCells := True;
  Self.Excel.WorkBooks[1].Sheets[1].Cells[1, 1] := Titulo;
  Self.Excel.WorkBooks[1].Sheets[1].Cells[1, 1].Font.Size  := 15;
  Self.Excel.WorkBooks[1].Sheets[1].Cells[1, 1].Font.Bold  := True;
  Self.Excel.WorkBooks[1].Sheets[1].Cells[1, 1].Font.Color := RGB(255, 255, 255);
  Self.Excel.WorkBooks[1].Sheets[1].Cells[1, 1].Interior.Color := RGB(35, 92, 145);
  Self.Excel.WorkBooks[1].Sheets[1].Cells[1, 1].HorizontalAlignment := 3;

  Self.Excel.Range['A2:' + Char((Self.DataSet.FieldCount - 1) + 64) + '2'].MergeCells := True;
  Self.Excel.WorkBooks[1].Sheets[1].Cells[2, 1] := Subtitulo;
  Self.Excel.WorkBooks[1].Sheets[1].Cells[2, 1].Font.Size := 13;
  Self.Excel.WorkBooks[1].Sheets[1].Cells[2, 1].Font.Bold := True;
  Self.Excel.WorkBooks[1].Sheets[1].Cells[2, 1].HorizontalAlignment := 3;

  Self.Excel.WorkBooks[1].Sheets[1].Cells[2, Self.DataSet.FieldCount] := 'Emissão: ' + FormatDateTime('dd/mm/yyyy', Now);
  Self.Excel.WorkBooks[1].Sheets[1].Cells[2, Self.DataSet.FieldCount].Font.Size := 10;
  Self.Excel.WorkBooks[1].Sheets[1].Cells[2, Self.DataSet.FieldCount].Font.Bold := True;
  Self.Excel.WorkBooks[1].Sheets[1].Cells[2, Self.DataSet.FieldCount].HorizontalAlignment := 3;

  for I := 0 to Self.DataSet.FieldCount - 1 do
  begin
    Self.Excel.WorkBooks[1].Sheets[1].Cells[3, I + 1] := Self.DataSet.Fields[I].DisplayName;
    Self.Excel.WorkBooks[1].Sheets[1].Cells[3, I + 1].Font.Size := 13;
    Self.Excel.WorkBooks[1].Sheets[1].Cells[3, I + 1].Font.Bold := True;
    Self.Excel.WorkBooks[1].Sheets[1].Cells[3, I + 1].Font.Color := RGB(255, 255, 255);
    Self.Excel.WorkBooks[1].Sheets[1].Cells[3, I + 1].Interior.Color := RGB(35, 92, 145);
  end;
end;

procedure TExcel.PrintConteudo(Linha: Integer = 2);
var
  I: Integer;
begin
  while not Self.DataSet.Eof do
  begin

    for I := 0 to Self.DataSet.FieldCount - 1 do
    begin
      if (Self.DataSet.Fields[I].DataType = FtInteger) then
        Self.Excel.WorkBooks[1].Sheets[1].Cells[Linha, I + 1] := Self.DataSet.Fields[I].AsInteger
      else if (Self.DataSet.Fields[I].DataType = FtFloat) then
        Self.Excel.WorkBooks[1].Sheets[1].Cells[Linha, I + 1] := Self.DataSet.Fields[I].AsFloat
      else if (Self.DataSet.Fields[I].DataType = FtDateTime) then
        Self.Excel.WorkBooks[1].Sheets[1].Cells[Linha, I + 1] := 'Data: ' + FormatDateTime('dd/mm/yyyy', Self.DataSet.Fields[I].AsDateTime)
      else
        Self.Excel.WorkBooks[1].Sheets[1].Cells[Linha, I + 1] := Self.DataSet.Fields[I].AsString
    end;

    Self.DataSet.Next();
    Linha := Linha + 1;

  end;

  Self.Excel.Columns.AutoFit;
end;

end.

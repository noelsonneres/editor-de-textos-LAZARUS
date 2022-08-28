unit Unit1;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, Forms, Controls, Graphics, Dialogs, StdCtrls, Menus,
  Buttons, ExtDlgs, PrintersDlgs, RichMemo, RichMemoHelpers, Windows, LCL, Printers;

type

  { TForm1 }

  TForm1 = class(TForm)
    Button10: TButton;
    Button11: TButton;
    Button12: TButton;
    Button13: TButton;
    Button15: TButton;
    Button16: TButton;
    Button17: TButton;
    Button18: TButton;
    Button19: TButton;
    Button20: TButton;
    Button5: TButton;
    Button6: TButton;
    Button7: TButton;
    Button8: TButton;
    Button9: TButton;
    ColorDialog1: TColorDialog;
    Edit1: TEdit;
    Edit2: TEdit;
    FontDialog1: TFontDialog;
    MainMenu1: TMainMenu;
    MenuItem1: TMenuItem;
    MenuItem2: TMenuItem;
    MenuItem3: TMenuItem;
    MenuItem4: TMenuItem;
    OpenDialog1: TOpenDialog;
    OpenPictureDialog1: TOpenPictureDialog;
    PageSetupDialog1: TPageSetupDialog;
    PrintDialog1: TPrintDialog;
    PrinterSetupDialog1: TPrinterSetupDialog;
    RichMemo1: TRichMemo;
    SaveDialog1: TSaveDialog;
    procedure Button10Click(Sender: TObject);
    procedure Button11Click(Sender: TObject);
    procedure Button12Click(Sender: TObject);
    procedure Button13Click(Sender: TObject);
    procedure Button14Click(Sender: TObject);
    procedure Button15Click(Sender: TObject);
    procedure Button16Click(Sender: TObject);
    procedure Button17Click(Sender: TObject);
    procedure Button18Click(Sender: TObject);
    procedure Button19Click(Sender: TObject);
    procedure Button1Click(Sender: TObject);
    procedure Button20Click(Sender: TObject);
    procedure Button21Click(Sender: TObject);
    procedure Button2Click(Sender: TObject);
    procedure Button3Click(Sender: TObject);
    procedure Button4Click(Sender: TObject);
    procedure Button5Click(Sender: TObject);
    procedure Button6Click(Sender: TObject);
    procedure Button7Click(Sender: TObject);
    procedure Button8Click(Sender: TObject);
    procedure Button9Click(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure MenuItem2Click(Sender: TObject);
    procedure MenuItem3Click(Sender: TObject);
    procedure MenuItem4Click(Sender: TObject);
    procedure RichMemo1Change(Sender: TObject);
  private

    procedure identar;

  public

  end;

var
  Form1: TForm1;

implementation

{$R *.lfm}

{ TForm1 }

procedure PageSetupToMargins(pg: TPageSetupDialog; var p: TPrintParams);
var
  md : Single; // modifies to tunr into inches
  pt : Single;
begin
  //if pg.Units=pmMillimeters then
  //  md:=0.254
  //else
  //  md:=0.001;

  p.Margins.Left:=25;
  p.Margins.Top:=50;
  p.Margins.Right:=24;
  p.Margins.Bottom:=50;

  //writeln(p.Margins.Left:0:2);

end;

procedure TForm1.RichMemo1Change(Sender: TObject);
begin

end;

procedure TForm1.identar;
var
  m: TParaMetric;
begin

  RichMemo1.GetParaMetric(RichMemo1.SelStart, m);

  m.FirstLine := 20;
  m.HeadIndent := 20;

  RichMemo1.Paragraph.FirstIndent := 10;
  RichMemo1.Paragraph.LeftIndent := 25;
  RichMemo1.Paragraph.RightIndent := 25;

  RichMemo1.SelStart := 0;

  RichMemo1.SetParaMetric(RichMemo1.SelStart, RichMemo1.SelLength, m);

end;

procedure TForm1.Button1Click(Sender: TObject);
begin

end;

procedure TForm1.Button20Click(Sender: TObject);
var
  prm : TPrintParams;
begin
  if not Assigned(Printer) then begin
    ShowMessage('No printer found.');
    Exit;
  end;
  InitPrintParams(prm);
  prm.SelStart:=RichMemo1.SelStart;
  prm.SelLength:=RichMemo1.SelLength;
  prm.JobTitle:='Rich Memo Printing';
  PageSetupToMargins(PageSetupDialog1, prm);

  RichMemo1.Print(prm);
end;

procedure TForm1.Button21Click(Sender: TObject);
begin
  RichMemo1.Lines.Add('*****************************');
end;

procedure TForm1.Button10Click(Sender: TObject);
begin
  RichMemo1.SetParaAlignment(RichMemo1.SelStart, RichMemo1.SelLength, paJustify);
end;

procedure TForm1.Button11Click(Sender: TObject);
begin

  RichMemo1.SetRangeParams(
    RichMemo1.SelStart, RichMemo1.SelLength,
    [tmm_Styles], '', 0, clNone, [fsBold], []);

end;

procedure TForm1.Button12Click(Sender: TObject);
begin
  RichMemo1.SetRangeParams(
    RichMemo1.SelStart, RichMemo1.SelLength,
    [tmm_Styles], '', 0, clNone, [fsItalic], []);
end;

procedure TForm1.Button13Click(Sender: TObject);
begin
  RichMemo1.SetRangeParams(
    RichMemo1.SelStart, RichMemo1.SelLength,
    [tmm_Styles], '', 0, clNone, [fsUnderline], []);
end;

procedure TForm1.Button14Click(Sender: TObject);
var
  ofs, len: integer;
begin
  RichMemo1.GetStyleRange(RichMemo1.SelStart, ofs, len);
  if (ofs = RichMEmo1.SelStart) and (len = RichMemo1.SelLength) then
  begin
    ofs := ofs + len;
    RichMemo1.GetStyleRange(ofs, ofs, len);
  end;
  RichMemo1.SelStart := ofs;
  RichMemo1.SelLength := len;
end;

procedure TForm1.Button15Click(Sender: TObject);
begin
  RichMemo1.Redo;
end;

procedure TForm1.Button16Click(Sender: TObject);
begin
  RichMemo1.Undo;
end;

procedure TForm1.Button17Click(Sender: TObject);
var
  Posicao: integer;
begin
  Posicao := 1;

  with RichMemo1 do
  begin
    Posicao := FindText(Edit1.Text, 0, Length(Text), []);

    SelStart := Posicao;
    SelLength := Length(Edit1.Text);
    //        SelText := TextoNovo;

    SetFocus;

  end;
end;

procedure TForm1.Button18Click(Sender: TObject);
var
  Posicao: integer;
begin
  Posicao := 1;

  with RichMemo1 do
  begin
    Posicao := FindText(Edit1.Text, 0, Length(Text), []);

    SelStart := Posicao;
    SelLength := Length(Edit1.Text);
    SelText := Edit2.Text;

    SetFocus;

  end;
end;

procedure TForm1.Button19Click(Sender: TObject);
begin
  if PrinterSetupDialog1.Execute then
    RichMemo1.Print('Printing with margins');

end;

procedure TForm1.Button2Click(Sender: TObject);
begin

end;

procedure TForm1.Button3Click(Sender: TObject);
begin
  RichMemo1.SetParaAlignment(RichMemo1.SelStart, RichMemo1.SelLength, paCenter);
end;

procedure TForm1.Button4Click(Sender: TObject);
var
  Posicao: integer;
begin
  Posicao := 1;

  with RichMemo1 do
    repeat

      Posicao := FindText('Novo', 0, Length(Text), []);

      if Posicao > 0 then
      begin
        SelStart := Posicao;
        SelLength := Length('Novo');
        SelText := 'Teste';

        SetFocus;
        Posicao := 0;
      end;
    until (Posicao <> 0);

end;

procedure TForm1.Button5Click(Sender: TObject);
begin
  if ColorDialog1.Execute then
  begin
    RichMemo1.SetRangeColor(RichMemo1.SelStart, RichMemo1.SelLength, ColorDialog1.Color);
  end;
end;

procedure TForm1.Button6Click(Sender: TObject);
begin
  if FontDialog1.Execute then

    RichMemo1.SetTextAttributes(RichMemo1.SelStart, RichMemo1.SelLength,
      FontDialog1.Font);
end;

procedure TForm1.Button7Click(Sender: TObject);
begin
  RichMemo1.SetParaAlignment(RichMemo1.SelStart, RichMemo1.SelLength, paLeft);
end;

procedure TForm1.Button8Click(Sender: TObject);
begin
  RichMemo1.SetParaAlignment(RichMemo1.SelStart, RichMemo1.SelLength, paCenter);
end;

procedure TForm1.Button9Click(Sender: TObject);
begin
  RichMemo1.SetParaAlignment(RichMemo1.SelStart, RichMemo1.SelLength, paRight);
end;

procedure TForm1.FormShow(Sender: TObject);
begin

  identar;

end;

procedure TForm1.MenuItem2Click(Sender: TObject);
var
  fs: TFileStream;
  tm: longword;
begin
  if OpenDialog1.Execute then
  begin
    fs := nil;
    try
      // Utf8ToAnsi is required for windows
      fs := TFileStream.Create(Utf8ToAnsi(OpenDialog1.FileName),
        fmOpenRead or fmShareDenyNone);
      RichMemo1.LoadRichText(fs);
    except
    end;
    fs.Free;
  end;

end;

procedure TForm1.MenuItem3Click(Sender: TObject);
begin
  RichMemo1.Clear;
  RichMemo1.Lines.Add('');
  identar;
end;

procedure TForm1.MenuItem4Click(Sender: TObject);
var
  fs: TFileStream;
begin
  if SaveDialog1.Execute then
  begin
    fs := nil;
    try
      fs := TFileStream.Create(Utf8ToAnsi(SaveDialog1.FileName), fmCreate);
      RichMemo1.SaveRichText(fs);
    except
    end;
    fs.Free;
  end;

end;

end.

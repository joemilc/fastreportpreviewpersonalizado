unit uRel_Preview;

interface

uses
   Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
   Dialogs, frxClass, frxPreview, StdCtrls, ExtCtrls, Buttons, frxChart,
   frxDCtrl, frxRich, frxCross, frxExportMail, frxExportPDF, frxBarcode,
   ComCtrls, frxExportXLS, frxExportRTF, StrUtils, frxExportBaseDialog, uDados;

type
   TfRel_Preview = class(TForm)
      Panel1: TPanel;
      Bevel1: TBevel;
      Bevel2: TBevel;
      FastRep: TfrxReport;
      frxBarCodeObject1: TfrxBarCodeObject;
      frxPDFExport1: TfrxPDFExport;
      frxCrossObject1: TfrxCrossObject;
      frxRichObject1: TfrxRichObject;
      frxDialogControls1: TfrxDialogControls;
      frxChartObject1: TfrxChartObject;
      StatusBar1: TStatusBar;
      btnImprimir: TSpeedButton;
      btnExportaPDF: TSpeedButton;
      btnFirst: TSpeedButton;
      btnPrior: TSpeedButton;
      btnNext: TSpeedButton;
      btnLast: TSpeedButton;
      btnFechar: TSpeedButton;
      Label1: TLabel;
      edPagina: TEdit;
      Label2: TLabel;
      edZoom: TComboBox;
      Bevel3: TBevel;
      Label3: TLabel;
      edLocalizar: TEdit;
      Bevel4: TBevel;
      btnLocalizar: TSpeedButton;
      btnLocalizarProx: TSpeedButton;
      Panel2: TPanel;
      frxPreview1: TfrxPreview;
      frxXLSExport1: TfrxXLSExport;
      btnExportaXLS: TSpeedButton;
      P1: TProgressBar;
      btnEditar: TSpeedButton;
      btnExportarRTF: TSpeedButton;
      frxRTFExport1: TfrxRTFExport;
      btnEmail: TSpeedButton;
      frxMailExport1: TfrxMailExport;
      procedure FormClose(Sender: TObject; var Action: TCloseAction);
      procedure btnImprimirClick(Sender: TObject);
      procedure btnFirstClick(Sender: TObject);
      procedure btnPriorClick(Sender: TObject);
      procedure btnNextClick(Sender: TObject);
      procedure btnLastClick(Sender: TObject);
      procedure btnFecharClick(Sender: TObject);
      procedure btnExportaPDFClick(Sender: TObject);
      procedure frxPreview1PageChanged(Sender: TfrxPreview; PageNo: Integer);
      procedure FormKeyPress(Sender: TObject; var Key: Char);
      procedure edPaginaKeyPress(Sender: TObject; var Key: Char);
      procedure FormShow(Sender: TObject);
      procedure edZoomChange(Sender: TObject);
      procedure FormKeyDown(Sender: TObject; var Key: Word;
         Shift: TShiftState);
      procedure btnLocalizarClick(Sender: TObject);
      procedure btnLocalizarProxClick(Sender: TObject);
      procedure FormMouseWheel(Sender: TObject; Shift: TShiftState;
         WheelDelta: Integer; MousePos: TPoint; var Handled: Boolean);
      procedure FormActivate(Sender: TObject);
      procedure frxPreview1Click(Sender: TObject);
      procedure edPaginaExit(Sender: TObject);
      procedure FormCreate(Sender: TObject);
      procedure btnExportaXLSClick(Sender: TObject);
      procedure FastRepProgress(Sender: TfrxReport;
         ProgressType: TfrxProgressType; Progress: Integer);
      procedure FastRepProgressStart(Sender: TfrxReport;
         ProgressType: TfrxProgressType; Progress: Integer);
      procedure FastRepProgressStop(Sender: TfrxReport;
         ProgressType: TfrxProgressType; Progress: Integer);
      procedure btnEditarClick(Sender: TObject);
      procedure btnExportarRTFClick(Sender: TObject);
      procedure edLocalizarKeyPress(Sender: TObject; var Key: Char);
      procedure btnEmailClick(Sender: TObject);
    function FastRepUserFunction(const MethodName: string;
      var Params: Variant): Variant;
   private
      { Private declarations }
   public
      { Public declarations }
   end;

var
   fRel_Preview: TfRel_Preview;

implementation

{$R *.dfm}

procedure TfRel_Preview.FormClose(Sender: TObject;
   var Action: TCloseAction);
begin
   Action := caFree;
   fRel_Preview := nil;
end;

procedure TfRel_Preview.btnImprimirClick(Sender: TObject);
begin
   frxPreview1.Print;
end;

procedure TfRel_Preview.btnFirstClick(Sender: TObject);
begin
   frxPreview1.First;
end;

procedure TfRel_Preview.btnPriorClick(Sender: TObject);
begin
   frxPreview1.Prior;
end;

procedure TfRel_Preview.btnNextClick(Sender: TObject);
begin
   frxPreview1.Next;
end;

procedure TfRel_Preview.btnLastClick(Sender: TObject);
begin
   frxPreview1.Last;
end;

procedure TfRel_Preview.btnFecharClick(Sender: TObject);
begin
   Close;
end;

procedure TfRel_Preview.btnExportaPDFClick(Sender: TObject);
begin
   frxPreview1.Export(frxPDFExport1);
end;

procedure TfRel_Preview.frxPreview1PageChanged(Sender: TfrxPreview;
   PageNo: Integer);
begin
   StatusBar1.Panels[1].Text := 'Pág. ' + IntToStr(PageNo);
   StatusBar1.Panels[0].Text := 'Páginas: ' + IntToSTr(frxPreview1.PageCount);
end;

procedure TfRel_Preview.FormKeyPress(Sender: TObject; var Key: Char);
begin
   if Key = #27 then
      btnFecharClick(Sender);
end;

procedure TfRel_Preview.edPaginaKeyPress(Sender: TObject; var Key: Char);
begin
   if Key = #13 then
   begin
      frxPreview1.SetPosition(StrToInt(edPagina.Text), 0);
      StatusBar1.Panels[0].Text := 'Páginas: ' + IntToSTr(frxPreview1.PageCount);
      Key := #0;
      frxPreview1.SetFocus;
   end;
end;

procedure TfRel_Preview.FormShow(Sender: TObject);
begin
   StatusBar1.Panels[0].Text := 'Páginas: ' + IntToSTr(frxPreview1.PageCount);
   frxPreview1.SetFocus;
end;

procedure TfRel_Preview.edZoomChange(Sender: TObject);
var
  i: Integer;
begin
  case edZoom.ItemIndex of
    0: frxPreview1.ZoomMode := zmDefault;
    1: frxPreview1.ZoomMode := zmWholePage;
  else
     frxPreview1.ZoomMode := zmDefault;
     i := Pos('%', edZoom.Text)-1;
     frxPreview1.Zoom := StrToInt(LeftStr(edZoom.Text, i)) / 100;
  end;
  frxPreview1.SetFocus;
end;

procedure TfRel_Preview.FormKeyDown(Sender: TObject; var Key: Word;
   Shift: TShiftState);
begin
   if ssCtrl in Shift then
   begin
      if KEY in [Ord('P'), Ord('p')] then
         btnImprimirClick(Sender);
      if KEY in [Ord('E'), Ord('e')] then
         btnExportaPDFClick(Sender);
      if KEY in [Ord('X'), Ord('x')] then
         btnExportaXLSClick(Sender);
   end;
end;

procedure TfRel_Preview.btnLocalizarClick(Sender: TObject);
begin
   frxPreview1.FindText(edLocalizar.Text, True, False);
   frxPreview1.SetFocus;
end;

procedure TfRel_Preview.btnLocalizarProxClick(Sender: TObject);
begin
   frxPreview1.FindNext;
   frxPreview1.SetFocus;
end;

procedure TfRel_Preview.FormMouseWheel(Sender: TObject; Shift: TShiftState;
   WheelDelta: Integer; MousePos: TPoint; var Handled: Boolean);
begin
   frxPreview1.MouseWheelScroll(WheelDelta, Shift, MousePos);
   frxPreview1.SetFocus;
end;

procedure TfRel_Preview.FormActivate(Sender: TObject);
begin
   frxPreview1.SetFocus;
end;

procedure TfRel_Preview.frxPreview1Click(Sender: TObject);
begin
   frxPreview1.SetFocus;
end;

procedure TfRel_Preview.edPaginaExit(Sender: TObject);
begin
   frxPreview1.SetFocus;
end;

procedure TfRel_Preview.FormCreate(Sender: TObject);
var Grupo: string;
begin
   { ------------------------

    ADICIONA AS SUAS FUNCOES

   -------------------------- }


   {with FastRep do
   begin

      Grupo := 'Minhas Funções';

      AddFunction('function DiaSemana(Data: TDate): string', Grupo,
         'Retorna o dia da Semana (Segunda, terça, ...)');

      AddFunction('function Extenso(xValor: Currency): string', Grupo,
         'Valor por Extenso');

      AddFunction('function MinutosParaHoras(Minutos: Integer): String', Grupo,
         'Converte Minutos para Horas');

      AddFunction('function RightStr(Texto: String; Quant: Integer): String', Grupo,
         'Caracteres a Direita da String');
   end;}
end;

procedure TfRel_Preview.btnExportaXLSClick(Sender: TObject);
begin
   frxPreview1.Export(frxXLSExport1);
end;

procedure TfRel_Preview.FastRepProgress(Sender: TfrxReport;
   ProgressType: TfrxProgressType; Progress: Integer);
begin
   P1.StepIt;
end;

procedure TfRel_Preview.FastRepProgressStart(Sender: TfrxReport;
   ProgressType: TfrxProgressType; Progress: Integer);
begin
   P1.Visible := True;
end;

procedure TfRel_Preview.FastRepProgressStop(Sender: TfrxReport;
   ProgressType: TfrxProgressType; Progress: Integer);
begin
   P1.Visible := False;
end;

function TfRel_Preview.FastRepUserFunction(const MethodName: string;
  var Params: Variant): Variant;
begin
   { ------------------------

    EXECUTA AS SUAS FUNCOES

   -------------------------- }

   {if MethodName = 'EXTENSO' then
      Result := Extenso(Params[0])
   else if MethodName = 'DIASEMANA' then
      Result := DiaSemana(Params[0])
   else if MethodName = 'MINUTOSPARAHORAS' then
      Result := MinutosParaHoras(Params[0])
   else if MethodName = 'RIGHTSTR' then
      Result := RightStr(Params[0], Params[1]);}

end;

procedure TfRel_Preview.btnEditarClick(Sender: TObject);
begin
   frxPreview1.Edit;
   frxPreview1.LoadFromFile;
end;

procedure TfRel_Preview.btnExportarRTFClick(Sender: TObject);
begin
   frxPreview1.Export(frxRTFExport1);
end;

procedure TfRel_Preview.edLocalizarKeyPress(Sender: TObject;
   var Key: Char);
begin
   if Key = #13 then
   begin
      Key := #0;
      btnLocalizar.Click;
   end;
end;

procedure TfRel_Preview.btnEmailClick(Sender: TObject);
var Arquivo, Email: string;
   st: TStrings;
begin
   { -----------------

     SUA ROTINA DE EMAIL

     ----------------- }

   st := TStringList.Create;
   st.Add('Segue documento em anexo.');
   st.Add('Obrigado por usar nossos serviços.');
   Arquivo := ExtractFilePath(ParamStr(0)) + 'Temp\' + ChangeFileExt(ExtractFileName(FastRep.FileName), '.pdf');
   frxPDFExport1.ShowDialog := False;
   frxPDFExport1.FileName := Arquivo;
   FastRep.Export(frxPDFExport1);
   Email := '';
   {if InputQuery('Email', 'Informe Email destinatário', Email) then
      EnviaEmail(ChangeFileExt(ExtractFileName(FastRep.FileName), ''),
         Email,
         Arquivo,
         st);}
   st.Free;
end;

end.

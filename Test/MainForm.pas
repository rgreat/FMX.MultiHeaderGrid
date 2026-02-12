unit MainForm;

interface

uses
  System.SysUtils, System.Types, System.UITypes, System.Classes, System.Variants,
  FMX.Types, FMX.Controls, FMX.Forms, FMX.Graphics, FMX.Dialogs, FMX.MultiHeaderGrid,
  FMX.Controls.Presentation,
  FMX.StdCtrls, FMX.Objects, FMX.Edit, FMX.TabControl;

type
  TForm1 = class(TForm)
    Panel1: TPanel;
    ButtonMergeCells: TButton;
    ButtonAutoSize: TButton;
    ButtonClearMergeCells: TButton;
    RowCountEdit: TEdit;
    Label1: TLabel;
    ButtonBenchmark: TButton;
    TabControl1: TTabControl;
    TabItem1: TTabItem;
    TabItem2: TTabItem;
    StatusBar1: TStatusBar;
    StatusText: TLabel;
    Grid1: TMultiHeaderGrid;
    Grid2: TMultiHeaderStringGrid;
    RowSelectCheckBox: TCheckBox;
    ButtonFillCells: TButton;
    WordWrapCheckBox: TCheckBox;
    procedure FormCreate(Sender: TObject);
    procedure ButtonMergeCellsClick(Sender: TObject);
    procedure ButtonAutoSizeClick(Sender: TObject);
    procedure ButtonClearMergeCellsClick(Sender: TObject);
    procedure GridSelectCell(Sender: TObject);
    procedure ButtonBenchmarkClick(Sender: TObject);

    procedure Grid1GetCellText(Sender: TObject; ACol, ARow: Integer; var Text: string);
    procedure Grid1SetCellText(Sender: TObject; ACol, ARow: Integer; const Text: string);
    procedure Grid1GetCellStyle(Sender: TObject; ACol, ARow: Integer; var Style: TCellStyle);
    procedure Grid1SetCellStyle(Sender: TObject; ACol, ARow: Integer; const Style: TCellStyle);
    procedure GridDrawCell(Sender: TObject; ACol, ARow: Integer; Canvas: TCanvas; const Rect: TRectF;
                           IsSelected: boolean; const Text: string; var Handled: Boolean);
    procedure RowSelectCheckBoxChange(Sender: TObject);
    procedure GridGridScroll(Sender: TObject; Left, Top: Integer);
    procedure RowCountEditChange(Sender: TObject);
    procedure ButtonFillCellsClick(Sender: TObject);
    procedure GridDblClick(Sender: TObject);
    procedure WordWrapCheckBoxChange(Sender: TObject);
  private
    Grid1CellTexts: array of array of string;
    Grid1CellStyles: array of array of TCellStyle;

    function ActiveGrid: TMultiHeaderGrid;
  public
    procedure InitGrid(Grid: TMultiHeaderGrid);
    procedure InitGridHeader(Grid: TMultiHeaderGrid);
    procedure InitGridParams(Grid: TMultiHeaderGrid);
    procedure FillGrid(Grid: TMultiHeaderGrid);
    procedure MergeCellsinGrid(Grid: TMultiHeaderGrid);
  end;

var
  Form1: TForm1;

implementation

uses
  System.Diagnostics, System.Math;

{$R *.fmx}

// ---------------------------------------------------------
// -------------------- Constructor ------------------------
// ---------------------------------------------------------

procedure TForm1.FormCreate(Sender: TObject);
begin
  RowCountEdit.Text:='1000';

  InitGrid(Grid1);
  InitGrid(Grid2);
end;

// ---------------------------------------------------------
// --------------- Create and Fill Grids -------------------
// ---------------------------------------------------------

procedure TForm1.InitGrid(Grid: TMultiHeaderGrid);
begin
  InitGridParams(Grid);
  InitGridHeader(Grid);
  FillGrid(Grid);
  MergeCellsinGrid(Grid);
  Grid.AutoSize;
end;

procedure TForm1.InitGridParams(Grid: TMultiHeaderGrid);
begin
  Grid.ColCount:=15;
  RowCountEditChange(nil);

  Grid.ColMinWidth[0]:=50;
  Grid.ColMaxWidth[0]:=200;

  Grid.ColTextHAlignment[3]:=TTextAlign.Trailing;
  Grid.ColTextHAlignment[6]:=TTextAlign.Center;
  Grid.ColTextHAlignment[13]:=TTextAlign.Center;
  Grid.ColTextHAlignment[14]:=TTextAlign.Center;

  Grid.CellFont.Size := 10;
  Grid.HeaderFont.Size := 11;
  Grid.HeaderFont.Style := [TFontStyle.fsBold];
  Grid.CellColorAlternate := TAlphaColors.Whitesmoke;
  Grid.GridLineColor := TAlphaColors.Gray;
end;

procedure TForm1.InitGridHeader(Grid: TMultiHeaderGrid);
begin
  // Настраиваем многоуровневые заголовки
  var Row:=Grid.Header.AddRow(30);
  var Cell:=Row.FillRow('Annual report');
  Cell.Style.FontSize:=14;
  Cell.Style.FontStyle:=[TFontStyle.fsBold];
  Cell.Style.CellColor:=TAlphaColors.Lightblue;

  Row:=Grid.Header.AddRow(30);
  Row.AddColumn('First Quarter',3).Style.CellColor:=TAlphaColors.Lightgreen;
  Row.AddColumn('Second Quarter',3).Style.CellColor:=TAlphaColors.Lightgreen;

  Cell:=Row.AddColumn('Summary for'#13#10'1st half of'#13#10+'the year',1,2);
  Cell.Style.FontSize:=14;
  Cell.Style.FontStyle:=[];
  Cell.Style.CellColor:=TAlphaColors.Lightsteelblue;
  Cell.Style.FontColor:=TAlphaColors.Darkred;

  Row.AddColumn('Third Quarter',3).Style.CellColor:=TAlphaColors.Lightgreen;
  Row.AddColumn('Forth Quarter',3).Style.CellColor:=TAlphaColors.Lightgreen;

  Cell:=Row.AddColumn('Summary for'#13#10'2nd half of'#13#10+'the year',1,2);
  Cell.Style.FontSize:=14;
  Cell.Style.FontStyle:=[];
  Cell.Style.CellColor:=TAlphaColors.Lightsteelblue;
  Cell.Style.FontColor:=TAlphaColors.Darkred;

  Cell:=Row.AddColumn('Year'#13#10'Totals',1,2);
  Cell.Style.FontSize:=14;
  Cell.Style.FontStyle:=[];
  Cell.Style.CellColor:=TAlphaColors.Lightsteelblue;
  Cell.Style.FontColor:=TAlphaColors.Darkred;

  Row:=Grid.Header.AddRow(30);
  var FormatSettings:=TFormatSettings.Create('en-us');
  for var i:=1 to 12 do begin
    Row.AddColumn(FormatSettings.LongMonthNames[i]);
  end;
end;

procedure TForm1.FillGrid(Grid: TMultiHeaderGrid);
begin
  for var Y:=0 to Grid.RowCount-1 do begin
    for var X:=0 to Grid.ColCount-1 do begin

      if (X+Y=0) or ((X+Y*2) mod 9 = 3) and (X<Grid.ColCount-1) then begin
        Grid.Cells[X,Y]:='Very Large Cell'#13#10'Number '+(X+Y*7).ToString;
      end else begin
        Grid.Cells[X,Y]:=Format('[%d,%d]', [X, Y]);
      end;

      var CellStyle:=Grid.CellStyle[X,Y];
      if (X+Y=0) or (((X+Y*2) mod 9 = 3) and (X<Grid.ColCount-1)) then begin
        case (X*2+Y*3) mod 3 of
          0: CellStyle.FontColor:=TAlphaColors.Blue;
          1: CellStyle.FontColor:=TAlphaColors.Green;
          2: CellStyle.FontColor:=TAlphaColors.Red;
        end;

        case (X+Y) mod 3 of
          0: CellStyle.CellColor:=TAlphaColors.Yellow;
          1: CellStyle.CellColor:=TAlphaColors.Chartreuse;
          2: CellStyle.CellColor:=TAlphaColors.Deepskyblue;
        end;
        case (X+Y) mod 3 of
          0: CellStyle.SelectedCellColor:=TAlphaColors.Red;
          1: CellStyle.SelectedCellColor:=TAlphaColors.Red;
          2: CellStyle.SelectedCellColor:=TAlphaColors.Red;
        end;

        case (X) mod 2 of
          0: CellStyle.FontSize:=14;
        end;

        case (X*3+Y*2) mod 3 of
          1: CellStyle.FontStyle:=[TFontStyle.fsBold];
          2: CellStyle.FontStyle:=[TFontStyle.fsItalic];
          3: CellStyle.FontStyle:=[TFontStyle.fsUnderline];
        end;
      end;
      if not CellStyle.IsEmpty then begin
        Grid.CellStyle[X,Y]:=CellStyle;
      end;

    end;
  end;
end;

procedure TForm1.MergeCellsinGrid(Grid: TMultiHeaderGrid);
begin
  Grid.ClearMergedCells;

  Grid.MergeCells(1, 1, 2, 3, 'Custom Draw');

  Grid.MergeCells(4, 2, 3, 2, 'var T:=TStopwatch.StartNew;'#13#10+
                              'var Cnt:=0;'#13#10+
                              'var MaxTop:=Max(Trunc(Grid.FullTableHeight-Grid.HeaderHeight-Grid.ViewPortDataHeight),0);'#13#10+
                              'var MaxLeft:=Max(Trunc(Grid.FullTableWidth-Grid.ViewPortDataWidth),0);');
  Grid.CellStyle[3,2].TextHAlignment:=TTextAlign.Leading;

  Grid.MergeCells(6, 6, 4, 1, 'var W:=Random(Grid.ColCount-3)+1'#13#10+
                              'var H:=Random(4)+1'#13#10+
                              'var L:=Random(Grid.ColCount-W);');

  var CellStyle:=Grid.CellStyle[6,6];
  CellStyle.TextHAlignment:=TTextAlign.Leading;
  Grid.CellStyle[6,6]:=CellStyle;



  for var T:=11 to Grid.RowCount-5 do begin
    if Random(10)<6 then Continue;

    var W:=Random(Grid.ColCount-3)+1;
    var H:=Random(4)+1;
    var L:=Random(Grid.ColCount-W);

    var Text:='Merged cell';
    var TH:=Random(4);
    for var i:=0 to TH do Text:=Text+#13#10+'Line '+(I+2).ToString;

    if Grid.MergeCells(L,T,W,H,Text) then begin
      CellStyle:=Grid.CellStyle[L,T];
      CellStyle.TextHAlignment:=TTextAlign.Center;
      Grid.CellStyle[L,T]:=CellStyle;
    end;
  end;
  Grid.Invalidate;
end;

// ---------------------------------------------------------
// ------------------- Handle Controls ---------------------
// ---------------------------------------------------------

function TForm1.ActiveGrid: TMultiHeaderGrid;
begin
  if TabControl1.ActiveTab=TabItem1 then begin
    Result:=Grid1;
  end else begin
    Result:=Grid2;
  end;
end;

procedure TForm1.RowCountEditChange(Sender: TObject);
begin
  Grid1.RowCount:=StrToIntDef(RowCountEdit.Text,Grid1.RowCount);
  Grid2.RowCount:=StrToIntDef(RowCountEdit.Text,Grid2.RowCount);
end;

procedure TForm1.ButtonFillCellsClick(Sender: TObject);
begin
  FillGrid(ActiveGrid);
end;

procedure TForm1.ButtonMergeCellsClick(Sender: TObject);
begin
  MergeCellsinGrid(ActiveGrid);
end;

procedure TForm1.ButtonAutoSizeClick(Sender: TObject);
begin
  ActiveGrid.AutoSize;
end;

procedure TForm1.ButtonClearMergeCellsClick(Sender: TObject);
begin
  ActiveGrid.ClearMergedCells;
end;

procedure TForm1.ButtonBenchmarkClick(Sender: TObject);
begin
  var Grid:=ActiveGrid;

  var T:=TStopwatch.StartNew;
  var Cnt:=0;
  var MaxTop:=Max(Trunc(Grid.FullTableHeight-Grid.HeaderHeight-Grid.ViewPortDataHeight),0);
  var MaxLeft:=Max(Trunc(Grid.FullTableWidth-Grid.ViewPortWidth),0);
  while T.ElapsedMilliseconds<5000 do begin
    Grid.ViewTop:=Random(MaxTop);
    Grid.ViewLeft:=Random(MaxLeft);
    inc(Cnt);
    StatusText.Text:='Rendering: '+round(T.ElapsedMilliseconds/5000*100).ToString+'%, '+Cnt.ToString+' renders.';

    Application.ProcessMessages;
  end;
  var MS:=T.ElapsedMilliseconds;

  Grid.ViewTop:=0;
  ShowMessage('Elapsed: '+MS.ToString+' ms.'#13#10+
              'Renders: '+Cnt.ToString+#13#10+
              'FPS: '+(Trunc(Cnt/(MS/10000))/10).ToString);
end;

procedure TForm1.RowSelectCheckBoxChange(Sender: TObject);
begin
  Grid1.RowSelect:=RowSelectCheckBox.IsChecked;
  Grid2.RowSelect:=RowSelectCheckBox.IsChecked;
end;

procedure TForm1.WordWrapCheckBoxChange(Sender: TObject);
begin
  Grid1.WordWrap:=WordWrapCheckBox.IsChecked;
  Grid2.WordWrap:=WordWrapCheckBox.IsChecked;
end;

// ---------------------------------------------------------
// -------------------- Grid Events ------------------------
// ---------------------------------------------------------

procedure TForm1.GridDblClick(Sender: TObject);
begin
  var Grid:=ActiveGrid;
  ShowMessage(Grid.Cells[Grid.Col,Grid.Row]);
end;

procedure TForm1.GridDrawCell(Sender: TObject; ACol, ARow: Integer; Canvas: TCanvas; const Rect: TRectF;
                              IsSelected: boolean; const Text: string; var Handled: Boolean);
begin
  // Custom Draw Cell

  If Text='Custom Draw' then begin
    if IsSelected then begin
      Canvas.Fill.Color:=TAlphaColorRec.Lightblue;
      Canvas.FillRect(Rect,1);
    end else begin
      Canvas.Fill.Color:=TAlphaColorRec.White;
      Canvas.FillRect(Rect,1);
    end;
    Rect.Inflate(-Grid1.GridLineWidth/4-2,-Grid1.GridLineWidth/4-2);
    Canvas.Fill.Color:=TAlphaColorRec.Darkgreen;
    Canvas.FillEllipse(Rect,0.2);
    Canvas.Stroke.Kind:=TBrushKind.Solid;
    Canvas.Stroke.Color:=TAlphaColorRec.Black;
    Canvas.DrawEllipse(Rect,1);
    Canvas.Font.Size:=14;
    Canvas.Fill.Color:=TAlphaColorRec.Red;
    Canvas.FillText(Rect, Text, False, 1, [], TTextAlign.Center, TTextAlign.Center);
    Handled:=True;
  end;
end;

procedure TForm1.Grid1GetCellStyle(Sender: TObject; ACol, ARow: Integer; var Style: TCellStyle);
begin
  if (ARow<0) or (ACol<0) then Exit;
  if (ARow>High(Grid1CellStyles)) or (ACol>High(Grid1CellStyles[ARow])) then Exit;

  Style:=Grid1CellStyles[ARow,ACol];
end;

procedure TForm1.Grid1SetCellStyle(Sender: TObject; ACol, ARow: Integer; const Style: TCellStyle);
begin
  if (ARow<0) or (ACol<0) then Exit;

  if ARow>High(Grid1CellStyles) then begin
    SetLength(Grid1CellStyles,ARow+1);
  end;
  if ACol>High(Grid1CellStyles[ARow]) then begin
    SetLength(Grid1CellStyles[ARow],Max(ACol,Grid1.ColCount));
  end;

  Grid1CellStyles[ARow,ACol]:=Style;
end;

procedure TForm1.Grid1GetCellText(Sender: TObject; ACol, ARow: Integer; var Text: string);
begin
  if (ARow<0) or (ACol<0) then Exit;
  if (ARow>High(Grid1CellTexts)) or (ACol>High(Grid1CellTexts[ARow])) then Exit;

  Text:=Grid1CellTexts[Arow,ACol];
end;

procedure TForm1.Grid1SetCellText(Sender: TObject; ACol, ARow: Integer; const Text: string);
begin
  if (ARow<0) or (ACol<0) then Exit;

  if ARow>High(Grid1CellTexts) then begin
    SetLength(Grid1CellTexts,ARow+1);
  end;
  if ACol>High(Grid1CellTexts[ARow]) then begin
    SetLength(Grid1CellTexts[ARow],Max(ACol,Grid1.ColCount));
  end;

  Grid1CellTexts[ARow,ACol]:=Text;
end;


procedure TForm1.GridSelectCell(Sender: TObject);
begin
  var Grid:=ActiveGrid;
  StatusText.Text:=Format('Selected Cell: [%d, %d, Left:%d, Top: %d], ViewLeft: %d, ViewTop: %d, ViewBottom: %d',
                          [Grid.SelectedCell.X, Grid.SelectedCell.Y,
                           Grid.ColLefts[Grid.SelectedCell.X],Grid.RowTops[Grid.SelectedCell.Y],
                           Grid.ViewLeft,Grid.ViewTop,Grid.ViewBottom]);
end;

procedure TForm1.GridGridScroll(Sender: TObject; Left, Top: Integer);
begin
  GridSelectCell(nil);
end;

end.

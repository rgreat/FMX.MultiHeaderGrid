unit MainForm;

interface

uses
  System.SysUtils, System.Types, System.UITypes, System.Classes, System.Variants,
  FMX.Types, FMX.Controls, FMX.Forms, FMX.Graphics, FMX.Dialogs, FMX.MultiHeaderGrid,
  FMX.Controls.Presentation,
{$IFDEF MSWINDOWS}
  MidasLib,
{$ENDIF}
  FMX.StdCtrls, FMX.Objects, FMX.Edit, FMX.TabControl, Data.DB, Datasnap.DBClient;

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
    TabGrid: TTabItem;
    TabStringGrid: TTabItem;
    StatusBar1: TStatusBar;
    StatusText: TLabel;
    Grid1: TMultiHeaderGrid;
    Grid2: TMultiHeaderStringGrid;
    Grid3: TMultiHeaderDBGrid;
    RowSelectCheckBox: TCheckBox;
    ButtonFillCells: TButton;
    WordWrapCheckBox: TCheckBox;
    TabDBGrid: TTabItem;
    CDS: TClientDataSet;
    DS: TDataSource;
    Panel2: TPanel;
    ButtonAddRow: TButton;
    ButtonDeleteRow: TButton;
    LimitWidthsCheckBox: TCheckBox;
    LineWidthEdit: TEdit;
    Label2: TLabel;
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
    procedure LimitWidthsCheckBoxChange(Sender: TObject);
    procedure ButtonDeleteRowClick(Sender: TObject);
    procedure ButtonAddRowClick(Sender: TObject);
    procedure TabControl1Change(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure LineWidthEditChange(Sender: TObject);
  private
    Grid1CellTexts: array of array of string;
    Grid1CellStyles: array of array of TCellStyle;

    function ActiveGrid: TMultiHeaderGrid;
  public
    procedure InitGrid(Grid: TMultiHeaderGrid);
    procedure InitDBGrid(Grid: TMultiHeaderDBGrid);
    procedure ApplyGrid3WidthLimits;
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
  InitGrid(Grid1);
  InitGrid(Grid2);
  InitGrid(Grid3);

  TabControl1Change(nil);
end;

procedure TForm1.FormShow(Sender: TObject);
begin
  var GridMax:=Max(Grid1.FullTableWidth,Max(Grid2.FullTableWidth,Grid3.FullTableWidth));

  Width:=Min(GridMax+46,Trunc(Screen.Width-40));

  Height:=Min(800,Trunc(Screen.Height-80));
  Left:=Max(Trunc(Screen.Width-Width) div 2,5);
  Top:=Max(Trunc(Screen.Height-Height) div 2,5);
end;

// ---------------------------------------------------------
// --------------- Create and Fill Grids -------------------
// ---------------------------------------------------------

procedure TForm1.InitGrid(Grid: TMultiHeaderGrid);
begin
  if Grid is TMultiHeaderDBGrid then begin
    InitDBGrid(TMultiHeaderDBGrid(Grid));
  end else begin
    InitGridParams(Grid);
    FillGrid(Grid);
    MergeCellsinGrid(Grid);
    Grid.AutoSize;
  end;
end;

procedure TForm1.InitGridParams(Grid: TMultiHeaderGrid);
begin
  Grid.ColCount:=15;

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

  InitGridHeader(Grid);

  Grid.RowCount:=1000;
end;

procedure TForm1.InitGridHeader(Grid: TMultiHeaderGrid);
begin
  Grid.Header.Clear;
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
// --------------------- Init DBGrid -----------------------
// ---------------------------------------------------------

procedure TForm1.InitDBGrid(Grid: TMultiHeaderDBGrid);
begin
  var Columns:=Grid.Columns;

  CDS.LoadFromFile('example.xml');

  Columns.Clear; // To make our own column sequence

  Columns.SetColumnProps('id',       'Internal'#13#10'identifier');

  Columns.SetColumnProps('klns',     'Service class number',  'Service class');
  Columns.SetColumnProps('klkod',    'Character code',        'Service class');
  Columns.SetColumnProps('klnazv',   'Name',                  'Service class');
  Columns.SetColumnProps('klname',   'Full name',             'Service class');

  Columns.SetColumnProps('kldotpn',  'Start',                 'Valid dates').DateTimeEditor:=dteDate;
  Columns.SetColumnProps('kldotpk',  'End',                   'Valid dates').DateTimeEditor:=dteDate;

  Columns.SetColumnProps('kltip',    'Car type');
  Columns.SetColumnProps('klabd',    'Class code in UZ');
  Columns.SetColumnProps('klsob',    'Carrier network code or 0');
  Columns.SetColumnProps('klntstp',  'Conditional comfort level for class comparison', '');

  Columns.SetColumnProps('klfl05',   'High comfort car',      'Flags;Comfort');
  Columns.SetColumnProps('klpr05',   'Linen on request',      'Flags;Comfort');
  Columns.SetColumnProps('klpr06',   'Business class',        'Flags;Comfort');
  Columns.SetColumnProps('klf2f1',   'Mandatory 4 seats',     'Flags;Order rules');
  Columns.SetColumnProps('klf2f2',   'Single passenger',      'Flags;Order rules');
  Columns.SetColumnProps('klfl04',   'Mandatory 2 seats',     'Flags;Order rules');

  Columns.SetColumnProps('cor_tip',  'Type',                  'Correction');
  Columns.SetColumnProps('cor_time', 'Time',                  'Correction').DisplayFormat:='DD.MM.YYYY HH:NN';
  Columns.SetColumnProps('test_time', 'Test Time');

  for var Column in Columns do begin
    var HMGColumn:=TMHGColumn(Column);
    var Field:=Grid.DataSet.FindField(HMGColumn.FieldName);
    if Field=nil then Continue;

    case Field.DataType of
      ftWideString: HMGColumn.Color:=$FFFFF8DC;
      ftInteger:    HMGColumn.Color:=$FFB3FFFF;
      ftBoolean:    HMGColumn.Color:=$FFFFF6C7;
      ftDate,ftTime,
      ftTimeStamp:  HMGColumn.Color:=$FFE5FFCA;
    end;

    if (Field.DataType=ftWideString) and (Field.Size>10) then begin
      HMGColumn.Alignment:=TTextAlign.Leading
    end else begin
      HMGColumn.Alignment:=TTextAlign.Center;
    end;
  end;

  ApplyGrid3WidthLimits;
  Grid.AutoSize;
end;

procedure TForm1.ApplyGrid3WidthLimits;
const
  Cols : array[0..3] of string  = ('klns', 'klkod', 'klnazv', 'klname');
  MinW : array[0..3] of Integer = (    50,      50,       50,      100);
  MaxW : array[0..3] of Integer = (   150,     150,      150,      350);
begin
  for var i:=0 to High(Cols) do begin
    var Column:=Grid3.Columns.FindByFieldName(Cols[i]);
    if not Assigned(Column) then Continue;

    Column.MinWidth:=IfThen(LimitWidthsCheckBox.IsChecked,MinW[i],0);
    Column.MaxWidth:=IfThen(LimitWidthsCheckBox.IsChecked,MaxW[i],0);
  end;
end;

// ---------------------------------------------------------
// ------------------- Handle Controls ---------------------
// ---------------------------------------------------------

function TForm1.ActiveGrid: TMultiHeaderGrid;
begin
  Result:=nil;
  if TabControl1.ActiveTab=TabGrid then begin
    Result:=Grid1;
  end;
  if TabControl1.ActiveTab=TabStringGrid then begin
    Result:=Grid2;
  end;
  if TabControl1.ActiveTab=TabDBGrid then begin
    Result:=Grid3;
  end;
end;

procedure TForm1.RowCountEditChange(Sender: TObject);
begin
  var Grid:=ActiveGrid;

  Grid.RowCount:=StrToIntDef(RowCountEdit.Text,Grid.RowCount);
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

procedure TForm1.ButtonAddRowClick(Sender: TObject);
begin
  CDS.Insert;
  CDS.FieldByName('id').AsInteger:=Random(MaxInt);
  CDS.Post;
end;

procedure TForm1.ButtonDeleteRowClick(Sender: TObject);
begin
  CDS.Delete;
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
  var Grid:=ActiveGrid;
  Grid.RowSelect:=RowSelectCheckBox.IsChecked;
end;

procedure TForm1.TabControl1Change(Sender: TObject);
begin
  var Grid:=Grid1;
  if TabControl1.ActiveTab=TabStringGrid then begin
    Grid:=Grid2;
  end;
  if TabControl1.ActiveTab=TabDBGrid then begin
    Grid:=Grid3;
  end;

  RowSelectCheckBox.IsChecked:=Grid.RowSelect;
  WordWrapCheckBox.IsChecked:=Grid.WordWrap;
  RowCountEdit.Text:=Grid.RowCount.ToString;
  LineWidthEdit.Text:=Grid.GridLineWidth.ToString;
  RowCountEdit.Enabled:=not (Grid is TMultiHeaderDBGrid);
end;

procedure TForm1.WordWrapCheckBoxChange(Sender: TObject);
begin
  var Grid:=ActiveGrid;
  Grid.WordWrap:=WordWrapCheckBox.IsChecked;
end;

procedure TForm1.LimitWidthsCheckBoxChange(Sender: TObject);
begin
  ApplyGrid3WidthLimits;
end;

procedure TForm1.LineWidthEditChange(Sender: TObject);
begin
  var Grid:=ActiveGrid;

  Grid.GridLineWidth:=StrToFloatDef(LineWidthEdit.Text,Grid.GridLineWidth);
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
  If Pos('Custom Draw',Text)>0 then begin
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
  var Rect:=Grid.GetCellRect(Grid.SelectedCell.X,Grid.SelectedCell.Y);
  StatusText.Text:=Format('Selected Cell: [%d, %d, Left:%d, Top: %d], ViewLeft: %d, ViewTop: %d, ViewBottom: %d. CellRect: [%f, %f, %f, %f]',
                          [Grid.SelectedCell.X, Grid.SelectedCell.Y,
                           Grid.ColLefts[Grid.SelectedCell.X],Grid.RowTops[Grid.SelectedCell.Y],
                           Grid.ViewLeft,Grid.ViewTop,Grid.ViewBottom,
                           Rect.Left,Rect.Top,Rect.Width,Rect.Height]);
end;

procedure TForm1.GridGridScroll(Sender: TObject; Left, Top: Integer);
begin
  GridSelectCell(nil);
end;

end.

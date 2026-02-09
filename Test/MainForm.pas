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
    CheckBox1: TCheckBox;
    ButtonFillCells: TButton;
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
    procedure CheckBox1Change(Sender: TObject);
    procedure GridGridScroll(Sender: TObject; Left, Top: Integer);
    procedure RowCountEditChange(Sender: TObject);
    procedure ButtonFillCellsClick(Sender: TObject);
    procedure GridDblClick(Sender: TObject);
  private
    Grid1CellTexts: array of array of string;
    Grid1CellStyles: array of array of TCellStyle;

    function ActiveGrid: TMultiHeaderGrid;

    procedure InitGridHeader(Grid: TMultiHeaderGrid);
    procedure InitGridParams(Grid: TMultiHeaderGrid);
    procedure FillGrid(Grid: TMultiHeaderGrid);

    procedure InitGrid(Grid: TMultiHeaderGrid);

  public
  end;

var
  Form1: TForm1;

implementation

uses
  System.Diagnostics, System.Math;

{$R *.fmx}

procedure TForm1.FormCreate(Sender: TObject);
begin
  RowCountEdit.Text:='10000';

  TabControl1.ActiveTab:=TabItem1;
  InitGrid(Grid1);
  InitGrid(Grid2);

  ButtonMergeCellsClick(nil);
  ButtonAutoSizeClick(nil);
end;

procedure TForm1.InitGrid(Grid: TMultiHeaderGrid);
begin
  InitGridParams(Grid);
  InitGridHeader(Grid);
  FillGrid(Grid);
end;

procedure TForm1.InitGridParams(Grid: TMultiHeaderGrid);
begin
  Grid.ColCount:=15;
  RowCountEditChange(nil);

  Grid.ColWidths[6] := 100;
  Grid.RowHeights[6] := 80;

  Grid.CellFont.Size := 10;
  Grid.HeaderFont.Size := 11;
  Grid.HeaderFont.Style := [TFontStyle.fsBold];
  Grid.CellColorAlternate := TAlphaColors.Whitesmoke;
  Grid.GridLineColor := TAlphaColors.Gray;
end;

procedure TForm1.InitGridHeader(Grid: TMultiHeaderGrid);
begin
  // Настраиваем многоуровневые заголовки
  var Row:=Grid.HeaderLevels.Add(30);
  var Cell:=Row.Add('Annual report',Grid.ColCount);
  Cell.Style.FontSize:=14;
  Cell.Style.FontStyle:=[TFontStyle.fsBold];
  Cell.Style.CellColor:=TAlphaColors.Lightblue;

  Row:=Grid.HeaderLevels.Add(30);
  Row.Add('First Quarter',3).Style.CellColor:=TAlphaColors.Lightgreen;
  Row.Add('Second Quarter',3).Style.CellColor:=TAlphaColors.Lightgreen;

  Cell:=Row.Add('Summary for'#13#10'1st half of'#13#10+'the year',1,2);
  Cell.Style.FontSize:=14;
  Cell.Style.FontStyle:=[];
  Cell.Style.CellColor:=TAlphaColors.Lightsteelblue;
  Cell.Style.FontColor:=TAlphaColors.Darkred;

  Row.Add('Third Quarter',3).Style.CellColor:=TAlphaColors.Lightgreen;
  Row.Add('Forth Quarter',3).Style.CellColor:=TAlphaColors.Lightgreen;

  Cell:=Row.Add('Summary for'#13#10'2nd half of'#13#10+'the year',1,2);
  Cell.Style.FontSize:=14;
  Cell.Style.FontStyle:=[];
  Cell.Style.CellColor:=TAlphaColors.Lightsteelblue;
  Cell.Style.FontColor:=TAlphaColors.Darkred;

  Cell:=Row.Add('Year'#13#10'Totals',1,2);
  Cell.Style.FontSize:=14;
  Cell.Style.FontStyle:=[];
  Cell.Style.CellColor:=TAlphaColors.Lightsteelblue;
  Cell.Style.FontColor:=TAlphaColors.Darkred;

  Row:=Grid.HeaderLevels.Add(30);
  var FormatSettings:=TFormatSettings.Create('en-us');
  for var i:=1 to 12 do begin
    Row.Add(FormatSettings.LongMonthNames[i]);
  end;
end;

procedure TForm1.FillGrid(Grid: TMultiHeaderGrid);
begin
  for var Y:=0 to Grid.RowCount-1 do begin
    for var X:=0 to Grid.ColCount-1 do begin

      if ((X+Y*2) mod 9 = 3) and (X<Grid.ColCount-1) then begin
        Grid.Cells[X,Y]:='Very Large Cell'#13#10'Number '+(X+Y*7).ToString;
      end else begin
        Grid.Cells[X,Y]:=Format('[%d,%d]', [X, Y]);
      end;

      var CellStyle:=Grid.CellStyle[X,Y];
      if ((X+Y*2) mod 9 = 3) and (X<Grid.ColCount-1) then begin
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

function TForm1.ActiveGrid: TMultiHeaderGrid;
begin
  if TabControl1.ActiveTab=TabItem1 then begin
    Result:=Grid1;
  end else begin
    Result:=Grid2;
  end;
end;

procedure TForm1.ButtonMergeCellsClick(Sender: TObject);
begin
  var Grid:=ActiveGrid;
  Grid.ClearMergedCells;

  Grid.MergeCells(0, 1, 2, 4, 'Merged Cell 1');
  Grid.MergeCells(3, 5, 3, 2, 'Text Block Line #1'#13#10+
                              'Text Block Line #2'#13#10+
                              'Text Block Line #3'#13#10+
                              'TForm1.GridGetCellText(Sender: TObject; ACol, ARow: Integer; var Text: string);');
  Grid.MergeCells(1, 10, 2, 1, 'Merged Cell 2');


  for var T:=11 to Grid.RowCount-5 do begin
    if Random(10)<6 then Continue;

    var W:=Random(Grid.ColCount-3)+1;
    var H:=Random(4)+1;
    var L:=Random(Grid.ColCount-W);

    var Text:='Общая ячейка';
    var TH:=Random(4);
    for var i:=0 to TH do Text:=Text+#13#10+'Строка '+(I+2).ToString;

    Grid.MergeCells(L,T,W,H,Text);
  end;
  Grid.Repaint;
end;

procedure TForm1.ButtonAutoSizeClick(Sender: TObject);
begin
  ActiveGrid.AutoSize;
end;

procedure TForm1.ButtonClearMergeCellsClick(Sender: TObject);
begin
  ActiveGrid.ClearMergedCells;
  ActiveGrid.Repaint;
end;

procedure TForm1.ButtonBenchmarkClick(Sender: TObject);
begin
  var Grid:=ActiveGrid;

  var T:=TStopwatch.StartNew;
  var Cnt:=0;
  var MaxTop:=Max(Trunc(Grid.FullTableHeight-Grid.HeaderHeight-Grid.ViewPortDataHeight),0);
  var MaxLeft:=Max(Trunc(Grid.FullTableWidth-Grid.ViewPortDataWidth),0);
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

procedure TForm1.GridDrawCell(Sender: TObject; ACol, ARow: Integer; Canvas: TCanvas; const Rect: TRectF;
                              IsSelected: boolean; const Text: string; var Handled: Boolean);
begin
  If (ACol=0) and (ARow=1) then begin
    Rect.Inflate(-1,-1);
    if IsSelected then begin
      Canvas.Fill.Color:=TAlphaColorRec.Lightblue;
      Canvas.FillRect(Rect,1);
    end else begin
      Canvas.Fill.Color:=TAlphaColorRec.White;
      Canvas.FillRect(Rect,1);
    end;
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


procedure TForm1.CheckBox1Change(Sender: TObject);
begin
  Grid1.RowSelect:=CheckBox1.IsChecked;
  Grid2.RowSelect:=CheckBox1.IsChecked;
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

procedure TForm1.RowCountEditChange(Sender: TObject);
begin
  Grid1.RowCount:=StrToIntDef(RowCountEdit.Text,Grid1.RowCount);
  Grid2.RowCount:=StrToIntDef(RowCountEdit.Text,Grid2.RowCount);
end;

procedure TForm1.ButtonFillCellsClick(Sender: TObject);
begin
  FillGrid(Grid1);
  FillGrid(Grid2);
  Grid1.Repaint;
  Grid2.Repaint;
end;

procedure TForm1.GridDblClick(Sender: TObject);
begin
  var Grid:=ActiveGrid;
  ShowMessage(Grid.Cells[Grid.Col,Grid.Row]);
end;

end.

unit FMX.MultiHeaderGrid;

interface

uses
  System.Classes, System.Types, System.UITypes, System.Generics.Collections,
  FMX.Types, FMX.Controls, FMX.Graphics, FMX.StdCtrls, FMX.Objects, FMX.Layouts;

type
  THeaderLevel = class;
  THeaderLevels = class;

  TCellStyle = record
  strict private
    FFontStyle              : TFontStyles;
    FFontSize               : Single;
    FCellColor              : TAlphaColor;
    FFontName               : string;
    FFontColor              : TAlphaColor;
    FSelectedCellColor      : TAlphaColor;
    FSelectedFontColor      : TAlphaColor;

    FParentCol              : Integer;
    FParentRow              : Integer;

    FFontStyleIsSet         : Boolean;
    FFontSizeIsSet          : Boolean;
    FCellColorIsSet         : Boolean;
    FFontNameIsSet          : Boolean;
    FFontColorIsSet         : Boolean;
    FSelectedCellColorIsSet : Boolean;
    FSelectedFontColorIsSet : Boolean;
    FIsMergedCell           : Boolean;

    procedure SetCellColor(const Value: TAlphaColor);
    procedure SetFontColor(const Value: TAlphaColor);
    procedure SetFontName(const Value: string);
    procedure SetFontSize(const Value: Single);
    procedure SetFontStyle(const Value: TFontStyles);
    procedure SetSelectedCellColor(const Value: TAlphaColor);
    procedure SetSelectedFontColor(const Value: TAlphaColor);
    procedure SetParentCol(const Value: Integer);
    procedure SetParentRow(const Value: Integer);
  public
    class operator Initialize(out Dest: TCellStyle);

    function IsEmpty: boolean;

    property FontName: string read FFontName write SetFontName;
    property FontSize: Single read FFontSize write SetFontSize;
    property FontStyle: TFontStyles read FFontStyle write SetFontStyle;
    property FontColor: TAlphaColor read FFontColor write SetFontColor;
    property CellColor: TAlphaColor read FCellColor write SetCellColor;
    property SelectedFontColor: TAlphaColor read FSelectedFontColor write SetSelectedFontColor;
    property SelectedCellColor: TAlphaColor read FSelectedCellColor write SetSelectedCellColor;

    property ParentCol: Integer read FParentCol write SetParentCol;
    property ParentRow: Integer read FParentRow write SetParentRow;
    property IsMergedCell: Boolean read FIsMergedCell;
    procedure SetMergedCell(ParenCol,ParenRow: integer);
    procedure ClearMergedCell;

    property FontNameIsSet: Boolean read FFontNameIsSet write FFontNameIsSet;
    property FontSizeIsSet: Boolean read FFontSizeIsSet write FFontSizeIsSet;
    property FontStyleIsSet: Boolean read FFontStyleIsSet write FFontStyleIsSet;
    property FontColorIsSet: Boolean read FFontColorIsSet write FFontColorIsSet;
    property CellColorIsSet: Boolean read FCellColorIsSet write FCellColorIsSet;
    property SelectedCellColorIsSet: Boolean read FSelectedCellColorIsSet write FSelectedCellColorIsSet;
    property SelectedFontColorIsSet: Boolean read FSelectedFontColorIsSet write FSelectedFontColorIsSet;
  end;

  TMergedCell = record
    ColSpan, RowSpan : Integer;
    CellStyle        : TCellStyle;

    function Col: Integer;
    function Row: Integer;
  end;

  THeaderElement = class
  private
    FLevel     : THeaderLevel;
    FCaption   : string;
    FColSpan   : Integer;
    FRowSpan   : Integer;
    FColSkip   : Integer;
    FStyle     : TCellStyle;

    procedure SetCaption(const Value: string);
    procedure SetStyle(const Value: TCellStyle);
  public
    constructor Create(HeaderLevel: THeaderLevel);
    destructor Destroy; reintroduce;

    property Caption: string read FCaption write SetCaption;

    property ColSkip: Integer read FColSkip write FColSkip default 0;
    property ColSpan: Integer read FColSpan write FColSpan default 1;
    property RowSpan: Integer read FRowSpan write FRowSpan default 1;
    property Style: TCellStyle read FStyle write SetStyle;
  end;

  THeaderLevel = class(TObjectList<THeaderElement>)
    FHeight: Integer;
    procedure SetHeight(const Value: Integer);
    function GetColumnsToSkip: TArray<Boolean>;
  public
    FLevels : THeaderLevels;
    constructor Create(HeaderLevels: THeaderLevels);

    function Add(Caption: string = '';
                 ColSpan: Integer = 1;
                 RowSpan: Integer = 1): THeaderElement; overload;


    property Height: Integer read FHeight write SetHeight;
  end;

  TMultiHeaderGrid = class;

  THeaderLevels = class(TObjectList<THeaderLevel>)
    Grid : TMultiHeaderGrid;

    constructor Create(Grid : TMultiHeaderGrid);

    function Add(Heigth: integer = 25): THeaderLevel; overload;
    function GetElementAtCell(ACol, ARow: integer): THeaderElement;
  end;

  TRowData = packed record
    Top    : integer;
    Height : Word;
  end;

  TStartEditingEvent = procedure(Sender: TObject; ACol, ARow: Integer; var InitialChar: Char) of object;
  TDrawCellEvent = procedure(Sender: TObject; ACol, ARow: Integer; Canvas: TCanvas; const Rect: TRectF; IsSelected: boolean; const Text: string; var Handled: Boolean) of object;
  TGetCellTextEvent = procedure(Sender: TObject; ACol, ARow: Integer; var Text: string) of object;
  TSetCellTextEvent = procedure(Sender: TObject; ACol, ARow: Integer; const Text: string) of object;
  TGetCellStyleEvent = procedure(Sender: TObject; ACol, ARow: Integer; var CellStyle: TCellStyle) of object;
  TSetCellStyleEvent = procedure(Sender: TObject; ACol, ARow: Integer; const CellStyle: TCellStyle) of object;
  TColumnsResizedEvent = procedure(Sender: TObject; StartRow, EndRow: integer) of object;
  TRowResizedEvent = procedure(Sender: TObject; ARow: Integer) of object;
  TGridScrollEvent = procedure(Sender: TObject; Left,Top: Integer) of object;

  TResizeMode = (rmNone,rmColumn,rmHeaderRow, rmGridRow);

  TMultiHeaderGrid = class(TControl)
  private
    VScrollBar: TScrollBar;
    HScrollBar: TScrollBar;
    HScrollPanel: TPaintBox;
    CornerPanel: TPanel;

    FColCount: Integer;
    FRowCount: Integer;
    FDefaultColWidth: integer;
    FDefaultRowHeight: integer;
    FColWidths: array of integer;
    FRowData: Array of TRowData;
    FHeaderLevels: THeaderLevels;
    FGridLines: Boolean;
    FGridLineColor: TAlphaColor;
    FGridLineWidth: Single;
    FSelectedCell: TPoint;
    FOnSelectCell: TNotifyEvent;
    FOnDrawCell: TDrawCellEvent;

    FOnGetCellText: TGetCellTextEvent;
    FOnSetCellText: TSetCellTextEvent;
    FOnGetCellStyle: TGetCellStyleEvent;
    FOnSetCellStyle: TSetCellStyleEvent;

    FOnStartEditing: TStartEditingEvent;
    FOnCellClick: TNotifyEvent;
    FOnHeaderClick: TNotifyEvent;

    FHeaderFont: TFont;
    FHeaderFontColor: TAlphaColor;
    FHeaderCellColor: TAlphaColor;

    FCellFont: TFont;

    FSelectedCellColor: TAlphaColor;
    FSelectedFontColor: TAlphaColor;

    FCellColorAlternate: TAlphaColor;
    FCellPadding: TBounds;
    FCellFontColor: TAlphaColor;
    FViewTop: Integer;
    FBackgroundColor: TAlphaColor;
    FViewLeft: Integer;
    FCellColor: TAlphaColor;

    FResizeEnabled: Boolean;
    FResizeHeaderRowEnabled: Boolean;
    FResizeColEnabled: Boolean;
    FResizeRowEnabled: Boolean;

    FResizeStartColumnIndex: Integer;
    FResizeEndColumnIndex: Integer;
    FResizeRowIndex: Integer;
    FResizeMode: TResizeMode;
    FResizeStartPos: TPointF;
    FResizeStartWidths: array of Integer;
    FResizeStartHeight: Integer;
    FResizeMargin: Integer;

    FOnColumnResized: TColumnsResizedEvent;
    FOnHeaderResized: TRowResizedEvent;
    FOnRowResized: TRowResizedEvent;
    FOnGridScroll: TGridScrollEvent;

    FRowSelect: Boolean;

    function ResizeStartWidth: Integer;

    procedure StartCellEditing(ACol, ARow: Integer; InitialChar: Char);
    procedure SetRowCount(Value: Integer);
    procedure SetColCount(Value: Integer);
    procedure SetDefaultColWidth(const Value: integer);
    procedure SetDefaultRowHeight(const Value: integer);
    procedure SetGridLines(const Value: Boolean);
    procedure SetGridLineColor(const Value: TAlphaColor);
    procedure SetGridLineWidth(const Value: Single);
    procedure SetSelectedCell(const Value: TPoint);
    procedure SetCellFont(const Value: TFont);
    procedure SetHeaderFont(const Value: TFont);
    procedure SetCellColorAlternate(const Value: TAlphaColor);
    function GetColLeft(Index: Integer): Integer;
    function GetColWidth(Index: Integer): Integer;
    procedure SetColWidth(Index: Integer; const Value: Integer);
    function GetRowHeight(Index: Integer): Integer;
    procedure SetRowHeight(Index: Integer; const Value: Integer);
    function GetCellRect(ACol, ARow: Integer): TRectF;
    function GetHeaderRect(ALevel, ACol: Integer): TRectF;
    procedure DrawGridLines(Canvas: TCanvas);
    procedure DrawCells(Canvas: TCanvas);
    procedure DrawHeaders(Canvas: TCanvas);
    procedure DrawCell(Canvas: TCanvas; ACol, ARow: Integer; ARect: TRectF);
    procedure DrawHeaderCell(Canvas: TCanvas; ALevel, ACol: Integer);
    procedure SetCellPadding(const Value: TBounds);
    procedure SetHeaderFontColor(const Value: TAlphaColor);
    procedure SetCellFontColor(const Value: TAlphaColor);
    procedure SetViewTop(const Value: Integer);
    procedure VScrollBarChange(Sender: TObject);
    procedure HScrollBarChange(Sender: TObject);

    procedure SetResizeEnabled(const Value: Boolean);
    function GetResizeMargin: Integer;
    procedure SetResizeMargin(const Value: Integer);
    function IsResizeArea(X, Y: Single; out AStartCol, AEndCol, ARow: Integer): TResizeMode;
    procedure StartColumnResize(StartCol, EndCol: Integer; X: Single);
    procedure StartHeaderRowResize(ARow: Integer; Y: Single);
    procedure StartGridRowResize(ARow: Integer; Y: Single);
    procedure UpdateGroupColumnWidth(StartCol, EndCol, TotalWidth: Integer);
    procedure UpdateColumnWidth(StartCol, EndCol: Integer; NewWidth: Integer);
    procedure UpdateHeaderRowHeight(ARow: Integer; NewHeight: Integer);
    procedure UpdateRowHeight(ARow: Integer; NewHeight: Integer);

    function GetCells(ACol, ARow: Integer): string;
    procedure SetCells(ACol, ARow: Integer; const Value: string);
    function GetCellStyle(ACol, ARow: Integer): TCellStyle;
    procedure SetCellStyle(ACol, ARow: Integer; const Value: TCellStyle);

    function GetRowTops(Index: Integer): Integer;
    procedure SetCol(const Value: Integer);
    procedure SetRow(const Value: Integer);
    procedure SetBackgroundColor(const Value: TAlphaColor);
    procedure SetViewLeft(const Value: Integer);
    procedure SetHeaderCellColor(const Value: TAlphaColor);
    procedure SetSelectedCellColor(const Value: TAlphaColor);
    procedure SetSelectedFontColor(const Value: TAlphaColor);
    procedure SetCellColor(const Value: TAlphaColor);
    procedure SetRowSelect(const Value: Boolean);
  protected
    function CanObserve(const ID: Integer): Boolean; override;

    procedure Paint; override;
    procedure Resize; override;
    procedure DoSelectCell; virtual;
    procedure DoDrawCell(ACol, ARow: Integer; Canvas: TCanvas; const Rect: TRectF; IsSelected: boolean; const Text: string; var Handled: Boolean); virtual;

    procedure DoGetCellText(ACol, ARow: Integer; var Text: string); virtual;
    procedure DoSetCellText(ACol, ARow: Integer; const Text: string); virtual;
    procedure DoGetCellStyle(ACol, ARow: Integer; var Style: TCellStyle); virtual;
    procedure DoSetCellStyle(ACol, ARow: Integer; const Style: TCellStyle); virtual;

    procedure DoCellClick(ACol, ARow: Integer); virtual;
    procedure DoHeaderClick(ALevel, ACol: Integer); virtual;

    procedure DoColumnResized; virtual;
    procedure DoHeaderRowResized; virtual;
    procedure DoRowResized; virtual;
    procedure DoGridScroll; virtual;

    procedure MouseDown(Button: TMouseButton; Shift: TShiftState; X, Y: Single); override;
    procedure MouseMove(Shift: TShiftState; X, Y: Single); override;
    procedure MouseUp(Button: TMouseButton; Shift: TShiftState; X, Y: Single); override;

    procedure KeyDown(var Key: Word; var KeyChar: WideChar; Shift: TShiftState); override;

    procedure UpdateSize;
    procedure ScrollToSelectedCell;
  public
    constructor Create(AOwner: TComponent); override;
    destructor Destroy; override;

    function IsMergedCell(ACol, ARow: Integer): Boolean; overload;
    function IsMergedCell(ACol, ARow: Integer; out MergedCell: TMergedCell): Boolean; overload;

    function MergeCells(ACol, ARow, AColSpan, ARowSpan: Integer): Boolean; overload;
    function MergeCells(ACol, ARow, AColSpan, ARowSpan: Integer; const ACaption: string): Boolean; overload;
    procedure UnMergeCells(ACol, ARow: Integer);
    procedure ClearMergedCells;

    procedure AutoSizeRows(ForcePrecise: boolean = False);
    procedure AutoSizeCols(ForcePrecise: boolean = False);
    procedure AutoSize(ForcePrecise: boolean = False);

    property ColLefts[Index: Integer]: Integer read GetColLeft;
    property ColWidths[Index: Integer]: Integer read GetColWidth write SetColWidth;
    property RowHeights[Index: Integer]: Integer read GetRowHeight write SetRowHeight;
    property RowTops[Index: Integer]: Integer read GetRowTops;
    property Cells[ACol, ARow: Integer]: string read GetCells write SetCells;
    property CellStyle[ACol, ARow: Integer]: TCellStyle read GetCellStyle write SetCellStyle;
    property SelectedCell: TPoint read FSelectedCell write SetSelectedCell;
    property HeaderLevels: THeaderLevels read FHeaderLevels;

    property ViewTop: Integer read FViewTop write SetViewTop;
    property ViewLeft: Integer read FViewLeft write SetViewLeft;

    function ViewBottom: Integer;
    function ViewCellsHeight: Integer;
    function HeaderHeight: Integer;
    function ViewPortHeight: Integer;
    function ViewPortDataHeight: Integer;
    function ViewPortDataWidth: Integer;
    function FullTableWidth: Integer;
    function FullTableHeight: Integer;

    function RowAtHeightCoord(Y: Integer): integer;

    property Col: Integer read FSelectedCell.X write SetCol;
    property Row: Integer read FSelectedCell.Y write SetRow;
  published
    property Align;
    property Anchors;
    property ClipChildren;
    property ClipParent;
    property Cursor;
    property DragMode;
    property EnableDragHighlight;
    property Enabled;
    property Locked;
    property Height;
    property HitTest;
    property Margins;
    property Opacity;
    property PopupMenu;
    property Position;
    property RotationAngle;
    property RotationCenter;
    property Scale;
    property Size;
    property TabStop;
    property Visible;
    property Width;

    property ColCount: Integer read FColCount write SetColCount default 5;
    property RowCount: Integer read FRowCount write SetRowCount default 10;
    property DefaultColWidth: integer read FDefaultColWidth write SetDefaultColWidth;
    property DefaultRowHeight: integer read FDefaultRowHeight write SetDefaultRowHeight;
    property BackgroundColor: TAlphaColor read FBackgroundColor write SetBackgroundColor default TAlphaColorRec.White;
    property GridLines: Boolean read FGridLines write SetGridLines default True;
    property GridLineColor: TAlphaColor read FGridLineColor write SetGridLineColor;
    property GridLineWidth: Single read FGridLineWidth write SetGridLineWidth;

    property HeaderFont: TFont read FHeaderFont write SetHeaderFont;
    property HeaderFontColor: TAlphaColor read FHeaderFontColor write SetHeaderFontColor default TAlphaColorRec.Black;
    property HeaderCellColor: TAlphaColor read FHeaderCellColor write SetHeaderCellColor default TAlphaColorRec.Lightgray;

    property CellFont: TFont read FCellFont write SetCellFont;
    property CellFontColor: TAlphaColor read FCellFontColor write SetCellFontColor default TAlphaColorRec.Black;
    property CellColor: TAlphaColor read FCellColor write SetCellColor default TAlphaColorRec.White;
    property CellColorAlternate: TAlphaColor read FCellColorAlternate write SetCellColorAlternate default TAlphaColorRec.Lightblue;

    property SelectedFontColor: TAlphaColor read FSelectedFontColor write SetSelectedFontColor default TAlphaColorRec.Black;
    property SelectedCellColor: TAlphaColor read FSelectedCellColor write SetSelectedCellColor default TAlphaColorRec.Lightblue;

    property RowSelect: Boolean read FRowSelect write SetRowSelect default False;

    property CellPadding: TBounds read FCellPadding write SetCellPadding;

    property OnSelectCell: TNotifyEvent read FOnSelectCell write FOnSelectCell;
    property OnDrawCell: TDrawCellEvent read FOnDrawCell write FOnDrawCell;

    property OnGetCellText: TGetCellTextEvent read FOnGetCellText write FOnGetCellText;
    property OnSetCellText: TSetCellTextEvent read FOnSetCellText write FOnSetCellText;
    property OnGetCellStyle: TGetCellStyleEvent read FOnGetCellStyle write FOnGetCellStyle;
    property OnSetCellStyle: TSetCellStyleEvent read FOnSetCellStyle write FOnSetCellStyle;

    property OnStartEditing: TStartEditingEvent read FOnStartEditing write FOnStartEditing;
    property OnCellClick: TNotifyEvent read FOnCellClick write FOnCellClick;
    property OnHeaderClick: TNotifyEvent read FOnHeaderClick write FOnHeaderClick;

    property ResizeEnabled: Boolean read FResizeEnabled write SetResizeEnabled default True;
    property ResizeRowEnabled: Boolean read FResizeRowEnabled write FResizeRowEnabled default True;
    property ResizeColEnabled: Boolean read FResizeColEnabled write FResizeColEnabled default True;
    property ResizeHeaderRowEnabled: Boolean read FResizeHeaderRowEnabled write FResizeHeaderRowEnabled default True;

    property ResizeMargin: Integer read GetResizeMargin write SetResizeMargin default 3;

    property OnColumnResized: TColumnsResizedEvent read FOnColumnResized write FOnColumnResized;
    property OnHeaderResized: TRowResizedEvent read FOnHeaderResized write FOnHeaderResized;
    property OnRowResized: TRowResizedEvent read FOnRowResized write FOnRowResized;
    property OnGridScroll: TGridScrollEvent read FOnGridScroll write FOnGridScroll;
  end;

  TMultiHeaderStringGrid = class(TMultiHeaderGrid)
  private
    FCellTexts: array of array of string;
    FCellStyles: array of array of TCellStyle;
  protected
    procedure DoGetCellText(ACol, ARow: Integer; var Text: string); override;
    procedure DoSetCellText(ACol, ARow: Integer; const Text: string); override;
    procedure DoGetCellStyle(ACol, ARow: Integer; var Style: TCellStyle); override;
    procedure DoSetCellStyle(ACol, ARow: Integer; const Style: TCellStyle); override;
  public
    constructor Create(AOwner: TComponent); override;
    destructor Destroy; override;
  end;


procedure Register;

implementation

uses
  System.SysUtils, System.Math, FMX.Platform;

procedure Register;
begin
  RegisterComponents('FMX', [TMultiHeaderGrid]);
  RegisterComponents('FMX', [TMultiHeaderStringGrid]);
end;

{ THeaderElement }

constructor THeaderElement.Create(HeaderLevel: THeaderLevel);
begin
  FLevel:=HeaderLevel;
  FColSpan:=1;
  FRowSpan:=1;
  HeaderLevel.Add(Self);
end;

destructor THeaderElement.Destroy;
begin
  inherited;
end;

procedure THeaderElement.SetCaption(const Value: string);
begin
  if FCaption<>Value then begin
    FCaption:=Value;
  end;
end;

procedure THeaderElement.SetStyle(const Value: TCellStyle);
begin
  FStyle:=Value;
end;

{ THeaderLevel }

function THeaderLevel.Add(Caption: string = '';
                          ColSpan: Integer = 1;
                          RowSpan: Integer = 1): THeaderElement;
begin
//  var SkipCols:=GetColumnsToSkip;
  Result:=THeaderElement.Create(Self);
  Result.Caption:=Caption;
  Result.ColSpan:=ColSpan;
  Result.RowSpan:=RowSpan;

  // Расчет текуещей колонки
  var CurCol:=0;
  for var Element in Self do begin
    inc(CurCol,Element.FColSpan);
  end;

  var CurRow:=0;
  for var Level in FLevels do begin
    if Level=Self then Break;
    inc(CurRow);
  end;

  // Расчет пропуска столбцов
  Result.FColSkip:=0;
  for var Row:=0 to FLevels.Count-1 do begin
    var Level:=FLevels[Row];
    if Level=Self then Break;
    var Col:=0;
    for var Element in Level do begin
      inc(Col,Element.FColSpan);
      if Col+Element.FColSkip>CurCol then Break;

      if Row+Element.FRowSpan>CurRow then begin
        inc(Result.FColSkip,Element.FColSpan);
      end;
    end;
  end;
end;

constructor THeaderLevel.Create(HeaderLevels: THeaderLevels);
begin
  inherited Create;

  FLevels:=HeaderLevels;
  HeaderLevels.Add(Self);
end;

procedure THeaderLevel.SetHeight(const Value: Integer);
begin
  FHeight:=Value;
end;

function THeaderLevel.GetColumnsToSkip: TArray<Boolean>;
var
  i, j, k, Col: Integer;
  Element: THeaderElement;
begin
  // Создаем массив для отслеживания занятых колонок
  SetLength(Result, FLevels.Grid.ColCount);
  for i := 0 to High(Result) do begin
    Result[i] := False;
  end;

  var LevelIndex:=FLevels.IndexOf(Self);

  // Проходим по всем уровням выше текущего
  for i := 0 to LevelIndex - 1 do begin
    for j := 0 to FLevels[i].Count - 1 do begin
      Element := FLevels[i][j];
      // Если элемент занимает несколько строк и его RowSpan распространяется на текущий уровень
      if (Element.RowSpan > 1) and (i + Element.RowSpan > LevelIndex) then begin
        // Помечаем все колонки, которые занимает этот элемент
        for k := 0 to Element.ColSpan - 1 do begin
          Col := Element.ColSkip + k;
          if Col < Length(Result) then begin
            Result[Col] := True;
          end;
        end;
      end;
    end;
  end;
end;

{ THeaderLevels }

function THeaderLevels.Add(Heigth: integer = 25): THeaderLevel;
begin
  Result:=THeaderLevel.Create(Self);
  Result.FHeight:=Heigth;
end;

constructor THeaderLevels.Create(Grid: TMultiHeaderGrid);
begin
  inherited Create;

  Self.Grid:=Grid;
end;

function THeaderLevels.GetElementAtCell(ACol, ARow: integer): THeaderElement;
begin
  Result:=nil;
  for var Row:=ARow downto 0 do begin
    var Level:=Items[Row];
    var Left:=0;
    for var Element in Level do begin
      if (ACol>=Left+Element.ColSkip) and (ACol<Left+Element.ColSpan+Element.ColSkip) and
         (ARow>=Row) and (ARow<Row+Element.RowSpan) then Exit(Element);
      inc(Left,Element.ColSpan);
    end;
  end;
end;

{ TMergedCell }

function TMergedCell.Col: Integer;
begin
  Result:=CellStyle.ParentCol;
end;

function TMergedCell.Row: Integer;
begin
  Result:=CellStyle.ParentRow;
end;

{ TCustomMultiHeaderGrid }

constructor TMultiHeaderGrid.Create(AOwner: TComponent);
var
 I: integer;
begin
  inherited;

  FColCount:=5;
  FRowCount:=10;
  FDefaultColWidth:=80;
  FDefaultRowHeight:=20;
  FGridLines:=True;
  FGridLineColor:=TAlphaColors.Gray;
  FGridLineWidth:=1;
  FSelectedCell:=Point(-1, -1);

  VScrollBar:=TScrollBar.Create(Self);
  VScrollBar.Orientation:=TOrientation.Vertical;
  VScrollBar.Parent:=Self;
  VScrollBar.OnChange:=VScrollBarChange;
  VScrollBar.Align:=TAlignLayout.Right;
  VScrollBar.Stored:=False;

  HScrollBar:=TScrollBar.Create(Self);
  HScrollBar.Orientation:=TOrientation.Horizontal;
  HScrollBar.Stored:=False;
  HScrollBar.Margins.Top:=1;

  HScrollPanel:=TPaintBox.Create(Self);
  HScrollPanel.Parent:=Self;
  HScrollPanel.Align:=TAlignLayout.MostBottom;
  HScrollBar.OnChange:=HScrollBarChange;
  HScrollPanel.Height:=HScrollBar.Height+1;
  HScrollPanel.Stored:=False;

  CornerPanel:=TPanel.Create(Self);
  CornerPanel.Parent:=HScrollPanel;
  CornerPanel.Align:=TAlignLayout.Right;
  CornerPanel.Height:=VScrollBar.Width;
  CornerPanel.Width:=HScrollBar.Height;
  CornerPanel.Stored:=False;

  HScrollBar.Parent:=HScrollPanel;
  HScrollBar.Align:=TAlignLayout.Client;

  FColWidths:=[];
  FRowData:=[];

  FCellFont:=TFont.Create;
  FCellFontColor:=TAlphaColors.Black;
  FCellColor:=TAlphaColors.White;
  FCellColorAlternate:=$FFEDFAFF;

  FHeaderFont:=TFont.Create;
  FHeaderFontColor:=TAlphaColors.Black;
  FHeaderCellColor:=TAlphaColors.Lightgray;

  FBackgroundColor:=TAlphaColors.White;

  FSelectedFontColor:=TAlphaColorRec.Black;
  FSelectedCellColor:=TAlphaColorRec.Lightblue;

  CanFocus:=True;

  CellPadding:=TBounds.Create(TRectF.Create(2,1,2,1));

  // Инициализация ширины колонок и высоты строк
  SetLength(FColWidths,FColCount);
  for I:=0 to FColCount-1 do begin
    FColWidths[i]:=FDefaultColWidth;
  end;

  var Top:=0;
  SetLength(FRowData,FRowCount);
  for I:=0 to FRowCount-1 do begin
    FRowData[I].Top:=Top;
    FRowData[I].Height:=trunc(FDefaultRowHeight);
    Top:=Top+trunc(FDefaultRowHeight);
  end;

  HitTest:=True;
  TabStop:=True;

  FHeaderLevels:=THeaderLevels.Create(Self);

  FResizeEnabled:=True;
  FResizeHeaderRowEnabled:=True;
  FResizeColEnabled:=True;
  FResizeRowEnabled:=True;

  FResizeStartColumnIndex:=-1;
  FResizeEndColumnIndex:=-1;
  FResizeRowIndex:=-1;
  FResizeMargin:=3;
end;

destructor TMultiHeaderGrid.Destroy;
begin
  ReleaseCapture;

  FHeaderLevels.Free;
  FCellFont.Free;
  FHeaderFont.Free;
  CellPadding.Free;

  VScrollBar.Free;
  HScrollBar.Free;
  CornerPanel.Free;
  HScrollPanel.Free;

  inherited;
end;

procedure TMultiHeaderGrid.SetCol(const Value: Integer);
begin
  if (FRowCount<=0) or (FColCount<=0) then Exit;

  if FSelectedCell.X=Value then Exit;

  FSelectedCell.X:=Value;
  if FSelectedCell.Y<0 then FSelectedCell.Y:=0;
  ScrollToSelectedCell;
  DoSelectCell;
end;

procedure TMultiHeaderGrid.SetColCount(Value: Integer);
var
  I: Integer;
begin
  if Value<0 then Value:=0;
  if FColCount<>Value then begin
    FColCount:=Value;
    SetLength(FColWidths,FColCount);
    for I:=0 to FColCount-1 do begin
      FColWidths[i]:=FDefaultColWidth;
    end;
    UpdateSize;
    InvalidateRect(LocalRect);

    HScrollBar.Max:=Value;
  end;
end;

procedure TMultiHeaderGrid.SetRow(const Value: Integer);
begin
  if (FRowCount<=0) or (FColCount<=0) then Exit;

  if FSelectedCell.Y=Value then Exit;

  FSelectedCell.Y:=Value;
  if FSelectedCell.X<0 then FSelectedCell.X:=0;
  ScrollToSelectedCell;
  DoSelectCell;
end;

procedure TMultiHeaderGrid.SetRowCount(Value: Integer);
begin
  if Value<0 then Value:=0;

  if FRowCount<>Value then begin
    var OldValue:=FRowCount;

    if (Value<FRowCount) and (Value>0) then begin
      for var I:=0 to FColCount-1 do begin
        UnMergeCells(i,Value-1);
      end;
    end;

    FRowCount:=Value;
    var Top:=0;
    if OldValue>0 then Top:=FRowData[OldValue-1].Top;
    SetLength(FRowData,FRowCount);
    for var I:=Max(OldValue-1,0) to FRowCount-1 do begin
      FRowData[I].Top:=Top;
      if i>=OldValue then begin
        FRowData[I].Height:=trunc(FDefaultRowHeight);
      end;
      Top:=Top+FRowData[I].Height;
    end;

    UpdateSize;
    InvalidateRect(LocalRect);

    VScrollBar.Max:=FullTableHeight-HeaderHeight;
    VScrollBar.ViewportSize:=ViewPortDataHeight;
  end;
end;

procedure TMultiHeaderGrid.SetDefaultColWidth(const Value: integer);
var
  I: Integer;
begin
  if FDefaultColWidth<>Value then
  begin
    FDefaultColWidth:=Value;
    SetLength(FColWidths,FColCount);
    for I:=0 to FColCount-1 do begin
      FColWidths[i]:=FDefaultColWidth;
    end;
    UpdateSize;
    InvalidateRect(LocalRect);
  end;
end;

procedure TMultiHeaderGrid.SetDefaultRowHeight(const Value: integer);
begin
  if FDefaultRowHeight<>Value then begin
    FDefaultRowHeight:=Value;

    var Top:=0;
    SetLength(FRowData,FRowCount);
    for var I:=0 to FRowCount-1 do begin
      FRowData[I].Top:=Top;
      FRowData[I].Height:=trunc(FDefaultRowHeight);
      Top:=Top+trunc(FDefaultRowHeight);
    end;

    UpdateSize;
    InvalidateRect(LocalRect);
  end;
end;

function TMultiHeaderGrid.GetColLeft(Index: Integer): Integer;
begin
  Result:=0;
  for var i:=0 to Index-1 do
    inc(Result,FColWidths[i]);
end;

function TMultiHeaderGrid.GetColWidth(Index: Integer): Integer;
begin
  if (Index>=0) and (Index<Length(FColWidths)) then
    Result:=FColWidths[Index]
  else
    Result:=FDefaultColWidth;
end;

function TMultiHeaderGrid.FullTableHeight: Integer;
begin
  if FRowCount=0 then Exit(0);

  var BottomCell:=FRowData[High(FRowData)];

  Result:=BottomCell.Top+BottomCell.Height+HeaderHeight;
end;

function TMultiHeaderGrid.FullTableWidth: Integer;
begin
  Result:=round(2*GridLineWidth);
  for var i:=0 to FColCount-1 do begin
    Result:=Result+FColWidths[i];
  end;
end;

procedure TMultiHeaderGrid.SetColWidth(Index: Integer; const Value: Integer);
begin
  if (Index>=0) and (Index<Length(FColWidths)) and (FColWidths[Index]<>Value) then begin
    FColWidths[Index]:=Value;
    UpdateSize;
    InvalidateRect(LocalRect);
  end;
end;

function TMultiHeaderGrid.GetRowHeight(Index: Integer): Integer;
begin
  if (Index>=0) and (Index<Length(FRowData)) then
    Result:=FRowData[Index].Height
  else
    Result:=FDefaultRowHeight;
end;

function TMultiHeaderGrid.GetRowTops(Index: Integer): Integer;
begin
  if (Index>=0) and (Index<Length(FRowData)) then
    Result:=FRowData[Index].Top
  else
    Result:=0;
end;

function TMultiHeaderGrid.ViewBottom: Integer;
begin
  Result:=ViewTop+ViewCellsHeight;
end;

procedure TMultiHeaderGrid.SetRowHeight(Index: Integer; const Value: Integer);
var
  I : integer;
begin
  if (Index>=0) and (Index<Length(FRowData)) and (FRowData[Index].Height<>Value) then
  begin
    FRowData[Index].Height:=Trunc(Value);
    for I:=Index+1 to FRowCount-1 do begin
      FRowData[I].Top:=FRowData[I-1].Top+FRowData[I-1].Height;
    end;
    UpdateSize;
    InvalidateRect(LocalRect);
  end;
end;

function TMultiHeaderGrid.GetCells(ACol, ARow: Integer): string;
begin
  Result:='';
  DoGetCellText(ACol, ARow, Result);
end;

function TMultiHeaderGrid.GetCellStyle(ACol, ARow: Integer): TCellStyle;
begin
  Result:=Default(TCellStyle);
  DoGetCellStyle(ACol, ARow, Result);
end;

function TMultiHeaderGrid.GetCellRect(ACol, ARow: Integer): TRectF;
var
  X, Y: Single;
  I: Integer;
begin
  // Проверяем валидность диапазона
  if (ACol<0) or (ARow<0) or (ACol>=FColCount) or (ARow>=FRowCount) then
    Exit(Default(TRectF));

  X:=FGridLineWidth-ViewLeft;
  Y:=FGridLineWidth+HeaderHeight-ViewTop;

  // Вычисляем позицию колонки
  for I:=0 to ACol-1 do
    X:=X+GetColWidth(I);

  // Вычисляем позицию строки
  Y:=Y+FRowData[ARow].Top;

  Result:=RectF(X, Y, X+GetColWidth(ACol), Y+GetRowHeight(ARow));
end;

function TMultiHeaderGrid.HeaderHeight: Integer;
begin
  Result:=round(FGridLineWidth);

  // Добавляем высоту заголовков
  for var I:=0 to FHeaderLevels.Count-1 do
    Result:=Result+FHeaderLevels[I].Height;
end;

function TMultiHeaderGrid.GetHeaderRect(ALevel, ACol: Integer): TRectF;
var
  X, Y          : Single;
  I             : Integer;
  Element       : THeaderElement;
  ActualColSpan : Integer;
  ActualRowSpan : Integer;
begin
  X:=FGridLineWidth-ViewLeft;
  Y:=FGridLineWidth;

  // Вычисляем позицию уровня заголовка
  for i:=0 to ALevel-1 do begin
    Y:=Y+FHeaderLevels[i].Height;
  end;

  // Вычисляем позицию колонки
  for I:=0 to ACol-1 do
    X:=X+GetColWidth(I);

  Element:=FHeaderLevels.GetElementAtCell(ACol,ALevel);

  if not Assigned(Element) then Exit(RectF(-1,-1,-1,-1));
  if FHeaderLevels.IndexOf(Element.FLevel)<>ALevel then Exit;

  // Определяем фактический ColSpan (не больше оставшихся колонок)
  ActualColSpan:=Element.ColSpan;
  if ACol+ActualColSpan>FColCount then
    ActualColSpan:=FColCount-ACol;

  // Вычисляем ширину ячейки заголовка
  Result:=RectF(X, Y, X, Y+ FHeaderLevels[ALevel].Height);
  for I:=0 to ActualColSpan-1 do
    Result.Right:=Result.Right+GetColWidth(ACol+I);

  // Определяем фактический RowSpan (не больше оставшихся строк)
  ActualRowSpan:=Element.RowSpan;
  if ALevel+ActualRowSpan>FHeaderLevels.Count then
    ActualRowSpan:=FHeaderLevels.Count-ALevel;

  // Вычисляем высоту ячейки заголовка
  for I:=1 to ActualRowSpan-1 do begin
    Result.Bottom:=Result.Bottom+FHeaderLevels[ALevel+I].Height;
  end;
end;

procedure TMultiHeaderGrid.UpdateSize;
begin
  if (FColCount+FRowCount>0) and (ViewPortHeight<FullTableHeight) then begin
    VScrollBar.Visible:=True;
    CornerPanel.Visible:=True;
    VScrollBar.Max:=FullTableHeight-HeaderHeight;
    VScrollBar.Value:=ViewTop;
    VScrollBar.ViewportSize:=ViewPortDataHeight;
    VScrollBar.SmallChange:=DefaultRowHeight;
    VScrollBar.OnChange:=VScrollBarChange;
  end else begin
    VScrollBar.Visible:=False;
    CornerPanel.Visible:=False;
    VScrollBar.OnChange:=nil;
  end;

  if (FColCount>0) and (FullTableWidth>ViewPortDataWidth) then begin
    HScrollPanel.Visible:=True;
    HScrollBar.Max:=FullTableWidth;
    HScrollBar.Value:=ViewLeft;
    HScrollBar.ViewportSize:=ViewPortDataWidth;
    HScrollBar.SmallChange:=FDefaultColWidth;
    HScrollBar.OnChange:=HScrollBarChange;
  end else begin
    HScrollPanel.Visible:=False;
    HScrollBar.OnChange:=nil;
  end;
end;

function TMultiHeaderGrid.ViewCellsHeight: Integer;
begin
  Result:=Round(LocalRect.Height-Margins.Bottom-HeaderHeight);
  if HScrollPanel.Visible then Result:=Round(Result-HScrollPanel.Height);
end;

function TMultiHeaderGrid.ViewPortDataHeight: Integer;
begin
  Result:=ViewPortHeight-HeaderHeight;
end;

function TMultiHeaderGrid.ViewPortDataWidth: Integer;
begin
  Result:=Round(LocalRect.Width);
  if VScrollBar.Visible then Result:=Round(Result-VScrollBar.Width);
end;

function TMultiHeaderGrid.ViewPortHeight: Integer;
begin
  Result:=Round(LocalRect.Height-2);
  if HScrollPanel.Visible then Result:=Round(Result-HScrollPanel.Height);
end;

procedure TMultiHeaderGrid.VScrollBarChange(Sender: TObject);
begin
  ViewTop:=Round(VScrollBar.Value);
  InvalidateRect(LocalRect);
end;

procedure TMultiHeaderGrid.HScrollBarChange(Sender: TObject);
begin
  ViewLeft:=Round(HScrollBar.Value);
  InvalidateRect(LocalRect);
end;

procedure TMultiHeaderGrid.Paint;
var
  Canvas: TCanvas;
  I: Integer;
  ResizeRect: TRectF;
  CellRect: TRectF;
begin
  Canvas:=Self.Canvas;

  if Canvas.BeginScene then begin
    var Save:=Canvas.SaveState;
    try
      // Устанавливаем область обрезки
      var R:=TRectF.Create(LocalRect.Left, LocalRect.Top,
                           LocalRect.Left+LocalRect.Width, LocalRect.Top+LocalRect.Height);
      if VScrollBar.Visible then R.Right:=R.Right-VScrollBar.Width;
      if HScrollPanel.Visible then R.Bottom:=R.Bottom-HScrollPanel.Height;
      Canvas.IntersectClipRect(R);

      Canvas.Fill.Kind:=TBrushKind.Solid;
      Canvas.Fill.Color:=FBackgroundColor;

      // Очищаем фон
      Canvas.FillRect(LocalRect, 0, 0, AllCorners, 1, Canvas.Fill);

      // Рисуем ячейки
      DrawCells(Canvas);

      // Рисуем линии сетки
      if FGridLines then
        DrawGridLines(Canvas);

      // Рисуем заголовки
      DrawHeaders(Canvas);
    finally
      Canvas.RestoreState(Save);
      Canvas.EndScene;
    end;
  end;

  inherited;
end;

procedure TMultiHeaderGrid.DrawHeaders(Canvas: TCanvas);
var
  I, J: Integer;
  Rect: TRectF;
begin
  Canvas.Stroke.Kind:=TBrushKind.Solid;
  Canvas.Stroke.Color:=FGridLineColor;
  Canvas.Stroke.Thickness:=FGridLineWidth;

  for i:=0 to FHeaderLevels.Count-1 do begin
    J:=0;
    for var Element in FHeaderLevels[i] do begin
      DrawHeaderCell(Canvas, I, J+Element.ColSkip);
      J:=J+Element.ColSpan;
    end;
  end;
end;


procedure TMultiHeaderGrid.DrawHeaderCell(Canvas: TCanvas; ALevel, ACol: Integer);
var
  Element : THeaderElement;
begin
  Element:=FHeaderLevels.GetElementAtCell(ACol,ALevel);
  if not Assigned(Element) then Exit;

  var ARect:=GetHeaderRect(ALevel, ACol);

  // Заливка фона
  if Element.Style.CellColorIsSet then begin
    // Задано явно для ячейки
    Canvas.Fill.Color:=Element.Style.CellColor;
  end else begin
    Canvas.Fill.Color:=FHeaderCellColor;
  end;
  Canvas.FillRect(ARect, 0, 0, AllCorners, 1);

  // Тескт
  Canvas.Font.Assign(FCellFont);

  if Element.Style.FontNameIsSet then begin
    Canvas.Font.Family:=Element.Style.FontName;
  end;
  if Element.Style.FontSizeIsSet then begin
    Canvas.Font.Size:=Element.Style.FontSize;
  end;
  if Element.Style.FontStyleIsSet then begin
    Canvas.Font.Style:=Element.Style.FontStyle;
  end;

  if Element.Style.FontColorIsSet then begin
    Canvas.Fill.Color:=Element.Style.FontColor;
  end else begin
    Canvas.Fill.Color:=FCellFontColor;
  end;
  Canvas.FillText(ARect, Element.Caption, False, 1, [], TTextAlign.Center, TTextAlign.Center);

  // Границы
  Canvas.DrawRect(ARect, 0, 0, AllCorners, 1);
end;

procedure TMultiHeaderGrid.DrawCells(Canvas: TCanvas);
var
  I, J       : Integer;
  Rect       : TRectF;
  MergedCell : TMergedCell;
begin
  var TopRow:=RowAtHeightCoord(ViewTop);
  var ViewBottomCell:=ViewBottom;

  if TopRow<0 then Exit;

  for J:=TopRow to FRowCount-1 do begin
    if FRowData[J].Top>ViewBottomCell then Break;
    for I:=0 to FColCount-1 do begin

      // Пропускаем объединенные ячейки (кроме первой)
      if IsMergedCell(I, J, MergedCell) then begin
        var FirstVisibleRow:=Max(MergedCell.Row,TopRow);

        if (I=MergedCell.Col) and (J=FirstVisibleRow) then begin
          Rect:=GetCellRect(MergedCell.Col, MergedCell.Row);
          // Корректируем прямоугольник для объединенной ячейки
          Rect.Right:=Rect.Left;
          Rect.Bottom:=Rect.Top;
          for var K:=0 to MergedCell.ColSpan-1 do
            Rect.Right:=Rect.Right+GetColWidth(I+K);
          for var K:=0 to MergedCell.RowSpan-1 do
            Rect.Bottom:=Rect.Bottom+GetRowHeight(J+K);

          DrawCell(Canvas, MergedCell.Col, MergedCell.Row, Rect);
        end;
        Continue;
      end;

      Rect:=GetCellRect(I, J);
      DrawCell(Canvas, I, J, Rect);
    end;
  end;
end;

procedure TMultiHeaderGrid.DrawCell(Canvas: TCanvas; ACol, ARow: Integer; ARect: TRectF);
var
  Handled    : Boolean;
  Text       : string;
  MergedCell : TMergedCell;
  IsSelected : Boolean;
begin
  Handled:=False;

  // Проверяем, является ли ячейка частью объединенной
  if IsMergedCell(ACol, ARow, MergedCell) then begin
    // Рисуем только первую ячейку объединенного блока
    if (ACol<>MergedCell.Col) or (ARow<>MergedCell.Row) then Exit;

    // Проверяем, выделена ли основная ячейка этого блока

    IsSelected:=((MergedCell.Col<=FSelectedCell.X) and (MergedCell.Col+MergedCell.ColSpan>FSelectedCell.X) or FRowSelect) and
                 (MergedCell.Row<=FSelectedCell.Y) and (MergedCell.Row+MergedCell.RowSpan>FSelectedCell.Y);

    // Выделенная ячейка
    if IsSelected then
      Canvas.Fill.Color:=TAlphaColors.Lightblue;

    // Корректируем прямоугольник для объединенной ячейки
    ARect.Right:=ARect.Left;
    ARect.Bottom:=ARect.Top;
    for var I:=0 to MergedCell.ColSpan-1 do
      ARect.Right:=ARect.Right+GetColWidth(ACol+I);
    for var I:=0 to MergedCell.RowSpan-1 do
      ARect.Bottom:=ARect.Bottom+GetRowHeight(ARow+I);

    // Текст
    Text:=Cells[MergedCell.Col,MergedCell.Row];
  end else begin
    // Обычная ячейка

    IsSelected:=((ACol=FSelectedCell.X) or FRowSelect) and (ARow=FSelectedCell.Y);

    Canvas.FillRect(ARect, 0, 0, AllCorners, 1);

    // Текст
    Text:=Cells[ACol, ARow];
  end;

  DoDrawCell(ACol, ARow, Canvas, ARect, IsSelected, Text, Handled);

  if not Handled then begin
    var CellStyle:=CellStyle[ACol, ARow];

    // Заливка
    if IsSelected then begin
      // Выделенная ячейка
      if CellStyle.SelectedCellColorIsSet then begin
        Canvas.Fill.Color:=CellStyle.SelectedCellColor;
      end else begin
        Canvas.Fill.Color:=FSelectedCellColor;
      end;
    end else begin
      // Обычная ячейка
      if CellStyle.CellColorIsSet then begin
        // Задано явно для ячейки
        Canvas.Fill.Color:=CellStyle.CellColor;
      end else begin
        // Чередование цветов строк
        if (ARow mod 2=1) and (FCellColorAlternate<>TAlphaColors.Null) then begin
          Canvas.Fill.Color:=FCellColorAlternate;
        end else begin
          Canvas.Fill.Color:=FCellColor;
        end;
      end;
    end;
    Canvas.FillRect(ARect, 0, 0, AllCorners, 1);

    // Тескт
    Canvas.Font.Assign(FCellFont);

    if CellStyle.FontNameIsSet then begin
      Canvas.Font.Family:=CellStyle.FontName;
    end;
    if CellStyle.FontSizeIsSet then begin
      Canvas.Font.Size:=CellStyle.FontSize;
    end;
    if CellStyle.FontStyleIsSet then begin
      Canvas.Font.Style:=CellStyle.FontStyle;
    end;

    if IsSelected then begin
      // Выделенная ячейка
      if CellStyle.SelectedFontColorIsSet then begin
        Canvas.Fill.Color:=CellStyle.SelectedFontColor;
      end else begin
        Canvas.Fill.Color:=FCellFontColor;
      end;
    end else begin
      // Обычная ячейка
      if CellStyle.FontColorIsSet then begin
        Canvas.Fill.Color:=CellStyle.FontColor;
      end else begin
        Canvas.Fill.Color:=FCellFontColor;
      end;
    end;

    // Корректировка области вывод для текста
    ARect.Left:=ARect.Left+CellPadding.Left+FGridLineWidth;
    ARect.Top:=ARect.Top+CellPadding.Top+FGridLineWidth;
    ARect.Right:=ARect.Right-CellPadding.Right-FGridLineWidth;
    ARect.Bottom:=ARect.Bottom-CellPadding.Bottom-FGridLineWidth;
    Canvas.FillText(ARect, Text, False, 1, [], TTextAlign.Center, TTextAlign.Center);
  end;
end;

procedure TMultiHeaderGrid.DrawGridLines(Canvas: TCanvas);
var
  I, J: Integer;
  X, Y: Single;
  StartX: Single;
  StartY: Single;
  MergedCell: TMergedCell;
begin
  if not FGridLines then Exit;
  if (FRowCount=0) or (FColCount=0) then Exit;

  Canvas.Stroke.Kind:=TBrushKind.Solid;
  Canvas.Stroke.Color:=FGridLineColor;
  Canvas.Stroke.Thickness:=FGridLineWidth;

  // Вычисляем начальную Y-координату для данных (после заголовков)
  StartX:=FGridLineWidth-ViewLeft;
  StartY:=FGridLineWidth+HeaderHeight-ViewTop;

  var ViewBottomCell:=ViewBottom;

  var TopRow:=RowAtHeightCoord(ViewTop);
  if TopRow<0 then Exit;

  // Рисуем линии для данных (как раньше)
  // Сначала рисуем все вертикальные линии
  X:=StartX;
  for I:=0 to FColCount do begin
    // Внутренние вертикальные линии
    for J:=TopRow to FRowCount-1 do begin
      Y:=StartY+FRowData[J].Top;
      if FRowData[J].Top>ViewBottomCell then Break;

      // Проверяем, не находится ли линия внутри объединенной ячейки
      var ShouldDraw:=True;

      // Проверяем ячейку слева
      if IsMergedCell(I-1, J, MergedCell) then begin
        if I<MergedCell.Col+MergedCell.ColSpan then
          ShouldDraw:=False;
      end;

      // Проверяем ячейку справа
      if ShouldDraw and IsMergedCell(I, J, MergedCell) then begin
        if I>MergedCell.Col then
          ShouldDraw:=False;
      end;

      if ShouldDraw then begin
        Canvas.DrawLine(PointF(X, Y), PointF(X, Y+GetRowHeight(J)), 1);
      end;
    end;

    if I<FColCount then
      X:=X+GetColWidth(I);
  end;

  // Затем рисуем все горизонтальные линии
  for I:=TopRow to FRowCount do begin
    if I=0 then begin
      Y:=StartY+FRowData[I].Top;
    end else begin
      Y:=StartY+FRowData[I-1].Top+FRowData[I-1].Height;
    end;

    if (I>0) and (FRowData[I-1].Top>ViewBottomCell) then Break;

    X:=StartX;
    for J:=0 to FColCount-1 do begin
      // Проверяем, не находится ли линия внутри объединенной ячейки
      var ShouldDraw:=True;
      // Внутренние горизонтальные линии

      // Проверяем ячейку сверху
      if IsMergedCell(J, I-1, MergedCell) then begin
        if I<MergedCell.Row+MergedCell.RowSpan then begin
          ShouldDraw:=False;
        end;
      end;

      // Проверяем ячейку снизу
      if ShouldDraw and IsMergedCell(J, I, MergedCell) then begin
        if I>MergedCell.Row then begin
          ShouldDraw:=False;
        end;
      end;

      if ShouldDraw then begin
        Canvas.DrawLine(PointF(X, Y), PointF(X+GetColWidth(J), Y), 1);
      end;

      X:=X+GetColWidth(J);
    end;
  end;
end;

function TMultiHeaderGrid.MergeCells(ACol, ARow, AColSpan, ARowSpan: Integer): Boolean;
begin
  // Проверяем валидность диапазона
  if (ACol<0) or (ARow<0) or (ACol+AColSpan>FColCount) or (ARow+ARowSpan>FRowCount) then
    Exit(False);

  for var Y:=ARow to ARow+ARowSpan-1 do begin
    for var X:=ACol to ACol+AColSpan-1 do begin
      if IsMergedCell(X,Y) then begin
        Exit(False);
      end;
    end;
  end;

  for var Y:=ARow to ARow+ARowSpan-1 do begin
    for var X:=ACol to ACol+AColSpan-1 do begin
      var Style:=CellStyle[X,Y];
      Style.SetMergedCell(ACol,ARow);
      CellStyle[X,Y]:=Style;
    end;
  end;

  Result:=True;
end;

function TMultiHeaderGrid.MergeCells(ACol, ARow, AColSpan, ARowSpan: Integer; const ACaption: string): Boolean;
begin
  Result:=MergeCells(ACol,ARow,AColSpan,ARowSpan);
  if not Result then Exit;

  Cells[ACol, ARow]:=ACaption;
end;

procedure TMultiHeaderGrid.UnmergeCells(ACol, ARow: Integer);
var
  MergedCell: TMergedCell;
begin
  var Style:=CellStyle[ACol,ARow];
  if not Style.IsMergedCell then Exit;

  for var Y:=Style.ParentRow to FRowCount-1 do begin
    var IsMergedCell:=False;
    for var X:=Style.ParentCol to FColCount-1 do begin
      var TmpStyle:=CellStyle[X,Y];
      IsMergedCell:=TmpStyle.IsMergedCell and (TmpStyle.ParentCol=Style.ParentCol) and (TmpStyle.ParentRow=Style.ParentRow);
      if IsMergedCell then begin
        TmpStyle.ClearMergedCell;
        CellStyle[X,Y]:=TmpStyle;
      end else begin
        if X>Style.ParentCol then begin
          Break;
        end else begin
          Exit;
        end;
      end;
    end;
  end;
end;

procedure TMultiHeaderGrid.ClearMergedCells;
begin
  for var Y:=0 to FRowCount-1 do begin
    for var X:=0 to FColCount-1 do begin
      var TmpStyle:=CellStyle[X,Y];
      if TmpStyle.IsMergedCell then begin
        TmpStyle.ClearMergedCell;
        CellStyle[X,Y]:=TmpStyle;
      end;
    end;
  end;
end;

function TMultiHeaderGrid.IsMergedCell(ACol, ARow: Integer): Boolean;
var
  MergedCell: TMergedCell;
begin
  Result:=IsMergedCell(ACol, ARow, MergedCell);
end;

function TMultiHeaderGrid.IsMergedCell(ACol, ARow: Integer; out MergedCell: TMergedCell): Boolean;
begin
  var Style:=CellStyle[ACol,ARow];
  Result:=Style.IsMergedCell;

  if not Result then begin
    MergedCell:=Default(TMergedCell);
    Exit;
  end;

  MergedCell.CellStyle:=CellStyle[Style.ParentCol,Style.ParentRow];
  MergedCell.ColSpan:=1;
  MergedCell.RowSpan:=1;

  for var X:=Style.ParentCol+1 to FColCount-1 do begin
    var TmpStyle:=CellStyle[X,Style.ParentRow];
    var IsMergedCell:=TmpStyle.IsMergedCell and (TmpStyle.ParentCol=Style.ParentCol) and (TmpStyle.ParentRow=Style.ParentRow);
    if not IsMergedCell then Break;
    inc(MergedCell.ColSpan);
  end;

  for var Y:=Style.ParentRow+1 to FRowCount-1 do begin
    var TmpStyle:=CellStyle[Style.ParentCol,Y];
    var IsMergedCell:=TmpStyle.IsMergedCell and (TmpStyle.ParentCol=Style.ParentCol) and (TmpStyle.ParentRow=Style.ParentRow);
    if not IsMergedCell then Break;
    inc(MergedCell.RowSpan);
  end;
end;

procedure TMultiHeaderGrid.AutoSize(ForcePrecise: boolean = False);
begin
  AutoSizeCols(ForcePrecise);
  AutoSizeRows(ForcePrecise);
  UpdateSize;
  InvalidateRect(LocalRect);
end;

function CountLines(const Text: string): Integer;
var
  PosStart, PosFound: Integer;
begin
  if Text='' then begin
    Result:=0;
    Exit;
  end;

  Result:=1; // Минимум одна строка
  PosStart:=1;

  while True do begin
    PosFound:=Pos(#13#10, Text, PosStart);
    if PosFound=0 then Break;

    Inc(Result);
    PosStart:=PosFound+2; // Пропускаем #13#10
  end;
end;

function GetMaxLineLength(const Text: string): Integer;
var
  StartPos, EndPos, LineLength: Integer;
begin
  if Text='' then begin
    Result:=0;
    Exit;
  end;

  Result:=0;
  StartPos:=1;
  while StartPos<=Length(Text) do begin
    // Ищем конец строки
    EndPos:=Pos(#13#10, Text, StartPos);

    if EndPos = 0 then begin
      // Это последняя строка
      LineLength:=Length(Text)-StartPos+1;
      if LineLength>Result then
        Result:=LineLength;
      Break;
    end else begin
      // Вычисляем длину текущей строки
      LineLength:=EndPos-StartPos;
      if LineLength>Result then
        Result:=LineLength;

      // Переходим к следующей строке
      StartPos:=EndPos+1;

      // Пропускаем возможный #10 после #13
      if (EndPos<=Length(Text)) and (Text[EndPos]=#13) and
         (StartPos<=Length(Text)) and (Text[StartPos]=#10) then
        Inc(StartPos);
    end;
  end;
end;

function GetMaxLine(const Text: string): string;
var
  StartPos, EndPos, LineLength: Integer;
begin
  if Text='' then begin
    Exit;
  end;

  StartPos:=1;
  while StartPos<=Length(Text) do begin
    // Ищем конец строки
    EndPos:=Pos(#13#10, Text, StartPos);

    if EndPos = 0 then begin
      // Это последняя строка
      LineLength:=Length(Text)-StartPos+1;
      if LineLength>Length(Result) then
        Result:=Copy(Text,StartPos,LineLength);
      Break;
    end else begin
      // Вычисляем длину текущей строки
      LineLength:=EndPos-StartPos;
      if LineLength>Length(Result) then
        Result:=Copy(Text,StartPos,LineLength);

      // Переходим к следующей строке
      StartPos:=EndPos+1;

      // Пропускаем возможный #10 после #13
      if (EndPos<=Length(Text)) and (Text[EndPos]=#13) and
         (StartPos<=Length(Text)) and (Text[StartPos]=#10) then
        Inc(StartPos);
    end;
  end;
end;

procedure TMultiHeaderGrid.AutoSizeCols(ForcePrecise: boolean = False);
var
  I, J: Integer;
  MaxWidth: Single;
  Text: string;
begin
  var CellPaddingWidth:=CellPadding.Left+CellPadding.Right;
  var CellDelimterWidth:=2*FGridLineWidth;
  var CellPaddingFull:=CellPaddingWidth+CellDelimterWidth;
  var FastMode:=not ForcePrecise and (FRowCount*FColCount>10000);

  for I:=0 to FColCount-1 do begin
    MaxWidth:=0;

    // Проверяем заголовки
    for J:=0 to FHeaderLevels.Count-1 do begin
      var Element:=FHeaderLevels.GetElementAtCell(I,J);
      if not Assigned(Element) then Continue;

      var Lines:=Element.Caption.Split([#13#10]);

      Canvas.Font.Assign(FCellFont);
      if Element.Style.FontNameIsSet then begin
        Canvas.Font.Family:=Element.Style.FontName;
      end;
      if Element.Style.FontSizeIsSet then begin
        Canvas.Font.Size:=Element.Style.FontSize;
      end;
      if Element.Style.FontStyleIsSet then begin
        Canvas.Font.Style:=Element.Style.FontStyle;
      end;

      for var Line in Lines do begin
        MaxWidth:=Max(MaxWidth,Canvas.TextWidth(Line)/Element.FColSpan-(Element.FColSpan-1)*CellDelimterWidth);
      end;
    end;

    Canvas.Font.Assign(Self.CellFont);
    Canvas.Font.Assign(FCellFont);
    var FLetterWidth:=Canvas.TextWidth('V');

    // Проверяем ячейки

    var MaxLetters:=0.0;
    for J:=0 to FRowCount-1 do begin
      var ColSpan:=1;
      var MergedCell: TMergedCell;
      if IsMergedCell(I,J,MergedCell) then begin
        Text:=Cells[MergedCell.Col,MergedCell.Row];
        ColSpan:=MergedCell.ColSpan;
      end else begin
        Text:=Cells[I, J];
      end;

      if FastMode then begin
        var Letters:=GetMaxLineLength(Text)/ColSpan;
        if Letters>MaxLetters then begin
          MaxLetters:=Letters;
          var Line:=GetMaxLine(Text);

          var CellStyle:=CellStyle[I, J];
          Canvas.Font.Assign(FCellFont);
          if CellStyle.FontNameIsSet then begin
            Canvas.Font.Family:=CellStyle.FontName;
          end;
          if CellStyle.FontSizeIsSet then begin
            Canvas.Font.Size:=CellStyle.FontSize;
          end;
          if CellStyle.FontStyleIsSet then begin
            Canvas.Font.Style:=CellStyle.FontStyle;
          end;

          MaxWidth:=Max(MaxWidth, Canvas.TextWidth(Line)/ColSpan-(ColSpan-1)*CellDelimterWidth);
        end;
      end else begin
        var CellStyle:=CellStyle[I, J];
        Canvas.Font.Assign(FCellFont);
        if CellStyle.FontNameIsSet then begin
          Canvas.Font.Family:=CellStyle.FontName;
        end;
        if CellStyle.FontSizeIsSet then begin
          Canvas.Font.Size:=CellStyle.FontSize;
        end;
        if CellStyle.FontStyleIsSet then begin
          Canvas.Font.Style:=CellStyle.FontStyle;
        end;

        var Lines:=Text.Split([#13#10]);
        for var Line in Lines do begin
          MaxWidth:=Max(MaxWidth, Canvas.TextWidth(Line)/ColSpan-(ColSpan-1)*CellDelimterWidth);
        end;
      end;
    end;

    // Добавляем отступы
    MaxWidth:=MaxWidth+CellPaddingFull;

    // Устанавливаем новую ширину
    FColWidths[I]:=Trunc(MaxWidth);
  end;
end;

procedure TMultiHeaderGrid.AutoSizeRows(ForcePrecise: boolean = False);
begin
  Canvas.Font.Assign(FCellFont);

  var CellPaddingHeight:=CellPadding.Top+CellPadding.Bottom;
  var CellDelimterHeight:=2*FGridLineWidth;
  var CellPaddingHeightFull:=CellPaddingHeight+CellDelimterHeight/2;
  var FastMode:=not ForcePrecise and (FRowCount*FColCount>10000);


  if FastMode then begin
    Canvas.Font.Assign(FCellFont);
    var TH:=Canvas.TextHeight('А');
    for var Row:=0 to FRowCount-1 do begin
      var MaxHeight:=TH+CellDelimterHeight;
      for var Col:=0 to FColCount-1 do begin
        var RowSpan:=1;
        var MergedCell: TMergedCell;
        var TextLines: Single;
        if IsMergedCell(Col,Row,MergedCell) then begin
          TextLines:=CountLines(Cells[MergedCell.Col,MergedCell.Row]);
          RowSpan:=MergedCell.RowSpan;
        end else begin
          TextLines:=CountLines(Cells[Col,Row]);
        end;

        var CellStyle:=CellStyle[Col,Row];
        var ResTH:=TH;
        if CellStyle.FontSizeIsSet then begin
          ResTH:=ResTH/Canvas.Font.Size*CellStyle.FontSize;
        end;
        if CellStyle.FontStyleIsSet then begin
          if (TFontStyle.fsBold in CellStyle.FontStyle) xor (TFontStyle.fsBold in Canvas.Font.Style) then begin
            if TFontStyle.fsBold in CellStyle.FontStyle then begin
              ResTH:=ResTH*1.1;
            end else begin
              ResTH:=ResTH*0.9;
            end;
          end;
        end;

        MaxHeight:=Max(MaxHeight,TextLines*ResTH/RowSpan-(RowSpan-1)*CellDelimterHeight);
      end;
      FRowData[Row].Height:=Trunc(MaxHeight+CellPaddingHeightFull);
    end;
  end else begin
    for var Row:=0 to FRowCount-1 do begin
      var MaxHeight:=0.0;
      for var Col:=0 to FColCount-1 do begin
        var RowSpan:=1;
        var MergedCell: TMergedCell;
        var TextLines: Single;
        var TH: Single;
        if IsMergedCell(Col,Row,MergedCell) then begin
          TextLines:=CountLines(Cells[MergedCell.Col,MergedCell.Row]);
          RowSpan:=MergedCell.RowSpan;
          TH:=Canvas.TextHeight('А');
        end else begin
          TextLines:=CountLines(Cells[Col,Row]);

          var CellStyle:=CellStyle[Col,Row];
          Canvas.Font.Assign(FCellFont);
          if CellStyle.FontNameIsSet then begin
            Canvas.Font.Family:=CellStyle.FontName;
          end;
          if CellStyle.FontSizeIsSet then begin
            Canvas.Font.Size:=CellStyle.FontSize;
          end;
          if CellStyle.FontStyleIsSet then begin
            Canvas.Font.Style:=CellStyle.FontStyle;
          end;
          TH:=Canvas.TextHeight('А');
        end;

        MaxHeight:=Max(MaxHeight,TextLines*TH/RowSpan-(RowSpan-1)*CellDelimterHeight);
      end;

      FRowData[Row].Height:=Trunc(MaxHeight+CellPaddingHeightFull);
    end;
  end;

  for var Row:=1 to FRowCount-1 do begin
    FRowData[Row].Top:=FRowData[Row-1].Top+FRowData[Row-1].Height;
  end;
end;

procedure TMultiHeaderGrid.ScrollToSelectedCell;
begin
  if (FSelectedCell.X>=0) and (FSelectedCell.Y>=0) then begin

    var Rect: TRectF;
    var MergedCell: TMergedCell;
    if IsMergedCell(FSelectedCell.X,FSelectedCell.Y,MergedCell) then begin
      Rect:=GetCellRect(MergedCell.Col, MergedCell.Row);
      // Корректируем прямоугольник для объединенной ячейки
      Rect.Right:=Rect.Left;
      Rect.Bottom:=Rect.Top;
      for var K:=0 to MergedCell.ColSpan-1 do
        Rect.Right:=Rect.Right+GetColWidth(MergedCell.Col+K);
      for var K:=0 to MergedCell.RowSpan-1 do
        Rect.Bottom:=Rect.Bottom+GetRowHeight(MergedCell.Row+K);
    end else begin
      Rect:=GetCellRect(FSelectedCell.X,FSelectedCell.Y);
    end;
    Rect.Offset(ViewLeft,ViewTop-HeaderHeight);

    if Rect.Left<ViewLeft then begin
      ViewLeft:=Trunc(Rect.Left-FGridLineWidth);
    end;

    if Rect.Right>ViewLeft+ViewPortDataWidth then begin
      ViewLeft:=Trunc(Rect.Right-ViewPortDataWidth+FGridLineWidth);
    end;

    if Rect.Top<ViewTop then begin
      ViewTop:=Trunc(Rect.Top);
    end;

    if (Rect.Bottom>ViewBottom-2*FGridLineWidth) then begin
      ViewTop:=Round(Rect.Bottom-ViewCellsHeight-2*FGridLineWidth);
    end;
  end;

  InvalidateRect(LocalRect);
end;


procedure TMultiHeaderGrid.MouseDown(Button: TMouseButton; Shift: TShiftState; X, Y: Single);
var
  StartCol, EndCol : Integer;
  Row              : Integer;
  I, J             : Integer;
  Rect             : TRectF;
  ClickedHeader    : Boolean;
  MergedCell       : TMergedCell;
begin
  inherited;

 // Устанавливаем фокус при клике
  if CanFocus and (Button=TMouseButton.mbLeft) then begin
    SetFocus; // Устанавливаем фокус
    InvalidateRect(LocalRect);
  end;

  if Button=TMouseButton.mbLeft then begin
    ClickedHeader:=False;

    if FResizeEnabled then begin
      case IsResizeArea(X, Y, StartCol, EndCol, Row) of
        TResizeMode.rmColumn: begin
          StartColumnResize(StartCol, EndCol, X);

          // Захватываем мышь для продолжения перетаскивания за пределами компонента
          Capture;
          Exit;
        end;
        TResizeMode.rmHeaderRow: begin
          StartHeaderRowResize(Row, Y);

          // Захватываем мышь для продолжения перетаскивания за пределами компонента
          Capture;
          Exit;
        end;
        TResizeMode.rmGridRow: begin
          StartGridRowResize(Row, Y);

          // Захватываем мышь для продолжения перетаскивания за пределами компонента
          Capture;
          Exit;
        end;
      end;
    end;

    // Проверяем клик по заголовкам
    for I:=0 to FHeaderLevels.Count-1 do begin
      J:=0;
      while J<FColCount do begin
        Rect:=GetHeaderRect(I, J);
        if Rect.Contains(PointF(X, Y)) then begin
          // Проверяем, является ли это объединенной ячейкой
          DoHeaderClick(I, J);
          ClickedHeader:=True;
          Break;
        end;
        J:=J+FHeaderLevels.GetElementAtCell(j,I).ColSpan;
      end;
      if ClickedHeader then
        Break;
    end;

    // Если не кликнули по заголовку, проверяем ячейки данных
    if not ClickedHeader then begin

      var TopRow:=RowAtHeightCoord(ViewTop);
      if TopRow<0 then Exit;
      var ViewBottom:=ViewTop+LocalRect.Top+LocalRect.Height+Margins.Bottom-HeaderHeight;

      for J:=TopRow to FRowCount-1 do begin
        if FRowData[J].Top>ViewBottom then Break;
        for I:=0 to FColCount-1 do begin
          Rect:=GetCellRect(I, J);
          if Rect.Contains(PointF(X, Y)) then begin
            FSelectedCell:=Point(I, J);
            DoCellClick(FSelectedCell.X, FSelectedCell.Y);
            ScrollToSelectedCell;
            DoSelectCell;
            Exit;
          end;
        end;
      end;
    end;
  end;
end;

procedure TMultiHeaderGrid.MouseMove(Shift: TShiftState; X, Y: Single);
var
  Col, Row            : Integer;
  NewWidth, NewHeight : Integer;
  GlobalPos           : TPointF;
begin
  inherited;

  if FResizeEnabled then begin
    case FResizeMode of
      TResizeMode.rmColumn: begin
        if FResizeEndColumnIndex>=0 then begin
          // Используем глобальные координаты для точного отслеживания
          GlobalPos := LocalToAbsolute(PointF(X, Y));
          NewWidth := ResizeStartWidth + Round(GlobalPos.X - FResizeStartPos.X);
          UpdateColumnWidth(FResizeStartColumnIndex,FResizeEndColumnIndex, NewWidth);
          Cursor := crHSplit;
        end;
      end;
      TResizeMode.rmHeaderRow: begin
        if FResizeRowIndex>=0 then begin
          // Используем глобальные координаты для точного отслеживания
          GlobalPos := LocalToAbsolute(PointF(X, Y));
          NewHeight := FResizeStartHeight + Round(GlobalPos.Y - FResizeStartPos.Y);
          UpdateHeaderRowHeight(FResizeRowIndex, NewHeight);
          Cursor := crVSplit;
        end;
      end;
      TResizeMode.rmGridRow: begin
        if FResizeRowIndex>=0 then begin
          // Используем глобальные координаты для точного отслеживания
          GlobalPos := LocalToAbsolute(PointF(X, Y));
          NewHeight := FResizeStartHeight + Round(GlobalPos.Y - FResizeStartPos.Y);
          UpdateRowHeight(FResizeRowIndex, NewHeight);
          Cursor := crVSplit;
        end;
      end
      else begin
        case IsResizeArea(X, Y, Col, Col, Row) of
          TResizeMode.rmColumn: Cursor := crHSplit;
          TResizeMode.rmHeaderRow: Cursor := crVSplit;
          TResizeMode.rmGridRow: Cursor := crVSplit;
          else Cursor := crDefault;
        end;
      end;
    end;
  end else begin
    Cursor := crDefault;
  end;
end;

procedure TMultiHeaderGrid.MouseUp(Button: TMouseButton; Shift: TShiftState; X, Y: Single);
begin
  inherited;

  if (Button = TMouseButton.mbLeft) then begin
    // Завершаем изменение размера столбца
    if FResizeMode<>TResizeMode.rmNone then begin
      DoColumnResized;
      FResizeStartColumnIndex:=-1;
      FResizeEndColumnIndex:=-1;
      FResizeRowIndex:=-1;
      FResizeMode:=TResizeMode.rmNone;
      Cursor := crDefault;
    end;
  end;
end;

procedure CopyTextToClipboard(const AText: string);
var
  ClipboardService: IFMXClipboardService;
begin
  if TPlatformServices.Current.SupportsPlatformService(IFMXClipboardService, IInterface(ClipboardService)) then
  begin
    ClipboardService.SetClipboard(AText);
  end
  else
    raise Exception.Create('Сервис буфера обмена не поддерживается');
end;

procedure TMultiHeaderGrid.KeyDown(var Key: Word; var KeyChar: WideChar; Shift: TShiftState);
var
  NewCol, NewRow   : Integer;
  OldCol, OldRow   : Integer;
  MergedCellOld    : TMergedCell;
  MergedCellNew    : TMergedCell;
  Handled          : Boolean;
begin
  Handled:=False;

  if IsFocused then
  begin
    var DX:=0;
    var DY:=0;

    case Key of
      vkLeft: begin
        DX:=-1;
        Handled:=True;
      end;
      vkRight: begin
        DX:=1;
        Handled:=True;
      end;
      vkUp: begin
        DY:=-1;
        Handled:=True;
      end;
      vkDown: begin
        DY:=1;
        Handled:=True;
      end;
      vkHome: begin
        DY:=-RowCount;
        Handled:=True;
      end;
      vkEnd: begin
        DY:=RowCount;
        Handled:=True;
      end;
      vkPrior: begin  // Page Up
        DY:=-Trunc(ViewCellsHeight/DefaultRowHeight);
        Handled:=True;
      end;
      vkNext: begin   // Page Down
        DY:=Trunc(ViewCellsHeight/DefaultRowHeight);
        Handled:=True;
      end;
      vkReturn: begin // Enter
        DoCellClick(FSelectedCell.X, FSelectedCell.Y);
        Exit;
      end;
      vkEscape: begin
        FSelectedCell:=Point(-1, -1);
        DoSelectCell;
        InvalidateRect(LocalRect);
        Exit;
      end;
      vkC,vkInsert: begin
        if (Row>=0) and (Col>=0) then begin
          var MergedCell: TMergedCell;
          if IsMergedCell(Col,Row,MergedCell) then begin
            CopyTextToClipboard(Cells[MergedCell.Col,MergedCell.Row]);
          end else begin
            CopyTextToClipboard(Cells[Col, Row]);
          end;
        end;
        Exit;
      end;
    end;

    NewCol:=FSelectedCell.X;
    NewRow:=FSelectedCell.Y;
    repeat
      OldCol:=NewCol;
      OldRow:=NewRow;

      NewCol:=Max(0, Min(FColCount-1, NewCol+DX));
      NewRow:=Max(0, Min(FRowCount-1, NewRow+DY));

      if (NewCol<>FSelectedCell.X) or (NewRow<>FSelectedCell.Y) then begin

        if IsMergedCell(FSelectedCell.X, FSelectedCell.Y, MergedCellOld) and
           IsMergedCell(NewCol, NewRow, MergedCellNew) and
           (MergedCellOld.Col=MergedCellNew.Col) and
           (MergedCellOld.Row=MergedCellNew.Row) then Continue;

        Break;
      end;
    until (OldCol=NewCol) and (OldRow=NewRow);

    if (NewCol<>FSelectedCell.X) or (NewRow<>FSelectedCell.Y) then begin
      FSelectedCell:=Point(NewCol, NewRow);

      ScrollToSelectedCell;
      DoSelectCell;
    end;
  end;

  // Обработка ввода текста
  if IsFocused and (KeyChar>=' ') and (KeyChar<='~') and not (ssCtrl in Shift) then begin
    StartCellEditing(FSelectedCell.X, FSelectedCell.Y, KeyChar);
    KeyChar:=#0;
  end;
end;

procedure TMultiHeaderGrid.Resize;
begin
  inherited;

  UpdateSize;
end;

function TMultiHeaderGrid.RowAtHeightCoord(Y: Integer): integer;
var
  L, R, M: Integer;
  CurrentTop, CurrentBottom: Integer;
begin
  Result:=-1;

  if Length(FRowData) = 0 then
    Exit;

  // Быстрая проверка граничных случаев
  if Y < FRowData[0].Top then
    Exit;

  // Проверка последнего элемента
  CurrentBottom:=FRowData[High(FRowData)].Top + FRowData[High(FRowData)].Height;
  if Y >= CurrentBottom then
    Exit(High(FRowData)); // или -1, в зависимости от требований

  // Бинарный поиск с целочисленной арифметикой
  L:=0;
  R:=High(FRowData);

  while L <= R do
  begin
    M:=(L + R) shr 1; // Быстрее, чем div 2

    CurrentTop:=FRowData[M].Top;
    CurrentBottom:=CurrentTop + FRowData[M].Height;

    if Y >= CurrentTop then
    begin
      if Y < CurrentBottom then
      begin
        Result:=M;
        Exit;
      end
      else
        L:=M + 1;
    end
    else
      R:=M - 1;
  end;
end;

procedure TMultiHeaderGrid.DoSelectCell;
begin
  if Assigned(FOnSelectCell) then
    FOnSelectCell(Self);
end;

procedure TMultiHeaderGrid.DoDrawCell(ACol, ARow: Integer; Canvas: TCanvas; const Rect: TRectF; IsSelected: boolean; const Text: string; var Handled: Boolean);
begin
  if Assigned(FOnDrawCell) then
    FOnDrawCell(Self,ACol,ARow,Canvas,Rect,IsSelected,Text,Handled);
end;

procedure TMultiHeaderGrid.DoGetCellText(ACol, ARow: Integer; var Text: string);
begin
  if Assigned(FOnGetCellText) then
    FOnGetCellText(Self, ACol, ARow, Text);
end;

procedure TMultiHeaderGrid.DoSetCellText(ACol, ARow: Integer; const Text: string);
begin
  if Assigned(FOnSetCellText) then
    FOnSetCellText(Self, ACol, ARow, Text);
end;

procedure TMultiHeaderGrid.DoGetCellStyle(ACol, ARow: Integer; var Style: TCellStyle);
begin
  if Assigned(FOnGetCellStyle) then
    FOnGetCellStyle(Self, ACol, ARow, Style);
end;

procedure TMultiHeaderGrid.DoSetCellStyle(ACol, ARow: Integer; const Style: TCellStyle);
begin
  if Assigned(FOnSetCellStyle) then
    FOnSetCellStyle(Self, ACol, ARow, Style);
end;


procedure TMultiHeaderGrid.DoCellClick(ACol, ARow: Integer);
begin
  if Assigned(FOnCellClick) then
    FOnCellClick(Self);
end;

procedure TMultiHeaderGrid.DoHeaderClick(ALevel, ACol: Integer);
begin
  if Assigned(FOnHeaderClick) then
    FOnHeaderClick(Self);
end;

procedure TMultiHeaderGrid.SetGridLines(const Value: Boolean);
begin
  if FGridLines<>Value then begin
    FGridLines:=Value;
    InvalidateRect(LocalRect);
  end;
end;

procedure TMultiHeaderGrid.SetGridLineColor(const Value: TAlphaColor);
begin
  if FGridLineColor<>Value then begin
    FGridLineColor:=Value;
    InvalidateRect(LocalRect);
  end;
end;

procedure TMultiHeaderGrid.SetGridLineWidth(const Value: Single);
begin
  if FGridLineWidth<>Value then begin
    FGridLineWidth:=Value;
    InvalidateRect(LocalRect);
  end;
end;

procedure TMultiHeaderGrid.SetSelectedCell(const Value: TPoint);
begin
  if (FSelectedCell.X<>Value.X) or (FSelectedCell.Y<>Value.Y) then begin
    FSelectedCell:=Value;
    DoSelectCell;
    InvalidateRect(LocalRect);
  end;
end;

procedure TMultiHeaderGrid.SetSelectedCellColor(const Value: TAlphaColor);
begin
  if FSelectedCellColor<>Value then begin
    FSelectedCellColor:=Value;
    InvalidateRect(LocalRect);
  end;
end;

procedure TMultiHeaderGrid.SetSelectedFontColor(const Value: TAlphaColor);
begin
  if FSelectedFontColor<>Value then begin
    FSelectedFontColor:=Value;
    InvalidateRect(LocalRect);
  end;
end;

procedure TMultiHeaderGrid.SetCellFont(const Value: TFont);
begin
  if FCellFont<>Value then begin
    FCellFont.Assign(Value);
    InvalidateRect(LocalRect);
  end;
end;

procedure TMultiHeaderGrid.SetCellFontColor(const Value: TAlphaColor);
begin
  if FCellFontColor<>Value then begin
    FCellFontColor:=Value;
    InvalidateRect(LocalRect);
  end;
end;

procedure TMultiHeaderGrid.SetCellPadding(const Value: TBounds);
begin
  if FCellPadding<>Value then begin
    FCellPadding:=Value;
    InvalidateRect(LocalRect);
  end;
end;

procedure TMultiHeaderGrid.SetCells(ACol, ARow: Integer; const Value: string);
begin
  DoSetCellText(ACol, ARow, Value);
end;

procedure TMultiHeaderGrid.SetCellStyle(ACol, ARow: Integer; const Value: TCellStyle);
begin
  DoSetCellStyle(ACol, ARow, Value);
end;

procedure TMultiHeaderGrid.SetCellColor(const Value: TAlphaColor);
begin
  if FCellColor<>Value then begin
    FCellColor:=Value;
    InvalidateRect(LocalRect);
  end;
end;

procedure TMultiHeaderGrid.SetHeaderFont(const Value: TFont);
begin
  if FHeaderFont<>Value then begin
    FHeaderFont.Assign(Value);
    InvalidateRect(LocalRect);
  end;
end;

procedure TMultiHeaderGrid.SetHeaderCellColor(const Value: TAlphaColor);
begin
  if FHeaderCellColor<>Value then begin
    FHeaderCellColor:=Value;
    InvalidateRect(LocalRect);
  end;
end;

procedure TMultiHeaderGrid.SetHeaderFontColor(const Value: TAlphaColor);
begin
  if FHeaderFontColor<>Value then begin
    FHeaderFontColor:=Value;
    InvalidateRect(LocalRect);
  end;
end;

procedure TMultiHeaderGrid.SetCellColorAlternate(const Value: TAlphaColor);
begin
  if FCellColorAlternate<>Value then begin
    FCellColorAlternate:=Value;
    InvalidateRect(LocalRect);
  end;
end;

procedure TMultiHeaderGrid.SetBackgroundColor(const Value: TAlphaColor);
begin
  if FBackgroundColor<>Value then begin
    FBackgroundColor:=Value;
    InvalidateRect(LocalRect);
  end;
end;

procedure TMultiHeaderGrid.SetViewLeft(const Value: integer);
begin
  var LastColLeft:=FullTableWidth-FColWidths[FColCount-1];
  FViewLeft:=Min(LastColLeft,Max(0,Value));

  HScrollBar.Value:=Value;

  DoGridScroll;

  InvalidateRect(LocalRect);
end;

procedure TMultiHeaderGrid.SetViewTop(const Value: integer);
begin
  var BottomCellTop:=FRowData[High(FRowData)].Top;
  FViewTop:=Min(BottomCellTop,Max(0,Value));

  VScrollBar.Value:=Value;

  DoGridScroll;

  InvalidateRect(LocalRect);
end;

function TMultiHeaderGrid.CanObserve(const ID: Integer): Boolean;
begin
  Result:=False;
  if ID=TObserverMapping.EditLinkID then
    Result:=True
  else if ID=TObserverMapping.ControlValueID then
    Result:=True;
end;

procedure TMultiHeaderGrid.StartCellEditing(ACol, ARow: Integer; InitialChar: Char);
begin
  // Реализация редактирования ячейки
  // Можно создать TEdit поверх ячейки
  if Assigned(FOnStartEditing) then
    FOnStartEditing(Self, ACol, ARow, InitialChar);
end;

{ TCellStyle }


procedure TCellStyle.ClearMergedCell;
begin
  FParentCol:=-1;
  FParentRow:=-1;
  FIsMergedCell:=False;
end;

class operator TCellStyle.Initialize(out Dest: TCellStyle);
begin
  Dest.FFontStyle:=[];
  Dest.FSelectedCellColor:=TAlphaColors.Null;
  Dest.FFontSize:=0;
  Dest.FCellColor:=TAlphaColors.Null;
  Dest.FSelectedFontColor:=TAlphaColors.Null;
  Dest.FFontColor:=TAlphaColors.Null;
  Dest.FParentCol:=-1;
  Dest.FParentRow:=-1;

  Dest.FFontStyleIsSet:=False;
  Dest.FSelectedCellColorIsSet:=False;
  Dest.FFontSizeIsSet:=False;
  Dest.FCellColorIsSet:=False;
  Dest.FFontNameIsSet:=False;
  Dest.FSelectedFontColorIsSet:=False;
  Dest.FFontColorIsSet:=False;
  Dest.FIsMergedCell:=False;
end;

function TCellStyle.IsEmpty: boolean;
begin
  Result:=not (FFontStyleIsSet or
               FSelectedCellColorIsSet or
               FFontSizeIsSet or
               FCellColorIsSet or
               FFontNameIsSet or
               FSelectedFontColorIsSet or
               FFontColorIsSet);
end;

procedure TCellStyle.SetCellColor(const Value: TAlphaColor);
begin
  FCellColor:=Value;
  FCellColorIsSet:=True;
end;

procedure TCellStyle.SetFontColor(const Value: TAlphaColor);
begin
  FFontColor:=Value;
  FFontColorIsSet:=True;
end;

procedure TCellStyle.SetFontName(const Value: string);
begin
  FFontName:=Value;
  FFontNameIsSet:=True;
end;

procedure TCellStyle.SetFontSize(const Value: Single);
begin
  FFontSize:=Value;
  FFontSizeIsSet:=True;
end;

procedure TCellStyle.SetFontStyle(const Value: TFontStyles);
begin
  FFontStyle:=Value;
  FFontStyleIsSet:=True;
end;

procedure TCellStyle.SetMergedCell(ParenCol, ParenRow: integer);
begin
  FParentCol:=ParenCol;
  FParentRow:=ParenRow;
  FIsMergedCell:=True;
end;

procedure TCellStyle.SetParentCol(const Value: Integer);
begin
  FParentCol:=Value;
  if FParentRow<=0 then FParentRow:=1;
  FIsMergedCell:=True;
end;

procedure TCellStyle.SetParentRow(const Value: Integer);
begin
  FParentRow:=Value;
  if FParentCol<=0 then FParentCol:=1;
  FIsMergedCell:=True;
end;

procedure TCellStyle.SetSelectedCellColor(const Value: TAlphaColor);
begin
  FSelectedCellColor:=Value;
  FSelectedCellColorIsSet:=True;
end;

procedure TCellStyle.SetSelectedFontColor(const Value: TAlphaColor);
begin
  FSelectedFontColor:=Value;
  FSelectedFontColorIsSet:=True;
end;

{ TMultiHeaderStringGrid }

constructor TMultiHeaderStringGrid.Create(AOwner: TComponent);
begin
  inherited;
end;

destructor TMultiHeaderStringGrid.Destroy;
begin
  inherited;
end;

procedure TMultiHeaderStringGrid.DoGetCellStyle(ACol, ARow: Integer; var Style: TCellStyle);
begin
  if (ARow<0) or (ACol<0) then Exit;
  if (ARow>High(FCellStyles)) or (ACol>High(FCellStyles[ARow])) then Exit;

  Style:=FCellStyles[Arow,ACol];

  inherited;
end;

procedure TMultiHeaderStringGrid.DoGetCellText(ACol, ARow: Integer; var Text: string);
begin
  if (ARow<0) or (ACol<0) then Exit;
  if (ARow>High(FCellTexts)) or (ACol>High(FCellTexts[ARow])) then Exit;

  Text:=FCellTexts[Arow,ACol];

  inherited;
end;

procedure TMultiHeaderStringGrid.DoSetCellStyle(ACol, ARow: Integer; const Style: TCellStyle);
begin
  inherited;

  if (ARow<0) or (ACol<0) then Exit;

  if ARow>High(FCellStyles) then begin
    SetLength(FCellStyles,ARow+1);
  end;
  if ACol>High(FCellStyles[ARow]) then begin
    SetLength(FCellStyles[ARow],Max(ACol,ColCount));
  end;

  FCellStyles[ARow,ACol]:=Style;
end;

procedure TMultiHeaderStringGrid.DoSetCellText(ACol, ARow: Integer; const Text: string);
begin
  inherited;

  if (ARow<0) or (ACol<0) then Exit;

  if ARow>High(FCellTexts) then begin
    SetLength(FCellTexts,ARow+1);
  end;
  if ACol>High(FCellTexts[ARow]) then begin
    SetLength(FCellTexts[ARow],Max(ACol,ColCount));
  end;

  FCellTexts[ARow,ACol]:=Text;
end;

procedure TMultiHeaderGrid.SetResizeEnabled(const Value: Boolean);
begin
  if FResizeEnabled <> Value then begin
    FResizeEnabled := Value;
    // Сбрасываем состояние изменения размера при отключении
    if not Value then begin
      FResizeStartColumnIndex:=-1;
      FResizeEndColumnIndex:=-1;
      FResizeRowIndex := -1;
      ReleaseCapture; // Освобождаем захват мыши
      Cursor := crDefault;
    end;
  end;
end;

function TMultiHeaderGrid.GetResizeMargin: Integer;
begin
  Result := FResizeMargin;
end;

procedure TMultiHeaderGrid.SetResizeMargin(const Value: Integer);
begin
  if FResizeMargin <> Value then
  begin
    FResizeMargin := Value;
    if FResizeMargin < 1 then
      FResizeMargin := 1;
  end;
end;

function TMultiHeaderGrid.IsResizeArea(X, Y: Single; out AStartCol, AEndCol, ARow: Integer): TResizeMode;
var
  I,J        : Integer;
  Col         : integer;
  CellRect   : TRectF;
  ResizeRect : TRectF;
begin
  Result := TResizeMode.rmNone;
  AStartCol := -1;
  AEndCol := -1;
  ARow := -1;

  if not FResizeEnabled then Exit;

  if Y<=HeaderHeight then begin
    // Проверяем изменение ширины столбцов заголовка

    if FResizeColEnabled then begin
      for i:=0 to FHeaderLevels.Count-1 do begin
        Col:=0;
        for j:=0 to FHeaderLevels[i].Count-1 do begin
          CellRect:=GetHeaderRect(I, Col);

          ResizeRect := TRectF.Create(
            CellRect.Right - FResizeMargin,
            CellRect.Top,
            CellRect.Right + FResizeMargin,
            CellRect.Bottom
          );

          if ResizeRect.Contains(PointF(X, Y)) then begin
            AStartCol := Col;
            AEndCol := Col+FHeaderLevels[i][j].ColSpan-1;
            Result := TResizeMode.rmColumn;
            Exit;
          end;

          Col:=Col+FHeaderLevels[i][j].ColSpan;
        end;
      end;
    end;

    if FResizeHeaderRowEnabled then begin
      // Проверяем изменение высоты строк заголовка
      for i:=0 to FHeaderLevels.Count-1 do begin
        Col:=0;
        for j:=0 to FHeaderLevels[i].Count-1 do begin
          CellRect:=GetHeaderRect(I, Col);

          ResizeRect := TRectF.Create(
            CellRect.Left,
            CellRect.Bottom-FResizeMargin,
            CellRect.Right,
            CellRect.Bottom+FResizeMargin
          );

          if ResizeRect.Contains(PointF(X, Y)) then begin
            ARow := I;
            Result := TResizeMode.rmHeaderRow;
            Exit;
          end;

          Col:=Col+FHeaderLevels[i][j].ColSpan;
        end;
      end;
    end;
  end else begin
    if FResizeRowEnabled then begin
      // Проверяем изменение высоты строк грида
      for I := 0 to FRowCount - 1 do begin
        CellRect := GetCellRect(0, I); // Получаем прямоугольник строки
        if not CellRect.IsEmpty then begin
          // Проверяем нижнюю границу строки
          ResizeRect := TRectF.Create(
            CellRect.Left,
            CellRect.Bottom - FResizeMargin,
            Min(CellRect.Right,CellRect.Left+FResizeMargin),
            CellRect.Bottom + FResizeMargin
          );

          if ResizeRect.Contains(PointF(X, Y)) then
          begin
            ARow := I;
            Result := TResizeMode.rmGridRow;
            Exit;
          end;
        end;
      end;
    end;
  end;
end;

procedure TMultiHeaderGrid.StartColumnResize(StartCol, EndCol: Integer; X: Single);
var
  GlobalPos: TPointF;
begin
  FResizeStartColumnIndex:=StartCol;
  FResizeEndColumnIndex:=EndCol;
  FResizeRowIndex:=-1;
  FResizeMode:=TResizeMode.rmColumn;

  // Сохраняем глобальные координаты для точного отслеживания
  GlobalPos:=LocalToAbsolute(PointF(X, 0));
  FResizeStartPos:=PointF(GlobalPos.X, 0);

  SetLength(FResizeStartWidths,EndCol-StartCol+1);
  for var i:=StartCol to EndCol do begin
    FResizeStartWidths[i-StartCol]:=GetColWidth(i);
  end;
  Cursor:=crHSplit;
end;

procedure TMultiHeaderGrid.StartHeaderRowResize(ARow: Integer; Y: Single);
var
  GlobalPos: TPointF;
begin
  FResizeRowIndex:=ARow;
  FResizeStartColumnIndex:=-1;
  FResizeEndColumnIndex:=-1;
  FResizeMode:=TResizeMode.rmHeaderRow;

  // Сохраняем глобальные координаты для точного отслеживания
  GlobalPos:=LocalToAbsolute(PointF(0, Y));
  FResizeStartPos:=PointF(0, GlobalPos.Y);
  FResizeStartHeight:=FHeaderLevels[ARow].Height;
  Cursor:=crVSplit;
end;

procedure TMultiHeaderGrid.StartGridRowResize(ARow: Integer; Y: Single);
var
  GlobalPos: TPointF;
begin
  FResizeRowIndex:=ARow;
  FResizeStartColumnIndex:=-1;
  FResizeEndColumnIndex:=-1;
  FResizeMode:=TResizeMode.rmGridRow;

  // Сохраняем глобальные координаты для точного отслеживания
  GlobalPos:=LocalToAbsolute(PointF(0, Y));
  FResizeStartPos:=PointF(0, GlobalPos.Y);
  FResizeStartHeight:=GetRowHeight(ARow);
  Cursor:=crVSplit;
end;

procedure TMultiHeaderGrid.UpdateGroupColumnWidth(StartCol, EndCol: Integer; TotalWidth: Integer);
var
  I: Integer;
  GroupColCount: Integer;
  NewWidth, RemainingWidth: Integer;
  ProportionalWidths: array of Integer;
begin
  if TotalWidth < 10 * (EndCol - StartCol + 1) then
    TotalWidth := 10 * (EndCol - StartCol + 1); // Минимальная общая ширина

  GroupColCount := EndCol - StartCol + 1;
  SetLength(ProportionalWidths, GroupColCount);

  // Вычисляем пропорциональные ширины на основе исходных соотношений
  RemainingWidth := TotalWidth;

  // Первый проход: вычисляем пропорциональные ширины
  var FResizeStartWidth:=ResizeStartWidth;
  for I := 0 to GroupColCount - 1 do begin
    ProportionalWidths[I] := Round(FResizeStartWidths[I] * TotalWidth / FResizeStartWidth);
    Dec(RemainingWidth, ProportionalWidths[I]);
  end;

  // Второй проход: распределяем остаток
  if RemainingWidth <> 0 then begin
    for I := 0 to GroupColCount - 1 do begin
      if RemainingWidth > 0 then begin
        Inc(ProportionalWidths[I]);
        Dec(RemainingWidth);
      end else if RemainingWidth < 0 then begin
        Dec(ProportionalWidths[I]);
        Inc(RemainingWidth);
      end else begin
        Break;
      end;
    end;
  end;

  // Применяем новые ширины
  for I := StartCol to EndCol do begin
    NewWidth := ProportionalWidths[I - StartCol];
    if NewWidth < 10 then
      NewWidth := 10; // Минимальная ширина столбца

    SetColWidth(I, NewWidth);
  end;

  InvalidateRect(LocalRect);

  // Обновляем горизонтальную прокрутку
  if Assigned(HScrollBar) then
    HScrollBar.Value := FViewLeft;
end;

procedure TMultiHeaderGrid.UpdateColumnWidth(StartCol, EndCol: Integer; NewWidth: Integer);
begin
  if StartCol<EndCol then begin
    // Ресайз группы колонок
    UpdateGroupColumnWidth(StartCol, EndCol, NewWidth);
  end else begin
    // Ресайз одной колонки
    if NewWidth < 10 then begin
      NewWidth := 10; // Минимальная ширина
    end;

    SetColWidth(EndCol, NewWidth);

    InvalidateRect(LocalRect);

    // Обновляем горизонтальную прокрутку
    if Assigned(HScrollBar) then begin
      HScrollBar.Value := FViewLeft;
    end;
  end;
end;

procedure TMultiHeaderGrid.UpdateHeaderRowHeight(ARow: Integer; NewHeight: Integer);
begin
  if NewHeight < 10 then
    NewHeight := 10; // Минимальная высота

  FHeaderLevels[ARow].Height:=NewHeight;

  InvalidateRect(LocalRect);

  // Обновляем вертикальную прокрутку
  if Assigned(VScrollBar) then
    VScrollBar.Value := FViewTop;
end;

procedure TMultiHeaderGrid.UpdateRowHeight(ARow: Integer; NewHeight: Integer);
begin
  if NewHeight < 10 then
    NewHeight := 10; // Минимальная высота

  SetRowHeight(ARow, NewHeight);
  InvalidateRect(LocalRect);

  // Обновляем вертикальную прокрутку
  if Assigned(VScrollBar) then
    VScrollBar.Value := FViewTop;
end;

procedure TMultiHeaderGrid.DoColumnResized;
begin
  if Assigned(FOnColumnResized) then
    FOnColumnResized(Self,FResizeStartColumnIndex,FResizeEndColumnIndex);
end;

procedure TMultiHeaderGrid.DoHeaderRowResized;
begin
  if Assigned(FOnColumnResized) then
    FOnHeaderResized(Self,FResizeRowIndex);
end;

procedure TMultiHeaderGrid.DoRowResized;
begin
  if Assigned(FOnRowResized) then
    FOnRowResized(Self,FResizeRowIndex);
end;

procedure TMultiHeaderGrid.DoGridScroll;
begin
  if Assigned(FOnGridScroll) then
    FOnGridScroll(Self,ViewLeft,ViewTop);
end;

function TMultiHeaderGrid.ResizeStartWidth: Integer;
begin
  Result:=0;
  for var Width in FResizeStartWidths do begin
    Inc(Result,Width);
  end;
end;

procedure TMultiHeaderGrid.SetRowSelect(const Value: Boolean);
begin
  if FRowSelect<>Value then begin
    FRowSelect:=Value;
    InvalidateRect(LocalRect);
  end;
end;

end.


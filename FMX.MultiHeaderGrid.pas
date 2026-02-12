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
    FTextVAlignment         : TTextAlign;
    FTextHAlignment         : TTextAlign;
    FSelectedCellColor      : TAlphaColor;
    FSelectedFontColor      : TAlphaColor;
    FWordWrap               : Boolean;

    FFontStyleIsSet         : Boolean;
    FFontSizeIsSet          : Boolean;
    FCellColorIsSet         : Boolean;
    FFontNameIsSet          : Boolean;
    FFontColorIsSet         : Boolean;
    FTextVAlignmentIsSet    : Boolean;
    FTextHAlignmentIsSet    : Boolean;
    FSelectedCellColorIsSet : Boolean;
    FSelectedFontColorIsSet : Boolean;
    FIsMergedCell           : Boolean;
    FWordWrapIsSet          : Boolean;

    FParentCol              : Integer;
    FParentRow              : Integer;

    procedure SetCellColor(const Value: TAlphaColor);
    procedure SetFontColor(const Value: TAlphaColor);
    procedure SetFontName(const Value: string);
    procedure SetFontSize(const Value: Single);
    procedure SetFontStyle(const Value: TFontStyles);
    procedure SetSelectedCellColor(const Value: TAlphaColor);
    procedure SetSelectedFontColor(const Value: TAlphaColor);
    procedure SetParentCol(const Value: Integer);
    procedure SetParentRow(const Value: Integer);
    procedure SetCellColorIsSet(const Value: Boolean);
    procedure SetFontColorIsSet(const Value: Boolean);
    procedure SetFontNameIsSet(const Value: Boolean);
    procedure SetFontSizeIsSet(const Value: Boolean);
    procedure SetFontStyleIsSet(const Value: Boolean);
    procedure SetSelectedCellColorIsSet(const Value: Boolean);
    procedure SetSelectedFontColorIsSet(const Value: Boolean);
    procedure SetTextHAlignment(const Value: TTextAlign);
    procedure SetTextHAlignmentIsSet(const Value: Boolean);
    procedure SetTextVAlignment(const Value: TTextAlign);
    procedure SetTextVAlignmentIsSet(const Value: Boolean);
  private
    procedure SetWordWrap(const Value: Boolean);
    procedure SetWordWrapIsSet(const Value: Boolean);
  public
    class operator Initialize(out Dest: TCellStyle);

    function IsEmpty: boolean;

    property IsMergedCell: Boolean read FIsMergedCell;
    procedure SetMergedCell(ParenCol,ParenRow: integer);
    procedure ClearMergedCell;

    property FontName: string read FFontName write SetFontName;
    property FontSize: Single read FFontSize write SetFontSize;
    property FontStyle: TFontStyles read FFontStyle write SetFontStyle;
    property FontColor: TAlphaColor read FFontColor write SetFontColor;

    property CellColor: TAlphaColor read FCellColor write SetCellColor;

    property SelectedFontColor: TAlphaColor read FSelectedFontColor write SetSelectedFontColor;
    property SelectedCellColor: TAlphaColor read FSelectedCellColor write SetSelectedCellColor;

    property TextVAlignment: TTextAlign read FTextVAlignment write SetTextVAlignment;
    property TextHAlignment: TTextAlign read FTextHAlignment write SetTextHAlignment;

    property FontNameIsSet: Boolean read FFontNameIsSet write SetFontNameIsSet;
    property FontSizeIsSet: Boolean read FFontSizeIsSet write SetFontSizeIsSet;
    property FontStyleIsSet: Boolean read FFontStyleIsSet write SetFontStyleIsSet;
    property FontColorIsSet: Boolean read FFontColorIsSet write SetFontColorIsSet;
    property CellColorIsSet: Boolean read FCellColorIsSet write SetCellColorIsSet;

    property TextVAlignmentIsSet: Boolean read FTextVAlignmentIsSet write SetTextVAlignmentIsSet;
    property TextHAlignmentIsSet: Boolean read FTextHAlignmentIsSet write SetTextHAlignmentIsSet;

    property SelectedCellColorIsSet: Boolean read FSelectedCellColorIsSet write SetSelectedCellColorIsSet;
    property SelectedFontColorIsSet: Boolean read FSelectedFontColorIsSet write SetSelectedFontColorIsSet;

    property WordWrap: Boolean read FWordWrap write SetWordWrap;
    property WordWrapIsSet: Boolean read FWordWrapIsSet write SetWordWrapIsSet;

    property ParentCol: Integer read FParentCol write SetParentCol;
    property ParentRow: Integer read FParentRow write SetParentRow;
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

    function AddColumn(Caption: string = '';
                       ColSpan: Integer = 1;
                       RowSpan: Integer = 1): THeaderElement;

    function FillRow(Caption: string = ''): THeaderElement;

    property Height: Integer read FHeight write SetHeight;
  end;

  TMultiHeaderGrid = class;

  THeaderLevels = class(TObjectList<THeaderLevel>)
    Grid : TMultiHeaderGrid;

    constructor Create(Grid : TMultiHeaderGrid);

    function AddRow: THeaderLevel; overload;
    function AddRow(Heigth: integer): THeaderLevel; overload;
    function GetElementAtCell(ACol, ARow: integer): THeaderElement;
  end;

  FColData = record
    Widths         : integer;
    MinWidth       : integer;
    MaxWidth       : integer;
    TextVAlignment : TTextAlign;
    TextHAlignment : TTextAlign;
    WordWrap       : Boolean;
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

  TResizeMode = (rmNone, rmColumn, rmHeaderRow, rmGridRow);

  TScrollShowMode = (smAuto, smShow, smHide);

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

    FColData: array of FColData;
    FRowData: Array of TRowData;

    FHeaderLevels: THeaderLevels;
    FGridLines: Boolean;
    FHeaderLineColor: TAlphaColor;
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
    FLastClickIsOnCell: boolean;

    FRowSelect: Boolean;
    FWordWrap: Boolean;
    FGridCellsHasWordWrap: boolean;
    FVerticalScroll: TScrollShowMode;
    FHorisontalScroll: TScrollShowMode;

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
    function GetCellRect(ACol, ARow: Integer): TRectF; overload;
    function GetCellRect(ACol, ARow: Integer; out MergedCell: TMergedCell): TRectF; overload;
    function GetMergedCellRect(ACol, ARow: Integer): TRectF;
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
    function GetColTextHAlignment(Index: Integer): TTextAlign;
    function GetColTextVAlignment(Index: Integer): TTextAlign;
    procedure SetColTextHAlignment(Index: Integer; const Value: TTextAlign);
    procedure SetColTextVAlignment(Index: Integer; const Value: TTextAlign);
    function GetColWordWrap(Index: Integer): Boolean;
    procedure SetColWordWrap(Index: Integer; const Value: Boolean);
    procedure SetWordWrap(const Value: Boolean);
    procedure SetHeaderLineColor(const Value: TAlphaColor);
    procedure SetVerticalScroll(const Value: TScrollShowMode);
    procedure SetHorisontalScroll(const Value: TScrollShowMode);
    function GetColMaxWidth(Index: Integer): integer;
    function GetColMinWidth(Index: Integer): integer;
    procedure SetColMaxWidth(Index: Integer; const Value: integer);
    procedure SetColMinWidth(Index: Integer; const Value: integer);
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

    procedure DblClick; override;

    procedure DoColumnResized; virtual;
    procedure DoHeaderRowResized; virtual;
    procedure DoRowResized; virtual;
    procedure DoGridScroll; virtual;

    procedure MouseDown(Button: TMouseButton; Shift: TShiftState; X, Y: Single); override;
    procedure MouseMove(Shift: TShiftState; X, Y: Single); override;
    procedure MouseUp(Button: TMouseButton; Shift: TShiftState; X, Y: Single); override;
    procedure MouseWheel(Shift: TShiftState; WheelDelta: Integer; var Handled: Boolean); override;

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

    procedure AutoSizeCols(ForcePrecise: boolean = False);
    procedure AutoSizeRows(ForcePrecise: boolean = False); overload;
    procedure AutoSizeRows(FromRow, ToRow: integer; ForcePrecise: boolean = False); overload;
    procedure AutoSizeVisibleRows;
    procedure AutoSize(ForcePrecise: boolean = False);

    procedure ClearSelection;

    property ColLefts[Index: Integer]: Integer read GetColLeft;
    property ColWidths[Index: Integer]: Integer read GetColWidth write SetColWidth;
    property ColTextHAlignment[Index: Integer]: TTextAlign read GetColTextHAlignment write SetColTextHAlignment;
    property ColTextVAlignment[Index: Integer]: TTextAlign read GetColTextVAlignment write SetColTextVAlignment;
    property ColMinWidth[Index: Integer]: integer read GetColMinWidth write SetColMinWidth;
    property ColMaxWidth[Index: Integer]: integer read GetColMaxWidth write SetColMaxWidth;
    property ColWordWrap[Index: Integer]: Boolean read GetColWordWrap write SetColWordWrap;

    property RowHeights[Index: Integer]: Integer read GetRowHeight write SetRowHeight;
    property RowTops[Index: Integer]: Integer read GetRowTops;

    property Cells[ACol, ARow: Integer]: string read GetCells write SetCells;
    property CellStyle[ACol, ARow: Integer]: TCellStyle read GetCellStyle write SetCellStyle;

    property SelectedCell: TPoint read FSelectedCell write SetSelectedCell;
    property Header: THeaderLevels read FHeaderLevels;

    property ViewTop: Integer read FViewTop write SetViewTop;
    property ViewLeft: Integer read FViewLeft write SetViewLeft;

    function ViewBottom: Integer;
    function ViewCellsHeight: Integer;
    function HeaderHeight: Integer;
    function ViewPortWidth: Integer;
    function ViewPortHeight: Integer;
    function ViewPortDataHeight: Integer;
    function FullTableWidth: Integer;
    function FullTableHeight: Integer;

    procedure Invalidate;

    function RowAtHeightCoord(Y: Integer): integer;
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
    property OnDblClick;
    property OnResize;

    property ColCount: Integer read FColCount write SetColCount default 5;
    property RowCount: Integer read FRowCount write SetRowCount default 10;

    property Col: Integer read FSelectedCell.X write SetCol;
    property Row: Integer read FSelectedCell.Y write SetRow;

    property DefaultColWidth: integer read FDefaultColWidth write SetDefaultColWidth;
    property DefaultRowHeight: integer read FDefaultRowHeight write SetDefaultRowHeight;
    property BackgroundColor: TAlphaColor read FBackgroundColor write SetBackgroundColor default TAlphaColorRec.White;
    property GridLines: Boolean read FGridLines write SetGridLines default True;
    property GridLineColor: TAlphaColor read FGridLineColor write SetGridLineColor default TAlphaColorRec.Gray;
    property GridLineWidth: Single read FGridLineWidth write SetGridLineWidth;

    property HeaderLineColor: TAlphaColor read FHeaderLineColor write SetHeaderLineColor default TAlphaColorRec.Black;

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
    property WordWrap: Boolean read FWordWrap write SetWordWrap default False;

    property HorisontalScroll: TScrollShowMode read FHorisontalScroll write SetHorisontalScroll default TScrollShowMode.smAuto;
    property VerticalScroll: TScrollShowMode read FVerticalScroll write SetVerticalScroll default TScrollShowMode.smAuto;

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

    property ResizeMargin: Integer read GetResizeMargin write SetResizeMargin default 2;

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

function THeaderLevel.AddColumn(Caption: string = '';
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

function THeaderLevel.FillRow(Caption: string): THeaderElement;
begin
  Result:=AddColumn(Caption,-1);
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
  for i:=0 to High(Result) do begin
    Result[i]:=False;
  end;

  var LevelIndex:=FLevels.IndexOf(Self);

  // Проходим по всем уровням выше текущего
  for i:=0 to LevelIndex-1 do begin
    for j:=0 to FLevels[i].Count-1 do begin
      Element:=FLevels[i][j];
      // Если элемент занимает несколько строк и его RowSpan распространяется на текущий уровень
      if (Element.RowSpan>1) and (i+Element.RowSpan>LevelIndex) then begin
        // Помечаем все колонки, которые занимает этот элемент
        for k:=0 to Element.ColSpan-1 do begin
          Col:=Element.ColSkip+k;
          if Col<Length(Result) then begin
            Result[Col]:=True;
          end;
        end;
      end;
    end;
  end;
end;

{ THeaderLevels }

function THeaderLevels.AddRow(Heigth: integer): THeaderLevel;
begin
  Result:=THeaderLevel.Create(Self);
  Result.FHeight:=Heigth;
end;

function THeaderLevels.AddRow: THeaderLevel;
begin
  Result:=AddRow(25);
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
      var ColSpan:=Element.ColSpan;
      if ColSpan<0 then ColSpan:=Grid.FColCount-Left;

      if (ACol>=Left+Element.ColSkip) and (ACol<Left+ColSpan+Element.ColSkip) and
         (ARow>=Row) and (ARow<Row+Element.RowSpan) then Exit(Element);
      inc(Left,ColSpan);
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
 i: integer;
begin
  inherited;

  FColCount:=5;
  FRowCount:=10;
  FDefaultColWidth:=80;
  FDefaultRowHeight:=20;
  FGridLines:=True;
  FGridLineColor:=TAlphaColors.Gray;
  FHeaderLineColor:=TAlphaColors.Black;
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
  SetLength(FColData,FColCount);
  for i:=0 to FColCount-1 do begin
    FColData[i].Widths:=FDefaultColWidth;
    FColData[i].TextVAlignment:=TTextAlign.Center;
    FColData[i].TextHAlignment:=TTextAlign.Leading;
    FColData[i].WordWrap:=False;
    FColData[i].MinWidth:=0;
    FColData[i].MaxWidth:=MaxInt;
  end;

  var Top:=0;
  SetLength(FRowData,FRowCount);
  for i:=0 to FRowCount-1 do begin
    FRowData[i].Top:=Top;
    FRowData[i].Height:=trunc(FDefaultRowHeight);
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
  FResizeMargin:=2;

  if csDesigning in ComponentState then begin
    Header.AddRow.FillRow('Header');
    Width:=401;
    Height:=225;
    VScrollBar.Visible:=False;
    HScrollPanel.Visible:=False;
  end;
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
  i: Integer;
begin
  if Value<0 then Value:=0;
  if FColCount<>Value then begin
    var OldValue:=FColCount;

    FColCount:=Value;

    SetLength(FColData,FColCount);
    for i:=OldValue to FColCount-1 do begin
      FColData[i].Widths:=FDefaultColWidth;
      FColData[i].TextVAlignment:=TTextAlign.Center;
      FColData[i].TextHAlignment:=TTextAlign.Leading;
      FColData[i].MinWidth:=0;
      FColData[i].MaxWidth:=MaxInt;
      FColData[i].WordWrap:=False;
    end;

    if FColCount=0 then FGridCellsHasWordWrap:=False;

    UpdateSize;
    Invalidate;

    HScrollBar.Max:=Value;
  end;
end;

procedure TMultiHeaderGrid.SetColMaxWidth(Index: Integer; const Value: integer);
begin
  if (Index>=0) and (Index<FColCount) then begin
    FColData[Index].MaxWidth:=Max(Value,FColData[Index].MinWidth);
    FColData[Index].Widths:=Max(FColData[Index].Widths,FColData[Index].MaxWidth);
    Invalidate;
  end;
end;

procedure TMultiHeaderGrid.SetColMinWidth(Index: Integer; const Value: integer);
begin
  if (Index>=0) and (Index<FColCount) then begin
    FColData[Index].MinWidth:=Min(Value,FColData[Index].MaxWidth);
    FColData[Index].Widths:=Min(FColData[Index].Widths,FColData[Index].MinWidth);
    Invalidate;
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
      for var i:=0 to FColCount-1 do begin
        UnMergeCells(i,Value-1);
      end;
    end;

    FRowCount:=Value;
    var Top:=0;
    if OldValue>0 then Top:=FRowData[OldValue-1].Top;
    SetLength(FRowData,FRowCount);
    for var i:=Max(OldValue-1,0) to FRowCount-1 do begin
      FRowData[i].Top:=Top;
      if i>=OldValue then begin
        FRowData[i].Height:=trunc(FDefaultRowHeight);
      end;
      Top:=Top+FRowData[i].Height;
    end;

    if FColCount=0 then FGridCellsHasWordWrap:=False;

    UpdateSize;
    Invalidate;

    VScrollBar.Max:=FullTableHeight-HeaderHeight;
    VScrollBar.ViewportSize:=ViewPortDataHeight;
  end;
end;

procedure TMultiHeaderGrid.SetDefaultColWidth(const Value: integer);
var
  i: Integer;
begin
  if FDefaultColWidth<>Value then begin
    var OldValue:=FDefaultColWidth;
    FDefaultColWidth:=Value;

    for i:=0 to FColCount-1 do begin
      if FColData[i].Widths=OldValue then begin
        ColWidths[i]:=FDefaultColWidth;
      end;
    end;
    UpdateSize;
    Invalidate;
  end;
end;

procedure TMultiHeaderGrid.SetDefaultRowHeight(const Value: integer);
begin
  if FDefaultRowHeight<>Value then begin
    FDefaultRowHeight:=Value;

    var Top:=0;
    SetLength(FRowData,FRowCount);
    for var i:=0 to FRowCount-1 do begin
      FRowData[i].Top:=Top;
      FRowData[i].Height:=trunc(FDefaultRowHeight);
      Top:=Top+trunc(FDefaultRowHeight);
    end;

    UpdateSize;
    Invalidate;
  end;
end;

function TMultiHeaderGrid.GetColLeft(Index: Integer): Integer;
begin
  Result:=0;
  for var i:=0 to Index-1 do
    inc(Result,FColData[i].Widths);
end;

function TMultiHeaderGrid.GetColMaxWidth(Index: Integer): integer;
begin
  if (Index>=0) and (Index<Length(FColData)) then
    Result:=FColData[Index].MaxWidth
  else
    Result:=MaxInt;
end;

function TMultiHeaderGrid.GetColMinWidth(Index: Integer): integer;
begin
  if (Index>=0) and (Index<Length(FColData)) then
    Result:=FColData[Index].MinWidth
  else
    Result:=0;
end;

function TMultiHeaderGrid.GetColWidth(Index: Integer): Integer;
begin
  if (Index>=0) and (Index<Length(FColData)) then
    Result:=FColData[Index].Widths
  else
    Result:=FDefaultColWidth;
end;

function TMultiHeaderGrid.FullTableHeight: Integer;
begin
  if FRowCount=0 then Exit(0);

  var BottomCell:=FRowData[FRowCount-1];

  Result:=BottomCell.Top+BottomCell.Height+HeaderHeight-Trunc(FGridLineWidth/2);
end;

function TMultiHeaderGrid.FullTableWidth: Integer;
begin
  Result:=round(FGridLineWidth/2);
  for var i:=0 to FColCount-1 do begin
    Result:=Result+FColData[i].Widths;
  end;
end;

procedure TMultiHeaderGrid.SetColWidth(Index: Integer; const Value: Integer);
begin
  if (Index>=0) and (Index<Length(FColData)) and (FColData[Index].Widths<>Value) then begin
    FColData[Index].Widths:=Max(Min(Value,FColData[Index].MaxWidth),FColData[Index].MinWidth);
    UpdateSize;
    Invalidate;
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
  i : integer;
begin
  if (Index>=0) and (Index<Length(FRowData)) and (FRowData[Index].Height<>Value) then begin
    FRowData[Index].Height:=Trunc(Value);
    for i:=Index+1 to FRowCount-1 do begin
      FRowData[i].Top:=FRowData[i-1].Top+FRowData[i-1].Height;
    end;
    UpdateSize;
    Invalidate;
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

function TMultiHeaderGrid.GetCellRect(ACol, ARow: Integer; out MergedCell: TMergedCell): TRectF;
var
  X, Y : Single;
  i    : Integer;
begin
  // Проверяем валидность диапазона
  if (ACol<0) or (ARow<0) or (ACol>=FColCount) or (ARow>=FRowCount) then
    Exit(Default(TRectF));

  X:=-ViewLeft;
  Y:=HeaderHeight-ViewTop;

  if IsMergedCell(ACol,ARow,MergedCell) then begin
    // Корректируем прямоугольник для объединенной ячейки

    Result.Left:=X+GetColLeft(MergedCell.Col);
    Result.Top:=Y+FRowData[MergedCell.Row].Top;
    Result.Right:=Result.Left;
    Result.Bottom:=Result.Top;
    for i:=0 to MergedCell.ColSpan-1 do
      Result.Right:=Result.Right+GetColWidth(MergedCell.Col+i);
    for i:=0 to MergedCell.RowSpan-1 do
      Result.Bottom:=Result.Bottom+GetRowHeight(MergedCell.Row+i);
  end else begin
    // Прямой рассчет

    Result.Left:=X+GetColLeft(ACol);
    Result.Top:=Y+FRowData[ARow].Top;
    Result.Right:=Result.Left+FColData[ACol].Widths;
    Result.Bottom:=Result.Top+FRowData[ARow].Height;
  end;
end;

function TMultiHeaderGrid.GetCellRect(ACol, ARow: Integer): TRectF;
var
  X, Y: Single;
  i: Integer;
begin
  // Проверяем валидность диапазона
  if (ACol<0) or (ARow<0) or (ACol>=FColCount) or (ARow>=FRowCount) then
    Exit(Default(TRectF));

  X:=FGridLineWidth/4-ViewLeft;
  Y:=-FGridLineWidth/4+HeaderHeight-ViewTop;

  // Вычисляем позицию колонки
  for i:=0 to ACol-1 do
    X:=X+GetColWidth(i);

  // Вычисляем позицию строки
  Y:=Y+FRowData[ARow].Top;

  Result:=RectF(X, Y, X+GetColWidth(ACol), Y+GetRowHeight(ARow));
end;


function TMultiHeaderGrid.GetMergedCellRect(ACol, ARow: Integer): TRectF;
begin
  var MergedCell: TMergedCell;
  Result:=GetCellRect(ACol,ARow,MergedCell);
end;

function TMultiHeaderGrid.HeaderHeight: Integer;
begin
  Result:=round(FGridLineWidth/2);

  // Добавляем высоту заголовков
  for var i:=0 to FHeaderLevels.Count-1 do
    Result:=Result+FHeaderLevels[i].Height;
end;

function TMultiHeaderGrid.GetHeaderRect(ALevel, ACol: Integer): TRectF;
var
  X, Y          : Single;
  i             : Integer;
  Element       : THeaderElement;
  ActualColSpan : Integer;
  ActualRowSpan : Integer;
begin
  X:=FGridLineWidth/4-ViewLeft;
  Y:=FGridLineWidth/4;

  // Вычисляем позицию уровня заголовка
  for i:=0 to ALevel-1 do begin
    Y:=Y+FHeaderLevels[i].Height;
  end;

  // Вычисляем позицию колонки
  for i:=0 to ACol-1 do
    X:=X+GetColWidth(i);

  Element:=FHeaderLevels.GetElementAtCell(ACol,ALevel);

  if not Assigned(Element) then Exit(RectF(-1,-1,-1,-1));
  if FHeaderLevels.IndexOf(Element.FLevel)<>ALevel then Exit;

  // Определяем фактический ColSpan (не больше оставшихся колонок)
  ActualColSpan:=Element.ColSpan;
  if ActualColSpan<0 then begin
    ActualColSpan:=FColCount-ACol;
  end;
  if ACol+ActualColSpan>FColCount then begin
    ActualColSpan:=FColCount-ACol;
  end;

  // Вычисляем ширину ячейки заголовка
  Result:=RectF(X, Y, X, Y+ FHeaderLevels[ALevel].Height);
  for i:=0 to ActualColSpan-1 do
    Result.Right:=Result.Right+GetColWidth(ACol+i);

  // Определяем фактический RowSpan (не больше оставшихся строк)
  ActualRowSpan:=Element.RowSpan;
  if ALevel+ActualRowSpan>FHeaderLevels.Count then
    ActualRowSpan:=FHeaderLevels.Count-ALevel;

  // Вычисляем высоту ячейки заголовка
  for i:=1 to ActualRowSpan-1 do begin
    Result.Bottom:=Result.Bottom+FHeaderLevels[ALevel+i].Height;
  end;
end;

procedure TMultiHeaderGrid.UpdateSize;
begin
  if (FColCount+FRowCount>0) and (ViewPortHeight<FullTableHeight-FGridLineWidth/2-1) then begin
    VScrollBar.Visible:=FVerticalScroll<>TScrollShowMode.smHide;
    CornerPanel.Visible:=True;
    VScrollBar.Max:=FullTableHeight-HeaderHeight-FGridLineWidth/2;
    VScrollBar.Value:=ViewTop;
    VScrollBar.ViewportSize:=ViewPortDataHeight;
    VScrollBar.SmallChange:=DefaultRowHeight;
    VScrollBar.OnChange:=VScrollBarChange;
    VScrollBar.Enabled:=True;
  end else begin
    VScrollBar.Visible:=FVerticalScroll=TScrollShowMode.smShow;
    CornerPanel.Visible:=FVerticalScroll=TScrollShowMode.smShow;
    VScrollBar.OnChange:=nil;
    VScrollBar.Max:=1;
    VScrollBar.ViewportSize:=1;
    VScrollBar.Enabled:=False;
  end;

  if (FColCount>0) and (FullTableWidth>=ViewPortWidth) then begin
    HScrollPanel.Visible:=FHorisontalScroll<>TScrollShowMode.smHide;
    HScrollBar.Max:=FullTableWidth+1;
    HScrollBar.Value:=ViewLeft;
    HScrollBar.ViewportSize:=ViewPortWidth;
    HScrollBar.SmallChange:=FDefaultColWidth;
    HScrollBar.OnChange:=HScrollBarChange;
    HScrollBar.Enabled:=True;
  end else begin
    HScrollPanel.Visible:=FHorisontalScroll=TScrollShowMode.smShow;
    HScrollBar.OnChange:=nil;
    HScrollBar.Max:=1;
    HScrollBar.ViewportSize:=1;
    HScrollBar.Enabled:=False;
  end;
end;

function TMultiHeaderGrid.ViewCellsHeight: Integer;
begin
  Result:=Round(LocalRect.Height-HeaderHeight);
  if HScrollPanel.Visible then Result:=Round(Result-HScrollPanel.Height);
end;

function TMultiHeaderGrid.ViewPortDataHeight: Integer;
begin
  Result:=ViewPortHeight-HeaderHeight;
end;

function TMultiHeaderGrid.ViewPortWidth: Integer;
begin
  Result:=Round(LocalRect.Width);
  if VScrollBar.Visible then Result:=Round(Result-VScrollBar.Width);
end;

function TMultiHeaderGrid.ViewPortHeight: Integer;
begin
  Result:=Round(LocalRect.Height-GridLineWidth);
  if HScrollPanel.Visible then Result:=Round(Result-HScrollPanel.Height);
end;

procedure TMultiHeaderGrid.VScrollBarChange(Sender: TObject);
begin
  ViewTop:=Round(VScrollBar.Value);
  Invalidate;
end;

procedure TMultiHeaderGrid.HScrollBarChange(Sender: TObject);
begin
  ViewLeft:=Round(HScrollBar.Value);
  Invalidate;
end;

procedure TMultiHeaderGrid.Paint;
var
  Canvas: TCanvas;
begin
  Canvas:=Self.Canvas;

  if FWordWrap then begin
    AutoSizeVisibleRows;
  end;

  if Canvas.BeginScene then begin
    var Save:=Canvas.SaveState;
    try
      // Устанавливаем область обрезки
      var R:=TRectF.Create(LocalRect.Left,LocalRect.Top,
                           LocalRect.Left+LocalRect.Width,LocalRect.Top+LocalRect.Height);
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
  i, j: Integer;
begin
  Canvas.Stroke.Kind:=TBrushKind.Solid;
  Canvas.Stroke.Color:=FHeaderLineColor;
  Canvas.Stroke.Thickness:=FGridLineWidth/2;

  for i:=0 to FHeaderLevels.Count-1 do begin
    j:=0;
    for var Element in FHeaderLevels[i] do begin
      DrawHeaderCell(Canvas, i, j+Element.ColSkip);
      j:=j+Element.ColSpan;
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

  var HAlignment:=TTextAlign.Center;
  if Element.Style.TextHAlignmentIsSet then begin
    HAlignment:=Element.Style.TextHAlignment;
  end;
  var VAlignment:=TTextAlign.Center;
  if Element.Style.TextVAlignmentIsSet then begin
    HAlignment:=Element.Style.TextVAlignment;
  end;

  if Element.Style.FontColorIsSet then begin
    Canvas.Fill.Color:=Element.Style.FontColor;
  end else begin
    Canvas.Fill.Color:=FCellFontColor;
  end;
  Canvas.FillText(ARect, Element.Caption, False, 1, [], HAlignment, VAlignment);

  // Границы
  Canvas.DrawRect(ARect, 0, 0, AllCorners, 1);
end;

procedure TMultiHeaderGrid.DrawCells(Canvas: TCanvas);
var
  i, j       : Integer;
  Rect       : TRectF;
  MergedCell : TMergedCell;
begin
  var TopRow:=RowAtHeightCoord(ViewTop);
  var ViewBottomCell:=ViewBottom;

  if TopRow<0 then Exit;

  for j:=TopRow to FRowCount-1 do begin
    if FRowData[j].Top>ViewBottomCell then Break;
    for i:=0 to FColCount-1 do begin

      // Пропускаем объединенные ячейки (кроме первой)
      if IsMergedCell(i, j, MergedCell) then begin
        var FirstVisibleRow:=Max(MergedCell.Row,TopRow);

        if (i=MergedCell.Col) and (j=FirstVisibleRow) then begin
          Rect:=GetCellRect(MergedCell.Col, MergedCell.Row);
          // Корректируем прямоугольник для объединенной ячейки
          Rect.Right:=Rect.Left;
          Rect.Bottom:=Rect.Top;
          for var K:=0 to MergedCell.ColSpan-1 do
            Rect.Right:=Rect.Right+GetColWidth(i+K);
          for var K:=0 to MergedCell.RowSpan-1 do
            Rect.Bottom:=Rect.Bottom+GetRowHeight(j+K);

          DrawCell(Canvas, MergedCell.Col, MergedCell.Row, Rect);
        end;
        Continue;
      end;

      Rect:=GetCellRect(i, j);
      DrawCell(Canvas, i, j, Rect);
    end;
  end;
end;

procedure TMultiHeaderGrid.DrawCell(Canvas: TCanvas; ACol, ARow: Integer; ARect: TRectF);
var
  Handled         : Boolean;
  Text            : string;
  MergedCell      : TMergedCell;
  IsSelected      : Boolean;
  WordWrapEnabled : Boolean; // Добавлено: флаг переноса слов
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
    for var i:=0 to MergedCell.ColSpan-1 do
      ARect.Right:=ARect.Right+GetColWidth(ACol+i);
    for var i:=0 to MergedCell.RowSpan-1 do
      ARect.Bottom:=ARect.Bottom+GetRowHeight(ARow+i);

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

    // Определяем, нужно ли использовать перенос слов
    // Добавлено: логика определения WordWrap
    if CellStyle.WordWrapIsSet then begin
      WordWrapEnabled:=CellStyle.WordWrap;
    end else if (ACol>=0) and (ACol<FColCount) then begin
      WordWrapEnabled:=FWordWrap or FColData[ACol].WordWrap;
    end else begin
      WordWrapEnabled:=FWordWrap;
    end;

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

    // Текст
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

    var HAlignment:=ColTextHAlignment[ACol];
    if CellStyle.TextHAlignmentIsSet then begin
      HAlignment:=CellStyle.TextHAlignment;
    end;
    var VAlignment:=ColTextVAlignment[ACol];
    if CellStyle.TextVAlignmentIsSet then begin
      VAlignment:=CellStyle.TextVAlignment; // Исправлено: было HAlignment
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

    // Корректировка области вывода для текста
    ARect.Left:=ARect.Left+CellPadding.Left+FGridLineWidth/4;
    ARect.Top:=ARect.Top+CellPadding.Top+FGridLineWidth/4-1;
    ARect.Right:=ARect.Right-CellPadding.Right-FGridLineWidth/4;
    if WordWrapEnabled then begin
      case HAlignment of
        TTextAlign.Center: begin
          ARect.Left:=ARect.Left-1;
          ARect.Right:=ARect.Right+1;
        end;
        TTextAlign.Leading: ARect.Right:=ARect.Right+2;
        TTextAlign.Trailing: ARect.Left:=ARect.Left-2;
      end;
    end;

    ARect.Bottom:=ARect.Bottom-CellPadding.Bottom-FGridLineWidth/4;

    // Изменено: добавлен параметр WordWrapEnabled
    Canvas.FillText(ARect, Text, WordWrapEnabled, 1, [], HAlignment, VAlignment);
  end;
end;

procedure TMultiHeaderGrid.DrawGridLines(Canvas: TCanvas);
var
  i, j: Integer;
  X, Y: Single;
  StartX: Single;
  StartY: Single;
  MergedCell: TMergedCell;
begin
  if not FGridLines then Exit;
  if (FRowCount=0) or (FColCount=0) then Exit;

  Canvas.Stroke.Kind:=TBrushKind.Solid;
  Canvas.Stroke.Color:=FGridLineColor;
  Canvas.Stroke.Thickness:=FGridLineWidth/2;

  // Вычисляем начальную Y-координату для данных (после заголовков)
  StartX:=FGridLineWidth/4-ViewLeft;
  StartY:=-FGridLineWidth/4+HeaderHeight-ViewTop;

  var ViewBottomCell:=ViewBottom;

  var TopRow:=RowAtHeightCoord(ViewTop);
  if TopRow<0 then Exit;

  // Рисуем линии для данных (как раньше)
  // Сначала рисуем все вертикальные линии
  X:=StartX;
  for i:=0 to FColCount do begin
    // Внутренние вертикальные линии
    for j:=TopRow to FRowCount-1 do begin
      Y:=StartY+FRowData[j].Top;
      if FRowData[j].Top>ViewBottomCell then Break;

      // Проверяем, не находится ли линия внутри объединенной ячейки
      var ShouldDraw:=True;

      // Проверяем ячейку слева
      if IsMergedCell(i-1, j, MergedCell) then begin
        if i<MergedCell.Col+MergedCell.ColSpan then
          ShouldDraw:=False;
      end;

      // Проверяем ячейку справа
      if ShouldDraw and IsMergedCell(i, j, MergedCell) then begin
        if i>MergedCell.Col then
          ShouldDraw:=False;
      end;

      if ShouldDraw then begin
        Canvas.DrawLine(PointF(X, Y), PointF(X, Y+GetRowHeight(j)), 1);
      end;
    end;

    if i<FColCount then
      X:=X+GetColWidth(i);
  end;

  // Затем рисуем все горизонтальные линии
  for i:=TopRow to FRowCount do begin
    if i=0 then begin
      Y:=StartY+FRowData[i].Top;
    end else begin
      Y:=StartY+FRowData[i-1].Top+FRowData[i-1].Height;
    end;

    if (i>0) and (FRowData[i-1].Top>ViewBottomCell) then Break;

    X:=StartX;
    for j:=0 to FColCount-1 do begin
      // Проверяем, не находится ли линия внутри объединенной ячейки
      var ShouldDraw:=True;
      // Внутренние горизонтальные линии

      // Проверяем ячейку сверху
      if IsMergedCell(j, i-1, MergedCell) then begin
        if i<MergedCell.Row+MergedCell.RowSpan then begin
          ShouldDraw:=False;
        end;
      end;

      // Проверяем ячейку снизу
      if ShouldDraw and IsMergedCell(j, i, MergedCell) then begin
        if i>MergedCell.Row then begin
          ShouldDraw:=False;
        end;
      end;

      if ShouldDraw then begin
        Canvas.DrawLine(PointF(X, Y), PointF(X+GetColWidth(j), Y), 1);
      end;

      X:=X+GetColWidth(j);
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
    var IsMergedCell: Boolean;
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
  Invalidate;
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

    if EndPos=0 then begin
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

    if EndPos=0 then begin
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

procedure TMultiHeaderGrid.AutoSizeCols(ForcePrecise:boolean=False);
var
  i,j:Integer;
  MaxWidth:Single;
  Text:string;
begin
  var CellPaddingWidth:=CellPadding.Left+CellPadding.Right;
  var CellDelimterWidth:=FGridLineWidth/2;
  var CellPaddingFull:=CellPaddingWidth+CellDelimterWidth+1;

  // Оптимизация: если нет WordWrap, используем быстрый режим
  var UseFastMode:=not ForcePrecise and (FRowCount*FColCount>10000);

  for i:=0 to FColCount-1 do begin
    MaxWidth:=0;

    // Проверяем заголовки
    for j:=0 to FHeaderLevels.Count-1 do begin
      var Element:=FHeaderLevels.GetElementAtCell(i,j);
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

    // Оптимизация: если в этом столбце нет WordWrap, используем оригинальную быструю логику
    var ColHasWordWrap:=FWordWrap or FColData[i].WordWrap or FGridCellsHasWordWrap;

    if not ColHasWordWrap and UseFastMode then begin
      // Быстрый режим
      Canvas.Font.Assign(FCellFont);
      var FLetterWidth:=Canvas.TextWidth('V');
      var MaxLetters:=0.0;

      for j:=0 to FRowCount-1 do begin
        var ColSpan:=1;
        var MergedCell:TMergedCell;
        if IsMergedCell(i,j,MergedCell) then begin
          Text:=Cells[MergedCell.Col,MergedCell.Row];
          ColSpan:=MergedCell.ColSpan;
        end else begin
          Text:=Cells[i,j];
        end;

        var Letters:=GetMaxLineLength(Text)/ColSpan;
        if Letters>MaxLetters then begin
          MaxLetters:=Letters;
          var Line:=GetMaxLine(Text);

          var CellStyle:=CellStyle[i,j];
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

          MaxWidth:=Max(MaxWidth,Canvas.TextWidth(Line)/ColSpan-(ColSpan-1)*CellDelimterWidth);
        end;
      end;
    end else begin
      // Полный режим (с проверкой WordWrap)
      for j:=0 to FRowCount-1 do begin
        var ColSpan:=1;
        var MergedCell:TMergedCell;
        var IsMerged:=IsMergedCell(i,j,MergedCell);
        var ActualCol:Integer;

        if IsMerged then begin
          Text:=Cells[MergedCell.Col,MergedCell.Row];
          ColSpan:=MergedCell.ColSpan;
          ActualCol:=MergedCell.Col;
        end else begin
          Text:=Cells[i,j];
          ActualCol:=i;
        end;

        var CellStyle:=CellStyle[ActualCol,j];

        // Определяем, включен ли перенос слов для этой ячейки
        var WordWrapEnabled: Boolean;
        if CellStyle.WordWrapIsSet then begin
          WordWrapEnabled:=CellStyle.WordWrap;
        end else begin
          WordWrapEnabled:=FWordWrap or FColData[i].WordWrap;
        end;

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

        if WordWrapEnabled then begin
          // При WordWrap измеряем ширину самого длинного слова
          var Words:=Text.Split([' ', ':', ';', ',', '.', '!', '?', '-', '+', '*', '/', '\', '|', #9]);
          var MaxWordWidth:=0.0;
          for var Word in Words do begin
            if Word<>'' then begin
              MaxWordWidth:=Max(MaxWordWidth,Canvas.TextWidth(Word));
            end;
          end;
          MaxWidth:=Max(MaxWidth,MaxWordWidth/ColSpan-(ColSpan-1)*CellDelimterWidth/2);
        end else begin
          // Без WordWrap измеряем все строки
          var Lines:=Text.Split([#13#10]);
          for var Line in Lines do begin
            MaxWidth:=Max(MaxWidth,Canvas.TextWidth(Line)/ColSpan-(ColSpan-1)*CellDelimterWidth/2);
          end;
        end;
      end;
    end;

    // Добавляем отступы
    MaxWidth:=MaxWidth+CellPaddingFull;

    // Устанавливаем новую ширину
    var NewWidth:=Trunc(MaxWidth);
    if NewWidth<FDefaultColWidth then
      NewWidth:=FDefaultColWidth;

    ColWidths[i]:=NewWidth;
  end;
  Invalidate;
end;

procedure TMultiHeaderGrid.AutoSizeRows(ForcePrecise: boolean = False);
begin
  AutoSizeRows(0, FRowCount-1, ForcePrecise);
end;

type
  TSizeComputeMode = (cmFast,cmSlow,cmFull);

procedure TMultiHeaderGrid.AutoSizeRows(FromRow, ToRow: integer; ForcePrecise: boolean = False);
begin
  Canvas.Font.Assign(FCellFont);

  var CellPaddingHeight:=CellPadding.Top+CellPadding.Bottom;
  var CellDelimterHeight:=FGridLineWidth;
  var ViewBottomCell:=ViewBottom;

  var ComputeMode:=TSizeComputeMode.cmSlow;
  if FRowCount*FColCount>10000 then begin
    ComputeMode:=TSizeComputeMode.cmFast;
  end;
  if ForcePrecise then begin
    ComputeMode:=TSizeComputeMode.cmSlow;
  end;
  if FWordWrap or FGridCellsHasWordWrap then begin
    ComputeMode:=TSizeComputeMode.cmFull;
  end;

  Canvas.Font.Assign(FCellFont);
  var TH:=Canvas.TextHeight('А');
  var Text:='';
  for var Row:=FromRow to FRowCount-1 do begin
    if ((ToRow<0) and (Row>0) and (FRowData[Row-1].Top>ViewBottomCell)) or
       ((ToRow>=0) and (Row>ToRow)) then Break;
    var MaxHeight:=0.0;
    for var Col:=0 to FColCount-1 do begin

      var RowSpan:=1;
      var MergedCell: TMergedCell;
      var TextLines: Single;
      if IsMergedCell(Col,Row,MergedCell) then begin
        TextLines:=Max(CountLines(Cells[MergedCell.Col,MergedCell.Row]),1);
        RowSpan:=MergedCell.RowSpan;
      end else begin
        TextLines:=Max(CountLines(Cells[Col,Row]),1);
      end;

      case ComputeMode of
        cmFast: begin
          // Быстрый рассчет

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

          MaxHeight:=Max(MaxHeight,TextLines*ResTH/RowSpan-(RowSpan-1)*CellDelimterHeight/4);
        end;
        cmSlow: begin
          // Обычный рассчет

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

          MaxHeight:=Max(MaxHeight,TextLines*Canvas.TextHeight('А')/RowSpan-(RowSpan-1)*CellDelimterHeight/4);
        end;

        cmFull: begin
          // Рассчет с учетом WordWrap, медленно.

          var AvailableWidth:=GetCellRect(Col,Row,MergedCell).Width;

          var ActualCol: integer;
          var ActualRow: integer;
          if (MergedCell.ColSpan>1) or (MergedCell.RowSpan>1) then begin
            Text:=Cells[MergedCell.Col,MergedCell.Row];
            RowSpan:=MergedCell.RowSpan;
            ActualCol:=MergedCell.Col;
            ActualRow:=MergedCell.Row;
          end else begin
            Text:=Cells[Col,Row];
            ActualCol:=Col;
            ActualRow:=Row;
          end;

          var CellStyle:=CellStyle[ActualCol,ActualRow];

          // Определяем, включен ли перенос слов
          var WordWrapEnabled: Boolean;
          if CellStyle.WordWrapIsSet then begin
            WordWrapEnabled:=CellStyle.WordWrap;
          end else begin
            WordWrapEnabled:=FWordWrap or FColData[ActualCol].WordWrap;
          end;

          // Настраиваем шрифт
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

          var TextHeight:Single;
          if WordWrapEnabled and (Text<>'') then begin
            // Используем MeasureText для расчета
            var MeasureRect:=TRectF.Create(0,0,AvailableWidth,10000);
            Canvas.MeasureText(MeasureRect,Text,True,[],TTextAlign.Leading,TTextAlign.Leading);
            TextHeight:=MeasureRect.Height;
          end else begin
            // Без WordWrap считаем переводы строк
            var LineCount:=CountLines(Text);
            TextHeight:=Canvas.TextHeight('А')*LineCount-LineCount*FGridLineWidth;
          end;

          // Корректируем высоту
          MaxHeight:=Max(MaxHeight,TextHeight/RowSpan-(RowSpan-1)*CellDelimterHeight/4);
        end;
      end;
    end;

    FRowData[Row].Height:=Trunc(MaxHeight+CellPaddingHeight+CellDelimterHeight/2);
    if Row>0 then begin
      FRowData[Row].Top:=FRowData[Row-1].Top+FRowData[Row-1].Height;
    end;
  end;

  for var Row:=1 to FRowCount-1 do begin
    FRowData[Row].Top:=FRowData[Row-1].Top+FRowData[Row-1].Height;
  end;
  Invalidate;
end;

procedure TMultiHeaderGrid.AutoSizeVisibleRows;
begin
  AutoSizeRows(RowAtHeightCoord(ViewTop), -1, True);
end;


procedure TMultiHeaderGrid.ScrollToSelectedCell;
begin
  if (FSelectedCell.X>=0) and (FSelectedCell.Y>=0) then begin

    var Rect:=GetMergedCellRect(FSelectedCell.X,FSelectedCell.Y);
    Rect.Right:=Rect.Right+FGridLineWidth/2;
    Rect.Offset(ViewLeft,ViewTop-HeaderHeight);

    if Rect.Left<ViewLeft then begin
      ViewLeft:=Trunc(Rect.Left);
    end;

    if Rect.Right>ViewLeft+ViewPortWidth then begin
      ViewLeft:=Trunc(Rect.Right-ViewPortWidth)+1;
    end;

    if Rect.Top<ViewTop then begin
      ViewTop:=Trunc(Rect.Top);
    end;

    if (Rect.Bottom>ViewBottom-FGridLineWidth) then begin
      ViewTop:=Round(Rect.Bottom-ViewCellsHeight-FGridLineWidth/2)+1;
    end;

  end;
  Invalidate;
end;


procedure TMultiHeaderGrid.MouseDown(Button: TMouseButton; Shift: TShiftState; X, Y: Single);
var
  StartCol, EndCol : Integer;
  Row              : Integer;
  i, j             : Integer;
  Rect             : TRectF;
  MergedCell       : TMergedCell;
begin
  inherited;
  FLastClickIsOnCell:=False;

  if CanFocus and (Button=TMouseButton.mbLeft) then begin
    // Устанавливаем фокус при клике
    SetFocus;
    Invalidate;
  end;

  if Button=TMouseButton.mbLeft then begin
    // Проверяем клик по заголовкам
    if FResizeEnabled then begin
      var ResizeMode:=IsResizeArea(X, Y, StartCol, EndCol, Row);

      case ResizeMode of
        TResizeMode.rmColumn: StartColumnResize(StartCol, EndCol, X);
        TResizeMode.rmHeaderRow: StartHeaderRowResize(Row, Y);
        TResizeMode.rmGridRow: StartGridRowResize(Row, Y);
      end;

      if ResizeMode<>TResizeMode.rmNone then begin
        // Захватываем мышь для продолжения перетаскивания за пределами компонента
        Capture;
        Exit;
      end;
    end;

    // Проверяем клик по заголовкам
    for i:=0 to FHeaderLevels.Count-1 do begin
      j:=0;
      while j<FColCount do begin
        Rect:=GetHeaderRect(i, j);
        if Rect.Contains(PointF(X, Y)) then begin
          // Проверяем, является ли это объединенной ячейкой
          DoHeaderClick(i, j);
          Exit;
        end;
        var ColSpan:=FHeaderLevels.GetElementAtCell(j,i).ColSpan;
        if ColSpan<0 then ColSpan:=FColCount;
        j:=j+ColSpan;
      end;
    end;

    // Проверяем ячейки данных
    var TopRow:=RowAtHeightCoord(ViewTop);
    if TopRow<0 then Exit;
    var ViewBottom:=ViewTop+LocalRect.Top+LocalRect.Height+Margins.Bottom-HeaderHeight;

    for j:=TopRow to FRowCount-1 do begin
      if FRowData[j].Top>ViewBottom then Break;
      for i:=0 to FColCount-1 do begin
        Rect:=GetCellRect(i, j);
        if Rect.Contains(PointF(X, Y)) then begin
          FLastClickIsOnCell:=True;
          FSelectedCell:=Point(i, j);
          DoCellClick(FSelectedCell.X, FSelectedCell.Y);
          ScrollToSelectedCell;
          DoSelectCell;
          Exit;
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
          GlobalPos:=LocalToAbsolute(PointF(X, Y));
          NewWidth:=ResizeStartWidth+Round(GlobalPos.X-FResizeStartPos.X);
          UpdateColumnWidth(FResizeStartColumnIndex,FResizeEndColumnIndex, NewWidth);
          Cursor:=crHSplit;
        end;
      end;
      TResizeMode.rmHeaderRow: begin
        if FResizeRowIndex>=0 then begin
          // Используем глобальные координаты для точного отслеживания
          GlobalPos:=LocalToAbsolute(PointF(X, Y));
          NewHeight:=FResizeStartHeight+Round(GlobalPos.Y-FResizeStartPos.Y);
          UpdateHeaderRowHeight(FResizeRowIndex, NewHeight);
          Cursor:=crVSplit;
        end;
      end;
      TResizeMode.rmGridRow: begin
        if FResizeRowIndex>=0 then begin
          // Используем глобальные координаты для точного отслеживания
          GlobalPos:=LocalToAbsolute(PointF(X, Y));
          NewHeight:=FResizeStartHeight+Round(GlobalPos.Y-FResizeStartPos.Y);
          UpdateRowHeight(FResizeRowIndex, NewHeight);
          Cursor:=crVSplit;
        end;
      end
      else begin
        case IsResizeArea(X, Y, Col, Col, Row) of
          TResizeMode.rmColumn: Cursor:=crHSplit;
          TResizeMode.rmHeaderRow: Cursor:=crVSplit;
          TResizeMode.rmGridRow: Cursor:=crVSplit;
          else Cursor:=crDefault;
        end;
      end;
    end;
  end else begin
    Cursor:=crDefault;
  end;
end;

procedure TMultiHeaderGrid.MouseUp(Button: TMouseButton; Shift: TShiftState; X, Y: Single);
begin
  inherited;

  if (Button=TMouseButton.mbLeft) then begin
    // Завершаем изменение размера столбца
    if FResizeMode<>TResizeMode.rmNone then begin
      DoColumnResized;
      FResizeStartColumnIndex:=-1;
      FResizeEndColumnIndex:=-1;
      FResizeRowIndex:=-1;
      FResizeMode:=TResizeMode.rmNone;
      Cursor:=crDefault;
    end;
  end;
end;

procedure TMultiHeaderGrid.MouseWheel(Shift: TShiftState; WheelDelta: Integer;
  var Handled: Boolean);
begin
  inherited;
  if Handled then Exit;

  var WheelDeltaDivider:=2.0;
  if ssShift in Shift then begin
    WheelDeltaDivider:=0.5;
  end;

  WheelDelta:=-Max(1,round(Abs(WheelDelta/WheelDeltaDivider)))*Sign(WheelDelta);

  if ssHorizontal in Shift then begin
    ViewLeft:=ViewLeft+WheelDelta;
  end else begin
    ViewTop:=Min(ViewTop+WheelDelta,FullTableHeight);
  end;
  Invalidate;
end;

procedure CopyTextToClipboard(const AText: string);
var
  ClipboardService: IFMXClipboardService;
begin
  if TPlatformServices.Current.SupportsPlatformService(IFMXClipboardService, IInterface(ClipboardService)) then begin
    ClipboardService.SetClipboard(AText);
  end else
    raise Exception.Create('Сервис буфера обмена не поддерживается');
end;

procedure TMultiHeaderGrid.KeyDown(var Key: Word; var KeyChar: WideChar; Shift: TShiftState);
var
  NewCol, NewRow   : Integer;
  OldCol, OldRow   : Integer;
  MergedCellOld    : TMergedCell;
  MergedCellNew    : TMergedCell;
begin

  if IsFocused then begin
    var DX:=0;
    var DY:=0;

    case Key of
      vkLeft: begin
        DX:=-1;
      end;
      vkRight: begin
        DX:=1;
      end;
      vkUp: begin
        DY:=-1;
      end;
      vkDown: begin
        DY:=1;
      end;
      vkHome: begin
        DY:=-RowCount;
      end;
      vkEnd: begin
        DY:=RowCount;
      end;
      vkPrior: begin  // Page Up
        DY:=-Trunc(ViewCellsHeight/DefaultRowHeight);
      end;
      vkNext: begin   // Page Down
        DY:=Trunc(ViewCellsHeight/DefaultRowHeight);
      end;
      vkReturn: begin // Enter
        DoCellClick(FSelectedCell.X, FSelectedCell.Y);
        Exit;
      end;
      vkEscape: begin
        FSelectedCell:=Point(-1, -1);
        DoSelectCell;
        Invalidate;
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

  if Length(FRowData)=0 then
    Exit;

  // Быстрая проверка граничных случаев
  if Y<FRowData[0].Top then
    Exit;

  // Проверка последнего элемента
  CurrentBottom:=FRowData[High(FRowData)].Top+FRowData[High(FRowData)].Height;
  if Y>=CurrentBottom then
    Exit(High(FRowData)); // или -1, в зависимости от требований

  // Бинарный поиск с целочисленной арифметикой
  L:=0;
  R:=High(FRowData);

  while L<=R do begin
    M:=(L+R) shr 1; // Быстрее, чем div 2

    CurrentTop:=FRowData[M].Top;
    CurrentBottom:=CurrentTop+FRowData[M].Height;

    if Y>=CurrentTop then begin
      if Y<CurrentBottom then begin
        Result:=M;
        Exit;
      end else
        L:=M+1;
    end else
      R:=M-1;
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
  if Assigned(FOnSetCellStyle) then begin
    FOnSetCellStyle(Self, ACol, ARow, Style);
    if Style.WordWrapIsSet and Style.WordWrap then FGridCellsHasWordWrap:=True;
  end;
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
    Invalidate;
  end;
end;

procedure TMultiHeaderGrid.SetGridLineColor(const Value: TAlphaColor);
begin
  if FGridLineColor<>Value then begin
    FGridLineColor:=Value;
    Invalidate;
  end;
end;

procedure TMultiHeaderGrid.SetGridLineWidth(const Value: Single);
begin
  if FGridLineWidth<>Value then begin
    FGridLineWidth:=Value;
    Invalidate;
  end;
end;

procedure TMultiHeaderGrid.SetSelectedCell(const Value: TPoint);
begin
  if (FSelectedCell.X<>Value.X) or (FSelectedCell.Y<>Value.Y) then begin
    FSelectedCell:=Value;
    DoSelectCell;
    Invalidate;
  end;
end;

procedure TMultiHeaderGrid.SetSelectedCellColor(const Value: TAlphaColor);
begin
  if FSelectedCellColor<>Value then begin
    FSelectedCellColor:=Value;
    Invalidate;
  end;
end;

procedure TMultiHeaderGrid.SetSelectedFontColor(const Value: TAlphaColor);
begin
  if FSelectedFontColor<>Value then begin
    FSelectedFontColor:=Value;
    Invalidate;
  end;
end;

procedure TMultiHeaderGrid.SetCellFont(const Value: TFont);
begin
  if FCellFont<>Value then begin
    FCellFont.Assign(Value);
    Invalidate;
  end;
end;

procedure TMultiHeaderGrid.SetCellFontColor(const Value: TAlphaColor);
begin
  if FCellFontColor<>Value then begin
    FCellFontColor:=Value;
    Invalidate;
  end;
end;

procedure TMultiHeaderGrid.SetCellPadding(const Value: TBounds);
begin
  if FCellPadding<>Value then begin
    FCellPadding:=Value;
    Invalidate;
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
    Invalidate;
  end;
end;

procedure TMultiHeaderGrid.SetHeaderFont(const Value: TFont);
begin
  if FHeaderFont<>Value then begin
    FHeaderFont.Assign(Value);
    Invalidate;
  end;
end;

procedure TMultiHeaderGrid.SetHeaderCellColor(const Value: TAlphaColor);
begin
  if FHeaderCellColor<>Value then begin
    FHeaderCellColor:=Value;
    Invalidate;
  end;
end;

procedure TMultiHeaderGrid.SetHeaderFontColor(const Value: TAlphaColor);
begin
  if FHeaderFontColor<>Value then begin
    FHeaderFontColor:=Value;
    Invalidate;
  end;
end;

procedure TMultiHeaderGrid.SetHeaderLineColor(const Value: TAlphaColor);
begin
  if FHeaderLineColor<>Value then begin
    FHeaderLineColor:=Value;
    Invalidate;
  end;
end;

procedure TMultiHeaderGrid.SetHorisontalScroll(const Value: TScrollShowMode);
begin
  if FHorisontalScroll<>Value then begin
    FHorisontalScroll:=Value;
    UpdateSize;
    Invalidate;
  end;
end;

procedure TMultiHeaderGrid.SetCellColorAlternate(const Value: TAlphaColor);
begin
  if FCellColorAlternate<>Value then begin
    FCellColorAlternate:=Value;
    Invalidate;
  end;
end;

procedure TMultiHeaderGrid.SetBackgroundColor(const Value: TAlphaColor);
begin
  if FBackgroundColor<>Value then begin
    FBackgroundColor:=Value;
    Invalidate;
  end;
end;

procedure TMultiHeaderGrid.SetVerticalScroll(const Value: TScrollShowMode);
begin
  if FVerticalScroll<>Value then begin
    FVerticalScroll:=Value;
    UpdateSize;
    Invalidate;
  end;
end;

procedure TMultiHeaderGrid.SetViewLeft(const Value: integer);
begin
  var LastColLeft:=FullTableWidth-FColData[FColCount-1].Widths;
  FViewLeft:=Min(LastColLeft,Max(0,Value));

  var Event:=HScrollBar.OnChange;
  HScrollBar.OnChange:=nil;
  HScrollBar.Value:=Value;
  HScrollBar.OnChange:=Event;

  DoGridScroll;

  Invalidate;
end;

procedure TMultiHeaderGrid.SetViewTop(const Value: integer);
begin
  var BottomCellTop:=FRowData[High(FRowData)].Top;
  FViewTop:=Min(BottomCellTop,Max(0,Value));

  var Event:=VScrollBar.OnChange;
  VScrollBar.OnChange:=nil;
  VScrollBar.Value:=Value;
  VScrollBar.OnChange:=Event;

  DoGridScroll;

  Invalidate;
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

procedure TCellStyle.SetCellColorIsSet(const Value: Boolean);
begin
  FCellColorIsSet:=Value;
end;

procedure TCellStyle.SetFontColorIsSet(const Value: Boolean);
begin
  FFontColorIsSet:=Value;
end;

procedure TCellStyle.SetFontNameIsSet(const Value: Boolean);
begin
  FFontNameIsSet:=Value;
end;

procedure TCellStyle.SetFontSizeIsSet(const Value: Boolean);
begin
  FFontSizeIsSet:=Value;
end;

procedure TCellStyle.SetFontStyleIsSet(const Value: Boolean);
begin
  FFontStyleIsSet:=Value;
end;

procedure TCellStyle.SetSelectedCellColorIsSet(const Value: Boolean);
begin
  FSelectedCellColorIsSet:=Value;
end;

procedure TCellStyle.SetSelectedFontColorIsSet(const Value: Boolean);
begin
  FSelectedFontColorIsSet:=Value;
end;

procedure TCellStyle.SetTextHAlignment(const Value: TTextAlign);
begin
  FTextHAlignment:=Value;
  FTextHAlignmentIsSet:=True;
end;

procedure TCellStyle.SetTextHAlignmentIsSet(const Value: Boolean);
begin
  FTextHAlignmentIsSet:=Value;
end;

procedure TCellStyle.SetTextVAlignment(const Value: TTextAlign);
begin
  FTextVAlignment:=Value;
  FTextVAlignmentIsSet:=True;
end;

procedure TCellStyle.SetTextVAlignmentIsSet(const Value: Boolean);
begin
  FTextVAlignmentIsSet:=Value;
end;

procedure TCellStyle.SetWordWrap(const Value: Boolean);
begin
  FWordWrap:=Value;
  FWordWrapIsSet:=True;
end;

procedure TCellStyle.SetWordWrapIsSet(const Value: Boolean);
begin
  FWordWrapIsSet:=Value;
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
  if FResizeEnabled<>Value then begin
    FResizeEnabled:=Value;
    // Сбрасываем состояние изменения размера при отключении
    if not Value then begin
      FResizeStartColumnIndex:=-1;
      FResizeEndColumnIndex:=-1;
      FResizeRowIndex:=-1;
      ReleaseCapture; // Освобождаем захват мыши
      Cursor:=crDefault;
    end;
  end;
end;

function TMultiHeaderGrid.GetResizeMargin: Integer;
begin
  Result:=FResizeMargin;
end;

procedure TMultiHeaderGrid.SetResizeMargin(const Value: Integer);
begin
  if FResizeMargin<>Value then begin
    FResizeMargin:=Value;
    if FResizeMargin<1 then
      FResizeMargin:=1;
  end;
end;

function TMultiHeaderGrid.IsResizeArea(X, Y: Single; out AStartCol, AEndCol, ARow: Integer): TResizeMode;
var
  i,j        : Integer;
  Col         : integer;
  CellRect   : TRectF;
  ResizeRect : TRectF;
begin
  Result:=TResizeMode.rmNone;
  AStartCol:=-1;
  AEndCol:=-1;
  ARow:=-1;

  if not FResizeEnabled then Exit;

  if Y<=HeaderHeight then begin
    // Проверяем изменение ширины столбцов заголовка

    if FResizeColEnabled then begin
      for i:=0 to FHeaderLevels.Count-1 do begin
        Col:=0;
        for j:=0 to FHeaderLevels[i].Count-1 do begin
          CellRect:=GetHeaderRect(i, Col);

          ResizeRect:=TRectF.Create(
            CellRect.Right-FResizeMargin-FGridLineWidth/4,
            CellRect.Top,
            CellRect.Right+FResizeMargin+FGridLineWidth/4,
            CellRect.Bottom
          );

          if ResizeRect.Contains(PointF(X, Y)) then begin
            AStartCol:=Col;
            AEndCol:=Col+FHeaderLevels[i][j].ColSpan-1;
            Result:=TResizeMode.rmColumn;
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
          CellRect:=GetHeaderRect(i, Col);

          ResizeRect:=TRectF.Create(
            CellRect.Left,
            CellRect.Bottom-FResizeMargin-FGridLineWidth/4,
            CellRect.Right,
            CellRect.Bottom+FResizeMargin+FGridLineWidth/4
          );

          if ResizeRect.Contains(PointF(X, Y)) then begin
            ARow:=i;
            Result:=TResizeMode.rmHeaderRow;
            Exit;
          end;

          Col:=Col+FHeaderLevels[i][j].ColSpan;
        end;
      end;
    end;
  end else begin
    if FResizeRowEnabled then begin
      // Проверяем изменение высоты строк грида

      for i:=0 to FRowCount-1 do begin
        CellRect:=GetCellRect(0, i); // Получаем прямоугольник строки
        if not CellRect.IsEmpty then begin
          // Проверяем нижнюю границу строки
          ResizeRect:=TRectF.Create(
            CellRect.Left-FGridLineWidth/4,
            CellRect.Bottom-FResizeMargin-FGridLineWidth/4,
            Min(CellRect.Right,CellRect.Left+FResizeMargin)+FGridLineWidth/2,
            CellRect.Bottom+FResizeMargin+FGridLineWidth/4
          );

          if ResizeRect.Contains(PointF(X, Y)) then begin
            ARow:=i;
            Result:=TResizeMode.rmGridRow;
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
  i: Integer;
  GroupColCount: Integer;
  NewWidth, RemainingWidth: Integer;
  ProportionalWidths: array of Integer;
begin
  if TotalWidth<10*(EndCol-StartCol+1) then
    TotalWidth:=10*(EndCol-StartCol+1); // Минимальная общая ширина

  GroupColCount:=EndCol-StartCol+1;
  SetLength(ProportionalWidths, GroupColCount);

  // Вычисляем пропорциональные ширины на основе исходных соотношений
  RemainingWidth:=TotalWidth;

  // Первый проход: вычисляем пропорциональные ширины
  var FResizeStartWidth:=ResizeStartWidth;
  for i:=0 to GroupColCount-1 do begin
    ProportionalWidths[i]:=Round(FResizeStartWidths[i]*TotalWidth/FResizeStartWidth);
    Dec(RemainingWidth, ProportionalWidths[i]);
  end;

  // Второй проход: распределяем остаток
  if RemainingWidth<>0 then begin
    for i:=0 to GroupColCount-1 do begin
      if RemainingWidth>0 then begin
        Inc(ProportionalWidths[i]);
        Dec(RemainingWidth);
      end else if RemainingWidth<0 then begin
        Dec(ProportionalWidths[i]);
        Inc(RemainingWidth);
      end else begin
        Break;
      end;
    end;
  end;

  // Применяем новые ширины
  for i:=StartCol to EndCol do begin
    NewWidth:=ProportionalWidths[i-StartCol];
    if NewWidth<10 then
      NewWidth:=10; // Минимальная ширина столбца

    ColWidths[i]:=NewWidth;
  end;

  Invalidate;

  // Обновляем горизонтальную прокрутку
  if Assigned(HScrollBar) then
    HScrollBar.Value:=FViewLeft;
end;

procedure TMultiHeaderGrid.UpdateColumnWidth(StartCol, EndCol: Integer; NewWidth: Integer);
begin
  if StartCol<EndCol then begin
    // Ресайз группы колонок
    UpdateGroupColumnWidth(StartCol, EndCol, NewWidth);
  end else begin
    // Ресайз одной колонки
    if NewWidth<10 then begin
      NewWidth:=10; // Минимальная ширина
    end;

    ColWidths[EndCol]:=NewWidth;

    Invalidate;

    // Обновляем горизонтальную прокрутку
    if Assigned(HScrollBar) then begin
      HScrollBar.Value:=FViewLeft;
    end;
  end;
end;

procedure TMultiHeaderGrid.UpdateHeaderRowHeight(ARow: Integer; NewHeight: Integer);
begin
  if NewHeight<10 then
    NewHeight:=10; // Минимальная высота

  FHeaderLevels[ARow].Height:=NewHeight;

  Invalidate;

  // Обновляем вертикальную прокрутку
  if Assigned(VScrollBar) then begin
    VScrollBar.BeginUpdate;
    VScrollBar.Value:=FViewTop;
    VScrollBar.EndUpdate;
  end;
end;

procedure TMultiHeaderGrid.UpdateRowHeight(ARow: Integer; NewHeight: Integer);
begin
  if NewHeight<10 then
    NewHeight:=10; // Минимальная высота

  SetRowHeight(ARow, NewHeight);
  Invalidate;

  // Обновляем вертикальную прокрутку
  if Assigned(VScrollBar) then begin
    VScrollBar.BeginUpdate;
    VScrollBar.Value:=FViewTop;
    VScrollBar.EndUpdate;
  end;
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
    Invalidate;
  end;
end;

procedure TMultiHeaderGrid.DblClick;
begin
  if not FLastClickIsOnCell then Exit;

  inherited;
end;

function TMultiHeaderGrid.GetColTextHAlignment(Index: Integer): TTextAlign;
begin
  if (Index>=0) and (Index<Length(FColData)) then
    Result:=FColData[Index].TextHAlignment
  else
    Result:=TTextAlign.Leading;
end;

function TMultiHeaderGrid.GetColTextVAlignment(Index: Integer): TTextAlign;
begin
  if (Index>=0) and (Index<Length(FColData)) then
    Result:=FColData[Index].TextVAlignment
  else
    Result:=TTextAlign.Center;
end;

procedure TMultiHeaderGrid.SetColTextHAlignment(Index: Integer; const Value: TTextAlign);
begin
  if (Index>=0) and (Index<Length(FColData)) and (FColData[Index].TextHAlignment<>Value) then begin
    FColData[Index].TextHAlignment:=Value;
    UpdateSize;
    Invalidate;
  end;
end;

procedure TMultiHeaderGrid.SetColTextVAlignment(Index: Integer; const Value: TTextAlign);
begin
  if (Index>=0) and (Index<Length(FColData)) and (FColData[Index].TextVAlignment<>Value) then begin
    FColData[Index].TextVAlignment:=Value;
    UpdateSize;
    Invalidate;
  end;
end;

function TMultiHeaderGrid.GetColWordWrap(Index: Integer): Boolean;
begin
  if (Index>=0) and (Index<FColCount) then
    Result:=FColData[Index].WordWrap
  else
    Result:=False;
end;

procedure TMultiHeaderGrid.SetColWordWrap(Index: Integer; const Value: Boolean);
begin
  if (Index>=0) and (Index<FColCount) then begin
    FColData[Index].WordWrap:=Value;
    Invalidate;
  end;
end;

procedure TMultiHeaderGrid.ClearSelection;
begin
  SelectedCell:=Point(-1,-1);
end;

procedure TMultiHeaderGrid.SetWordWrap(const Value: Boolean);
begin
  if FWordWrap<>Value then begin
    FWordWrap:=Value;
    Invalidate;
  end;
end;

procedure TMultiHeaderGrid.Invalidate;
begin
  InvalidateRect(LocalRect);
end;

end.


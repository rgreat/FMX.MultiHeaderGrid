unit FMX.MultiHeaderGrid;

interface

uses
  System.Classes, System.Types, System.UITypes, System.Generics.Collections,
  FMX.Types, FMX.Controls, FMX.Graphics, FMX.StdCtrls, FMX.Objects, FMX.Layouts,
  FMX.Memo, FMX.ListBox, FMX.DateTimeCtrls, FMX.Edit, FMX.EditBox,
  FMX.NumberBox, FMX.Text, Data.DB;

type

  THeaderLevel = class;
  THeaderLevels = class;

  // Kind of inplace editor the DB grid uses for a column, selected from the
  // bound field's DataType / FieldKind. See TMultiHeaderDBGrid.EditorKindForField.
  TMHGEditorKind = (
    ekNone,      // not editable (binary/structured/calculated/read-only)
    ekMemo,      // TMemo - text fields, value round-trips via Field.AsString
    ekNumber,    // TNumberBox - integer / float fields (integer vs decimal mode)
    ekCheckBox,  // ftBoolean - toggled in place (no editor control)
    ekComboBox,  // TComboBox - lookup fields (fkLookup)
    ekDate,      // TDateEdit - ftDate
    ekTime,      // TTimeEdit - ftTime
    ekDateTime   // TDateEdit + TTimeEdit composite - ftDateTime/ftTimeStamp
  );

  // Per-column choice of how a datetime-typed field is edited. Lets the
  // column author override the default (composite date+time) editor.
  //   dteDateTime : date + time composite (default for datetime fields)
  //   dteDate     : date only (time component left untouched on commit)
  TMHGDateTimeEditKind = (
    dteDateTime,
    dteDate
  );

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
    constructor Create(HeaderLevels: THeaderLevels; InsertBefore: integer = -1);

    function AddColumn(Caption: string = '';
                       ColSpan: Integer = 1;
                       RowSpan: Integer = 1): THeaderElement;

    function FillRow(Caption: string = ''): THeaderElement;

    property Height: Integer read FHeight write SetHeight;
  end;

  TMultiHeaderGrid = class;
  TMHGHeaderColumns = class;

  // Base, non-data-bound column descriptor shared by all grids. Carries
  // everything needed to build a (possibly grouped, multi-row) header and
  // to lay out a column: Title, GroupHeader/Separator, widths, word wrap
  // and the four alignments. The data-bound TMHGColumn (DB grid) extends
  // this with FieldName and per-column data Color.
  TMHGHeaderColumn = class(TCollectionItem)
  private
    FTitle                 : string;
    FWidth                 : Integer;
    FMinWidth              : Integer;
    FMaxWidth              : Integer;
    FVisible               : Boolean;
    FWordWrap              : Boolean;
    FAlignment             : TTextAlign;        // data cells, horizontal
    FVertAlignment         : TTextAlign;        // data cells, vertical
    FHeaderAlignment       : TTextAlign;        // header, horizontal
    FHeaderVertAlignment   : TTextAlign;        // header, vertical
    FGroupHeader           : string;
    FGroupHeaderSeparator  : Char;

    procedure SetTitle(const Value: string);
    procedure SetWidth(const Value: Integer);
    procedure SetMinWidth(const Value: Integer);
    procedure SetMaxWidth(const Value: Integer);
    procedure SetVisible(const Value: Boolean);
    procedure SetWordWrap(const Value: Boolean);
    procedure SetAlignment(const Value: TTextAlign);
    procedure SetVertAlignment(const Value: TTextAlign);
    procedure SetHeaderAlignment(const Value: TTextAlign);
    procedure SetHeaderVertAlignment(const Value: TTextAlign);
    procedure SetGroupHeader(const Value: string);
    procedure SetGroupHeaderSeparator(const Value: Char);
  protected
    function GetDisplayName: string; override;
    // Splits GroupHeader into its individual group-level captions.
    function GroupPath: TArray<string>;
    procedure Changed;
  public
    constructor Create(Collection: TCollection); override;
    procedure Assign(Source: TPersistent); override;

    // Sets several common properties at once; returns Self for chaining.
    // Width/MinWidth/MaxWidth of -1 mean "leave unchanged".
    function SetProps(const ATitle: string;
                      const AGroupHeader: string = '';
                      AWidth: Integer = -1;
                      AAlignment: TTextAlign = TTextAlign.Leading;
                      AMinWidth: Integer = -1;
                      AMaxWidth: Integer = -1): TMHGHeaderColumn;
  published
    property Title: string read FTitle write SetTitle;
    property Width: Integer read FWidth write SetWidth default 80;
    property MinWidth: Integer read FMinWidth write SetMinWidth default 0;
    property MaxWidth: Integer read FMaxWidth write SetMaxWidth default 0;
    property Visible: Boolean read FVisible write SetVisible default True;
    property WordWrap: Boolean read FWordWrap write SetWordWrap default False;
    property Alignment: TTextAlign read FAlignment write SetAlignment default TTextAlign.Leading;
    property VertAlignment: TTextAlign read FVertAlignment write SetVertAlignment default TTextAlign.Center;
    property HeaderAlignment: TTextAlign read FHeaderAlignment write SetHeaderAlignment default TTextAlign.Center;
    property HeaderVertAlignment: TTextAlign read FHeaderVertAlignment write SetHeaderVertAlignment default TTextAlign.Center;
    property GroupHeader: string read FGroupHeader write SetGroupHeader;
    property GroupHeaderSeparator: Char read FGroupHeaderSeparator write SetGroupHeaderSeparator default ';';
  end;

  TMHGHeaderColumns = class(TOwnedCollection)
  private
    FGrid: TMultiHeaderGrid;
    function GetItem(Index: Integer): TMHGHeaderColumn;
    procedure SetItem(Index: Integer; const Value: TMHGHeaderColumn);
  protected
    procedure Update(Item: TCollectionItem); override;
  public
    // AItemClass lets the data-bound TMHGColumns reuse this collection
    // with its richer item class (TMHGColumn) without a virtual call
    // during construction.
    constructor Create(AGrid: TMultiHeaderGrid;
                       AItemClass: TCollectionItemClass = nil); reintroduce; overload;
    function Add: TMHGHeaderColumn;
    function AddColumn(const ATitle: string = '';
                       const AGroupHeader: string = ''): TMHGHeaderColumn;
    property Grid: TMultiHeaderGrid read FGrid;
    property Items[Index: Integer]: TMHGHeaderColumn read GetItem write SetItem; default;
  end;

  THeaderLevels = class(TObjectList<THeaderLevel>)
    Grid : TMultiHeaderGrid;

    constructor Create(Grid : TMultiHeaderGrid);

    function AddRow: THeaderLevel; overload;
    function AddRow(Heigth: integer): THeaderLevel; overload;
    function AddRowOnTop: THeaderLevel; overload;
    function AddRowOnTop(Heigth: integer): THeaderLevel; overload;
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
    type
      TRowData = packed record
        Top    : integer;
        Height : Word;
      end;

    var
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
      FHeaderWordWrap: Boolean;
      FHeaderColumns: TMHGHeaderColumns;
      FRebuildingHeaderColumns: Boolean;
      FGridCellsHasWordWrap: boolean;
      FVerticalScroll: TScrollShowMode;
      FHorisontalScroll: TScrollShowMode;

      // Inplace cell editor (basic grids). A single reusable TMemo is moved
      // over the cell being edited.
      FEditor: TMemo;
      FEditorBack: TRectangle;
      FEditorHost: TLayout;
      // Currently active inplace editor control. For the base/string grids
      // this is always FEditor (the TMemo). The DB grid may swap in a typed
      // control (TComboBox/TDateEdit/TTimeEdit/...) per field type. (Boolean
      // fields use no editor - they toggle in place.)
      FActiveEditor: TControl;
      FEditing: Boolean;
      FEditCol: Integer;
      FEditRow: Integer;
      FReadOnly: Boolean;

    function ResizeStartWidth: Integer;
    procedure SetHeaderWordWrap(const Value: Boolean);
    procedure SetHeaderColumns(const Value: TMHGHeaderColumns);
    // Builds a (possibly grouped, multi-row) header + applies column
    // geometry/alignment/wordwrap from a TMHGHeaderColumns collection.
    // Shared by the base grid (HeaderColumns) and reused conceptually by
    // the DB grid. AVisible returns the visible items in order.
    procedure BuildHeaderFromColumns(const ACols: array of TMHGHeaderColumn);
    procedure RebuildFromHeaderColumns;
    // True when wrapping applies to the header element at (ALevel,ACol).
    function HeaderCellWordWrap(AElement: THeaderElement): Boolean;
    // Width available to a header element's text (merged rect, padded).
    function HeaderElementTextWidth(ALevel, ACol: Integer): Single;

    procedure StartCellEditing(ACol, ARow: Integer; InitialChar: Char);
    // Inplace TMemo editor helpers.
    procedure EnsureEditor(TextAlign: TTextAlign);
    procedure PositionEditor; virtual;
    procedure CommitEditing;
    procedure CancelEditing; virtual;
    procedure EditorKeyDown(Sender: TObject; var Key: Word; var KeyChar: Char;
                            Shift: TShiftState);
    procedure EditorExit(Sender: TObject);
    // Strips control from painted border/background so the inplace editor
    // blends seamlessly into the cell.
    procedure ApplyStyle(Sender: TObject);
  protected
    // --- Inplace editor extensibility hooks -----------------------------
    // The base/string grids always edit through the shared TMemo (FEditor).
    // The DB grid overrides these to select a typed editor control per field
    // (TMemo / TComboBox / TDateEdit / TTimeEdit / composite). Boolean fields
    // are not edited through a control - they toggle in place.
    //
    // PrepareCellEditor: choose/create the control for (ACol,ARow), parent it,
    // wire OnKeyDown/OnExit, and return it. Base returns the shared TMemo.
    function  PrepareCellEditor(ACol, ARow: Integer): TControl; virtual;
    // Load the cell's current value into the active editor (InitialChar<>#0
    // means the user started typing - replace content with that char).
    procedure LoadEditorValue(ACol, ARow: Integer; InitialChar: Char); virtual;
    // Read the active editor's value back as display text.
    function  GetEditorText: string; virtual;
    // Persist the editor's value into the cell/field. Base writes Cells[].
    // Returns True if the value was accepted (commit succeeded).
    function  CommitEditorValue(ACol, ARow: Integer): Boolean; virtual;
    // Active editor control (FActiveEditor, or the TMemo as a fallback).
    function  ActiveEditorControl: TControl; virtual;
    // After the editor is shown, place the caret at the end of its text (no
    // selection). Base handles the shared TMemo; descendants override for
    // their own typed editors.
    procedure PlaceEditorCaretAtEnd(Ed: TControl); virtual;
    // Minimum column width (px) the chosen editor needs to be usable for the
    // cell at (ACol,ARow). 0 means "no requirement" (the memo wraps/scrolls
    // and is fine in a narrow column). StartCellEditing widens the column to
    // at least this before positioning the editor. Descendants override to
    // request room for fixed-size controls (combo/date/time/datetime).
    function  EditorMinColWidth(ACol, ARow: Integer): single; virtual;
    procedure SetReadOnly(const Value: Boolean);
    function  CanEditCell(ACol, ARow: Integer): Boolean; virtual;
    // Boolean (toggle) cells are not edited through an inplace control: they
    // flip in place on a click / Space / Enter. CellIsToggle reports such a
    // cell; ToggleCell flips its value and persists it. The base grid has no
    // such cells (both are no-ops); TMultiHeaderDBGrid overrides them for
    // ftBoolean fields.
    function  CellIsToggle(ACol, ARow: Integer): Boolean; virtual;
    function  ToggleCell(ACol, ARow: Integer): Boolean; virtual;
    // Glyph rectangle for a toggle cell's checkbox, shared by drawing and
    // mouse hit-testing so a click on the box (vs. around it) is detected
    // consistently with what is painted.
    function  ToggleGlyphRect(ACol, ARow: Integer; const ARect: TRectF): TRectF; virtual;
    // Paints the built-in checkbox glyph for a toggle cell. Self-contained
    // (vector drawing, no image list). Descendants may override to restyle.
    procedure DrawToggleCell(Canvas: TCanvas; ACol, ARow: Integer;
                             const ARect: TRectF; IsSelected, AChecked: Boolean); virtual;
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
    function HeaderCellIsFiller(ALevel, ACol: Integer): Boolean;
    function HeaderMergedRect(ALevel, ACol: Integer): TRectF;
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
    function GridHaveWordWrap: boolean;
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
    // Sizes each header level's height to fit its (optionally wrapped or
    // multi-line) captions. Shared by all grid descendants.
    procedure AutoSizeHeaders;
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
    // When True the inplace cell editor is disabled and the grid is view-only.
    property ReadOnly: Boolean read FReadOnly write SetReadOnly default False;
    property WordWrap: Boolean read FWordWrap write SetWordWrap default False;
    // When set, header captions wrap to the cell width during drawing and
    // are accounted for by AutoSizeHeaders. On by default.
    property HeaderWordWrap: Boolean read FHeaderWordWrap write SetHeaderWordWrap default True;
    // Declarative header/column definitions for the non-data-bound grids.
    // When populated they build the (grouped) header and set column
    // widths/alignment/word wrap. When empty the grid keeps whatever
    // header was created procedurally via Header.AddRow.
    property HeaderColumns: TMHGHeaderColumns read FHeaderColumns write SetHeaderColumns;

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
  end;

  TMultiHeaderDBGrid = class;
  TMHGColumns = class;

  // Single column of TMultiHeaderDBGrid. Extends the shared
  // TMHGHeaderColumn (Title, GroupHeader, widths, alignments, word wrap)
  // with the data-bound bits: FieldName and a per-column data Color.
  //
  // Multi-line grouped headers are produced exactly as in UniGUI:
  // GroupHeader holds one or more group captions joined with
  // GroupHeaderSeparator (default ';'). Adjacent columns that share the
  // same leading group path are merged into spanning header cells.
  TMHGColumn = class(TMHGHeaderColumn)
  private
    FFieldName  : string;
    FColor      : TAlphaColor;
    FColorIsSet : Boolean;
    FDateTimeEditor : TMHGDateTimeEditKind;

    function GetGrid: TMultiHeaderDBGrid;
    procedure SetFieldName(const Value: string);
    procedure SetColor(const Value: TAlphaColor);
    function IsColorStored: Boolean;
  protected
    function GetDisplayName: string; override;
  public
    constructor Create(Collection: TCollection); override;
    procedure Assign(Source: TPersistent); override;

    property Grid: TMultiHeaderDBGrid read GetGrid;
    // Effective text colour flag (Color only matters when explicitly set).
    property ColorIsSet: Boolean read FColorIsSet;
  published
    // Name of the DataSet field shown in this column.
    property FieldName: string read FFieldName write SetFieldName;
    // Background colour of the data cells of this column.
    property Color: TAlphaColor read FColor write SetColor stored IsColorStored;
    // For datetime fields: whether this column edits date+time (default) or
    // date only. Ignored for non-datetime fields.
    property DateTimeEditor: TMHGDateTimeEditKind read FDateTimeEditor
      write FDateTimeEditor default dteDateTime;
  end;

  TMHGColumns = class(TOwnedCollection)
  private
    FGrid: TMultiHeaderDBGrid;
    function GetItem(Index: Integer): TMHGColumn;
    procedure SetItem(Index: Integer; const Value: TMHGColumn);
  protected
    procedure Update(Item: TCollectionItem); override;
  public
    constructor Create(AGrid: TMultiHeaderDBGrid);

    function Add: TMHGColumn;
    function AddColumn(const AFieldName: string;
                       const ATitle: string = '';
                       const AGroupHeader: string = ''): TMHGColumn;
    function FindByFieldName(const AFieldName: string): TMHGColumn;
    // Alters several basic properties of the column with the given field
    // name in one call (one rebuild). Width/MinWidth/MaxWidth of -1 mean
    // "leave unchanged". Returns the column (nil if the field name is not
    // found).
    function SetColumnProps(const AFieldName: string;
                            const ATitle: string;
                            const AGroupHeader: string = '';
                            AWidth: Integer = -1;
                            AAlignment: TTextAlign = TTextAlign.Leading;
                            AMinWidth: Integer = -1;
                            AMaxWidth: Integer = -1): TMHGColumn;

    property Grid: TMultiHeaderDBGrid read FGrid;
    property Items[Index: Integer]: TMHGColumn read GetItem write SetItem; default;
  end;

  // TDataLink through which TMultiHeaderDBGrid tracks changes
  // in DataSet/DataSource - open/close, cursor movement,
  // changes to the record set (filter, refresh, insert/delete) and
  // changes to field values of the current record.
  TMHGDataLink = class(TDataLink)
  private
    FGrid: TMultiHeaderDBGrid;
  protected
    procedure ActiveChanged; override;
    procedure LayoutChanged; override;
    procedure DataSetChanged; override;
    procedure DataSetScrolled(Distance: Integer); override;
    procedure RecordChanged(Field: TField); override;
  public
    constructor Create(AGrid: TMultiHeaderDBGrid);
  end;

  TMultiHeaderDBGrid = class(TMultiHeaderGrid)
  private
    type
      TDBRowData = record
        LastUsage : TDateTime;
        Cells     : TArray<string>;
        Styles    : TArray<TCellStyle>;
      end;
      TRowsData = TDictionary<integer,TDBRowData>;

    var
      FDataLink: TMHGDataLink;
      FCellTexts: TRowsData;
      FRowCacheSize: integer;
      FUpdatingRow: Boolean; // guards against recursion when syncing Row <-> DataSet.RecNo
      FColumns: TMHGColumns;
      // Signature of the DataSet/table we last auto-created columns for, so
      // we only auto-create on a genuinely new source or a newly opened table
      // (different field set), not on every re-open of the same one.
      FLastAutoSig: string;
      // Signature of the column layout (ordered field names) realised by the
      // last ResetTable. When unchanged, a rebuild is a pure re-apply (e.g. a
      // Min/MaxWidth tweak) and live column widths are preserved instead of
      // being reset to the design-time Col.Width.
      FLastColLayout: string;
      FRebuildingHeader: Boolean; // guards ResetTable re-entrancy from column changes
      // Effective column-index -> collection item, rebuilt by ResetTable.
      // Lets DoGetCellStyle read a column's Color in O(1) per cell.
      FColMap: TArray<TMHGColumn>;

      // --- Typed inplace editor controls --------------------------------
      // Created lazily, reused across edits. Only one is active at a time
      // (selected per field type); the active one is FActiveEditor.
      // Boolean fields use no editor control - they toggle in place
      // (see CellIsToggle / ToggleCell).
      FComboEditor: TComboBox;
      FNumberEditor: TNumberBox;
      FDateEditor : TDateEdit;
      FTimeEditor : TTimeEdit;
      // For ftDateTime the date control hosts the time control beside it; the
      // composite is positioned as a unit (see PositionEditor override).
      FEditField  : TField; // field bound to the editor currently open

    function GetDataSource: TDataSource;
    procedure SetDataSource(const Value: TDataSource);
    procedure SetRowCacheSize(const Value: integer);
    procedure SetColumns(const Value: TMHGColumns);
    procedure CleanUpCache;
    function GetVisibleFields: TArray<TField>;
    function GetRowData(ARow: Integer): TDBRowData;
    procedure SetFieldValue(DS: TDataSet; Field: TField; const Text: string);
    // Builds the multi-row grouped header from the Columns collection.
    procedure BuildGroupedHeader(const AFields: TArray<TField>;
                                 const ACols: TArray<TMHGColumn>);
    // Resolves the effective field list strictly from the Columns
    // collection (no fallback): an empty collection yields no columns.
    function ResolveColumns(out AFields: TArray<TField>): TArray<TMHGColumn>;
    // A string that identifies the currently open dataset/table: the DataSet
    // instance plus its field-name list. Changes when a different DataSet is
    // assigned OR the same DataSet opens a different set of fields.
    function DatasetSignature: string;
    // Called when the dataset becomes active. Auto-creates columns only when
    // the grid has none AND this is a genuinely new source/table (a signature
    // we have not auto-created for before), then rebuilds.
    procedure HandleActiveChanged;
  protected
    procedure DoGetCellText(ACol, ARow: Integer; var Text: string); override;
    procedure DoSetCellText(ACol, ARow: Integer; const Text: string); override;
    procedure DoGetCellStyle(ACol, ARow: Integer; var Style: TCellStyle); override;
    procedure DoSetCellStyle(ACol, ARow: Integer; const Style: TCellStyle); override;
    procedure DoSelectCell; override;
    // Data-bound inplace editing: a typed editor is chosen per field.
    function CanEditCell(ACol, ARow: Integer): Boolean; override;
    // Boolean fields toggle in place instead of opening an editor.
    function CellIsToggle(ACol, ARow: Integer): Boolean; override;
    function ToggleCell(ACol, ARow: Integer): Boolean; override;
    // --- Typed editor overrides (see base hooks) ------------------------
    function  PrepareCellEditor(ACol, ARow: Integer): TControl; override;
    procedure PositionEditor; override;
    procedure CancelEditing; override;
    procedure LoadEditorValue(ACol, ARow: Integer; InitialChar: Char); override;
    function  GetEditorText: string; override;
    function  CommitEditorValue(ACol, ARow: Integer): Boolean; override;
    procedure PlaceEditorCaretAtEnd(Ed: TControl); override;
    function  EditorMinColWidth(ACol, ARow: Integer): single; override;
    // Field bound to a given grid column (nil if out of range / no dataset).
    function  FieldForCol(ACol: Integer): TField;
    // Editor kind for a field, from DataType / FieldKind / ReadOnly.
    function  EditorKindForField(Field: TField): TMHGEditorKind;
    // True for whole-number field types (drives the number editor's integer
    // vs decimal mode).
    function  FieldIsInteger(Field: TField): Boolean;
    // Min/Max range the number editor should allow for a field's data type.
    procedure NumberRangeForField(Field: TField; out AMin, AMax: Double);
    // The Columns collection item that drives a given field (matched by
    // FieldName), or nil if the field isn't bound to a column.
    function  ColumnForField(Field: TField): TMHGColumn;
    // Lazily create & wire the typed editor controls.
    procedure EnsureTypedEditors;
    // OnApplyStyleLookup for the number editor: hides its spin buttons.
    procedure NumberEditorApplyStyle(Sender: TObject);
  public
    constructor Create(AOwner: TComponent); override;
    destructor Destroy; override;

    // Full grid rebuild: column headers (based on the DataSet's visible fields)
    // and the row count. Called automatically when DataSource changes,
    // the dataset is opened/closed, or the field list changes.
    procedure ResetTable;

    // Recalculates RowCount from DataSet.RecordCount, clears the row cache
    // and syncs the selected row with the current DataSet record.
    // Called automatically when the record set changes (filter,
    // Refresh, record insertion/deletion, etc.).
    procedure UpdateRowCount;

    // Moves DataSet.RecNo to the selected grid row (Row).
    // Called automatically when the DataSet cursor moves.
    procedure SyncSelectedRowFromDataSet;

    // Discards the cached values of the current DataSet record
    // (called automatically when fields of the current record change).
    procedure InvalidateCurrentRow;

    function DataSet: TDataSet;
    function Column(FieldNum: integer): THeaderElement; overload;
    function Column(FieldName: string): THeaderElement; overload;

    // Re-reads the DataSet's field list and (re)creates a default
    // Columns entry for every visible field that is not yet present.
    // Existing columns (and their grouping) are preserved. Useful after
    // opening a DataSet when columns were not predefined at design time.
    procedure AutoCreateColumns;
  published
    property DataSource: TDataSource read GetDataSource write SetDataSource;
    property RowCacheSize: integer read FRowCacheSize write SetRowCacheSize default 1000;
    // Declarative column definitions. The collection is authoritative:
    // the grid shows exactly these columns, in this order. An empty
    // collection means an empty grid - use AutoCreateColumns (or the
    // design-time editor's Import button) to (re)populate it from the
    // DataSet's visible fields.
    property Columns: TMHGColumns read FColumns write SetColumns;
  end;

procedure Register;

implementation

uses
  System.SysUtils, System.Math, FMX.Platform, System.StrUtils;

// Forward declarations for unit-local text helpers used by methods that
// appear earlier in the implementation than the functions themselves.
function CountLines(const Text: string): Integer; forward;

procedure Register;
begin
  RegisterComponents('FMX', [TMultiHeaderGrid]);
  RegisterComponents('FMX', [TMultiHeaderStringGrid]);
  RegisterComponents('FMX', [TMultiHeaderDBGrid]);
end;

{ THeaderElement }

constructor THeaderElement.Create(HeaderLevel: THeaderLevel);
begin
  FLevel:=HeaderLevel;
  FColSpan:=1;
  FRowSpan:=1;
  HeaderLevel.Add(Self);
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

  // Calculate the current column
  var CurCol:=0;
  for var Element in Self do begin
    inc(CurCol,Element.FColSpan);
  end;

  var CurRow:=0;
  for var Level in FLevels do begin
    if Level=Self then Break;
    inc(CurRow);
  end;

  // Calculate the column skip
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

constructor THeaderLevel.Create(HeaderLevels: THeaderLevels; InsertBefore: integer = -1);
begin
  inherited Create;

  FLevels:=HeaderLevels;

  if InsertBefore>=0 then begin
    HeaderLevels.Insert(InsertBefore,Self);
  end else begin
    HeaderLevels.Add(Self);
  end;
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
  // Create an array to track occupied columns
  SetLength(Result, FLevels.Grid.ColCount);
  for i:=0 to High(Result) do begin
    Result[i]:=False;
  end;

  var LevelIndex:=FLevels.IndexOf(Self);

  // Iterate over all levels above the current one
  for i:=0 to LevelIndex-1 do begin
    for j:=0 to FLevels[i].Count-1 do begin
      Element:=FLevels[i][j];
      // If the element spans multiple rows and its RowSpan reaches the current level
      if (Element.RowSpan>1) and (i+Element.RowSpan>LevelIndex) then begin
        // Mark all columns occupied by this element
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

function THeaderLevels.AddRowOnTop(Heigth: integer): THeaderLevel;
begin
  Result:=THeaderLevel.Create(Self,0);
  Result.FHeight:=Heigth;
end;

function THeaderLevels.AddRowOnTop: THeaderLevel;
begin
  Result:=AddRowOnTop(25);
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

{ TMHGHeaderColumn }

constructor TMHGHeaderColumn.Create(Collection: TCollection);
begin
  inherited;
  FWidth:=80;
  FMinWidth:=0;
  FMaxWidth:=0;
  FVisible:=True;
  FWordWrap:=False;
  FAlignment:=TTextAlign.Leading;
  FVertAlignment:=TTextAlign.Center;
  FHeaderAlignment:=TTextAlign.Center;
  FHeaderVertAlignment:=TTextAlign.Center;
  FGroupHeaderSeparator:=';';
end;

procedure TMHGHeaderColumn.Assign(Source: TPersistent);
begin
  if Source is TMHGHeaderColumn then begin
    var C:=TMHGHeaderColumn(Source);
    FTitle:=C.FTitle;
    FWidth:=C.FWidth;
    FMinWidth:=C.FMinWidth;
    FMaxWidth:=C.FMaxWidth;
    FVisible:=C.FVisible;
    FWordWrap:=C.FWordWrap;
    FAlignment:=C.FAlignment;
    FVertAlignment:=C.FVertAlignment;
    FHeaderAlignment:=C.FHeaderAlignment;
    FHeaderVertAlignment:=C.FHeaderVertAlignment;
    FGroupHeader:=C.FGroupHeader;
    FGroupHeaderSeparator:=C.FGroupHeaderSeparator;
    Changed;
  end else
    inherited;
end;

function TMHGHeaderColumn.SetProps(const ATitle: string; const AGroupHeader: string;
  AWidth: Integer; AAlignment: TTextAlign; AMinWidth: Integer; AMaxWidth: Integer): TMHGHeaderColumn;
begin
  var Coll:=Collection;
  if Coll<>nil then Coll.BeginUpdate;
  try
    FTitle:=ATitle;
    FGroupHeader:=AGroupHeader;
    FAlignment:=AAlignment;
    if AWidth>=0 then FWidth:=AWidth;
    if AMinWidth>=0 then FMinWidth:=AMinWidth;
    if AMaxWidth>=0 then FMaxWidth:=AMaxWidth;
    Changed;
  finally
    if Coll<>nil then Coll.EndUpdate;
  end;
  Result:=Self;
end;

function TMHGHeaderColumn.GetDisplayName: string;
begin
  if FTitle<>'' then Result:=FTitle else Result:=inherited GetDisplayName;
end;

function TMHGHeaderColumn.GroupPath: TArray<string>;
begin
  if FGroupHeader='' then Exit(nil);
  if FGroupHeaderSeparator=#0 then Exit(TArray<string>.Create(FGroupHeader));
  Result:=FGroupHeader.Split([FGroupHeaderSeparator]);
  var NLen:=Length(Result);
  while (NLen>0) and (Result[NLen-1]='') do Dec(NLen);
  SetLength(Result,NLen);
end;

procedure TMHGHeaderColumn.Changed;
begin
  inherited Changed(False);
end;

procedure TMHGHeaderColumn.SetTitle(const Value: string);
begin if FTitle<>Value then begin FTitle:=Value; Changed; end; end;
procedure TMHGHeaderColumn.SetWidth(const Value: Integer);
begin if FWidth<>Value then begin FWidth:=Value; Changed; end; end;
procedure TMHGHeaderColumn.SetMinWidth(const Value: Integer);
begin if FMinWidth<>Value then begin FMinWidth:=Value; Changed; end; end;
procedure TMHGHeaderColumn.SetMaxWidth(const Value: Integer);
begin if FMaxWidth<>Value then begin FMaxWidth:=Value; Changed; end; end;
procedure TMHGHeaderColumn.SetVisible(const Value: Boolean);
begin if FVisible<>Value then begin FVisible:=Value; Changed; end; end;
procedure TMHGHeaderColumn.SetWordWrap(const Value: Boolean);
begin if FWordWrap<>Value then begin FWordWrap:=Value; Changed; end; end;
procedure TMHGHeaderColumn.SetAlignment(const Value: TTextAlign);
begin if FAlignment<>Value then begin FAlignment:=Value; Changed; end; end;
procedure TMHGHeaderColumn.SetVertAlignment(const Value: TTextAlign);
begin if FVertAlignment<>Value then begin FVertAlignment:=Value; Changed; end; end;
procedure TMHGHeaderColumn.SetHeaderAlignment(const Value: TTextAlign);
begin if FHeaderAlignment<>Value then begin FHeaderAlignment:=Value; Changed; end; end;
procedure TMHGHeaderColumn.SetHeaderVertAlignment(const Value: TTextAlign);
begin if FHeaderVertAlignment<>Value then begin FHeaderVertAlignment:=Value; Changed; end; end;
procedure TMHGHeaderColumn.SetGroupHeader(const Value: string);
begin if FGroupHeader<>Value then begin FGroupHeader:=Value; Changed; end; end;
procedure TMHGHeaderColumn.SetGroupHeaderSeparator(const Value: Char);
begin if FGroupHeaderSeparator<>Value then begin FGroupHeaderSeparator:=Value; Changed; end; end;

{ TMHGHeaderColumns }

constructor TMHGHeaderColumns.Create(AGrid: TMultiHeaderGrid;
  AItemClass: TCollectionItemClass = nil);
begin
  if AItemClass=nil then AItemClass:=TMHGHeaderColumn;
  inherited Create(AGrid, AItemClass);
  FGrid:=AGrid;
end;

function TMHGHeaderColumns.GetItem(Index: Integer): TMHGHeaderColumn;
begin
  Result:=TMHGHeaderColumn(inherited Items[Index]);
end;

procedure TMHGHeaderColumns.SetItem(Index: Integer; const Value: TMHGHeaderColumn);
begin
  inherited Items[Index]:=Value;
end;

procedure TMHGHeaderColumns.Update(Item: TCollectionItem);
begin
  inherited;
  if FGrid<>nil then FGrid.RebuildFromHeaderColumns;
end;

function TMHGHeaderColumns.Add: TMHGHeaderColumn;
begin
  Result:=TMHGHeaderColumn(inherited Add);
end;

function TMHGHeaderColumns.AddColumn(const ATitle, AGroupHeader: string): TMHGHeaderColumn;
begin
  BeginUpdate;
  try
    Result:=Add;
    Result.Title:=ATitle;
    Result.GroupHeader:=AGroupHeader;
  finally
    EndUpdate;
  end;
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
  FHeaderWordWrap:=True;
  FHeaderColumns:=TMHGHeaderColumns.Create(Self);
  FGridLineColor:=TAlphaColors.Gray;
  FHeaderLineColor:=TAlphaColors.Black;
  FGridLineWidth:=1;
  FSelectedCell:=Point(-1, -1);

  Margins.Rect:=RectF(4,4,4,4);

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

  // Initialize column widths and row heights
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
    Header.AddRow.FillRow(IfThen(Name='',ClassName,Name));
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
  FHeaderColumns.Free;
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
    // A max constraint only shrinks an over-wide column; columns already
    // within range are left untouched.
    if FColData[Index].Widths>FColData[Index].MaxWidth then
      FColData[Index].Widths:=FColData[Index].MaxWidth;
    Invalidate;
  end;
end;

procedure TMultiHeaderGrid.SetColMinWidth(Index: Integer; const Value: integer);
begin
  if (Index>=0) and (Index<FColCount) then begin
    FColData[Index].MinWidth:=Min(Value,FColData[Index].MaxWidth);
    // A min constraint only grows an under-wide column; columns already
    // within range are left untouched.
    if FColData[Index].Widths<FColData[Index].MinWidth then
      FColData[Index].Widths:=FColData[Index].MinWidth;
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

function TMultiHeaderGrid.GridHaveWordWrap: boolean;
begin
  if FWordWrap or FGridCellsHasWordWrap then Exit(True);

  for var ColData in FColData do begin
    if ColData.WordWrap then Exit(True)
  end;

  Result:=False;
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
  // Check the range is valid
  if (ACol<0) or (ARow<0) or (ACol>=FColCount) or (ARow>=FRowCount) then
    Exit(Default(TRectF));

  X:=-ViewLeft;
  Y:=HeaderHeight-ViewTop;

  if IsMergedCell(ACol,ARow,MergedCell) then begin
    // Adjust the rectangle for a merged cell

    Result.Left:=X+GetColLeft(MergedCell.Col);
    Result.Top:=Y+FRowData[MergedCell.Row].Top;
    Result.Right:=Result.Left;
    Result.Bottom:=Result.Top;
    for i:=0 to MergedCell.ColSpan-1 do
      Result.Right:=Result.Right+GetColWidth(MergedCell.Col+i);
    for i:=0 to MergedCell.RowSpan-1 do
      Result.Bottom:=Result.Bottom+GetRowHeight(MergedCell.Row+i);
  end else begin
    // Direct calculation

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
  // Check the range is valid
  if (ACol<0) or (ARow<0) or (ACol>=FColCount) or (ARow>=FRowCount) then
    Exit(Default(TRectF));

  X:=FGridLineWidth/4-ViewLeft;
  Y:=-FGridLineWidth/4+HeaderHeight-ViewTop;

  // Calculate the column position
  for i:=0 to ACol-1 do
    X:=X+GetColWidth(i);

  // Calculate the row position
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

  // Add the header height
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

  // Calculate the header level position
  for i:=0 to ALevel-1 do begin
    Y:=Y+FHeaderLevels[i].Height;
  end;

  // Calculate the column position
  for i:=0 to ACol-1 do
    X:=X+GetColWidth(i);

  Element:=FHeaderLevels.GetElementAtCell(ACol,ALevel);

  if not Assigned(Element) then Exit(RectF(-1,-1,-1,-1));
  if FHeaderLevels.IndexOf(Element.FLevel)<>ALevel then Exit;

  // Determine the effective ColSpan (no more than the remaining columns)
  ActualColSpan:=Element.ColSpan;
  if ActualColSpan<0 then begin
    ActualColSpan:=FColCount-ACol;
  end;
  if ACol+ActualColSpan>FColCount then begin
    ActualColSpan:=FColCount-ACol;
  end;

  // Calculate the header cell width
  Result:=RectF(X, Y, X, Y+ FHeaderLevels[ALevel].Height);
  for i:=0 to ActualColSpan-1 do
    Result.Right:=Result.Right+GetColWidth(ACol+i);

  // Determine the effective RowSpan (no more than the remaining rows)
  ActualRowSpan:=Element.RowSpan;
  if ALevel+ActualRowSpan>FHeaderLevels.Count then
    ActualRowSpan:=FHeaderLevels.Count-ALevel;

  // Calculate the header cell height
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

  if GridHaveWordWrap then begin
    AutoSizeVisibleRows;
  end;

  if Canvas.BeginScene then begin
    var Save:=Canvas.SaveState;
    try
      // Set the clipping region
      var R:=TRectF.Create(LocalRect.Left,LocalRect.Top,
                           LocalRect.Left+LocalRect.Width,LocalRect.Top+LocalRect.Height);
      if VScrollBar.Visible then R.Right:=R.Right-VScrollBar.Width;
      if HScrollPanel.Visible then R.Bottom:=R.Bottom-HScrollPanel.Height;
      Canvas.IntersectClipRect(R);

      Canvas.Fill.Kind:=TBrushKind.Solid;
      Canvas.Fill.Color:=FBackgroundColor;

      // Clear the background
      Canvas.FillRect(LocalRect, 0, 0, AllCorners, 1, Canvas.Fill);

      // Draw the cells
      DrawCells(Canvas);

      // Draw the grid lines
      if FGridLines then
        DrawGridLines(Canvas);

      // Draw the headers
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
      // Blank fillers are not drawn here; the title cell beneath the
      // stack draws the whole merged region in one piece (see
      // DrawHeaderCell / HeaderCellIsFiller). This removes the internal
      // hairlines and lets the title text centre over the full height.
      if not HeaderCellIsFiller(i, j+Element.ColSkip) then
        DrawHeaderCell(Canvas, i, j+Element.ColSkip);
      j:=j+Element.ColSpan;
    end;
  end;
end;


function TMultiHeaderGrid.HeaderCellIsFiller(ALevel, ACol: Integer): Boolean;
// A "filler" is an auto-generated, empty-caption group cell that sits
// above a column whose grouping is shallower than the deepest group
// path. Fillers exist only to keep every header row fully tiled (the
// header model has no horizontal-gap support); visually they belong to
// the title cell beneath them. The title row itself is never a filler.
begin
  if (ALevel<0) or (ALevel>=FHeaderLevels.Count-1) then Exit(False);
  var Element:=FHeaderLevels.GetElementAtCell(ACol,ALevel);
  Result:=Assigned(Element) and (Element.Caption='') and (Element.ColSpan=1) and
          (FHeaderLevels.IndexOf(Element.FLevel)=ALevel);
end;

function TMultiHeaderGrid.HeaderMergedRect(ALevel, ACol: Integer): TRectF;
// Rect of the header cell at (ALevel,ACol), extended upward to swallow
// any stack of blank fillers directly above it, so a short column's
// title is drawn (and its text centred) over the full combined height.
begin
  Result:=GetHeaderRect(ALevel, ACol);
  var L:=ALevel-1;
  while (L>=0) and HeaderCellIsFiller(L, ACol) do begin
    var R:=GetHeaderRect(L, ACol);
    if R.Top<Result.Top then Result.Top:=R.Top;
    Dec(L);
  end;
end;

procedure TMultiHeaderGrid.DrawHeaderCell(Canvas: TCanvas; ALevel, ACol: Integer);
var
  Element : THeaderElement;
begin
  Element:=FHeaderLevels.GetElementAtCell(ACol,ALevel);
  if not Assigned(Element) then Exit;

  // Fillers are never drawn on their own - the title beneath the stack
  // paints the whole region (see DrawHeaders). If asked directly, no-op.
  if HeaderCellIsFiller(ALevel, ACol) then Exit;

  // If this (non-filler) cell has blank fillers stacked above it, grow
  // the rect upward to cover them so it renders as one merged cell.
  var ARect: TRectF;
  if (ALevel>0) and HeaderCellIsFiller(ALevel-1, ACol) then
    ARect:=HeaderMergedRect(ALevel, ACol)
  else
    ARect:=GetHeaderRect(ALevel, ACol);

  // Background fill
  if Element.Style.CellColorIsSet then begin
    // Explicitly set for the cell
    Canvas.Fill.Color:=Element.Style.CellColor;
  end else begin
    Canvas.Fill.Color:=FHeaderCellColor;
  end;
  Canvas.FillRect(ARect, 0, 0, AllCorners, 1);

  // Text
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
    VAlignment:=Element.Style.TextVAlignment;
  end;

  if Element.Style.FontColorIsSet then begin
    Canvas.Fill.Color:=Element.Style.FontColor;
  end else begin
    Canvas.Fill.Color:=FCellFontColor;
  end;

  var TextRect:=ARect;
  TextRect.Inflate(-CellPadding.Left-FGridLineWidth/2,-CellPadding.Top-FGridLineWidth/2,
                   -CellPadding.Right-FGridLineWidth/2,-CellPadding.Bottom-FGridLineWidth/2);
  Canvas.FillText(TextRect, Element.Caption, HeaderCellWordWrap(Element), 1, [],
                  HAlignment, VAlignment);

  // Borders (single rect around the whole merged region)
  Canvas.Stroke.Kind:=TBrushKind.Solid;
  Canvas.Stroke.Color:=FHeaderLineColor;
  Canvas.Stroke.Thickness:=FGridLineWidth/2;
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

      // Skip merged cells (except the first one)
      if IsMergedCell(i, j, MergedCell) then begin
        var FirstVisibleRow:=Max(MergedCell.Row,TopRow);

        if (i=MergedCell.Col) and (j=FirstVisibleRow) then begin
          Rect:=GetCellRect(MergedCell.Col, MergedCell.Row);
          // Adjust the rectangle for a merged cell
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
  WordWrapEnabled : Boolean; // Added: word-wrap flag
begin
  Handled:=False;

  // Check whether the cell is part of a merged block
  if IsMergedCell(ACol, ARow, MergedCell) then begin
    // Draw only the first cell of the merged block
    if (ACol<>MergedCell.Col) or (ARow<>MergedCell.Row) then Exit;

    // Check whether the master cell of this block is selected
    IsSelected:=((MergedCell.Col<=FSelectedCell.X) and (MergedCell.Col+MergedCell.ColSpan>FSelectedCell.X) or FRowSelect) and
                 (MergedCell.Row<=FSelectedCell.Y) and (MergedCell.Row+MergedCell.RowSpan>FSelectedCell.Y);

    // Selected cell
    if IsSelected then
      Canvas.Fill.Color:=TAlphaColors.Lightblue;

    // Adjust the rectangle for a merged cell
    ARect.Right:=ARect.Left;
    ARect.Bottom:=ARect.Top;
    for var i:=0 to MergedCell.ColSpan-1 do
      ARect.Right:=ARect.Right+GetColWidth(ACol+i);
    for var i:=0 to MergedCell.RowSpan-1 do
      ARect.Bottom:=ARect.Bottom+GetRowHeight(ARow+i);

    // Text
    Text:=Cells[MergedCell.Col,MergedCell.Row];
  end else begin
    // Regular cell
    IsSelected:=((ACol=FSelectedCell.X) or FRowSelect) and (ARow=FSelectedCell.Y);

    Canvas.FillRect(ARect, 0, 0, AllCorners, 1);

    // Text
    Text:=Cells[ACol, ARow];
  end;

  DoDrawCell(ACol, ARow, Canvas, ARect, IsSelected, Text, Handled);

  // Built-in checkbox glyph for boolean (toggle) cells, unless a user
  // OnDrawCell handler already drew the cell.
  if not Handled then begin
    var MC2: TMergedCell;
    if (not IsMergedCell(ACol, ARow, MC2)) and CellIsToggle(ACol, ARow) then begin
      DrawToggleCell(Canvas, ACol, ARow, ARect, IsSelected,
                     SameText(Text, 'True') or (Text='1'));
      Handled:=True;
    end;
  end;

  if not Handled then begin
    var CellStyle:=CellStyle[ACol, ARow];

    // Determine whether word wrap should be used
    // Added: word-wrap detection logic
    if CellStyle.WordWrapIsSet then begin
      WordWrapEnabled:=CellStyle.WordWrap;
    end else if (ACol>=0) and (ACol<FColCount) then begin
      WordWrapEnabled:=FWordWrap or FColData[ACol].WordWrap;
    end else begin
      WordWrapEnabled:=FWordWrap;
    end;

    // Fill
    if IsSelected then begin
      // Selected cell
      if CellStyle.SelectedCellColorIsSet then begin
        Canvas.Fill.Color:=CellStyle.SelectedCellColor;
      end else begin
        Canvas.Fill.Color:=FSelectedCellColor;
      end;
    end else begin
      // Regular cell
      if CellStyle.CellColorIsSet then begin
        // Explicitly set for the cell
        Canvas.Fill.Color:=CellStyle.CellColor;
      end else begin
        // Alternating row colors
        if (ARow mod 2=1) and (FCellColorAlternate<>TAlphaColors.Null) then begin
          Canvas.Fill.Color:=FCellColorAlternate;
        end else begin
          Canvas.Fill.Color:=FCellColor;
        end;
      end;
    end;
    Canvas.FillRect(ARect, 0, 0, AllCorners, 1);

    // Text
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
      VAlignment:=CellStyle.TextVAlignment; 
    end;

    if IsSelected then begin
      // Selected cell
      if CellStyle.SelectedFontColorIsSet then begin
        Canvas.Fill.Color:=CellStyle.SelectedFontColor;
      end else begin
        Canvas.Fill.Color:=FCellFontColor;
      end;
    end else begin
      // Regular cell
      if CellStyle.FontColorIsSet then begin
        Canvas.Fill.Color:=CellStyle.FontColor;
      end else begin
        Canvas.Fill.Color:=FCellFontColor;
      end;
    end;

    // Adjust the text output area
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

    // Changed: added the WordWrapEnabled parameter
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

  // Calculate the starting Y coordinate for the data (after the headers)
  StartX:=FGridLineWidth/4-ViewLeft;
  StartY:=-FGridLineWidth/4+HeaderHeight-ViewTop;

  var ViewBottomCell:=ViewBottom;

  var TopRow:=RowAtHeightCoord(ViewTop);
  if TopRow<0 then Exit;

  // Draw the data lines (as before)
  // First draw all vertical lines
  X:=StartX;
  for i:=0 to FColCount do begin
    // Inner vertical lines
    for j:=TopRow to FRowCount-1 do begin
      Y:=StartY+FRowData[j].Top;
      if FRowData[j].Top>ViewBottomCell then Break;

      // Check whether the line falls inside a merged cell
      var ShouldDraw:=True;

      // Check the cell to the left
      if IsMergedCell(i-1, j, MergedCell) then begin
        if i<MergedCell.Col+MergedCell.ColSpan then
          ShouldDraw:=False;
      end;

      // Check the cell to the right
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

  // Then draw all horizontal lines
  for i:=TopRow to FRowCount do begin
    if i=0 then begin
      Y:=StartY+FRowData[i].Top;
    end else begin
      Y:=StartY+FRowData[i-1].Top+FRowData[i-1].Height;
    end;

    if (i>0) and (FRowData[i-1].Top>ViewBottomCell) then Break;

    X:=StartX;
    for j:=0 to FColCount-1 do begin
      // Check whether the line falls inside a merged cell
      var ShouldDraw:=True;
      // Inner horizontal lines

      // Check the cell above
      if IsMergedCell(j, i-1, MergedCell) then begin
        if i<MergedCell.Row+MergedCell.RowSpan then begin
          ShouldDraw:=False;
        end;
      end;

      // Check the cell below
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
  // Check the range is valid
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

function TMultiHeaderGrid.ActiveEditorControl: TControl;
begin
  if FActiveEditor<>nil then
    Result:=FActiveEditor
  else
    Result:=FEditor;
end;

procedure TMultiHeaderGrid.AutoSize(ForcePrecise: boolean = False);
begin
  AutoSizeCols(ForcePrecise); // also re-fits header heights to the new widths
  AutoSizeRows(ForcePrecise);
  UpdateSize;
end;

procedure TMultiHeaderGrid.SetHeaderWordWrap(const Value: Boolean);
begin
  if FHeaderWordWrap<>Value then begin
    FHeaderWordWrap:=Value;
    Invalidate;
  end;
end;

procedure TMultiHeaderGrid.SetHeaderColumns(const Value: TMHGHeaderColumns);
begin
  FHeaderColumns.Assign(Value);
end;

procedure TMultiHeaderGrid.RebuildFromHeaderColumns;
// Re-applies the HeaderColumns collection to the grid. When empty the
// procedural header (Header.AddRow...) is left untouched.
begin
  if FRebuildingHeaderColumns then Exit;
  if FHeaderColumns=nil then Exit;
  if FHeaderColumns.Count=0 then Exit;

  FRebuildingHeaderColumns:=True;
  try
    // Gather visible items in order.
    var Vis: array of TMHGHeaderColumn;
    SetLength(Vis,FHeaderColumns.Count);
    var Cnt:=0;
    for var i:=0 to FHeaderColumns.Count-1 do
      if FHeaderColumns[i].Visible then begin
        Vis[Cnt]:=FHeaderColumns[i];
        Inc(Cnt);
      end;
    SetLength(Vis,Cnt);

    BuildHeaderFromColumns(Vis);
  finally
    FRebuildingHeaderColumns:=False;
  end;
end;

procedure TMultiHeaderGrid.BuildHeaderFromColumns(const ACols: array of TMHGHeaderColumn);
// Builds a (possibly grouped, multi-row) header from a flat list of
// header columns and applies their geometry / alignment / word wrap.
// This is the shared grouping algorithm used by both the non-data-bound
// grids (HeaderColumns) and, conceptually, the DB grid. Every group row
// is fully tiled (one element per column position) so the header model's
// left-to-right accumulation stays perfectly column-aligned.
var
  Paths: TArray<TArray<string>>;
begin
  var N:=Length(ACols);
  if N=0 then Exit;

  ColCount:=N;
  Header.Clear;

  SetLength(Paths,N);
  var MaxDepth:=0;
  for var i:=0 to N-1 do begin
    Paths[i]:=ACols[i].GroupPath;
    if Length(Paths[i])>MaxDepth then MaxDepth:=Length(Paths[i]);
  end;

  for var Level:=0 to MaxDepth-1 do begin
    var Row:=Header.AddRow(30);
    var Col:=0;
    while Col<N do begin
      if Length(Paths[Col])<=Level then begin
        Row.AddColumn('',1);
        Inc(Col);
        Continue;
      end;
      var Span:=1;
      while (Col+Span<N) and (Length(Paths[Col+Span])>Level) do begin
        var Same:=True;
        for var L:=0 to Level do
          if Paths[Col+Span][L]<>Paths[Col][L] then begin Same:=False; Break; end;
        if not Same then Break;
        Inc(Span);
      end;
      Row.AddColumn(Paths[Col][Level],Span);
      Inc(Col,Span);
    end;
  end;

  // Title row.
  var TitleRow:=Header.AddRow(30);
  for var i:=0 to N-1 do begin
    var El:=TitleRow.AddColumn(ACols[i].Title);
    El.Style.TextHAlignment:=ACols[i].HeaderAlignment;
    El.Style.TextHAlignmentIsSet:=True;
    El.Style.TextVAlignment:=ACols[i].HeaderVertAlignment;
    El.Style.TextVAlignmentIsSet:=True;
  end;

  // Apply per-column geometry/alignment/word wrap.
  for var i:=0 to N-1 do begin
    if ACols[i].Width>0 then ColWidths[i]:=ACols[i].Width;
    if ACols[i].MinWidth>0 then ColMinWidth[i]:=ACols[i].MinWidth;
    if ACols[i].MaxWidth>0 then ColMaxWidth[i]:=ACols[i].MaxWidth;
    ColWordWrap[i]:=ACols[i].WordWrap;
    ColTextHAlignment[i]:=ACols[i].Alignment;
    ColTextVAlignment[i]:=ACols[i].VertAlignment;
  end;

  Invalidate;
end;

function TMultiHeaderGrid.HeaderCellWordWrap(AElement: THeaderElement): Boolean;
// Per-element override (Style.WordWrap) wins; otherwise the grid default.
begin
  if not Assigned(AElement) then Exit(FHeaderWordWrap);
  if AElement.Style.WordWrapIsSet then
    Result:=AElement.Style.WordWrap
  else
    Result:=FHeaderWordWrap;
end;

function TMultiHeaderGrid.HeaderElementTextWidth(ALevel, ACol: Integer): Single;
// Width available for an element's text: the merged rect (covering any
// ColSpan and any blank-filler stack) minus horizontal padding.
begin
  var R:=HeaderMergedRect(ALevel, ACol);
  Result:=R.Width-CellPadding.Left-CellPadding.Right-FGridLineWidth;
  if Result<1 then Result:=1;
end;

procedure TMultiHeaderGrid.AutoSizeHeaders;
// Sizes every header level's height to fit the tallest caption on that
// level, honouring explicit #13#10 line breaks and (when enabled) word
// wrapping against the element's current column width. RowSpan/filler
// stacks are handled by distributing a tall caption's required height
// across the levels it covers.
var
  i, j: Integer;
begin
  if FHeaderLevels.Count=0 then Exit;

  Canvas.Font.Assign(FCellFont);
  var BaseTH:=Canvas.TextHeight('A');
  var CellPaddingHeight:=CellPadding.Top+CellPadding.Bottom;
  var CellDelimterHeight:=FGridLineWidth;

  // Pass 1: minimum height each level needs from its content. Levels start at
  // just padding (NOT a full text line) - a level that ends up holding only
  // blank filler cells must not reserve a line of text height. Real captions,
  // and titles distributing through the filler stack above them, raise the
  // levels they cover below.
  var MinLevel:=CellPaddingHeight+CellDelimterHeight/2;
  var LevelHeight: TArray<Single>;
  SetLength(LevelHeight,FHeaderLevels.Count);
  for i:=0 to FHeaderLevels.Count-1 do
    LevelHeight[i]:=MinLevel;

  for i:=0 to FHeaderLevels.Count-1 do begin
    var Col:=0;
    for j:=0 to FHeaderLevels[i].Count-1 do begin
      var Element:=FHeaderLevels[i][j];
      var DrawCol:=Col+Element.ColSkip;

      // Blank filler cells carry no text; they take their height from the
      // title distributing up into them (below) or from a real caption sharing
      // the level. Skip them so they reserve no line of their own.
      if (Element.Caption='') or HeaderCellIsFiller(i, DrawCol) then begin
        Col:=Col+Element.ColSpan;
        Continue;
      end;

      Canvas.Font.Assign(FCellFont);
      if Element.Style.FontNameIsSet then Canvas.Font.Family:=Element.Style.FontName;
      if Element.Style.FontSizeIsSet then Canvas.Font.Size:=Element.Style.FontSize;
      if Element.Style.FontStyleIsSet then Canvas.Font.Style:=Element.Style.FontStyle;

      var NeededText: Single;
      if HeaderCellWordWrap(Element) and (Element.Caption<>'') then begin
        var W:=HeaderElementTextWidth(i, DrawCol);
        var MeasureRect:=TRectF.Create(0,0,W,100000);
        Canvas.MeasureText(MeasureRect,Element.Caption,True,[],
                           TTextAlign.Leading,TTextAlign.Leading);
        // Convert the measured (possibly leading-padded) height into a whole
        // number of text lines, then size by BaseTH. This keeps the wrapped
        // branch as tight as the explicit-#13#10 branch and avoids the bottom
        // title row ballooning from MeasureText's internal line spacing.
        var Lines:=Max(1,Ceil(MeasureRect.Height/Max(BaseTH,1)));
        NeededText:=Lines*BaseTH;
      end else begin
        NeededText:=Max(CountLines(Element.Caption),1)*BaseTH;
      end;

      var Needed:=NeededText+CellPaddingHeight+CellDelimterHeight/2;

      // The cell visually occupies its declared RowSpan AND any blank-filler
      // levels stacked directly above it (a short column's title is drawn over
      // the combined filler+title rect - see HeaderMergedRect). Count those
      // fillers so a tall/wrapped title spreads its height across the whole
      // stack instead of forcing its own single level to grow (which made the
      // header rows too tall and left the filler levels unaccounted for).
      var TopLevel:=i;
      while (TopLevel>0) and HeaderCellIsFiller(TopLevel-1, DrawCol) do
        Dec(TopLevel);

      var Span:=Element.RowSpan;
      if Span<1 then Span:=1;
      var BottomLevel:=Min(i+Span-1,FHeaderLevels.Count-1);

      // Spread the requirement evenly across every level the cell covers
      // (filler levels above + its own + any RowSpan below). Each level is
      // bumped to at least its share, so the levels' SUM (which is how the
      // merged height is drawn) covers Needed.
      var CoveredLevels:=BottomLevel-TopLevel+1;
      if CoveredLevels<1 then CoveredLevels:=1;
      var Share:=Needed/CoveredLevels;
      for var L:=TopLevel to BottomLevel do
        LevelHeight[L]:=Max(LevelHeight[L],Share);

      Col:=Col+Element.ColSpan;
    end;
  end;

  for i:=0 to FHeaderLevels.Count-1 do
    FHeaderLevels[i].Height:=Trunc(LevelHeight[i]);

  Invalidate;
end;

function CountLines(const Text: string): Integer;
var
  PosStart, PosFound: Integer;
begin
  if Text='' then begin
    Result:=0;
    Exit;
  end;

  Result:=1; // At least one row
  PosStart:=1;

  while True do begin
    PosFound:=Pos(#13#10, Text, PosStart);
    if PosFound=0 then Break;

    Inc(Result);
    PosStart:=PosFound+2; // Skip #13#10
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
    // Find the end of the line
    EndPos:=Pos(#13#10, Text, StartPos);

    if EndPos=0 then begin
      // This is the last line
      LineLength:=Length(Text)-StartPos+1;
      if LineLength>Result then
        Result:=LineLength;
      Break;
    end else begin
      // Calculate the length of the current line
      LineLength:=EndPos-StartPos;
      if LineLength>Result then
        Result:=LineLength;

      // Move to the next line
      StartPos:=EndPos+1;

      // Skip a possible #10 after #13
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
    // Find the end of the line
    EndPos:=Pos(#13#10, Text, StartPos);

    if EndPos=0 then begin
      // This is the last line
      LineLength:=Length(Text)-StartPos+1;
      if LineLength>Length(Result) then
        Result:=Copy(Text,StartPos,LineLength);
      Break;
    end else begin
      // Calculate the length of the current line
      LineLength:=EndPos-StartPos;
      if LineLength>Length(Result) then
        Result:=Copy(Text,StartPos,LineLength);

      // Move to the next line
      StartPos:=EndPos+1;

      // Skip a possible #10 after #13
      if (EndPos<=Length(Text)) and (Text[EndPos]=#13) and
         (StartPos<=Length(Text)) and (Text[StartPos]=#10) then
        Inc(StartPos);
    end;
  end;
end;

procedure TMultiHeaderGrid.AutoSizeCols(ForcePrecise:boolean=False);
var
  i,j:Integer;
  Text:string;
begin
  var CellPaddingWidth:=CellPadding.Left+CellPadding.Right;
  var CellDelimterWidth:=FGridLineWidth/2;
  var CellPaddingFull:=CellPaddingWidth+CellDelimterWidth+1;

  // Upper bound on how wide a SINGLE unbreakable word may push a column, so
  // one pathological token (a long URL/id with no spaces) cannot blow the
  // column out. Past the cap we accept a mid-word break as the lesser evil.
  var VP:=ViewPortWidth;
  if VP<=0 then VP:=600;
  var WordWidthCap:Single:=Max(200, VP*0.6);

  // Optimization: if there's no WordWrap, use the fast path
  var UseFastMode:=not ForcePrecise and (FRowCount*FColCount>10000);

  for i:=0 to FColCount-1 do begin
    // HeaderFullW  - widest header caption laid out on a SINGLE line.
    // HeaderWordW  - widest single word (the tightest a wrapping header can be).
    // DataW        - widest cell content (honouring cell word-wrap as before).
    // CanWrapHdr   - any header element over this column allows word wrap.
    var HeaderFullW:Single:=0;
    var HeaderWordW:Single:=0;
    var DataW:Single:=0;
    var CanWrapHdr:Boolean:=False;

    // Check the headers
    for j:=0 to FHeaderLevels.Count-1 do begin
      var Element:=FHeaderLevels.GetElementAtCell(i,j);
      if not Assigned(Element) then Continue;

      Canvas.Font.Assign(FCellFont);
      if Element.Style.FontNameIsSet then Canvas.Font.Family:=Element.Style.FontName;
      if Element.Style.FontSizeIsSet then Canvas.Font.Size:=Element.Style.FontSize;
      if Element.Style.FontStyleIsSet then Canvas.Font.Style:=Element.Style.FontStyle;

      if HeaderCellWordWrap(Element) then CanWrapHdr:=True;

      // Full one-line width (per spanned column).
      var Lines:=Element.Caption.Split([#13#10]);
      for var Line in Lines do
        HeaderFullW:=Max(HeaderFullW,
          Canvas.TextWidth(Line)/Element.FColSpan-(Element.FColSpan-1)*CellDelimterWidth);

      // Widest single unbreakable token - the hard floor for a wrapped
      // column: a whole word must fit within ONE column's width, so it is NOT
      // divided by the ColSpan (dividing would let a word split across the
      // span). Only true whitespace/line-breaks delimit tokens, so a word is
      // never broken mid-characters.
      var Words:=Element.Caption.Split([' ', #9, #13, #10]);
      for var Word in Words do
        if Word<>'' then
          HeaderWordW:=Max(HeaderWordW,Canvas.TextWidth(Word));
    end;

    // Optimization: if this column has no WordWrap, use the original fast logic
    var ColHasWordWrap:=FWordWrap or FColData[i].WordWrap or FGridCellsHasWordWrap;

    if not ColHasWordWrap and UseFastMode then begin
      // Fast path
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
          if CellStyle.FontNameIsSet then Canvas.Font.Family:=CellStyle.FontName;
          if CellStyle.FontSizeIsSet then Canvas.Font.Size:=CellStyle.FontSize;
          if CellStyle.FontStyleIsSet then Canvas.Font.Style:=CellStyle.FontStyle;

          DataW:=Max(DataW,Canvas.TextWidth(Line)/ColSpan-(ColSpan-1)*CellDelimterWidth);
        end;
      end;
    end else begin
      // Full path (with WordWrap check)
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

        // Determine whether word wrap is enabled for this cell
        var WordWrapEnabled: Boolean;
        if CellStyle.WordWrapIsSet then
          WordWrapEnabled:=CellStyle.WordWrap
        else
          WordWrapEnabled:=FWordWrap or FColData[i].WordWrap;

        Canvas.Font.Assign(FCellFont);
        if CellStyle.FontNameIsSet then Canvas.Font.Family:=CellStyle.FontName;
        if CellStyle.FontSizeIsSet then Canvas.Font.Size:=CellStyle.FontSize;
        if CellStyle.FontStyleIsSet then Canvas.Font.Style:=CellStyle.FontStyle;

        if WordWrapEnabled then begin
          // Word-wrap sizing: normally fit just the widest word and let the
          // rest wrap. But do NOT wrap a cell whose whole content already fits
          // in a small width - wrapping a short value onto several lines looks
          // bad. Measure the full one-line width; only fall back to the
          // widest-word width when the content is genuinely wide (> 120px).
          const DataWrapGateW = 120;
          var FullLineW:=0.0;
          var Lines:=Text.Split([#13#10]);
          for var Line in Lines do
            FullLineW:=Max(FullLineW,Canvas.TextWidth(Line));

          if FullLineW>DataWrapGateW then begin
            // Wide content: size to the widest single word, let it wrap.
            var Words:=Text.Split([' ', ':', ';', ',', '.', '!', '?', '-', '+', '*', '/', '\', '|', #9]);
            var MaxWordWidth:=0.0;
            for var Word in Words do
              if Word<>'' then
                MaxWordWidth:=Max(MaxWordWidth,Canvas.TextWidth(Word));
            DataW:=Max(DataW,MaxWordWidth/ColSpan-(ColSpan-1)*CellDelimterWidth/2);
          end else begin
            // Already small enough: keep it on one line.
            DataW:=Max(DataW,FullLineW/ColSpan-(ColSpan-1)*CellDelimterWidth/2);
          end;
        end else begin
          // Without WordWrap, measure each line
          var Lines:=Text.Split([#13#10]);
          for var Line in Lines do
            DataW:=Max(DataW,Canvas.TextWidth(Line)/ColSpan-(ColSpan-1)*CellDelimterWidth/2);
        end;
      end;
    end;

    // Decide the column width.
    //
    // Default: fit whichever of header / data is wider (header stays 1 line).
    //
    // Smart wrap: when the header is much wider than the data - header width
    // > 50px AND header exceeds data by > 20px - and the header is allowed to
    // wrap, size the column to the DATA instead and let the caption wrap onto
    // extra lines (AutoSizeHeaders then grows the height).
    var Wrapped:=CanWrapHdr and (HeaderFullW>50) and (HeaderFullW-DataW>20);

    var TargetW:Single;
    if Wrapped then
      // Wrap header to data width, but never below the widest whole word so
      // words are not split - except an unusually long word is capped at
      // WordWidthCap (it then character-breaks rather than blow the column out).
      TargetW:=Max(DataW, Min(HeaderWordW, WordWidthCap))
    else
      TargetW:=Max(HeaderFullW,DataW);

    // Add the padding
    TargetW:=TargetW+CellPaddingFull;

    var NewWidth:=Ceil(TargetW);

    // When the header wraps, AutoSizeHeaders gives the text only
    // (ColWidth - Left - Right - FGridLineWidth) - which subtracts a bit more
    // than CellPaddingFull added (full grid line vs half, no +1). A column
    // sized to exactly the widest word can then come up a pixel short and split
    // it. Guarantee the realised text width covers the widest whole word
    // (unless over the cap) with a small safety margin.
    if Wrapped then begin
      var WordFloor:=Min(HeaderWordW, WordWidthCap);
      var MinColForWord:=Ceil(WordFloor + CellPadding.Left + CellPadding.Right +
                              FGridLineWidth + 2);
      if NewWidth<MinColForWord then
        NewWidth:=MinColForWord;
    end;

    // Floor at the column's own MinWidth (or a small absolute minimum), so a
    // genuinely narrow column (e.g. a short type/flag field) can autosize
    // tight instead of being padded out to the global default width.
    var Floor:=Max(10,FColData[i].MinWidth);
    if NewWidth<Floor then
      NewWidth:=Floor;

    ColWidths[i]:=NewWidth;
  end;

  // Header heights must follow the (possibly wrap-narrowed) column widths.
  AutoSizeHeaders;
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
  if GridHaveWordWrap then begin
    ComputeMode:=TSizeComputeMode.cmFull;
  end;

  Canvas.Font.Assign(FCellFont);
  var TH:=Canvas.TextHeight('A');
  var Text:='';
  for var Row:=FromRow to FRowCount-1 do begin
    if ((ToRow<0) and (Row>0) and (FRowData[Row-1].Top>ViewBottomCell)) or
       ((ToRow>=0) and (Row>ToRow)) then Break;

    var MaxHeight:=TH;
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
          // Fast calculation

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
          // Regular calculation

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

          MaxHeight:=Max(MaxHeight,TextLines*Canvas.TextHeight('A')/RowSpan-(RowSpan-1)*CellDelimterHeight/4);
        end;

        cmFull: begin
          // Calculation accounting for WordWrap, slow.

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

          // Determine whether word wrap is enabled
          var WordWrapEnabled: Boolean;
          if CellStyle.WordWrapIsSet then begin
            WordWrapEnabled:=CellStyle.WordWrap;
          end else begin
            WordWrapEnabled:=FWordWrap or FColData[ActualCol].WordWrap;
          end;

          // Set up the font
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
            // Use MeasureText for the calculation
            var MeasureRect:=TRectF.Create(0,0,AvailableWidth,10000);
            Canvas.MeasureText(MeasureRect,Text,True,[],TTextAlign.Leading,TTextAlign.Leading);
            TextHeight:=MeasureRect.Height;
          end else begin
            // Without WordWrap, count line breaks
            var LineCount:=CountLines(Text);
            TextHeight:=Canvas.TextHeight('A')*LineCount-LineCount*FGridLineWidth;
          end;

          // Adjust the height
          MaxHeight:=Max(MaxHeight,TextHeight/RowSpan-(RowSpan-1)*CellDelimterHeight/4);
        end;
      end;
    end;

    if Row>=0 then begin
      FRowData[Row].Height:=Trunc(MaxHeight+CellPaddingHeight+CellDelimterHeight/2);
      if Row>0 then begin
        FRowData[Row].Top:=FRowData[Row-1].Top+FRowData[Row-1].Height;
      end;
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
    // Set focus on click
    SetFocus;
    Invalidate;
  end;

  if Button=TMouseButton.mbLeft then begin
    // Check for a click on the headers
    if FResizeEnabled then begin
      var ResizeMode:=IsResizeArea(X, Y, StartCol, EndCol, Row);

      case ResizeMode of
        TResizeMode.rmColumn: StartColumnResize(StartCol, EndCol, X);
        TResizeMode.rmHeaderRow: StartHeaderRowResize(Row, Y);
        TResizeMode.rmGridRow: StartGridRowResize(Row, Y);
      end;

      if ResizeMode<>TResizeMode.rmNone then begin
        // Capture the mouse to continue dragging outside the control
        Capture;
        Exit;
      end;
    end;

    // Check for a click on the headers
    for i:=0 to FHeaderLevels.Count-1 do begin
      j:=0;
      while j<FColCount do begin
        Rect:=GetHeaderRect(i, j);
        if Rect.Contains(PointF(X, Y)) then begin
          // Check whether this is a merged cell
          DoHeaderClick(i, j);
          Exit;
        end;
        // With grouped headers a header cell can be empty (no element
        // covers it - e.g. the blank band above an ungrouped column),
        // so GetElementAtCell may return nil. Advance one column then.
        var Element:=FHeaderLevels.GetElementAtCell(j,i);
        if Element=nil then begin
          Inc(j);
          Continue;
        end;
        var ColSpan:=Element.ColSpan;
        if ColSpan<0 then ColSpan:=FColCount;
        if ColSpan<1 then ColSpan:=1; // never advance by 0 -> no infinite loop
        j:=j+ColSpan;
      end;
    end;

    // Check the data cells
    var TopRow:=RowAtHeightCoord(ViewTop);
    if TopRow<0 then Exit;
    var ViewBottom:=ViewTop+LocalRect.Top+LocalRect.Height+Margins.Bottom-HeaderHeight;

    for j:=TopRow to FRowCount-1 do begin
      if FRowData[j].Top>ViewBottom then Break;
      for i:=0 to FColCount-1 do begin
        Rect:=GetCellRect(i, j);
        if Rect.Contains(PointF(X, Y)) then begin
          FLastClickIsOnCell:=True;
          // Remember whether the click landed on the checkbox glyph BEFORE
          // selection changes the row. Clicking on the box toggles; clicking
          // anywhere else in the cell only selects.
          var HitToggle:=CellIsToggle(i, j) and
                         ToggleGlyphRect(i, j, Rect).Contains(PointF(X, Y));

          FSelectedCell:=Point(i, j);
          DoCellClick(FSelectedCell.X, FSelectedCell.Y);
          ScrollToSelectedCell;
          // Move the dataset cursor to the clicked row first, so the toggle
          // acts on the correct record (and the right row gets selected).
          DoSelectCell;

          if HitToggle then
            ToggleCell(i, j);
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
          // Use global coordinates for accurate tracking
          GlobalPos:=LocalToAbsolute(PointF(X, Y));
          NewWidth:=ResizeStartWidth+Round(GlobalPos.X-FResizeStartPos.X);
          UpdateColumnWidth(FResizeStartColumnIndex,FResizeEndColumnIndex, NewWidth);
          Cursor:=crHSplit;
        end;
      end;
      TResizeMode.rmHeaderRow: begin
        if FResizeRowIndex>=0 then begin
          // Use global coordinates for accurate tracking
          GlobalPos:=LocalToAbsolute(PointF(X, Y));
          NewHeight:=FResizeStartHeight+Round(GlobalPos.Y-FResizeStartPos.Y);
          UpdateHeaderRowHeight(FResizeRowIndex, NewHeight);
          Cursor:=crVSplit;
        end;
      end;
      TResizeMode.rmGridRow: begin
        if FResizeRowIndex>=0 then begin
          // Use global coordinates for accurate tracking
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
    // Finish the column resize
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

procedure TMultiHeaderGrid.MouseWheel(Shift: TShiftState; WheelDelta: Integer; var Handled: Boolean);
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
    ViewTop:=Min(ViewTop+WheelDelta,FullTableHeight-ViewPortDataHeight-HeaderHeight);
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
    raise Exception.Create('Clipboard service is not supported');
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
        if CellIsToggle(FSelectedCell.X, FSelectedCell.Y) then
          ToggleCell(FSelectedCell.X, FSelectedCell.Y)
        else if CanEditCell(FSelectedCell.X, FSelectedCell.Y) then
          StartCellEditing(FSelectedCell.X, FSelectedCell.Y, #0)
        else
          DoCellClick(FSelectedCell.X, FSelectedCell.Y);
        Exit;
      end;
      vkSpace: begin // Space toggles a boolean cell
        if CellIsToggle(FSelectedCell.X, FSelectedCell.Y) then begin
          ToggleCell(FSelectedCell.X, FSelectedCell.Y);
          KeyChar:=#0;
          Exit;
        end;
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

  // F2 edits the current cell's existing content.
  if IsFocused and (Key=vkF2) then begin
    if CellIsToggle(FSelectedCell.X, FSelectedCell.Y) then
      ToggleCell(FSelectedCell.X, FSelectedCell.Y)
    else
      StartCellEditing(FSelectedCell.X, FSelectedCell.Y, #0);
    Key:=0;
    Exit;
  end;

  // Handle text input (never enters edit mode on a toggle/boolean cell).
  // Space on a toggle cell flips it (the vkSpace case above doesn't fire on
  // platforms that deliver space only as KeyChar=' ' with Key=0).
  if IsFocused and (KeyChar>=' ') and (KeyChar<='~') and not (ssCtrl in Shift) then begin
    if CellIsToggle(FSelectedCell.X, FSelectedCell.Y) then begin
      if KeyChar=' ' then
        ToggleCell(FSelectedCell.X, FSelectedCell.Y);
    end else begin
      // A keystroke opens the editor. Editors that can use the char (text /
      // number) seed it; others (date/time) ignore it in LoadEditorValue.
      StartCellEditing(FSelectedCell.X, FSelectedCell.Y, KeyChar);
    end;
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

  // Quick edge-case check
  if Y<FRowData[0].Top then
    Exit;

  // Check the last element
  CurrentBottom:=FRowData[High(FRowData)].Top+FRowData[High(FRowData)].Height;
  if Y>=CurrentBottom then
    Exit(High(FRowData)); // or -1, depending on requirements

  // Binary search using integer arithmetic
  L:=0;
  R:=High(FRowData);

  while L<=R do begin
    M:=(L+R) shr 1; // Faster than div 2

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

function TMultiHeaderGrid.CellIsToggle(ACol, ARow: Integer): Boolean;
begin
  // The plain grid has no toggle cells; the DB grid overrides for ftBoolean.
  Result:=False;
end;

function TMultiHeaderGrid.ToggleCell(ACol, ARow: Integer): Boolean;
begin
  Result:=False;
end;

function TMultiHeaderGrid.ToggleGlyphRect(ACol, ARow: Integer; const ARect: TRectF): TRectF;
const
  BoxSize = 13; // nominal glyph size (px); clamped to the cell
begin
  // Centered square box, clamped so it always fits the cell.
  var S:=Min(BoxSize, Min(ARect.Width, ARect.Height)-2);
  if S<6 then S:=Min(ARect.Width, ARect.Height);
  Result:=TRectF.Create(0, 0, S, S);
  Result.Offset(ARect.Left+(ARect.Width-S)/2, ARect.Top+(ARect.Height-S)/2);
end;

procedure TMultiHeaderGrid.DrawToggleCell(Canvas: TCanvas; ACol, ARow: Integer;
  const ARect: TRectF; IsSelected, AChecked: Boolean);
begin
  // Cell background (matches the regular fill / selection highlight).
  if IsSelected then
    Canvas.Fill.Color:=FSelectedCellColor
  else
    Canvas.Fill.Color:=FCellColor;
  Canvas.FillRect(ARect, 0, 0, AllCorners, 1);

  var Box:=ToggleGlyphRect(ACol, ARow, ARect);
  var S:=Box.Width;

  // Selected state is made distinct: a thicker, darker (blue) box outline.
  var BorderColor: TAlphaColor;
  var BorderW: Single;
  if IsSelected then begin
    BorderColor:=TAlphaColors.Royalblue;
    BorderW:=2;
  end else begin
    BorderColor:=TAlphaColors.Gray;
    BorderW:=1;
  end;

  if AChecked then begin
    // Checked: filled box (blue when selected, sea-green otherwise) with a
    // white check mark - clearly readable at the smaller glyph size.
    if IsSelected then
      Canvas.Fill.Color:=TAlphaColors.Royalblue
    else
      Canvas.Fill.Color:=TAlphaColors.Seagreen;
    Canvas.FillRect(Box, 2, 2, AllCorners, 1);

    Canvas.Stroke.Kind:=TBrushKind.Solid;
    Canvas.Stroke.Thickness:=BorderW;
    Canvas.Stroke.Color:=BorderColor;
    Canvas.DrawRect(Box, 2, 2, AllCorners, 1);

    Canvas.Stroke.Color:=TAlphaColors.White;
    Canvas.Stroke.Thickness:=Max(1.5, S/7);
    Canvas.Stroke.Cap:=TStrokeCap.Round;
    var P1:=PointF(Box.Left+S*0.22, Box.Top+S*0.52);
    var P2:=PointF(Box.Left+S*0.42, Box.Top+S*0.72);
    var P3:=PointF(Box.Left+S*0.78, Box.Top+S*0.26);
    Canvas.DrawLine(P1, P2, 1);
    Canvas.DrawLine(P2, P3, 1);
    Canvas.Stroke.Cap:=TStrokeCap.Flat;
  end else begin
    // Unchecked: white interior + outline.
    Canvas.Fill.Color:=TAlphaColors.White;
    Canvas.FillRect(Box, 2, 2, AllCorners, 1);
    Canvas.Stroke.Kind:=TBrushKind.Solid;
    Canvas.Stroke.Thickness:=BorderW;
    Canvas.Stroke.Color:=BorderColor;
    Canvas.DrawRect(Box, 2, 2, AllCorners, 1);
  end;
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
    // Moving to another cell commits any in-progress edit on the old one.
    if FEditing and ((FEditCol<>Value.X) or (FEditRow<>Value.Y)) then
      CommitEditing;
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

function TMultiHeaderGrid.CanEditCell(ACol, ARow: Integer): Boolean;
begin
  // Base/string grids are editable unless ReadOnly. Descendants (e.g. the DB
  // grid) override to apply their own rules.
  Result:=(not FReadOnly) and
          (ACol>=0) and (ACol<FColCount) and (ARow>=0) and (ARow<FRowCount);
end;

procedure TMultiHeaderGrid.EnsureEditor(TextAlign: TTextAlign);
begin
  if FEditor<>nil then begin
    FEditor.TextSettings.HorzAlign:=TextAlign;
    Exit;
  end;

  // A clipping host sized to the cell interior. Both the backing fill and the
  // memo are parented here and clipped to it, so text/fill never bleed into
  // neighbouring cells or fight the gridlines.
  FEditorHost:=TLayout.Create(Self);
  FEditorHost.Parent:=Self;
  FEditorHost.Visible:=False;
  FEditorHost.ClipChildren:=True;
  FEditorHost.HitTest:=False;

  // Opaque backing rectangle drawn behind the borderless editor so the
  // underlying cell text never shows through.
  FEditorBack:=TRectangle.Create(Self);
  FEditorBack.Parent:=FEditorHost;
  FEditorBack.Align:=TAlignLayout.Client;
  FEditorBack.HitTest:=False;
  FEditorBack.Stroke.Kind:=TBrushKind.None;
  FEditorBack.Fill.Kind:=TBrushKind.Solid;
  FEditorBack.Fill.Color:=TAlphaColors.White;
  FEditorBack.XRadius:=0;
  FEditorBack.YRadius:=0;

  FEditor:=TMemo.Create(Self);
  FEditor.Parent:=FEditorHost;
  FEditor.Align:=TAlignLayout.Client;
  FEditor.WordWrap:=False;
  FEditor.AutoSelect:=False;
  FEditor.StyledSettings:=[];
  FEditor.ShowScrollBars:=False;
  FEditor.DisableMouseWheel:=True;
  FEditor.OnKeyDown:=EditorKeyDown;
  FEditor.OnExit:=EditorExit;
  FEditor.OnApplyStyleLookup:=ApplyStyle;
  FEditor.TextSettings.HorzAlign:=TextAlign;
end;

procedure TMultiHeaderGrid.ApplyStyle(Sender: TObject);
var
  Obj: TFmxObject;
begin
  // Hide the styled background entirely (this is what paints the
  // border). The opaque FEditorBack rectangle supplies the fill instead.
  if not Assigned(Sender) or not (Sender is TStyledControl) then Exit;

  Obj:=TStyledControl(Sender).FindStyleResource('background');
  if Obj is TControl then
    TControl(Obj).Visible:=False;
end;

procedure TMultiHeaderGrid.LoadEditorValue(ACol, ARow: Integer; InitialChar: Char);
begin
  // Base: the active editor is the TMemo.
  if InitialChar>=' ' then
    FEditor.Text:=InitialChar          // typing replaces the cell content
  else
    FEditor.Text:=Cells[ACol,ARow];    // F2 / double-click edits existing text
end;

function TMultiHeaderGrid.GetEditorText: string;
begin
  if FEditor<>nil then
    Result:=FEditor.Text
  else
    Result:='';
end;

function TMultiHeaderGrid.CommitEditorValue(ACol, ARow: Integer): Boolean;
begin
  // Base: write the editor text straight into the cell.
  var NewText:=GetEditorText;
  Result:=True;
  if (ACol>=0) and (ACol<FColCount) and (ARow>=0) and (ARow<FRowCount) then
    if Cells[ACol,ARow]<>NewText then begin
      Cells[ACol,ARow]:=NewText; // SetCells fires OnSetCellText / invalidates
      // The new content may need more (or fewer) lines, so re-fit the row
      // height. A merged cell spans several rows - autosize that whole range.
      var MergedCell: TMergedCell;
      if IsMergedCell(ACol,ARow,MergedCell) then
        AutoSizeRows(MergedCell.Row, MergedCell.Row+MergedCell.RowSpan-1, True)
      else
        AutoSizeRows(ARow, ARow, True);
      UpdateSize; // total height changed -> refresh scrollbar range
    end;
end;

procedure TMultiHeaderGrid.PositionEditor;
begin
  var Ed:=ActiveEditorControl;
  if (Ed=nil) or (not FEditing) then Exit;

  // Build the rect EXACTLY as DrawCells does, so the editor lines up with the
  // painted cell regardless of GridLineWidth. The plain GetCellRect overload
  // carries the FGridLineWidth/4 inset that the merged-cell overload omits;
  // using the same source keeps them aligned. FEditCol/FEditRow are already
  // the merge anchor (resolved in StartCellEditing).
  var R:=GetCellRect(FEditCol,FEditRow);
  var MergedCell: TMergedCell;
  if IsMergedCell(FEditCol,FEditRow,MergedCell) then begin
    R.Right:=R.Left;
    R.Bottom:=R.Top;
    for var K:=0 to MergedCell.ColSpan-1 do
      R.Right:=R.Right+GetColWidth(FEditCol+K);
    for var K:=0 to MergedCell.RowSpan-1 do
      R.Bottom:=R.Bottom+GetRowHeight(FEditRow+K);
  end;

  // Clip to the cells viewport so the editor never overlaps the header.
  var Top:=HeaderHeight;
  if R.Top<Top then R.Top:=Top;

  // Clamp the right/bottom edges to the visible cells area (before the
  // scrollbars). Without this an editor wider/taller than the viewport - e.g.
  // after a column was widened past the grid's right edge - paints outside the
  // grid and leaves rendering artefacts at the border.
  var MaxRight:=ViewPortWidth;
  var MaxBottom:=HeaderHeight+ViewCellsHeight;
  if R.Right>MaxRight then R.Right:=MaxRight;
  if R.Bottom>MaxBottom then R.Bottom:=MaxBottom;

  // Only show the editor if a positive area remains after clamping.
  if (R.Width<=0) or (R.Height<=0) then begin
    Ed.Visible:=False;
    Exit;
  end;

  Ed.SetBounds(R.Left+FGridLineWidth/4, R.Top+FGridLineWidth/4,
               R.Width-FGridLineWidth/2, R.Height-FGridLineWidth/2);


  // Match the cell font (only controls that expose TextSettings).
  if Ed is TMemo then
    TMemo(Ed).TextSettings.Font.Assign(FCellFont);

  // Hide the editor entirely if the cell has scrolled out of the visible area.
  var Vis:=(R.Bottom>HeaderHeight) and (R.Top<Height) and
           (R.Right>0) and (R.Left<Width);

  Ed.Visible:=Vis;
  if Vis then Ed.BringToFront;
  if FEditorHost<>nil then begin
    // Host fills the cell interior; backing + memo align Client and are
    // clipped to it.
    FEditorHost.SetBounds(R.Left+FGridLineWidth/4, R.Top+FGridLineWidth/4,
                          R.Width-FGridLineWidth/2, R.Height-FGridLineWidth/2);
    FEditorHost.Visible:=Vis;
    if Vis then FEditorHost.BringToFront;
  end;
end;

function TMultiHeaderGrid.PrepareCellEditor(ACol, ARow: Integer): TControl;
begin
  // Base/string grids always use the shared TMemo.

  EnsureEditor(ColTextHAlignment[ACol]);
  FActiveEditor:=FEditor;
  Result:=FEditor;
end;

procedure TMultiHeaderGrid.StartCellEditing(ACol, ARow: Integer; InitialChar: Char);
begin
  // If the cell is part of a merge, edit the merge's top-left anchor - that
  // is where the value lives and what gets drawn.
  var MergedCell: TMergedCell;
  if IsMergedCell(ACol, ARow, MergedCell) then begin
    ACol:=MergedCell.Col;
    ARow:=MergedCell.Row;
  end;

  // Let a handler veto or transform the trigger char (back-compat).
  if Assigned(FOnStartEditing) then
    FOnStartEditing(Self, ACol, ARow, InitialChar);

  if not CanEditCell(ACol,ARow) then Exit;

  // Commit any edit already in progress before starting a new one.
  if FEditing then CommitEditing;

  EnsureEditor(ColTextHAlignment[ACol]);
  FEditCol:=ACol;
  FEditRow:=ARow;
  FEditing:=True;

  // Let the (overridable) factory pick/create the editor control for this
  // cell. Base returns the shared TMemo; the DB grid may return a typed one.
  var Ed:=PrepareCellEditor(ACol, ARow);
  if Ed=nil then Exit; // no editor available for this cell

  FEditCol:=ACol;
  FEditRow:=ARow;
  FEditing:=True;

  // If the chosen editor needs more room than the (possibly merged) cell
  // currently offers, widen the anchor column before positioning so the
  // control isn't clipped in a narrow column.
  var MinW:=EditorMinColWidth(ACol, ARow);
  if MinW>0 then begin
    var SpanW:=GetColWidth(ACol);
    var Anchor:=ACol;
    var MC: TMergedCell;
    if IsMergedCell(ACol, ARow, MC) then begin
      Anchor:=MC.Col;
      SpanW:=0;
      for var K:=0 to MC.ColSpan-1 do
        SpanW:=SpanW+GetColWidth(MC.Col+K);
    end;
    if SpanW<MinW then begin
      // Add the whole deficit to the anchor column. SetColWidth clamps to the
      // column's MaxWidth, so an explicit MaxWidth still wins.
      ColWidths[Anchor]:=Trunc(GetColWidth(Anchor)+(MinW-SpanW));
      // Widening can push the cell's right edge past the viewport. Bring the
      // (now wider) cell fully back into view so the editor isn't drawn half
      // outside the grid, which leaves rendering artefacts at the border.
      FSelectedCell:=Point(FEditCol, FEditRow);
      ScrollToSelectedCell;
    end;
  end;

  LoadEditorValue(ACol, ARow, InitialChar);

  PositionEditor;
  Ed.Visible:=True;
  Ed.BringToFront;
  if Ed.CanFocus then
    Ed.SetFocus;
  // Place the caret at the end of the editor's text (no selection), both when
  // entering with the existing value and when a typed char seeded it.
  PlaceEditorCaretAtEnd(Ed);
end;

procedure TMultiHeaderGrid.PlaceEditorCaretAtEnd(Ed: TControl);
begin
  // Base grid only has the shared TMemo. Descendants override to handle their
  // own typed editors (see TMultiHeaderDBGrid).
  if Ed is TMemo then
    TMemo(Ed).GoToTextEnd;
end;

procedure TMultiHeaderGrid.CommitEditing;
begin
  if not FEditing then Exit;
  FEditing:=False; // clear first so EditorExit re-entry is a no-op
  var Col:=FEditCol;
  var Row:=FEditRow;
  if FEditorHost<>nil then FEditorHost.Visible:=False;

  var Ed:=ActiveEditorControl;
  if Ed<>nil then Ed.Visible:=False;

  // Single persistence point. The base writes the editor text into Cells[]
  // (which fires DoSetCellText); the DB grid override assigns the bound field
  // with the correct type under its own try/except guard. (Previously the
  // memo text was ALSO pushed through Cells[] here, which double-wrote and,
  // for a typed editor, forced an unparsed string into a numeric/datetime
  // field - raising on commit.)
  CommitEditorValue(Col, Row);

  FActiveEditor:=nil;
  if CanFocus then SetFocus;
end;

procedure TMultiHeaderGrid.CancelEditing;
begin
  if not FEditing then Exit;
  FEditing:=False;
  var Ed:=ActiveEditorControl;
  if Ed<>nil then Ed.Visible:=False;
  if FEditorHost<>nil then FEditorHost.Visible:=False;
  FActiveEditor:=nil;
  if CanFocus then SetFocus;
  Invalidate;
end;

procedure TMultiHeaderGrid.EditorKeyDown(Sender: TObject; var Key: Word;
  var KeyChar: Char; Shift: TShiftState);
begin
  case Key of
    vkReturn:
      begin
        // Plain Enter commits and exits the editor. Enter with any modifier
        // (Shift/Ctrl/Alt) inserts a line break instead.
        var WantNewline:=(ssShift in Shift) or (ssCtrl in Shift) or (ssAlt in Shift);
        if WantNewline then begin
          // Modifier + Enter inserts a line break at the caret. SelText is
          // read-only in FMX, so compute the absolute caret offset from
          // CaretPosition (SelStart is unreliable right after focus / text
          // changes and can collapse to 0, inserting at the start).
          var Cp:=FEditor.CaretPosition;
          var S:=FEditor.Text;

          // Absolute char offset of the caret within S.
          var Offset:=0;
          for var Ln:=0 to Cp.Line-1 do
            Offset:=Offset+FEditor.Lines[Ln].Length+Length(sLineBreak);
          Offset:=Offset+Cp.Pos;
          if Offset>S.Length then Offset:=S.Length;

          // Replace any active selection, mirroring normal typing behaviour.
          var L:=FEditor.SelLength;
          if L>0 then begin
            var SelP:=FEditor.SelStart;
            if (SelP>=0) and (SelP<=S.Length) then begin
              Delete(S, SelP+1, L);
              Offset:=SelP;
            end;
          end;

          Insert(sLineBreak, S, Offset+1);
          FEditor.Text:=S;

          // Restore caret just after the inserted break.
          var NewPos:=Offset+Length(sLineBreak);
          FEditor.SelStart:=NewPos;
          FEditor.SelLength:=0;
        end
        else begin
          // Plain Enter commits and exits the editor.
          CommitEditing;
        end;
        Key:=0;
        KeyChar:=#0;
      end;
    vkEscape:
      begin
        CancelEditing;
        Key:=0;
        KeyChar:=#0;
      end;
    vkTab:
      begin
        CommitEditing;
        // Move selection to the next/previous cell.
        var NewCol:=FSelectedCell.X+IfThen(ssShift in Shift,-1,1);
        if (NewCol>=0) and (NewCol<FColCount) then
          SelectedCell:=Point(NewCol,FSelectedCell.Y);
        Key:=0;
        KeyChar:=#0;
      end;
  end;
end;

function TMultiHeaderGrid.EditorMinColWidth(ACol, ARow: Integer): single;
begin
  // Base/string grids edit with a TMemo, which is fine in a narrow column.
  Result:=0;
end;

procedure TMultiHeaderGrid.EditorExit(Sender: TObject);
begin
  // Losing focus commits the edit (unless we already cancelled/committed).
  if FEditing then CommitEditing;
end;

procedure TMultiHeaderGrid.SetReadOnly(const Value: Boolean);
begin
  if FReadOnly<>Value then begin
    FReadOnly:=Value;
    if FReadOnly and FEditing then CancelEditing;
  end;
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
    // Reset the resize state when disabled
    if not Value then begin
      FResizeStartColumnIndex:=-1;
      FResizeEndColumnIndex:=-1;
      FResizeRowIndex:=-1;
      ReleaseCapture; // Release the mouse capture
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
    // Check for a header column width change

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
      // Check for a header row height change

      for i:=0 to FHeaderLevels.Count-1 do begin
        Col:=0;
        for j:=0 to FHeaderLevels[i].Count-1 do begin
          // The bottom edge of a blank filler is an internal, visually
          // suppressed border of a merged title stack - don't expose a
          // row-resize handle there.
          if HeaderCellIsFiller(i, Col+FHeaderLevels[i][j].ColSkip) then begin
            Col:=Col+FHeaderLevels[i][j].ColSpan;
            Continue;
          end;

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
      // Check for a grid row height change

      for i:=0 to FRowCount-1 do begin
        CellRect:=GetCellRect(0, i); // Get the row's rectangle
        if not CellRect.IsEmpty then begin
          // Check the row's bottom edge
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

  // Save the global coordinates for accurate tracking
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

  // Save the global coordinates for accurate tracking
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

  // Save the global coordinates for accurate tracking
  GlobalPos:=LocalToAbsolute(PointF(0, Y));
  FResizeStartPos:=PointF(0, GlobalPos.Y);
  FResizeStartHeight:=GetRowHeight(ARow);
  Cursor:=crVSplit;
end;

procedure TMultiHeaderGrid.UpdateGroupColumnWidth(StartCol, EndCol: Integer; TotalWidth: Integer);
var
  i, n           : Integer;
  GroupColCount  : Integer;
  Widths         : array of Integer;
  MinW, MaxW     : array of Integer;
  FResizeStartWidth: Integer;
begin
  GroupColCount:=EndCol-StartCol+1;
  if GroupColCount<=0 then Exit;

  SetLength(Widths, GroupColCount);
  SetLength(MinW,   GroupColCount);
  SetLength(MaxW,   GroupColCount);

  // Effective per-column bounds (at least 10px, honouring Min/MaxWidth).
  var SumMin:=0;
  var SumMax:=0;
  for i:=0 to GroupColCount-1 do begin
    var Lo:=Max(10, FColData[StartCol+i].MinWidth);
    var Hi:=FColData[StartCol+i].MaxWidth;
    if Hi<Lo then Hi:=Lo;
    MinW[i]:=Lo;
    MaxW[i]:=Hi;
    Inc(SumMin, Lo);
    if SumMax<MaxInt then begin
      if Hi>=MaxInt-SumMax then SumMax:=MaxInt else Inc(SumMax, Hi);
    end;
  end;

  // The achievable total is bounded by the sum of per-column min/max widths.
  // Clamping the *target* here (rather than letting individual columns clamp
  // silently) is what keeps the dragged border locked to the mouse: the grid
  // never claims a width it cannot actually realise.
  if TotalWidth<SumMin then TotalWidth:=SumMin;
  if TotalWidth>SumMax then TotalWidth:=SumMax;

  // Start from the proportional distribution of the original widths.
  FResizeStartWidth:=ResizeStartWidth;
  if FResizeStartWidth<=0 then FResizeStartWidth:=GroupColCount;
  for i:=0 to GroupColCount-1 do
    Widths[i]:=Round(FResizeStartWidths[i]*TotalWidth/FResizeStartWidth);

  // Water-filling: clamp to bounds, measure how far the realised total is from
  // the (achievable) target, and push the difference one pixel at a time onto
  // the columns that can still move in that direction. Repeat until the total
  // matches or every column is pinned. Because TotalWidth was pre-clamped to
  // [SumMin..SumMax], this always converges and the realised total equals the
  // target - so the dragged border stays locked to the mouse.
  for var Pass:=0 to GroupColCount do begin
    var Cur:=0;
    for i:=0 to GroupColCount-1 do begin
      Widths[i]:=Min(Max(Widths[i],MinW[i]),MaxW[i]);
      Inc(Cur, Widths[i]);
    end;

    var Diff:=TotalWidth-Cur;       // >0 need to grow, <0 need to shrink
    if Diff=0 then Break;

    // Count columns that can still move the needed way.
    n:=0;
    for i:=0 to GroupColCount-1 do
      if ((Diff>0) and (Widths[i]<MaxW[i])) or
         ((Diff<0) and (Widths[i]>MinW[i])) then Inc(n);
    if n=0 then Break;

    // One-pixel-per-movable-column passes until Diff is consumed.
    var Dir:=Sign(Diff);
    var Left:=Abs(Diff);
    i:=0;
    while (Left>0) do begin
      if ((Dir>0) and (Widths[i]<MaxW[i])) or
         ((Dir<0) and (Widths[i]>MinW[i])) then begin
        Inc(Widths[i], Dir);
        Dec(Left);
      end;
      Inc(i);
      if i>=GroupColCount then i:=0;
      // Safety: if nothing movable remains, stop (handled by next Pass too).
      if Left>0 then begin
        var AnyMovable:=False;
        for var k:=0 to GroupColCount-1 do
          if ((Dir>0) and (Widths[k]<MaxW[k])) or
             ((Dir<0) and (Widths[k]>MinW[k])) then begin AnyMovable:=True; Break; end;
        if not AnyMovable then Break;
      end;
    end;
  end;

  // Commit.
  for i:=0 to GroupColCount-1 do
    ColWidths[StartCol+i]:=Widths[i];

  Invalidate;
  if Assigned(HScrollBar) then
    HScrollBar.Value:=FViewLeft;
end;

procedure TMultiHeaderGrid.UpdateColumnWidth(StartCol, EndCol: Integer; NewWidth: Integer);
begin
  if StartCol<EndCol then begin
    // Resize a group of columns
    UpdateGroupColumnWidth(StartCol, EndCol, NewWidth);
  end else begin
    // Resize a single column
    if NewWidth<10 then begin
      NewWidth:=10; // Minimum width
    end;

    ColWidths[EndCol]:=NewWidth;

    Invalidate;

    // Update the horizontal scrollbar
    if Assigned(HScrollBar) then begin
      HScrollBar.Value:=FViewLeft;
    end;
  end;

  // Narrower columns wrap their captions onto more lines, so the header height
  // must follow the new widths. Only needed when header word-wrap is in play.
  if FHeaderWordWrap then
    AutoSizeHeaders;
end;

procedure TMultiHeaderGrid.UpdateHeaderRowHeight(ARow: Integer; NewHeight: Integer);
begin
  if NewHeight<10 then
    NewHeight:=10; // Minimum height

  FHeaderLevels[ARow].Height:=NewHeight;

  Invalidate;

  // Update the vertical scrollbar
  if Assigned(VScrollBar) then begin
    VScrollBar.BeginUpdate;
    VScrollBar.Value:=FViewTop;
    VScrollBar.EndUpdate;
  end;
end;

procedure TMultiHeaderGrid.UpdateRowHeight(ARow: Integer; NewHeight: Integer);
begin
  if NewHeight<10 then
    NewHeight:=10; // Minimum height

  SetRowHeight(ARow, NewHeight);
  Invalidate;

  // Update the vertical scrollbar
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
  // Keep the inplace editor glued to its cell as the view scrolls.
  if FEditing then PositionEditor;
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
  if not FLastClickIsOnCell then begin
    inherited;
    Exit;
  end;

  inherited;  // fires user OnDblClick (may show a modal)

  // Start editing only if a cell is editable. Defer it so it runs after the
  // double-click sequence (and any modal dialog raised by a user OnDblClick
  // handler) has fully settled - otherwise the modal steals focus from the
  // freshly-focused editor and EditorExit commits/closes it immediately.
  if CanEditCell(FSelectedCell.X,FSelectedCell.Y) then begin
    var C:=FSelectedCell.X;
    var R:=FSelectedCell.Y;
    TThread.ForceQueue(nil,
      procedure
      begin
        if (not FEditing) and CanEditCell(C,R) and
           (C=FSelectedCell.X) and (R=FSelectedCell.Y) then
          StartCellEditing(C,R,#0);
      end);
  end;
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

{ TMHGColumn }

constructor TMHGColumn.Create(Collection: TCollection);
begin
  inherited; // base initialises widths, alignments, separator, etc.
  FColor:=TAlphaColorRec.White;
  FColorIsSet:=False;
  FDateTimeEditor:=dteDateTime;
end;

procedure TMHGColumn.Assign(Source: TPersistent);
begin
  inherited; // copies all shared header-column fields and calls Changed
  if Source is TMHGColumn then begin
    var C:=TMHGColumn(Source);
    FFieldName:=C.FFieldName;
    FColor:=C.FColor;
    FColorIsSet:=C.FColorIsSet;
    FDateTimeEditor:=C.FDateTimeEditor;
    Changed;
  end;
end;

function TMHGColumn.GetGrid: TMultiHeaderDBGrid;
begin
  if Collection is TMHGColumns then
    Result:=TMHGColumns(Collection).Grid
  else
    Result:=nil;
end;

function TMHGColumn.GetDisplayName: string;
begin
  if Title<>'' then
    Result:=Title
  else if FFieldName<>'' then
    Result:=FFieldName
  else
    Result:=inherited GetDisplayName;
end;

function TMHGColumn.IsColorStored: Boolean;
begin
  Result:=FColorIsSet;
end;

procedure TMHGColumn.SetFieldName(const Value: string);
begin
  if FFieldName<>Value then begin
    FFieldName:=Value;
    Changed;
  end;
end;

procedure TMHGColumn.SetColor(const Value: TAlphaColor);
begin
  if (FColor<>Value) or (not FColorIsSet) then begin
    FColor:=Value;
    FColorIsSet:=True;
    Changed;
  end;
end;

{ TMHGColumns }

constructor TMHGColumns.Create(AGrid: TMultiHeaderDBGrid);
begin
  inherited Create(AGrid, TMHGColumn);
  FGrid:=AGrid;
end;

function TMHGColumns.GetItem(Index: Integer): TMHGColumn;
begin
  Result:=TMHGColumn(inherited Items[Index]);
end;

procedure TMHGColumns.SetItem(Index: Integer; const Value: TMHGColumn);
begin
  inherited Items[Index]:=Value;
end;

procedure TMHGColumns.Update(Item: TCollectionItem);
begin
  inherited;
  // Any add/remove/reorder/property change rebuilds the grid header.
  if FGrid<>nil then
    FGrid.ResetTable;
end;

function TMHGColumns.Add: TMHGColumn;
begin
  Result:=TMHGColumn(inherited Add);
end;

function TMHGColumns.AddColumn(const AFieldName: string;
                               const ATitle: string = '';
                               const AGroupHeader: string = ''): TMHGColumn;
begin
  BeginUpdate;
  try
    Result:=Add;
    Result.FieldName:=AFieldName;
    Result.Title:=ATitle;
    Result.GroupHeader:=AGroupHeader;
  finally
    EndUpdate;
  end;
end;

function TMHGColumns.FindByFieldName(const AFieldName: string): TMHGColumn;
begin
  for var i:=0 to Count-1 do begin
    if SameText(Items[i].FieldName,AFieldName) then Exit(Items[i]);
  end;
  Result:=nil;
end;

function TMHGColumns.SetColumnProps(const AFieldName: string;
                                    const ATitle: string;
                                    const AGroupHeader: string = '';
                                    AWidth: Integer = -1;
                                    AAlignment: TTextAlign = TTextAlign.Leading;
                                    AMinWidth: Integer = -1;
                                    AMaxWidth: Integer = -1): TMHGColumn;
begin
  Result:=FindByFieldName(AFieldName);
  if not Assigned(Result) then begin
    Result:=AddColumn(AFieldName);
  end;
  Result.SetProps(ATitle,AGroupHeader,AWidth,AAlignment,AMinWidth,AMaxWidth);
end;

{ TMHGDataLink }

constructor TMHGDataLink.Create(AGrid: TMultiHeaderDBGrid);
begin
  inherited Create;

  FGrid:=AGrid;
  VisualControl:=True;
end;

procedure TMHGDataLink.ActiveChanged;
begin
  if FGrid=nil then Exit;
  FGrid.HandleActiveChanged;
end;

procedure TMHGDataLink.LayoutChanged;
begin
  if FGrid<>nil then FGrid.ResetTable;
end;

procedure TMHGDataLink.DataSetChanged;
begin
  if FGrid<>nil then FGrid.UpdateRowCount;
end;

procedure TMHGDataLink.DataSetScrolled(Distance: Integer);
begin
  if FGrid<>nil then FGrid.SyncSelectedRowFromDataSet;
end;

procedure TMHGDataLink.RecordChanged(Field: TField);
begin
  if FGrid<>nil then FGrid.InvalidateCurrentRow;
end;

{ TMultiHeaderDBGrid }

function TMultiHeaderDBGrid.Column(FieldNum: integer): THeaderElement;
begin
  if DataSet=nil then Exit(nil);

  Result:=FHeaderLevels.Last[FieldNum];
end;

{$WARN USE_BEFORE_DEF OFF}
procedure TMultiHeaderDBGrid.CleanUpCache;
begin
  while FCellTexts.Count>FRowCacheSize do begin
    var MinTime: TDateTime;
    var MinRow:=-1;
    var IsFirst:=True;
    for var Item in FCellTexts do begin
      if IsFirst or (Item.Value.LastUsage<MinTime) then begin
        MinTime:=Item.Value.LastUsage;
        MinRow:=Item.Key;
        IsFirst:=False;
      end;
    end;
    if MinRow>=0 then begin
      FCellTexts.Remove(MinRow);
    end else begin
      Break;
    end;
  end;
end;
{$WARN USE_BEFORE_DEF ON}

function TMultiHeaderDBGrid.Column(FieldName: string): THeaderElement;
begin
  if DataSet=nil then Exit(nil);

  var Fields:=GetVisibleFields;
  for var i:=0 to High(Fields) do begin
    if SameText(Fields[i].FieldName,FieldName) then Exit(Column(i));
  end;
  Result:=nil;
end;

constructor TMultiHeaderDBGrid.Create(AOwner: TComponent);
begin
  inherited;

  FRowCacheSize:=1000;
  FCellTexts:=TRowsData.Create(FRowCacheSize);
  FDataLink:=TMHGDataLink.Create(Self);
  FColumns:=TMHGColumns.Create(Self);
end;

destructor TMultiHeaderDBGrid.Destroy;
begin
  FDataLink.Free;
  FCellTexts.Free;
  FColumns.Free;

  inherited;
end;

procedure TMultiHeaderDBGrid.DoGetCellStyle(ACol, ARow: Integer; var Style: TCellStyle);
begin
  // Per-column default appearance from the Columns collection
  // (applies to the data cells, beneath any per-row overrides).
  // FColMap maps the effective column index to its collection item and
  // is refreshed by ResetTable, so this stays O(1) per cell.
  if (ACol>=0) and (ACol<Length(FColMap)) then begin
    var Col:=FColMap[ACol];
    if Col<>nil then begin
      if Col.ColorIsSet and (not Style.CellColorIsSet) then
        Style.CellColor:=Col.Color;
      // Surface the column's word-wrap so per-cell style consumers
      // (drawing, autosize) wrap consistently. Horizontal alignment is
      // applied through ColTextHAlignment in ResetTable.
      if Col.WordWrap and (not Style.WordWrapIsSet) then begin
        Style.WordWrap:=True;
        Style.WordWrapIsSet:=True;
      end;
      // Vertical alignment: column default, overridable per cell.
      if not Style.TextVAlignmentIsSet then begin
        Style.TextVAlignment:=Col.VertAlignment;
        Style.TextVAlignmentIsSet:=True;
      end;
    end;
  end;

  // Don't create a cache entry just to check the style -
  // it's enough to see whether the row is already cached.
  if (ARow>=0) and (ACol>=0) then begin
    var RowData: TDBRowData;
    if FCellTexts.TryGetValue(ARow,RowData) then begin
      if (ACol<=High(RowData.Styles)) and (not RowData.Styles[ACol].IsEmpty) then begin
        Style:=RowData.Styles[ACol];
      end;
    end;
  end;

  inherited;
end;

procedure TMultiHeaderDBGrid.DoGetCellText(ACol, ARow: Integer; var Text: string);
begin
  if (ARow<0) or (ACol<0) then Exit;
  if (ARow>=RowCount) or (ACol>=ColCount) then Exit;

  var DS:=DataSet;
  if (DS=nil) or (not DS.Active) then Exit;

  var RowData:=GetRowData(ARow);
  if ACol<=High(RowData.Cells) then begin
    Text:=RowData.Cells[ACol];
  end;

  inherited;
end;

procedure TMultiHeaderDBGrid.DoSetCellStyle(ACol, ARow: Integer; const Style: TCellStyle);
begin
  inherited;

  if (ARow<0) or (ACol<0) then Exit;
  if (ARow>=RowCount) or (ACol>=ColCount) then Exit;

  var DS:=DataSet;
  if (DS=nil) or (not DS.Active) then Exit;

  // The style isn't part of the DataSet's data - we store it as a
  // presentation-only overlay in the same per-row cache used for
  // the cell text. The cache entry is created on demand
  // (loading the cell text from the DataSet if needed).
  var RowData:=GetRowData(ARow);

  if ACol>High(RowData.Styles) then begin
    SetLength(RowData.Styles,Max(ACol+1,ColCount));
  end;

  RowData.Styles[ACol]:=Style;
  RowData.LastUsage:=Now;

  // RowData is a copy of the cache entry; write the copy back
  // to the dictionary so the change is persisted.
  FCellTexts[ARow]:=RowData;

  Invalidate;
end;

procedure TMultiHeaderDBGrid.DoSetCellText(ACol, ARow: Integer; const Text: string);
begin
  inherited;

  if (ARow<0) or (ACol<0) then Exit;
  if (ARow>=RowCount) or (ACol>=ColCount) then Exit;

  var DS:=DataSet;
  if (DS=nil) or (not DS.Active) then Exit;

  var Fields:=GetVisibleFields;
  if ACol>High(Fields) then Exit;

  var Field:=Fields[ACol];
  if Field.ReadOnly then Exit;

  if ARow=DS.RecNo-1 then begin
    // The DataSet's current record already matches this grid row -
    // no Bookmark navigation is needed.
    SetFieldValue(DS,Field,Text);
  end else begin
    DS.DisableControls;
    try
      var Bookmark:=DS.Bookmark;
      try
        DS.RecNo:=ARow+1;
        SetFieldValue(DS,Field,Text);
      finally
        if DS.BookmarkValid(Bookmark) then begin
          DS.Bookmark:=Bookmark;
        end;
      end;
    finally
      DS.EnableControls;
    end;
  end;

  // The cached row text is now stale - drop it so that
  // DoGetCellText re-reads the current field values.
  // The style cache (DoSetCellStyle) for this row is lost too -
  // see the chat notes about possible follow-up improvements.
  FCellTexts.Remove(ARow);
  Invalidate;
end;

procedure TMultiHeaderDBGrid.SetFieldValue(DS: TDataSet; Field: TField; const Text: string);
begin
  if not (DS.State in [dsEdit,dsInsert]) then begin
    DS.Edit;
  end;
  try
    Field.AsString:=Text;
    DS.Post;
  except
    DS.Cancel;
    raise;
  end;
end;

function TMultiHeaderDBGrid.DataSet: TDataSet;
begin
  Result:=FDataLink.DataSet;
end;

function TMultiHeaderDBGrid.GetDataSource: TDataSource;
begin
  Result:=FDataLink.DataSource;
end;

procedure TMultiHeaderDBGrid.SetColumns(const Value: TMHGColumns);
begin
  FColumns.Assign(Value);
end;

function TMultiHeaderDBGrid.ResolveColumns(out AFields: TArray<TField>): TArray<TMHGColumn>;
// Produces the ordered list of effective columns and the matching fields.
// The Columns collection is authoritative: it always drives the layout
// (column order, titles, grouping). An empty Columns collection means an
// empty grid - use AutoCreateColumns / the editor's Import to (re)populate
// from the DataSet's visible fields.
begin
  AFields:=nil;
  Result:=nil;

  var DS:=DataSet;
  if DS=nil then Exit;

  // Columns-driven, always. Skip invisible columns and columns whose
  // FieldName does not resolve to a real field.
  SetLength(Result,FColumns.Count);
  SetLength(AFields,FColumns.Count);
  var Cnt:=0;
  for var i:=0 to FColumns.Count-1 do begin
    var Col:=FColumns[i];
    if not Col.Visible then Continue;
    var F:=DS.FindField(Col.FieldName);
    if F=nil then Continue;
    Result[Cnt]:=Col;
    AFields[Cnt]:=F;
    Inc(Cnt);
  end;
  SetLength(Result,Cnt);
  SetLength(AFields,Cnt);
end;

function TMultiHeaderDBGrid.GetVisibleFields: TArray<TField>;
begin
  ResolveColumns(Result);
end;

function TMultiHeaderDBGrid.GetRowData(ARow: Integer): TDBRowData;
begin
  if FCellTexts.TryGetValue(ARow,Result) then begin
    // Result is a copy of the entry (Cells/Styles are refcounted
    // dynamic arrays - the data is shared with the dictionary entry,
    // but Result itself is independent). LastUsage is updated and
    // the copy is written back for correct LRU behaviour.
    Result.LastUsage:=Now;
    FCellTexts[ARow]:=Result;
    Exit;
  end;

  var DS:=DataSet;
  var Fields:=GetVisibleFields;

  SetLength(Result.Cells,Length(Fields));

  if ARow=DS.RecNo-1 then begin
    for var i:=0 to High(Fields) do begin
      Result.Cells[i]:=Fields[i].AsString;
    end;
  end else begin
    DS.DisableControls;
    try
      var Bookmark:=DS.Bookmark;
      try
        DS.RecNo:=ARow+1;
        for var i:=0 to High(Fields) do begin
          Result.Cells[i]:=Fields[i].AsString;
        end;
      finally
        DS.Bookmark:=Bookmark;
      end;
    finally
      DS.EnableControls;
    end;
  end;

  Result.LastUsage:=Now;
  FCellTexts.Add(ARow,Result);

  // Result is an independent copy (with its own references to
  // Cells/Styles), so CleanUpCache can safely remove the entry from
  // the dictionary - even evicting ARow itself won't affect Result.
  CleanUpCache;
end;

procedure TMultiHeaderDBGrid.ResetTable;
begin
  // Guard against re-entrancy: building the header changes ColCount/
  // column widths, none of which should trigger another rebuild.
  if FRebuildingHeader then Exit;
  FRebuildingHeader:=True;
  try
    FCellTexts.Clear;

    var DS:=DataSet;
    if (DS=nil) or (not DS.Active) then begin
      FColMap:=nil;
      ColCount:=5;
      RowCount:=1;
      Header.Clear;
      Header.AddRow.FillRow(IfThen(Name='',ClassName,Name));
      Exit;
    end;

    var Fields: TArray<TField>;
    var Cols:=ResolveColumns(Fields);
    FColMap:=Cols; // cache for O(1) per-cell colour lookup in DoGetCellStyle

    ColCount:=Length(Fields);
    if ColCount=0 then begin
      FColMap:=nil;
      Header.Clear;
      Header.AddRow.FillRow(IfThen(Name='',ClassName,Name));
      UpdateRowCount;
      Exit;
    end;

    // Compute the layout signature (which fields, in what order) BEFORE the
    // rebuild, so we know whether this is a structural change or a pure
    // re-apply (e.g. a Min/MaxWidth toggle).
    var Layout:='';
    for var i:=0 to High(Fields) do
      Layout:=Layout+Fields[i].FieldName+';';
    var LayoutChanged:=Layout<>FLastColLayout;
    FLastColLayout:=Layout;

    // Preserve the current header row heights so a non-structural rebuild does
    // not flatten them back to the fixed build-time default (which made every
    // header row the same size when the Limit-widths checkbox was toggled).
    var SavedHeights: TArray<Integer>;
    if not LayoutChanged then begin
      SetLength(SavedHeights, FHeaderLevels.Count);
      for var i:=0 to FHeaderLevels.Count-1 do
        SavedHeights[i]:=FHeaderLevels[i].Height;
    end;

    BuildGroupedHeader(Fields, Cols);

    // Apply per-column geometry/alignment taken from the Columns
    // collection. Column *data-cell* colour is handled separately in
    // DoGetCellStyle via FColMap, so it is intentionally not set here.
    if Length(Cols)=ColCount then begin
      for var i:=0 to ColCount-1 do begin
        var Col:=Cols[i];
        if Col=nil then Continue;
        // Initial width only on a fresh/changed layout; otherwise the live
        // (user-dragged) width stands.
        if LayoutChanged and (Col.Width>0) then ColWidths[i]:=Col.Width;
        if Col.MinWidth>0 then ColMinWidth[i]:=Col.MinWidth;
        if Col.MaxWidth>0 then ColMaxWidth[i]:=Col.MaxWidth;
        // When a column drops its limits (Min/Max back to 0), restore the
        // permissive defaults so a previously clamped width can grow again.
        if Col.MinWidth<=0 then ColMinWidth[i]:=0;
        if Col.MaxWidth<=0 then ColMaxWidth[i]:=MaxInt;
        ColWordWrap[i]:=Col.WordWrap;
        ColTextHAlignment[i]:=Col.Alignment;
        ColTextVAlignment[i]:=Col.VertAlignment;
      end;
    end;

    if LayoutChanged then
      // Fresh structure: fit header heights to the new captions/widths.
      AutoSizeHeaders
    else
      // Pure re-apply: keep the header heights exactly as they were.
      for var i:=0 to Min(High(SavedHeights),FHeaderLevels.Count-1) do
        FHeaderLevels[i].Height:=SavedHeights[i];

    UpdateRowCount;
  finally
    FRebuildingHeader:=False;
  end;
end;

procedure TMultiHeaderDBGrid.BuildGroupedHeader(const AFields: TArray<TField>;
                                                const ACols: TArray<TMHGColumn>);
// Generates the stacked group-header rows (UniGUI-style) plus the
// final title row from the Columns collection.
//
// AFields and ACols are the already-resolved (visible, field-matched)
// parallel arrays produced by ResolveColumns, of equal length.
var
  Paths: TArray<TArray<string>>;
begin
  Header.Clear;

  var N:=Length(AFields);
  if N=0 then Exit;

  var ColObjs: TArray<TMHGColumn> := ACols;
  if Length(ColObjs)<>N then ColObjs:=nil;

  // Collect the group path of each column and the deepest path length.
  SetLength(Paths,N);
  var MaxDepth:=0;
  for var i:=0 to N-1 do begin
    if ColObjs<>nil then
      Paths[i]:=ColObjs[i].GroupPath
    else
      Paths[i]:=nil;
    if Length(Paths[i])>MaxDepth then MaxDepth:=Length(Paths[i]);
  end;

  // One header level per group depth, drawn top (level 0) to bottom.
  //
  // The underlying header model positions a row's elements purely by
  // accumulating ColSpan left-to-right (plus a ColSkip that only
  // accounts for RowSpans coming from rows *above*). It has no concept
  // of a horizontal gap inside a row. So every group row is *fully
  // tiled*: each column position gets exactly one element - either a
  // merged group caption spanning its run of same-path columns, or a
  // blank 1-wide filler where a column has no caption at this level.
  // This keeps every row perfectly column-aligned regardless of how
  // groups, ungrouped columns and differing depths are interleaved.
  for var Level:=0 to MaxDepth-1 do begin
    var Row:=Header.AddRow(30);
    var Col:=0;
    while Col<N do begin
      // No group caption for this column at this level -> blank filler.
      if Length(Paths[Col])<=Level then begin
        Row.AddColumn('',1);
        Inc(Col);
        Continue;
      end;

      // Merge consecutive columns that share the same path prefix
      // [0..Level] into a single spanning group cell.
      var Span:=1;
      while (Col+Span<N) and (Length(Paths[Col+Span])>Level) do begin
        var Same:=True;
        for var L:=0 to Level do begin
          if Paths[Col+Span][L]<>Paths[Col][L] then begin
            Same:=False;
            Break;
          end;
        end;
        if not Same then Break;
        Inc(Span);
      end;

      Row.AddColumn(Paths[Col][Level],Span);
      Inc(Col,Span);
    end;
  end;

  // Final row: the column titles themselves. One element per column,
  // so column index lines up with header element index (Column(i)
  // continues to return the title element of column i).
  var TitleRow:=Header.AddRow(30);
  for var i:=0 to N-1 do begin
    var Caption:=AFields[i].DisplayLabel;
    if (ColObjs<>nil) and (ColObjs[i].Title<>'') then
      Caption:=ColObjs[i].Title;
    var El:=TitleRow.AddColumn(Caption);
    if ColObjs<>nil then begin
      El.Style.TextHAlignment:=ColObjs[i].HeaderAlignment;
      El.Style.TextHAlignmentIsSet:=True;
      El.Style.TextVAlignment:=ColObjs[i].HeaderVertAlignment;
      El.Style.TextVAlignmentIsSet:=True;
    end;
  end;
end;

function TMultiHeaderDBGrid.DatasetSignature: string;
begin
  Result:='';
  var DS:=DataSet;
  if (DS=nil) or (not DS.Active) then Exit;

  // DataSet instance identity (so a different DataSet object always differs)
  // plus the ordered field-name list (so the same DataSet reopened against a
  // different table/query - hence a different field set - also differs, while
  // a plain close/reopen of the same table yields the same signature).
  var SB:=TStringBuilder.Create;
  try
    SB.Append(IntToHex(NativeUInt(DS),SizeOf(Pointer)*2));
    SB.Append('|');
    for var i:=0 to DS.FieldCount-1 do begin
      SB.Append(DS.Fields[i].FieldName);
      SB.Append(';');
    end;
    Result:=SB.ToString;
  finally
    SB.Free;
  end;
end;

procedure TMultiHeaderDBGrid.HandleActiveChanged;
begin
  if not ((DataSet<>nil) and DataSet.Active) then begin
    // Closed: keep current columns; just rebuild (shows placeholder if empty).
    ResetTable;
    Exit;
  end;

  var Sig:=DatasetSignature;
  // Auto-create columns only when there are none AND this is a source/table we
  // have not auto-created for yet. This fills a fresh grid or a newly opened
  // table, but never resurrects columns the user deleted on the same table.
  if (FColumns.Count=0) and (Sig<>FLastAutoSig) then begin
    FLastAutoSig:=Sig;
    AutoCreateColumns; // rebuilds via the collection's Update -> ResetTable
    // If the table genuinely had no visible fields, AutoCreateColumns adds
    // nothing and no rebuild is triggered; ensure the grid is still reset.
    if FColumns.Count=0 then ResetTable;
  end else begin
    // Remember the current table so a later delete-all on it stays empty.
    FLastAutoSig:=Sig;
    ResetTable;
  end;
end;

procedure TMultiHeaderDBGrid.AutoCreateColumns;
begin
  var DS:=DataSet;
  if (DS=nil) or (not DS.Active) then Exit;

  FColumns.BeginUpdate;
  try
    for var i:=0 to DS.FieldCount-1 do begin
      var F:=DS.Fields[i];
      if not F.Visible then Continue;
      if FColumns.FindByFieldName(F.FieldName)=nil then
        FColumns.AddColumn(F.FieldName,F.DisplayLabel);
    end;
  finally
    FColumns.EndUpdate;
  end;
end;

procedure TMultiHeaderDBGrid.UpdateRowCount;
begin
  var DS:=DataSet;
  if (DS=nil) or (not DS.Active) then begin
    RowCount:=0;
    Exit;
  end;

  FCellTexts.Clear;

  if DS.RecordCount>=0 then begin
    RowCount:=DS.RecordCount;
  end else begin
    // The dataset cannot report the number of records
    // (unidirectional cursor) - the grid cannot show one.
    RowCount:=0;
  end;

  SyncSelectedRowFromDataSet;
  Invalidate;
end;

procedure TMultiHeaderDBGrid.SyncSelectedRowFromDataSet;
begin
  var DS:=DataSet;
  if (DS=nil) or (not DS.Active) or DS.IsEmpty then Exit;
  if DS.RecNo<=0 then Exit;

  var NewRow:=DS.RecNo-1;
  if NewRow=Row then Exit;
  if (NewRow<0) or (NewRow>=RowCount) then Exit;

  FUpdatingRow:=True;
  try
    Row:=NewRow;
  finally
    FUpdatingRow:=False;
  end;
end;

procedure TMultiHeaderDBGrid.InvalidateCurrentRow;
begin
  var DS:=DataSet;
  if (DS=nil) or (not DS.Active) or DS.IsEmpty then Exit;
  if DS.RecNo<=0 then Exit;

  FCellTexts.Remove(DS.RecNo-1);
  Invalidate;
end;

procedure TMultiHeaderDBGrid.DoSelectCell;
begin
  inherited;

  if FUpdatingRow then Exit;

  var DS:=DataSet;
  if (DS=nil) or (not DS.Active) then Exit;
  if (Row<0) or (Row>=DS.RecordCount) then Exit;
  if DS.RecNo=Row+1 then Exit;

  FUpdatingRow:=True;
  try
    try
      DS.RecNo:=Row+1;
    except
      // The dataset does not support arbitrary RecNo navigation.
    end;
  finally
    FUpdatingRow:=False;
  end;
end;

function TMultiHeaderDBGrid.FieldForCol(ACol: Integer): TField;
begin
  Result:=nil;
  var DS:=DataSet;
  if (DS=nil) or (not DS.Active) then Exit;
  var Fields:=GetVisibleFields;
  if (ACol<0) or (ACol>High(Fields)) then Exit;
  Result:=Fields[ACol];
end;

function TMultiHeaderDBGrid.EditorKindForField(Field: TField): TMHGEditorKind;
begin
  Result:=ekNone;
  if Field=nil then Exit;

  // Read-only and non-data (calculated/internal-calc) fields are never edited.
  if Field.ReadOnly then Exit;
  if Field.FieldKind in [fkCalculated, fkInternalCalc] then Exit;

  // Lookup fields -> combobox of the looked-up values.
  if Field.FieldKind=fkLookup then
    Exit(ekComboBox);

  case Field.DataType of
    ftBoolean:
      Result:=ekCheckBox;
    ftDate:
      Result:=ekDate;
    ftTime:
      Result:=ekTime;
    ftDateTime, ftTimeStamp, ftOraTimeStamp:
      begin
        // The owning column decides date-only vs date+time (default).
        var Col:=ColumnForField(Field);
        if (Col<>nil) and (Col.DateTimeEditor=dteDate) then
          Result:=ekDate
        else
          Result:=ekDateTime;
      end;

    // Whitelisted text types: edit as text through a TMemo and round-trip
    // via Field.AsString (guarded by try/except at commit).
    ftString, ftWideString, ftFixedChar, ftFixedWideChar, ftGuid,
    ftMemo, ftWideMemo, ftFmtMemo, ftOraClob:
      Result:=ekMemo;

    // Numeric types: edited through a dedicated TNumberBox (integer mode for
    // whole-number fields, decimal mode for floating/scaled fields).
    ftSmallint, ftInteger, ftWord, ftLargeint, ftLongWord, ftShortint, ftByte,
    ftFloat, ftCurrency, ftBCD, ftFMTBcd, ftSingle, ftExtended:
      Result:=ekNumber;
  else
    // BLOB / binary / structured / cursor / array / stream etc. - no editor.
    Result:=ekNone;
  end;
end;

function TMultiHeaderDBGrid.FieldIsInteger(Field: TField): Boolean;
// Whole-number field types -> the number editor runs in integer mode.
begin
  Result:=(Field<>nil) and
          (Field.DataType in [ftSmallint, ftInteger, ftWord, ftLargeint,
                              ftLongWord, ftShortint, ftByte]);
end;

procedure TMultiHeaderDBGrid.NumberRangeForField(Field: TField; out AMin, AMax: Double);
// Min/Max the number editor should allow for a given field's data type, so it
// never clamps a valid stored value. For float/64-bit types a wide finite
// range is used (TNumberBox.Min/Max are Double).
const
  // A wide but finite range for floating-point fields. Large enough for any
  // realistic stored value while staying well clear of Double's precision
  // limits and any infinity handling in the control.
  FLOAT_LIMIT = 1.0E15;
begin
  AMin:=-FLOAT_LIMIT;
  AMax:= FLOAT_LIMIT;
  if Field=nil then Exit;

  case Field.DataType of
    ftByte:     begin AMin:=0;          AMax:=255;         end;
    ftShortint: begin AMin:=-128;       AMax:=127;         end;
    ftWord:     begin AMin:=0;          AMax:=65535;       end;
    ftSmallint: begin AMin:=-32768;     AMax:=32767;       end;
    ftLongWord: begin AMin:=0;          AMax:=4294967295;  end;
    ftInteger:  begin AMin:=-2147483648; AMax:=2147483647; end;
    // 64-bit integers exceed Double's exact-integer range; use the wide
    // finite limit rather than the exact (unrepresentable) Int64 bounds.
    ftLargeint: begin AMin:=-FLOAT_LIMIT; AMax:=FLOAT_LIMIT; end;
  else
    // ftFloat / ftCurrency / ftBCD / ftFMTBcd / ftSingle / ftExtended.
    AMin:=-FLOAT_LIMIT;
    AMax:= FLOAT_LIMIT;
  end;
end;

function TMultiHeaderDBGrid.ColumnForField(Field: TField): TMHGColumn;
begin
  Result:=nil;
  if Field=nil then Exit;
  for var i:=0 to FColumns.Count-1 do
    if SameText(FColumns[i].FieldName, Field.FieldName) then
      Exit(FColumns[i]);
end;

procedure TMultiHeaderDBGrid.EnsureTypedEditors;

  procedure WireCommon(C: TControl);
  begin
    C.Parent:=Self;
    C.Visible:=False;
    C.OnKeyDown:=EditorKeyDown;
    C.OnExit:=EditorExit;
  end;

begin
  if FComboEditor=nil then begin
    FComboEditor:=TComboBox.Create(Self);
    WireCommon(FComboEditor);
  end;
  if FNumberEditor=nil then begin
    FNumberEditor:=TNumberBox.Create(Self);
    // Plain numeric box: disable mouse-drag increments and hide the spin
    // up/down buttons (done in NumberEditorApplyStyle, which finds the button
    // style resources - portable across Delphi versions). Value mode (integer
    // vs decimal) and decimal digits are set per field in LoadEditorValue.
    FNumberEditor.HorzIncrement:=0;
    FNumberEditor.VertIncrement:=0;
    FNumberEditor.OnApplyStyleLookup:=NumberEditorApplyStyle;
    WireCommon(FNumberEditor);
  end;
  if FDateEditor=nil then begin
    FDateEditor:=TDateEdit.Create(Self);
    // Compact, fixed-width date so the control doesn't reserve space for a
    // long localized format (which left empty padding inside the box).
    FDateEditor.Format:='dd.MM.yyyy';
    WireCommon(FDateEditor);
  end;
  if FTimeEditor=nil then begin
    FTimeEditor:=TTimeEdit.Create(Self);
    // Show seconds; compact 24h layout keeps the box tight.
    FTimeEditor.Format:='HH:nn:ss';
    WireCommon(FTimeEditor);
  end;
end;

procedure TMultiHeaderDBGrid.NumberEditorApplyStyle(Sender: TObject);

  // Hide a spin-button style sub-control by name, if present.
  procedure HideButton(Ctrl: TStyledControl; const AName: string);
  begin
    var Obj:=Ctrl.FindStyleResource(AName);
    if Obj is TControl then
      TControl(Obj).Visible:=False;
  end;

begin
  // The TNumberBox spin arrows are style sub-controls; their names vary a
  // little across themes/versions, so hide every known variant. Missing ones
  // are simply skipped, leaving a plain numeric edit box.
  if not (Sender is TStyledControl) then Exit;
  var Ctrl:=TStyledControl(Sender);
  HideButton(Ctrl, 'plusbutton');
  HideButton(Ctrl, 'minusbutton');
  HideButton(Ctrl, 'upbutton');
  HideButton(Ctrl, 'downbutton');
  HideButton(Ctrl, 'spinup');
  HideButton(Ctrl, 'spindown');
  HideButton(Ctrl, 'buttonsbox');   // some themes group both arrows here
  HideButton(Ctrl, 'updownbuttons');
end;

function TMultiHeaderDBGrid.EditorMinColWidth(ACol, ARow: Integer): single;
const
  // Usable minimum widths (px) for the fixed-size typed editors. The memo and
  // checkbox are comfortable in a narrow column, so they keep the default 0.
  MINW_COMBO    = 120;
  MINW_DATE     = 83;  // single 'dd.MM.yyyy' date editor
  MINW_TIME     = 95;  // HH:nn:ss + picker button
  MINW_DATETIME = 156; // date + time composite (date +5 / time -3 vs even split)
begin
  Result:=0;
  var Field:=FieldForCol(ACol);
  if Field=nil then Exit;

  case EditorKindForField(Field) of
    ekComboBox: Result:=MINW_COMBO+FGridLineWidth/2;
    ekDate:     Result:=MINW_DATE+FGridLineWidth/2;
    ekTime:     Result:=MINW_TIME+FGridLineWidth/2;
    ekDateTime: Result:=MINW_DATETIME+FGridLineWidth/2;
  else
    Result:=0; // ekMemo / ekNumber / ekCheckBox / ekNone
  end;
end;

function TMultiHeaderDBGrid.CanEditCell(ACol, ARow: Integer): Boolean;
begin
  Result:=False;
  if FReadOnly then Exit;
  if (ACol<0) or (ARow<0) or (ARow>=RowCount) then Exit;

  var Field:=FieldForCol(ACol);
  if Field=nil then Exit;

  // Editable through an inplace editor only when a typed editor exists.
  // Boolean fields are NOT edited this way - they toggle in place
  // (CellIsToggle / ToggleCell), so they are excluded here.
  var Kind:=EditorKindForField(Field);
  Result:=(Kind<>ekNone) and (Kind<>ekCheckBox);
end;

function TMultiHeaderDBGrid.CellIsToggle(ACol, ARow: Integer): Boolean;
begin
  Result:=False;
  if FReadOnly then Exit;
  if (ACol<0) or (ARow<0) or (ARow>=RowCount) then Exit;

  var Field:=FieldForCol(ACol);
  Result:=(Field<>nil) and (EditorKindForField(Field)=ekCheckBox);
end;

function TMultiHeaderDBGrid.ToggleCell(ACol, ARow: Integer): Boolean;
// Flips a boolean field's value in place and posts it, with the same
// off-row navigation / bookmark handling used by CommitEditorValue.
var
  DS: TDataSet;
  Field: TField;

  procedure ApplyToCurrentRecord;
  begin
    if not (DS.State in [dsEdit, dsInsert]) then DS.Edit;
    try
      Field.AsBoolean:=not Field.AsBoolean;
      DS.Post;
    except
      DS.Cancel;
      raise;
    end;
  end;

begin
  Result:=False;
  Field:=FieldForCol(ACol);
  if Field=nil then Exit;
  if Field.ReadOnly then Exit;

  DS:=DataSet;
  if (DS=nil) or (not DS.Active) then Exit;

  try
    if ARow=DS.RecNo-1 then
      ApplyToCurrentRecord
    else begin
      DS.DisableControls;
      try
        var Bookmark:=DS.Bookmark;
        try
          DS.RecNo:=ARow+1;
          ApplyToCurrentRecord;
        finally
          if DS.BookmarkValid(Bookmark) then
            DS.Bookmark:=Bookmark;
        end;
      finally
        DS.EnableControls;
      end;
    end;
    Result:=True;
  except
    Result:=False;
  end;

  // The cached row text is now stale - drop it so DoGetCellText re-reads.
  FCellTexts.Remove(ARow);
  Invalidate;
end;

function TMultiHeaderDBGrid.PrepareCellEditor(ACol, ARow: Integer): TControl;
begin
  Result:=nil;
  var Field:=FieldForCol(ACol);
  if Field=nil then Exit;

  FEditField:=Field;
  var Kind:=EditorKindForField(Field);

  case Kind of
    ekMemo:
      // Reuse the shared TMemo via the base implementation.
      Result:=inherited PrepareCellEditor(ACol, ARow);

    ekNumber:
      begin
        EnsureTypedEditors;
        // Follow the column's horizontal text alignment, like the TMemo does.
        // Remove 'Other' from StyledSettings so the manual HorzAlign is honored
        // instead of being overridden by the control's style.
        FNumberEditor.StyledSettings:=FNumberEditor.StyledSettings-[TStyledSetting.Other];
        FNumberEditor.TextSettings.HorzAlign:=ColTextHAlignment[ACol];
        FActiveEditor:=FNumberEditor;
        Result:=FNumberEditor;
      end;

    ekComboBox:
      begin
        EnsureTypedEditors;
        // Populate the list from the lookup dataset's display values.
        FComboEditor.Items.Clear;
        if (Field.FieldKind=fkLookup) and (Field.LookupDataSet<>nil) then begin
          var LDS:=Field.LookupDataSet;
          var BM:=LDS.Bookmark;
          LDS.DisableControls;
          try
            LDS.First;
            while not LDS.Eof do begin
              FComboEditor.Items.Add(LDS.FieldByName(Field.LookupResultField).AsString);
              LDS.Next;
            end;
          finally
            if LDS.BookmarkValid(BM) then LDS.Bookmark:=BM;
            LDS.EnableControls;
          end;
        end;
        FActiveEditor:=FComboEditor;
        Result:=FComboEditor;
      end;

    ekDate:
      begin
        EnsureTypedEditors;
        FActiveEditor:=FDateEditor;
        Result:=FDateEditor;
      end;

    ekTime:
      begin
        EnsureTypedEditors;
        FActiveEditor:=FTimeEditor;
        Result:=FTimeEditor;
      end;

    ekDateTime:
      begin
        EnsureTypedEditors;
        // Composite: the date control is the focused/active editor; the time
        // control rides alongside it (positioned in DoPositionComposite below
        // via the base PositionEditor, then nudged). For layout simplicity the
        // time editor is shown to the right of the date editor.
        FActiveEditor:=FDateEditor;
        Result:=FDateEditor;
      end;
  else
    Result:=nil; // ekNone - not editable
  end;
end;

procedure TMultiHeaderDBGrid.PlaceEditorCaretAtEnd(Ed: TControl);
begin
  // Typed text editors (number / date / time) descend from TCustomEdit and
  // expose GoToTextEnd; the combo box has no caret. Fall back to the base for
  // the shared TMemo.
  if Ed is TCustomEdit then begin
    var E:=TCustomEdit(Ed);
    E.GoToTextEnd;
    // Belt-and-suspenders: set the caret index explicitly and clear any
    // selection, since GoToTextEnd can be a no-op right after SetFocus on
    // some platforms / before the style finishes applying.
    E.SelLength:=0;
    E.CaretPosition:=E.Text.Length;
  end else
    inherited;
end;

procedure TMultiHeaderDBGrid.LoadEditorValue(ACol, ARow: Integer; InitialChar: Char);
begin
  var Field:=FEditField;
  if Field=nil then begin
    inherited; // memo fallback
    Exit;
  end;

  case EditorKindForField(Field) of
    ekMemo:
      inherited;  // load text into the shared TMemo

    ekNumber:
      begin
        // Integer fields -> integer mode (no decimals); float/scaled fields ->
        // float mode with a sensible number of decimal places.
        if FieldIsInteger(Field) then begin
          FNumberEditor.ValueType:=TNumValueType.Integer;
          FNumberEditor.DecimalDigits:=0;
        end else begin
          FNumberEditor.ValueType:=TNumValueType.Float;
          var Digits:=4;
          if Field.DataType in [ftCurrency, ftBCD, ftFMTBcd] then Digits:=2;
          FNumberEditor.DecimalDigits:=Digits;
        end;
        // Set the editor's allowed range to the field type's full range.
        // Without this TNumberBox keeps its default (0..100) and clamps/
        // corrupts the value.
        var MinV, MaxV: Double;
        NumberRangeForField(Field, MinV, MaxV);
        FNumberEditor.Min:=MinV;
        FNumberEditor.Max:=MaxV;
        // Typing a digit/sign starts the value fresh from that char; F2 /
        // double-click (InitialChar=#0) loads the field's current value.
        if (InitialChar>=' ') and CharInSet(InitialChar, ['0'..'9','-','+','.',',']) then
          FNumberEditor.Text:=InitialChar
        else if Field.IsNull then
          FNumberEditor.Value:=0
        else
          FNumberEditor.Value:=Field.AsFloat;
      end;

    ekComboBox:
      FComboEditor.ItemIndex:=FComboEditor.Items.IndexOf(Field.AsString);

    ekDate:
      if Field.IsNull then FDateEditor.Date:=Now
      else FDateEditor.Date:=Field.AsDateTime;

    ekTime:
      if Field.IsNull then FTimeEditor.Time:=Now
      else FTimeEditor.Time:=Field.AsDateTime;

    ekDateTime:
      begin
        var V: TDateTime;
        if Field.IsNull then V:=Now else V:=Field.AsDateTime;
        FDateEditor.Date:=V;
        FTimeEditor.Time:=V;
        FTimeEditor.Visible:=True;
      end;
  end;
end;

function TMultiHeaderDBGrid.GetEditorText: string;
begin
  var Field:=FEditField;
  if Field=nil then Exit(inherited GetEditorText);

  case EditorKindForField(Field) of
    ekMemo:      Result:=inherited GetEditorText;
    ekNumber:    if FieldIsInteger(FEditField) then Result:=Trunc(FNumberEditor.Value).ToString
                 else Result:=FNumberEditor.Value.ToString;
    ekComboBox:  if FComboEditor.Selected<>nil then Result:=FComboEditor.Selected.Text else Result:='';
    ekDate:      Result:=DateToStr(FDateEditor.Date);
    ekTime:      Result:=TimeToStr(FTimeEditor.Time);
    ekDateTime:  Result:=DateTimeToStr(Trunc(FDateEditor.Date)+Frac(FTimeEditor.Time));
  else
    Result:='';
  end;
end;

function TMultiHeaderDBGrid.CommitEditorValue(ACol, ARow: Integer): Boolean;
// Writes the typed editor value into the bound field, with a try/except guard
// that reverts the dataset edit on a conversion/validation failure.
var
  DS: TDataSet;
  Field: TField;

  procedure ApplyToCurrentRecord;
  begin
    if not (DS.State in [dsEdit, dsInsert]) then DS.Edit;
    try
      case EditorKindForField(Field) of
        ekNumber:
          // Assign through the matching typed setter so the field keeps its
          // native precision (integer vs float).
          if FieldIsInteger(Field) then
            Field.AsLargeInt:=Trunc(FNumberEditor.Value)
          else
            Field.AsFloat:=FNumberEditor.Value;
        ekComboBox:
          if FComboEditor.ItemIndex>=0 then
            Field.AsString:=FComboEditor.Selected.Text;
        ekDate:
          // Pure ftDate -> midnight. A datetime field shown date-only keeps
          // its existing time component (only the date part is replaced).
          if Field.DataType in [ftDateTime, ftTimeStamp, ftOraTimeStamp] then begin
            var OldTime: TDateTime;
            if Field.IsNull then OldTime:=0 else OldTime:=Frac(Field.AsDateTime);
            Field.AsDateTime:=Trunc(FDateEditor.Date)+OldTime;
          end else
            Field.AsDateTime:=Trunc(FDateEditor.Date);
        ekTime:
          Field.AsDateTime:=Frac(FTimeEditor.Time);
        ekDateTime:
          Field.AsDateTime:=Trunc(FDateEditor.Date)+Frac(FTimeEditor.Time);
      else
        // ekMemo and any text path: round-trip via AsString.
        Field.AsString:=GetEditorText;
      end;
      DS.Post;
    except
      DS.Cancel;
      raise;
    end;
  end;

begin
  Result:=False;
  Field:=FEditField;
  if Field=nil then Exit;

  DS:=DataSet;
  if (DS=nil) or (not DS.Active) then Exit;
  if Field.ReadOnly then Exit;

  try
    if ARow=DS.RecNo-1 then
      ApplyToCurrentRecord
    else begin
      DS.DisableControls;
      try
        var Bookmark:=DS.Bookmark;
        try
          DS.RecNo:=ARow+1;
          ApplyToCurrentRecord;
        finally
          if DS.BookmarkValid(Bookmark) then
            DS.Bookmark:=Bookmark;
        end;
      finally
        DS.EnableControls;
      end;
    end;
    Result:=True;
  except
    // Conversion/validation failed: keep the old value, swallow so the grid
    // simply stays put. (Change to `raise` to surface the error instead.)
    Result:=False;
  end;

  // The cached row text is now stale - drop it so DoGetCellText re-reads.
  FCellTexts.Remove(ARow);
  FEditField:=nil;
  if FTimeEditor<>nil then FTimeEditor.Visible:=False;
  Invalidate;
end;

procedure TMultiHeaderDBGrid.CancelEditing;
begin
  inherited;
  FEditField:=nil;
  if FTimeEditor<>nil then FTimeEditor.Visible:=False;
end;

procedure TMultiHeaderDBGrid.PositionEditor;
begin
  inherited; // positions the active editor (date control for the composite)

  // For the ftDateTime composite, lay the time editor immediately to the
  // right of the date editor, splitting the cell width between them.
  if FEditing and (FEditField<>nil) and
     (EditorKindForField(FEditField)=ekDateTime) and
     (FDateEditor<>nil) and (FTimeEditor<>nil) and FDateEditor.Visible then begin
    // Split the cell between the two controls, biased so the date editor is a
    // little wider than the time editor. At the composite's minimum width
    // (MINW_DATETIME=156) this yields date=82, time=74 (date +5 / time -3
    // versus the previous even 77/77 split).
    var Total:=FDateEditor.Width;
    var DateW:=Total/2+4;
    var TimeW:=Total-DateW;        // = Total/2 - 4
    FDateEditor.Width:=DateW;
    FTimeEditor.SetBounds(FDateEditor.Position.X+DateW, FDateEditor.Position.Y,
                          TimeW, FDateEditor.Height);
    FTimeEditor.Visible:=True;
    FTimeEditor.BringToFront;
  end;
end;

procedure TMultiHeaderDBGrid.SetDataSource(const Value: TDataSource);
begin
  if FDataLink.DataSource<>Value then begin
    FDataLink.DataSource:=Value;
    // Treat the assignment like an active-state change: if the new source's
    // DataSet is already open, this auto-creates columns for it (when empty)
    // via the same new-source/table check used by ActiveChanged.
    HandleActiveChanged;
  end;
end;

procedure TMultiHeaderDBGrid.SetRowCacheSize(const Value: integer);
begin
  if FRowCacheSize<>Value then begin
    FRowCacheSize:=Value;

    CleanUpCache;
  end;
end;

end.


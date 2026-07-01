unit FMX.MultiHeaderGrid;

interface

uses
  System.Classes, System.Types, System.UITypes, System.Generics.Collections,
  FMX.Types, FMX.Controls, FMX.Graphics, FMX.StdCtrls, FMX.Objects, FMX.Layouts,
  FMX.Memo, FMX.ListBox, FMX.DateTimeCtrls, FMX.DateTimeCtrls.Types, FMX.Edit, FMX.EditBox,
  FMX.NumberBox, FMX.Text, Data.DB;

type

  THeaderLevel = class;
  THeaderLevels = class;

  // Kind of inplace editor the DB grid uses for a column, selected from the
  // bound field's DataType / FieldKind. See TMultiHeaderDBGrid.EditorKindForField.
  TMHGEditorKind = (
    ekNone,      // not editable (binary/structured/calculated/read-only)
    ekMemo,      // TMemo - text fields, value round-trips via Field.AsString
    ekNumber,    // TEdit - numeric text (holds '' for SQL NULL)
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
  // Fired before a row is inserted/appended/deleted via the keyboard shortcuts
  // (Insert / Down-on-last-row / Ctrl+Del). Set Allow:=False to veto the change.
  TRowModifyEvent = procedure(Sender: TObject; ARow: Integer; var Allow: Boolean) of object;

  TResizeMode = (rmNone, rmColumn, rmHeaderRow, rmGridRow);
  TResizeQuality = (rqNoChange, rqPrecise, rqFast);
  TScrollShowMode = (smAuto, smShow, smHide);

  TMultiHeaderGrid = class(TControl)
  private
    type
      TRowData = packed record
        Top       : integer;
        Height    : Word;
        AutoSized : boolean;
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

      // VCL-TDBGrid-like column focus events. FLastColEvent holds the column
      // index OnColEnter last fired for, so OnColExit/OnColEnter fire once per
      // genuine column change regardless of the path (mouse, keyboard or
      // programmatic Col:=). -1 means "no column entered yet".
      FOnColEnter: TNotifyEvent;
      FOnColExit: TNotifyEvent;
      FLastColEvent: Integer;

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
      FOnInsertRow: TRowModifyEvent;
      FOnAppendRow: TRowModifyEvent;
      FOnDeleteRow: TRowModifyEvent;
      FLastClickIsOnCell: boolean;
      FDrawRect : TRectF;

      FRowSelect: Boolean;
      FWordWrap: Boolean;
      FHeaderWordWrap: Boolean;
      // When word wrap is on, AutoSizeCols first sizes wrapped columns down to
      // their widest-word width (so text wraps). With this on (the default), a
      // follow-up pass reconciles the total against the viewport: leftover
      // horizontal space is handed back to wrapped columns (growing them toward
      // their natural one-line width, so they wrap less), and an overflow shrinks
      // columns proportionally toward their word-width floor to fit. When off,
      // columns are sized to word width regardless of available space.
      FConservativeWrap: Boolean;
      FHeaderColumns: TMHGHeaderColumns;
      FRebuildingHeaderColumns: Boolean;
      // Raised by bulk operations (header/column rebuilds) that apply word-wrap
      // to many columns at once, so the per-setter row re-fit doesn't run once
      // per column. The caller does a single sizing pass when finished.
      FSuppressAutoSize: Boolean;
      // True only while AutoSizeRows runs. Row-provider descendants check this
      // and skip extending RowCount during the measurement sweep (it iterates
      // every row).
      FInLayout: Boolean;
      FGridCellsHasWordWrap: boolean;
      // Precision of the last explicit AutoSize the user requested. Internal
      // re-sizes (the inplace editor's row re-fit, commit re-fit, the visible-
      // rows pass) reuse this so editing matches whatever mode the grid was last
      // sized in - keeping fast grids fast and precise grids pixel-consistent.
      FAutoSizePrecise: Boolean;
      FVerticalScroll: TScrollShowMode;
      FHorisontalScroll: TScrollShowMode;

      // Inplace cell editor (basic grids). A single reusable TMemo is moved
      // over the cell being edited.
      FEditor: TMemo;
      FEditorHost: TLayout;
      // Currently active inplace editor control. For the base/string grids
      // this is always FEditor (the TMemo). The DB grid may swap in a typed
      // control (TComboBox/TDateEdit/TTimeEdit/...) per field type. (Boolean
      // fields use no editor - they toggle in place.)
      FActiveEditor: TControl;
      FEditing: Boolean;
      FEditCol: Integer;
      FEditRow: Integer;
      // Anchor column widened for the editor and its width before editing
      // started, so a transient widen can be shrunk back (but not below the
      // original) when editing ends. FEditWidenAnchor<0 means "none".
      FEditWidenAnchor: Integer;
      FEditWidenStartW: Integer;
      // When True a column transiently widened to fit the inplace editor keeps
      // its enlarged width when editing ends; when False (default) it is fully
      // restored to its pre-edit width.
      FKeepEditorWidenedColumn: Boolean;
      FReadOnly: Boolean;
      FPainted: boolean;
      // Set while the base handles vkEnd so the shared down-navigation skips the
      // synchronous on-demand grow; the DB descendant pages to Eof separately.
      FEndJump: boolean;
      FMaxColumnAutoWidth: integer;

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
    // Effective word-wrap for a data cell, matching the cell-draw logic
    // (per-cell style override, else grid/column flag).
    function EffectiveCellWordWrap(ACol, ARow: Integer): Boolean;
    // Effective cell style for a data cell, matching the cell-draw logic
    // (per-cell style override, else grid/column flag).
    function EffectiveCellStyle(ACol, ARow: Integer): TCellStyle;
    // Raw height (before padding/gridline) the cell's text needs. Single source
    // of truth shared by AutoSizeRows (committed text) and the inplace editor
    // (uncommitted FEditor.Text) so the two never disagree by a fractional
    // pixel. AText is the text to measure; RowSpan divides a merged cell's
    // height across its rows.
    function MeasureCellTextHeight(ACol, ARow: Integer; const AText: string;
      RowSpan: Integer): Single;
    // Width available to a header element's text (merged rect, padded).
    function HeaderElementTextWidth(ALevel, ACol: Integer): Single;

    procedure StartCellEditing(ACol, ARow: Integer; InitialChar: Char);
    // Widens the editing cell's (anchor) column if EditorMinColWidth needs more
    // room than it currently has. Called at edit start and re-callable while a
    // non-wrapping editor's content grows (e.g. typing a long number).
    procedure WidenColForEditor(ACol, ARow: Integer); virtual;
    // Called when editing ends: shrinks a transiently-widened anchor column to
    // ATargetW (the final value's needed width), but never below its pre-edit
    // width. ATargetW<0 restores the pre-edit width (cancel).
    procedure FinishEditColWidth(ATargetW: Integer);
    procedure EnsureEditor(TextAlign: TTextAlign);
    // Parents a control into FEditorHost and wires the shared key/exit
    // handlers, so the memo and the DB grid's typed editors all sit in the host
    // over the opaque backing, clipped to the cell.
    procedure WireEditor(C: TControl);
    procedure PositionEditor; virtual;
    // Lays the active editor out inside the already-positioned host, given the
    // host interior size, and shows it. Base fills the host with the
    // Client-aligned TMemo; the DB grid overrides for its typed editors.
    procedure LayoutActiveEditor(AWidth, AHeight: Single); virtual;
    // Takes down the editor UI: hides the host (and its FEditorBack child) and
    // the active editor. Descendants override to hide their extra editors too.
    procedure HideEditor; virtual;
    procedure CommitEditing;
    procedure CancelEditing; virtual;
    procedure EditorKeyDown(Sender: TObject; var Key: Word; var KeyChar: Char;
                            Shift: TShiftState);
    procedure EditorExit(Sender: TObject); virtual;

    // Grows (or shrinks) the row being edited so it fits the editor's current text,
    // then re-lays the editor out within the resized cell. Fires on every keystroke
    // via FEditor.OnChange while the cell content is still uncommitted, so we
    // measure the live editor rather than the stored cell text.
    procedure MemoEditorTextChanged(Sender: TObject);
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
    // CanEditCell: an inplace editor may OPEN for this cell (by type/bounds).
    // It is independent of ReadOnly so a read-only grid still opens the editor
    // for selecting/copying text. CellIsModifiable gates actual writes: it is
    // False when ReadOnly (and, in the DB grid, when the dataset can't modify),
    // and drives the editor's own ReadOnly state plus the row-edit shortcuts.
    function  CanEditCell(ACol, ARow: Integer): Boolean; virtual;
    function  CellIsModifiable(ACol, ARow: Integer): Boolean; virtual;
    // Applies CellIsModifiable to the active editor control so a read-only cell
    // opens an editor that allows selection/copy but not changes.
    procedure ApplyEditorReadOnly(Ed: TControl; AModifiable: Boolean); virtual;
    // Keyboard row shortcuts. Insert inserts a blank row before ARow, AppendRow
    // adds one after the last row (Down on the last row), DeleteRow removes ARow.
    // All are no-ops when the grid/dataset is not modifiable. The base grid acts
    // on the string-grid row model; the DB grid overrides to drive the dataset.
    procedure InsertRow(ARow: Integer); virtual;
    procedure AppendRow; virtual;
    // Handles a plain Down keypress. Returns True if it fully handled the key
    // (so KeyDown does no further navigation). The base appends when Down is
    // pressed on the last row of a modifiable grid; the DB grid overrides to
    // follow VCL TDBGrid's NextRow semantics (append at Eof, cancel an
    // unmodified pending insert). Returns False to fall through to plain move.
    function  HandleDownKey: Boolean; virtual;
    procedure DeleteRow(ARow: Integer); virtual;
    // Fire the veto events; return True when the change may proceed.
    function  DoInsertRow(ARow: Integer): Boolean;
    function  DoAppendRow(ARow: Integer): Boolean;
    function  DoDeleteRow(ARow: Integer): Boolean;
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
    procedure SetViewTop(const Value: Integer); virtual;
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
    procedure SetConservativeWrap(const Value: Boolean);
    // ConservativeWrap helper: after AutoSizeCols has sized wrapped columns to
    // their word width, redistribute against the viewport - give spare width
    // back to wrapped columns (up to their natural one-line width), or shrink
    // them proportionally toward their word floor when the total overflows.
    procedure ReconcileWrappedColumns(const AIsWrapped: TArray<Boolean>;
                                      const AWordFloor, ANaturalW: TArray<Single>;
                                      AViewportW: Integer);
    // After the header heights are known, find columns whose caption wrapped onto
    // substantially more lines than the typical column (average line count rounded
    // up) and widen them just enough to bring them down toward typical height -
    // not all the way to one line (which would leave blank header lines). Capped
    // by the column's MaxWidth / MaxColumnAutoWidth; if the total exceeds the
    // viewport the grid scrolls horizontally. Active whenever HeaderWordWrap is
    // on. Re-fits header heights itself. Returns True if any width changed.
    function BalanceHeaderColumnWidths: Boolean;
    // Number of wrapped lines the caption of the element governing (ALevel,ACol)
    // needs at the column's current merged width (1 when not wrapping).
    function HeaderElementLineCount(ALevel, ACol: Integer): Integer;
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

    // Pre-paint hook. The base does nothing; descendants (the DB grid) use it
    // to flush a deferred, coalesced rebuild exactly once before drawing, so a
    // burst of column-property changes triggers a single recompute instead of
    // one per change. Call this before any operation that must observe an
    // up-to-date layout (paint, size queries, autosize).
    procedure EnsureLayout; virtual;

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
    // Called from keyboard navigation before moving down, so the target row can
    // be provided and the move is not capped at the current RowCount.
    procedure EnsureRowAvailable(ARow: Integer); virtual;
    // True only while AutoSizeRows runs (read by row-provider descendants).
    function  InLayout: Boolean;
    // Upper row bound (exclusive) for the AutoSizeCols content scan. A base grid
    // scans all rows; a row-provider descendant overrides this to the rows it has
    // available so auto-fit does not force the whole set to be provided.
    function  ColScanRowLimit: Integer; virtual;
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

    procedure AutoSizeCols;

    procedure AutoSizeRows(FromRow: integer = 0; ToRow: integer = -1;
                           FResizeStartColumnIndex: integer = -1; FResizeEndColumnIndex: integer = -1;
                           TryOptimise: boolean = False);

    procedure AutoSizeVisibleRows(FResizeStartColumnIndex: integer = -1; FResizeEndColumnIndex: integer = -1);
    // Sizes each header level's height to fit its (optionally wrapped or
    // multi-line) captions. Shared by all grid descendants.
    procedure AutoSizeHeaders;
    procedure AutoSize(Precision: TResizeQuality = rqNoChange);

    procedure ClearSelection;

    property ColLefts[Index: Integer]: Integer read GetColLeft;
    property ColWidths[Index: Integer]: Integer read GetColWidth write SetColWidth;
    property ColTextHAlignment[Index: Integer]: TTextAlign read GetColTextHAlignment write SetColTextHAlignment;
    property ColTextVAlignment[Index: Integer]: TTextAlign read GetColTextVAlignment write SetColTextVAlignment;
    property ColMinWidth[Index: Integer]: integer read GetColMinWidth write SetColMinWidth;
    property ColMaxWidth[Index: Integer]: integer read GetColMaxWidth write SetColMaxWidth;
    property ColWordWrap[Index: Integer]: Boolean read GetColWordWrap write SetColWordWrap;
    function GetCellRect(ACol, ARow: Integer): TRectF; overload;
    function GetCellRect(ACol, ARow: Integer; out MergedCell: TMergedCell): TRectF; overload;
    function GetMergedCellRect(ACol, ARow: Integer): TRectF;
    function GetHeaderRect(ALevel, ACol: Integer): TRectF;

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
    // --- Layer 1: events the grid does not intercept (fire natively) ---
    property OnEnter;
    property OnExit;
    property OnMouseEnter;
    property OnMouseLeave;
    property OnMouseWheel;
    property OnGesture;
    property OnDragEnter;
    property OnDragLeave;
    property OnDragOver;
    property OnDragDrop;
    property OnDragEnd;
    // --- Layer 2: events the grid intercepts; dispatch ensured in the
    //     overridden MouseMove / KeyDown (inherited now called) ---
    property OnMouseDown;
    property OnMouseMove;
    property OnMouseUp;
    property OnKeyDown;
    property OnKeyUp;

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
    // A column may be transiently widened to fit the inplace editor. When False
    // (default) the column returns to its pre-edit width when editing ends; when
    // True it keeps the enlarged width.
    property KeepEditorWidenedColumn: Boolean read FKeepEditorWidenedColumn write FKeepEditorWidenedColumn default False;
    property WordWrap: Boolean read FWordWrap write SetWordWrap default False;
    // When word wrap is on, controls whether AutoSizeCols reconciles wrapped
    // column widths against the available viewport: grows wrapped columns back
    // toward their one-line width when there's spare horizontal space (less
    // wrapping), and shrinks proportionally toward word width on overflow.
    // On by default; turn off to size wrapped columns purely to word width.
    property ConservativeWrap: Boolean read FConservativeWrap write SetConservativeWrap default True;
    // When set, header captions wrap to the cell width during drawing and
    // are accounted for by AutoSizeHeaders. On by default.
    property HeaderWordWrap: Boolean read FHeaderWordWrap write SetHeaderWordWrap default True;
    // Maximum width of the column for autowidth computing
    // Can be overriden by MaxWidth property of  the column.
    property MaxColumnAutoWidth: integer read FMaxColumnAutoWidth write FMaxColumnAutoWidth default 400;
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
    // VCL TDBGrid-like column focus events (fire on a genuine column change
    // via mouse, keyboard or programmatic Col:=).
    property OnColEnter: TNotifyEvent read FOnColEnter write FOnColEnter;
    property OnColExit: TNotifyEvent read FOnColExit write FOnColExit;

    property ResizeEnabled: Boolean read FResizeEnabled write SetResizeEnabled default True;
    property ResizeRowEnabled: Boolean read FResizeRowEnabled write FResizeRowEnabled default True;
    property ResizeColEnabled: Boolean read FResizeColEnabled write FResizeColEnabled default True;
    property ResizeHeaderRowEnabled: Boolean read FResizeHeaderRowEnabled write FResizeHeaderRowEnabled default True;

    property ResizeMargin: Integer read GetResizeMargin write SetResizeMargin default 2;

    property OnColumnResized: TColumnsResizedEvent read FOnColumnResized write FOnColumnResized;
    property OnHeaderResized: TRowResizedEvent read FOnHeaderResized write FOnHeaderResized;
    property OnRowResized: TRowResizedEvent read FOnRowResized write FOnRowResized;
    property OnGridScroll: TGridScrollEvent read FOnGridScroll write FOnGridScroll;
    // Fired (with a veto flag) before the keyboard row shortcuts modify rows.
    property OnInsertRow: TRowModifyEvent read FOnInsertRow write FOnInsertRow;
    property OnAppendRow: TRowModifyEvent read FOnAppendRow write FOnAppendRow;
    property OnDeleteRow: TRowModifyEvent read FOnDeleteRow write FOnDeleteRow;
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
    FDisplayFormat : string;

    function GetGrid: TMultiHeaderDBGrid;
    procedure SetFieldName(const Value: string);
    procedure SetColor(const Value: TAlphaColor);
    procedure SetDisplayFormat(const Value: string);
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
    // Per-column display format for the field -> cell-string conversion. When
    // empty (the default), date/time/datetime fields use the grid's DateFormat
    // / TimeFormat and all other fields use Field.AsString. When set, it is
    // passed verbatim to FormatDateTime for date/time fields (e.g. 'YYYY-MM-DD',
    // 'HH:NN'); ignored for non-datetime fields.
    property DisplayFormat: string read FDisplayFormat write SetDisplayFormat;
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
      // When True (default), DeleteRow shows a confirmation prompt before
      // deleting - mirrors VCL TDBGrid's dgConfirmDelete option.
      FConfirmDelete: Boolean;
      FUpdatingRow: Boolean; // guards against recursion when syncing Row <-> DataSet.RecNo
      // Grid row of the pending insert record during dsInsert (VCL buffer-style
      // phantom row); -1 otherwise.
      FInsertRowIndex: Integer;
      // --- On-demand rows: native RecNo/RecordCount + live growth --------
      // The grid always uses the dataset's native RecNo/RecordCount. When the
      // dataset fetches on demand (FireDAC FetchOptions.Mode is fmManual/
      // fmOnDemand), RecordCount only counts rows fetched so far, so the cursor
      // is stepped with Next (bounded by Eof) as the grid navigates/scrolls to
      // materialize more rows. Fully-materialized cursors need no growth.
      FFetchesOnDemand: Boolean;            // cached: dataset fetches rows lazily (FetchOptions.Mode)
      FInUpdateRowCount: Boolean;           // re-entrancy guard for UpdateRowCount
      // Re-entrancy guard for native-path live growth (stepping the cursor with
      // Next fires DataSetScrolled/-Changed -> UpdateRowCount).
      FNativeFetching: Boolean;
      // Set once a native-path fetch has reached Eof (the whole set is loaded and
      // RowCount is final). Gates the native grow paths so they stop re-fetching
      // and re-positioning the cursor - which would fire DataSetScrolled and move
      // the selection (the "End needs two presses" bug).
      FNativeAllFetched: Boolean;
      // End-key on-demand fetch state. FEndFetching is set while End is paging to
      // Eof; any real key or mouse click sets FEndFetchCancelled to stop it.
      // FEndSynthetic marks the loop's own Page Down presses so they are not
      // treated as user input that cancels.
      FEndFetching: Boolean;
      FEndFetchCancelled: Boolean;
      FEndSynthetic: Boolean;
      FColumns: TMHGColumns;
      // Grid-wide default formats for the field -> cell-string conversion of
      // date / time / datetime fields, used when a column has no explicit
      // DisplayFormat. Passed to FormatDateTime. Defaults: 'DD.MM.YYYY' and
      // 'HH:NN:SS'; datetime fields use DateFormat + ' ' + TimeFormat.
      FDateFormat: string;
      FTimeFormat: string;
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
      // When True, a rebuild has been requested but deferred. It is coalesced
      // and flushed once by EnsureLayout, just before the next paint (or before
      // any synchronous query that must see an up-to-date layout). This means a
      // burst of column-property changes costs a single ResetTable, not one per
      // change.
      FRebuildPending: Boolean;
      // Effective column-index -> collection item, rebuilt by ResetTable.
      // Lets DoGetCellStyle read a column's Color in O(1) per cell.
      FColMap: TArray<TMHGColumn>;

      // --- Typed inplace editor controls --------------------------------
      // Created lazily, reused across edits. Only one is active at a time
      // (selected per field type); the active one is FActiveEditor.
      // Boolean fields use no editor control - they toggle in place
      // (see CellIsToggle / ToggleCell).
      FComboEditor: TComboBox;
      FNumberEditor: TEdit; // plain edit: holds '' for SQL NULL (a TNumberBox
                            // cannot represent an empty/null value)
      FDateEditor : TDateEdit;
      FTimeEditor : TTimeEdit;
      // For ftDateTime the date and time controls share the cell side by side.
      FEditField  : TField; // field bound to the editor currently open
      // Natural single-line height of the typed editors, captured at creation
      // before any cell layout stretches them.
      FTypedEditorHeight: Single;
      // Transient extra width (clear-button footprint) added to a date/time
      // editor's measured min column width during WidenColForEditor only.
      FEditorExtraWidth: Single;

    function GetDataSource: TDataSource;
    procedure SetDataSource(const Value: TDataSource);
    procedure SetRowCacheSize(const Value: integer);
    procedure SetColumns(const Value: TMHGColumns);
    procedure SetDateFormat(const Value: string);
    procedure SetTimeFormat(const Value: string);
    // Converts a field's value to its displayed cell string, applying the
    // per-column DisplayFormat or the grid DateFormat/TimeFormat for
    // date/time/datetime fields; falls back to Field.AsString otherwise.
    function FieldToText(ACol: TMHGColumn; AField: TField): string;
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
    // --- On-demand row helpers (native RecNo + live growth) ------------
    // True when the dataset fetches rows on demand, so RecordCount only counts
    // rows fetched so far and the true end is unknown until stepped. Detected via
    // the FireDAC FetchOptions.Mode RTTI property. False when not detectable.
    function  DataSetFetchesOnDemand(DS: TDataSet): Boolean;
    // True when the End key must discover the end by fetching on demand (an
    // on-demand cursor not yet fully fetched). False when the end is already
    // known, so End jumps to it.
    function  EndIsOnDemand: Boolean;
    // Runs the cancellable, repainting page-down loop driving the dataset to Eof
    // (or until the user cancels). Invoked from KeyDown on a real End press.
    procedure RunEndFetch;
    // 0-based index of the current record: RecNo-1, or the pending insert row
    // during dsInsert. -1 if unknown.
    function  CurrentRowIndex: Integer;
    // Native-path live growth: step the cursor forward (bounded by Eof) until
    // grid row ATargetRow is materialized, or to Eof when ATargetRow < 0, then
    // adopt the grown RecordCount. Shared by EnsureRowAvailable /
    // EnsureRowsForViewport and the End-key fetch. No-op once fully loaded.
    procedure GrowNativeTo(ATargetRow: Integer);
  protected
    procedure DoGetCellText(ACol, ARow: Integer; var Text: string); override;
    procedure DoSetCellText(ACol, ARow: Integer; const Text: string); override;
    procedure DoGetCellStyle(ACol, ARow: Integer; var Style: TCellStyle); override;
    procedure DoSetCellStyle(ACol, ARow: Integer; const Style: TCellStyle); override;
    procedure DoSelectCell; override;
    // Row-provider overrides: the DB grid supplies rows on demand from the
    // dataset, growing RowCount as the grid navigates/scrolls/scans.
    procedure DoGridScroll; override;
    procedure EnsureRowsForViewport(ATargetTop: Integer);
    procedure EnsureRowAvailable(ARow: Integer); override;
    function  ColScanRowLimit: Integer; override;
    // End-key handling and on-demand End-fetch cancellation live here (not in the
    // base, which is dataset-agnostic): End on an on-demand cursor pages to Eof,
    // cancellable by any other key or a mouse click.
    procedure KeyDown(var Key: Word; var KeyChar: WideChar; Shift: TShiftState); override;
    procedure MouseDown(Button: TMouseButton; Shift: TShiftState; X, Y: Single); override;
    // Data-bound inplace editing: a typed editor is chosen per field.
    function CanEditCell(ACol, ARow: Integer): Boolean; override;
    // A DB cell is modifiable only when not ReadOnly, the dataset CanModify,
    // and the bound field is not read-only.
    function CellIsModifiable(ACol, ARow: Integer): Boolean; override;
    // Row shortcuts driven through the dataset (Insert/Append/Delete).
    procedure InsertRow(ARow: Integer); override;
    procedure AppendRow; override;
    function  HandleDownKey: Boolean; override;
    procedure DeleteRow(ARow: Integer); override;
    // Extends the base to cover the typed date/time editors (not TCustomEdit
    // descendants) so a read-only date/time cell opens for copy without edits.
    procedure ApplyEditorReadOnly(Ed: TControl; AModifiable: Boolean); override;
    // Boolean fields toggle in place instead of opening an editor.
    function CellIsToggle(ACol, ARow: Integer): Boolean; override;
    function ToggleCell(ACol, ARow: Integer): Boolean; override;
    // --- Typed editor overrides (see base hooks) ------------------------
    function  PrepareCellEditor(ACol, ARow: Integer): TControl; override;
    procedure LayoutActiveEditor(AWidth, AHeight: Single); override;
    // Self-contained placement of the shared TMemo for ekMemo cells. A private
    // copy of the base logic (the DB grid does NOT call inherited for the memo),
    // so the DB and base/string grids can be fixed independently.
    procedure LayoutMemoEditor(AWidth, AHeight: Single);
    procedure HideEditor; override;
    // Footprint of the styled clear button (its Width plus left/right margins),
    // read from the captured 'clearbutton' style instance so column widening tracks
    // whatever the style defines. The date control's button serves ekDate and the
    // ekDateTime composite; the time control's button serves ekTime. Returns 0 when
    // the style hasn't resolved yet (first frame) or no button applies - the column
    // re-measures on the next layout once the instance is captured.
    function ClearButtonWidth(Field: TField): single;
    procedure CancelEditing; override;
    procedure EditorExit(Sender: TObject); override;
    procedure OnDateEditorClear(Sender: TObject);
    procedure OnTimeEditorClear(Sender: TObject);
    procedure LoadEditorValue(ACol, ARow: Integer; InitialChar: Char); override;
    function  GetEditorText: string; override;
    function  CommitEditorValue(ACol, ARow: Integer): Boolean; override;
    procedure PlaceEditorCaretAtEnd(Ed: TControl); override;
    procedure SetViewTop(const Value: Integer); override;
    function  EditorMinColWidth(ACol, ARow: Integer): single; override;
    // The clear button on nullable TDateEdit/TTimeEdit editors takes extra
    // horizontal room that the base EditorMinColWidth (via MINW_*) doesn't
    // account for. Widen the column by that footprint so the button isn't
    // clipped. (WidenColForEditor only runs in the base, so override here.)
    procedure WidenColForEditor(ACol, ARow: Integer); override;
    // Field bound to a given grid column (nil if out of range / no dataset).
    function  FieldForCol(ACol: Integer): TField;
    // Editor kind for a field, from DataType / FieldKind / ReadOnly.
    function  EditorKindForField(Field: TField): TMHGEditorKind; overload;
    // As above but, when AIgnoreReadOnly is True, returns the type-based editor
    // kind even for a read-only field, so a read-only cell can still open that
    // editor in read-only mode for selecting/copying text.
    function  EditorKindForField(Field: TField;
                                 AIgnoreReadOnly: Boolean): TMHGEditorKind; overload;
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
    // Hides every typed editor except Keep; KeepComposite preserves the
    // ftDateTime date+time pair. Used by LayoutActiveEditor and HideEditor.
    procedure HideTypedEditors(Keep: TControl; KeepComposite: Boolean);
    // OnKeyDown for the number editor: rejects characters that can't form a
    // valid number, so the TEdit only ever holds a numeric string (or '').
    procedure NumberEditorKeyDown(Sender: TObject; var Key: Word;
      var KeyChar: Char; Shift: TShiftState);
    // OnKeyDown for the ftDateTime composite date/time controls. Lets the caret
    // cross from the date editor's last position to the time editor's first
    // (right arrow) and back (left arrow), while still routing Enter/Tab/Esc
    // through the shared EditorKeyDown.
    procedure DateTimeEditorKeyDown(Sender: TObject; var Key: Word;
      var KeyChar: Char; Shift: TShiftState);
    // True when C is one of the two controls of the active ftDateTime composite.
    function  IsCompositeEditor(C: TObject): Boolean;
    // True when the date/time control's focused format part is the last/first
    // section (e.g. the year of dd.MM.yyyy), used to decide arrow-key hand-off.
    function  DateTimePartIsLast(Ed: TCustomDateTimeEdit): Boolean;
    function  DateTimePartIsFirst(Ed: TCustomDateTimeEdit): Boolean;
    // Moves the focused format part of a date/time control to its first or last
    // section without changing which control has focus. Used both for arrow
    // hand-off and to re-init the editors when an edit starts on a new cell.
    procedure DateTimeGoToFirstPart(Ed: TCustomDateTimeEdit);
    procedure DateTimeGoToLastPart(Ed: TCustomDateTimeEdit);
    // Resets caret/selection/format-part of every typed editor to a known state
    // so a fresh edit never inherits the previous cell's cursor position. Called
    // from PlaceEditorCaretAtEnd at the start of each edit.
    procedure ResetTypedEditorCursors;
    // Moves focus to the composite's other control, placing the caret at its
    // start (ToTime=True) or end (ToTime=False).
    procedure FocusCompositePart(ToTime: Boolean);
    // OnChangeTracking for the number editor: widens the column as the value
    // grows (numbers don't wrap), so the editor never clips while typing.
    procedure NumberEditorChangeTracking(Sender: TObject);
  public
    constructor Create(AOwner: TComponent); override;
    destructor Destroy; override;

    // Full grid rebuild: column headers (based on the DataSet's visible fields)
    // and the row count. Called automatically when DataSource changes,
    // the dataset is opened/closed, or the field list changes.
    procedure ResetTable;

    // Requests a rebuild but defers the work: the actual ResetTable runs once,
    // coalesced, on the next EnsureLayout (i.e. just before paint). Prefer this
    // over ResetTable for change notifications, so N property changes in a row
    // cost one rebuild instead of N. Use ResetTable directly only when an
    // up-to-date layout is needed synchronously right now.
    procedure InvalidateTable;

    // Flushes a pending deferred rebuild (see InvalidateTable). Called by Paint
    // and by synchronous layout queries; safe to call when nothing is pending.
    procedure EnsureLayout; override;

    // Recalculates RowCount from DataSet.RecordCount, clears the row cache
    // and syncs the selected row with the current DataSet record.
    // Called automatically when the record set changes (filter,
    // Refresh, record insertion/deletion, etc.).
    procedure UpdateRowCount;

    // Drops the cached cell strings (so they are re-rendered from the DataSet,
    // e.g. after a display-format change) and repaints. Does not rebuild the
    // header or change row count.
    procedure RefreshCells;

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
    // When True (default) DeleteRow asks for confirmation before deleting the
    // record, like VCL TDBGrid's dgConfirmDelete option.
    property ConfirmDelete: Boolean read FConfirmDelete write FConfirmDelete default True;
    // Default display format for date and time field -> cell-string conversion,
    // used when a column has no explicit DisplayFormat. Datetime fields combine
    // them as DateFormat + ' ' + TimeFormat (or DateFormat alone when the
    // column's DateTimeEditor is dteDate). Defaults: 'DD.MM.YYYY' / 'HH:NN:SS'.
    property DateFormat: string read FDateFormat write SetDateFormat;
    property TimeFormat: string read FTimeFormat write SetTimeFormat;
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
  System.SysUtils, System.Math, System.Rtti, FMX.Platform, FMX.Forms, System.StrUtils, FMX.Styles, FMX.Styles.Objects, FMX.DialogService.Sync;

// Forward declarations for unit-local text helpers used by methods that
// appear earlier in the implementation than the functions themselves.
function CountLines(const Text: string): Integer; forward;

var
  MemoStyle   : TLayout;
  NumberStyle : TLayout;
  ComboStyle  : TLayout;
  DateStyle   : TLayout;
  TimeStyle   : TLayout;

const
  // Usable minimum widths (px) for the fixed-size typed editors. The memo and
  // checkbox are comfortable in a narrow column, so they keep the default 0.
  MINW_COMBO    = 120;
  MINW_DATE     = 74;  // 'dd.MM.yyyy'
  MINW_TIME     = 64;  // 'HH:nn:ss' plus picker button
  MINW_DATETIME = MINW_DATE+MINW_TIME; // date + time composite

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
begin
  if FTitle <> Value then begin
    FTitle := Value;
    Changed;
  end;
end;

procedure TMHGHeaderColumn.SetWidth(const Value: Integer);
begin
  if FWidth <> Value then begin
    FWidth := Value;
    Changed;
  end;
end;

procedure TMHGHeaderColumn.SetMinWidth(const Value: Integer);
begin
  if FMinWidth <> Value then begin
    FMinWidth := Value;
    Changed;
  end;
end;

procedure TMHGHeaderColumn.SetMaxWidth(const Value: Integer);
begin
  if FMaxWidth <> Value then begin
    FMaxWidth := Value;
    Changed;
  end;
end;

procedure TMHGHeaderColumn.SetVisible(const Value: Boolean);
begin
  if FVisible <> Value then begin
    FVisible := Value;
    Changed;
  end;
end;

procedure TMHGHeaderColumn.SetWordWrap(const Value: Boolean);
begin
  if FWordWrap <> Value then begin
    FWordWrap := Value;
    Changed;
  end;
end;

procedure TMHGHeaderColumn.SetAlignment(const Value: TTextAlign);
begin
  if FAlignment <> Value then begin
    FAlignment := Value;
    Changed;
  end;
end;

procedure TMHGHeaderColumn.SetVertAlignment(const Value: TTextAlign);
begin
  if FVertAlignment <> Value then begin
    FVertAlignment := Value;
    Changed;
  end;
end;

procedure TMHGHeaderColumn.SetHeaderAlignment(const Value: TTextAlign);
begin
  if FHeaderAlignment <> Value then begin
    FHeaderAlignment := Value;
    Changed;
  end;
end;

procedure TMHGHeaderColumn.SetHeaderVertAlignment(const Value: TTextAlign);
begin
  if FHeaderVertAlignment <> Value then begin
    FHeaderVertAlignment := Value;
    Changed;
  end;
end;

procedure TMHGHeaderColumn.SetGroupHeader(const Value: string);
begin
  if FGroupHeader <> Value then begin
    FGroupHeader := Value;
    Changed;
  end;
end;

procedure TMHGHeaderColumn.SetGroupHeaderSeparator(const Value: Char);
begin
  if FGroupHeaderSeparator <> Value then begin
    FGroupHeaderSeparator := Value;
    Changed;
  end;
end;

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

  FLastColEvent:=-1; // no column entered yet (OnColEnter/OnColExit guard)
  FEditWidenAnchor:=-1;
  FAutoSizePrecise:=False; // fast until the user requests a precise AutoSize
  FColCount:=5;
  FRowCount:=10;
  FDefaultColWidth:=80;
  FDefaultRowHeight:=20;
  FGridLines:=True;
  FHeaderWordWrap:=True;
  FConservativeWrap:=True;
  FHeaderColumns:=TMHGHeaderColumns.Create(Self);
  FGridLineColor:=TAlphaColors.Gray;
  FHeaderLineColor:=TAlphaColors.Black;
  FGridLineWidth:=1;
  FSelectedCell:=Point(-1, -1);
  FMaxColumnAutoWidth:=400;

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
    FColData[i].MaxWidth:=0;
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
  FCellPadding.Free;

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
      FColData[i].MaxWidth:=0;
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
    Result:=0;
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
    FColData[Index].Widths:=Value;
    if FColData[Index].MaxWidth>0 then begin
      FColData[Index].Widths:=Min(FColData[Index].Widths,FColData[Index].MaxWidth);
    end;
    if FColData[Index].MinWidth>0 then begin
      FColData[Index].Widths:=Max(FColData[Index].Widths,FColData[Index].MinWidth);
    end;
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
    FRowData[Index].Height:=Value;
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

procedure TMultiHeaderGrid.EnsureLayout;
begin
  // Base grids have no deferred layout; the DB grid overrides this.
end;

procedure TMultiHeaderGrid.Paint;
var
  Canvas: TCanvas;
begin
  // Flush any pending deferred rebuild so we draw an up-to-date layout.
  EnsureLayout;

  Canvas:=Self.Canvas;

  if Canvas.BeginScene then begin
    var Save:=Canvas.SaveState;
    try
      // Set the clipping region
      FDrawRect:=TRectF.Create(LocalRect.Left,LocalRect.Top,
                              LocalRect.Left+LocalRect.Width,LocalRect.Top+LocalRect.Height);
      if VScrollBar.Visible then FDrawRect.Right:=FDrawRect.Right-VScrollBar.Width;
      if HScrollPanel.Visible then FDrawRect.Bottom:=FDrawRect.Bottom-HScrollPanel.Height;
      Canvas.IntersectClipRect(FDrawRect);

      Canvas.Fill.Kind:=TBrushKind.Solid;
      Canvas.Fill.Color:=FBackgroundColor;

      if WordWrap then begin
        AutoSizeVisibleRows;
      end;

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
  FPainted:=True;
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

  if not FDrawRect.IntersectsWith(ARect) then Exit;

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

  var RowDataLen:=Length(FRowData);
  if RowDataLen=0 then Exit;
  if TopRow>=RowDataLen then Exit;
  var LastRow:=FRowCount-1;
  if LastRow>RowDataLen-1 then LastRow:=RowDataLen-1;
  for j:=TopRow to LastRow do begin
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

          if not FDrawRect.IntersectsWith(Rect) then Continue;

          DrawCell(Canvas, MergedCell.Col, MergedCell.Row, Rect);
        end;
        Continue;
      end;

      Rect:=GetCellRect(i, j);
      if not FDrawRect.IntersectsWith(Rect) then Continue;

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

  if FEditing and (FEditCol=ACol) and (FEditRow=ARow) then begin
    Text:='';
  end;

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

    // When vertically centered, fall back to top (Leading) alignment if the
    // text is taller than the available cell height - otherwise centering
    // clips the text symmetrically and hides its first lines. The measuring
    // font is the one already assigned above, so it matches what is drawn.
    if VAlignment=TTextAlign.Center then begin
      var TextH: Single;
      if WordWrapEnabled then begin
        var MeasureRect:=ARect;
        MeasureRect.Bottom:=MeasureRect.Top+100000;
        Canvas.MeasureText(MeasureRect, Text, True, [], TTextAlign.Leading, TTextAlign.Leading);
        TextH:=MeasureRect.Height;
      end else
        TextH:=Canvas.TextHeight(Text);
      if TextH>ARect.Height then
        VAlignment:=TTextAlign.Leading;
    end;

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
    if (X>=FDrawRect.Left) and (X<=FDrawRect.Right) then begin
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

procedure TMultiHeaderGrid.AutoSize(Precision: TResizeQuality = rqNoChange);
begin
  // Remember the mode the user asked for, so in-house re-sizes (editor re-fit,
  // commit re-fit, visible-rows pass) reuse it - fast stays fast, precise stays
  // pixel-consistent with the editor.

  case Precision of
    rqPrecise: FAutoSizePrecise:=True;
    rqFast: FAutoSizePrecise:=False;
  end;

  AutoSizeCols;
  AutoSizeRows;
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

  // Apply per-column geometry/alignment/word wrap. Suppress the per-setter row
  // re-fit while looping; a single AutoSize follows at the call site / below.
  FSuppressAutoSize:=True;
  try
    for var i:=0 to N-1 do begin
      if ACols[i].Width>0 then ColWidths[i]:=ACols[i].Width;
      if ACols[i].MinWidth>0 then ColMinWidth[i]:=ACols[i].MinWidth;
      if ACols[i].MaxWidth>0 then ColMaxWidth[i]:=ACols[i].MaxWidth;
      ColWordWrap[i]:=ACols[i].WordWrap;
      ColTextHAlignment[i]:=ACols[i].Alignment;
      ColTextVAlignment[i]:=ACols[i].VertAlignment;
    end;
  finally
    FSuppressAutoSize:=False;
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

function TMultiHeaderGrid.EffectiveCellWordWrap(ACol, ARow: Integer): Boolean;
// Mirrors the cell-draw decision: a per-cell style override wins, otherwise the
// grid-wide flag combined with the column's flag.
begin
  var Style:=CellStyle[ACol, ARow];
  if Style.WordWrapIsSet then
    Result:=Style.WordWrap
  else if (ACol>=0) and (ACol<FColCount) then
    Result:=FWordWrap or FColData[ACol].WordWrap
  else
    Result:=FWordWrap;
end;

function TMultiHeaderGrid.EffectiveCellStyle(ACol, ARow: Integer): TCellStyle;
// Resolves the style used to render a data cell into a fully-populated
// TCellStyle holding the concrete values DrawCell would have applied. Mirrors
// the cell-draw decision: a per-cell override wins; for a merged block the
// master cell's style is used; remaining gaps fall back to grid-wide defaults
// combined with the column's settings.
begin
  // For a member of a merged block the visible style is the master cell's
  // (DrawCell only paints the master, stretched across the span).
  Result:=CellStyle[ACol,ARow];
  if Result.IsMergedCell then begin
    ACol:=Result.ParentCol;
    ARow:=Result.ParentRow;
    Result:=CellStyle[ACol,ARow];
  end;

  // Fill colour: explicit cell colour, else alternating-row colour, else base.
  if not Result.CellColorIsSet then
    if (ARow mod 2=1) and (FCellColorAlternate<>TAlphaColors.Null) then
      Result.CellColor:=FCellColorAlternate
    else
      Result.CellColor:=FCellColor;

  // Selected-state colours: cell overrides, else grid-wide defaults.
  if not Result.SelectedCellColorIsSet then
    Result.SelectedCellColor:=FSelectedCellColor;
  if not Result.SelectedFontColorIsSet then
    Result.SelectedFontColor:=FCellFontColor;

  // Font colour: cell override, else grid-wide cell font colour.
  if not Result.FontColorIsSet then
    Result.FontColor:=FCellFontColor;

  // Font face/size/style: cell override, else the grid's CellFont.
  if not Result.FontNameIsSet then
    Result.FontName:=FCellFont.Family;
  if not Result.FontSizeIsSet then
    Result.FontSize:=FCellFont.Size;
  if not Result.FontStyleIsSet then
    Result.FontStyle:=FCellFont.Style;

  // Alignment: cell override, else the column's alignment.
  if not Result.TextHAlignmentIsSet then
    Result.TextHAlignment:=ColTextHAlignment[ACol];
  if not Result.TextVAlignmentIsSet then
    Result.TextVAlignment:=ColTextVAlignment[ACol];

  // Word wrap: cell override, else grid-wide OR the column's flag.
  if not Result.WordWrapIsSet then
    if (ACol>=0) and (ACol<FColCount) then
      Result.WordWrap:=FWordWrap or FColData[ACol].WordWrap
    else
      Result.WordWrap:=FWordWrap;
end;

function TMultiHeaderGrid.MeasureCellTextHeight(ACol, ARow: Integer;
                                                const AText: string;
                                                RowSpan: Integer): Single;
// Single source of truth for the height a data cell's text needs (the raw text
// block height, BEFORE padding/gridline are added). Used by both the inplace
// editor (uncommitted FEditor.Text) and any caller that must match the row
// sizing exactly, so the two can never disagree by a fractional pixel. Mirrors
// AutoSizeRows' cmFull branch: cell-style font, wrap via MeasureText else line
// count, divided across a merged RowSpan.
var
  Style      : TCellStyle;
  TextHeight : Single;
  MergedCell : TMergedCell;
begin
  if RowSpan<1 then RowSpan:=1;

  Style:=EffectiveCellStyle(ACol,ARow);

  Canvas.Font.Assign(FCellFont);
  Canvas.Font.Family:=Style.FontName;
  Canvas.Font.Size  :=Style.FontSize;
  Canvas.Font.Style :=Style.FontStyle;

  if Style.WordWrap and (AText<>'') then begin
    var AvailWidth:=GetCellRect(ACol,ARow,MergedCell).Width-FGridLineWidth-FCellPadding.Left-FCellPadding.Right;
    var MeasureRect:=TRectF.Create(0,0,AvailWidth,10000);
    Canvas.MeasureText(MeasureRect,AText,True,[],
                       TTextAlign.Leading,TTextAlign.Leading);
    TextHeight:=MeasureRect.Height;
  end else begin
    // Plain (no-wrap) text: height is simply the line count times the line
    // height. No gridline adjustment is applied per text line, because
    // gridlines sit *between rows*, not between text lines inside a cell;
    // this keeps every sizing path (fast-path, AutoSize, editor) in agreement.
    var LineCount:=CountLines(AText);
    TextHeight:=Canvas.TextHeight('A')*LineCount;
  end;

  var CellDelimterHeight:=FGridLineWidth;
  Result:=Max(Canvas.TextHeight('A'),
              TextHeight/RowSpan-(RowSpan-1)*CellDelimterHeight/4);
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
      // stack instead of forcing its own single level to grow.
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

procedure TMultiHeaderGrid.AutoSizeCols;
var
  i,j:Integer;
  Text:string;
begin
  // If a column rebuild is pending (deferred), do it first so we auto-size the
  // up-to-date layout rather than the stale one. No-op on grids without
  // deferred layout.
  EnsureLayout;

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
  var UseFastMode:=not FAutoSizePrecise and (FRowCount*FColCount>10000);

  // Rows to scan for content width: all rows normally, or only fetched rows on
  // an on-demand cursor (C-fit-fetched).
  var ScanRows:=ColScanRowLimit;

  // Per-column data captured for the conservative-wrap reconciliation pass
  // (only meaningful for columns that were wrap-narrowed):
  //   ColIsWrapped  - this column's width was reduced to let content wrap.
  //   ColWordFloor  - the smallest width that still keeps whole words intact
  //                   (incl. padding); never shrink a wrapped column below this.
  //   ColNaturalW   - the width that would show the widest content on ONE line
  //                   (incl. padding); the cap we grow a wrapped column back to.
  var ColIsWrapped: TArray<Boolean>; SetLength(ColIsWrapped, FColCount);
  var ColWordFloor: TArray<Single>;  SetLength(ColWordFloor, FColCount);
  var ColNaturalW:  TArray<Single>;  SetLength(ColNaturalW,  FColCount);

  for i:=0 to FColCount-1 do begin
    // HeaderFullW  - widest header caption laid out on a SINGLE line.
    // HeaderWordW  - widest single word (the tightest a wrapping header can be).
    // DataW        - widest cell content (honouring cell word-wrap as before).
    // DataFullW    - widest cell content on ONE line (ignores wrap shrink); used
    //                as the natural-width cap when reconciling against viewport.
    // CanWrapHdr   - any header element over this column allows word wrap.
    var HeaderFullW:Single:=0;
    var HeaderWordW:Single:=0;
    var DataW:Single:=0;
    var DataFullW:Single:=0;
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

    // Optimization: if this column has no WordWrap, use the fast logic
    var ColHasWordWrap:=FWordWrap or FColData[i].WordWrap or FGridCellsHasWordWrap;

    if not ColHasWordWrap and UseFastMode then begin
      // Fast path
      Canvas.Font.Assign(FCellFont);
      var FLetterWidth:=Canvas.TextWidth('V');
      var MaxLetters:=0.0;

      for j:=0 to ScanRows-1 do begin
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
      for j:=0 to ScanRows-1 do begin
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

          // Natural one-line width (the cap we may grow back toward), kept
          // separately from the possibly-collapsed DataW.
          DataFullW:=Max(DataFullW,FullLineW/ColSpan-(ColSpan-1)*CellDelimterWidth/2);

          if FullLineW>DataWrapGateW then begin
            // Wide content: size to the widest single word, let it wrap.
            var Words:=Text.Split([' ', ':', ';', ',', '.', '!', '?', '-', '+', '*', '/', '\', '|', #9]);
            var MaxWordWidth:=0.0;
            for var Word in Words do
              if Word<>'' then
                MaxWordWidth:=Max(MaxWordWidth,Canvas.TextWidth(Word));

            if ConservativeWrap then begin
              // Try to keep Cell Width/Height as 5/1
              var TH:=Canvas.TextHeight('A');
              var D:=DataFullW/MaxWordWidth*TH;
              if D*5>MaxWordWidth then begin
                MaxWordWidth:=Min(Sqrt(5*MaxWordWidth*D),DataFullW);
              end;
            end;

            DataW:=Max(DataW,MaxWordWidth/ColSpan-(ColSpan-1)*CellDelimterWidth/2);
          end else begin
            // Already small enough: keep it on one line.
            DataW:=Max(DataW,FullLineW/ColSpan-(ColSpan-1)*CellDelimterWidth/2);
          end;
        end else begin
          // Without WordWrap, measure each line
          var Lines:=Text.Split([#13#10]);
          for var Line in Lines do begin
            var LineW:=Canvas.TextWidth(Line)/ColSpan-(ColSpan-1)*CellDelimterWidth/2;
            DataW:=Max(DataW,LineW);
            DataFullW:=Max(DataFullW,LineW);
          end;
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

    // --- Conservative-wrap bookkeeping -------------------------------------
    // The word-width floor: the narrowest this column may become without
    // splitting whole words (data words and header words both considered),
    // used only when SHRINKING to fit.
    var WordFloorW:=Max(DataW, Min(HeaderWordW, WordWidthCap))+CellPaddingFull;
    ColWordFloor[i]:=Max(WordFloorW, Floor);
    // The natural one-line width is DATA-only: a wide header is allowed to wrap
    // and must NOT pull the column wider. We only ever grow a column back to the
    // width its data needs on one line - never to fit the caption.
    ColNaturalW[i]:=DataFullW+CellPaddingFull;
    if ColNaturalW[i]<NewWidth then ColNaturalW[i]:=NewWidth;
    // A column has reclaimable slack only when its DATA was cut: the data's
    // one-line width exceeds the width we settled on. Header wrapping alone does
    // not count - growing to un-wrap a caption is not wanted.
    ColIsWrapped[i]:=(DataFullW>DataW+1) and (ColNaturalW[i]>NewWidth+1);

    // Limit width
    NewWidth:=Min(NewWidth,Max(FMaxColumnAutoWidth,FColData[i].MaxWidth));

    ColWidths[i]:=NewWidth;
  end;

  // --- Conservative-wrap reconciliation against the viewport ---------------
  // Wrapped columns above were sized down to word width so text wraps. Now
  // reconcile the total against the available viewport width: hand spare
  // horizontal space back to wrapped columns (so they wrap less), or, on
  // overflow, shrink them proportionally toward their word floor to fit.
  if FConservativeWrap and GridHaveWordWrap and (FColCount>0) then
    ReconcileWrappedColumns(ColIsWrapped, ColWordFloor, ColNaturalW, VP);

  // Header heights must follow the (possibly wrap-narrowed) column widths.
  AutoSizeHeaders;

  // Balance an over-wrapped header band: when the header word-wraps, every
  // column whose caption wrapped onto more lines than typical is widened in a
  // single pass straight to its one-line width, so the over-tall header band
  // shrinks back toward balance. The pass re-fits header heights itself.
  if FHeaderWordWrap and (FHeaderLevels.Count>0) and (FColCount>0) then
    BalanceHeaderColumnWidths;

  Invalidate;
end;

procedure TMultiHeaderGrid.ReconcileWrappedColumns(
  const AIsWrapped: TArray<Boolean>;
  const AWordFloor, ANaturalW: TArray<Single>;
  AViewportW: Integer);
var
  i: Integer;
begin
  if AViewportW<=0 then Exit;

  // Total currently allocated, and how much the wrapped columns can flex.
  var Total:=0;
  var GrowHeadroom:Single:=0;  // sum of (natural - current) over wrapped cols
  var ShrinkHeadroom:Single:=0;// sum of (current - floor)   over wrapped cols
  for i:=0 to FColCount-1 do begin
    Total:=Total+FColData[i].Widths;
    if AIsWrapped[i] then begin
      var Cur:=FColData[i].Widths;
      var Cap:=ANaturalW[i];
      // Respect an explicit per-column MaxWidth as the real upper bound.
      if (FColData[i].MaxWidth>0) and (Cap>FColData[i].MaxWidth) then
        Cap:=Max(FMaxColumnAutoWidth,FColData[i].MaxWidth);
      if Cap>Cur then GrowHeadroom:=GrowHeadroom+(Cap-Cur);
      if Cur>AWordFloor[i] then ShrinkHeadroom:=ShrinkHeadroom+(Cur-AWordFloor[i]);
    end;
  end;

  // Leave a hair of slack so we don't trip the horizontal scrollbar by 1px.
  var Avail:=AViewportW-2;

  if (Total<Avail) and (GrowHeadroom>0) then begin
    // Spare space: grow wrapped columns back toward their one-line width,
    // proportional to each column's remaining headroom. Never exceed natural
    // width (or MaxWidth), so we stop wrapping rather than over-stretch.
    var Spare:Single:=Avail-Total;
    if Spare>GrowHeadroom then Spare:=GrowHeadroom; // don't grow past natural
    for i:=0 to FColCount-1 do begin
      if not AIsWrapped[i] then Continue;
      var Cur:=FColData[i].Widths;
      var Cap:=Min(ANaturalW[i],Max(FMaxColumnAutoWidth,FColData[i].MaxWidth));
      var Room:=Cap-Cur;
      if Room<=0 then Continue;
      var Add:=Round(Spare*(Room/GrowHeadroom));
      if Add>Room then Add:=Round(Room);
      if Add>0 then ColWidths[i]:=Cur+Add;
    end;
  end else if (Total>Avail) and (ShrinkHeadroom>0) then begin
    // Overflow: shrink wrapped columns proportionally toward their word floor,
    // only as much as needed to fit (never below the floor, so whole words stay
    // intact and the rest wraps).
    var Excess:Single:=Total-Avail;
    if Excess>ShrinkHeadroom then Excess:=ShrinkHeadroom; // can't reclaim more
    for i:=0 to FColCount-1 do begin
      if not AIsWrapped[i] then Continue;
      var Cur:=FColData[i].Widths;
      var Room:=Cur-AWordFloor[i];
      if Room<=0 then Continue;
      var Cut:=Round(Excess*(Room/ShrinkHeadroom));
      if Cut>Room then Cut:=Round(Room);
      if Cut>0 then ColWidths[i]:=Cur-Cut;
    end;
  end;
end;

function TMultiHeaderGrid.HeaderElementLineCount(ALevel, ACol: Integer): Integer;
// How many wrapped text lines the caption at (ALevel,ACol) needs at the column's
// current merged width. Mirrors AutoSizeHeaders' wrapped-branch measurement so
// the balancing pass reasons about the same line counts the heights came from.
begin
  Result:=0;
  var Element:=FHeaderLevels.GetElementAtCell(ACol,ALevel);
  if not Assigned(Element) then Exit;
  if Element.Caption='' then Exit;
  // Only count for the cell that actually owns this position (not span members).
  if FHeaderLevels.IndexOf(Element.FLevel)<>ALevel then Exit;
  if HeaderCellIsFiller(ALevel, ACol) then Exit;

  Canvas.Font.Assign(FCellFont);
  if Element.Style.FontNameIsSet then Canvas.Font.Family:=Element.Style.FontName;
  if Element.Style.FontSizeIsSet then Canvas.Font.Size:=Element.Style.FontSize;
  if Element.Style.FontStyleIsSet then Canvas.Font.Style:=Element.Style.FontStyle;

  var BaseTH:=Canvas.TextHeight('A');

  if HeaderCellWordWrap(Element) then begin
    var W:=HeaderElementTextWidth(ALevel, ACol);
    var MeasureRect:=TRectF.Create(0,0,W,100000);
    Canvas.MeasureText(MeasureRect,Element.Caption,True,[],
                       TTextAlign.Leading,TTextAlign.Leading);
    Result:=Max(1,Ceil(MeasureRect.Height/Max(BaseTH,1)));
  end else
    Result:=Max(CountLines(Element.Caption),1);
end;

function TMultiHeaderGrid.BalanceHeaderColumnWidths: Boolean;
// In a single pass, find columns whose header caption wrapped onto substantially
// more lines than the typical column (the average line count, rounded up) and
// widen them just enough to bring them down toward that typical height - NOT all
// the way to a single line, which would leave the header cell with blank, unused
// lines. Capped by the caption's true one-line width and the column's MaxWidth /
// MaxColumnAutoWidth. No viewport budget and no borrowing from other columns - if
// the result exceeds the viewport the grid scrolls horizontally. Re-fits header
// heights and returns True if any width changed.
var
  i, lvl: Integer;
begin
  Result:=False;
  if FColCount<=0 then Exit;

  var CellPaddingWidth:=CellPadding.Left+CellPadding.Right;
  var CellDelimterWidth:=FGridLineWidth/2;
  var CellPaddingFull:=CellPaddingWidth+CellDelimterWidth+1;

  // Per column: worst wrapped-line count over all its header levels, plus the
  // single-line width its widest-wrapping caption would need (the growth cap).
  // Also remember WHICH element drives that worst count, so we can re-measure its
  // caption at trial widths when deciding how far to widen.
  var ColLines: TArray<Integer>;  SetLength(ColLines, FColCount);
  var ColOneLineW: TArray<Single>; SetLength(ColOneLineW, FColCount);
  var ColElement: TArray<THeaderElement>; SetLength(ColElement, FColCount);
  // Width contributed by the OTHER columns of the worst element's span (0 for a
  // single-column header). Added to a trial column width to get the element's
  // total text width when re-measuring.
  var ColSpanOther: TArray<Single>; SetLength(ColSpanOther, FColCount);
  var MaxLines:=1;
  var SumLines:=0;

  for i:=0 to FColCount-1 do begin
    ColLines[i]:=1;
    ColOneLineW[i]:=0;
    ColElement[i]:=nil;
    ColSpanOther[i]:=0;
    for lvl:=0 to FHeaderLevels.Count-1 do begin
      var Lines:=HeaderElementLineCount(lvl, i);
      if Lines>=ColLines[i] then begin
        if Lines>ColLines[i] then ColLines[i]:=Lines;
        // Width the widest caption line of this element needs on ONE line, so we
        // know how far it is worth growing the column (never beyond this).
        var Element:=FHeaderLevels.GetElementAtCell(i,lvl);
        if Assigned(Element) then begin
          ColElement[i]:=Element;
          // The element's current total text width minus this column's share -
          // i.e. what the rest of the span already contributes. For a single
          // column header this is ~0. Uses the SAME text inset the renderer and
          // HeaderElementTextWidth apply (Left+Right padding + full grid line).
          var TextInset:=CellPadding.Left+CellPadding.Right+FGridLineWidth;
          var TotalTextW:=HeaderElementTextWidth(lvl, i);
          ColSpanOther[i]:=Max(0, TotalTextW-(FColData[i].Widths-TextInset));

          Canvas.Font.Assign(FCellFont);
          if Element.Style.FontNameIsSet then Canvas.Font.Family:=Element.Style.FontName;
          if Element.Style.FontSizeIsSet then Canvas.Font.Size:=Element.Style.FontSize;
          if Element.Style.FontStyleIsSet then Canvas.Font.Style:=Element.Style.FontStyle;
          var OneLine:Single:=0;
          for var Ln in Element.Caption.Split([#13#10]) do
            OneLine:=Max(OneLine,Canvas.TextWidth(Ln));
          // Per-spanned-column share of the one-line caption width.
          var Span:=Element.FColSpan; if Span<1 then Span:=1;
          ColOneLineW[i]:=Max(ColOneLineW[i],
            OneLine/Span+CellPaddingFull);
        end;
      end;
    end;
    if ColLines[i]>MaxLines then MaxLines:=ColLines[i];
    SumLines:=SumLines+ColLines[i];
  end;

  // Nothing wraps onto multiple lines - the band is already balanced.
  if MaxLines<=1 then Exit;

  // The "typical" line count is the AVERAGE rounded up. A column must wrap onto
  // substantially more lines than this to be considered over-tall.
  var Threshold:=Ceil(SumLines/FColCount);
  if Threshold<1 then Threshold:=1;

  // A column is treated as over-tall only when it wraps onto more than
  // (Threshold + AllowedExcess) lines. When we widen one, we bring it down only
  // to that boundary - NOT all the way to the typical height, which for a long
  // caption would demand a huge width (e.g. 9 lines -> 2 needs ~4.5x the width).
  // Landing at the boundary keeps the column only as wide as needed to stop being
  // an outlier.
  const AllowedExcess = 3;
  var TargetLines:=Max(2,Threshold+AllowedExcess);

  for i:=0 to FColCount-1 do begin
    if ColLines[i]<=Threshold+AllowedExcess then Continue; // not over-tall
    if ColLines[i]<3 then Continue;
    if not Assigned(ColElement[i]) then Continue;

    var Cur:=FColData[i].Widths;

    var Cap:=ColOneLineW[i];
    var ColMax:=Max(FMaxColumnAutoWidth,FColData[i].MaxWidth);
    if Cap>ColMax then Cap:=ColMax;
    if Cap<=Cur+1 then Continue;                  // already at/over its cap

    // Measure how many lines this column's governing caption needs at a trial
    // COLUMN width - by laying the caption out at the matching text width. This
    // is exact (no inverse-proportional guess), so a long single word or uneven
    // lines can't fool us into over-widening.
    var Element:=ColElement[i];
    Canvas.Font.Assign(FCellFont);
    if Element.Style.FontNameIsSet then Canvas.Font.Family:=Element.Style.FontName;
    if Element.Style.FontSizeIsSet then Canvas.Font.Size:=Element.Style.FontSize;
    if Element.Style.FontStyleIsSet then Canvas.Font.Style:=Element.Style.FontStyle;
    var BaseTH:=Canvas.TextHeight('A');
    var SpanOther:=ColSpanOther[i];

    // Lines needed if this column were width W. Uses the same text inset the
    // renderer applies, so the measured count matches what is actually drawn.
    var TextInset:=CellPadding.Left+CellPadding.Right+FGridLineWidth;
    var LinesAt: TFunc<Single,Integer> :=
      function(W: Single): Integer
      begin
        var TextW:=(W-TextInset)+SpanOther;
        if TextW<1 then TextW:=1;
        var R:=TRectF.Create(0,0,TextW,100000);
        Canvas.MeasureText(R,Element.Caption,True,[],
                           TTextAlign.Leading,TTextAlign.Leading);
        Result:=Max(1,Ceil(R.Height/Max(BaseTH,1)));
      end;

    // Binary-search the SMALLEST column width in (Cur, Cap] whose caption fits in
    // TargetLines or fewer. If even Cap can't reach TargetLines (very long text),
    // we still grow only to Cap.
    var Hi:Single:=Cap;
    if LinesAt(Hi)<=TargetLines then begin
      // Narrow down to ~1px. Hi always satisfies, Lo never does.
      var Lo:Single:=Cur;
      while Hi-Lo>1 do begin
        var Mid:=(Lo+Hi)/2;
        if LinesAt(Mid)<=TargetLines then Hi:=Mid else Lo:=Mid;
      end;
    end;

    var Target:=Hi;
    if Target>Cur+1 then begin
      ColWidths[i]:=Round(Target);
      if FColData[i].Widths<>Round(Cur) then Result:=True;
    end;
  end;

  if Result then
    AutoSizeHeaders;
end;

type
  TSizeComputeMode = (cmFast,cmSlow,cmFull);

procedure TMultiHeaderGrid.AutoSizeRows(FromRow: integer = 0; ToRow: integer = -1;
                                        FResizeStartColumnIndex: integer = -1; FResizeEndColumnIndex: integer = -1;
                                        TryOptimise: boolean = False);
begin
  // If a column rebuild is pending (deferred), do it first so we auto-size the
  // up-to-date layout rather than the stale one. No-op on grids without
  // deferred layout.
  EnsureLayout;

  if ToRow<0 then begin
    ToRow:=FRowCount-1;
  end;

  if (FResizeStartColumnIndex<>-1) or (FResizeEndColumnIndex<>-1) then begin
    TryOptimise:=False;
  end;

  // Suppress on-demand fetching for the measurement sweep below: it iterates
  // every row, which on an on-demand cursor would force-fetch / mutate FRowData
  // mid-loop. Fetching is driven by paint/scroll, not layout.
  FInLayout:=True;
  try
    Canvas.Font.Assign(FCellFont);

    var CellPaddingHeight:=CellPadding.Top+CellPadding.Bottom;
    var CellDelimterHeight:=FGridLineWidth;
    var ViewBottomCell:=ViewBottom;

    var ComputeMode:=TSizeComputeMode.cmSlow;
    if FRowCount*FColCount>10000 then begin
      ComputeMode:=TSizeComputeMode.cmFast;
    end;
    if FAutoSizePrecise then begin
      ComputeMode:=TSizeComputeMode.cmSlow;
    end;
    if GridHaveWordWrap then begin
      ComputeMode:=TSizeComputeMode.cmFull;
    end;

    if FResizeStartColumnIndex<0 then begin
      FResizeStartColumnIndex:=0;
    end;
    if FResizeEndColumnIndex<0 then begin
      FResizeEndColumnIndex:=FColCount-1;
    end;

    Canvas.Font.Assign(FCellFont);
    var TH:=Canvas.TextHeight('A');
    var Changed:=False;
    for var Row:=FromRow to FRowCount-1 do begin
      if ((ToRow<0) and (Row>0) and (FRowData[Row-1].Top>ViewBottomCell)) or
         ((ToRow>=0) and (Row>ToRow)) then Break;

      if TryOptimise and FRowData[Row].AutoSized then Continue;

      var MaxHeight:=TH;

      for var Col:=FResizeStartColumnIndex to FResizeEndColumnIndex do begin

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

            MaxHeight:=Max(MaxHeight,
                           TextLines*ResTH/RowSpan-(RowSpan-1)*CellDelimterHeight/4);
          end;

          cmSlow,
          cmFull: begin
            // Precise calculation. Both the no-wrap (cmSlow) and word-wrap (cmFull)
            // cases delegate to MeasureCellTextHeight - the same routine the inplace
            // editor uses for its live row-fit. Sharing one measurement (identical
            // font assignment, gridline subtraction and single-line floor) is what
            // makes the committed-text row height and the editor's height agree for
            // custom-sized/styled fonts; the previous inline formulas drifted from
            // it and made styled rows shrink a pixel on edit start.
            var ActualCol:=Col;
            var ActualRow:=Row;
            if IsMergedCell(Col,Row,MergedCell) then begin
              ActualCol:=MergedCell.Col;
              ActualRow:=MergedCell.Row;
            end;
            MaxHeight:=Max(MaxHeight,
                           MeasureCellTextHeight(ActualCol,ActualRow,Cells[ActualCol,ActualRow],RowSpan));
          end;
        end;
      end;

      if (Row>=0) and (Row<Length(FRowData)) then begin
        // Ceil (not Trunc) so this matches MemoEditorTextChanged's row-fit, which
        // also Ceils. Flooring here while the editor Ceils made each edit re-floor
        // the row a sub-pixel shorter than the editor had grown it, so repeated
        // edits slowly shrank the row. Rounding both the same way removes that.
        var NewSize:=Ceil(MaxHeight+CellPaddingHeight+CellDelimterHeight/2);
        if FRowData[Row].Height<>NewSize then begin
          FRowData[Row].Height:=NewSize;
          Changed:=True;
        end;

        if Row>0 then begin
          FRowData[Row].Top:=FRowData[Row-1].Top+FRowData[Row-1].Height;
        end;
      end;
      if TryOptimise then begin
        FRowData[Row].AutoSized:=True;
      end;
    end;

    if Changed then begin
      for var Row:=1 to Min(FRowCount-1,High(FRowData)) do begin
        FRowData[Row].Top:=FRowData[Row-1].Top+FRowData[Row-1].Height;
      end;
    end;
  finally
    FInLayout:=False;
  end;
end;

procedure TMultiHeaderGrid.AutoSizeVisibleRows(FResizeStartColumnIndex: integer = -1; FResizeEndColumnIndex: integer = -1);
// Re-fits only the rows currently on screen (first..last visible row). Used for
// live feedback during a column drag, where re-measuring every row below the
// viewport on each MouseMove would be needlessly expensive.
begin
  var First:=RowAtHeightCoord(ViewTop);
  if First<0 then First:=0;
  var Last:=RowAtHeightCoord(ViewBottom);
  if Last<First then Last:=FRowCount-1; // viewport past the last row
  AutoSizeRows(First, Last, FResizeStartColumnIndex, FResizeEndColumnIndex, True);

  UpdateSize;
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
  inherited;  // fire user OnMouseMove before the grid's own resize handling

  if FResizeEnabled then begin
    case FResizeMode of
      TResizeMode.rmColumn: begin
        if FResizeEndColumnIndex>=0 then begin
          // Use global coordinates for accurate tracking
          GlobalPos:=LocalToAbsolute(PointF(X, Y));
          NewWidth:=ResizeStartWidth+Round(GlobalPos.X-FResizeStartPos.X);
          UpdateColumnWidth(FResizeStartColumnIndex,FResizeEndColumnIndex, NewWidth);
          Cursor:=crHSplit;

          if FPainted then begin
            FPainted:=False;
            // Wrapped text reflows as the column narrows/widens, changing row
            // heights. Re-fit only the on-screen rows for live feedback; the full
            // pass runs once on MouseUp (drag end).
            if GridHaveWordWrap then begin
              AutoSizeVisibleRows(FResizeStartColumnIndex,FResizeEndColumnIndex);
            end;
          end;
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
      var WasColumnResize:=(FResizeMode=TResizeMode.rmColumn);
      DoColumnResized;
      FResizeRowIndex:=-1;
      FResizeMode:=TResizeMode.rmNone;
      Cursor:=crDefault;

      // Column width is now final; do the authoritative full re-fit of every
      // row so off-screen rows (skipped during the live drag) match the new
      // wrap. No-op when the grid doesn't word-wrap.
      if WasColumnResize and GridHaveWordWrap then begin
        AutoSizeRows(0,-1, FResizeStartColumnIndex, FResizeEndColumnIndex);
        UpdateSize;
      end;

      FResizeStartColumnIndex:=-1;
      FResizeEndColumnIndex:=-1;
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
  if not FPainted then Exit;

  FEndJump:=False;

  // Fire the user's OnKeyDown first (VCL-like). A handler may consume the key
  // by setting Key:=0 and KeyChar:=#0, in which case the grid skips its own
  // navigation / editing handling below.
  inherited;
  if (Key=0) and (KeyChar=#0) then Exit;

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
        // Let the (overridable) Down handler run first. If it fully handled the
        // key (e.g. appended a record or cancelled a pending insert), stop;
        // otherwise fall through to a plain one-row move.
        if (not (ssCtrl in Shift)) and HandleDownKey then begin
          Key:=0; KeyChar:=#0;
          Exit;
        end;
        DY:=1;
      end;
      vkHome: begin
        DY:=-RowCount;
      end;
      vkEnd: begin
        // Jump to the last row already fetched. On-demand descendants extend past
        // this to the true Eof in their own KeyDown (page-by-page, repainting),
        // so End must not synchronously grow the cursor here - that blocks the UI.
        DY:=RowCount;
        FEndJump:=True;
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
      vkC: begin // Ctrl+C / C copies the current cell
        if (Row>=0) and (Col>=0) then begin
          var MergedCell: TMergedCell;
          if IsMergedCell(Col,Row,MergedCell) then
            CopyTextToClipboard(Cells[MergedCell.Col,MergedCell.Row])
          else
            CopyTextToClipboard(Cells[Col, Row]);
        end;
        Exit;
      end;
      vkInsert: begin
        // Ctrl+Insert keeps the legacy copy shortcut. A plain Insert inserts a
        // new row before the current one (when the grid is modifiable); if it
        // isn't, fall back to copy so Insert still does something useful.
        if (ssCtrl in Shift) or FReadOnly then begin
          if (Row>=0) and (Col>=0) then begin
            var MergedCell: TMergedCell;
            if IsMergedCell(Col,Row,MergedCell) then
              CopyTextToClipboard(Cells[MergedCell.Col,MergedCell.Row])
            else
              CopyTextToClipboard(Cells[Col, Row]);
          end;
        end else begin
          InsertRow(Max(0, Row));
        end;
        Key:=0; KeyChar:=#0;
        Exit;
      end;
      vkDelete: begin
        // Ctrl+Delete deletes the current row (when modifiable).
        if (ssCtrl in Shift) and (not FReadOnly) and (Row>=0) then begin
          DeleteRow(Row);
          Key:=0; KeyChar:=#0;
          Exit;
        end;
      end;
    end;

    NewCol:=FSelectedCell.X;
    NewRow:=FSelectedCell.Y;
    repeat
      OldCol:=NewCol;
      OldRow:=NewRow;

      NewCol:=Max(0, Min(FColCount-1, NewCol+DX));
      // On-demand grids only know FRowCount rows fetched so far. When moving
      // down, give a descendant the chance to fetch the target row first so the
      // clamp below does not stop at the fetched boundary. No-op on base grids.
      // End is excluded: it must land on the last fetched row without a blocking
      // grow - the on-demand descendant pages on to Eof afterward.
      if (DY>0) and (not FEndJump) then EnsureRowAvailable(NewRow+DY);
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
      FPainted:=False;
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
  // Column-focus change: OnColExit for the old column, OnColEnter for the new.
  // Fired here (not in SetCol) so keyboard/mouse navigation - which set
  // FSelectedCell.X directly - are covered too. FLastColEvent guards against
  // duplicate firing when only the row changed.
  if FSelectedCell.X<>FLastColEvent then begin
    if (FLastColEvent>=0) and Assigned(FOnColExit) then
      FOnColExit(Self);
    FLastColEvent:=FSelectedCell.X;
    if (FSelectedCell.X>=0) and Assigned(FOnColEnter) then
      FOnColEnter(Self);
  end;

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
  if High(FRowData)>=ARow then begin
    FRowData[ARow].AutoSized:=False;
  end;

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
  if High(FRowData)>=ARow then begin
    FRowData[ARow].AutoSized:=False;
  end;

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
  // Did this cell already wrap before the new style is applied?
  var WrappedBefore:=EffectiveCellWordWrap(ACol, ARow);

  DoSetCellStyle(ACol, ARow, Value);

  // If the new per-cell style turns word-wrap on where it was off, the row's
  // height can change; fit just that row now. Comparing before/after means
  // merge/clear operations that merely rewrite an already-wrapped style don't
  // trigger a needless pass.
  if (not FSuppressAutoSize) and (ARow>=0) and (ARow<FRowCount) and
     (not WrappedBefore) and EffectiveCellWordWrap(ACol, ARow) then begin
    AutoSizeRows(ARow, ARow);
    UpdateSize;
  end;
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
  // An editor may open on any in-range cell, even when ReadOnly, so its text
  // can be selected and copied. Whether edits are accepted is CellIsModifiable.
  Result:=(ACol>=0) and (ACol<FColCount) and (ARow>=0) and (ARow<FRowCount);
end;

function TMultiHeaderGrid.CellIsModifiable(ACol, ARow: Integer): Boolean;
begin
  // Base/string grids accept edits unless ReadOnly. Descendants tighten this
  // (the DB grid also checks DataSet.CanModify and the field's ReadOnly).
  Result:=(not FReadOnly) and
          (ACol>=0) and (ACol<FColCount) and (ARow>=0) and (ARow<FRowCount);
end;

procedure TMultiHeaderGrid.ApplyEditorReadOnly(Ed: TControl; AModifiable: Boolean);
// Puts the active editor control into read-only mode when the cell is not
// modifiable, so the user can still select/copy the text but not change it.
// Covers every control the grids use as an inplace editor.
begin
  if Ed is TCustomMemo then
    TCustomMemo(Ed).ReadOnly:=not AModifiable
  else if Ed is TCustomEdit then
    TCustomEdit(Ed).ReadOnly:=not AModifiable
  else if Ed is TCustomComboBox then
    TCustomComboBox(Ed).Enabled:=AModifiable;
end;

function TMultiHeaderGrid.DoInsertRow(ARow: Integer): Boolean;
begin
  Result:=True;
  if Assigned(FOnInsertRow) then FOnInsertRow(Self, ARow, Result);
end;

function TMultiHeaderGrid.DoAppendRow(ARow: Integer): Boolean;
begin
  Result:=True;
  if Assigned(FOnAppendRow) then FOnAppendRow(Self, ARow, Result);
end;

function TMultiHeaderGrid.DoDeleteRow(ARow: Integer): Boolean;
begin
  Result:=True;
  if Assigned(FOnDeleteRow) then FOnDeleteRow(Self, ARow, Result);
end;

procedure TMultiHeaderGrid.InsertRow(ARow: Integer);
// Base/string-grid model: grow RowCount by one, then shift cell contents down
// from ARow so the inserted row is blank. SetRowCount keeps FRowData in sync.
begin
  if FReadOnly then Exit;
  if ARow<0 then ARow:=0;
  if ARow>FRowCount then ARow:=FRowCount;
  if not DoInsertRow(ARow) then Exit;

  if FEditing then CancelEditing;
  RowCount:=FRowCount+1;
  for var R:=FRowCount-1 downto ARow+1 do
    for var C:=0 to FColCount-1 do begin
      Cells[C,R]:=Cells[C,R-1];
      CellStyle[C,R]:=CellStyle[C,R-1];
    end;
  for var C:=0 to FColCount-1 do
    Cells[C,ARow]:='';

  FSelectedCell:=Point(Min(FSelectedCell.X,FColCount-1), ARow);
  AutoSizeRows;
  UpdateSize;
  ScrollToSelectedCell;
  DoSelectCell;
  Invalidate;
end;

procedure TMultiHeaderGrid.AppendRow;
// Adds a blank row after the current last row and selects it.
begin
  if FReadOnly then Exit;
  if not DoAppendRow(FRowCount) then Exit;

  if FEditing then CommitEditing;
  RowCount:=FRowCount+1;
  for var C:=0 to FColCount-1 do
    Cells[C,FRowCount-1]:='';

  FSelectedCell:=Point(Min(FSelectedCell.X,FColCount-1), FRowCount-1);
  AutoSizeRows;
  UpdateSize;
  ScrollToSelectedCell;
  DoSelectCell;
  Invalidate;
end;

function TMultiHeaderGrid.HandleDownKey: Boolean;
// Base/string grids: pressing Down on the last row appends a blank row.
begin
  Result:=False;
  if FReadOnly then Exit;
  if (FRowCount>0) and (Row=FRowCount-1) then begin
    AppendRow;
    Result:=True;
  end;
end;

procedure TMultiHeaderGrid.DeleteRow(ARow: Integer);
// Removes row ARow, shifting subsequent rows up. The last remaining row is
// kept (a string grid is never reduced below one row).
begin
  if FReadOnly then Exit;
  if (ARow<0) or (ARow>=FRowCount) then Exit;
  if FRowCount<=1 then Exit;
  if not DoDeleteRow(ARow) then Exit;

  if FEditing then CancelEditing;
  for var R:=ARow to FRowCount-2 do
    for var C:=0 to FColCount-1 do begin
      Cells[C,R]:=Cells[C,R+1];
      CellStyle[C,R]:=CellStyle[C,R+1];
    end;
  RowCount:=FRowCount-1;

  FSelectedCell:=Point(Min(FSelectedCell.X,FColCount-1), Min(ARow,FRowCount-1));
  AutoSizeRows;
  UpdateSize;
  ScrollToSelectedCell;
  DoSelectCell;
  Invalidate;
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

  FEditor:=TMemo.Create(Self);
  FEditor.Align:=TAlignLayout.None;
  FEditor.WordWrap:=False;
  FEditor.AutoSelect:=False;
  FEditor.StyledSettings:=[];
  FEditor.ShowScrollBars:=False;
  FEditor.DisableMouseWheel:=True;
  FEditor.TextSettings.HorzAlign:=TextAlign;
  FEditor.OnChange:=MemoEditorTextChanged;
  FEditor.StyleLookup:='mhg_editor_memo';

  WireEditor(FEditor); // parent into host + wire OnKeyDown/OnExit
end;

procedure TMultiHeaderGrid.MemoEditorTextChanged(Sender: TObject);
// Live re-fit of the edited row. A row's height is the max over ALL columns, so
// (1) AutoSizeRows sizes it from every cell's committed text (this also shrinks
// the row when the other cells allow it), then (2) we grow the anchor row if
// the uncommitted editor text needs more. Both steps measure through
// MeasureCellTextHeight, so the editing height equals the final committed
// height to the pixel. The model is never written here - the value stays
// uncommitted until CommitEditorValue.
var
  MergedCell : TMergedCell;
  RowSpan    : Integer;
  FromRow    : Integer;
  ToRow      : Integer;
  EditH      : Integer;
begin
  if (not FEditing) or (FEditor=nil) then Exit;

  RowSpan:=1;
  FromRow:=FEditRow;
  ToRow  :=FEditRow;
  if IsMergedCell(FEditCol,FEditRow,MergedCell) then begin
    RowSpan:=MergedCell.RowSpan;
    FromRow:=MergedCell.Row;
    ToRow  :=MergedCell.Row+MergedCell.RowSpan-1;
  end;

  // 1. Floor from all committed cells in the row/span (the edited cell still
  //    holds its pre-edit text - step 2 accounts for the live editor text).
  //    Use the grid's sticky AutoSize mode so a fast grid stays fast; step 2
  //    measures the edited cell precisely regardless, so the editor aligns.
  AutoSizeRows(FromRow,ToRow);

  // 2. Height the uncommitted editor text needs, measured identically, then
  //    converted to a whole-pixel row height with Ceil (never Trunc - flooring
  //    leaves the cell a sub-pixel too short and clips the text).
  var CellPaddingHeight:=CellPadding.Top+CellPadding.Bottom;
  var CellDelimterHeight:=FGridLineWidth;
  var TextH:=MeasureCellTextHeight(FEditCol,FEditRow,FEditor.Text,RowSpan);
  EditH:=Ceil(TextH+CellPaddingHeight+CellDelimterHeight/2);

  // 3. Grow the anchor row only if the editor needs more than the committed
  //    cells already gave it. Never shrink below what other cells require
  //    (step 1 set that floor); shrinking on delete is handled by step 1.
  if EditH>GetRowHeight(FEditRow) then
    SetRowHeight(FEditRow,EditH);

  UpdateSize;     // total height may have changed in step 1 or 3
  PositionEditor; // re-place editor inside the resized cell
end;

procedure TMultiHeaderGrid.WireEditor(C: TControl);
begin
  // Editors live inside FEditorHost, sitting on the FEditorBack backing and
  // clipped to the cell. The host must already exist (EnsureEditor creates it
  // before any editor).
  C.Parent:=FEditorHost;
  C.Visible:=False;
  C.OnKeyDown:=EditorKeyDown;
  C.OnExit:=EditorExit;
end;

procedure TMultiHeaderGrid.LoadEditorValue(ACol, ARow: Integer; InitialChar: Char);
begin
  // Match the cell's wrapping before assigning text so the editor lays the
  // content out the same way the cell does.
  FEditor.WordWrap:=EffectiveCellWordWrap(ACol, ARow);
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
      if IsMergedCell(ACol,ARow,MergedCell) then begin
        AutoSizeRows(MergedCell.Row, MergedCell.Row+MergedCell.RowSpan-1)
      end else begin
        AutoSizeRows(ARow, ARow);
      end;
      UpdateSize; // total height changed -> refresh scrollbar range
    end;
end;

procedure TMultiHeaderGrid.PositionEditor;
begin
  var Ed:=ActiveEditorControl;
  if (Ed=nil) or (not FEditing) or (FEditorHost=nil) then Exit;

  // The cell rect, built the same way DrawCells does so the editor lines up
  // with the painted cell at any GridLineWidth. FEditCol/FEditRow are the
  // merge anchor; for a merged cell the rect spans the merged range.
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

  // Keep the editor below the header.
  if R.Top<HeaderHeight then R.Top:=HeaderHeight;

  // Clamp to the visible cells area. Without this an editor wider/taller than
  // the viewport paints outside the grid and leaves artefacts at the border.
  var MaxRight:=ViewPortWidth;
  var MaxBottom:=HeaderHeight+ViewCellsHeight;
  if R.Right>MaxRight then R.Right:=MaxRight;
  if R.Bottom>MaxBottom then R.Bottom:=MaxBottom;

  // Show only while a positive area of the cell is within the visible region.
  var Vis:=(R.Width>0) and (R.Height>0) and
           (R.Bottom>HeaderHeight) and (R.Top<Height) and
           (R.Right>0) and (R.Left<Width);

  if not Vis then begin
    HideEditor;
    Exit;
  end;

  // Position only the host; its children (the opaque backing and the active
  // editor) are laid out by LayoutActiveEditor. The host is the single source
  // of truth for the editor's on-screen rect.
  var HostW:=R.Width-FGridLineWidth/2;
  var HostH:=R.Height-FGridLineWidth/2;

  FEditorHost.SetBounds(R.Left+FGridLineWidth/4, R.Top+FGridLineWidth/4, HostW, HostH);
  FEditorHost.Visible:=True;
  FEditorHost.BringToFront;

  LayoutActiveEditor(HostW, HostH);
end;

procedure TMultiHeaderGrid.LayoutActiveEditor(AWidth, AHeight: Single);
begin
  // Size and place the shared TMemo. ApplyStyle zeroes the memo's own inset, so
  // its text starts at its top-left corner; offset the control by the cell's
  // text inset so the editor text lands exactly where the cell text was. The
  // control extends to the host edges (host clips any overflow).
  if FEditor<>nil then begin
    var Style:=EffectiveCellStyle(FEditCol, FEditRow);

    FEditor.TextSettings.HorzAlign  :=Style.TextHAlignment;
    FEditor.TextSettings.Font.Family:=Style.FontName;
    FEditor.TextSettings.Font.Size  :=Style.FontSize;
    FEditor.TextSettings.Font.Style :=Style.FontStyle;
    FEditor.WordWrap:=Style.WordWrap;

    // Vertical placement: the memo always lays text from its own top, so to
    // honour TextVAlignment we shift the whole control down by the slack
    // between the cell height and the text height. Center -> half the slack,
    // Trailing -> all of it, Leading -> none.

    var TextH:=FEditor.ContentBounds.Height;
    var Slack:=AHeight - TextH;
    if Slack<0 then Slack:=0;

    var OffY: Single;
    case Style.TextVAlignment of
      TTextAlign.Center:   OffY:=Slack/2;
      TTextAlign.Trailing: OffY:=Slack;
    else
      OffY:=0; // Leading
    end;

    // The -3 / +8 keep the existing single-line baseline nudge and overscan.
    FEditor.SetBounds(0, OffY-3, AWidth, AHeight-OffY+8);

    FEditor.Visible:=True;
    FEditor.BringToFront;
  end;
end;

procedure TMultiHeaderGrid.HideEditor;
begin
  // Hiding the host hides its FEditorBack child too. Descendants override to
  // also hide any extra editors they parented into the host.
  if FEditorHost<>nil then FEditorHost.Visible:=False;
  var Ed:=ActiveEditorControl;
  if Ed<>nil then Ed.Visible:=False;
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
  // Edit a merged cell through its top-left anchor, where the value lives.
  var MergedCell: TMergedCell;
  if IsMergedCell(ACol, ARow, MergedCell) then begin
    ACol:=MergedCell.Col;
    ARow:=MergedCell.Row;
  end;

  if Assigned(FOnStartEditing) then
    FOnStartEditing(Self, ACol, ARow, InitialChar);

  if not CanEditCell(ACol,ARow) then Exit;

  if FEditing then CommitEditing;

  EnsureEditor(ColTextHAlignment[ACol]);
  FEditCol:=ACol;
  FEditRow:=ARow;
  FEditing:=True;
  FEditWidenAnchor:=-1; // no transient widen yet this edit

  // The factory picks the control for this cell: the shared TMemo in the base,
  // a typed editor in the DB grid.
  var Ed:=PrepareCellEditor(ACol, ARow);
  if Ed=nil then Exit;

  // If the chosen editor needs more room than the (possibly merged) cell
  // offers, widen the anchor column so the control isn't clipped.
  WidenColForEditor(ACol, ARow);

  // If this cell can't be modified the editor opens read-only purely so its
  // text can be selected and copied: never seed it with a typed character
  // (that would look like an edit), always load the existing value.
  var Modifiable:=CellIsModifiable(ACol, ARow);
  if not Modifiable then InitialChar:=#0;

  LoadEditorValue(ACol, ARow, InitialChar);

  // Read-only cells: editor allows copy but not changes.
  ApplyEditorReadOnly(Ed, Modifiable);

  // Hide Value in edited cell
  Invalidate;

  // Fit the row to the just-loaded content. OnChange (MemoEditorTextChanged)
  // only fires when the text actually changes, so reopening a cell whose value
  // equals the memo's current text would otherwise skip the row re-fit and the
  // row height wouldn't update. Call it explicitly for the shared memo.
  if ActiveEditorControl=FEditor then
    MemoEditorTextChanged(FEditor);

  // PositionEditor sizes and shows the host; only focus and caret remain.
  PositionEditor;
  if Ed.CanFocus then
    Ed.SetFocus;
  PlaceEditorCaretAtEnd(Ed);
end;

procedure TMultiHeaderGrid.WidenColForEditor(ACol, ARow: Integer);
// Widens the (anchor) column so EditorMinColWidth(ACol,ARow) fits. Safe to call
// repeatedly while editing - it only ever grows the column, and re-positions
// the editor over the resized cell. A merged cell grows its anchor column by
// the shortfall across the whole span.
begin
  var MinW:=EditorMinColWidth(ACol, ARow);
  if MinW<=0 then Exit;

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
    // Remember the anchor and its width the first time we widen during this
    // edit, so editing-end can shrink back to fit (but not below this).
    if FEditing and (FEditWidenAnchor<0) then begin
      FEditWidenAnchor:=Anchor;
      FEditWidenStartW:=GetColWidth(Anchor);
    end;
    // SetColWidth clamps to the column's MaxWidth, so an explicit MaxWidth
    // still wins.
    ColWidths[Anchor]:=Trunc(GetColWidth(Anchor)+(MinW-SpanW));
    // Widening can push the cell's right edge past the viewport; bring the
    // cell fully back into view so the editor isn't drawn half outside.
    FSelectedCell:=Point(FEditCol, FEditRow);
    ScrollToSelectedCell;
    if FEditing then PositionEditor; // re-place the editor over the resized cell
  end;
end;

procedure TMultiHeaderGrid.PlaceEditorCaretAtEnd(Ed: TControl);
begin
  // Only the base TMemo here; the DB grid overrides for its typed editors.
  // Clear any inherited selection and put the caret at the end, keeping SelStart
  // consistent with the caret so a fresh edit never starts with stale selection
  // or a mismatched SelStart.
  if Ed is TMemo then begin
    var M:=TMemo(Ed);
    M.GoToTextEnd;
    M.SelLength:=0;
    M.SelStart:=M.Text.Length; // align SelStart with the caret position
  end;
end;

procedure TMultiHeaderGrid.FinishEditColWidth(ATargetW: Integer);
// Called when editing ends. If the anchor column was transiently widened for
// the editor:
//   - ATargetW<0 (cancel): always restore the pre-edit width.
//   - otherwise (commit): if KeepEditorWidenedColumn is True the enlarged width
//     is kept; if False (default) the column is restored to its pre-edit width.
// This avoids the inconsistent "shrink partway back" behaviour: the column
// either fully returns to its original width or fully keeps the enlargement.
begin
  if FEditWidenAnchor<0 then Exit; // never widened this edit

  var Anchor:=FEditWidenAnchor;
  FEditWidenAnchor:=-1; // consume

  // Keep the enlarged width on a normal commit when requested.
  if (ATargetW>=0) and FKeepEditorWidenedColumn then Exit;

  // Otherwise restore the pre-edit width.
  if FEditWidenStartW<GetColWidth(Anchor) then
    ColWidths[Anchor]:=FEditWidenStartW; // SetColWidth clamps to Min/MaxWidth
end;

procedure TMultiHeaderGrid.CommitEditing;
begin
  if not FEditing then Exit;
  FEditing:=False; // clear first so EditorExit re-entry is a no-op
  var Col:=FEditCol;
  var Row:=FEditRow;
  HideEditor;

  // Single persistence point. The base writes the editor text into Cells[]
  // (firing DoSetCellText); the DB grid override assigns the bound field with
  // the correct type. The editor text is not also pushed through Cells[], which
  // would double-write and force an unparsed string into a typed field.
  // A read-only cell's editor was opened only for copy, so nothing is written.
  if CellIsModifiable(Col, Row) then
    CommitEditorValue(Col, Row);

  // Editing widened the column to fit the in-progress value. By default the
  // column now returns to its pre-edit width; set KeepEditorWidenedColumn to
  // keep the enlargement instead. (ATargetW>=0 selects the commit branch.)
  FinishEditColWidth(Trunc(EditorMinColWidth(Col, Row)));

  FActiveEditor:=nil;
  if CanFocus then SetFocus;
end;

procedure TMultiHeaderGrid.CancelEditing;
begin
  if not FEditing then Exit;
  FEditing:=False;
  HideEditor;
  FinishEditColWidth(-1); // undo any transient widen
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
    var Hi:=IfThen(FColData[StartCol+i].MaxWidth>0,FColData[StartCol+i].MaxWidth,MaxInt);
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

procedure TMultiHeaderGrid.EnsureRowAvailable(ARow: Integer);
begin
  // A base grid owns all its rows; a row-provider descendant overrides this.
end;

function TMultiHeaderGrid.InLayout: Boolean;
begin
  Result:=FInLayout;
end;

function TMultiHeaderGrid.ColScanRowLimit: Integer;
begin
  Result:=FRowCount; // base grids hold all rows
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
  if (Index>=0) and (Index<FColCount) and (FColData[Index].WordWrap<>Value) then begin
    FColData[Index].WordWrap:=Value;
    // This column's cells now wrap (or stop wrapping), changing row heights.
    // Skipped during a bulk rebuild (it applies wrap to every column in a loop
    // and sizes once at the end).
    if not FSuppressAutoSize then begin
      AutoSizeRows;
      UpdateSize;
    end;
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
    // Row heights depend on wrap, so re-fit them now. Skipped during a
    // bulk rebuild, which does its own single sizing pass at the end.
    if not FSuppressAutoSize and Value then begin
      AutoSizeRows;
      UpdateSize;
    end;
    Invalidate;
  end;
end;

procedure TMultiHeaderGrid.SetConservativeWrap(const Value: Boolean);
begin
  if FConservativeWrap<>Value then begin
    FConservativeWrap:=Value;
    // Only matters while wrapping; re-fit so the new policy takes effect.
    if not FSuppressAutoSize and GridHaveWordWrap then begin
      AutoSize;
    end;
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
    FDisplayFormat:=C.FDisplayFormat;
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
    // Colour is read live per cell from FColMap (the same column instance) by
    // DoGetCellStyle, and it does not affect the column set, order, widths or
    // header layout. So a full rebuild is unnecessary - just repaint.
    var G:=Grid;
    if G<>nil then G.Invalidate;
  end;
end;

procedure TMHGColumn.SetDisplayFormat(const Value: string);
begin
  if FDisplayFormat<>Value then begin
    FDisplayFormat:=Value;
    // Only the rendered text of this column's data cells changes - not the
    // column set/order/geometry. Drop the cached strings so they re-render
    // with the new format, then repaint. No header rebuild needed.
    var G:=Grid;
    if G<>nil then G.RefreshCells;
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
  // Any add/remove/reorder/property change rebuilds the grid header. Defer and
  // coalesce: a burst of changes (e.g. setting Min/MaxWidth on several columns)
  // then costs one rebuild, flushed before the next paint.
  if FGrid<>nil then
    FGrid.InvalidateTable;
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
  if FGrid<>nil then FGrid.InvalidateTable;
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
  FConfirmDelete:=True;
  FInsertRowIndex:=-1;
  FDateFormat:='DD.MM.YYYY';
  FTimeFormat:='HH:NN:SS';
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
  // DoGetCellText re-reads the current field values. The style cache for this
  // row is dropped with it.
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

function TMultiHeaderDBGrid.FieldToText(ACol: TMHGColumn; AField: TField): string;
begin
  if AField=nil then Exit('');
  if AField.IsNull then Exit('');

  // Only date/time-family fields are formatted; everything else keeps the
  // field's native AsString (numbers, strings, booleans, etc.).
  case AField.DataType of
    ftDate, ftTime, ftDateTime, ftTimeStamp, ftOraTimeStamp: ;
  else
    Exit(AField.AsString);
  end;

  var DT:=AField.AsDateTime;

  // A per-column DisplayFormat, when set, overrides the grid defaults and is
  // applied verbatim regardless of the date/time sub-kind.
  if (ACol<>nil) and (ACol.DisplayFormat<>'') then
    Exit(FormatDateTime(ACol.DisplayFormat,DT));

  // Otherwise build the format from the grid defaults per field sub-kind.
  var Fmt: string;
  case AField.DataType of
    ftDate:
      Fmt:=FDateFormat;
    ftTime:
      Fmt:=FTimeFormat;
  else
    // ftDateTime / ftTimeStamp / ftOraTimeStamp: date+time, unless the column
    // is configured as date-only.
    if (ACol<>nil) and (ACol.DateTimeEditor=dteDate) then
      Fmt:=FDateFormat
    else
      Fmt:=Trim(FDateFormat+' '+FTimeFormat);
  end;

  // Empty format falls back to the field's own rendering.
  if Fmt='' then Exit(AField.AsString);
  Result:=FormatDateTime(Fmt,DT);
end;

procedure TMultiHeaderDBGrid.SetDateFormat(const Value: string);
begin
  if FDateFormat<>Value then begin
    FDateFormat:=Value;
    RefreshCells; // cached strings used the old format
  end;
end;

procedure TMultiHeaderDBGrid.SetTimeFormat(const Value: string);
begin
  if FTimeFormat<>Value then begin
    FTimeFormat:=Value;
    RefreshCells;
  end;
end;

procedure TMultiHeaderDBGrid.SetViewTop(const Value: Integer);
begin
  // Fetch more rows before we clamp to the current last row,
  // otherwise the viewport can never grow past materialized rows.
  EnsureRowsForViewport(Value);

  inherited;
end;

procedure TMultiHeaderDBGrid.RefreshCells;
begin
  FCellTexts.Clear;
  Invalidate;
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
  var Fields: TArray<TField>;
  var Cols:=ResolveColumns(Fields); // Cols[i] pairs with Fields[i]

  SetLength(Result.Cells,Length(Fields));

  if ARow=CurrentRowIndex then begin
    // Current record buffer (also the pending insert row).
    for var i:=0 to High(Fields) do begin
      Result.Cells[i]:=FieldToText(Cols[i],Fields[i]);
    end;
  end else if (DS.State=dsInsert) and (FInsertRowIndex>=0) then begin
    // VCL buffer model: read the posted record for this grid row via the
    // DataLink's active buffer, which already accounts for the phantom insert
    // row's slot - no RecNo arithmetic. Map grid row -> buffer index relative
    // to the pending record's ActiveRecord position.
    var BufIdx:=FDataLink.ActiveRecord+(ARow-FInsertRowIndex);
    if (BufIdx>=0) and (BufIdx<FDataLink.BufferCount) then begin
      var Old:=FDataLink.ActiveRecord;
      try
        FDataLink.ActiveRecord:=BufIdx;
        for var i:=0 to High(Fields) do
          Result.Cells[i]:=FieldToText(Cols[i],Fields[i]);
      finally
        FDataLink.ActiveRecord:=Old;
      end;
    end else begin
      Exit;
    end;
  end else begin
    DS.DisableControls;
    try
      var Bookmark:=DS.Bookmark;
      try
        DS.RecNo:=ARow+1;
        for var i:=0 to High(Fields) do
          Result.Cells[i]:=FieldToText(Cols[i],Fields[i]);
      finally
        if DS.BookmarkValid(Bookmark) then
          DS.Bookmark:=Bookmark;
      end;
    finally
      DS.EnableControls;
    end;
  end;

  Result.LastUsage:=Now;

  FCellTexts.AddOrSetValue(ARow,Result);

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
    // not flatten them back to the fixed build-time default.
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
      FSuppressAutoSize:=True;
      try
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
          ColWordWrap[i]:=Col.WordWrap;
          ColTextHAlignment[i]:=Col.Alignment;
          ColTextVAlignment[i]:=Col.VertAlignment;
        end;
      finally
        FSuppressAutoSize:=False;
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

    // Rows just got default heights from UpdateRowCount; if the grid wraps,
    // fit them to content now.
    if GridHaveWordWrap then begin
      AutoSizeRows;
      UpdateSize;
    end;
  finally
    FRebuildingHeader:=False;
  end;

  // A real rebuild just happened, so any deferred request is now satisfied.
  FRebuildPending:=False;
end;

procedure TMultiHeaderDBGrid.InvalidateTable;
begin
  // Mark the layout dirty and schedule a repaint. The actual rebuild is
  // deferred to EnsureLayout (called from Paint), which coalesces a burst of
  // change notifications into a single ResetTable.
  FRebuildPending:=True;
  Invalidate;
end;

procedure TMultiHeaderDBGrid.EnsureLayout;
begin
  if FRebuildPending and (not FRebuildingHeader) then
    ResetTable; // clears FRebuildPending
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
  // A close or (re)open invalidates cached on-demand state.
  FFetchesOnDemand:=False;
  FNativeAllFetched:=False;

  if not ((DataSet<>nil) and DataSet.Active) then begin
    // Closed: keep current columns; just rebuild (shows placeholder if empty).
    ResetTable;
    Exit;
  end;

  FFetchesOnDemand:=DataSetFetchesOnDemand(DataSet); // probe once per open

  var Sig:=DatasetSignature;
  // Auto-create columns only when there are none AND this is a source/table we
  // have not auto-created for yet. This fills a fresh grid or a newly opened
  // table, but never resurrects columns the user deleted on the same table.
  if Sig<>FLastAutoSig then begin
    FLastAutoSig:=Sig;
    // The dataset structure changed (different table/query), so any existing
    // columns were built against the OLD field set and are now meaningless:
    // fields shared by both tables would otherwise keep their old (earlier)
    // collection index and stay in front of the new table's columns. Clear
    // them so AutoCreateColumns rebuilds in the NEW table's field order.
    FColumns.BeginUpdate;
    try
      FColumns.Clear;
    finally
      FColumns.EndUpdate;
    end;
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

function TMultiHeaderDBGrid.DataSetFetchesOnDemand(DS: TDataSet): Boolean;
// FireDAC datasets expose FetchOptions.Mode (TFDFetchMode = fmManual, fmOnDemand,
// fmAll, fmExactRecsMax). fmManual/fmOnDemand fetch rows lazily, so RecordCount is
// only the rows fetched so far; fmAll/fmExactRecsMax materialize everything up
// front. Probe it via RTTI so the grid keeps no compile-time FireDAC dependency.
// Any dataset without FetchOptions (or where the probe fails) is treated as not
// on-demand.
begin
  Result:=False;
  if (DS=nil) or (not DS.Active) then Exit;
  try
    var Ctx:=TRttiContext.Create;
    try
      var T:=Ctx.GetType(DS.ClassType);
      if T=nil then Exit;
      var FetchProp:=T.GetProperty('FetchOptions');
      if FetchProp=nil then Exit;               // not a FireDAC dataset
      var FetchVal:=FetchProp.GetValue(DS);
      if FetchVal.IsEmpty or (not FetchVal.IsObject) then Exit;
      var FetchObj:=FetchVal.AsObject;
      if FetchObj=nil then Exit;
      var ModeProp:=Ctx.GetType(FetchObj.ClassType).GetProperty('Mode');
      if ModeProp=nil then Exit;
      // TFDFetchMode = (fmManual=0, fmOnDemand=1, fmAll=2, fmExactRecsMax=3).
      // fmManual and fmOnDemand fetch lazily (RecordCount is only rows fetched so
      // far); fmAll and fmExactRecsMax materialize the whole set up front.
      Result:=ModeProp.GetValue(FetchObj).AsOrdinal<2;
    finally
      Ctx.Free;
    end;
  except
    Result:=False; // probe failed - assume not on-demand
  end;
end;

function TMultiHeaderDBGrid.CurrentRowIndex: Integer;
begin
  Result:=-1;
  var DS:=DataSet;
  if (DS=nil) or (not DS.Active) then Exit;

  if (DS.State=dsInsert) and (FInsertRowIndex>=0) then Exit(FInsertRowIndex);
  if DS.IsEmpty then Exit;

  if DS.RecNo>0 then Result:=DS.RecNo-1;
end;

procedure TMultiHeaderDBGrid.UpdateRowCount;
begin
  if FNativeFetching then Exit;     // ignore callbacks during cursor stepping
  if FInUpdateRowCount then Exit;  // re-entrancy guard (nested DataSetChanged)
  FInUpdateRowCount:=True;
  try
    var DS:=DataSet;
    if (DS=nil) or (not DS.Active) then begin
      RowCount:=0;
      Exit;
    end;

    FCellTexts.Clear;

    // VCL buffer model: size the DataLink window to cover the visible rows so
    // ActiveRecord can read any visible record (incl. the pending insert).
    var VisRows:=Trunc(ViewPortDataHeight/Max(1,DefaultRowHeight))+2;
    if VisRows<1 then VisRows:=1;
    if FDataLink.BufferCount<>VisRows then FDataLink.BufferCount:=VisRows;

    var Inserting:=(DS.State=dsInsert);
    if not Inserting and (FInsertRowIndex>=0) then FInsertRowIndex:=-1;

    // Native path. VCL buffer model: during dsInsert the pending record is an
    // extra visible row; grow RowCount by one (through the setter so FRowData
    // stays consistent). FInsertRowIndex holds its grid position.
    if Inserting then begin
      if FInsertRowIndex<0 then FInsertRowIndex:=DS.RecordCount;
      RowCount:=DS.RecordCount+1;
    end else
      RowCount:=DS.RecordCount;

    SyncSelectedRowFromDataSet;
    Invalidate;
  finally
    FInUpdateRowCount:=False;
  end;
end;

procedure TMultiHeaderDBGrid.SyncSelectedRowFromDataSet;
begin
  if FNativeFetching then Exit;

  var DS:=DataSet;
  if (DS=nil) or (not DS.Active) then Exit;
  if DS.IsEmpty and (DS.State<>dsInsert) then Exit;

  var NewRow:=CurrentRowIndex;
  if NewRow<0 then Exit;
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
  if FNativeFetching then Exit;

  var DS:=DataSet;
  if (DS=nil) or (not DS.Active) or DS.IsEmpty then Exit;

  var Idx:=CurrentRowIndex;
  if Idx<0 then Exit;

  FCellTexts.Remove(Idx);
  Invalidate;
end;

procedure TMultiHeaderDBGrid.DoSelectCell;
begin
  inherited;

  if FUpdatingRow then Exit;

  var DS:=DataSet;
  if (DS=nil) or (not DS.Active) then Exit;

  // Leaving a pending insert row commits it (VCL behaviour). If the post fails
  // (e.g. a required field is still empty) keep the caret on the row so the
  // user can fix it, and let the error surface. Escape clears the selection
  // (Row < 0) - that cancels the pending record instead of posting it.
  if (DS.State=dsInsert) and (FInsertRowIndex>=0) and (Row<>FInsertRowIndex) then begin
    if Row<0 then begin
      if FEditing then CancelEditing
      else begin FInsertRowIndex:=-1; DS.Cancel; end;
      Exit;
    end;
    if FEditing then CommitEditing;
    try
      DS.Post;
    except
      FUpdatingRow:=True;
      try
        Row:=FInsertRowIndex;   // snap back to the still-open insert row
      finally
        FUpdatingRow:=False;
      end;
      raise;
    end;
    FInsertRowIndex:=-1;
    // RowCount collapses on the DataSetChanged that follows the post.
  end;

  if (Row<0) or (Row>=RowCount) then Exit;
  if CurrentRowIndex=Row then Exit;

  FUpdatingRow:=True;
  try
    var RecRow:=Row;
    if (DS.State=dsInsert) and (FInsertRowIndex>=0) and (Row>FInsertRowIndex) then
      RecRow:=Row-1;
    try DS.RecNo:=RecRow+1; except end;
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

procedure TMultiHeaderDBGrid.DoGridScroll;
begin
  inherited; // editor glue + OnGridScroll

  var LastVisible:=RowAtHeightCoord(ViewBottom);
  if LastVisible<0 then LastVisible:=RowCount-1;

  // On-demand cursor with a partial RecordCount: pull the next packet(s) when
  // the last visible row reaches the known count.
  if LastVisible>=RowCount-1 then EnsureRowAvailable(LastVisible+1);
end;

procedure TMultiHeaderDBGrid.GrowNativeTo(ATargetRow: Integer);
// Native-path growth shared by the navigation/scroll/End paths. Steps the cursor
// forward from the fetched edge with Next - bounded by Eof - until ATargetRow is
// materialized (ATargetRow < 0 means "to Eof"), then adopts the grown
// RecordCount. Avoids DS.Last (which can fetch the whole set at once on a
// partial cursor). Captures the count while still at the fetched position, since
// restoring the bookmark can revert RecordCount on some FireDAC cursors.
begin
  var DS:=DataSet;
  if (DS=nil) or (not DS.Active) then Exit;
  if FNativeAllFetched then Exit;
  if FNativeFetching then Exit;
  if (ATargetRow>=0) and (ATargetRow<RowCount) then Exit; // already known

  FNativeFetching:=True;
  try
    try
      DS.DisableControls;
      var FinalCount:=RowCount;
      try
        var BM:=DS.Bookmark;
        try
          if DS.RecordCount>0 then DS.RecNo:=DS.RecordCount; // resume at fetched edge
          var Guard:=0;
          while (not DS.Eof) and ((ATargetRow<0) or (DS.RecordCount<=ATargetRow))
                and (not FEndFetchCancelled) do begin // stop promptly when End fetch is cancelled
            DS.Next;
            Inc(Guard);
            if Guard>10000000 then Break;
          end;
          if DS.RecordCount>=0 then FinalCount:=DS.RecordCount;
          if DS.Eof then FNativeAllFetched:=True;
        finally
          if DS.BookmarkValid(BM) then DS.Bookmark:=BM;
        end;
      finally
        DS.EnableControls;
      end;
      if RowCount<>FinalCount then RowCount:=FinalCount;
    except
      // ignore - navigation will simply clamp to the current RowCount
    end;
  finally
    FNativeFetching:=False;
  end;
end;

function TMultiHeaderDBGrid.EndIsOnDemand: Boolean;
// End must page on demand (the cancellable, repainting loop) whenever the true
// end is not yet known: a cursor that fetches lazily (FetchOptions.Mode is
// fmManual/fmOnDemand) and is not yet fully fetched. A fully-materialized cursor
// returns False, so End jumps straight to the last row.
begin
  Result:=FFetchesOnDemand and (not FNativeAllFetched);
end;

procedure TMultiHeaderDBGrid.RunEndFetch;
// Page down repeatedly - reusing the base vkNext handling (grow + move + scroll +
// OnSelectCell) - until the selection stops advancing (Eof) or the user cancels.
// The synthetic presses are marked FEndSynthetic so the cancel check in KeyDown
// ignores them; any real key or mouse click sets FEndFetchCancelled.
begin
  FEndFetchCancelled:=False;
  FEndFetching:=True;
  try
    var PrevRow:=-1;
    FPainted:=True;               // first synthetic press is allowed through the paint gate
    while (FSelectedCell.Y<>PrevRow) and EndIsOnDemand and (not FEndFetchCancelled)
          and FPainted do begin   // only advance once the previous Paint has completed
      PrevRow:=FSelectedCell.Y;
      FEndSynthetic:=True;        // mark the next press as our own (not user input)
      var K: Word:=vkNext; var KC: WideChar:=#0;
      inherited KeyDown(K,KC,[]); // one page down via the base navigation (clears FPainted)
      FEndSynthetic:=False;
      Repaint;                      // request a paint of the new rows
      // Yield once so the requested paint can run before the next press; this
      // also delivers any pending user key/click that sets FEndFetchCancelled.
      Application.ProcessMessages;
    end;
  finally
    FEndFetching:=False;
    FEndSynthetic:=False;
  end;
end;

procedure TMultiHeaderDBGrid.KeyDown(var Key: Word; var KeyChar: WideChar; Shift: TShiftState);
begin
  // While an End fetch is running, swallow every real key (including another End,
  // which would otherwise re-enter RunEndFetch and corrupt the loop state) except
  // our own synthetic Page Down presses. Any swallowed key cancels the fetch.
  if FEndFetching and (not FEndSynthetic) then begin
    FEndFetchCancelled:=True;
    Key:=0; KeyChar:=#0;
    Exit;
  end;

  // On-demand End: let the base run first (fires OnKeyDown and jumps to the last
  // row fetched so far), then - if the user did not consume the key - page on to
  // the true Eof. A fully-loaded cursor needs no extension; the base jump suffices.
  if (Key=vkEnd) and (not FEndSynthetic) and FPainted and IsFocused and EndIsOnDemand then begin
    inherited;
    if Key<>0 then RunEndFetch; // not consumed by a user OnKeyDown handler
    Key:=0; KeyChar:=#0;
    Exit;
  end;

  inherited;
end;

procedure TMultiHeaderDBGrid.MouseDown(Button: TMouseButton; Shift: TShiftState; X, Y: Single);
begin
  // A mouse click cancels an in-progress End fetch.
  if FEndFetching then begin
    FEndFetchCancelled:=True;
    Exit;
  end;
  inherited;
end;

procedure TMultiHeaderDBGrid.EnsureRowAvailable(ARow: Integer);
// Ensures grid row ARow is reachable when navigating down: step the cursor toward
// ARow (GrowNativeTo) - an on-demand FireDAC cursor reports only the fetched
// RecordCount until rows are pulled. Synchronous - KeyDown is not a paint, so
// resizing here is safe.
begin
  var DS:=DataSet;
  if (DS=nil) or (not DS.Active) then Exit;
  GrowNativeTo(ARow);
end;

procedure TMultiHeaderDBGrid.EnsureRowsForViewport(ATargetTop: Integer);
// From SetViewTop before clamping. If the requested top reaches the bottom of
// fetched rows and more may exist, pull packet(s) and grow RowCount
// synchronously (scroll event, not paint) so the clamp sees the new extent.
begin
  if FNativeAllFetched then Exit; // whole set loaded; nothing to grow
  if Length(FRowData)=0 then Exit;

  // If the requested top reaches the last known row, pull more so the clamp in
  // SetViewTop can move further.
  var LastTop:=FRowData[High(FRowData)].Top;
  if ATargetTop+ViewCellsHeight>=LastTop then begin
    var TargetRow:=RowCount-1+Max(1,(ViewCellsHeight div Max(1,DefaultRowHeight)));
    GrowNativeTo(TargetRow);
  end;
end;

function TMultiHeaderDBGrid.ColScanRowLimit: Integer;
begin
  // Bound the column-width scan to rows fetched so far (C-fit-fetched) -
  // RowCount is the fetched RecordCount.
  Result:=Min(FRowCount, RowCount);
  if Result<0 then Result:=0;
end;

function TMultiHeaderDBGrid.EditorKindForField(Field: TField): TMHGEditorKind;
begin
  Result:=EditorKindForField(Field, False);
end;

function TMultiHeaderDBGrid.EditorKindForField(Field: TField;
  AIgnoreReadOnly: Boolean): TMHGEditorKind;
begin
  Result:=ekNone;
  if Field=nil then Exit;

  // Read-only and non-data (calculated/internal-calc) fields are never edited.
  // With AIgnoreReadOnly the read-only gate is skipped so the cell can still
  // open the type's editor in read-only mode (copy only).
  if Field.ReadOnly and (not AIgnoreReadOnly) then Exit;
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

    // Numeric types: edited through a numeric-filtered TEdit. A plain edit is
    // used (not TNumberBox) so the box can hold '' to represent SQL NULL.
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
    ftByte:     begin
      AMin:=0;
      AMax:=255;
    end;
    ftShortint: begin
      AMin:=-128;
      AMax:=127;
    end;
    ftWord:     begin
      AMin:=0;
      AMax:=65535;
    end;
    ftSmallint: begin
      AMin:=-32768;
      AMax:=32767;
    end;
    ftLongWord: begin
      AMin:=0;
      AMax:=4294967295;
    end;
    ftInteger:  begin
      AMin:=-2147483648;
      AMax:=2147483647;
    end;
    // 64-bit integers exceed Double's exact-integer range; use the wide
    // finite limit rather than the exact (unrepresentable) Int64 bounds.
    ftLargeint: begin
      AMin:=-FLOAT_LIMIT;
      AMax:=FLOAT_LIMIT;
    end;
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

procedure TMultiHeaderDBGrid.OnDateEditorClear(Sender: TObject);
begin
  FDateEditor.IsEmpty:=True;
  EditorExit(Sender);
end;

procedure TMultiHeaderDBGrid.OnTimeEditorClear(Sender: TObject);
begin
  if FEditing then begin
    if IsCompositeEditor(Sender) then begin
      FDateEditor.IsEmpty:=True;
      FTimeEditor.IsEmpty:=True;
    end else begin
      FTimeEditor.IsEmpty:=True;
    end;
  end;
  EditorExit(Sender);
end;

procedure TMultiHeaderDBGrid.EnsureTypedEditors;
begin
  if FComboEditor=nil then begin
    FComboEditor:=TComboBox.Create(Self);
    WireEditor(FComboEditor);
    FComboEditor.StyleLookup:='mhg_editor_combo';
  end;
  if FNumberEditor=nil then begin
    FNumberEditor:=TEdit.Create(Self);
    WireEditor(FNumberEditor);
    FNumberEditor.OnKeyDown:=NumberEditorKeyDown;
    FNumberEditor.OnChangeTracking:=NumberEditorChangeTracking;
    FNumberEditor.StyleLookup:='mhg_editor_number';
  end;
  if FDateEditor=nil then begin
    FDateEditor:=TDateEdit.Create(Self);
    FDateEditor.Format:='dd.MM.yyyy';
    WireEditor(FDateEditor);
    FDateEditor.OnKeyDown:=DateTimeEditorKeyDown;
    FDateEditor.StyleLookup:='mhg_editor_date';

    var Button:=TButton(FDateEditor.FindStyleResource('mhgclearbutton'));
    if Assigned(Button) then begin
      Button.OnClick:=OnDateEditorClear;
    end;
  end;
  if FTimeEditor=nil then begin
    FTimeEditor:=TTimeEdit.Create(Self);
    FTimeEditor.Format:='HH:nn:ss';
    WireEditor(FTimeEditor);
    FTimeEditor.OnKeyDown:=DateTimeEditorKeyDown;
    FTimeEditor.StyleLookup:='mhg_editor_time';

    var Button:=TButton(FTimeEditor.FindStyleResource('mhgclearbutton'));
    if Assigned(Button) then begin
      Button.OnClick:=OnTimeEditorClear;
    end;
  end;

  // Capture the editors' default single-line height once, before any cell
  // layout stretches them, so a tall cell can be height-capped reliably.
  if FTypedEditorHeight<=0 then
    FTypedEditorHeight:=FComboEditor.Height;
end;

procedure TMultiHeaderDBGrid.NumberEditorKeyDown(Sender: TObject; var Key: Word;
  var KeyChar: Char; Shift: TShiftState);
// Filters keystrokes so the TEdit only accepts characters that can build a
// number: digits, a sign, and a decimal separator (integer fields drop the
// separator). Control keys (navigation, backspace, Ctrl+combos) pass through,
// as does everything the shared EditorKeyDown handles (Enter/Tab/Esc). An empty
// box is allowed - that's how NULL is entered.
begin
  if KeyChar<>#0 then begin
    var Allowed:=CharInSet(KeyChar, ['0'..'9']) or
                 (KeyChar='-') or (KeyChar='+');
    // Decimal separator only for non-integer fields, and only one of them.
    if not Allowed and CharInSet(KeyChar, ['.', ',']) and
       (FEditField<>nil) and not FieldIsInteger(FEditField) then
      Allowed:=(Pos('.', FNumberEditor.Text)=0) and (Pos(',', FNumberEditor.Text)=0);

    if not Allowed then begin
      KeyChar:=#0; // swallow the character
      Key:=0;
      Exit;        // don't forward a rejected char to the shared handler
    end;
  end;

  // Let the shared handler process Enter/Tab/Esc and anything we allowed.
  EditorKeyDown(Sender, Key, KeyChar, Shift);
end;

function TMultiHeaderDBGrid.IsCompositeEditor(C: TObject): Boolean;
begin
  Result:=(FEditField<>nil) and (EditorKindForField(FEditField)=ekDateTime)
          and ((C=FDateEditor) or (C=FTimeEditor));
end;

type
  // Surfaces the protected format-part navigation of TCustomDateTimeEdit so the
  // composite editor can tell when the caret is on the first/last section
  // (e.g. the year of dd.MM.yyyy) and hand off to the sibling control.
  TDateTimeEditAccess = class(TCustomDateTimeEdit);

procedure TMultiHeaderDBGrid.DateTimeGoToFirstPart(Ed: TCustomDateTimeEdit);
// Walk left until GoToPreviousFormatPart stops changing the focused part.
var
  Before, After: TDTFormatPart;
begin
  if Ed=nil then Exit;
  var Cracker:=TDateTimeEditAccess(Ed);
  while Cracker.FindCurrentFormatPart(Before) do begin
    Cracker.GoToPreviousFormatPart;
    if not Cracker.FindCurrentFormatPart(After) then Break;
    if CompareMem(@Before, @After, SizeOf(TDTFormatPart)) then Break;
  end;
end;

procedure TMultiHeaderDBGrid.DateTimeGoToLastPart(Ed: TCustomDateTimeEdit);
// Walk right until GoToNextFormatPart stops changing the focused part.
var
  Before, After: TDTFormatPart;
begin
  if Ed=nil then Exit;
  var Cracker:=TDateTimeEditAccess(Ed);
  while Cracker.FindCurrentFormatPart(Before) do begin
    Cracker.GoToNextFormatPart;
    if not Cracker.FindCurrentFormatPart(After) then Break;
    if CompareMem(@Before, @After, SizeOf(TDTFormatPart)) then Break;
  end;
end;

procedure TMultiHeaderDBGrid.FocusCompositePart(ToTime: Boolean);
// Moves focus between the two halves of the ftDateTime composite. ToTime=True
// jumps to the time control and selects its first format part; ToTime=False
// jumps to the date control and selects its last part. FActiveEditor is updated
// so commit/layout keep treating the focused half as active.
begin
  if ToTime then begin
    if (FTimeEditor=nil) or not FTimeEditor.CanFocus then Exit;
    FActiveEditor:=FTimeEditor;
    FTimeEditor.SetFocus;
    DateTimeGoToFirstPart(FTimeEditor);  // first section of the time control
  end
  else begin
    if (FDateEditor=nil) or not FDateEditor.CanFocus then Exit;
    FActiveEditor:=FDateEditor;
    FDateEditor.SetFocus;
    DateTimeGoToLastPart(FDateEditor);   // last section of the date control
  end;
end;

function TMultiHeaderDBGrid.DateTimePartIsLast(
  Ed: TCustomDateTimeEdit): Boolean;
// True when the focused format part is the rightmost one (no next part to move
// to). Detected by probing: remember the current part, step forward, and see
// whether the current part actually changed; if not, we were already at the
// end. Any forward step is undone so probing never alters the visible caret.
var
  Before, After: TDTFormatPart;
  HadBefore: Boolean;
begin
  Result:=False;
  if Ed=nil then Exit;
  var Cracker:=TDateTimeEditAccess(Ed);
  HadBefore:=Cracker.FindCurrentFormatPart(Before);
  if not HadBefore then Exit;        // nothing focused -> treat as "not last"
  Cracker.GoToNextFormatPart;
  if Cracker.FindCurrentFormatPart(After) then begin
    // CompareMem keeps this independent of TDTFormatPart's exact fields, which
    // differ across FMX versions.
    Result:=CompareMem(@Before, @After, SizeOf(TDTFormatPart));
    if not Result then
      Cracker.GoToPreviousFormatPart; // undo the probe step
  end
  else
    Result:=True;                     // no part after the move -> was last
end;

function TMultiHeaderDBGrid.DateTimePartIsFirst(
  Ed: TCustomDateTimeEdit): Boolean;
// Mirror of DateTimePartIsLast for the leftmost format part.
var
  Before, After: TDTFormatPart;
  HadBefore: Boolean;
begin
  Result:=False;
  if Ed=nil then Exit;
  var Cracker:=TDateTimeEditAccess(Ed);
  HadBefore:=Cracker.FindCurrentFormatPart(Before);
  if not HadBefore then Exit;
  Cracker.GoToPreviousFormatPart;
  if Cracker.FindCurrentFormatPart(After) then begin
    Result:=CompareMem(@Before, @After, SizeOf(TDTFormatPart));
    if not Result then
      Cracker.GoToNextFormatPart;     // undo the probe step
  end
  else
    Result:=True;
end;

procedure TMultiHeaderDBGrid.DateTimeEditorKeyDown(Sender: TObject;
  var Key: Word; var KeyChar: Char; Shift: TShiftState);
// Section-crossing between the date and time halves of an ftDateTime cell.
// These controls navigate by format part (day/month/year, hour/min/sec), not a
// character caret, so "edge" means the focused part is the first/last section.
//   - Right arrow on the date control's last part  -> first part of time control.
//   - Left  arrow on the time control's first part -> last part of date control.
// Tab/Shift+Tab also hop between the two halves before leaving the cell.
// Everything else (Enter/Esc/normal editing) defers to the shared handler.
begin
  if IsCompositeEditor(Sender) then begin
    case Key of
      vkRight:
        if (Sender=FDateEditor) and DateTimePartIsLast(FDateEditor) then begin
          FocusCompositePart(True);   // cross into the time control
          Key:=0; KeyChar:=#0;
          Exit;
        end;

      vkLeft:
        if (Sender=FTimeEditor) and DateTimePartIsFirst(FTimeEditor) then begin
          FocusCompositePart(False);  // cross back into the date control
          Key:=0; KeyChar:=#0;
          Exit;
        end;

      vkTab:
        begin
          // Tab moves date -> time; Shift+Tab moves time -> date. Only when the
          // hop stays inside the cell; otherwise fall through to the shared
          // handler, which moves to the next/previous grid cell.
          if (ssShift in Shift) and (Sender=FTimeEditor) then begin
            FocusCompositePart(False);
            Key:=0; KeyChar:=#0;
            Exit;
          end
          else if not (ssShift in Shift) and (Sender=FDateEditor) then begin
            FocusCompositePart(True);
            Key:=0; KeyChar:=#0;
            Exit;
          end;
        end;
    end;
  end;

  // Enter/Esc/Tab-out and anything not handled above.
  EditorKeyDown(Sender, Key, KeyChar, Shift);
end;

procedure TMultiHeaderDBGrid.EditorExit(Sender: TObject);
// For the ftDateTime composite the date and time controls form one logical
// editor, so focus moving from one to the other must NOT commit/close the cell.
// Only commit when focus has actually left both halves. All other editors keep
// the base behaviour (any focus loss commits).
begin
  if FEditing and IsCompositeEditor(Sender) then begin
    // Defer the decision: after this OnExit, focus may land on the sibling
    // control (intra-cell hop) or somewhere outside (real exit). Check on the
    // next message cycle which of the two halves, if either, now holds focus.
    TThread.ForceQueue(nil,
      procedure
      begin
        if not FEditing then Exit;
        if (FDateEditor<>nil) and FDateEditor.IsFocused then Exit;
        if (FTimeEditor<>nil) and FTimeEditor.IsFocused then Exit;
        CommitEditing; // focus left the whole composite
      end);
    Exit;
  end;

  inherited; // non-composite editors: base commits on focus loss
end;

procedure TMultiHeaderDBGrid.NumberEditorChangeTracking(Sender: TObject);
// Widen the column while typing so a growing number is never clipped (numeric
// cells don't wrap). EditorMinColWidth measures the live editor text for the
// cell being edited, so WidenColForEditor grows the column to fit it.
begin
  if FEditing and (FActiveEditor=FNumberEditor) then
    WidenColForEditor(FEditCol, FEditRow);
end;

function TMultiHeaderDBGrid.EditorMinColWidth(ACol, ARow: Integer): single;
begin
  Result:=0;
  var Field:=FieldForCol(ACol);
  if Field=nil then Exit;

  case EditorKindForField(Field) of
    ekComboBox: Result:=MINW_COMBO+FGridLineWidth/2;
    ekDate:     Result:=MINW_DATE+FGridLineWidth/2+FEditorExtraWidth;
    ekTime:     Result:=MINW_TIME+FGridLineWidth/2+FEditorExtraWidth;
    ekDateTime: Result:=MINW_DATETIME+FGridLineWidth/2+FEditorExtraWidth;
    ekNumber:
      begin
        // Numbers must not wrap: make the column wide enough to show the whole
        // value. While this cell is being edited, measure the live editor text
        // (so the column grows as the user types); otherwise the cell's text.
        const EDITOR_INSET = 8; // TEdit's own internal text margin + caret room
        var Style:=EffectiveCellStyle(ACol, ARow);
        Canvas.Font.Assign(FCellFont);
        Canvas.Font.Family:=Style.FontName;
        Canvas.Font.Size  :=Style.FontSize;
        Canvas.Font.Style :=Style.FontStyle;
        var S: string;
        if FEditing and (FActiveEditor=FNumberEditor) and
           (FEditCol=ACol) and (FEditRow=ARow) then
          S:=FNumberEditor.Text
        else
          S:=Cells[ACol, ARow];
        var W:=Canvas.TextWidth(S);
        if W>0 then
          Result:=Ceil(W)+CellPadding.Left+CellPadding.Right+
                  FGridLineWidth+EDITOR_INSET;
      end;
  else
    Result:=0; // ekMemo / ekCheckBox / ekNone
  end;
end;

procedure TMultiHeaderDBGrid.WidenColForEditor(ACol, ARow: Integer);
// EditorMinColWidth returns the MINW_* base footprint for date/time editors,
// but a nullable field also shows a clear button (see LayoutActiveEditor) that
// EditorMinColWidth doesn't include. Temporarily fold that footprint into the
// measurement so the base widening logic accounts for it, then restore.
var
  Field: TField;
  Extra: single;
begin
  Extra:=0;
  Field:=FieldForCol(ACol);
  if Field<>nil then begin
    var Nullable:=(not Field.Required) and (not Field.ReadOnly);
    if Nullable then
      case EditorKindForField(Field) of
        ekDate, ekTime, ekDateTime: Extra:=ClearButtonWidth(Field);
      end;
  end;

  if Extra<=0 then begin
    inherited WidenColForEditor(ACol, ARow);
    Exit;
  end;

  FEditorExtraWidth:=Extra;
  try
    inherited WidenColForEditor(ACol, ARow);
  finally
    FEditorExtraWidth:=0;
  end;
end;

function TMultiHeaderDBGrid.CanEditCell(ACol, ARow: Integer): Boolean;
begin
  // An editor may open even when the grid/field is read-only, so its text can
  // be selected and copied (the editor itself is set read-only by the base via
  // CellIsModifiable). Boolean fields are excluded - they toggle in place.
  Result:=False;
  if (ACol<0) or (ARow<0) or (ARow>=RowCount) then Exit;

  var Field:=FieldForCol(ACol);
  if Field=nil then Exit;

  var Kind:=EditorKindForField(Field, True); // ignore ReadOnly: open for copy
  Result:=(Kind<>ekNone) and (Kind<>ekCheckBox);
end;

function TMultiHeaderDBGrid.CellIsModifiable(ACol, ARow: Integer): Boolean;
begin
  // A DB cell accepts edits only when the grid isn't ReadOnly, the dataset can
  // modify, and the bound field is itself writable.
  Result:=False;
  if FReadOnly then Exit;
  if (ACol<0) or (ARow<0) or (ARow>=RowCount) then Exit;

  var DS:=DataSet;
  if (DS=nil) or (not DS.Active) or (not DS.CanModify) then Exit;

  var Field:=FieldForCol(ACol);
  if (Field=nil) or Field.ReadOnly then Exit;

  var Kind:=EditorKindForField(Field); // honours field ReadOnly
  Result:=(Kind<>ekNone) and (Kind<>ekCheckBox);
end;

procedure TMultiHeaderDBGrid.ApplyEditorReadOnly(Ed: TControl; AModifiable: Boolean);
begin
  // TDateEdit / TTimeEdit are not TCustomEdit descendants, so handle them here;
  // everything else (memo, number edit, combo) goes through the base.
  if Ed is TCustomDateEdit then
    TCustomDateEdit(Ed).ReadOnly:=not AModifiable
  else if Ed is TCustomTimeEdit then
    TCustomTimeEdit(Ed).ReadOnly:=not AModifiable
  else
    inherited;

  // The ftDateTime composite opens the date control as the active editor but
  // also shows the time control beside it; keep both in the same state.
  if (Ed=FDateEditor) and (FTimeEditor<>nil) and
     (FEditField<>nil) and (EditorKindForField(FEditField, True)=ekDateTime) then
    FTimeEditor.ReadOnly:=not AModifiable;
end;


function TMultiHeaderDBGrid.HandleDownKey: Boolean;
// Mirrors VCL TDBGrid.NextRow for a plain Down press:
//   - On an unmodified pending insert at Eof: do nothing (stay).
//   - On the last record (Eof reached after the move): append a new record.
//   - Otherwise: fall through to a normal one-row move.
begin
  Result:=False;
  var DS:=DataSet;
  if (DS=nil) or (not DS.Active) then Exit;

  var AtLast:=(RowCount>0) and (Row>=RowCount-1);

  // An unmodified, freshly-inserted record sitting at the end: Down keeps the
  // caret there instead of appending another empty record.
  if (DS.State=dsInsert) and (not DS.Modified) and AtLast then
    Exit(True);

  // Moving down from the last record appends a new (open) record - but only
  // when modifiable and the true end is known (no pending on-demand fetch).
  if AtLast and DS.CanModify and (not FReadOnly) and (not EndIsOnDemand) then begin
    AppendRow;
    Exit(True);
  end;

  // Not at the end (or not modifiable): let the base move down one row.
  Result:=False;
end;

procedure TMultiHeaderDBGrid.InsertRow(ARow: Integer);
// Puts the dataset into insert mode at grid row ARow and leaves the new record
// open for the user to fill in - exactly like VCL TDBGrid. It does NOT post:
// required fields (keys, generators, etc.) are the application's responsibility,
// via the dataset's OnNewRecord or the grid's OnInsertRow event.
begin
  var DS:=DataSet;
  if (DS=nil) or (not DS.Active) or (not DS.CanModify) then Exit;
  if not DoInsertRow(ARow) then Exit;

  if FEditing then CancelEditing;
  try DS.RecNo:=Max(0, ARow)+1; except end;
  FInsertRowIndex:=Max(0, ARow); // phantom row shown before the current row
  DS.Insert;
end;

procedure TMultiHeaderDBGrid.AppendRow;
// Puts the dataset into insert mode at the end (Append) and leaves the new
// record open for editing. Like InsertRow, it does NOT post - the application
// fills required fields. Mirrors VCL TDBGrid's append behaviour.
begin
  var DS:=DataSet;
  if (DS=nil) or (not DS.Active) or (not DS.CanModify) then Exit;
  if not DoAppendRow(RowCount) then Exit;

  if FEditing then CancelEditing;
  FInsertRowIndex:=DS.RecordCount; // phantom row shown after the last row
  DS.Append;
end;

procedure TMultiHeaderDBGrid.DeleteRow(ARow: Integer);
// Deletes the record at grid row ARow. When ConfirmDelete is set it first asks
// for confirmation (OK/Cancel), matching VCL TDBGrid's dgConfirmDelete prompt.
begin
  var DS:=DataSet;
  if (DS=nil) or (not DS.Active) or (not DS.CanModify) then Exit;
  if (ARow<0) or (ARow>=RowCount) then Exit;
  if DS.IsEmpty then Exit;
  if not DoDeleteRow(ARow) then Exit;

  if FConfirmDelete and
     (TDialogServiceSync.MessageDialog('Delete record?',
        TMsgDlgType.mtConfirmation, [TMsgDlgBtn.mbOK, TMsgDlgBtn.mbCancel],
        TMsgDlgBtn.mbOK, 0)=mrCancel) then Exit;

  if FEditing then CancelEditing;
  HideEditor;

  try
    DS.RecNo:=ARow+1;
    DS.Delete;
  except
  end;
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
  var Kind:=EditorKindForField(Field, True); // open even for read-only (copy)

  case Kind of
    ekMemo:
      // Reuse the shared TMemo via the base implementation.
      Result:=inherited PrepareCellEditor(ACol, ARow);

    ekNumber:
      begin
        EnsureTypedEditors;
        // Follow the column's horizontal text alignment, like the TMemo does.
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
        // control sits beside it, both laid out by LayoutActiveEditor.
        FActiveEditor:=FDateEditor;
        Result:=FDateEditor;
      end;
  else
    Result:=nil; // ekNone: not editable
  end;
end;

procedure TMultiHeaderDBGrid.ResetTypedEditorCursors;
// Re-init caret / selection / format-part of every typed editor to a known
// state so a fresh edit on a new cell never inherits the previous cell's cursor
// position or selection. The shared memo and the number/edit controls have
// their selection cleared and SelStart aligned to the caret (end of text, just
// loaded); the combo box has no caret; the date/time controls reset to their
// first format part. Both composite halves are reset, not just the active one.
begin
  // Shared TMemo (base field, reused for ekMemo cells).
  if FEditor<>nil then begin
    FEditor.GoToTextEnd;
    FEditor.SelLength:=0;
    FEditor.SelStart:=FEditor.Text.Length;
  end;

  // Plain TEdit for numeric cells.
  if FNumberEditor<>nil then begin
    FNumberEditor.GoToTextEnd;
    FNumberEditor.SelLength:=0;
    FNumberEditor.SelStart:=FNumberEditor.Text.Length;   // SelStart = caret
    FNumberEditor.CaretPosition:=FNumberEditor.Text.Length;
  end;

  // Combo box: no text caret/selection to reset (its value is ItemIndex), so
  // it is intentionally left untouched here.

  // The date control starts on its first section (day); the time control on its
  // first too, so a later hand-off into it begins cleanly at the hour.
  DateTimeGoToFirstPart(FDateEditor);
  DateTimeGoToFirstPart(FTimeEditor);
end;

procedure TMultiHeaderDBGrid.PlaceEditorCaretAtEnd(Ed: TControl);
begin
  // Always re-init every typed editor first, so a new edit on a different cell
  // starts from a known cursor state regardless of which control is active.
  ResetTypedEditorCursors;

  // Picker date/time controls don't expose a usable character caret; their
  // position is set by ResetTypedEditorCursors (format-part) above, so nothing
  // more to do here for them.
  if (Ed=FDateEditor) or (Ed=FTimeEditor) then Exit;

  // The number editor (a plain TEdit) and any other TCustomEdit: caret to end,
  // selection cleared, SelStart aligned to the caret. The combo box has no
  // caret; fall back to the base for the shared TMemo (which also clears
  // selection and aligns SelStart).
  if Ed is TCustomEdit then begin
    var E:=TCustomEdit(Ed);
    E.GoToTextEnd;
    E.SelLength:=0;
    E.CaretPosition:=E.Text.Length;
    E.SelStart:=E.Text.Length; // keep SelStart consistent with the caret
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
        // Plain TEdit holding a numeric string. Typing a digit/sign starts
        // fresh from that char; F2 / double-click (InitialChar=#0) loads the
        // field's current value. A NULL field loads as '' so it stays NULL
        // unless the user types something (see GetEditorText / commit).
        if (InitialChar>=' ') and CharInSet(InitialChar, ['0'..'9','-','+','.',',']) then
          FNumberEditor.Text:=InitialChar
        else if Field.IsNull then
          FNumberEditor.Text:=''
        else if FieldIsInteger(Field) then
          FNumberEditor.Text:=Field.AsLargeInt.ToString
        else
          FNumberEditor.Text:=Field.AsString; // keeps the field's native formatting
      end;

    ekComboBox:
      FComboEditor.ItemIndex:=FComboEditor.Items.IndexOf(Field.AsString);

    ekDate: begin
      if Field.IsNull then FDateEditor.Date:=Now
      else FDateEditor.Date:=Field.AsDateTime;
      FDateEditor.IsEmpty:=False;
    end;
    ekTime: begin
      if Field.IsNull then FTimeEditor.Time:=Now
      else FTimeEditor.Time:=Field.AsDateTime;
      FTimeEditor.IsEmpty:=False;
    end;
    ekDateTime: begin
        var V: TDateTime;
        if Field.IsNull then V:=Now else V:=Field.AsDateTime;
        FDateEditor.Date:=V;
        FTimeEditor.Time:=V;
        FDateEditor.IsEmpty:=False;
        FTimeEditor.IsEmpty:=False;
      end;
  end;
end;

function TMultiHeaderDBGrid.GetEditorText: string;
begin
  var Field:=FEditField;
  if Field=nil then Exit(inherited GetEditorText);

  case EditorKindForField(Field) of
    ekMemo:      Result:=inherited GetEditorText;
    ekNumber:    Result:=Trim(FNumberEditor.Text); // '' means SQL NULL
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
          begin
            // An empty box is SQL NULL; otherwise parse the text and assign
            // through the matching typed setter so the field keeps its native
            // precision (integer vs float). The parsed value is clamped to the
            // field type's range so an out-of-range entry can't overflow/wrap
            // into a garbage value. A bad parse raises and the outer try/except
            // cancels the edit.
            var NumText:=Trim(FNumberEditor.Text);
            if NumText='' then
              Field.Clear
            else begin
              var MinV, MaxV: Double;
              NumberRangeForField(Field, MinV, MaxV);
              // The key filter accepts both '.' and ',' as the decimal point;
              // normalize to the locale separator so StrToFloat parses either.
              NumText:=NumText.Replace('.', FormatSettings.DecimalSeparator)
                              .Replace(',', FormatSettings.DecimalSeparator);
              // Parse as float (not StrToInt64) so a value beyond Int64 range
              // doesn't raise before it can be clamped - clamp first, then
              // assign through the matching typed setter.
              var F:=StrToFloat(NumText);
              if F<MinV then F:=MinV
              else if F>MaxV then F:=MaxV;
              if FieldIsInteger(Field) then
                Field.AsLargeInt:=Round(F)
              else
                Field.AsFloat:=F;
            end;
          end;
        ekComboBox:
          if FComboEditor.ItemIndex>=0 then
            Field.AsString:=FComboEditor.Selected.Text;
        ekDate:
          if FDateEditor.IsEmpty then begin
            Field.Clear;
          end else begin
            Field.AsDateTime:=Trunc(FDateEditor.Date);
          end;
        ekTime:
          if FTimeEditor.IsEmpty then begin
            Field.Clear;
          end else begin
            Field.AsDateTime:=Frac(FTimeEditor.Time);
          end;
        ekDateTime:
          if FTimeEditor.IsEmpty then begin
            Field.Clear;
          end else begin
            Field.AsDateTime:=Trunc(FDateEditor.Date)+Frac(FTimeEditor.Time);
          end;
      else
        // ekMemo and any text path: round-trip via AsString.
        Field.AsString:=GetEditorText;
      end;
      // During an insert, keep the record open (don't post per field) so the
      // user can fill required fields; the row is committed when it's left.
      // Editing an existing record posts immediately as before.
      if DS.State<>dsInsert then
        DS.Post;
    except
      if DS.State<>dsInsert then DS.Cancel;
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
    if (ARow=CurrentRowIndex) or (ARow=DS.RecNo-1) then
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
    // Conversion/validation failed: keep the old value and stay put.
    Result:=False;
  end;

  // Drop the cached row text so DoGetCellText re-reads the committed value.
  FCellTexts.Remove(ARow);
  FEditField:=nil;

  // The committed value may need a different number of wrapped lines, so re-fit
  // the row height (the base CommitEditorValue does the same for Cells[]).
  // Merged cells span a range. Done before Invalidate so the new height paints
  // immediately.
  if Result and (ARow>=0) and (ARow<FRowCount) then begin
    var MergedCell: TMergedCell;
    if IsMergedCell(ACol,ARow,MergedCell) then begin
      AutoSizeRows(MergedCell.Row, MergedCell.Row+MergedCell.RowSpan-1)
    end else begin
      AutoSizeRows(ARow, ARow);
    end;
    UpdateSize;
  end;

  Invalidate;
end;

procedure TMultiHeaderDBGrid.CancelEditing;
begin
  inherited;
  FEditField:=nil;
  // Escape on a pending insert discards the new record (VCL behaviour).
  var DS:=DataSet;
  if (DS<>nil) and DS.Active and (DS.State=dsInsert) then begin
    FInsertRowIndex:=-1;
    DS.Cancel;
  end;
end;

function TMultiHeaderDBGrid.ClearButtonWidth(Field: TField): single;
// Footprint of the styled clear button (its Width plus left/right margins),
// read from the captured 'clearbutton' style instance so column widening tracks
// whatever the style defines. The date control's button serves ekDate and the
// ekDateTime composite; the time control's button serves ekTime. Returns 0 when
// the style hasn't resolved yet (first frame) or no button applies - the column
// re-measures on the next layout once the instance is captured.
var
  Control: TControl;
begin
  Result:=0;
  if Field=nil then Exit;
  case EditorKindForField(Field) of
    ekDate, ekDateTime: Control:=FDateEditor;
    ekTime:             Control:=FTimeEditor;
  else
    Control:=nil;
  end;
  if Control<>nil then begin
    var Button:=TButton(Control.FindStyleResource('mhgclearbutton'));
    if Assigned(Button) then begin
      Result:=Button.Width+Button.Margins.Left+Button.Margins.Right;
    end;
  end;
end;

function GetClearButtonWidth(Control: TControl): single;
begin
  var Button:=TButton(Control.FindStyleResource('mhgclearbutton'));
  if Assigned(Button) then begin
    Result:=Button.Width+Button.Margins.Left+Button.Margins.Right;
  end else begin
    Result:=0;
  end;
end;

procedure SetEditorClearButton(Control: TControl; Visible: boolean);
begin
  var Button:=TButton(Control.FindStyleResource('mhgclearbutton'));
  if Assigned(Button) then begin
    Button.Visible:=Visible;
  end;
end;

procedure TMultiHeaderDBGrid.LayoutActiveEditor(AWidth, AHeight: Single);
begin
  if (FEditorHost=nil) or not FEditorHost.Visible then Exit;

  var Ed:=ActiveEditorControl;
  if Ed=nil then Exit;

  // ekMemo cells reuse the shared TMemo (FActiveEditor=FEditor). The memo
  // placement is intentionally a self-contained copy of the base's logic rather
  // than a call to `inherited`: the base and DB grids size the memo differently
  // enough (style/ContentBounds timing) that sharing one path made fixing one
  // grid regress the other. Keeping them separate lets each be tuned alone.
  if Ed=FEditor then begin
    HideTypedEditors(nil, False);
    LayoutMemoEditor(AWidth, AHeight);
    Exit;
  end;

  // Show only the active typed editor (and, for ftDateTime, the time control);
  // hide the memo and the others so nothing lingers over FEditorBack.
  var Composite:=(FEditField<>nil) and (EditorKindForField(FEditField)=ekDateTime);
  HideTypedEditors(Ed, Composite);
  if FEditor<>nil then FEditor.Visible:=False;

  var EdH:=AHeight;
  if (FTypedEditorHeight>0) and (FTypedEditorHeight<EdH) then begin
    EdH:=FTypedEditorHeight;
  end;

  // The plain-edit number editor mirrors the cell's font the way the base does
  // for the TMemo (face/size/style/colour from the effective cell style), so
  // editing looks identical to the painted cell. The date/time/combo controls
  // keep their own styling.
  if Ed=FNumberEditor then begin
    var NStyle:=EffectiveCellStyle(FEditCol, FEditRow);
    FNumberEditor.TextSettings.Font.Family:=NStyle.FontName;
    FNumberEditor.TextSettings.Font.Size  :=NStyle.FontSize;
    FNumberEditor.TextSettings.Font.Style :=NStyle.FontStyle;
    FNumberEditor.TextSettings.HorzAlign  :=NStyle.TextHAlignment;
  end;

  // Single-line typed editors are one line tall, so the slack is simply the
  // gap between the cell and that line. Offset by TextVAlignment the same way
  // the base positions the memo (Leading -> top, Center -> middle, Trailing ->
  // bottom). Whole-pixel offset avoids sub-pixel rounding jitter. The -2 / +2
  // keep the existing single-line baseline nudge and overscan.
  var Slack:=AHeight-EdH;
  if Slack<0 then Slack:=0;
  var OffY: Single;
  case EffectiveCellStyle(FEditCol,FEditRow).TextVAlignment of
    TTextAlign.Center:   OffY:=Int(Slack/2);
    TTextAlign.Trailing: OffY:=Int(Slack);
  else
    OffY:=0; // Leading
  end;

  var ClearButtonWidth:=0.0;
  case EditorKindForField(FEditField) of
    ekDate: begin
      ClearButtonWidth:=GetClearButtonWidth(FDateEditor);
    end;
    ekTime, ekDateTime: begin
      ClearButtonWidth:=GetClearButtonWidth(FTimeEditor);
    end;
  end;

  var FieldIsNullable:=(FEditField<>nil) and (not FEditField.Required) and (not FEditField.ReadOnly);

  if not FieldIsNullable then ClearButtonWidth:=0;

  if Composite and (FDateEditor<>nil) and (FTimeEditor<>nil) then begin
    // Split the host width, biasing the date editor a little wider than time.

    SetEditorClearButton(FDateEditor,False);
    SetEditorClearButton(FTimeEditor,FieldIsNullable);

    var Center:=AWidth/2+(MINW_DATE-(MINW_TIME+ClearButtonWidth))/2;

    FDateEditor.SetBounds(Center-MINW_DATE, OffY, MINW_DATE, EdH);
    FTimeEditor.SetBounds(Center, OffY, MINW_TIME+ClearButtonWidth, EdH);

    FDateEditor.Visible:=True;
    FTimeEditor.Visible:=True;
    FDateEditor.BringToFront;
    FTimeEditor.BringToFront;
  end else begin
    var X:=0.0;
    var DY:=0.0;
    var W:=AWidth;

    if Ed=FNumberEditor then begin
      DY:=-0.5;
    end;
    if Ed=FDateEditor then begin
      SetEditorClearButton(FDateEditor,FieldIsNullable);
      W:=MINW_DATE+IfThen(FieldIsNullable,ClearButtonWidth,0);

      var Center:=(AWidth+W)/2;
      X:=Center-W;
    end;
    if Ed=FTimeEditor then begin
      SetEditorClearButton(FTimeEditor,FieldIsNullable);

      W:=MINW_TIME+IfThen(FieldIsNullable,ClearButtonWidth,0);

      var Center:=(AWidth+W)/2;
      X:=Center-W;
    end;

    Ed.SetBounds(X, OffY+DY, W, EdH+DY); // fill width, natural height
    Ed.Visible:=True;
    Ed.BringToFront;
  end;
end;

procedure TMultiHeaderDBGrid.LayoutMemoEditor(AWidth, AHeight: Single);
// DB grid's own copy of the shared-TMemo placement for ekMemo cells. Kept
// separate from TMultiHeaderGrid.LayoutActiveEditor (no `inherited`) so the two
// grids can be tuned independently. This currently mirrors the base logic that
// was working well for the DB grid; adjust here without affecting the basic
// grids.
begin
  if FEditor=nil then Exit;

  var Style:=EffectiveCellStyle(FEditCol, FEditRow);

  FEditor.TextSettings.HorzAlign  :=Style.TextHAlignment;
  FEditor.TextSettings.Font.Family:=Style.FontName;
  FEditor.TextSettings.Font.Size  :=Style.FontSize;
  FEditor.TextSettings.Font.Style :=Style.FontStyle;
  FEditor.WordWrap:=Style.WordWrap;

  var RowSpan:=1;
  var MC: TMergedCell;
  if IsMergedCell(FEditCol,FEditRow,MC) then RowSpan:=MC.RowSpan;
  var TextH:=MeasureCellTextHeight(FEditCol,FEditRow,FEditor.Text,RowSpan);
  var Slack:=AHeight - TextH;
  if Slack<0 then Slack:=0;

  var OffY: Single;
  case Style.TextVAlignment of
    TTextAlign.Center:   OffY:=Slack/2;
    TTextAlign.Trailing: OffY:=Slack;
  else
    OffY:=0; // Leading
  end;

  FEditor.SetBounds(0, OffY-3, AWidth+3, AHeight-OffY+8);

  FEditor.Visible:=True;
  FEditor.BringToFront;
end;

procedure TMultiHeaderDBGrid.HideTypedEditors(Keep: TControl; KeepComposite: Boolean);
// Hides every typed editor except Keep. KeepComposite preserves the ftDateTime
// date+time pair, whose active control is only the date editor.
begin
  if (FComboEditor<>nil)  and (FComboEditor<>Keep)  then FComboEditor.Visible:=False;
  if (FNumberEditor<>nil) and (FNumberEditor<>Keep) then FNumberEditor.Visible:=False;
  if (FDateEditor<>nil)   and (FDateEditor<>Keep)   and not KeepComposite then FDateEditor.Visible:=False;
  if (FTimeEditor<>nil)   and (FTimeEditor<>Keep)   and not KeepComposite then FTimeEditor.Visible:=False;
end;

procedure TMultiHeaderDBGrid.HideEditor;
begin
  inherited; // hides the host (with FEditorBack) and the active editor
  // The base hides only ActiveEditorControl; clear the composite's time control
  // and any inactive typed editor still parented in the host.
  HideTypedEditors(nil, False);
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

function CreateMemoStyle: TLayout;
begin
  Result := TLayout.Create(nil);
  Result.StyleName := 'mhg_editor_memo';

  var Content := TLayout.Create(Result);
  Content.Parent := Result;
  Content.StyleName := 'content';
  Content.Align := TAlignLayout.Client;
  Content.Margins.Left:=2;
  Content.Margins.Right:=2;
  Content.Margins.Top:=3;
  Content.Margins.Bottom:=2;
end;

function CreateNumberStyle: TLayout;
begin
  Result := TLayout.Create(nil);
  Result.StyleName := 'mhg_editor_number';

  var Content := TLayout.Create(Result);
  Content.Parent := Result;
  Content.StyleName := 'content';
  Content.Align := TAlignLayout.Client;
  Content.Margins.Top:=1;
end;

function CreateComboStyle: TLayout;
begin
  Result := TLayout.Create(nil);
  Result.StyleName := 'mhg_editor_combo';

  var Content := TLayout.Create(Result);
  Content.StyleName := 'content';
  Content.Align := TAlignLayout.Client;
  Content.Margins.Right := 16; // leave room for the arrow

  var Arrow := TPath.Create(nil);
  Arrow.Parent := Result;
  Arrow.StyleName := 'arrow';
  Arrow.Align := TAlignLayout.Right;
  Arrow.Width := 16;
  Arrow.Data.Data := 'M0,0 L8,0 L4,5 z';
  Arrow.Fill.Color := TAlphaColors.Gray;
  Arrow.WrapMode := TPathWrapMode.Fit;
end;

function CreateDateStyle: TLayout;
begin
  Result := TLayout.Create(nil);
  Result.StyleName := 'mhg_editor_date';

  var Text := TActiveStyleTextObject.Create(Result);
  Text.Parent := Result;
  Text.StyleName := 'Text';
  Text.Align := TAlignLayout.Client;
  Text.TextSettings.HorzAlign := TTextAlign.Trailing;
  Text.Locked := True;
  Text.Margins.Left := 2;
  Text.Margins.Right := 0;
  Text.Margins.Top := 0;
  Text.Margins.Bottom := 0;
  Text.Size.PlatformDefault := False;
  Text.Text := 'Text';
  Text.ShadowVisible := False;
  Text.ActiveTrigger := TStyleTrigger.Focused;
  Text.ActiveColor := TAlphaColorRec.Black;
  Text.Cursor:=crIBeam;

  var Selection:=TBrushObject.Create(Result);
  Selection.StyleName := 'selection';
  Selection.Brush.Color := $7F2A96FF;

  var CalendarButton := TButton.Create(Result);
  CalendarButton.Parent := Result;
  CalendarButton.StyleName := 'arrow';
  CalendarButton.Align := TAlignLayout.Right;
  CalendarButton.CanFocus := False;
  CalendarButton.Margins.Top := 1;
  CalendarButton.Margins.Bottom := 1;
  CalendarButton.Width := 12;
  CalendarButton.Cursor := crHandPoint;

  var Glyph := TPath.Create(CalendarButton);
  Glyph.Parent := CalendarButton;
  Glyph.StyleName := 'arrowglyph';
  Glyph.Align := TAlignLayout.Client;
  Glyph.HitTest := False;
  Glyph.Margins.Left := 3;
  Glyph.Margins.Right := 3;
  Glyph.Margins.Top := 4;
  Glyph.Margins.Bottom := 4;
  Glyph.Data.Data := 'M0,0 L8,0 L4,5 z';
  Glyph.Fill.Color := TAlphaColors.Gray;
  Glyph.WrapMode := TPathWrapMode.Fit;

  var ClearBtn := TButton.Create(Result);
  ClearBtn.Parent := Result;
  ClearBtn.StyleName := 'mhgclearbutton';
  ClearBtn.Align := TAlignLayout.MostRight;
  ClearBtn.CanFocus := False;
  ClearBtn.Visible := True;
  ClearBtn.Margins.Top := 1;
  ClearBtn.Margins.Bottom := 1;
  ClearBtn.Margins.Right := 1;
  ClearBtn.Width := 14;
  ClearBtn.Cursor := crHandPoint;
  ClearBtn.Hint := 'Clear';

  var ClearGlyph := TPath.Create(ClearBtn);
  ClearGlyph.Parent := ClearBtn;
  ClearGlyph.StyleName := 'clearglyph';
  ClearGlyph.Align := TAlignLayout.Center;
  ClearGlyph.Width := 9;
  ClearGlyph.Height := 9;
  ClearGlyph.HitTest := False;
  ClearGlyph.WrapMode := TPathWrapMode.Stretch;
  ClearGlyph.Fill.Kind := TBrushKind.None;
  ClearGlyph.Stroke.Kind := TBrushKind.Solid;
  ClearGlyph.Stroke.Color := TAlphaColorRec.Gray;
  ClearGlyph.Stroke.Thickness := 1.5;
  ClearGlyph.Data.Data := 'M0,0 L9,9 M9,0 L0,9';
end;

function CreateTimeStyle: TLayout;
begin
  Result := TLayout.Create(nil);
  Result.StyleName := 'mhg_editor_time';

  var Text := TActiveStyleTextObject.Create(Result);
  Text.Parent := Result;
  Text.StyleName := 'Text';
  Text.Align := TAlignLayout.Client;
  Text.TextSettings.HorzAlign := TTextAlign.Trailing;
  Text.Locked := True;
  Text.Margins.Left := 2;
  Text.Margins.Top := 0;
  Text.Margins.Right := 0;
  Text.Margins.Bottom := 0;
  Text.Size.PlatformDefault := False;
  Text.Text := 'Text';
  Text.ShadowVisible := False;
  Text.ActiveTrigger := TStyleTrigger.Focused;
  Text.ActiveColor := TAlphaColorRec.Black;
  Text.Cursor:=crIBeam;

  var Selection:=TBrushObject.Create(Result);
  Selection.StyleName := 'selection';
  Selection.Brush.Color := $7F2A96FF;

  var UpDown:=TGridLayout.Create(Result);
  UpDown.Parent := Result;
  UpDown.Align := TAlignLayout.Right;
  UpDown.ItemHeight := -1;
  UpDown.Orientation := TOrientation.Vertical;
  UpDown.Margins.Top := 1;
  UpDown.Margins.Bottom := 1;
  UpDown.Margins.Right := 1;
  UpDown.Width := 14;

  var TopLayout := TLayout.Create(UpDown);
  TopLayout.Parent := UpDown;

  var UpButton := TButton.Create(TopLayout);
  UpButton.Parent := TopLayout;
  UpButton.StyleName := 'upbutton';
  UpButton.Align := TAlignLayout.Client;
  UpButton.CanFocus := False;

  var UpArrow := TPath.Create(UpButton);
  UpArrow.Parent := UpButton;
  UpArrow.Align := TAlignLayout.Center;
  UpArrow.Width := 7;
  UpArrow.Height := 4;
  UpArrow.HitTest := False;
  UpArrow.Data.Data := 'M0,4 L3.5,0 L7,4 Z';
  UpArrow.Fill.Color := TAlphaColorRec.Black;
  UpArrow.Stroke.Kind := TBrushKind.None;
  UpArrow.Cursor := crHandPoint;

  var BottomLayout := TLayout.Create(UpDown);
  BottomLayout.Parent := UpDown;
  BottomLayout.Position.Y := 8;

  var DownButton := TButton.Create(BottomLayout);
  DownButton.Parent := BottomLayout;
  DownButton.StyleName := 'downbutton';
  DownButton.Align := TAlignLayout.Client;
  DownButton.CanFocus := False;
  DownButton.StyleLookup := 'spindownbutton';
  DownButton.Cursor := crHandPoint;

  var DownArrow := TPath.Create(DownButton);
  DownArrow.Parent := DownButton;
  DownArrow.Align := TAlignLayout.Center;
  DownArrow.Width := 7;
  DownArrow.Height := 4;
  DownArrow.HitTest := False;
  DownArrow.Data.Data := 'M0,0 L7,0 L3.5,4 Z';
  DownArrow.Fill.Color := TAlphaColorRec.Black;
  DownArrow.Stroke.Kind := TBrushKind.None;

  var ClearBtn := TButton.Create(Result);
  ClearBtn.Parent := Result;
  ClearBtn.StyleName := 'mhgclearbutton';
  ClearBtn.Align := TAlignLayout.MostRight;
  ClearBtn.CanFocus := False;
  ClearBtn.Visible := True;
  ClearBtn.Margins.Top := 1;
  ClearBtn.Margins.Bottom := 1;
  ClearBtn.Margins.Right := 1;
  ClearBtn.Width := 14;
  ClearBtn.Cursor := crHandPoint;
  ClearBtn.Hint := 'Clear';

  var ClearGlyph := TPath.Create(ClearBtn);
  ClearGlyph.Parent := ClearBtn;
  ClearGlyph.StyleName := 'clearglyph';
  ClearGlyph.Align := TAlignLayout.Center;
  ClearGlyph.Width := 9;
  ClearGlyph.Height := 9;
  ClearGlyph.HitTest := False;
  ClearGlyph.WrapMode := TPathWrapMode.Stretch;
  ClearGlyph.Fill.Kind := TBrushKind.None;
  ClearGlyph.Stroke.Kind := TBrushKind.Solid;
  ClearGlyph.Stroke.Color := TAlphaColorRec.Gray;
  ClearGlyph.Stroke.Thickness := 1.5;
  ClearGlyph.Data.Data := 'M0,0 L9,9 M9,0 L0,9';
end;

initialization
  MemoStyle:=CreateMemoStyle;
  NumberStyle:=CreateNumberStyle;
  ComboStyle:=CreateComboStyle;
  DateStyle:=CreateDateStyle;
  TimeStyle:=CreateTimeStyle;

finalization
  MemoStyle.Free;
  NumberStyle.Free;
  ComboStyle.Free;
  DateStyle.Free;
  TimeStyle.Free;

end.


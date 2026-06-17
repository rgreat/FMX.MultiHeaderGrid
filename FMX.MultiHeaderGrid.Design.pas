unit FMX.MultiHeaderGrid.Design;

{
  Design-time support for FMX.MultiHeaderGrid.

  Both grid families get a custom, NON-MODAL collection editor
  (TMHGColumnsEditorForm) instead of the stock Delphi ColnEdit dialog, so
  that:

    * the editor is modeless - you can pick a column here and edit its
      properties in the Object Inspector at the same time (exactly like the
      stock Fields/Columns editor);
    * columns can be multi-selected and copied / pasted - between grids, or
      out to / in from plain text - via the clipboard (Delphi component
      text); and
    * the layout has a sensible default size, a compact toolbar and no
      horizontal splitter.

  One editor instance is kept alive per edited collection; reopening focuses
  the existing window instead of leaking forms. The window frees itself on
  close (caFree) and on grid/designer teardown.

  - TMultiHeaderDBGrid: edits the data-bound Columns collection (items are
    TMHGColumn: FieldName, Color + everything from the base column). Verbs:
    "Edit Columns...", "Auto Create Columns".

  - TMultiHeaderGrid / TMultiHeaderStringGrid: not data bound; they edit the
    HeaderColumns collection (TMHGHeaderColumn: captions, grouping,
    alignment, word wrap, widths - everything except FieldName).

  Uses DesignIntf / DesignEditors and a VCL form, so this unit MUST live in a
  design-time-only package (MultiHeaderGridDsgn.dpk). Do NOT add it to the
  run-time package.
}

interface

procedure Register;

implementation

uses
  System.Classes, System.SysUtils, System.Types, System.UITypes,
  System.Generics.Collections,
  Winapi.Windows, Winapi.Messages,
  Vcl.Forms, Vcl.Controls, Vcl.StdCtrls, Vcl.ExtCtrls, Vcl.Buttons,
  Vcl.Graphics, Vcl.Clipbrd, Vcl.Dialogs, Vcl.ImgList, Vcl.Menus,
  Vcl.ActnList, System.Actions,
  DesignIntf, DesignEditors,
  FMX.MultiHeaderGrid;

const
  // Clipboard format for an exact round-trip of selected columns. The payload
  // is Delphi component-text of a throw-away carrier holding a collection of
  // the copied columns.
  MHG_CLIP_FORMAT_NAME = 'FMX.MultiHeaderGrid.Columns';

  // Glyph indices in FGlyphs (drawn at run time in BuildGlyphs).
  GI_ADD    = 0;
  GI_DELETE = 1;
  GI_UP     = 2;
  GI_DOWN   = 3;
  GI_COPY   = 4;
  GI_PASTE  = 5;
  GI_IMPORT = 6;
  GLYPH_SIZE = 16;

var
  MHGClipFormat: Cardinal = 0;

type
  // Minimal streamable wrapper: a TComponent that owns a collection of the
  // requested item class.
  TMHGClipCarrier = class(TComponent)
  private
    FItems: TCollection;
  published
    property Items: TCollection read FItems write FItems;
  end;

{ ---- the editor form (built entirely in code, no .dfm) -------------------- }

type
  TMHGColumnsEditorForm = class(TForm)
  private
    FDesigner   : IDesigner;
    FGrid       : TComponent;
    FCollection : TCollection;
    FItemClass  : TCollectionItemClass;
    FTearingDown: Boolean;
    FList       : TListBox;
    FToolBar    : TPanel;
    FGlyphs     : TImageList;
    FActions    : TActionList;
    FActAdd, FActDelete, FActUp, FActDown,
    FActCopy, FActPaste, FActImport, FActSelectAll: TAction;
    procedure BuildUI;
    procedure BuildGlyphs;
    function  MakeAction(const ACaption: string; AGlyph: Integer;
                         AShortCut, AShortCut2: TShortCut;
                         AOnExecute, AOnUpdate: TNotifyEvent): TAction;
    procedure MakeButton(AAction: TAction; var X: Integer;
                         AY: Integer = 4; AExtraGap: Integer = 0);
    procedure RefreshList(KeepIndex: Integer = -1);
    procedure NotifyModified;
    procedure SyncInspector;

    function  ClipboardHasColumns: Boolean;
    function  SelectionToText: string;
    function  AppendFromText(const AText: string): Integer;

    // Action handlers (single source of truth for both buttons and shortcuts).
    procedure DoAdd(Sender: TObject);
    procedure DoDelete(Sender: TObject);
    procedure DoUp(Sender: TObject);
    procedure DoDown(Sender: TObject);
    procedure DoCopy(Sender: TObject);
    procedure DoPaste(Sender: TObject);
    procedure DoImport(Sender: TObject);
    procedure DoSelectAll(Sender: TObject);
    // Action OnUpdate handlers (enable/disable).
    procedure UpdHasSel(Sender: TObject);
    procedure UpdUp(Sender: TObject);
    procedure UpdDown(Sender: TObject);
    procedure UpdPaste(Sender: TObject);
    procedure UpdImport(Sender: TObject);
    procedure UpdHasItems(Sender: TObject);
    procedure ListClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
  public
    constructor CreateEditor(ADesigner: IDesigner; AGrid: TComponent;
      ACollection: TCollection; AItemClass: TCollectionItemClass;
      const ACaption: string); reintroduce;
    property Collection: TCollection read FCollection;
  end;

{ One open editor per collection, so reopening focuses the existing window. }
var
  OpenEditors: TDictionary<TCollection, TMHGColumnsEditorForm> = nil;

constructor TMHGColumnsEditorForm.CreateEditor(ADesigner: IDesigner;
  AGrid: TComponent; ACollection: TCollection;
  AItemClass: TCollectionItemClass; const ACaption: string);
begin
  inherited CreateNew(Application);
  FDesigner   := ADesigner;
  FGrid       := AGrid;
  FCollection := ACollection;
  FItemClass  := AItemClass;

  Caption     := ACaption;
  // Sensible default size; remembers nothing, but starts usable.
  ClientWidth  := 320;
  ClientHeight := 600;
  Constraints.MinWidth  := 260;
  Constraints.MinHeight := 300;
  Position    := poScreenCenter;
  BorderStyle := bsSizeable;
  OnClose     := FormClose;
  BuildUI;
  RefreshList;
end;

procedure TMHGColumnsEditorForm.BuildUI;
var
  X: Integer;
begin
  BuildGlyphs;

  // Actions are the single source of truth: each drives its toolbar button
  // AND its keyboard shortcut(s), with one OnExecute and one OnUpdate.
  FActions := TActionList.Create(Self);
  FActions.Images := FGlyphs;

  FActAdd    := MakeAction('Add',    GI_ADD,    ShortCut(Ord('N'), [ssCtrl]), 0,
                           DoAdd,    nil);   // Add is always enabled

  FActDelete := MakeAction('Delete', GI_DELETE,
                           ShortCut(VK_DELETE, []),         // plain Del
                           ShortCut(VK_DELETE, [ssCtrl]),   // Ctrl+Del
                           DoDelete, UpdHasSel);
  FActUp     := MakeAction('Up',     GI_UP,     ShortCut(VK_UP, [ssCtrl]), 0,
                           DoUp,     UpdUp);
  FActDown   := MakeAction('Down',   GI_DOWN,   ShortCut(VK_DOWN, [ssCtrl]), 0,
                           DoDown,   UpdDown);
  FActCopy   := MakeAction('Copy',   GI_COPY,
                           ShortCut(Ord('C'), [ssCtrl]),
                           ShortCut(VK_INSERT, [ssCtrl]),
                           DoCopy,   UpdHasSel);
  FActPaste  := MakeAction('Paste',  GI_PASTE,
                           ShortCut(Ord('V'), [ssCtrl]),
                           ShortCut(VK_INSERT, [ssShift]),
                           DoPaste,  UpdPaste);
  // Import: refill columns from the DataSet's fields. DB grid only.
  FActImport := MakeAction('Import', GI_IMPORT, ShortCut(Ord('I'), [ssCtrl]), 0,
                           DoImport, UpdImport);
  // Select All: Ctrl+A. No toolbar button - shortcut only.
  FActSelectAll := MakeAction('Select All', -1, ShortCut(Ord('A'), [ssCtrl]), 0,
                              DoSelectAll, UpdHasItems);

  FToolBar := TPanel.Create(Self);
  FToolBar.Parent     := Self;
  FToolBar.Align      := alTop;
  FToolBar.Height     := 62;          // two rows of buttons
  FToolBar.BevelOuter := bvNone;
  FToolBar.ParentBackground := False;

  // Row 1 (Y=4): item operations.
  X := 6;
  MakeButton(FActAdd,    X, 4);
  MakeButton(FActDelete, X, 4);
  MakeButton(FActUp,     X, 4);
  MakeButton(FActDown,   X, 4);

  // Row 2 (Y=32): clipboard + import, restarting from the left.
  X := 6;
  MakeButton(FActCopy,   X, 32);
  MakeButton(FActPaste,  X, 32);
  // Import only makes sense for the data-bound grid (it has a DataSet).
  if FGrid is TMultiHeaderDBGrid then
    MakeButton(FActImport, X, 32);

  FList := TListBox.Create(Self);
  FList.Parent           := Self;
  FList.Align            := alClient;
  FList.MultiSelect      := True;
  FList.ExtendedSelect   := True;
  FList.OnClick          := ListClick;
  // No splitter, no gap: the list sits directly under the toolbar.
end;

// Draws seven flat 16x16 monochrome glyphs into FGlyphs at run time, so the
// design package needs no external .res / image files.
procedure TMHGColumnsEditorForm.BuildGlyphs;

  function NewGlyph: TBitmap;
  begin
    Result := TBitmap.Create;
    Result.PixelFormat := pf24bit;
    Result.SetSize(GLYPH_SIZE, GLYPH_SIZE);
    // Fuchsia background becomes the transparency mask via AddMasked.
    Result.Canvas.Brush.Color := clFuchsia;
    Result.Canvas.Brush.Style := bsSolid;
    Result.Canvas.FillRect(Rect(0, 0, GLYPH_SIZE, GLYPH_SIZE));
    Result.Canvas.Pen.Color := clBlack;
    Result.Canvas.Pen.Width := 1;
    Result.Canvas.Brush.Style := bsClear;
  end;

  procedure AddGlyph(B: TBitmap);
  begin
    try
      FGlyphs.AddMasked(B, clFuchsia);
    finally
      B.Free;
    end;
  end;

var
  B: TBitmap;
begin
  FGlyphs := TImageList.Create(Self);
  FGlyphs.Width  := GLYPH_SIZE;
  FGlyphs.Height := GLYPH_SIZE;

  // 0: Add - green plus
  B := NewGlyph;
  B.Canvas.Pen.Color := clGreen; B.Canvas.Pen.Width := 2;
  B.Canvas.MoveTo(8, 3);  B.Canvas.LineTo(8, 13);
  B.Canvas.MoveTo(3, 8);  B.Canvas.LineTo(13, 8);
  AddGlyph(B);

  // 1: Delete - red X
  B := NewGlyph;
  B.Canvas.Pen.Color := clRed; B.Canvas.Pen.Width := 2;
  B.Canvas.MoveTo(4, 4);  B.Canvas.LineTo(12, 12);
  B.Canvas.MoveTo(12, 4); B.Canvas.LineTo(4, 12);
  AddGlyph(B);

  // 2: Up - filled triangle
  B := NewGlyph;
  B.Canvas.Brush.Style := bsSolid; B.Canvas.Brush.Color := clNavy;
  B.Canvas.Pen.Color := clNavy;
  B.Canvas.Polygon([Point(8, 3), Point(13, 11), Point(3, 11)]);
  AddGlyph(B);

  // 3: Down - filled triangle
  B := NewGlyph;
  B.Canvas.Brush.Style := bsSolid; B.Canvas.Brush.Color := clNavy;
  B.Canvas.Pen.Color := clNavy;
  B.Canvas.Polygon([Point(3, 5), Point(13, 5), Point(8, 13)]);
  AddGlyph(B);

  // 4: Copy - two overlapping pages
  B := NewGlyph;
  B.Canvas.Pen.Color := clBlack; B.Canvas.Pen.Width := 1;
  B.Canvas.Brush.Style := bsSolid; B.Canvas.Brush.Color := clWhite;
  B.Canvas.Rectangle(3, 3, 10, 11);    // back page
  B.Canvas.Rectangle(6, 6, 13, 14);    // front page
  AddGlyph(B);

  // 5: Paste - clipboard
  B := NewGlyph;
  B.Canvas.Pen.Color := clBlack; B.Canvas.Pen.Width := 1;
  B.Canvas.Brush.Style := bsSolid; B.Canvas.Brush.Color := $00C8C8C8;
  B.Canvas.Rectangle(3, 4, 13, 14);    // board
  B.Canvas.Brush.Color := clWhite;
  B.Canvas.Rectangle(5, 6, 11, 13);    // sheet
  B.Canvas.Brush.Color := $00808080;
  B.Canvas.Rectangle(6, 2, 10, 5);     // clip
  AddGlyph(B);

  // 6: Import - database cylinder with a green down arrow (refill from data)
  B := NewGlyph;
  B.Canvas.Pen.Color := clBlack; B.Canvas.Pen.Width := 1;
  B.Canvas.Brush.Style := bsSolid; B.Canvas.Brush.Color := $00E0E0E0;
  B.Canvas.Ellipse(2, 2, 10, 5);       // cylinder top
  B.Canvas.Rectangle(2, 3, 10, 10);    // cylinder body
  B.Canvas.Brush.Color := $00E0E0E0;
  B.Canvas.Ellipse(2, 8, 10, 11);      // cylinder bottom rim
  B.Canvas.Pen.Color := clGreen; B.Canvas.Pen.Width := 2;
  B.Canvas.MoveTo(12, 5);  B.Canvas.LineTo(12, 12);    // arrow shaft
  B.Canvas.MoveTo(9, 9);   B.Canvas.LineTo(12, 13);    // arrow head left
  B.Canvas.MoveTo(15, 9);  B.Canvas.LineTo(12, 13);    // arrow head right
  AddGlyph(B);
end;

// Creates an action carrying its caption, glyph, up to two shortcuts and its
// execute/update handlers. Owned by FActions.
function TMHGColumnsEditorForm.MakeAction(const ACaption: string;
  AGlyph: Integer; AShortCut, AShortCut2: TShortCut;
  AOnExecute, AOnUpdate: TNotifyEvent): TAction;
begin
  Result := TAction.Create(FActions);
  Result.ActionList  := FActions;
  Result.Caption     := ACaption;
  Result.Hint        := ACaption;
  Result.ImageIndex  := AGlyph;
  Result.ShortCut    := AShortCut;
  if AShortCut2 <> 0 then
    Result.SecondaryShortCuts.Add(ShortCutToText(AShortCut2));
  Result.OnExecute   := AOnExecute;
  Result.OnUpdate    := AOnUpdate;
end;

// Builds a compact, action-driven toolbar button. The action supplies the
// caption, glyph, enabled state and the click/shortcut behaviour - so there
// is exactly one code path per command.
procedure TMHGColumnsEditorForm.MakeButton(AAction: TAction; var X: Integer;
  AY: Integer; AExtraGap: Integer);
var
  TextW, BtnW: Integer;
  Btn: TSpeedButton;
begin
  Inc(X, AExtraGap);
  Btn := TSpeedButton.Create(Self);
  Btn.Parent   := FToolBar;
  Btn.Flat     := True;
  Btn.Images   := FGlyphs;
  Btn.Layout   := blGlyphLeft;
  Btn.Spacing  := 2;
  Btn.ShowHint := True;
  Btn.Action   := AAction;   // caption, glyph index, OnClick, enabled - all via action
  // Compact width = glyph + spacing + measured caption + small padding.
  // (TSpeedButton.AutoSize and TPanel.Canvas are protected, so we measure on
  // an independent bitmap canvas.)
  var Measure := TBitmap.Create;
  try
    Measure.Canvas.Font.Assign(Btn.Font);
    TextW := Measure.Canvas.TextWidth(AAction.Caption);
  finally
    Measure.Free;
  end;
  BtnW := GLYPH_SIZE + Btn.Spacing + TextW + 12;
  Btn.SetBounds(X, AY, BtnW, 26);
  Inc(X, Btn.Width + 2);
end;

procedure TMHGColumnsEditorForm.RefreshList(KeepIndex: Integer);
var
  I: Integer;
  Item: TCollectionItem;
begin
  FList.Items.BeginUpdate;
  try
    FList.Items.Clear;
    for I := 0 to FCollection.Count - 1 do
    begin
      Item := FCollection.Items[I];
      FList.Items.Add(Format('%d - %s', [I, Item.DisplayName]));
    end;
  finally
    FList.Items.EndUpdate;
  end;
  if (KeepIndex >= 0) and (KeepIndex < FList.Items.Count) then
  begin
    // In a multi-select listbox, ItemIndex only moves the focus rectangle;
    // it does not select. Select the item too, so SelCount reflects it and
    // the selection-dependent actions (Up/Down/Delete/Copy) stay enabled.
    FList.ClearSelection;
    FList.Selected[KeepIndex] := True;
    FList.ItemIndex := KeepIndex;
  end
  else if FList.Items.Count > 0 then
  begin
    FList.ClearSelection;
    FList.Selected[0] := True;
    FList.ItemIndex := 0;
  end;
end;

{ Action OnUpdate handlers - the actions enable/disable themselves on idle. }

procedure TMHGColumnsEditorForm.UpdHasSel(Sender: TObject);
begin
  (Sender as TAction).Enabled := FList.SelCount > 0;
end;

procedure TMHGColumnsEditorForm.UpdUp(Sender: TObject);
begin
  (Sender as TAction).Enabled := (FList.SelCount = 1) and (FList.ItemIndex > 0);
end;

procedure TMHGColumnsEditorForm.UpdDown(Sender: TObject);
begin
  (Sender as TAction).Enabled := (FList.SelCount = 1) and
    (FList.ItemIndex >= 0) and (FList.ItemIndex < FList.Items.Count - 1);
end;

procedure TMHGColumnsEditorForm.UpdPaste(Sender: TObject);
begin
  (Sender as TAction).Enabled := ClipboardHasColumns;
end;

procedure TMHGColumnsEditorForm.UpdImport(Sender: TObject);
var
  G: TMultiHeaderDBGrid;
begin
  // Enabled only when editing a DB grid whose DataSet is open (so there are
  // fields to import).
  if FGrid is TMultiHeaderDBGrid then
  begin
    G := TMultiHeaderDBGrid(FGrid);
    (Sender as TAction).Enabled := (G.DataSet <> nil) and G.DataSet.Active;
  end
  else
    (Sender as TAction).Enabled := False;
end;

procedure TMHGColumnsEditorForm.UpdHasItems(Sender: TObject);
begin
  (Sender as TAction).Enabled := FList.Items.Count > 0;
end;

procedure TMHGColumnsEditorForm.NotifyModified;
begin
  if FDesigner <> nil then
    FDesigner.Modified;
end;

procedure TMHGColumnsEditorForm.SyncInspector;
var
  Comps: IDesignerSelections;
  I: Integer;
begin
  // Push the selected column(s) into the Object Inspector so the user can
  // edit their properties while this modeless window stays open.
  if FDesigner = nil then Exit;
  Comps := CreateSelectionList;
  for I := 0 to FCollection.Count - 1 do
    if FList.Selected[I] then
      Comps.Add(FCollection.Items[I]);
  if Comps.Count > 0 then
    FDesigner.SetSelections(Comps);
end;

procedure TMHGColumnsEditorForm.ListClick(Sender: TObject);
begin
  SyncInspector;
end;

procedure TMHGColumnsEditorForm.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
  if not FTearingDown and (OpenEditors <> nil) and (FCollection <> nil) then
    OpenEditors.Remove(FCollection);
  Action := caFree;
end;

{ ---- toolbar actions ----------------------------------------------------- }

procedure TMHGColumnsEditorForm.DoAdd(Sender: TObject);
var
  Item: TCollectionItem;
begin
  Item := FCollection.Add;
  NotifyModified;
  RefreshList(Item.Index);
  SyncInspector;
end;

procedure TMHGColumnsEditorForm.DoDelete(Sender: TObject);
var
  I, First: Integer;
begin
  if FList.SelCount = 0 then Exit;
  First := FList.ItemIndex;
  // Delete from the back so indices stay valid.
  for I := FCollection.Count - 1 downto 0 do
    if FList.Selected[I] then
      FCollection.Items[I].Free;
  NotifyModified;
  RefreshList(First);
  SyncInspector;
end;

procedure TMHGColumnsEditorForm.DoUp(Sender: TObject);
var
  Idx: Integer;
begin
  Idx := FList.ItemIndex;
  if Idx <= 0 then Exit;
  FCollection.Items[Idx].Index := Idx - 1;
  NotifyModified;
  RefreshList(Idx - 1);
  SyncInspector;
end;

procedure TMHGColumnsEditorForm.DoDown(Sender: TObject);
var
  Idx: Integer;
begin
  Idx := FList.ItemIndex;
  if (Idx < 0) or (Idx >= FCollection.Count - 1) then Exit;
  FCollection.Items[Idx].Index := Idx + 1;
  NotifyModified;
  RefreshList(Idx + 1);
  SyncInspector;
end;

{ ---- clipboard helpers --------------------------------------------------- }

function TMHGColumnsEditorForm.SelectionToText: string;
var
  Carrier: TMHGClipCarrier;
  Bin, Txt: TMemoryStream;
  SS: TStringStream;
  I: Integer;
begin
  Result := '';
  if FList.SelCount = 0 then Exit;

  Carrier := TMHGClipCarrier.Create(nil);
  Bin := TMemoryStream.Create;
  Txt := TMemoryStream.Create;
  SS  := TStringStream.Create('', TEncoding.UTF8);
  try
    Carrier.FItems := TCollection.Create(FItemClass);
    try
      for I := 0 to FCollection.Count - 1 do
        if FList.Selected[I] then
          Carrier.FItems.Add.Assign(FCollection.Items[I]);
      Bin.WriteComponent(Carrier);
      Bin.Position := 0;
      ObjectBinaryToText(Bin, Txt);
      Txt.Position := 0;
      SS.CopyFrom(Txt, Txt.Size);
      Result := SS.DataString;
    finally
      Carrier.FItems.Free;
    end;
  finally
    SS.Free;
    Txt.Free;
    Bin.Free;
    Carrier.Free;
  end;
end;

function TMHGColumnsEditorForm.AppendFromText(const AText: string): Integer;
var
  Carrier: TMHGClipCarrier;
  Txt, Bin: TMemoryStream;
  Bytes: TBytes;
  I: Integer;
begin
  Result := FList.ItemIndex;
  if Trim(AText) = '' then Exit;

  Bytes := TEncoding.UTF8.GetBytes(AText);
  Carrier := TMHGClipCarrier.Create(nil);
  Txt := TMemoryStream.Create;
  Bin := TMemoryStream.Create;
  try
    Carrier.FItems := TCollection.Create(FItemClass);
    try
      if Length(Bytes) > 0 then
        Txt.WriteBuffer(Bytes[0], Length(Bytes));
      Txt.Position := 0;
      try
        ObjectTextToBinary(Txt, Bin);
        Bin.Position := 0;
        Bin.ReadComponent(Carrier);
      except
        on E: Exception do
        begin
          ShowMessage('Clipboard does not contain valid column data:'#13#10 +
                      E.Message);
          Exit;
        end;
      end;

      for I := 0 to Carrier.FItems.Count - 1 do
      begin
        var Dst := FCollection.Add;
        Dst.Assign(Carrier.FItems.Items[I]);
        Result := Dst.Index;
      end;
    finally
      Carrier.FItems.Free;
    end;
  finally
    Txt.Free;
    Bin.Free;
    Carrier.Free;
  end;
end;

procedure TMHGColumnsEditorForm.DoCopy(Sender: TObject);
var
  Txt: string;
  Bytes: TBytes;
  Data: THandle;
  P: Pointer;
begin
  Txt := SelectionToText;
  if Txt = '' then Exit;

  Bytes := TEncoding.UTF8.GetBytes(Txt + #0);
  Clipboard.Open;
  try
    Data := GlobalAlloc(GMEM_MOVEABLE or GMEM_DDESHARE, Length(Bytes));
    if Data <> 0 then
    begin
      P := GlobalLock(Data);
      try
        if P <> nil then
          Move(Bytes[0], P^, Length(Bytes));
      finally
        GlobalUnlock(Data);
      end;
      SetClipboardData(MHGClipFormat, Data);
    end;
    Clipboard.AsText := Txt;   // also CF_TEXT, for paste-as-text anywhere
  finally
    Clipboard.Close;
  end;
end;

procedure TMHGColumnsEditorForm.DoPaste(Sender: TObject);
var
  S: string;
  Data: THandle;
  P: PAnsiChar;
  NewIdx: Integer;
begin
  S := '';
  if Clipboard.HasFormat(MHGClipFormat) then
  begin
    Clipboard.Open;
    try
      Data := GetClipboardData(MHGClipFormat);
      if Data <> 0 then
      begin
        P := GlobalLock(Data);
        try
          if P <> nil then
            S := UTF8ToString(RawByteString(P));
        finally
          GlobalUnlock(Data);
        end;
      end;
    finally
      Clipboard.Close;
    end;
  end;

  if (S = '') and Clipboard.HasFormat(CF_TEXT) then
    S := Clipboard.AsText;
  if Trim(S) = '' then Exit;

  NewIdx := AppendFromText(S);
  NotifyModified;
  RefreshList(NewIdx);
  SyncInspector;
end;

procedure TMHGColumnsEditorForm.DoImport(Sender: TObject);
var
  G: TMultiHeaderDBGrid;
  Before: Integer;
begin
  if not (FGrid is TMultiHeaderDBGrid) then Exit;
  G := TMultiHeaderDBGrid(FGrid);
  if (G.DataSet = nil) or (not G.DataSet.Active) then
  begin
    ShowMessage('Cannot import columns: the grid''s DataSet is not open.');
    Exit;
  end;

  // AutoCreateColumns adds one column per visible field, skipping fields that
  // are already present - so Import both fills an empty grid and tops up a
  // partial one.
  Before := FCollection.Count;
  G.AutoCreateColumns;
  if FCollection.Count = Before then
  begin
    ShowMessage('Nothing to import: every visible field already has a column.');
    Exit;
  end;
  NotifyModified;
  RefreshList(FCollection.Count - 1);
  SyncInspector;
end;

procedure TMHGColumnsEditorForm.DoSelectAll(Sender: TObject);
var
  I: Integer;
begin
  if FList.Items.Count = 0 then Exit;
  FList.Items.BeginUpdate;
  try
    for I := 0 to FList.Items.Count - 1 do
      FList.Selected[I] := True;
  finally
    FList.Items.EndUpdate;
  end;
  SyncInspector;   // push the full selection to the Object Inspector
end;

function TMHGColumnsEditorForm.ClipboardHasColumns: Boolean;
begin
  Result := Clipboard.HasFormat(MHGClipFormat) or Clipboard.HasFormat(CF_TEXT);
end;

{ ---- show helper: one modeless instance per collection ------------------- }

procedure ShowColumnsEditor(const ADesigner: IDesigner; AGrid: TComponent;
  ACollection: TCollection; AItemClass: TCollectionItemClass;
  const ACaption: string);
var
  F: TMHGColumnsEditorForm;
begin
  if OpenEditors = nil then
    OpenEditors := TDictionary<TCollection, TMHGColumnsEditorForm>.Create;

  if OpenEditors.TryGetValue(ACollection, F) then
  begin
    F.BringToFront;
    F.Show;
    Exit;
  end;

  F := TMHGColumnsEditorForm.CreateEditor(ADesigner, AGrid, ACollection,
         AItemClass, ACaption);
  OpenEditors.Add(ACollection, F);
  F.Show;   // modeless - Object Inspector stays usable (request #4)
end;

{ ---- editors / property editors ------------------------------------------ }

type
  TMHGColumnsProperty = class(TPropertyEditor)
  public
    function GetAttributes: TPropertyAttributes; override;
    function GetValue: string; override;
    procedure Edit; override;
  end;

  TMHGDBGridEditor = class(TComponentEditor)
  public
    procedure Edit; override;
    function GetVerbCount: Integer; override;
    function GetVerb(Index: Integer): string; override;
    procedure ExecuteVerb(Index: Integer); override;
  end;

  TMHGGridEditor = class(TComponentEditor)
  public
    procedure Edit; override;
    function GetVerbCount: Integer; override;
    function GetVerb(Index: Integer): string; override;
    procedure ExecuteVerb(Index: Integer); override;
  end;

procedure EditDBColumns(const ADesigner: IDesigner; AGrid: TMultiHeaderDBGrid);
begin
  ShowColumnsEditor(ADesigner, AGrid, AGrid.Columns, TMHGColumn,
    Format('Editing %s.Columns', [AGrid.Name]));
end;

procedure EditHeaderColumns(const ADesigner: IDesigner; AGrid: TMultiHeaderGrid);
begin
  ShowColumnsEditor(ADesigner, AGrid, AGrid.HeaderColumns, TMHGHeaderColumn,
    Format('Editing %s.HeaderColumns', [AGrid.Name]));
end;

{ TMHGColumnsProperty }

function TMHGColumnsProperty.GetAttributes: TPropertyAttributes;
begin
  Result := [paDialog, paReadOnly];
end;

function TMHGColumnsProperty.GetValue: string;
begin
  Result := '(MultiHeaderGrid Columns)';
end;

procedure TMHGColumnsProperty.Edit;
var
  Comp: TPersistent;
begin
  Comp := GetComponent(0);
  if Comp is TMultiHeaderDBGrid then
    EditDBColumns(Designer, TMultiHeaderDBGrid(Comp))
  else if Comp is TMultiHeaderGrid then
    EditHeaderColumns(Designer, TMultiHeaderGrid(Comp));
end;

{ TMHGDBGridEditor }

procedure TMHGDBGridEditor.Edit;
begin
  EditDBColumns(Designer, Component as TMultiHeaderDBGrid);
end;

function TMHGDBGridEditor.GetVerbCount: Integer;
begin
  Result := 2;
end;

function TMHGDBGridEditor.GetVerb(Index: Integer): string;
begin
  case Index of
    0: Result := 'Edit Columns...';
    1: Result := 'Auto Create Columns';
  else
    Result := '';
  end;
end;

procedure TMHGDBGridEditor.ExecuteVerb(Index: Integer);
var
  Grid: TMultiHeaderDBGrid;
begin
  Grid := Component as TMultiHeaderDBGrid;
  case Index of
    0: EditDBColumns(Designer, Grid);
    1: begin
         Grid.AutoCreateColumns;
         if Designer <> nil then Designer.Modified;
       end;
  end;
end;

{ TMHGGridEditor }

procedure TMHGGridEditor.Edit;
begin
  EditHeaderColumns(Designer, Component as TMultiHeaderGrid);
end;

function TMHGGridEditor.GetVerbCount: Integer;
begin
  Result := 1;
end;

function TMHGGridEditor.GetVerb(Index: Integer): string;
begin
  if Index = 0 then Result := 'Edit Header...' else Result := '';
end;

procedure TMHGGridEditor.ExecuteVerb(Index: Integer);
begin
  if Index = 0 then
    EditHeaderColumns(Designer, Component as TMultiHeaderGrid);
end;

{ ---- registration -------------------------------------------------------- }

procedure Register;
begin
  RegisterComponentEditor(TMultiHeaderDBGrid, TMHGDBGridEditor);
  RegisterPropertyEditor(TypeInfo(TMHGColumns), TMultiHeaderDBGrid,
                         'Columns', TMHGColumnsProperty);

  RegisterComponentEditor(TMultiHeaderGrid, TMHGGridEditor);
  RegisterComponentEditor(TMultiHeaderStringGrid, TMHGGridEditor);
  RegisterPropertyEditor(TypeInfo(TMHGHeaderColumns), TMultiHeaderGrid,
                         'HeaderColumns', TMHGColumnsProperty);
end;

initialization
  MHGClipFormat := RegisterClipboardFormat(MHG_CLIP_FORMAT_NAME);
finalization
  if OpenEditors <> nil then
  begin
    // Detach from the registry first (FormClose would otherwise mutate the
    // dictionary while we iterate), then free the still-open windows.
    var Forms := OpenEditors.Values.ToArray;
    OpenEditors.Clear;
    for var F in Forms do
    begin
      F.FTearingDown := True;
      F.Free;
    end;
    FreeAndNil(OpenEditors);
  end;
end.

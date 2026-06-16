unit FMX.MultiHeaderGrid.Design;

{
  Design-time support for FMX.MultiHeaderGrid.

  - TMultiHeaderDBGrid: wires the Columns collection to the standard VCL
    collection editor (TCollectionProperty). Double-click on the grid in
    the form designer opens that editor (the "field editor"). Context-menu
    verbs: "Edit Columns..." and "Auto Create Columns".

  - TMultiHeaderGrid / TMultiHeaderStringGrid: these grids are not data
    bound and build their headers procedurally (Header.AddRow / AddColumn),
    so there is no field list and no FieldName. They get a component editor
    whose double-click opens the standard Fields/Columns-style collection
    editor for their HeaderColumns helper collection (captions, grouping,
    alignment, word wrap, widths - everything except FieldName).

  This unit uses DesignIntf / DesignEditors / ColnEdit and therefore MUST
  live in a design-time-only package (MultiHeaderGridDsgn.dpk). Do NOT add
  it to the run-time package.
}

interface

procedure Register;

implementation

uses
  System.Classes, System.SysUtils,
  DesignIntf, DesignEditors, ColnEdit,
  FMX.MultiHeaderGrid;

type
  TMHGColumnsProperty = class(TCollectionProperty);
  TMHGHeaderColumnsProperty = class(TCollectionProperty);

  // Data-bound grid editor; double-click opens the Columns (field) editor.
  TMHGDBGridEditor = class(TComponentEditor)
  public
    procedure Edit; override;
    function GetVerbCount: Integer; override;
    function GetVerb(Index: Integer): string; override;
    procedure ExecuteVerb(Index: Integer); override;
  end;

  // Non-data-bound grid editor; double-click opens the HeaderColumns editor.
  TMHGGridEditor = class(TComponentEditor)
  public
    procedure Edit; override;
    function GetVerbCount: Integer; override;
    function GetVerb(Index: Integer): string; override;
    procedure ExecuteVerb(Index: Integer); override;
  end;

{ TMHGDBGridEditor }

procedure TMHGDBGridEditor.Edit;
begin
  var Grid:=Component as TMultiHeaderDBGrid;
  ShowCollectionEditor(Designer, Grid, Grid.Columns, 'Columns');
end;

function TMHGDBGridEditor.GetVerbCount: Integer;
begin
  Result:=2;
end;

function TMHGDBGridEditor.GetVerb(Index: Integer): string;
begin
  case Index of
    0: Result:='Edit Columns...';
    1: Result:='Auto Create Columns';
  else
    Result:='';
  end;
end;

procedure TMHGDBGridEditor.ExecuteVerb(Index: Integer);
begin
  var Grid:=Component as TMultiHeaderDBGrid;
  case Index of
    0: ShowCollectionEditor(Designer, Grid, Grid.Columns, 'Columns');
    1: begin
         Grid.AutoCreateColumns;
         if Designer<>nil then Designer.Modified;
       end;
  end;
end;

{ TMHGGridEditor }

procedure TMHGGridEditor.Edit;
begin
  var Grid:=Component as TMultiHeaderGrid;
  ShowCollectionEditor(Designer, Grid, Grid.HeaderColumns, 'HeaderColumns');
end;

function TMHGGridEditor.GetVerbCount: Integer;
begin
  Result:=1;
end;

function TMHGGridEditor.GetVerb(Index: Integer): string;
begin
  if Index=0 then Result:='Edit Header...' else Result:='';
end;

procedure TMHGGridEditor.ExecuteVerb(Index: Integer);
begin
  var Grid:=Component as TMultiHeaderGrid;
  if Index=0 then
    ShowCollectionEditor(Designer, Grid, Grid.HeaderColumns, 'HeaderColumns');
end;

procedure Register;
begin
  RegisterComponentEditor(TMultiHeaderDBGrid, TMHGDBGridEditor);
  RegisterPropertyEditor(TypeInfo(TMHGColumns), TMultiHeaderDBGrid,
                         'Columns', TMHGColumnsProperty);

  RegisterComponentEditor(TMultiHeaderGrid, TMHGGridEditor);
  RegisterComponentEditor(TMultiHeaderStringGrid, TMHGGridEditor);
  RegisterPropertyEditor(TypeInfo(TMHGHeaderColumns), TMultiHeaderGrid,
                         'HeaderColumns', TMHGHeaderColumnsProperty);
end;

end.

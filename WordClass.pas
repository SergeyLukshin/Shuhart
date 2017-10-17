unit WordClass;

interface

uses ComObj, System.Variants, Vcl.Dialogs, WinApi.Windows, SysUtils, ActiveX;

type
  TCellVerticalAlignment = (CellAlignVerticalTop = 0, CellAlignVerticalCenter = 1, CellAlignVerticalBottom = 3);
  TParagraphAlignment = (AlignParagraphLeft = 0, AlignParagraphCenter = 1, AlignParagraphRight = 2, AlignParagraphJustify = 3);

  TPair = record
  public
    strName: string;
    strValue: string;
  end;

  TWord = class
    private
      m_pApp : variant;
      m_pDoc : variant;
      m_pTable : variant;
      m_pSelection : variant;
      old_table_index: integer;

      procedure WordDocumentToFront(wd:OLEVariant);
      procedure GetTableReference(indTable: integer);

    public
      constructor Create;
      Destructor  Destroy; override;
      function Start(strFileName : string) : boolean;
      procedure SetVisible(bVisible: boolean);
      procedure Quit();
      procedure Release();
      procedure SetBookmarkText(var mBookMark: array of TPair);
      procedure InsertRowsInTable (indTable: integer; nCount: integer); // индексы с 1
      procedure SetTextInCell(Column : integer; Row: integer; strText: string; indTable: integer); // индексы с 1
      procedure ConvertToTable(strText: string; row_cnt, col_cnt: integer; bSetBorders: boolean; bookmark: string);
      procedure DeleteColumn(indCol: integer; indTable: integer);
      procedure SetCellsMerge(FirstColumn: integer; LastColumn: integer;
        FirstRow: integer; LastRow: integer; vAlignment: TCellVerticalAlignment; indTable: integer  );
      procedure InsertPicture(strFileName: string; hAlign: TParagraphAlignment; bookmark: string);
      procedure SetColumnWidth(Column: integer; Width_: Single; indTable: integer);
      function VarFromSingle(const Value: Single): Variant;
end;

implementation

constructor TWord.Create();
begin
  //
end;

destructor TWord.Destroy;
begin
  // Освобождение памяти, если она получена
  if (not m_pApp.Visible) then
    Quit()
  else
    Release();

  // Всегда вызывайте родительский деструктор после выполнения вашего собственного кода
  inherited;
end;

procedure TWord.Quit();
begin
  if (NOT VarIsEmpty(m_pApp)) then begin
		m_pApp.Quit(false);
	end;
end;

procedure TWord.Release();
begin
  if (NOT VarIsEmpty(m_pSelection)) then begin
		m_pSelection := UnAssigned;
	end;

  if (NOT VarIsEmpty(m_pTable)) then begin
		m_pTable := UnAssigned;
	end;

  if (NOT VarIsEmpty(m_pDoc)) then begin
		m_pDoc := UnAssigned;
	end;

  if (NOT VarIsEmpty(m_pApp)) then begin
		m_pApp := UnAssigned;
	end;
end;

function TWord.Start(strFileName : string) : boolean;
begin
  m_pApp := UnAssigned;
  m_pDoc := UnAssigned;
  m_pTable := UnAssigned;
  m_pSelection := UnAssigned;

  old_table_index := 0;

  try
    m_pApp := CreateOleObject('Word.Application');

    if (strFileName = '') then begin
      m_pDoc := m_pApp.Documents.Add;
    end
    else begin
      m_pDoc := m_pApp.Documents.Add(strFileName);
    end;

    result := true;
  except
    on E: EOleException do begin
        ShowMessage(e.Message);
        result := false;
    end
    else
      result := false;
  end;
end;

procedure TWord.WordDocumentToFront(wd:OLEVariant);
var
  s: string;
  h: HWnd;
begin
  s := wd.activedocument.ActiveWindow.Caption + ' - ' + wd.Caption;
  h := FindWindow(nil,pchar(s));
  if h<>0 then SetForegroundWindow(h);
end;

procedure TWord.SetVisible(bVisible: boolean);
begin
	if (VarIsEmpty(m_pApp)) then exit;
  WordDocumentToFront(m_pApp);
	m_pApp.Visible := bVisible;
	m_pApp.Activate;
end;

procedure TWord.SetBookmarkText(var mBookMark: array of TPair);
var i: integer;
begin
	if (Length(mBookMark) = 0) then exit;

	if (VarIsEmpty(m_pSelection)) then begin
		if (VarIsEmpty(m_pDoc)) then m_pDoc := m_pApp.GetActiveDocument();
		if (VarIsEmpty(m_pDoc)) then exit;

		m_pDoc.Select;
		m_pSelection := m_pApp.Selection;
		if (VarIsEmpty(m_pSelection)) then exit;
	end;

	for i := 0 to Length(mBookMark) - 1 do begin
    if (mBookMark[i].strName = '') then continue;

		m_pSelection.GoTo(-1, unassigned, unassigned, mBookMark[i].strName);
		m_pSelection.Text := mBookMark[i].strValue;
	end;
end;

procedure TWord.InsertRowsInTable (indTable: integer; nCount: integer);
var
  pSel, pTables: variant;
  i: integer;
begin
	if (nCount <= 0) then exit;

	if (not VarIsEmpty(m_pApp)) then begin
		if (VarIsEmpty(m_pDoc)) then m_pDoc := m_pApp.GetActiveDocument();
		if (VarIsEmpty(m_pDoc)) then exit;

		pTables := m_pDoc.Tables;
		if (VarIsEmpty(pTables)) then exit;
		if (pTables.Count = 0) then exit;

    //i := pTables.Count;

		m_pTable := pTables.Item(indTable);
		if (VarIsEmpty(m_pTable)) then exit;

		m_pTable.Select;

		pSel := m_pApp.Selection;
		pSel.InsertRowsBelow(nCount);

		m_pDoc.Select;
	  m_pSelection := m_pApp.Selection;
		if (VarIsEmpty(m_pSelection)) then exit;
		m_pSelection.Collapse(0);
	end;
end;

procedure TWord.SetTextInCell(Column : integer; Row: integer; strText: string; indTable: integer);
var
  top, left: integer;
  pCols, pRows, pCell, pSel: variant;
begin
	GetTableReference(indTable);

	// получаем границы диапазона
	top := 1;
  left := 1;

	pCols := m_pTable.Columns;
	if (VarIsEmpty(pCols)) then exit;
	pRows := m_pTable.Rows;
	if (VarIsEmpty(pRows)) then exit;

	if (Column >= 0) then left := Column;
	if (Row >= 0) then top := Row;

	// проверяем правильность установки границ диапазона
	// получаем верхнюю левую ячейку диапазона
	pCell := UnAssigned;
	pCell := m_pTable.Cell[top, left];
	if (VarIsEmpty(pCell)) then exit;

	pCell.Select;
	pSel := m_pApp.Selection;
	pSel.EndKey;

	pSel.Text := strText;
	pSel.Collapse(0);
end;

procedure TWord.GetTableReference(indTable: integer);
var
  pTables: variant;
begin
	if (indTable = -1) then begin
		if (VarIsEmpty(m_pDoc)) then m_pDoc := m_pApp.GetActiveDocument();
		if (VarIsEmpty(m_pDoc)) then exit;

		pTables := m_pDoc.Tables;
		if (VarIsEmpty(pTables)) then exit;
		if (pTables.Count = 0) then exit;

		m_pTable := pTables.Item(pTables.Count);
		if (VarIsEmpty(m_pTable)) then exit;

		old_table_index := pTables.Count;
		exit;
	end;

	if (indTable <> old_table_index) then begin
		if (VarIsEmpty(m_pDoc)) then m_pDoc := m_pApp.GetActiveDocument();
		if (VarIsEmpty(m_pDoc)) then exit;

		pTables := m_pDoc.Tables;
		if (VarIsEmpty(pTables)) then exit;
		if (pTables.Count = 0) or (pTables.Count < indTable) then exit;

		m_pTable := pTables.Item(indTable);
		if (VarIsEmpty(m_pTable)) then exit;

		old_table_index := indTable;
		exit;
	end;

	if (VarIsEmpty(m_pTable)) then begin
		if (VarIsEmpty(m_pDoc)) then m_pDoc := m_pApp.GetActiveDocument();
		if (VarIsEmpty(m_pDoc)) then exit;

		pTables := m_pDoc.Tables;
		if (VarIsEmpty(pTables)) then exit;
		if (pTables.Count = 0) then exit;

		m_pTable := pTables.Item(pTables.Count);
		if (VarIsEmpty(m_pTable)) then exit;

		old_table_index := pTables.Count;
		exit;
	end;
end;

procedure TWord.ConvertToTable(strText: string; row_cnt, col_cnt: integer; bSetBorders: boolean; bookmark: string);
var
  pRows, pBorders: variant;
begin

  if (bookmark = '') then begin
 		if (VarIsEmpty(m_pDoc)) then m_pDoc := m_pApp.GetActiveDocument();
		if (VarIsEmpty(m_pDoc)) then exit;

		m_pDoc.Select;
		m_pSelection := m_pApp.Selection;
		if (VarIsEmpty(m_pSelection)) then exit;

  	m_pSelection.Collapse(0);
  end
  else begin
    m_pSelection.GoTo(-1, unassigned, unassigned, bookmark);
  end;

	m_pSelection.Text := strText;

	m_pTable := m_pSelection.ConvertToTable(1, row_cnt, col_cnt, 0, 0,
		unassigned, unassigned, unassigned, unassigned, unassigned, unassigned, unassigned, unassigned,
		unassigned, 0, 0);

	// выравниваем талицу (не содержимое!!!)
	if (VarIsEmpty(m_pTable)) then exit;

	pRows := m_pSelection.Rows;
	pRows.AllowBreakAcrossPages := false;

	m_pSelection.Collapse(0);

	pBorders := m_pTable.Borders;
	if (VarIsEmpty(pBorders)) then exit;
	pBorders.Enable := bSetBorders;
end;

procedure TWord.DeleteColumn(indCol: integer; indTable: integer);
var
  pCols, pCol : variant;
begin
  GetTableReference(indTable);
  pCols := m_pTable.Columns;
  if (VarIsEmpty(pCols)) then exit;
  pCol := pCols.Item[indCol];
  pCol.Delete;
end;

procedure TWord.SetCellsMerge(FirstColumn: integer; LastColumn: integer;
      FirstRow: integer; LastRow: integer; vAlignment: TCellVerticalAlignment; indTable: integer);
var
  pCellBegin, pCellEnd: variant;
begin
	GetTableReference(indTable);

	pCellBegin := m_pTable.Cell[FirstRow, FirstColumn];
	pCellEnd := m_pTable.Cell[LastRow, LastColumn];
	if (VarIsEmpty(pCellBegin)) then exit;
  if (VarIsEmpty(pCellEnd)) then exit;

	pCellBegin.Merge(pCellEnd);
	pCellBegin.VerticalAlignment := vAlignment;
end;

procedure TWord.InsertPicture(strFileName: string; hAlign: TParagraphAlignment; bookmark: string);
var pParFormat, shapes: variant;
begin

	if (bookmark = '') then begin
 		if (VarIsEmpty(m_pDoc)) then m_pDoc := m_pApp.GetActiveDocument();
		if (VarIsEmpty(m_pDoc)) then exit;

		m_pDoc.Select;
		m_pSelection := m_pApp.Selection;
		if (VarIsEmpty(m_pSelection)) then exit;

  	m_pSelection.Collapse(0);
  end
  else begin
    m_pSelection.GoTo(-1, unassigned, unassigned, bookmark);
  end;

	pParFormat := m_pSelection.ParagraphFormat;
	if (VarIsEmpty(pParFormat)) then exit;
  pParFormat.Alignment := hAlign;
	shapes := m_pSelection.InlineShapes;
  if (VarIsEmpty(shapes)) then exit;
  shapes.AddPicture(strFileName, false, true);
end;

function TWord.VarFromSingle(const Value: Single): Variant;
begin
  VarClear(Result);
  TVarData(Result).VSingle := Value;
  TVarData(Result).VType := varSingle;
end;

procedure TWord.SetColumnWidth(Column: integer; Width_: Single; indTable: integer);
var
  pColumns, pColumn: variant;
  cntColumns: integer;
  w: Single;
  retval: HRESULT;
  disp: IDispatch;
  Params: TDispParams;
  param, result: variant;
begin
	GetTableReference(indTable);

	pColumns := m_pTable.Columns;
	if (VarIsEmpty(pColumns)) then exit;
	cntColumns := pColumns.Count;
  if (Column >= 1) and (Column <= cntColumns) then begin
		pColumn := pColumns.Item(Column);
		if (VarIsEmpty(pColumn)) then exit;
    //w := 2.835 * Width_;//m_pApp.MillimetersToPoints(Width);
    disp := IDispatch(m_pApp);

    param := VarFromSingle(Width_ / 10);
    Params := Default(TDispParams);
    Params.cArgs := 1;
    Params.rgvarg := @param;
    retval := disp.Invoke(
      371,//CentimetersToPoints
      GUID_NULL,
      LOCALE_USER_DEFAULT,
      DISPATCH_METHOD or DISPATCH_PROPERTYGET,
      Params,
      @Result,
      nil,
      nil
    );
    w := TVarData(Result).VSingle;
	  pColumn.Width := w;
  end;
end;

end.

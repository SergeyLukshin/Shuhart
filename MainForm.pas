unit MainForm;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, cxGraphics, cxControls, cxLookAndFeels,
  cxLookAndFeelPainters, cxStyles, dxSkinsCore, cxCustomData,
  cxFilter, cxData, cxDataStorage, cxEdit, cxNavigator, Data.DB, cxDBData,
  cxGridCustomTableView, cxGridTableView, cxGridDBTableView, cxGridLevel,
  cxClasses, cxGridCustomView, cxGrid, cxMaskEdit, cxContainer, cxTextEdit,
  VclTee.TeeGDIPlus, VCLTee.TeEngine, VCLTee.Series, Vcl.ExtCtrls,
  VCLTee.TeeProcs, VCLTee.Chart, cxPCdxBarPopupMenu, cxPC, cxGridBandedTableView,
  cxCurrencyEdit, Vcl.StdCtrls, cxDropDownEdit, cxGroupBox, cxCheckBox,
  Vcl.Menus, Xml.xmldom, Xml.XMLIntf, Xml.Win.msxmldom, Xml.XMLDoc, AboutForm,
  dxSkinBlack, dxSkinBlue, dxSkinBlueprint, dxSkinCaramel, dxSkinCoffee,
  dxSkinDarkRoom, dxSkinDarkSide, dxSkinDevExpressDarkStyle,
  dxSkinDevExpressStyle, dxSkinFoggy, dxSkinGlassOceans, dxSkinHighContrast,
  dxSkiniMaginary, dxSkinLilian, dxSkinLiquidSky, dxSkinLondonLiquidSky,
  dxSkinMcSkin, dxSkinMoneyTwins, dxSkinOffice2007Black, dxSkinOffice2007Blue,
  dxSkinOffice2007Green, dxSkinOffice2007Pink, dxSkinOffice2007Silver,
  dxSkinOffice2010Black, dxSkinOffice2010Blue, dxSkinOffice2010Silver,
  dxSkinOffice2013White, dxSkinPumpkin, dxSkinSeven, dxSkinSevenClassic,
  dxSkinSharp, dxSkinSharpPlus, dxSkinSilver, dxSkinSpringTime, dxSkinStardust,
  dxSkinSummer2008, dxSkinTheAsphaltWorld, dxSkinsDefaultPainters,
  dxSkinValentine, dxSkinVS2010, dxSkinWhiteprint, dxSkinXmas2008Blue,
  dxSkinscxPCPainter, WordClass, Vcl.Clipbrd, Vcl.ImgList, System.Actions,
  Vcl.ActnList, Vcl.ToolWin, Vcl.ActnMan, Vcl.ActnCtrls, Vcl.ActnMenus,
  Vcl.PlatformDefaultStyleActnCtrls, Vcl.ComCtrls, dxCore, cxDateUtils,
  cxCalendar, DateUtils, cxMemo, LoginForm;

const
  X_avg_col_index = 21;
  S_col_index = 22;
  Me_col_index = 23;
  R_col_index = 24;
  CV_col_index = 25;

  month_str:array [1..12] of string=('января','февраля','марта','апреля','мая',
  'июня','июля','августа','сентября','октября','ноября','декабря');

type

  TMCIDataInfo = record
    strLaboratoryName: string;
    strOKName: string;
    val1: variant;
    val2: variant;
    val3: variant;
    val4: variant;
    val5: variant;
  end;

  TMCILaboratoryInfo = record
    strLaboratoryName: string;
    bResult: boolean;
  end;

  TCalcResult = record
    C_val, CDelta_val, Avg_val: Currency;
    Kohren_val, SKO_Povt_val, SKO_Prec_val, SKO_Povt_otn_val, SKO_Prec_otn_val, Move_val,
    SKO_System_Pogr_val, Student_val, System_Pogr_val, Verify_val, AnalysPogr_val,
    MethodPogr_val: double;
    Limit_Povt_val, Limit_Prec_val, Precision_val: integer;
    bKohren_val, bSKO_Prec_val, bSKO_System_Pogr_val, bStudent_val,
    bVerify_val, bMethodPogr_val, bSKO_Povt_otn_val, bSKO_Prec_otn_val,
    bLimit_Povt_val, bLimit_Prec_val, bPrecision_val: boolean;

    Kohren_tab_val, Student_tab_val: Currency;
    bKohren_tab_val, bStudent_tab_val: boolean;
    bPCI_val: boolean; PCI_val: double;
    bWarningShuhart: boolean;
    bCriticalShuhart: boolean;
    strZakl: string;
    strZaklForFill: string;
    arrMCILabInfo: array of TMCILaboratoryInfo;
  end;

  TCurrencyValue = class
    private
      // Поля данных этого нового класса
      value   : Currency;

    public
      // Свойства для чтения значений этих данных
      property val : Currency
          read value;

      // Коструктрор
      constructor Create(const value : Currency);
  end;

  TCriteriaInfo = class
    private
      // Поля данных этого нового класса
      index_beg   : integer;
      index_end   : integer;

    public
      // Свойства для чтения значений этих данных
      property ind_beg : integer
          read index_beg;
      property ind_end : integer
          read index_end;

      // Коструктрор
      constructor Create(const ind_beg : integer; const ind_end : integer);
  end;

  TShuhartForm = class(TForm)
    cxStyleRepository1: TcxStyleRepository;
    cxStyle1: TcxStyle;
    cxPageControl2: TcxPageControl;
    cxTabSheet2: TcxTabSheet;
    cxGrid: TcxGrid;
    cxGridLevel1: TcxGridLevel;
    cxTabSheet4: TcxTabSheet;
    cxTabSheet5: TcxTabSheet;
    cxStyle2: TcxStyle;
    cxGridView: TcxGridBandedTableView;
    cxColumn1: TcxGridBandedColumn;
    cxColumnX: TcxGridBandedColumn;
    cxStyle3: TcxStyle;
    cxTabSheet1: TcxTabSheet;
    cxTabSheet3: TcxTabSheet;
    cxTabSheet6: TcxTabSheet;
    cxColumnX1: TcxGridBandedColumn;
    cxColumnX2: TcxGridBandedColumn;
    cxColumnX3: TcxGridBandedColumn;
    cxColumnX4: TcxGridBandedColumn;
    cxColumnX5: TcxGridBandedColumn;
    cxColumnS: TcxGridBandedColumn;
    cxColumnMe: TcxGridBandedColumn;
    cxColumnR: TcxGridBandedColumn;
    R_Chart: TChart;
    R_Chart_Data: TLineSeries;
    R_Chart_Mid: TLineSeries;
    R_Chart_UCL: TLineSeries;
    R_Chart_LCL: TLineSeries;
    XMLDoc: TXMLDocument;
    DlgOpen: TOpenDialog;
    DlgSave: TSaveDialog;
    cxStyle4: TcxStyle;
    X_Chart: TChart;
    X_Chart_Data: TLineSeries;
    X_Chart_Mid: TLineSeries;
    X_Chart_UCL: TLineSeries;
    X_Chart_LCL: TLineSeries;
    S_Chart: TChart;
    S_Chart_Data: TLineSeries;
    S_Chart_Mid: TLineSeries;
    S_Chart_UCL: TLineSeries;
    S_Chart_LCL: TLineSeries;
    Me_Chart: TChart;
    Me_Chart_Data: TLineSeries;
    Me_Chart_Mid: TLineSeries;
    Me_Chart_UCL: TLineSeries;
    Me_Chart_LCL: TLineSeries;
    X_Chart_CH: TLineSeries;
    X_Chart_CL: TLineSeries;
    X_Chart_BH: TLineSeries;
    X_Chart_BL: TLineSeries;
    cxTabSheet7: TcxTabSheet;
    cxGrid1: TcxGrid;
    cxGridLevel2: TcxGridLevel;
    cxGridCriteriaView: TcxGridTableView;
    cxGridCriteriaViewColumn1: TcxGridColumn;
    cxGridCriteriaViewColumn2: TcxGridColumn;
    cxGridCriteriaViewColumn3: TcxGridColumn;
    cxGridCriteriaViewColumn4: TcxGridColumn;
    cxGroupBox1: TcxGroupBox;
    PCI_label: TLabel;
    cxColumnX6: TcxGridBandedColumn;
    cxColumnX7: TcxGridBandedColumn;
    cxColumnX8: TcxGridBandedColumn;
    cxColumnX9: TcxGridBandedColumn;
    cxColumnX10: TcxGridBandedColumn;
    cxColumnX11: TcxGridBandedColumn;
    cxColumnX12: TcxGridBandedColumn;
    cxColumnX13: TcxGridBandedColumn;
    cxColumnX14: TcxGridBandedColumn;
    cxColumnX15: TcxGridBandedColumn;
    cxColumnX16: TcxGridBandedColumn;
    cxColumnX17: TcxGridBandedColumn;
    cxColumnX18: TcxGridBandedColumn;
    cxColumnX19: TcxGridBandedColumn;
    cxColumnX20: TcxGridBandedColumn;
    cxImageList1: TcxImageList;
    cxColumnCV: TcxGridBandedColumn;
    cxPageControl1: TcxPageControl;
    cxTabSheet8: TcxTabSheet;
    cxTabSheet9: TcxTabSheet;
    cxGroupBoxC: TcxGroupBox;
    Label12: TLabel;
    cxC: TcxCurrencyEdit;
    Label13: TLabel;
    cxCDelta: TcxCurrencyEdit;
    lUnit: TLabel;
    lUnit2: TLabel;
    cxTabSheet10: TcxTabSheet;
    cxGroupBox2: TcxGroupBox;
    Label18: TLabel;
    cxKohren_val: TcxTextEdit;
    lKohren_tab: TLabel;
    Label20: TLabel;
    cxSKO_povt_val: TcxTextEdit;
    lUnit6: TLabel;
    Label22: TLabel;
    cxSKO_prec_val: TcxTextEdit;
    lUnit5: TLabel;
    Label24: TLabel;
    cxAvg_val: TcxTextEdit;
    lUnit7: TLabel;
    Label26: TLabel;
    cxSKO_povt_otn_val: TcxTextEdit;
    Label27: TLabel;
    Label28: TLabel;
    cxSKO_prec_otn_val: TcxTextEdit;
    Label29: TLabel;
    cxGroupBox3: TcxGroupBox;
    Label30: TLabel;
    lUnit8: TLabel;
    Label32: TLabel;
    lStudent_tab: TLabel;
    Label34: TLabel;
    lUnit9: TLabel;
    Label36: TLabel;
    lUnit10: TLabel;
    Label38: TLabel;
    Label40: TLabel;
    lUnit11: TLabel;
    cxMove_val: TcxTextEdit;
    cxSKO_System_pogr_val: TcxTextEdit;
    cxStudent_val: TcxTextEdit;
    cxSystem_Pogr_val: TcxTextEdit;
    cxVerify_val: TcxTextEdit;
    cxAnalysPogr_val: TcxTextEdit;
    Label39: TLabel;
    cxMethodPogr_val: TcxTextEdit;
    Label42: TLabel;
    cxGroupBox4: TcxGroupBox;
    Label43: TLabel;
    Label44: TLabel;
    Label45: TLabel;
    Label46: TLabel;
    Label47: TLabel;
    Label48: TLabel;
    cxLimit_Povt_val: TcxTextEdit;
    cxLimit_Prec_val: TcxTextEdit;
    cxPrecision_val: TcxTextEdit;
    S_Chart_CH: TLineSeries;
    S_Chart_CL: TLineSeries;
    S_Chart_BH: TLineSeries;
    S_Chart_BL: TLineSeries;
    Me_Chart_CH: TLineSeries;
    Me_Chart_CL: TLineSeries;
    Me_Chart_BH: TLineSeries;
    Me_Chart_BL: TLineSeries;
    cxTabSheet11: TcxTabSheet;
    SKO_Chart: TChart;
    Conv_Chart: TLineSeries;
    Sigma_Chart_Min: TLineSeries;
    Sigma_Chart_Mid: TLineSeries;
    Sigma_Chart_Max: TLineSeries;
    cxPageControl3: TcxPageControl;
    cxTabSheet12: TcxTabSheet;
    cxTabSheet13: TcxTabSheet;
    cxGrid2: TcxGrid;
    cxGridLevel3: TcxGridLevel;
    cxGroupBox5: TcxGroupBox;
    Label19: TLabel;
    cxLaboratoryCnt: TcxComboBox;
    Label21: TLabel;
    cxOKCnt: TcxComboBox;
    cxGridLaboratoryView: TcxGridTableView;
    cxColumnLaboratory: TcxGridColumn;
    cxGrid3: TcxGrid;
    cxGridResultView: TcxGridTableView;
    cxGridColumn1: TcxGridColumn;
    cxColumnOK: TcxGridColumn;
    cxGridColumn3: TcxGridColumn;
    cxGridLevel4: TcxGridLevel;
    cxColumnNumber: TcxGridColumn;
    cxGridResultViewColumn1: TcxGridColumn;
    cxGridResultViewColumn2: TcxGridColumn;
    cxGridResultViewColumn3: TcxGridColumn;
    cxGridResultViewColumn4: TcxGridColumn;
    cxGridResultViewColumn5: TcxGridColumn;
    cxGridResultViewColumn6: TcxGridColumn;
    cxStyle5: TcxStyle;
    cxGridResultViewColumn7: TcxGridColumn;
    cxGroupBox6: TcxGroupBox;
    Label17: TLabel;
    Label25: TLabel;
    Label31: TLabel;
    cxGroupBox7: TcxGroupBox;
    cxGroupBox8: TcxGroupBox;
    cxGroupBox9: TcxGroupBox;
    MainMenu1: TMainMenu;
    N1: TMenuItem;
    mNew: TMenuItem;
    mOpen: TMenuItem;
    mSave: TMenuItem;
    N2: TMenuItem;
    mReport1: TMenuItem;
    mReport2: TMenuItem;
    mAbout: TMenuItem;
    Label54: TLabel;
    cxMCIUnitName: TcxTextEdit;
    cxGroupBox11: TcxGroupBox;
    Label3: TLabel;
    cxUnitName: TcxTextEdit;
    cxMeasureName: TcxComboBox;
    Label2: TLabel;
    Label56: TLabel;
    cxLaboratoryType: TcxComboBox;
    cxCardName: TcxTextEdit;
    Label1: TLabel;
    cxGroupBox10: TcxGroupBox;
    Label57: TLabel;
    Label58: TLabel;
    Label59: TLabel;
    Label60: TLabel;
    cxStandardValues: TcxCheckBox;
    cxStandardValuesGroup: TcxGroupBox;
    Label10: TLabel;
    Label9: TLabel;
    Label8: TLabel;
    Label7: TLabel;
    cxR0: TcxCurrencyEdit;
    cxX0: TcxCurrencyEdit;
    cxS0: TcxCurrencyEdit;
    cxSigma0: TcxCurrencyEdit;
    cxGroupBox12: TcxGroupBox;
    Label14: TLabel;
    Label16: TLabel;
    cxC_val: TcxTextEdit;
    lUnit3: TLabel;
    lUnit4: TLabel;
    cxCDelta_val: TcxTextEdit;
    Label61: TLabel;
    cxGroupBox13: TcxGroupBox;
    Label62: TLabel;
    cxRContr_val: TcxMaskEdit;
    Label63: TLabel;
    Label64: TLabel;
    cxR2Contr_val: TcxMaskEdit;
    Label65: TLabel;
    Label55: TLabel;
    cxRContr: TcxCurrencyEdit;
    cxR2Contr: TcxCurrencyEdit;
    cxGroupBox14: TcxGroupBox;
    lZakl: TLabel;
    Label33: TLabel;
    Label35: TLabel;
    Label37: TLabel;
    Label41: TLabel;
    Label49: TLabel;
    Label50: TLabel;
    Label51: TLabel;
    Label52: TLabel;
    Label53: TLabel;
    cxGroupBox15: TcxGroupBox;
    Label5: TLabel;
    cxDigitCnt: TcxComboBox;
    Label6: TLabel;
    cxMeasureCnt: TcxComboBox;
    Label11: TLabel;
    cxRowCnt: TcxMaskEdit;
    Label23: TLabel;
    cxChartFontSize: TcxComboBox;
    Label66: TLabel;
    Label67: TLabel;
    cxDateBegin: TcxDateEdit;
    cxDateEnd: TcxDateEdit;
    Label68: TLabel;
    Label4: TLabel;
    Label15: TLabel;
    cxCardType: TcxComboBox;
    cxAnalyzeType: TcxComboBox;
    Label69: TLabel;
    cxMCILaboratoryType: TcxComboBox;
    Label70: TLabel;
    cxMCIMeasureName: TcxComboBox;
    Label71: TLabel;
    Label72: TLabel;
    cxMCIDateBegin: TcxDateEdit;
    Label73: TLabel;
    cxMCIDateEnd: TcxDateEdit;
    Label74: TLabel;
    cxMCIModel: TcxComboBox;
    Label75: TLabel;
    Label76: TLabel;
    cxMCIDescript: TcxMemo;
    cxMeasureMethod: TcxMemo;
    cxMCIMeasureMethod: TcxMemo;
    procedure RecalcRowData(row: integer);
    procedure RecalcResultRowData(row: integer);
    procedure EnableStandardValues();
    procedure InitGrid();
    procedure InitLaboratoryGrid();
    procedure InitResultMCIGrid(bRecalcData: boolean);
    procedure InitData(bInitGrid: boolean);
    procedure InitDataMCI(bInitGrid: boolean);
    procedure CalcDataKoef();
    procedure CalcResultMCI();
    procedure CalcCriteries();
    procedure CalcVerification();
    procedure FillVerification();
    procedure DrawRCard();
    procedure DrawXCard();
    procedure DrawSCard();
    procedure DrawMeCard();
    function VerifyCriteria_1(var arr_val : array of Currency; index: integer): boolean;
    function VerifyCriteria_2(var arr_val : array of Currency; index: integer): boolean;
    function VerifyCriteria_3(var arr_val : array of Currency; index: integer): boolean;
    function VerifyCriteria_4(var arr_val : array of Currency; index: integer): boolean;
    function VerifyCriteria_5(var arr_val : array of Currency; index: integer): boolean;
    function VerifyCriteria_6(var arr_val : array of Currency; index: integer): boolean;
    function VerifyCriteria_7(var arr_val : array of Currency; index: integer): boolean;
    function VerifyCriteria_8(var arr_val : array of Currency; index: integer): boolean;
    function InsertClipboardInGrid(row_beg: integer; col_beg: integer): integer;
    function InsertClipboardInGridMCI(row_beg: integer; col_beg: integer): integer;
    procedure DelimitedText(mmString: string; DelimitChar: char; var ResultList: TStringList);
    procedure DrawConvChart();
    function VerifyDoubleLaboratory(): boolean;

    //procedure PrintHeader(var word: TWord; bShuhart: boolean);
    procedure PrintReport1(var word: TWord; DIR: string);
    procedure PrintReport2(var word: TWord);

    procedure FormShow(Sender: TObject);
    procedure cxStandardValuesClick(Sender: TObject);
    procedure cxMeasureCntFocusChanged(Sender: TObject);
    procedure cxPageControl2Change(Sender: TObject);
    procedure cxDigitCntFocusChanged(Sender: TObject);
    procedure cxColumnX1PropertiesValidate(Sender: TObject;
      var DisplayValue: Variant; var ErrorText: TCaption; var Error: Boolean);
    procedure FormCreate(Sender: TObject);
    procedure cxX0FocusChanged(Sender: TObject);
    procedure cxR0FocusChanged(Sender: TObject);
    procedure cxS0PropertiesChange(Sender: TObject);
    procedure cxSigma0FocusChanged(Sender: TObject);
    procedure mNewClick(Sender: TObject);
    procedure mLoadClick(Sender: TObject);
    procedure mSaveClick(Sender: TObject);
    procedure cxS0FocusChanged(Sender: TObject);
    procedure cxGridCriteriaViewCustomDrawCell(Sender: TcxCustomGridTableView;
      ACanvas: TcxCanvas; AViewInfo: TcxGridTableDataCellViewInfo;
      var ADone: Boolean);
    procedure cxCardTypeFocusChanged(Sender: TObject);
    procedure N3Click(Sender: TObject);
    procedure cxRowCntFocusChanged(Sender: TObject);
    procedure X_ChartMouseMove(Sender: TObject; Shift: TShiftState; X,
      Y: Integer);
    procedure X_ChartAfterDraw(Sender: TObject);
    procedure cxGridViewKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure cxGridViewEditKeyDown(Sender: TcxCustomGridTableView;
      AItem: TcxCustomGridTableItem; AEdit: TcxCustomEdit; var Key: Word;
      Shift: TShiftState);
    procedure cxCFocusChanged(Sender: TObject);
    procedure cxCDeltaFocusChanged(Sender: TObject);
    procedure cxAnalyzeTypeFocusChanged(Sender: TObject);
    procedure Report1Execute(Sender: TObject);
    procedure cxLaboratoryCntPropertiesEditValueChanged(Sender: TObject);
    procedure cxPageControl3Change(Sender: TObject);
    procedure cxGridResultViewColumn1PropertiesValidate(Sender: TObject;
      var DisplayValue: Variant; var ErrorText: TCaption; var Error: Boolean);
    procedure cxChartFontSizePropertiesEditValueChanged(Sender: TObject);
    procedure Report2Execute(Sender: TObject);
    //procedure Report3Execute(Sender: TObject);
    procedure cxGridResultViewEditKeyDown(Sender: TcxCustomGridTableView;
      AItem: TcxCustomGridTableItem; AEdit: TcxCustomEdit; var Key: Word;
      Shift: TShiftState);
    procedure cxGridResultViewKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure cxGridResultViewCustomDrawCell(Sender: TcxCustomGridTableView;
      ACanvas: TcxCanvas; AViewInfo: TcxGridTableDataCellViewInfo;
      var ADone: Boolean);
    procedure cxMCIUnitNamePropertiesEditValueChanged(Sender: TObject);
    procedure cxUnitNamePropertiesEditValueChanged(Sender: TObject);
    procedure cxMeasureNamePropertiesEditValueChanged(Sender: TObject);
    procedure cxLaboratoryTypePropertiesEditValueChanged(Sender: TObject);
    procedure cxR2ContrFocusChanged(Sender: TObject);
    procedure cxRContrFocusChanged(Sender: TObject);
    procedure cxMCILaboratoryTypePropertiesEditValueChanged(Sender: TObject);
    procedure cxMCIMeasureNamePropertiesEditValueChanged(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure FormActivate(Sender: TObject);
    procedure cxAnalyzeTypePropertiesChange(Sender: TObject);
  private
    { Private declarations }
  public
    var A_koef : array of array of Currency;
    var B_koef : array of array of Currency;
    var D_koef : array of array of Currency;
    var Q_koef : array of Currency;
    var C_koef_middle : array of array of Currency;
    var d_koef_middle : array of array of Currency;
    var Student_koef : array of Currency;
    var Kohren_koef : array of array of Currency;
    var SCO_koef : array of array of Currency;
    var A_Me_koef: array of Currency;
    //var m_iCntRows: integer;
    var X_avg_middle, X_avg_UCL, X_avg_LCL, X_AVG_AVG: Currency;
    var R_middle, R_UCL, R_LCL: Currency;
    var S_middle, S_UCL, S_LCL, S_MAX_2, S_SUM_2, S_AVG_2: Currency;
    var Me_middle, Me_UCL, Me_LCL: Currency;
    var CH, CL, BH, BL, AH, AL, Mid: Currency;
    var arrCriteries: array of TList;

    var gx, gy: integer;
    var metka: string;
    var result : TCalcResult;
    dataMeasureName: TStrings;
  end;

var
  ShuhartForm: TShuhartForm;

implementation

{$R *.dfm}

constructor TCurrencyValue.Create(const value : Currency);
begin
  // Сохранение переданных параметров
  self.value := value;
end;

constructor TCriteriaInfo.Create(const ind_beg : integer; const ind_end : integer);
begin
  // Сохранение переданных параметров
  self.index_beg := ind_beg;
  self.index_end := ind_end;
end;

function compare(Item1 : Pointer; Item2 : Pointer) : Integer;
var
  val1, val2 : TCurrencyValue;
begin
  val1 := TCurrencyValue(Item1);
  val2 := TCurrencyValue(Item2);

  // Теперь сравнение строк
  if      val1.val > val2.val
  then Result := 1
  else if val1.val = val2.val
  then Result := 0
  else Result := -1;
end;

procedure TShuhartForm.DrawRCard();
var
  i, j, iCnt: integer;
  bAllNull: boolean;
  y: double;
begin
  R_Chart_Data.Clear;
  R_Chart_Mid.Clear;
  R_Chart_UCL.Clear;
  R_Chart_LCL.Clear;

  {if (R_LCL <= 0) then Series4.Visible := false
  else Series4.Visible := true;}


  if (cxCardType.ItemIndex = 0) then
    iCnt := cxMeasureCnt.ItemIndex + 2
  else
    iCnt := 1;
  y := 1;
  for i := 0 to cxGridView.DataController.RecordCount - 1 do begin
    bAllNull := true;
    for j := 0 to iCnt - 1 do begin
      if (cxGridView.DataController.Values[i, 1 + j] <> Null) then begin
        bAllNull := false;
        break;
      end
    end;
    if (not bAllNull) AND (cxGridView.DataController.Values[i, R_col_index] <> Null) then  begin
      R_Chart_Data.AddXY(y, cxGridView.DataController.Values[i, R_col_index], '', clGreen);
    end;
    if (not bAllNull) then  begin
      R_Chart_Mid.AddXY(y, R_middle, '', clBlack);
      R_Chart_UCL.AddXY(y, R_UCL, '', clRed);
      R_Chart_LCL.AddXY(y, R_LCL, '', clRed);
      y := y + 1;
    end;
  end;

  R_Chart.Title.Text.Clear;

  R_Chart.Title.Text.Add('Карта размахов');
  R_Chart.Title.Text.Add('(CL = ' + CurrToStrF(R_middle, ffFixed, cxDigitCnt.ItemIndex + 1)
    + '; UCL = ' + CurrToStrF(R_UCL, ffFixed, cxDigitCnt.ItemIndex + 1)
    + '; LCL = ' + CurrToStrF(R_LCL, ffFixed, cxDigitCnt.ItemIndex + 1) + ')');

end;

procedure TShuhartForm.DrawXCard();
var
  i, j, iCnt: integer;
  bAllNull: boolean;
  y: double;
begin
  X_Chart_Data.Clear;
  X_Chart_Mid.Clear;
  X_Chart_UCL.Clear;
  X_Chart_LCL.Clear;

  X_Chart_CH.Clear;
  X_Chart_CL.Clear;
  X_Chart_BH.Clear;
  X_Chart_BL.Clear;

  if (cxAnalyzeType.ItemIndex = 0) then begin
    X_Chart_CH.Visible := true;
    X_Chart_CL.Visible := true;
    X_Chart_BH.Visible := true;
    X_Chart_BL.Visible := true;
  end
  else begin
    X_Chart_CH.Visible := false;
    X_Chart_CL.Visible := false;
    X_Chart_BH.Visible := false;
    X_Chart_BL.Visible := false;
  end;
  {if (R_LCL <= 0) then Series4.Visible := false
  else Series4.Visible := true;}


  if (cxCardType.ItemIndex = 0) then
    iCnt := cxMeasureCnt.ItemIndex + 2
  else
    iCnt := 1;

  y := 1;
  for i := 0 to cxGridView.DataController.RecordCount - 1 do begin
    bAllNull := true;
    for j := 0 to iCnt - 1 do begin
      if (cxGridView.DataController.Values[i, 1 + j] <> Null) then begin
        bAllNull := false;
        break;
      end
    end;
    if (not bAllNull) then  begin
      if (cxCardType.ItemIndex = 0) then
        X_Chart_Data.AddXY(y, cxGridView.DataController.Values[i, X_avg_col_index], '', clGreen)
      else
        X_Chart_Data.AddXY(y, cxGridView.DataController.Values[i, 1], '', clGreen);
      X_Chart_Mid.AddXY(y, X_avg_middle, '', clBlack);
      X_Chart_UCL.AddXY(y, X_avg_UCL, '', clRed);
      X_Chart_LCL.AddXY(y, X_avg_LCL, '', clRed);

      if (cxAnalyzeType.ItemIndex = 0) then begin
        X_Chart_CH.AddXY(y, CH, '', clBlue);
        X_Chart_CL.AddXY(y, CL, '', clBlue);
        X_Chart_BH.AddXY(y, BH, '', clBlue);
        X_Chart_BL.AddXY(y, BL, '', clBlue);
      end;

      y := y + 1;
    end;
  end;

  X_Chart.Title.Text.Clear;

  X_Chart.Title.Text.Add('Карта средних значений');
  X_Chart.Title.Text.Add('(CL = ' + CurrToStrF(X_avg_middle, ffFixed, cxDigitCnt.ItemIndex + 1)
    + '; UCL = ' + CurrToStrF(X_avg_UCL, ffFixed, cxDigitCnt.ItemIndex + 1)
    + '; LCL = ' + CurrToStrF(X_avg_LCL, ffFixed, cxDigitCnt.ItemIndex + 1) + ')');
end;

procedure TShuhartForm.DrawConvChart();
var SKO_Prec_Min, SKO_Prec_Max, SKO_Prec_Mid: Currency;
  iCnt, i, j, y: integer;
  bAllNull: boolean;
  arrXavg: array of Currency;
begin

  iCnt := cxMeasureCnt.ItemIndex + 2;

  Conv_Chart.Clear;
  Sigma_Chart_Min.Clear;
  Sigma_Chart_Mid.Clear;
  Sigma_Chart_Max.Clear;

  y := 0;
  SetLength(arrXavg, cxGridView.DataController.RecordCount);
  if ((iCnt - 2) < Length(SCO_koef[0])) then begin
    SKO_Prec_Min := SCO_koef[0, iCnt - 2] * result.SKO_Prec_val;
    SKO_Prec_Mid := SCO_koef[1, iCnt - 2] * result.SKO_Prec_val;
    SKO_Prec_Max := SCO_koef[2, iCnt - 2] * result.SKO_Prec_val;

    for i := 0 to cxGridView.DataController.RecordCount - 1 do begin
      bAllNull := true;
      for j := 0 to iCnt - 1 do begin
        if (cxGridView.DataController.Values[i, 1 + j] <> Null) then begin
          bAllNull := false;
          break;
        end
      end;
      if (not bAllNull) then  begin
        arrXavg[y] := cxGridView.DataController.Values[i, X_avg_col_index];
        y := y + 1;
      end;
    end;

    SetLength(arrXavg, y);

    if (y >= 2) then begin
      for i := 0 to Length(arrXavg) - 2 do begin
        arrXavg[i] := arrXavg[i] - arrXavg [i + 1];
        if (arrXavg[i] < 0) then arrXavg[i] := - arrXavg[i];
        Conv_Chart.AddXY(i + 1, arrXavg[i], '', clGreen);
        Sigma_Chart_Min.AddXY(i + 1, SKO_Prec_Min, '', clBlack);
        Sigma_Chart_Mid.AddXY(i + 1, SKO_Prec_Mid, '', clBlue);
        Sigma_Chart_Max.AddXY(i + 1, SKO_Prec_Max, '', clRed);
      end;
    end;

  end;

end;

procedure TShuhartForm.DrawSCard();
var
  i, j, iCnt: integer;
  bAllNull: boolean;
  y: double;
begin
  S_Chart_Data.Clear;
  S_Chart_Mid.Clear;
  S_Chart_UCL.Clear;
  S_Chart_LCL.Clear;

  if (cxAnalyzeType.ItemIndex = 1) then begin
    S_Chart_CH.Visible := true;
    S_Chart_CL.Visible := true;
    S_Chart_BH.Visible := true;
    S_Chart_BL.Visible := true;
  end
  else begin
    S_Chart_CH.Visible := false;
    S_Chart_CL.Visible := false;
    S_Chart_BH.Visible := false;
    S_Chart_BL.Visible := false;
  end;
  {if (R_LCL <= 0) then Series4.Visible := false
  else Series4.Visible := true;}


  iCnt := cxMeasureCnt.ItemIndex + 2;
  y := 1;
  for i := 0 to cxGridView.DataController.RecordCount - 1 do begin
    bAllNull := true;
    for j := 0 to iCnt - 1 do begin
      if (cxGridView.DataController.Values[i, 1 + j] <> Null) then begin
        bAllNull := false;
        break;
      end
    end;
    if (not bAllNull) then  begin
      S_Chart_Data.AddXY(y, cxGridView.DataController.Values[i, S_col_index], '', clGreen);
      S_Chart_Mid.AddXY(y, S_middle, '', clBlack);
      S_Chart_UCL.AddXY(y, S_UCL, '', clRed);
      S_Chart_LCL.AddXY(y, S_LCL, '', clRed);

      if (cxAnalyzeType.ItemIndex = 1) then begin
        S_Chart_CH.AddXY(y, CH, '', clBlue);
        S_Chart_CL.AddXY(y, CL, '', clBlue);
        S_Chart_BH.AddXY(y, BH, '', clBlue);
        S_Chart_BL.AddXY(y, BL, '', clBlue);
      end;

      y := y + 1;
    end;
  end;

  S_Chart.Title.Text.Clear;

  S_Chart.Title.Text.Add('Карта выборочных СКО');
  S_Chart.Title.Text.Add('(CL = ' + CurrToStrF(S_middle, ffFixed, cxDigitCnt.ItemIndex + 1)
    + '; UCL = ' + CurrToStrF(S_UCL, ffFixed, cxDigitCnt.ItemIndex + 1)
    + '; LCL = ' + CurrToStrF(S_LCL, ffFixed, cxDigitCnt.ItemIndex + 1) + ')');
end;

procedure TShuhartForm.DrawMeCard();
var
  i, j, iCnt: integer;
  bAllNull: boolean;
  y: double;
begin
  Me_Chart_Data.Clear;
  Me_Chart_Mid.Clear;
  Me_Chart_UCL.Clear;
  Me_Chart_LCL.Clear;

  if (cxAnalyzeType.ItemIndex = 2) then begin
    Me_Chart_CH.Visible := true;
    Me_Chart_CL.Visible := true;
    Me_Chart_BH.Visible := true;
    Me_Chart_BL.Visible := true;
  end
  else begin
    Me_Chart_CH.Visible := false;
    Me_Chart_CL.Visible := false;
    Me_Chart_BH.Visible := false;
    Me_Chart_BL.Visible := false;
  end;
  {if (R_LCL <= 0) then Series4.Visible := false
  else Series4.Visible := true;}


  iCnt := cxMeasureCnt.ItemIndex + 2;
  y := 1;
  for i := 0 to cxGridView.DataController.RecordCount - 1 do begin
    bAllNull := true;
    for j := 0 to iCnt - 1 do begin
      if (cxGridView.DataController.Values[i, 1 + j] <> Null) then begin
        bAllNull := false;
        break;
      end
    end;
    if (not bAllNull) then  begin
      Me_Chart_Data.AddXY(y, cxGridView.DataController.Values[i, Me_col_index], '', clGreen);
      Me_Chart_Mid.AddXY(y, Me_middle, '', clBlack);
      Me_Chart_UCL.AddXY(y, Me_UCL, '', clRed);
      Me_Chart_LCL.AddXY(y, Me_LCL, '', clRed);

      if (cxAnalyzeType.ItemIndex = 2) then begin
        Me_Chart_CH.AddXY(y, CH, '', clBlue);
        Me_Chart_CL.AddXY(y, CL, '', clBlue);
        Me_Chart_BH.AddXY(y, BH, '', clBlue);
        Me_Chart_BL.AddXY(y, BL, '', clBlue);
      end;

      y := y + 1;
    end;
  end;

  Me_Chart.Title.Text.Clear;

  Me_Chart.Title.Text.Add('Карта медиан');
  Me_Chart.Title.Text.Add('(CL = ' + CurrToStrF(Me_middle, ffFixed, cxDigitCnt.ItemIndex + 1)
    + '; UCL = ' + CurrToStrF(Me_UCL, ffFixed, cxDigitCnt.ItemIndex + 1)
    + '; LCL = ' + CurrToStrF(Me_LCL, ffFixed, cxDigitCnt.ItemIndex + 1) + ')');
end;

function TShuhartForm.VerifyCriteria_1(var arr_val : array of Currency; index: integer): boolean;
begin
  // 1. Одна точка вне зоны А
  if ((arr_val[index] > AH{X_avg_UCL}) or (arr_val[index] < AL{X_avg_LCL})) then result := true
  else result := false;
end;

function TShuhartForm.VerifyCriteria_2(var arr_val : array of Currency; index: integer): boolean;
var i, iCntElements: integer;
begin
  result := true;
  iCntElements := 9;
  // 2. Девять точек подряд в зоне C
  if (index + iCntElements - 1 >= Length(arr_val)) then begin
    result := false;
    exit;
  end;

  for i := index to index + iCntElements - 1 do begin
    if (arr_val[i] > CH) or (arr_val[i] < CL) then begin
      result := false;
      break;
    end;
  end;

  // или по одну сторону от центральной линии
  if (result = false) then begin
    result := true;

    for i := index to index + iCntElements - 1 do begin
      if (arr_val[i] > Mid{X_avg_middle}) then begin
        result := false;
        break;
      end;
    end;

    if (result = false) then begin
      result := true;

      for i := index to index + iCntElements - 1 do begin
        if (arr_val[i] < Mid{X_avg_middle}) then begin
          result := false;
          exit;
        end;
      end;
    end;

  end;
end;

function TShuhartForm.VerifyCriteria_3(var arr_val : array of Currency; index: integer): boolean;
var i, iCntElements: integer;
  prev_val : Currency;
begin
  result := true;
  iCntElements := 6;
  // 3. Шесть возрастающих точек подряд
  if (index + iCntElements - 1 >= Length(arr_val)) then begin
    result := false;
    exit;
  end;

  prev_val := arr_val[index];
  for i := index + 1 to index + iCntElements - 1 do begin
    if (arr_val[i] < prev_val) then begin
      result := false;
      break;
    end
    else prev_val := arr_val[i];
  end;

  if (result = false) then begin
    result := true;
    // или убывающих точек подряд
    prev_val := arr_val[index];
    for i := index + 1 to index + iCntElements - 1 do begin
      if (arr_val[i] > prev_val) then begin
        result := false;
        exit;
      end
      else prev_val := arr_val[i];
    end;
  end;
end;

function TShuhartForm.VerifyCriteria_4(var arr_val : array of Currency; index: integer): boolean;
var i, iCntElements: integer;
  prev_val : Currency;
  bVal: boolean;
begin
  result := true;
  iCntElements := 14;
  // 4. Четырнадцать попеременно возрастающих и убывающих точек
  if (index + iCntElements - 1 >= Length(arr_val)) then begin
    result := false;
    exit;
  end;

  if (arr_val[index] > arr_val[index + 1]) then bVal := true
  else bVal := false;

  prev_val := arr_val[index];
  for i := index + 1 to index + iCntElements - 1 do begin
    if not((arr_val[i] > prev_val) and not (bVal) or (arr_val[i] <= prev_val) and (bVal)) then begin
      result := false;
      exit;
    end
    else begin
      prev_val := arr_val[i];
      bVal := not bVal;
    end;
  end;
end;

function TShuhartForm.VerifyCriteria_5(var arr_val : array of Currency; index: integer): boolean;
var i, iCnt, iCnt2, iCntElements: integer;
begin
  result := true;
  iCntElements := 3;
  // 5. Две из трех последовательных точек находятся с одной стороны от центральной линии в зоне A или выше
  if (index + iCntElements - 1 >= Length(arr_val)) then begin
    result := false;
    exit;
  end;

  iCnt := 0;
  iCnt2 := 0;
  if (arr_val[index] >= Mid{X_avg_middle}) then begin
    for i := index to index + iCntElements - 1 do begin
      if (arr_val[i] > BH) then iCnt := iCnt + 1;
      if (arr_val[i] >= Mid{X_avg_middle}) then iCnt2 := iCnt2 + 1;
    end;
    if (iCnt >= 2) AND (iCnt2 = 3) then begin
      result := true;
      exit;
    end;
  end;

  iCnt := 0;
  if (arr_val[index] <= Mid{X_avg_middle}) then begin
    for i := index to index + iCntElements - 1 do begin
      if (arr_val[i] < BL) then iCnt := iCnt + 1;
      if (arr_val[i] <= Mid{X_avg_middle}) then iCnt2 := iCnt2 + 1;
    end;
    if (iCnt >= 2) AND (iCnt2 = 3) then begin
      result := true;
      exit;
    end;
  end;

  result := false;
end;

function TShuhartForm.VerifyCriteria_6(var arr_val : array of Currency; index: integer): boolean;
var i, iCnt, iCnt2, iCntElements: integer;
begin
  result := true;
  iCntElements := 5;
  // 6. Четыре из пяти последовательных точек находятся с одной стороны от центральной линии в зоне B или выше
  if (index + iCntElements - 1 >= Length(arr_val)) then begin
    result := false;
    exit;
  end;

  iCnt := 0;
  iCnt2 := 0;
  if (arr_val[index] >= Mid{X_avg_middle}) then begin
    for i := index to index + iCntElements - 1 do begin
      if (arr_val[i] > CH) then iCnt := iCnt + 1;
      if (arr_val[i] >= Mid{X_avg_middle}) then iCnt2 := iCnt2 + 1;
    end;
    if (iCnt >= 4) AND (iCnt2 = 5) then begin
      result := true;
      exit;
    end;
  end;

  iCnt := 0;
  iCnt2 := 0;
  if (arr_val[index] <= Mid{X_avg_middle}) then begin
    for i := index to index + iCntElements - 1 do begin
      if (arr_val[i] < CL) then iCnt := iCnt + 1;
      if (arr_val[i] <= Mid{X_avg_middle}) then iCnt2 := iCnt2 + 1;
    end;
    if (iCnt >= 4) AND (iCnt2 = 5) then begin
      result := true;
      exit;
    end;
  end;

  result := false;
end;

function TShuhartForm.VerifyCriteria_7(var arr_val : array of Currency; index: integer): boolean;
var i, iCntElements: integer;
  bUpperMiddle, bLowerMiddle: boolean;
begin
  result := true;
  iCntElements := 15;
  // 7. Пятнадцать последовательных точек в зоне C выше и ниже центральной линии
  if (index + iCntElements - 1 >= Length(arr_val)) then begin
    result := false;
    exit;
  end;

  bUpperMiddle := false;
  bLowerMiddle := false;
  for i := index to index + iCntElements - 1 do begin
    if (arr_val[i] > CH) OR (arr_val[i] < CL) then begin
      result := false;
      exit;
    end;

    if (arr_val[i] > Mid{X_avg_middle}) then bUpperMiddle := true;
    if (arr_val[i] < Mid{X_avg_middle}) then bLowerMiddle := true;

  end;

  if (not bUpperMiddle) or (not bLowerMiddle) then begin
    result := false;
    exit;
  end;

end;

function TShuhartForm.VerifyCriteria_8(var arr_val : array of Currency; index: integer): boolean;
var i, iCntElements: integer;
  bUpperMiddle, bLowerMiddle: boolean;
begin
  result := true;
  iCntElements := 8;
  // 8. Восемь последовательных точек по обеим сторонам центральной линии и ни одной в зоне C
  if (index + iCntElements - 1 >= Length(arr_val)) then begin
    result := false;
    exit;
  end;

  bUpperMiddle := false;
  bLowerMiddle := false;
  for i := index to index + iCntElements - 1 do begin
    if (arr_val[i] <= CH) AND (arr_val[i] >= CL) then begin
      result := false;
      exit;
    end;

    if (arr_val[i] > Mid{X_avg_middle}) then bUpperMiddle := true;
    if (arr_val[i] < Mid{X_avg_middle}) then bLowerMiddle := true;

  end;

  if (not bUpperMiddle) or (not bLowerMiddle) then begin
    result := false;
    exit;
  end;
end;

procedure TShuhartForm.X_ChartAfterDraw(Sender: TObject);
var i: integer;
  h, w: integer;
  rect: TRect;
  x_max, y_mid, y_CH, y_CL, y_BH, y_BL, y_AH, y_AL: integer;
begin
 with (Sender as TChart).Canvas do begin

  if ((Sender as TChart).Name = 'X_Chart') AND (X_Chart_Mid.Count > 0) AND (cxAnalyzeType.ItemIndex = 0)
    OR ((Sender as TChart).Name = 'S_Chart') AND (S_Chart_Mid.Count > 0) AND (cxAnalyzeType.ItemIndex = 1)
    OR ((Sender as TChart).Name = 'Me_Chart') AND (Me_Chart_Mid.Count > 0) AND (cxAnalyzeType.ItemIndex = 2)
  then begin
    //
    if (cxAnalyzeType.ItemIndex = 0) then begin
      x_max := X_Chart_Mid.CalcXPos(X_Chart_Mid.Count - 1);

      y_mid := X_Chart_Mid.CalcYPos(X_Chart_Mid.Count - 1);
      y_CH := X_Chart_CH.CalcYPos(X_Chart_CH.Count - 1);
      y_CL := X_Chart_CL.CalcYPos(X_Chart_CL.Count - 1);
      y_BH := X_Chart_BH.CalcYPos(X_Chart_BH.Count - 1);
      y_BL := X_Chart_BL.CalcYPos(X_Chart_BL.Count - 1);
      y_AH := X_Chart_UCL.CalcYPos(X_Chart_UCL.Count - 1);
      y_AL := X_Chart_LCL.CalcYPos(X_Chart_LCL.Count - 1);
    end;
    if (cxAnalyzeType.ItemIndex = 1) then begin
      x_max := S_Chart_Mid.CalcXPos(S_Chart_Mid.Count - 1);

      y_mid := S_Chart_Mid.CalcYPos(S_Chart_Mid.Count - 1);
      y_CH := S_Chart_CH.CalcYPos(S_Chart_CH.Count - 1);
      y_CL := S_Chart_CL.CalcYPos(S_Chart_CL.Count - 1);
      y_BH := S_Chart_BH.CalcYPos(S_Chart_BH.Count - 1);
      y_BL := S_Chart_BL.CalcYPos(S_Chart_BL.Count - 1);
      y_AH := S_Chart_UCL.CalcYPos(S_Chart_UCL.Count - 1);
      y_AL := S_Chart_LCL.CalcYPos(S_Chart_LCL.Count - 1);
    end;
    if (cxAnalyzeType.ItemIndex = 2) then begin
      x_max := Me_Chart_Mid.CalcXPos(X_Chart_Mid.Count - 1);

      y_mid := Me_Chart_Mid.CalcYPos(Me_Chart_Mid.Count - 1);
      y_CH := Me_Chart_CH.CalcYPos(Me_Chart_CH.Count - 1);
      y_CL := Me_Chart_CL.CalcYPos(Me_Chart_CL.Count - 1);
      y_BH := Me_Chart_BH.CalcYPos(Me_Chart_BH.Count - 1);
      y_BL := Me_Chart_BL.CalcYPos(Me_Chart_BL.Count - 1);
      y_AH := Me_Chart_UCL.CalcYPos(Me_Chart_UCL.Count - 1);
      y_AL := Me_Chart_LCL.CalcYPos(Me_Chart_LCL.Count - 1);
    end;

    Font.Size:= 8; Font.Color:= clBlack;
    Brush.Color := RGB(255, 255, 225);//clRed;//clInfoBk;
    Pen.Style := psSolid;
    pen.Color := clBlack;
    pen.Width := 1;

    i := round((y_AH + y_BH) / 2);
    rect := TRect.Create(x_max - 8, i - 8, x_max + 8, i + 8);
    RoundRect(rect, 5, 5);
    TextOut(x_max - 5, i - 8, 'A');

    Font.Size:= 8; Font.Color:= clBlack;
    Brush.Color := RGB(255, 255, 225);//clRed;//clInfoBk;
    Pen.Style := psSolid;
    pen.Color := clBlack;
    pen.Width := 1;

    i := round((y_BH + y_CH) / 2);
    rect := TRect.Create(x_max - 8, i - 8, x_max + 8, i + 8);
    RoundRect(rect, 5, 5);
    TextOut(x_max - 5, i - 8, 'B');

    Font.Size:= 8; Font.Color:= clBlack;
    Brush.Color := RGB(255, 255, 225);//clRed;//clInfoBk;
    Pen.Style := psSolid;
    pen.Color := clBlack;
    pen.Width := 1;

    i := round((y_CH + y_mid) / 2);
    rect := TRect.Create(x_max - 8, i - 8, x_max + 8, i + 8);
    RoundRect(rect, 5, 5);
    TextOut(x_max - 5, i - 8, 'C');

    Font.Size:= 8; Font.Color:= clBlack;
    Brush.Color := RGB(255, 255, 225);//clRed;//clInfoBk;
    Pen.Style := psSolid;
    pen.Color := clBlack;
    pen.Width := 1;

    i := round((y_BL + y_AL) / 2);
    rect := TRect.Create(x_max - 8, i - 8, x_max + 8, i + 8);
    RoundRect(rect, 5, 5);
    TextOut(x_max - 5, i - 8, 'A');

    Font.Size:= 8; Font.Color:= clBlack;
    Brush.Color := RGB(255, 255, 225);//clRed;//clInfoBk;
    Pen.Style := psSolid;
    pen.Color := clBlack;
    pen.Width := 1;

    i := round((y_CL + y_BL) / 2);
    rect := TRect.Create(x_max - 8, i - 8, x_max + 8, i + 8);
    RoundRect(rect, 5, 5);
    TextOut(x_max - 5, i - 8, 'B');

    Font.Size:= 8; Font.Color:= clBlack;
    Brush.Color := RGB(255, 255, 225);//clRed;//clInfoBk;
    Pen.Style := psSolid;
    pen.Color := clBlack;
    pen.Width := 1;

    i := round((y_mid + y_CL) / 2);
    rect := TRect.Create(x_max - 8, i - 8, x_max + 8, i + 8);
    RoundRect(rect, 5, 5);
    TextOut(x_max - 5, i - 8, 'C');
  end;


  // прицел-
  if gx > 1 then begin
    Font.Size:= 8; Font.Color:= clBlack;
    Brush.Color := RGB(255, 255, 225);//clRed;//clInfoBk;
    Pen.Style := psSolid;
    pen.Color := clBlack;
    pen.Width := 1;

    h := TextHeight(metka);
    w := TextWidth(metka);
    rect := TRect.Create(gx+10 - 2, gy + font.Height-4, gx+10 + w + 2, gy + font.Height + h);
    RoundRect(rect, 10, 10);
    TextOut(gx+10,gy + font.Height-2, metka);
  end;
 end;
end;

procedure TShuhartForm.X_ChartMouseMove(Sender: TObject; Shift: TShiftState; X,
  Y: Integer);
var i, j: integer;
    tmpX,tmpY:Double;
    ch: TChart;
begin
  ch := (Sender as TChart);
  ch.Repaint;
  gx:= -1; gy:= -1;

  for i := 0 to ch.SeriesCount - 1 do begin
    if (ch.Series[i].GetCursorValueIndex <> -1) then begin
       gx:=x; gy:=y;

      //X_Chart.Series[0].GetCursorValues(tmpX, tmpY);
      tmpX := ch.Series[i].GetMarkValue(ch.Series[i].GetCursorValueIndex);
      metka := FloatToStr(tmpX);
      if ((Sender as TChart).Name <> 'SKO_Chart') then  begin
        if (i = 1) then metka := 'CL = ' + metka;
        if (i = 2) then metka := 'UCL = ' + metka;
        if (i = 3) then metka := 'LCL = ' + metka;
      end;
      //metka := ch.Series[i].ValueMarkText[ch.Series[i].GetCursorValueIndex];
      break;
    end;
  end;
end;

procedure TShuhartForm.CalcCriteries();
var
  i, j, iCnt, iCntMeasure, ind_beg, ind_end: integer;
  bAllNull: boolean;
  val: Currency;
  point_index: integer;
  arr_val: array of Currency;
  str: string;
  PCI: double;
begin
  // рассчет критериев
  for i := 0 to 7 do arrCriteries[i].Clear;

  if (cxCardType.ItemIndex = 0) then
    iCnt := cxMeasureCnt.ItemIndex + 2
  else
    iCnt := 1;

  point_index := 0;
  SetLength(arr_val, cxGridView.DataController.RecordCount);
  for i := 0 to cxGridView.DataController.RecordCount - 1 do begin
    bAllNull := true;
    for j := 0 to iCnt - 1 do begin
      if (cxGridView.DataController.Values[i, 1 + j] <> Null) then begin
        bAllNull := false;
        break;
      end
    end;
    if (not bAllNull) then  begin
      if (cxAnalyzeType.ItemIndex = 0) then begin
        if (cxCardType.ItemIndex = 0) then
          val := cxGridView.DataController.Values[i, X_avg_col_index]
        else
          val := cxGridView.DataController.Values[i, 1];
      end;
      if (cxAnalyzeType.ItemIndex = 1) then begin
        val := cxGridView.DataController.Values[i, S_col_index]
      end;
      if (cxAnalyzeType.ItemIndex = 2) then begin
        val := cxGridView.DataController.Values[i, Me_col_index]
      end;
      arr_val[point_index] := val;
      point_index := point_index + 1;
    end;
  end;
  if (point_index > 0) then
    SetLength(arr_val, point_index)
  else arr_val := nil;

  if (cxAnalyzeType.ItemIndex = 0) then begin
    AH := X_avg_UCL;
    AL := X_avg_LCL;
    Mid := X_avg_middle;
    cxGridCriteriaViewColumn4.Caption := 'Карта средних';
  end;
  if (cxAnalyzeType.ItemIndex = 1) then begin
    AH := S_UCL;
    AL := S_LCL;
    Mid := S_middle;
    cxGridCriteriaViewColumn4.Caption := 'Карта выборочных СКО';
  end;
  if (cxAnalyzeType.ItemIndex = 2) then begin
    AH := Me_UCL;
    AL := Me_LCL;
    Mid := Me_middle;
    cxGridCriteriaViewColumn4.Caption := 'Карта медиан';
  end;

  for i := 0 to Length(arr_val) - 1 do begin
    if VerifyCriteria_1(arr_val, i) then begin
      arrCriteries[0].Add(TCriteriaInfo.Create(i, i));
    end;

    if VerifyCriteria_2(arr_val, i) then begin
      arrCriteries[1].Add(TCriteriaInfo.Create(i, i + 9 - 1));
    end;

    if VerifyCriteria_3(arr_val, i) then begin
      arrCriteries[2].Add(TCriteriaInfo.Create(i, i + 6 - 1));
    end;

    if VerifyCriteria_4(arr_val, i) then begin
      arrCriteries[3].Add(TCriteriaInfo.Create(i, i + 14 - 1));
    end;

    if VerifyCriteria_5(arr_val, i) then begin
      arrCriteries[4].Add(TCriteriaInfo.Create(i, i + 3 - 1));
    end;

    if VerifyCriteria_6(arr_val, i) then begin
      arrCriteries[5].Add(TCriteriaInfo.Create(i, i + 5 - 1));
    end;

    if VerifyCriteria_7(arr_val, i) then begin
      arrCriteries[6].Add(TCriteriaInfo.Create(i, i + 15 - 1));
    end;

    if VerifyCriteria_8(arr_val, i) then begin
      arrCriteries[7].Add(TCriteriaInfo.Create(i, i + 8 - 1));
    end;
  end;

  result.bWarningShuhart := false;
  result.bCriticalShuhart := false;
  for i := 0 to 8 - 1 do begin
    if (arrCriteries[i].Count = 0) then
      cxGridCriteriaView.DataController.Values[i, 3] := 'Критерий не обнаружен'
    else begin
      if (cxGridCriteriaView.DataController.Values[i, 2] = 'Предупреждающий') then
        result.bWarningShuhart := true
      else
        result.bCriticalShuhart := true;
      str := '';
      for j := 0 to arrCriteries[i].Count - 1 do begin
        ind_beg := TCriteriaInfo(arrCriteries[i].Items[j]).ind_beg;
        ind_end := TCriteriaInfo(arrCriteries[i].Items[j]).ind_end;
        if (str <> '') then str := str + ', ';
        if (ind_beg <> ind_end) then
          str := str + IntToStr(ind_beg + 1) + '-' + IntToStr(ind_end + 1)
        else
          str := str + IntToStr(ind_beg + 1);
      end;

      cxGridCriteriaView.DataController.Values[i, 3] := str;
    end;

  end;

  if (cxCardType.ItemIndex = 0) then
    iCntMeasure := cxMeasureCnt.ItemIndex + 2
  else
    iCntMeasure := 2;

  result.PCI_val := 0;
  result.bPCI_val := false;

  if (R_middle <> 0) then begin
    PCI := (X_avg_UCL - X_avg_LCL) / (6.0 * R_middle / d_koef_middle[0, iCntMeasure - 2]);
    PCI_label.Caption := 'Индекс возможностей процесса (PCI) = ' + FloatToStrF(PCI, ffFixed, 6, 6);
    result.PCI_val := PCI;
    result.bPCI_val := true;
  end
  else
    PCI_label.Caption := 'Индекс возможностей процесса (PCI) = нет значения';

end;

procedure TShuhartForm.CalcResultMCI();
var i, index, iOKCnt: integer;
  bRes: boolean;
  str: string;
begin
  cxGridResultView.BeginUpdate();

  SetLength(result.arrMCILabInfo, cxLaboratoryCnt.ItemIndex);
  for i := 0 to Length(result.arrMCILabInfo) - 1 do begin
    result.arrMCILabInfo[i].strLaboratoryName := cxGridLaboratoryView.DataController.Values[i, 1];
    result.arrMCILabInfo[i].bResult := true;
  end;

  if (Length(result.arrMCILabInfo) > 0) then begin
    index := 0;
    iOKCnt := cxOKCnt.ItemIndex + 1;
    bRes := true;
    for i := 0 to cxGridResultView.DataController.RecordCount - 1 do begin
      RecalcResultRowData(i);

      str := cxGridResultView.DataController.Values[i, 9];
      if (cxGridResultView.DataController.Values[i, 9] <> 'Погрешность удовл.') then
        bRes := false;

      if (((i + 1) mod iOKCnt) = 0) then begin
        result.arrMCILabInfo[index].bResult := bRes;
        index := index + 1;
        bRes := true;
      end
    end;
  end;

  cxGridResultView.EndUpdate();
end;

procedure TShuhartForm.CalcDataKoef();
var
  i, j, iCnt, iCntMeasure, iCntRows: integer;
  bAllNull, bShuhart: boolean;
  X_avg_sum, R_sum, Me_sum, S_sum, tmp: Currency;
begin
  if (cxCardType.ItemIndex = 0) then begin
    iCntRows := 0;
    iCnt := cxMeasureCnt.ItemIndex + 2;

    X_avg_sum := 0;
    R_sum := 0;
    Me_sum := 0;
    S_sum := 0;
    S_MAX_2 := 0;
    S_SUM_2 := 0;
    S_AVG_2 := 0;
    // считаем кол-во ненулевых строк
    for i := 0 to cxGridView.DataController.RecordCount - 1 do begin
      bAllNull := true;
      for j := 0 to iCnt - 1 do begin
        if (cxGridView.DataController.Values[i, 1 + j] <> Null) then begin
          bAllNull := false;
          break;
        end
      end;
      if (not bAllNull) then  begin
        iCntRows := iCntRows + 1;
        X_avg_sum := X_avg_sum + cxGridView.DataController.Values[i, X_avg_col_index];
        S_sum := S_sum + cxGridView.DataController.Values[i, S_col_index];
        Me_sum := Me_sum + cxGridView.DataController.Values[i, Me_col_index];
        R_sum := R_sum + cxGridView.DataController.Values[i, R_col_index];

        tmp := cxGridView.DataController.Values[i, S_col_index];
        if (S_MAX_2 < tmp * tmp) then S_MAX_2 := tmp * tmp;
        S_SUM_2 := S_SUM_2 + tmp * tmp;
      end;
    end;

    iCntMeasure := cxMeasureCnt.ItemIndex + 2;

    if (iCntRows = 0) OR (iCntMeasure < 2) then begin
      X_avg_middle := 0;
      X_AVG_AVG := 0;
      R_middle := 0;
      S_middle := 0;
      Me_middle := 0;

      S_MAX_2 := 0;
      S_SUM_2 := 0;
      S_AVG_2 := 0;
      X_AVG_AVG := 0;

      X_avg_UCL := 0;
      R_UCL := 0;
      S_UCL := 0;
      Me_UCL := 0;

      X_avg_LCL := 0;
      R_LCL := 0;
      S_LCL := 0;
      Me_LCL := 0;

      CH := 0;
      CL := 0;

      BH := 0;
      BL := 0;

      cxGridCriteriaView.DataController.Values[0, 3] := '';
      cxGridCriteriaView.DataController.Values[1, 3] := '';
      cxGridCriteriaView.DataController.Values[2, 3] := '';
      cxGridCriteriaView.DataController.Values[3, 3] := '';
      cxGridCriteriaView.DataController.Values[4, 3] := '';
      cxGridCriteriaView.DataController.Values[5, 3] := '';
      cxGridCriteriaView.DataController.Values[6, 3] := '';
      cxGridCriteriaView.DataController.Values[7, 3] := '';

      PCI_label.Caption := 'Индекс возможностей процесса (PCI) = ';

      exit;
    end;


    // вычисляем коэффициенты для Xср
    if (not cxStandardValues.Checked) then begin
      // средняя линия
      X_avg_middle := X_avg_sum /  iCntRows;
      R_middle := R_sum /  iCntRows;
      S_middle := S_sum /  iCntRows;
      Me_middle := Me_sum /  iCntRows;

      // UCL
      X_avg_UCL := X_avg_middle + A_koef[2 - 1, iCntMeasure - 2] * R_middle;
      R_UCL := D_koef[4 - 1, iCntMeasure - 2] * R_middle;
      S_UCL := B_koef[2 - 1, iCntMeasure - 2] * S_middle;
      if ((iCntMeasure - 2) < Length(A_Me_koef)) then
        Me_UCL := Me_middle + A_Me_koef[iCntMeasure - 2] * R_middle
      else
        Me_UCL := 0;

      // LCL
      X_avg_LCL := X_avg_middle - A_koef[2 - 1, iCntMeasure - 2] * R_middle;
      R_LCL := D_koef[3 - 1, iCntMeasure - 2] * R_middle;
      S_LCL := B_koef[1 - 1, iCntMeasure - 2] * S_middle;
      if ((iCntMeasure - 2) < Length(A_Me_koef)) then
        Me_LCL := Me_middle - A_Me_koef[iCntMeasure - 2] * R_middle
      else
        Me_LCL := 0;
    end
    else begin
      // средняя линия
      X_avg_middle := cxX0.Value;
      R_middle := cxR0.Value;
      S_middle := cxS0.Value;
      Me_middle := Me_sum /  iCntRows;

      // UCL
      X_avg_UCL := X_avg_middle + A_koef[1 - 1, iCntMeasure - 2] * cxSigma0.Value;
      R_UCL := D_koef[2 - 1, iCntMeasure - 2] * cxSigma0.Value;
      S_UCL := B_koef[4 - 1, iCntMeasure - 2] * cxSigma0.Value;
      if ((iCntMeasure - 2) < Length(A_Me_koef)) then
        Me_UCL := Me_middle + A_Me_koef[iCntMeasure - 2] * R_middle
      else
        Me_UCL := 0;

      // LCL
      X_avg_LCL := X_avg_middle - A_koef[1 - 1, iCntMeasure - 2] * cxSigma0.Value;
      R_LCL := D_koef[1 - 1, iCntMeasure - 2] * cxSigma0.Value;
      S_LCL := B_koef[3 - 1, iCntMeasure - 2] * cxSigma0.Value;
      if ((iCntMeasure - 2) < Length(A_Me_koef)) then
        Me_LCL := Me_middle - A_Me_koef[iCntMeasure - 2] * R_middle
      else
        Me_UCL := 0;
    end;

    S_AVG_2 := S_SUM_2 / iCntRows;
    X_AVG_AVG := X_avg_sum /  iCntRows;
  end
  else begin
    X_avg_sum := 0;
    R_sum := 0;
    iCntRows := 0;
    // считаем кол-во ненулевых строк
    for i := 0 to cxGridView.DataController.RecordCount - 1 do begin
      if (cxGridView.DataController.Values[i, 1] <> Null) then begin
        iCntRows := iCntRows + 1;
        X_avg_sum := X_avg_sum + cxGridView.DataController.Values[i, 1];
        if (cxGridView.DataController.Values[i, R_col_index] <> Null) then
          R_sum := R_sum + cxGridView.DataController.Values[i, R_col_index];
      end;
    end;

    if (iCntRows <= 1) then begin
      X_avg_middle := 0;
      R_middle := 0;

      X_avg_UCL := 0;
      R_UCL := 0;

      X_avg_LCL := 0;
      R_LCL := 0;

      CH := 0;
      CL := 0;

      BH := 0;
      BL := 0;

      cxGridCriteriaView.DataController.Values[0, 3] := '';
      cxGridCriteriaView.DataController.Values[1, 3] := '';
      cxGridCriteriaView.DataController.Values[2, 3] := '';
      cxGridCriteriaView.DataController.Values[3, 3] := '';
      cxGridCriteriaView.DataController.Values[4, 3] := '';
      cxGridCriteriaView.DataController.Values[5, 3] := '';
      cxGridCriteriaView.DataController.Values[6, 3] := '';
      cxGridCriteriaView.DataController.Values[7, 3] := '';

      PCI_label.Caption := 'Индекс возможностей процесса (PCI) = ';

      exit;
    end;

    // вычисляем коэффициенты
    if (not cxStandardValues.Checked) then begin
      // средняя линия
      X_avg_middle := X_avg_sum /  iCntRows;
      R_middle := R_sum /  (iCntRows - 1);

      // UCL
      X_avg_UCL := X_avg_middle + 3 * d_koef_middle[1, 0] * R_middle;
      R_UCL := D_koef[4 - 1, 0] * R_middle;

      // LCL
      X_avg_LCL := X_avg_middle - 3 * d_koef_middle[1, 0] * R_middle;
      R_LCL := D_koef[3 - 1, 0] * R_middle;
    end
    else begin
      // средняя линия
      X_avg_middle := cxX0.Value;
      R_middle := cxR0.Value;

      // UCL
      X_avg_UCL := X_avg_middle + 3 * cxSigma0.Value;
      R_UCL := D_koef[2 - 1, 0] * cxSigma0.Value;

      // LCL
      X_avg_LCL := X_avg_middle - 3 * cxSigma0.Value;
      R_LCL := D_koef[1 - 1, 0] * cxSigma0.Value;
    end;
  end;


  // зоны A, B, C
  if (cxAnalyzeType.ItemIndex = 0) then begin
    CH := (X_avg_UCL - X_avg_middle) / 3 + X_avg_middle;
    CL := X_avg_middle - (X_avg_middle - X_avg_LCL) / 3;

    BH := 2 * (X_avg_UCL - X_avg_middle) / 3 + X_avg_middle;
    BL := X_avg_middle - 2 * (X_avg_middle - X_avg_LCL) / 3;
  end;
  if (cxAnalyzeType.ItemIndex = 1) then begin
    CH := (S_UCL - S_middle) / 3 + S_middle;
    CL := S_middle - (S_middle - S_LCL) / 3;

    BH := 2 * (S_UCL - S_middle) / 3 + S_middle;
    BL := S_middle - 2 * (S_middle - S_LCL) / 3;
  end;
  if (cxAnalyzeType.ItemIndex = 2) then begin
    CH := (Me_UCL - Me_middle) / 3 + Me_middle;
    CL := Me_middle - (Me_middle - Me_LCL) / 3;

    BH := 2 * (Me_UCL - Me_middle) / 3 + Me_middle;
    BL := Me_middle - 2 * (Me_middle - Me_LCL) / 3;
  end;

  CalcCriteries();
end;

procedure TShuhartForm.FillVerification();
begin

  cxC_val.Text := CurrToStr(result.C_val);
  cxCDelta_val.Text := CurrToStr(result.CDelta_val);

  cxRContr_val.Text := cxRContr.Text;
  cxR2Contr_val.Text := cxR2Contr.Text;

  if (result.bKohren_val) then begin
    cxKohren_val.Text := FloatToStrF(result.Kohren_val, ffFixed, 6, 6);
    cxKohren_val.Hint := '';
  end
  else begin
    cxKohren_val.Text := 'нет значения';
    cxKohren_val.Hint := 'значение рассчитать невозможно из-за ограничений в табличных коэффициентах или некорректности входных данных';
  end;

  cxSKO_povt_val.Text := FloatToStrF(result.SKO_povt_val, ffFixed, 6, 6);

  if (result.bSKO_Prec_val) then begin
    cxSKO_prec_val.Text := FloatToStrF(result.SKO_prec_val, ffFixed, 6, 6);
    cxSKO_prec_val.Hint := '';
  end
  else begin
    cxSKO_prec_val.Text := 'нет значения';
    cxSKO_prec_val.Hint := 'значение рассчитать невозможно из-за ограничений в табличных коэффициентах или некорректности входных данных';
  end;

  cxAvg_val.Text := FloatToStrF(result.Avg_val, ffFixed, 6, 6);

  if (result.bSKO_Povt_otn_val) then begin
    cxSKO_povt_otn_val.Text := FloatToStrF(result.SKO_povt_otn_val, ffFixed, 6, 6);
    cxSKO_povt_otn_val.Hint := '';
  end
  else begin
    cxSKO_povt_otn_val.Text := 'нет значения';
    cxSKO_povt_otn_val.Hint := 'значение рассчитать невозможно из-за ограничений в табличных коэффициентах или некорректности входных данных';
  end;

  if (result.bSKO_Prec_otn_val) then begin
    cxSKO_prec_otn_val.Text := FloatToStrF(result.SKO_prec_otn_val, ffFixed, 6, 6);
    cxSKO_prec_otn_val.Hint := '';
  end
  else begin
    cxSKO_prec_otn_val.Text := 'нет значения';
    cxSKO_prec_otn_val.Hint := 'значение рассчитать невозможно из-за ограничений в табличных коэффициентах или некорректности входных данных';
  end;

  cxMove_val.Text := FloatToStrF(result.Move_val, ffFixed, 6, 6);

  if (result.bSKO_System_Pogr_val) then begin
    cxSKO_System_Pogr_val.Text := FloatToStrF(result.SKO_System_Pogr_val, ffFixed, 6, 6);
    cxSKO_System_Pogr_val.Hint := '';
  end
  else begin
    cxSKO_System_Pogr_val.Text := 'нет значения';
    cxSKO_System_Pogr_val.Hint := 'значение рассчитать невозможно из-за ограничений в табличных коэффициентах или некорректности входных данных';
  end;

  if (result.bStudent_val) then begin
    cxStudent_val.Text := FloatToStrF(result.Student_val, ffFixed, 6, 6);
    cxStudent_val.Hint := '';
  end
  else begin
    cxStudent_val.Text := 'нет значения';
    cxStudent_val.Hint := 'значение рассчитать невозможно из-за ограничений в табличных коэффициентах или некорректности входных данных';
  end;

  cxSystem_Pogr_val.Text :=  FloatToStrF(result.System_Pogr_val, ffFixed, 6, 6);

  if (result.bVerify_val) then begin
    cxVerify_val.Text := FloatToStrF(result.Verify_val, ffFixed, 6, 6);
    cxVerify_val.Hint := '';
  end
  else begin
    cxVerify_val.Text := 'нет значения';
    cxVerify_val.Hint := 'значение рассчитать невозможно из-за ограничений в табличных коэффициентах или некорректности входных данных';
  end;

  cxAnalysPogr_val.Text := FloatToStrF(result.AnalysPogr_val, ffFixed, 6, 6);

  if (result.bMethodPogr_val) then begin
    cxMethodPogr_val.Text := FloatToStrF(result.MethodPogr_val, ffFixed, 6, 6);
    cxMethodPogr_val.Hint := '';
  end
  else begin
    cxMethodPogr_val.Text := 'нет значения';
    cxMethodPogr_val.Hint := 'значение рассчитать невозможно из-за ограничений в табличных коэффициентах или некорректности входных данных';
  end;

  if (result.bLimit_Povt_val) then begin
    cxLimit_Povt_val.Text := IntToStr(result.Limit_Povt_val);
    cxLimit_Povt_val.Hint := '';
  end
  else begin
    cxLimit_Povt_val.Text := 'нет значения';
    cxLimit_Povt_val.Hint := 'значение рассчитать невозможно из-за ограничений в табличных коэффициентах или некорректности входных данных';
  end;

  if (result.bLimit_Prec_val) then  begin
    cxLimit_Prec_val.Text := IntToStr(result.Limit_Prec_val);
    cxLimit_Prec_val.Hint := '';
  end
  else begin
    cxLimit_Prec_val.Text := 'нет значения';
    cxLimit_Prec_val.Hint := 'значение рассчитать невозможно из-за ограничений в табличных коэффициентах или некорректности входных данных';
  end;

  if (result.bPrecision_val) then begin
    cxPrecision_val.Text := IntToStr(result.Precision_val);
    cxPrecision_val.Hint := '';
  end
  else begin
    cxPrecision_val.Text := 'нет значения';
    cxPrecision_val.Hint := 'значение рассчитать невозможно из-за ограничений в табличных коэффициентах или некорректности входных данных';
  end;

  if (result.bKohren_tab_val) then
    lKohren_tab.Caption := '(Gтаб=' + CurrToStr(result.Kohren_tab_val) + ')'
  else
    lKohren_tab.Caption := '(Gтаб= )';

  if (result.bStudent_tab_val) then
    lStudent_tab.Caption := '(tтаб=' + CurrToStr(result.Student_tab_val) + ')'
  else
    lStudent_tab.Caption := '(tтаб= )';

  lZakl.Caption := result.strZaklForFill;
end;

procedure TShuhartForm.CalcVerification();
var i, j, iCnt, iCntRows: integer;
  bAllNull: boolean;
  X_avg, tmp: Currency;
begin
  iCnt := cxMeasureCnt.ItemIndex + 2;

  result.C_val := StrToCurr(cxC.Text);
  result.CDelta_val := StrToCurr(cxCDelta.Text);
  result.Avg_val := X_avg_avg;

  if (S_SUM_2 <> 0) then begin
    result.Kohren_val := S_MAX_2 / S_SUM_2;
    result.bKohren_val := true;
  end
  else begin
    result.Kohren_val := 0;
    result.bKohren_val := false;
  end;

  result.SKO_Povt_val := sqrt(S_AVG_2);

  // СКО прециз.
  result.SKO_Prec_val := 0;
  iCntRows := 0;
  for i := 0 to cxGridView.DataController.RecordCount - 1 do begin
    bAllNull := true;
    for j := 0 to iCnt - 1 do begin
      if (cxGridView.DataController.Values[i, 1 + j] <> Null) then begin
        bAllNull := false;
        break;
      end
    end;
    if (not bAllNull) then  begin
      X_avg := cxGridView.DataController.Values[i, X_avg_col_index];
      result.SKO_Prec_val := result.SKO_Prec_val + (X_avg - X_AVG_AVG) * (X_avg - X_AVG_AVG);
      iCntRows := iCntRows + 1;
    end;
  end;
  if (iCntRows > 1) then begin
    result.SKO_Prec_val := Sqrt(result.SKO_Prec_val / (iCntRows - 1));
    result.bSKO_Prec_val := true;
  end
  else begin
    result.SKO_Prec_val := 0;
    result.bSKO_Prec_val := false;
  end;

  //
  if ((iCnt - 1 - 1) < Length(Kohren_koef)) AND ((iCntRows - 2) >= 0) AND ((iCntRows - 2) <= Length(Kohren_koef[0])) then begin
    result.Kohren_tab_val := Kohren_koef[iCnt - 1 - 1, iCntRows - 2]; // n - 1 (c 1 до 5), L (с 2)
    result.bKohren_tab_val := true;
  end
  else begin
    result.Kohren_tab_val := 0;
    result.bKohren_tab_val := false;
  end;

  //
  if ((iCntRows - 1 - 1) >= 0) AND ((iCntRows - 1 - 1) < Length(Student_koef)) then begin
    result.Student_tab_val := Student_koef[iCntRows - 1 - 1];  // L - 1 (с 1)
    result.bStudent_tab_val := true;
  end
  else begin
    result.Student_tab_val := 0;
    result.bStudent_tab_val := false;
  end;

  //
  if (result.C_val <> 0) then begin
    result.SKO_Povt_otn_val := result.SKO_Povt_val * 100 / result.C_val;
    result.bSKO_Povt_otn_val := true;
  end
  else begin
    result.SKO_Povt_otn_val := 0;
    result.bSKO_Povt_otn_val := false;
  end;

  //
  if (result.C_val <> 0) then begin
    result.SKO_Prec_otn_val := result.SKO_Prec_val * 100 / result.C_val;
    result.bSKO_Prec_otn_val := true;
  end
  else begin
    result.SKO_Prec_otn_val := 0;
    result.bSKO_Prec_otn_val := false;
  end;

  //
  result.Move_val := abs(X_AVG_AVG - result.C_val);

  //
  if (iCntRows > 0) then begin
    result.SKO_System_Pogr_val := Sqrt(result.SKO_Prec_val * result.SKO_Prec_val / iCntRows + result.CDelta_val * result.CDelta_val / 3);
    result.bSKO_System_Pogr_val := true;
  end
  else begin
    result.SKO_System_Pogr_val := 0;
    result.bSKO_System_Pogr_val := false;
  end;
  //
  if (result.SKO_System_Pogr_val <> 0) then begin
    result.Student_val :=  abs(result.Move_val / result.SKO_System_Pogr_val);
    result.bStudent_val :=  true;
  end
  else begin
    result.Student_val := 0;
    result.bStudent_val :=  false;
  end;

  //
  result.System_Pogr_val := 1.96 * result.SKO_System_Pogr_val;

  //
  if (result.SKO_Prec_val <> 0) then begin
    result.Verify_val := result.System_Pogr_val / result.SKO_Prec_val;
    result.bVerify_val := true;
  end
  else begin
    result.Verify_val := 0;
    result.bVerify_val := false;
  end;

  //
  result.AnalysPogr_val := Sqrt(result.SKO_Prec_val * result.SKO_Prec_val + result.SKO_System_Pogr_val * result.SKO_System_Pogr_val) * 1.96;

  //
  if (result.C_val <> 0) then begin
    result.MethodPogr_val := result.AnalysPogr_val * 100 / result.C_val;
    result.bMethodPogr_val := true;
  end
  else begin
    result.MethodPogr_val := 0;
    result.bMethodPogr_val := false;
  end;

  //
  if ((iCnt - 2) < Length(Q_koef)) then begin
    result.Limit_Povt_val := Round(Q_koef[iCnt - 2] * result.SKO_Povt_otn_val);
    result.Limit_Prec_val := Round(Q_koef[iCnt - 2] * result.SKO_Prec_otn_val + 0.5);

    result.bLimit_Povt_val := true;
    result.bLimit_Prec_val := true;
  end
  else begin
    result.Limit_Povt_val := 0;
    result.Limit_Prec_val := 0;

    result.bLimit_Povt_val := false;
    result.bLimit_Prec_val := false;
  end;

  //
  result.Precision_val  := Round(result.MethodPogr_val);
  result.bPrecision_val := result.bMethodPogr_val;

  result.strZakl := '';
  result.strZaklForFill := '';
  if (result.bLimit_Povt_val) and (iCntRows > 0) then begin
    if (cxR2Contr.Value <> Null) then begin
      if (cxR2Contr.Value >= result.Limit_Povt_val) and (not result.bCriticalShuhart) then begin
        result.strZakl := 'Лаборатория показала показатели достоверности по методике выполнения измерений ' + cxMeasureMethod.Text + ', находящиеся в области разрешенных значений.';
        result.strZaklForFill := 'Лаборатория показала показатели достоверности по методике выполнения измерений ' + cxMeasureMethod.Text + ', находящиеся в области разрешенных значений.';
      end;
      if (cxR2Contr.Value >= result.Limit_Povt_val) and (result.bCriticalShuhart) then begin
        result.strZakl := 'Лаборатория показала показатели достоверности по методике выполнения измерений ' + cxMeasureMethod.Text + ', находящиеся в области разрешенных значений, ' + 'однако по результатам анализа контрольных карт данные измерения влагосодержания могут считаться ограниченно достоверными и лаборатории надлежит провести мероприятия по повышению качества измерений.';
        result.strZaklForFill := 'Лаборатория показала показатели достоверности по методике выполнения измерений ' + cxMeasureMethod.Text + ', находящиеся в области разрешенных значений, ' + 'однако по результатам анализа контрольных карт данные измерения влагосодержания могут считаться ограниченно достоверными и лаборатории надлежит провести мероприятия по повышению качества измерений.';
      end;
      if (cxR2Contr.Value < result.Limit_Povt_val) then begin
        result.strZakl := 'Лаборатория не подтвердила показатели достоверности по методике выполнения измерений ' + cxMeasureMethod.Text + '. Лаборатории надлежит разработать и ' + 'реализовать план мероприятий по повышению качества измерений.';
        result.strZaklForFill := 'Лаборатория не подтвердила показатели достоверности по методике выполнения измерений ' + cxMeasureMethod.Text + '. Лаборатории надлежит разработать и ' + 'реализовать план мероприятий по повышению качества измерений.';
      end;
    end;
  end;
end;

procedure TShuhartForm.RecalcResultRowData(row: integer);
var i: integer;
  bAllNull: boolean;
  tmp, tmp2, OK_val, OK_proc, IL_val, IL_proc: Currency;
  Pogr: integer;
begin

  bAllNull := true;
  for i := 2 to 5 do begin
    if (cxGridResultView.DataController.Values[row, i] <> Null) then begin
      bAllNull := false;
      break;
    end;
  end;
  if (bAllNull) then begin
    cxGridResultView.DataController.Values[row, 7] := '';
    cxGridResultView.DataController.Values[row, 8] := '';
    cxGridResultView.DataController.Values[row, 9] := '';
    exit;
  end;
  if (cxGridResultView.DataController.Values[row, 2] <> Null) then
    OK_val := cxGridResultView.DataController.Values[row, 2]
  else
    OK_val := 0;
  if (cxGridResultView.DataController.Values[row, 3] <> Null) then
    OK_proc := cxGridResultView.DataController.Values[row, 3]
  else
    OK_proc := 0;
  if (cxGridResultView.DataController.Values[row, 4] <> Null) then
    IL_val := cxGridResultView.DataController.Values[row, 4]
  else
    IL_val := 0;
  if (cxGridResultView.DataController.Values[row, 5] <> Null) then
    IL_proc := cxGridResultView.DataController.Values[row, 5]
  else
    IL_proc := 0;
  if (cxGridResultView.DataController.Values[row, 6] <> Null) then
    Pogr := cxGridResultView.DataController.Values[row, 6]
  else
    Pogr := 0;
  
  if (OK_val <> 0) then begin
    tmp := (IL_val - OK_val) * 100 / OK_val;
    if (tmp > 0) then
      cxGridResultView.DataController.Values[row, 8] := '+' + CurrToStrF(tmp, ffFixed, 2)
    else
      cxGridResultView.DataController.Values[row, 8] := CurrToStrF(tmp, ffFixed, 2);

    tmp2 := 2 * sqrt(((100 * OK_proc / OK_val) / 1.96) * ((100 * OK_proc / OK_val) / 1.96) + (Pogr / 1.96) * (Pogr / 1.96));
    cxGridResultView.DataController.Values[row, 7] := '±' + CurrToStrF(tmp2, ffFixed, 2);

    if (abs(tmp) < tmp2) then
      cxGridResultView.DataController.Values[row, 9] := 'Погрешность удовл.'
    else
      cxGridResultView.DataController.Values[row, 9] := 'Погрешность неудовл.';
  end
  else begin
    cxGridResultView.DataController.Values[row, 7] := 'нет значения';
    cxGridResultView.DataController.Values[row, 8] := 'нет значения';
    cxGridResultView.DataController.Values[row, 9] := '';
  end;

end;

procedure TShuhartForm.RecalcRowData(row: integer);
var
  //arr_val : array of Currency;
  bAllNull: boolean;
  iCnt, i, row_h, row_l: integer;
  minVal, maxVal, allSum, SKO, med, adv, tmp: Currency;
  list : TList;
  val: TCurrencyValue;
begin
  if (cxCardType.ItemIndex = 0) then begin
    iCnt := cxMeasureCnt.ItemIndex + 2;
    //SetLength(arr_val, iCnt);
    list := TList.Create;

    bAllNull := true;
    minVal := 1000000;
    maxVal := -1000000;
    allSum := 0;
    for i := 0 to iCnt - 1 do begin
      if (cxGridView.DataController.Values[row, 1 + i] <> Null) then begin
        bAllNull := false;
        tmp := cxGridView.DataController.Values[row, 1 + i];
      end
      else tmp := 0;

      // для подсчета медианы необходимо сортировать данные по возрастанию
      val := TCurrencyValue.Create(tmp);
      list.Add(val);

      if (minVal > tmp) then minVal := tmp;
      if (maxVal < tmp) then maxVal := tmp;
      allSum := allSum + tmp;
    end;

    list.Sort(compare);

    if (bAllNull) then begin
      cxGridView.DataController.Values[row, X_avg_col_index] := Null;
      cxGridView.DataController.Values[row, S_col_index] := Null;
      cxGridView.DataController.Values[row, Me_col_index] := Null;
      cxGridView.DataController.Values[row, R_col_index] := Null;
      cxGridView.DataController.Values[row, CV_col_index] := Null;
      exit;
    end;

    if ((iCnt mod 2) <> 0) then med := TCurrencyValue(list.Items[iCnt div 2]).val
    else med := (TCurrencyValue(list.Items[(iCnt div 2) - 1]).val + TCurrencyValue(list.Items[iCnt div 2]).val) / 2;

    adv := allSum / iCnt;
    SKO := 0;
    for i := 0 to iCnt - 1 do begin
      SKO := SKO + (adv - TCurrencyValue(list.Items[i]).val) * (adv - TCurrencyValue(list.Items[i]).val);
    end;
    SKO := sqrt(SKO / (iCnt - 1));
    // среднее значение
    cxGridView.DataController.Values[row, X_avg_col_index] := adv;
    // СКО
    cxGridView.DataController.Values[row, S_col_index] := SKO;
    // медиана
    cxGridView.DataController.Values[row, Me_col_index] := med;
    // размах
    cxGridView.DataController.Values[row, R_col_index] := maxVal - minVal;
    // CV
    if (adv <> 0) then
      cxGridView.DataController.Values[row, CV_col_index] := SKO * 100 / adv
    else
      cxGridView.DataController.Values[row, CV_col_index] := NULL;

    list.Free;
  end
  else begin
    // ищем верхнего соседа
    row_h := -1;
    for i := row - 1 downto 0 do begin
      if (cxGridView.DataController.Values[i, 1] <> Null) then begin
        row_h := i;
        break;
      end;
    end;

    // ищем нижнего соседа
    row_l := -1;
    for i := row + 1 to cxGridView.DataController.RecordCount - 1 do begin
      if (cxGridView.DataController.Values[i, 1] <> Null) then begin
        row_l := i;
        break;
      end;
    end;

    if (cxGridView.DataController.Values[row, 1] <> Null) then begin
      if (row_l >= 0) AND (row_h < 0) then begin
        tmp := cxGridView.DataController.Values[row_l, 1] - cxGridView.DataController.Values[row, 1];
        if (tmp < 0) then tmp := -tmp;
        cxGridView.DataController.Values[row_l, R_col_index] := tmp;
        cxGridView.DataController.Values[row, R_col_index] := NULL;
      end;
      if (row_l < 0) AND (row_h >= 0) then begin
        tmp := cxGridView.DataController.Values[row, 1] - cxGridView.DataController.Values[row_h, 1];
        if (tmp < 0) then tmp := -tmp;
        cxGridView.DataController.Values[row, R_col_index] := tmp;
      end;
      if (row_l >= 0) AND (row_h >= 0) then begin
        tmp := cxGridView.DataController.Values[row, 1] - cxGridView.DataController.Values[row_h, 1];
        if (tmp < 0) then tmp := -tmp;
        cxGridView.DataController.Values[row, R_col_index] := tmp;
        tmp := cxGridView.DataController.Values[row_l, 1] - cxGridView.DataController.Values[row, 1];
        if (tmp < 0) then tmp := -tmp;
        cxGridView.DataController.Values[row_l, R_col_index] := tmp;
      end;
    end
    else begin
      cxGridView.DataController.Values[row, R_col_index] := Null;
      if (row_l >= 0) AND (row_h >= 0) then begin
        tmp := cxGridView.DataController.Values[row_l, 1] - cxGridView.DataController.Values[row_h, 1];
        if (tmp < 0) then tmp := -tmp;
        cxGridView.DataController.Values[row_l, R_col_index] := tmp;
        cxGridView.DataController.Values[row, R_col_index] := NULL;
      end;
      if (row_l >= 0) AND (row_h < 0) then begin
        cxGridView.DataController.Values[row_l, R_col_index] := Null;
        cxGridView.DataController.Values[row, R_col_index] := NULL;
      end;
    end;

  end;
end;

{procedure TShuhartForm.PrintHeader (var word: TWord; bShuhart: boolean);
var
  mBookMark: array of TPair;
  i, j, iCntRows, iCnt: Integer;
  bAllNull: boolean;
  str: string;
begin

  SetLength(mBookMark, 6);

  mBookMark[0].strName := 'K';  //
  if (result.bPrecision_val) then
    mBookMark[0].strValue := IntToStr(result.Precision_val)
  else
    mBookMark[0].strValue := 'нет значения';

  mBookMark[1].strName := 'rrл'; //
  if (result.bLimit_Prec_val) then
    mBookMark[1].strValue := IntToStr(result.Limit_Prec_val)
  else
    mBookMark[1].strValue := 'нет значения';

  mBookMark[2].strName := 'rл'; //
  if (result.bLimit_Povt_val) then
    mBookMark[2].strValue := IntToStr(result.Limit_Povt_val)
  else
    mBookMark[2].strValue := 'нет значения';

  mBookMark[3].strName := 'Дата';
  mBookMark[3].strValue := FormatDateTime('yyyy', Now);

  mBookMark[4].strName := 'Заключение1';
  mBookMark[4].strValue := result.strZakl;

  str := '';
  for i := 0 to Length(result.arrMCILabInfo) - 1 do begin
    if (result.arrMCILabInfo[i].bResult) then begin
      if (i = 0) then
        str := 'Показатели прецизионности (воспроизводимость) ' + result.arrMCILabInfo[i].strLaboratoryName + ' по определению _____________ находится в пределах, установленных нормативами. Данные контроля, анализа испытаний и измерений отвечают требованиям достоверности.'
      else
        str := str + chr(13) + 'Показатели прецизионности (воспроизводимость) ' + result.arrMCILabInfo[i].strLaboratoryName + ' по определению _____________ находится в пределах, установленных нормативами. Данные контроля, анализа испытаний и измерений отвечают требованиям достоверности.';
    end
    else begin
      if (i = 0) then
        str := 'Показатели прецизионности ' + result.arrMCILabInfo[i].strLaboratoryName + ' по определению _____________ находится вне пределов, установленных нормативами. Данные контроля, анализа испытаний и измерений не отвечают требованиям достоверности.'
      else
        str := str + chr(13) + 'Показатели прецизионности ' + result.arrMCILabInfo[i].strLaboratoryName + ' по определению _____________ находится вне пределов, установленных нормативами. Данные контроля, анализа испытаний и измерений не отвечают требованиям достоверности.';
    end;
  end;

  mBookMark[5].strName := 'Заключение2';
  mBookMark[5].strValue := str;

  word.SetBookmarkText(mBookMark);
end;  }

procedure TShuhartForm.PrintReport2(var word: TWord);
var iCntRows : integer;
  i, j, ind, ind_table: integer;
  val1, val2: Currency;
  str, strLaboratory: string;
  mBookMark: array of TPair;
  strDateBegin, strDateEnd, strMCIMeasureName : string;
begin
  SetLength(mBookMark, 10);

  mBookMark[0].strName := 'ед_изм_МСИ';    //
  mBookMark[0].strValue := cxMCIUnitName.Text;

  mBookMark[1].strName := 'ед_изм_МСИ2';  //
  mBookMark[1].strValue := cxMCIUnitName.Text;

  mBookMark[2].strName := 'Дата';  //
  mBookMark[2].strValue := FormatDateTime('yyyy', Now);

  if (cxMCIMeasureName.ItemIndex >= 0) then
    strMCIMeasureName := dataMeasureName.Values[cxMCIMeasureName.Text]
  else
    strMCIMeasureName := cxMCIMeasureName.Text;

  str := '';
  strLaboratory := '';
  for i := 0 to Length(result.arrMCILabInfo) - 1 do begin
    if (i <> 0) then str := str + ';' + chr(13);
    if (i <> 0) then strLaboratory := strLaboratory + ';' + chr(13);
    if (result.arrMCILabInfo[i].bResult) then begin
      str := str + 'Данные испытаний и измерений, проведенных в ' + result.arrMCILabInfo[i].strLaboratoryName + ' по определению ' + strMCIMeasureName + ' по ' + cxMCIMeasureMethod.Text + ' отвечают требованиям достоверности (показатели прецизионности находятся в пределах, установленных нормативами)';
    end
    else begin
      str := str + 'Данные испытаний и измерений, проведенных в ' + result.arrMCILabInfo[i].strLaboratoryName + ' по определению ' + strMCIMeasureName + ' по ' + cxMCIMeasureMethod.Text + ' не отвечают требованиям достоверности (показатели прецизионности находятся вне пределов, установленных нормативами)';
    end;
    strLaboratory := strLaboratory + result.arrMCILabInfo[i].strLaboratoryName;
  end;
  str := str + '.';
  strLaboratory := strLaboratory + '.';

  mBookMark[3].strName := 'Заключение';  //
  mBookMark[3].strValue := str;

  mBookMark[4].strName := 'Контролируемый_параметр';  //
  mBookMark[4].strValue := strMCIMeasureName;

  mBookMark[5].strName := 'Метод_измерения';  //
  mBookMark[5].strValue := cxMCIMeasureMethod.Text;

  mBookMark[6].strName := 'Модель';  //
  case cxMCIModel.ItemIndex of
    0:mBookMark[6].strValue := 'модели параллельных измерений';
    1:mBookMark[6].strValue := 'модели последовательный измерений';
    2:mBookMark[6].strValue := 'интерпретационной модели';
    3:mBookMark[6].strValue := 'модели анализа проб (части процесса)';
    4:mBookMark[6].strValue := 'модели разделенных проб';
  end;

  mBookMark[7].strName := 'Описание';  //
  mBookMark[7].strValue := cxMCIDescript.Text;

  mBookMark[8].strName := 'Период';  //
  if (cxMCIDateBegin.Text <> '') then begin
    strDateBegin := '«' + FormatDateTime('dd', cxMCIDateBegin.Date) + '» ' + month_str[MonthOfTheYear(cxMCIDateBegin.Date)] + ' ' + FormatDateTime('yyyy', cxMCIDateBegin.Date) + ' г.';
  end;
  if (cxMCIDateEnd.Text <> '') then begin
    strDateEnd := '«' + FormatDateTime('dd', cxMCIDateEnd.Date) + '» ' + month_str[MonthOfTheYear(cxMCIDateEnd.Date)] + ' ' + FormatDateTime('yyyy', cxMCIDateEnd.Date) + ' г.';
  end;
  mBookMark[8].strValue := strDateBegin + ' по ' + strDateEnd;

  mBookMark[9].strName := 'Список_лабораторий';  //
  mBookMark[9].strValue := strLaboratory;

  word.SetBookmarkText(mBookMark);

  ind_table := 3;

  iCntRows := cxGridLaboratoryView.DataController.RecordCount + cxGridLaboratoryView.DataController.RecordCount * (cxOKCnt.ItemIndex + 1);
  if (iCntRows > 0) then begin
    word.InsertRowsInTable(ind_table, iCntRows - 1);

    ind := 2;
    for i := 0 to cxGridResultView.DataController.RecordCount - 1 do begin
      if ((i mod (cxOKCnt.ItemIndex + 1)) = 0) then begin
        word.SetTextInCell(1, ind, cxGridResultView.DataController.Values[i, 0], ind_table);
        word.SetCellsMerge(1, 7, ind, ind, TCellVerticalAlignment.CellAlignVerticalCenter, ind_table);
        ind := ind + 1;
      end;
      word.SetTextInCell(1, ind, cxGridResultView.DataController.Values[i, 1], ind_table);

      if (cxGridResultView.DataController.Values[i, 2] <> Null) then begin
        val1 := cxGridResultView.DataController.Values[i, 2];
        str := CurrToStr(val1);
      end
      else str := '';
      str := str + '±';
      if (cxGridResultView.DataController.Values[i, 3] <> Null) then begin
        val2 := cxGridResultView.DataController.Values[i, 3];
        str := str + CurrToStr(val1);
      end;
      word.SetTextInCell(2, ind, str, ind_table);

      if (cxGridResultView.DataController.Values[i, 4] <> Null) then begin
        val1 := cxGridResultView.DataController.Values[i, 4];
        str := CurrToStr(val1);
      end
      else str := '';
      str := str + '±';
      if (cxGridResultView.DataController.Values[i, 5] <> Null) then begin
        val2 := cxGridResultView.DataController.Values[i, 5];
        str := str + CurrToStr(val1);
      end;
      word.SetTextInCell(3, ind, str, ind_table);

      if (cxGridResultView.DataController.Values[i, 6] <> Null) then str := cxGridResultView.DataController.Values[i, 6] + '%'
      else str := '';
      word.SetTextInCell(4, ind, str, ind_table);

      if (cxGridResultView.DataController.Values[i, 7] <> Null) then str := cxGridResultView.DataController.Values[i, 7]
      else str := '';
      word.SetTextInCell(5, ind, str, ind_table);

      if (cxGridResultView.DataController.Values[i, 8] <> Null) then str := cxGridResultView.DataController.Values[i, 8]
      else str := '';
      word.SetTextInCell(6, ind, str, ind_table);

      if (cxGridResultView.DataController.Values[i, 9] <> Null) then str := cxGridResultView.DataController.Values[i, 9]
      else str := '';
      word.SetTextInCell(7, ind, str, ind_table);

      ind := ind + 1;
    end;

  end;
end;

procedure TShuhartForm.PrintReport1(var word: TWord; DIR: string);
var
  mBookMark: array of TPair;
  i, j, iCntRows, iCnt, ind_table: Integer;
  bAllNull: boolean;
  vecData: array of array of String;
  strDateBegin, strDateEnd: string;
  //Year: word;
  //Year, Month, Day: Word;
begin
  iCntRows := 0;
  iCnt := cxMeasureCnt.ItemIndex + 2;

  SetLength(vecData, cxGridView.DataController.RecordCount);

  for i := 0 to cxGridView.DataController.RecordCount - 1 do begin
    bAllNull := true;
    for j := 0 to iCnt - 1 do begin
      if (cxGridView.DataController.Values[i, 1 + j] <> Null) then begin
        bAllNull := false;
        break;
      end
    end;
    if (not bAllNull) then  begin
      SetLength(vecData[iCntRows], iCnt + 3);
      for j := 0 to iCnt - 1 do begin
        if (cxGridView.DataController.Values[i, 1 + j] <> Null) then
          vecData[i, j] := CurrToStr(cxGridView.DataController.Values[i, 1 + j])
        else
          vecData[i, j] := '0';
      end;

      vecData[i, iCnt] := CurrToStr(cxGridView.DataController.Values[i, X_avg_col_index]);
      vecData[i, iCnt + 1] := CurrToStr(cxGridView.DataController.Values[i, S_col_index]);
      vecData[i, iCnt + 2] := CurrToStr(cxGridView.DataController.Values[i, CV_col_index]);

      iCntRows := iCntRows + 1;
    end;
  end;

  SetLength(vecData, iCntRows);

  SetLength(mBookMark, 45);

  mBookMark[0].strName := 'G';
  if (result.bKohren_val) then
    mBookMark[0].strValue := FloatToStrF(result.Kohren_val, ffFixed, 8, 4)
  else
    mBookMark[0].strValue := 'нет значения';

  mBookMark[1].strName := 'Gтаб';
  if (result.bKohren_tab_val) then
    mBookMark[1].strValue := FloatToStrF(result.Kohren_tab_val, ffFixed, 8, 4)
  else
    mBookMark[1].strValue := 'нет значения';

  mBookMark[2].strName := 'K2';
  if (result.bKohren_tab_val) then
    mBookMark[2].strValue := FloatToStrF(result.Kohren_tab_val, ffFixed, 8, 4)
  else
    mBookMark[2].strValue := 'нет значения';

  mBookMark[3].strName := 'PCI';
  if (result.bPCI_val) then
    mBookMark[3].strValue := FloatToStrF(result.PCI_val, ffFixed, 12, 6)
  else
    mBookMark[3].strValue := 'нет значения';

  mBookMark[4].strName := 'rrл2';
  if (result.bLimit_Prec_val) then
    mBookMark[4].strValue := IntToStr(result.Limit_Prec_val)
  else
    mBookMark[4].strValue := 'нет значения';

  mBookMark[5].strName := 'rл2';
  if (result.bLimit_Povt_val) then
    mBookMark[5].strValue := IntToStr(result.Limit_Povt_val)
  else
    mBookMark[5].strValue := 'нет значения';

  mBookMark[6].strName := 'srrл';
  if (result.bSKO_Prec_val) then
    mBookMark[6].strValue := FloatToStrF(result.SKO_Prec_val, ffFixed, 8, 4)
  else
    mBookMark[6].strValue := 'нет значения';

  mBookMark[7].strName := 'srrотнл';
  if (result.bSKO_Prec_otn_val) then
    mBookMark[7].strValue := FloatToStrF(result.SKO_Prec_otn_val, ffFixed, 8, 4)
  else
    mBookMark[7].strValue := 'нет значения';

  mBookMark[8].strName := 'srл';
  mBookMark[8].strValue := FloatToStrF(result.SKO_Povt_val, ffFixed, 8, 4);

  mBookMark[9].strName := 'srотнл';
  if (result.bSKO_Povt_otn_val) then
    mBookMark[9].strValue := FloatToStrF(result.SKO_Povt_otn_val, ffFixed, 8, 4)
  else
    mBookMark[9].strValue := 'нет значения';

  mBookMark[10].strName := 'sсл';
  if (result.bSKO_System_Pogr_val) then
    mBookMark[10].strValue := FloatToStrF(result.SKO_System_Pogr_val, ffFixed, 8, 4)
  else
    mBookMark[10].strValue := 'нет значения';

  mBookMark[11].strName := 't';
  if (result.bStudent_val) then
    mBookMark[11].strValue := FloatToStrF(result.Student_val, ffFixed, 8, 4)
  else
    mBookMark[11].strValue := 'нет значения';

  mBookMark[12].strName := 'tтаб';
  if (result.bStudent_tab_val) then
    mBookMark[12].strValue := FloatToStrF(result.Student_tab_val, ffFixed, 8, 4)
  else
    mBookMark[12].strValue := 'нет значения';

  mBookMark[13].strName := 'атт_знач';
  mBookMark[13].strValue := CurrToStr(result.C_val) + '±' + CurrToStr(result.CDelta_val) + ' ' + cxUnitName.Text;

  mBookMark[14].strName := 'ед_изм';
  mBookMark[14].strValue := cxUnitName.Text;

  mBookMark[15].strName := 'ед_изм2';
  mBookMark[15].strValue := cxUnitName.Text;

  mBookMark[16].strName := 'ед_изм3';
  mBookMark[16].strValue := cxUnitName.Text;

  mBookMark[17].strName := 'ед_изм4';
  mBookMark[17].strValue := cxUnitName.Text;

  mBookMark[18].strName := 'ед_изм5';
  mBookMark[18].strValue := cxUnitName.Text;

  mBookMark[19].strName := 'ед_изм6';
  mBookMark[19].strValue := cxUnitName.Text;

  mBookMark[20].strName := 'ед_изм7';
  mBookMark[20].strValue := cxUnitName.Text;

  mBookMark[21].strName := 'ед_изм8';
  mBookMark[21].strValue := cxUnitName.Text;

  mBookMark[22].strName := '';
  mBookMark[22].strValue := '';

  mBookMark[23].strName := 'ед_изм10';
  mBookMark[23].strValue := cxUnitName.Text;

  mBookMark[24].strName := 'колво_результатов';
  mBookMark[24].strValue := IntToStr(iCnt);

  mBookMark[25].strName := 'Название_карты';
  if (cxAnalyzeType.ItemIndex = 0) then mBookMark[25].strValue := 'Карта средних значений';
  if (cxAnalyzeType.ItemIndex = 1) then mBookMark[25].strValue := 'Карта выборочных СКО';
  if (cxAnalyzeType.ItemIndex = 2) then mBookMark[25].strValue := 'Карта медиан';

  mBookMark[26].strName := 'Название_карты2';
  if (cxAnalyzeType.ItemIndex = 0) then mBookMark[26].strValue := 'Карта средних значений';
  if (cxAnalyzeType.ItemIndex = 1) then mBookMark[26].strValue := 'Карта выборочных СКО';
  if (cxAnalyzeType.ItemIndex = 2) then mBookMark[26].strValue := 'Карта медиан';

  mBookMark[27].strName := 'Погрешность_анализа';
  mBookMark[27].strValue := FloatToStrF(result.AnalysPogr_val, ffFixed, 8, 4);

  mBookMark[28].strName := 'Погрешность_методики';
  if (result.bMethodPogr_val) then
    mBookMark[28].strValue := FloatToStrF(result.MethodPogr_val, ffFixed, 8, 4)
  else
    mBookMark[28].strValue := 'нет значения';

  mBookMark[29].strName := 'Проверка_значимости';
  if (result.bVerify_val) then
    mBookMark[29].strValue := FloatToStrF(result.Verify_val, ffFixed, 8, 4)
  else
    mBookMark[29].strValue := 'нет значения';

  mBookMark[30].strName := 'Системная_погрешность';
  mBookMark[30].strValue := FloatToStrF(result.System_Pogr_val, ffFixed, 8, 4);

  mBookMark[31].strName := 'Смещение';
  mBookMark[31].strValue := FloatToStrF(result.Move_val, ffFixed, 8, 4);

  mBookMark[32].strName := 'Хсрср';
  mBookMark[32].strValue := CurrToStr(result.Avg_val);

  mBookMark[33].strName := 'Название_лаборатории2';
  mBookMark[33].strValue := cxCardName.Text;

  mBookMark[34].strName := 'Период';
  if (cxDateBegin.Text <> '') then begin
    strDateBegin := '«' + FormatDateTime('dd', cxDateBegin.Date) + '» ' + month_str[MonthOfTheYear(cxDateBegin.Date)] + ' ' + FormatDateTime('yyyy', cxDateBegin.Date) + ' г.';
  end;
  if (cxDateEnd.Text <> '') then begin
    strDateEnd := '«' + FormatDateTime('dd', cxDateEnd.Date) + '» ' + month_str[MonthOfTheYear(cxDateEnd.Date)] + ' ' + FormatDateTime('yyyy', cxDateEnd.Date) + ' г.';
  end;
  mBookMark[34].strValue := strDateBegin + ' по ' + strDateEnd;


  mBookMark[35].strName := 'Контрольный_параметр';
  if (cxMeasureName.ItemIndex >= 0) then
    mBookMark[35].strValue := dataMeasureName.Values[cxMeasureName.Text]
  else
    mBookMark[35].strValue := cxMeasureName.Text;

  mBookMark[36].strName := 'Метод_измерения';
  mBookMark[36].strValue := cxMeasureMethod.Text;

  mBookMark[37].strName := 'K';  //
  if (result.bPrecision_val) then
    mBookMark[37].strValue := IntToStr(result.Precision_val)
  else
    mBookMark[37].strValue := 'нет значения';

  mBookMark[38].strName := 'rrл'; //
  if (result.bLimit_Prec_val) then
    mBookMark[38].strValue := IntToStr(result.Limit_Prec_val)
  else
    mBookMark[38].strValue := 'нет значения';

  mBookMark[39].strName := 'rл'; //
  if (result.bLimit_Povt_val) then
    mBookMark[39].strValue := IntToStr(result.Limit_Povt_val)
  else
    mBookMark[39].strValue := 'нет значения';

  mBookMark[40].strName := 'Дата';
  mBookMark[40].strValue := FormatDateTime('yyyy', Now);

  mBookMark[41].strName := 'Название_лаборатории';
  mBookMark[41].strValue := cxCardName.Text;

  mBookMark[42].strName := 'Заключение1';
  mBookMark[42].strValue := result.strZakl;

  mBookMark[43].strName := 'r_контр';
  mBookMark[43].strValue := cxRContr.Text;

  mBookMark[44].strName := 'r2_контр';
  mBookMark[44].strValue := cxR2Contr.Text;

  // Окончание_измерений

  word.SetBookmarkText(mBookMark);

  ind_table := 4;

  if (iCnt > 4) then begin
    for i := iCnt + 1 to 20 do begin
      word.DeleteColumn(iCnt + 1 + 1, ind_table);
    end;

    for i := 1 to iCnt + 1 do begin // вместе со средним значением
      word.SetColumnWidth(i + 1, 20, ind_table);
    end;
  end
  else begin
    for i := iCnt + 1 to 4 do begin
      word.DeleteColumn(iCnt + 1 + 1, ind_table);
    end;

    if (iCnt < 4) then begin
      for i := 1 to iCnt + 1 do begin // вместе со средним значением
        word.SetColumnWidth(i + 1, (4 + 1) * 20.0 / (iCnt + 1), ind_table);
      end;
    end
  end;

  //for i := 1 to iCnt do word.SetTextInCell(i + 3, 3, cxUnitName.Text, ind_table);

  word.InsertRowsInTable(ind_table, iCntRows - 1);
  for i := 1 to iCntRows do begin
    word.SetTextInCell(1, i + 2, IntToStr(i), ind_table);
  end;

  for i := 0 to Length(vecData) - 1 do begin
    for j := 0 to Length(vecData[i]) - 1 do begin
      word.SetTextInCell(j + 1 + 1, i + 2 + 1, vecData[i, j], ind_table);
    end;
  end;

  //word.SetTextInCell(2, 4, CurrToStr(result.C_val), ind_table);
  //word.SetTextInCell(3, 4, CurrToStr(result.CDelta_val), ind_table);
  word.SetCellsMerge(2, 2 + iCnt - 1, 1, 1, CellAlignVerticalCenter, ind_table);

  // переносим критерии
  for i := 0 to cxGridCriteriaView.DataController.RecordCount - 1 do begin
    word.SetTextInCell(2, 2 + i, cxGridCriteriaView.DataController.Values[i, 1], ind_table + 4);
    word.SetTextInCell(3, 2 + i, cxGridCriteriaView.DataController.Values[i, 2], ind_table + 4);
    word.SetTextInCell(4, 2 + i, cxGridCriteriaView.DataController.Values[i, 3], ind_table + 4);
  end;


  if (cxAnalyzeType.ItemIndex = 0) then begin
    DrawXCard();
    X_Chart.Update;
    X_Chart.SaveToBitmapFile(DIR + '\\' + 'tmp.bmp');
  end;
  if (cxAnalyzeType.ItemIndex = 1) then begin
    DrawSCard();
    S_Chart.Update;
    S_Chart.SaveToBitmapFile(DIR + '\\' + 'tmp.bmp');
  end;
  if (cxAnalyzeType.ItemIndex = 2) then begin
    DrawMeCard();
    Me_Chart.Update;
    Me_Chart.SaveToBitmapFile(DIR + '\\' + 'tmp.bmp');
  end;
  word.InsertPicture(DIR + '\\' + 'tmp.bmp', AlignParagraphCenter, 'рис1');

  DrawRCard();
  R_Chart.Update;
  R_Chart.SaveToBitmapFile(DIR + '\\' + 'tmp.bmp');
  word.InsertPicture(DIR + '\\' + 'tmp.bmp', AlignParagraphCenter, 'рис2');
  DrawConvChart();
  SKO_Chart.Update;
  SKO_Chart.SaveToBitmapFile(DIR + '\\' + 'tmp.bmp');
  word.InsertPicture(DIR + '\\' + 'tmp.bmp', AlignParagraphCenter, 'рис3');
end;

procedure TShuhartForm.Report1Execute(Sender: TObject);
var word: TWord;
  DIR, strFileName: string;
  iCnt: Integer;
  bShuhart: boolean;
  result: integer;
begin

  if (cxCardType.ItemIndex <> 0) then begin
    MessageDLG('Для индивидуальных данных вывод отчета не поддерживается.',mtError,[mbOk], 0);
    exit;
  end;

  if (cxRContr.Text = '') or (cxR2Contr.Text = '') then begin
    result := MessageDLG('Не заданы контрольные значения - протокол не будет содержать заключения!',mtWarning,[mbYes,mbNo], 0);
    if (result = mrNo) then exit;
  end;

  iCnt := cxMeasureCnt.ItemIndex + 2;

  try
    DIR := ExtractFileDir(Application.ExeName);
    strFileName := DIR + '\\' + 'Протокол_внутрилаб_иссл.dot';
    if (iCnt > 4) then strFileName := DIR + '\\' + 'Протокол_внутрилаб_иссл_альб.dot';

    CalcDataKoef();
    CalcCriteries();
    CalcVerification();

    word := TWord.Create;
    if (word.Start(strFileName)) then begin
      PrintReport1(word, DIR);
      word.SetVisible(true);
    end;
  finally
    word.Free;
  end;
end;

procedure TShuhartForm.Report2Execute(Sender: TObject);
var word: TWord;
  DIR, strFileName: string;
begin

  if (cxCardType.ItemIndex <> 0) then begin
    MessageDLG('Для индивидуальных данных вывод отчета не поддерживается.',mtError,[mbOk], 0);
    exit;
  end;

    if (not VerifyDoubleLaboratory()) then begin
      cxPageControl3.ActivePageIndex := 0;
      MessageDLG('В списке лабораторий не должно быть лабораторий с повторяющимися или пустыми названиями.',mtError,[mbOk], 0);
      exit;
    end;

  try
    DIR := ExtractFileDir(Application.ExeName);
    strFileName := DIR + '\\' + 'Протокол_МСИ.dot';

    InitResultMCIGrid(false);
    CalcResultMCI();

    word := TWord.Create;
    if (word.Start(strFileName)) then begin
      PrintReport2(word);
      word.SetVisible(true);
    end;
  finally
    word.Free;
  end;
end;

{procedure TShuhartForm.Report3Execute(Sender: TObject);
var word: TWord;
  DIR, strFileName: string;
  iCnt: Integer;
  result: integer;
  bShuhart: boolean;
begin

  if (cxRContr.Text = '') or (cxR2Contr.Text = '') then begin
    result := MessageDLG('Не заданы контрольные значения - протокол не будет содержать заключения!',mtWarning,[mbYes,mbNo], 0);
    if (result = mrNo) then exit;
  end;

  if (cxCardType.ItemIndex <> 0) then begin
    MessageDLG('Для индивидуальных данных вывод отчета не поддерживается.',mtError,[mbOk], 0);
    exit;
  end;

  iCnt := cxMeasureCnt.ItemIndex + 2;

  try
    DIR := ExtractFileDir(Application.ExeName);
    strFileName := DIR + '\\' + 'Протокол.dot';
    if (iCnt > 4) then strFileName := DIR + '\\' + 'Протокол_альб.dot';

    CalcDataKoef();
    CalcCriteries();
    CalcVerification();
    DrawConvChart();

    InitResultMCIGrid(false);
    CalcResultMCI();

    word := TWord.Create;
    if (word.Start(strFileName)) then begin
      PrintHeader(word, bShuhart);
      PrintReport1(word, DIR, 4);
      PrintReport2(word, 9);
      word.SetVisible(true);
    end;
  finally
    word.Free;
  end;
end;   }

procedure TShuhartForm.cxAnalyzeTypeFocusChanged(Sender: TObject);
begin
  if (not cxAnalyzeType.IsFocused) then begin
    InitGrid();
  end;
end;

procedure TShuhartForm.cxAnalyzeTypePropertiesChange(Sender: TObject);
begin
  if (cxAnalyzeType.ItemIndex = 0) then begin
    cxGridCriteriaViewColumn4.Caption := 'Карта средних';
  end;
  if (cxAnalyzeType.ItemIndex = 1) then begin
    cxGridCriteriaViewColumn4.Caption := 'Карта выборочных СКО';
  end;
  if (cxAnalyzeType.ItemIndex = 2) then begin
    cxGridCriteriaViewColumn4.Caption := 'Карта медиан';
  end;
end;

procedure TShuhartForm.cxCardTypeFocusChanged(Sender: TObject);
var i: integer;
begin
  if (not cxCardType.IsFocused) then begin

    {for i := 1 to cxGridView.ColumnCount - 1 do begin
      for j := 0 to cxGridView.DataController.RecordCount - 1 do begin
        cxGridView.DataController.Values[j, i] := Null;
      end;
    end;}

    if (cxCardType.ItemIndex = 0) then begin
      cxMeasureCnt.Enabled := true;
    end
    else begin
      cxMeasureCnt.ItemIndex := 0;
      cxMeasureCnt.Enabled := false;
    end;
    InitGrid();

    cxGridView.DataController.BeginUpdate;
    for i := 0 to cxGridView.DataController.RecordCount - 1 do begin
      RecalcRowData(i);
    end;
    cxGridView.DataController.EndUpdate;
  end;
end;
  procedure TShuhartForm.cxCDeltaFocusChanged(Sender: TObject);
begin
  if (not cxCDelta.IsFocused) then begin
    if (cxCDelta.Text = '') then cxCDelta.Value := 0;
  end;
end;

procedure TShuhartForm.cxColumnX1PropertiesValidate(Sender: TObject;
  var DisplayValue: Variant; var ErrorText: TCaption; var Error: Boolean);
begin
  if DisplayValue = '' then
    cxGridView.DataController.Values[cxGridView.DataController.FocusedRowIndex, cxGridView.Controller.FocusedColumnIndex] := Null
  else
    cxGridView.DataController.Values[cxGridView.DataController.FocusedRowIndex, cxGridView.Controller.FocusedColumnIndex] := DisplayValue;

  cxGridView.DataController.BeginUpdate;
  RecalcRowData(cxGridView.DataController.FocusedRowIndex);
  cxGridView.DataController.EndUpdate;
end;

procedure TShuhartForm.cxDigitCntFocusChanged(Sender: TObject);
begin
  if (not cxDigitCnt.IsFocused) then begin
    InitGrid();
  end;
end;

procedure TShuhartForm.cxGridCriteriaViewCustomDrawCell(
  Sender: TcxCustomGridTableView; ACanvas: TcxCanvas;
  AViewInfo: TcxGridTableDataCellViewInfo; var ADone: Boolean);
var
  Cellvalue, Cellvalue2: string;
begin

  if (AViewInfo.Item.Index = 3) then begin
    if (not AViewInfo.Selected) then begin
      Cellvalue := AViewInfo.GridRecord.Values[2];
      Cellvalue2 := AViewInfo.GridRecord.Values[3];
      if (Cellvalue2 = 'Критерий не обнаружен') then begin
        ACanvas.Canvas.Font.Color := clBlack;
      end
      else begin
        if (Cellvalue = 'Критический') then
          ACanvas.Canvas.Font.Color := clRed
        else
          ACanvas.Canvas.Font.Color := clBlue;
      end;
    end;
  end;

end;

procedure TShuhartForm.cxGridResultViewColumn1PropertiesValidate(
  Sender: TObject; var DisplayValue: Variant; var ErrorText: TCaption;
  var Error: Boolean);
begin
  if DisplayValue = '' then
    cxGridResultView.DataController.Values[cxGridResultView.DataController.FocusedRecordIndex, cxGridResultView.Controller.FocusedColumnIndex + 1] := Null
  else
    cxGridResultView.DataController.Values[cxGridResultView.DataController.FocusedRecordIndex, cxGridResultView.Controller.FocusedColumnIndex + 1] := DisplayValue;

  cxGridResultView.DataController.BeginUpdate;
  RecalcResultRowData(cxGridResultView.DataController.FocusedRecordIndex);
  cxGridResultView.DataController.EndUpdate;
end;

procedure TShuhartForm.cxGridResultViewCustomDrawCell(
  Sender: TcxCustomGridTableView; ACanvas: TcxCanvas;
  AViewInfo: TcxGridTableDataCellViewInfo; var ADone: Boolean);
var Cellvalue: string;
begin
  if (AViewInfo.Item.Index = 9) then begin
    if (not AViewInfo.Selected) then begin
      Cellvalue := AViewInfo.GridRecord.Values[AViewInfo.Item.Index];
      if (Cellvalue = 'Погрешность неудовл.') then
        ACanvas.Canvas.Font.Color := clRed
      else
        ACanvas.Canvas.Font.Color := clBlack;
    end;
  end;
end;

procedure TShuhartForm.cxGridResultViewEditKeyDown(
  Sender: TcxCustomGridTableView; AItem: TcxCustomGridTableItem;
  AEdit: TcxCustomEdit; var Key: Word; Shift: TShiftState);
var rows, i: integer;
begin
  if ((Key = ord('V')) and (ssCtrl in Shift)) or ((Key = 45) and (ssShift in Shift)) then
  begin
    rows := InsertClipboardInGridMCI(cxGridResultView.DataController.FocusedRecordIndex, cxGridResultView.Controller.FocusedColumnIndex + 1);//добавляем данные из буфера

    cxGridResultView.DataController.BeginUpdate;
    for i := cxGridResultView.DataController.FocusedRecordIndex to cxGridResultView.DataController.FocusedRecordIndex + rows - 1 do begin
      RecalcResultRowData(i);
    end;
    cxGridResultView.DataController.EndUpdate;

    Key:=0;//зануляем нажатые клавиши, что бы исключить другую обработку этого нажатия
  end;
end;

procedure TShuhartForm.cxGridResultViewKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
var rows, i: integer;
begin
  if ((Key = ord('V')) and (ssCtrl in Shift)) or ((Key = 45) and (ssShift in Shift)) then
  begin
    rows := InsertClipboardInGridMCI(cxGridResultView.DataController.FocusedRecordIndex, cxGridResultView.Controller.FocusedColumnIndex + 1);//добавляем данные из буфера

    cxGridResultView.DataController.BeginUpdate;
    for i := cxGridResultView.DataController.FocusedRecordIndex to cxGridResultView.DataController.FocusedRecordIndex + rows - 1 do begin
      RecalcResultRowData(i);
    end;
    cxGridResultView.DataController.EndUpdate;

    Key:=0;//зануляем нажатые клавиши, что бы исключить другую обработку этого нажатия
  end;
end;

procedure TShuhartForm.cxGridViewEditKeyDown(Sender: TcxCustomGridTableView;
  AItem: TcxCustomGridTableItem; AEdit: TcxCustomEdit; var Key: Word;
  Shift: TShiftState);
var rows, i: integer;
begin
  if ((Key = ord('V')) and (ssCtrl in Shift)) or ((Key = 45) and (ssShift in Shift)) then
  begin
    rows := InsertClipboardInGrid(cxGridView.DataController.FocusedRowIndex, cxGridView.Controller.FocusedColumnIndex);//добавляем данные из буфера

    cxGridView.DataController.BeginUpdate;
    for i := cxGridView.DataController.FocusedRecordIndex to cxGridView.DataController.FocusedRecordIndex + rows - 1 do begin
      RecalcRowData(i);
    end;
    cxGridView.DataController.EndUpdate;

    Key:=0;//зануляем нажатые клавиши, что бы исключить другую обработку этого нажатия
  end;
end;

procedure TShuhartForm.DelimitedText(mmString: string; DelimitChar: char; var ResultList: TStringList);
var
  iLS, nLS: integer; //счетчик и количество символов в строке
  prVal: string;     //промежуточная строка
begin
  ResultList.Clear; //очищаем список
  nLS := Length(mmString); //узнаем длинну строки
  prVal := ''; //очищаем промежуточную переменную
  for iLS := 1 to nLS do //цикл по строке
  begin
    if mmString[iLS] = DelimitChar then //если текущий символ разделитель
    begin
      ResultList.Add(prVal);//то добавляем что ранне сложили в переменную
      prVal := '';          //и очистили
    end
    else    //иначе
    begin
      prVal := prVal + mmString[iLS];//в переменную ложим текущий символ
    end;
  end;
  if prVal <> '' then //если после прохода в переменной есть символы
    ResultList.Add(prVal);//тоже заносим их в список
end;

function TShuhartForm.InsertClipboardInGrid(row_beg: integer; col_beg: integer): integer;
var
  CCol, CRow: TStringList; //список данных со строками и столбцами
  i, j, iCntCols: Integer; //счетчик и количество столбцов в добавляемых данных
  val: Currency;
  rows: integer;
  str: string;
begin
  rows := 0;
  if Clipboard.HasFormat(CF_TEXT) then //если в буфере данные тектовые
  begin
    try
      CRow := TStringList.Create;      //создаем список строк
      CRow.Text := Clipboard.AsText;   //из буфера копируем туда все данные

      iCntCols := cxMeasureCnt.ItemIndex + 2;

      for i := 0 to CRow.Count - 1 do begin
        if ((i + row_beg) >= cxGridView.DataController.RecordCount) then break;
        try
          CCol := TStringList.Create; //создаем список столбцов
          // для этого текущую строку раскладываем в список, разделитель табуляция
          DelimitedText(CRow.Strings[i], chr(9), CCol);
          for j := 0 to CCol.Count - 1 do begin
            if ((j + col_beg) >= (1 + iCntCols)) or (j + col_beg < 1) then break;

            str := CCol[j];
            str := StringReplace(str, ',', FormatSettings.DecimalSeparator, [rfReplaceAll]);
            str := StringReplace(str, '.', FormatSettings.DecimalSeparator, [rfReplaceAll]);

            if (TryStrToCurr(str, val)) then
              cxGridView.DataController.Values[i + row_beg, j + col_beg] := val
            else
              cxGridView.DataController.Values[i + row_beg, j + col_beg] := Null;
          end;
        finally
          CCol.Free;//удаляем список по столбцам
          rows := rows + 1;
        end;
      end;
    finally
      CRow.Free;//удаляем список по строкам
    end;
  end;

  result := rows;
end;

function TShuhartForm.InsertClipboardInGridMCI(row_beg: integer; col_beg: integer): integer;
var
  CCol, CRow: TStringList; //список данных со строками и столбцами
  i, j, iCntCols: Integer; //счетчик и количество столбцов в добавляемых данных
  val: Currency;
  rows, valInt: integer;
  str: string;
begin
  rows := 0;
  if Clipboard.HasFormat(CF_TEXT) then //если в буфере данные тектовые
  begin
    try
      CRow := TStringList.Create;      //создаем список строк
      CRow.Text := Clipboard.AsText;   //из буфера копируем туда все данные

      iCntCols := 5;

      for i := 0 to CRow.Count - 1 do begin
        if ((i + row_beg) >= cxGridResultView.DataController.RecordCount) then break;
        try
          CCol := TStringList.Create; //создаем список столбцов
          // для этого текущую строку раскладываем в список, разделитель табуляция
          DelimitedText(CRow.Strings[i], chr(9), CCol);
          for j := 0 to CCol.Count - 1 do begin
            if ((j + col_beg) >= (2 + iCntCols)) or (j + col_beg < 2) then break;

            str := CCol[j];
            str := StringReplace(str, ',', FormatSettings.DecimalSeparator, [rfReplaceAll]);
            str := StringReplace(str, '.', FormatSettings.DecimalSeparator, [rfReplaceAll]);

            if ((j + col_beg) = 6) then begin
              if (TryStrToInt(str, valInt)) then
                cxGridResultView.DataController.Values[i + row_beg, j + col_beg] := valInt
              else
                cxGridResultView.DataController.Values[i + row_beg, j + col_beg] := Null;
            end
            else begin
              if (TryStrToCurr(str, val)) then
                cxGridResultView.DataController.Values[i + row_beg, j + col_beg] := val
              else
                cxGridResultView.DataController.Values[i + row_beg, j + col_beg] := Null;
            end;
          end;
        finally
          CCol.Free;//удаляем список по столбцам
          rows := rows + 1;
        end;
      end;
    finally
      CRow.Free;//удаляем список по строкам
    end;
  end;

  result := rows;
end;

procedure TShuhartForm.cxGridViewKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
var rows, i: integer;
begin
  if ((Key = ord('V')) and (ssCtrl in Shift)) or ((Key = 45) and (ssShift in Shift)) then
  begin
    rows := InsertClipboardInGrid(cxGridView.DataController.FocusedRecordIndex, cxGridView.Controller.FocusedColumnIndex);//добавляем данные из буфера

    cxGridView.DataController.BeginUpdate;
    for i := cxGridView.DataController.FocusedRecordIndex to cxGridView.DataController.FocusedRecordIndex + rows - 1 do begin
      RecalcRowData(i);
    end;
    cxGridView.DataController.EndUpdate;

    Key:=0;//зануляем нажатые клавиши, что бы исключить другую обработку этого нажатия
  end;
end;

procedure TShuhartForm.cxLaboratoryCntPropertiesEditValueChanged(
  Sender: TObject);
begin
  InitLaboratoryGrid();
end;

procedure TShuhartForm.cxLaboratoryTypePropertiesEditValueChanged(
  Sender: TObject);
begin
  cxMeasureName.Properties.Items.Clear;
  if (cxLaboratoryType.ItemIndex = 0) then begin
    cxMeasureName.Properties.Items.Add('Влагосодержание');
    cxMeasureName.Properties.Items.Add('Концентрация водорода H2');
    cxMeasureName.Properties.Items.Add('Концентрация кислорода О2');
    cxMeasureName.Properties.Items.Add('Концентрация метана СH4');
    cxMeasureName.Properties.Items.Add('Концентрация оксида углерода CO');
    cxMeasureName.Properties.Items.Add('Концентрация углекислого газа CO2');
    cxMeasureName.Properties.Items.Add('Концентрация этилена С2H4');
    cxMeasureName.Properties.Items.Add('Концентрация этана С2Н6');
    cxMeasureName.Properties.Items.Add('Концентрация ацетилена С2H2');
    cxMeasureName.Properties.Items.Add('Концентрация азота N2');
    cxMeasureName.Properties.Items.Add('Кислотное число');
    cxMeasureName.Properties.Items.Add('Температура вспышки');
    cxMeasureName.Properties.Items.Add('Пробивное напряжение');
    cxMeasureName.Properties.Items.Add('Тангенс угла диэлектрических потерь');
    cxMeasureName.Properties.Items.Add('Содержание присадок');
    cxMeasureName.Properties.Items.Add('Содержание растворенного шлама');
    cxMeasureName.Properties.Items.Add('Содержание фурановых производных (FAL)');
    cxMeasureName.Properties.Items.Add('Содержание фурановых производных (ACF)');
    cxMeasureName.Properties.Items.Add('Содержание фурановых производных (MEF)');
    cxMeasureName.Properties.Items.Add('Содержание фурановых производных (FOL)');
  end;
  if (cxLaboratoryType.ItemIndex = 1) then begin
    cxMeasureName.Properties.Items.Add('Сопротивления изоляции');
    cxMeasureName.Properties.Items.Add('Частичные разряды');
    cxMeasureName.Properties.Items.Add('Тангенс угла диэлектрических потерь');
    cxMeasureName.Properties.Items.Add('Коэффициент трансформации');
    cxMeasureName.Properties.Items.Add('Сопротивление КЗ');
  end;
  if (cxLaboratoryType.ItemIndex = 2) then begin
    cxMeasureName.Properties.Items.Add('Размер образца');
    cxMeasureName.Properties.Items.Add('Размер дефекта');
    cxMeasureName.Properties.Items.Add('Толщина металла');
    cxMeasureName.Properties.Items.Add('Твердость');
  end;

  cxMeasureName.ItemIndex := 0;

end;

{procedure TForm1.cxMaskEdit1PropertiesValidate(Sender: TObject;
  var DisplayValue: Variant; var ErrorText: TCaption; var Error: Boolean);
begin
  if (DisplayValue = '') then begin
    DisplayValue := '25';
    exit;
  end;
  if (StrToInt(DisplayValue) > 25) then DisplayValue := '25';
  if (StrToInt(DisplayValue) < 2) then DisplayValue := '2';

end;  }

procedure TShuhartForm.cxMCILaboratoryTypePropertiesEditValueChanged(
  Sender: TObject);
begin
  cxMCIMeasureName.Properties.Items.Clear;
  if (cxMCILaboratoryType.ItemIndex = 0) then begin
    cxMCIMeasureName.Properties.Items.Add('Влагосодержание');
    cxMCIMeasureName.Properties.Items.Add('Концентрация водорода H2');
    cxMCIMeasureName.Properties.Items.Add('Концентрация кислорода О2');
    cxMCIMeasureName.Properties.Items.Add('Концентрация метана СH4');
    cxMCIMeasureName.Properties.Items.Add('Концентрация оксида углерода CO');
    cxMCIMeasureName.Properties.Items.Add('Концентрация углекислого газа CO2');
    cxMCIMeasureName.Properties.Items.Add('Концентрация этилена С2H4');
    cxMCIMeasureName.Properties.Items.Add('Концентрация этана С2Н6');
    cxMCIMeasureName.Properties.Items.Add('Концентрация ацетилена С2H2');
    cxMCIMeasureName.Properties.Items.Add('Концентрация азота N2');
    cxMCIMeasureName.Properties.Items.Add('Кислотное число');
    cxMCIMeasureName.Properties.Items.Add('Температура вспышки');
    cxMCIMeasureName.Properties.Items.Add('Пробивное напряжение');
    cxMCIMeasureName.Properties.Items.Add('Тангенс угла диэлектрических потерь');
    cxMCIMeasureName.Properties.Items.Add('Содержание присадок');
    cxMCIMeasureName.Properties.Items.Add('Содержание растворенного шлама');
    cxMCIMeasureName.Properties.Items.Add('Содержание фурановых производных (FAL)');
    cxMCIMeasureName.Properties.Items.Add('Содержание фурановых производных (ACF)');
    cxMCIMeasureName.Properties.Items.Add('Содержание фурановых производных (MEF)');
    cxMCIMeasureName.Properties.Items.Add('Содержание фурановых производных (FOL)');
  end;
  if (cxMCILaboratoryType.ItemIndex = 1) then begin
    cxMCIMeasureName.Properties.Items.Add('Сопротивления изоляции');
    cxMCIMeasureName.Properties.Items.Add('Частичные разряды');
    cxMCIMeasureName.Properties.Items.Add('Тангенс угла диэлектрических потерь');
    cxMCIMeasureName.Properties.Items.Add('Коэффициент трансформации');
    cxMCIMeasureName.Properties.Items.Add('Сопротивление КЗ');
  end;
  if (cxMCILaboratoryType.ItemIndex = 2) then begin
    cxMCIMeasureName.Properties.Items.Add('Размер образца');
    cxMCIMeasureName.Properties.Items.Add('Размер дефекта');
    cxMCIMeasureName.Properties.Items.Add('Толщина металла');
    cxMCIMeasureName.Properties.Items.Add('Твердость');
  end;

  cxMCIMeasureName.ItemIndex := 0;
end;

procedure TShuhartForm.cxMCIMeasureNamePropertiesEditValueChanged(
  Sender: TObject);
begin
  if (cxMCILaboratoryType.ItemIndex = 0) then begin
    case cxMCIMeasureName.ItemIndex of
      -1:cxMCIUnitName.Text := '';
      0:cxMCIUnitName.Text := 'г/т';
      1:cxMCIUnitName.Text := 'мкл/л';
      2:cxMCIUnitName.Text := 'мкл/л';
      3:cxMCIUnitName.Text := 'мкл/л';
      4:cxMCIUnitName.Text := 'мкл/л';
      5:cxMCIUnitName.Text := 'мкл/л';
      6:cxMCIUnitName.Text := 'мкл/л';
      7:cxMCIUnitName.Text := 'мкл/л';
      8:cxMCIUnitName.Text := 'мкл/л';
      9:cxMCIUnitName.Text := 'мкл/л';
      10:cxMCIUnitName.Text := 'мг КОН/г';
      11:cxMCIUnitName.Text := '°C';
      12:cxMCIUnitName.Text := 'кВ';
      13:cxMCIUnitName.Text := '%';
      14:cxMCIUnitName.Text := '%';
      15:cxMCIUnitName.Text := '%';
      16:cxMCIUnitName.Text := 'мг/кг';
      17:cxMCIUnitName.Text := 'мг/кг';
      18:cxMCIUnitName.Text := 'мг/кг';
      19:cxMCIUnitName.Text := 'мг/кг';
    end;
  end;
  if (cxMCILaboratoryType.ItemIndex = 1) then begin
    case cxMCIMeasureName.ItemIndex of
      -1:cxMCIUnitName.Text := '';
      0:cxMCIUnitName.Text := 'МОм';
      1:cxMCIUnitName.Text := '';
      2:cxMCIUnitName.Text := '%';
      3:cxMCIUnitName.Text := '';
      4:cxMCIUnitName.Text := 'Ом';
    end;
  end;
  if (cxMCILaboratoryType.ItemIndex = 2) then begin
    case cxMCIMeasureName.ItemIndex of
      -1:cxMCIUnitName.Text := '';
      0:cxMCIUnitName.Text := 'мм';
      1:cxMCIUnitName.Text := 'мм';
      2:cxMCIUnitName.Text := 'мм';
      3:cxMCIUnitName.Text := '';
    end;
  end;
end;

procedure TShuhartForm.cxMCIUnitNamePropertiesEditValueChanged(Sender: TObject);
begin
  cxGridResultView.BeginUpdate();

  cxGridResultView.Columns[2].Caption := 'Аттестованное значение ОК (СО), ' + cxMCIUnitName.Text;
  cxGridResultView.Columns[3].Caption := 'Погрешность аттестованного ОК, ' + cxMCIUnitName.Text;
  cxGridResultView.Columns[4].Caption := 'Результат испытаний ОК в ИЛ, ' + cxMCIUnitName.Text;
  cxGridResultView.Columns[5].Caption := 'Погрешность результата испытаний в ИЛ, ' + cxMCIUnitName.Text;

  cxGridResultView.EndUpdate();
end;

procedure TShuhartForm.cxMeasureCntFocusChanged(Sender: TObject);
var
 i: integer;
begin
  if (not cxMeasureCnt.IsFocused) then begin
    InitGrid();

    cxGridView.DataController.BeginUpdate;
    for i := 0 to cxGridView.DataController.RecordCount - 1 do begin
      RecalcRowData(i);
    end;
    cxGridView.DataController.EndUpdate;
  end;
end;

procedure TShuhartForm.cxMeasureNamePropertiesEditValueChanged(Sender: TObject);
begin
  if (cxLaboratoryType.ItemIndex = 0) then begin
    case cxMeasureName.ItemIndex of
      -1:cxUnitName.Text := '';
      0:cxUnitName.Text := 'г/т';
      1:cxUnitName.Text := 'мкл/л';
      2:cxUnitName.Text := 'мкл/л';
      3:cxUnitName.Text := 'мкл/л';
      4:cxUnitName.Text := 'мкл/л';
      5:cxUnitName.Text := 'мкл/л';
      6:cxUnitName.Text := 'мкл/л';
      7:cxUnitName.Text := 'мкл/л';
      8:cxUnitName.Text := 'мкл/л';
      9:cxUnitName.Text := 'мкл/л';
      10:cxUnitName.Text := 'мг КОН/г';
      11:cxUnitName.Text := '°C';
      12:cxUnitName.Text := 'кВ';
      13:cxUnitName.Text := '%';
      14:cxUnitName.Text := '%';
      15:cxUnitName.Text := '%';
      16:cxUnitName.Text := 'мг/кг';
      17:cxUnitName.Text := 'мг/кг';
      18:cxUnitName.Text := 'мг/кг';
      19:cxUnitName.Text := 'мг/кг';
    end;
  end;
  if (cxLaboratoryType.ItemIndex = 1) then begin
    case cxMeasureName.ItemIndex of
      -1:cxUnitName.Text := '';
      0:cxUnitName.Text := 'МОм';
      1:cxUnitName.Text := '';
      2:cxUnitName.Text := '%';
      3:cxUnitName.Text := '';
      4:cxUnitName.Text := 'Ом';
    end;
  end;
  if (cxLaboratoryType.ItemIndex = 2) then begin
    case cxMeasureName.ItemIndex of
      -1:cxUnitName.Text := '';
      0:cxUnitName.Text := 'мм';
      1:cxUnitName.Text := 'мм';
      2:cxUnitName.Text := 'мм';
      3:cxUnitName.Text := '';
    end;
  end;

  if (cxUnitName.Text <> '') then begin
    cxGridView.Bands[1].Caption := cxMeasureName.Text + ', ' + cxUnitName.Text;
  end
  else begin
    cxGridView.Bands[1].Caption := cxMeasureName.Text;
  end;

end;

procedure TShuhartForm.cxPageControl2Change(Sender: TObject);
begin
  if (cxPageControl2.ActivePageIndex >= 2) then begin
    CalcDataKoef();
  end;

  if (cxPageControl2.ActivePageIndex = 2) then begin
    DrawXCard();
  end;

  if (cxPageControl2.ActivePageIndex = 3) then begin
    DrawSCard();
  end;

  if (cxPageControl2.ActivePageIndex = 4) then begin
    DrawMeCard();
  end;

  if (cxPageControl2.ActivePageIndex = 5) then begin
    DrawRCard();
  end;

  if (cxPageControl2.ActivePageIndex = 6) then begin
    CalcVerification();
    DrawConvChart();
  end;

  if (cxPageControl2.ActivePageIndex = 8) then begin
    CalcVerification();
    FillVerification();
  end;

end;

function TShuhartForm.VerifyDoubleLaboratory(): boolean;
var i, j: integer;
begin
  for i := 0 to cxGridLaboratoryView.DataController.RecordCount - 1 do begin
    if ( cxGridLaboratoryView.DataController.Values[i, 1] = null) then begin
      result := false;
      exit;
    end;
    for j := i + 1 to cxGridLaboratoryView.DataController.RecordCount - 1 do begin
      if (cxGridLaboratoryView.DataController.Values[i, 1] = cxGridLaboratoryView.DataController.Values[j, 1]) then begin
        result := false;
        exit;
      end;

    end;
  end;

  result := true;
end;

procedure TShuhartForm.cxPageControl3Change(Sender: TObject);
begin
    if (cxPageControl3.ActivePageIndex > 0) then begin
      if (not VerifyDoubleLaboratory()) then begin
        cxPageControl3.ActivePageIndex := 0;
        MessageDLG('В списке лабораторий не должно быть лабораторий с повторяющимися или пустыми названиями.',mtError,[mbOk], 0);
        exit;
      end;

      InitResultMCIGrid(true);
    end;
end;

procedure TShuhartForm.cxR0FocusChanged(Sender: TObject);
begin
  if (not cxR0.IsFocused) then begin
    if (cxR0.Text = '') then cxR0.Value := 0;
  end;
end;

procedure TShuhartForm.cxR2ContrFocusChanged(Sender: TObject);
begin
  if (not cxR2Contr.IsFocused) then begin
    if (cxR2Contr.Text <> '') and (StrToCurr(cxR2Contr.Text) > 100) then cxR2Contr.Value := 100;
  end;
end;

procedure TShuhartForm.cxRContrFocusChanged(Sender: TObject);
begin
  if (not cxRContr.IsFocused) then begin
    if (cxRContr.Text <> '') and (StrToCurr(cxRContr.Text) > 100) then cxRContr.Value := 100;
  end;
end;

procedure TShuhartForm.cxRowCntFocusChanged(Sender: TObject);
var i: integer;
begin
  if (not cxRowCnt.IsFocused) then begin
    if (cxRowCnt.Text = '') or (StrToInt(cxRowCnt.Text) < 2) then cxRowCnt.Text := '2';
    if (StrToInt(cxRowCnt.Text) > 200) then cxRowCnt.Text := '200';

    cxGridView.DataController.BeginUpdate;
    cxGridView.DataController.RecordCount := StrToInt(cxRowCnt.Text);
    cxGridView.DataController.EndUpdate;

    InitGrid();
  end;
end;

procedure TShuhartForm.cxS0FocusChanged(Sender: TObject);
begin
  if (not cxS0.IsFocused) then begin
    if (cxS0.Text = '') then cxS0.Value := 0;
  end;
end;

procedure TShuhartForm.cxS0PropertiesChange(Sender: TObject);
begin
  if (cxS0.Text = '') then cxS0.Value := 0;
end;

procedure TShuhartForm.cxSigma0FocusChanged(Sender: TObject);
begin
  if (not cxSigma0.IsFocused) then begin
    if (cxSigma0.Text = '') then cxSigma0.Value := 0;
  end;
end;

procedure TShuhartForm.cxStandardValuesClick(Sender: TObject);
begin
  EnableStandardValues();
end;

procedure TShuhartForm.cxUnitNamePropertiesEditValueChanged(Sender: TObject);
begin
    lUnit.Caption := cxUnitName.Text;
    lUnit2.Caption := cxUnitName.Text;
    lUnit3.Caption := cxUnitName.Text;
    lUnit4.Caption := cxUnitName.Text;
    lUnit5.Caption := cxUnitName.Text;
    lUnit6.Caption := cxUnitName.Text;
    lUnit7.Caption := cxUnitName.Text;
    lUnit8.Caption := cxUnitName.Text;
    lUnit9.Caption := cxUnitName.Text;
    lUnit10.Caption := cxUnitName.Text;
    lUnit11.Caption := cxUnitName.Text;

    if (cxUnitName.Text <> '') then begin
      cxGridView.Bands[1].Caption := cxMeasureName.Text + ', ' + cxUnitName.Text;
      cxGridView.Bands[2].Caption := 'Среднее значение, ' + cxUnitName.Text;
      cxGridView.Bands[3].Caption := 'СКО, ' + cxUnitName.Text;
      cxGridView.Bands[4].Caption := 'Медиана, ' + cxUnitName.Text;
      cxGridView.Bands[5].Caption := 'Размах, ' + cxUnitName.Text;
    end
    else begin
      cxGridView.Bands[1].Caption := cxMeasureName.Text;
      cxGridView.Bands[2].Caption := 'Среднее значение';
      cxGridView.Bands[3].Caption := 'СКО';
      cxGridView.Bands[4].Caption := 'Медиана';
      cxGridView.Bands[5].Caption := 'Размах';
    end;
end;

procedure TShuhartForm.cxX0FocusChanged(Sender: TObject);
begin
  if (not cxX0.IsFocused) then begin
    if (cxX0.Text = '') then cxX0.Value := 0;
  end;
end;

procedure TShuhartForm.cxCFocusChanged(Sender: TObject);
begin
  if (not cxC.IsFocused) then begin
    if (cxC.Text = '') then cxC.Value := 0;
  end;
end;

procedure TShuhartForm.cxChartFontSizePropertiesEditValueChanged(
  Sender: TObject);
begin
  X_Chart.LeftAxis.LabelsFormat.Font.Size := cxChartFontSize.ItemIndex + 7;
  X_Chart.BottomAxis.LabelsFormat.Font.Size := cxChartFontSize.ItemIndex + 7;
  X_Chart.Update;

  S_Chart.LeftAxis.LabelsFormat.Font.Size := cxChartFontSize.ItemIndex + 7;
  S_Chart.BottomAxis.LabelsFormat.Font.Size := cxChartFontSize.ItemIndex + 7;
  S_Chart.Update;

  R_Chart.LeftAxis.LabelsFormat.Font.Size := cxChartFontSize.ItemIndex + 7;
  R_Chart.BottomAxis.LabelsFormat.Font.Size := cxChartFontSize.ItemIndex + 7;
  R_Chart.Update;

  Me_Chart.LeftAxis.LabelsFormat.Font.Size := cxChartFontSize.ItemIndex + 7;
  Me_Chart.BottomAxis.LabelsFormat.Font.Size := cxChartFontSize.ItemIndex + 7;
  Me_Chart.Update;

  SKO_Chart.LeftAxis.LabelsFormat.Font.Size := cxChartFontSize.ItemIndex + 7;
  SKO_Chart.BottomAxis.LabelsFormat.Font.Size := cxChartFontSize.ItemIndex + 7;
  SKO_Chart.Update;
end;

procedure TShuhartForm.EnableStandardValues();
begin
  if (cxStandardValues.Checked) then begin
    cxX0.Enabled := true;
    cxR0.Enabled := true;
    cxS0.Enabled := true;
    cxSigma0.Enabled := true;
    cxStandardValuesGroup.Enabled := true;

    {cxX0.Value := 0;
    cxR0.Value := 0;
    cxS0.Value := 0;
    cxSigma0.Value := 0;}
  end
  else begin
    cxX0.Enabled := false;
    cxR0.Enabled := false;
    cxS0.Enabled := false;
    cxSigma0.Enabled := false;
    cxStandardValuesGroup.Enabled := false;

    {cxX0.Text := '';
    cxR0.Text := '';
    cxS0.Text := '';
    cxSigma0.Text := '';}
  end;

end;

procedure TShuhartForm.InitResultMCIGrid(bRecalcData: boolean);
var i, j, iRowIndex, iOldCntRows : integer;
  cVal, cProc: Currency;
  arr: array of TMCIDataInfo;
  strLaboratoryName: string;
  strOKName: string;
  di: TMCIDataInfo;
begin
  cxGridResultView.BeginUpdate();

  // записываем существующие данные
  iOldCntRows :=  cxGridResultView.DataController.RecordCount;
  if (iOldCntRows <> cxLaboratoryCnt.ItemIndex * (cxOKCnt.ItemIndex + 1)) then begin

    SetLength(arr, iOldCntRows);
    for i := 0 to iOldCntRows - 1 do begin
      di.strLaboratoryName := cxGridResultView.DataController.Values[i, 0];
      di.strOKName := cxGridResultView.DataController.Values[i, 1];
      di.val1 := cxGridResultView.DataController.Values[i, 2];
      di.val2 := cxGridResultView.DataController.Values[i, 3];
      di.val3 := cxGridResultView.DataController.Values[i, 4];
      di.val4 := cxGridResultView.DataController.Values[i, 5];
      di.val5 := cxGridResultView.DataController.Values[i, 6];

      arr[i] := di;
    end;

    cxGridResultView.DataController.RecordCount := 0;
  end;


  cxGridResultView.DataController.RecordCount := cxLaboratoryCnt.ItemIndex * (cxOKCnt.ItemIndex + 1);
  iRowIndex := 0;
  for i := 0 to cxLaboratoryCnt.ItemIndex - 1 do begin
    for j := 0 to cxOKCnt.ItemIndex do begin
      cxGridResultView.DataController.Values[iRowIndex, 0] := cxGridLaboratoryView.DataController.Values[i, 1];
      cxGridResultView.DataController.Values[iRowIndex, 1] := 'OK № ' + IntToStr(j + 1);
      if (cxGridResultView.DataController.Values[iRowIndex, 6] = Null) or (cxGridResultView.DataController.Values[iRowIndex, 6] = '') then
        cxGridResultView.DataController.Values[iRowIndex, 6] := '10';

      iRowIndex := iRowIndex + 1;
    end;
  end;

  // восстанавливаем существующие данные
  if (iOldCntRows <> cxLaboratoryCnt.ItemIndex * (cxOKCnt.ItemIndex + 1)) then begin
    for i := 0 to cxGridResultView.DataController.RecordCount - 1 do begin
      strLaboratoryName := cxGridResultView.DataController.Values[i, 0];
      strOKName := cxGridResultView.DataController.Values[i, 1];
      for j := 0 to Length(arr) - 1 do begin
        if (arr[j].strLaboratoryName = strLaboratoryName) and (arr[j].strOKName = strOKName) then begin
          cxGridResultView.DataController.Values[i, 2] := arr[j].val1;
          cxGridResultView.DataController.Values[i, 3] := arr[j].val2;
          cxGridResultView.DataController.Values[i, 4] := arr[j].val3;
          cxGridResultView.DataController.Values[i, 5] := arr[j].val4;
          cxGridResultView.DataController.Values[i, 6] := arr[j].val5;
          break;
        end;

      end;
    end;
  end;

  if (bRecalcData) then begin
    for i := 0 to cxGridResultView.DataController.RecordCount - 1 do begin
      RecalcResultRowData(i);
    end;
  end;

  cxGridResultView.EndUpdate();

  cxGridResultView.DataController.Groups.FullExpand;
  
end;

procedure TShuhartForm.InitLaboratoryGrid();
var i: integer;
begin
  cxGridLaboratoryView.BeginUpdate();
  cxGridLaboratoryView.DataController.RecordCount := cxLaboratoryCnt.ItemIndex;
  for i := 0 to cxLaboratoryCnt.ItemIndex - 1 do begin
    cxGridLaboratoryView.DataController.Values[i, 0] := IntToStr(i + 1);    
    if (cxGridLaboratoryView.DataController.Values[i, 1] = Null) or (cxGridLaboratoryView.DataController.Values[i, 1] = '') then begin
      cxGridLaboratoryView.DataController.Values[i, 1] := 'ИЛ-участница № ' + IntToStr(i + 1);
    end;
  end;
  cxGridLaboratoryView.EndUpdate();
end;

procedure TShuhartForm.InitGrid();
var
  i, cols: Integer;
begin
  cxGridView.BeginUpdate();

  for i := 0 to cxGridView.DataController.RecordCount - 1 do begin
    cxGridView.DataController.Values[i, 0] := i + 1;
  end;

  if (cxCardType.ItemIndex = 0) then begin // количественные данные
    cxColumnX1.Caption := 'X1';

    cols := cxMeasureCnt.ItemIndex + 2;
    for i := 2 to cols do begin
      cxGridView.Columns[i].Visible := true;
    end;
    for i := cols + 1 to 20 do begin
      cxGridView.Columns[i].Visible := false;
    end;

    if (cxUnitName.Text <> '') then
      cxGridView.Bands[1].Caption := cxMeasureName.Text + ', ' + cxUnitName.Text
    else
      cxGridView.Bands[1].Caption := cxMeasureName.Text;
    cxGridView.Bands[1].Width := 0;
    cxGridView.Bands[2].Visible := true;
    cxGridView.Bands[3].Visible := true;
    cxGridView.Bands[4].Visible := true;
    if (cxUnitName.Text <> '') then
      cxGridView.Bands[5].Caption := 'Размах' + ', ' + cxUnitName.Text
    else
      cxGridView.Bands[5].Caption := 'Размах';
    cxGridView.Bands[5].Width := 74;
    cxGridView.Bands[6].Visible := true;
    cxTabSheet3.Caption := 'Карта средних';
    cxTabSheet4.TabVisible := true;
    cxTabSheet5.TabVisible := true;
    cxTabSheet6.Caption := 'Карта размахов';
    cxTabSheet10.TabVisible := true;
    cxTabSheet11.TabVisible := true;

    lUnit.Caption := cxUnitName.Text;
    lUnit2.Caption := cxUnitName.Text;
    lUnit3.Caption := cxUnitName.Text;
    lUnit4.Caption := cxUnitName.Text;
    lUnit5.Caption := cxUnitName.Text;
    lUnit6.Caption := cxUnitName.Text;
    lUnit7.Caption := cxUnitName.Text;
    lUnit8.Caption := cxUnitName.Text;
    lUnit9.Caption := cxUnitName.Text;
    lUnit10.Caption := cxUnitName.Text;
    lUnit11.Caption := cxUnitName.Text;

    cxGroupBoxC.Visible := true;

    cxAnalyzeType.Enabled := true;
    if (cxAnalyzeType.ItemIndex = 0) then begin
      cxTabSheet3.TabVisible := true;
      cxTabSheet4.TabVisible := false;
      cxTabSheet5.TabVisible := false;
      cxGridView.Bands[4].Visible := false;
    end;
    if (cxAnalyzeType.ItemIndex = 1) then begin
      cxTabSheet3.TabVisible := false;
      cxTabSheet4.TabVisible := true;
      cxTabSheet5.TabVisible := false;
      cxGridView.Bands[4].Visible := false;
    end;
    if (cxAnalyzeType.ItemIndex = 2) then begin
      cxTabSheet3.TabVisible := false;
      cxTabSheet4.TabVisible := false;
      cxTabSheet5.TabVisible := true;
      cxGridView.Bands[4].Visible := true;

      cxTabSheet7.TabVisible := true;
      if (cxMeasureCnt.ItemIndex >= Length(A_Me_koef)) then begin {коэффициенты для медиан только от 2 до 10}
        cxGridView.Bands[4].Visible := false;
        cxTabSheet5.TabVisible := false;

        if (cxAnalyzeType.ItemIndex = 2) then
          cxTabSheet7.TabVisible := false;
      end
    end;

  end
  else begin // индивидуальные данные
    cxColumnX1.Caption := 'X';

    for i := 2 to 20 do begin
      cxGridView.Columns[i].Visible := false;
    end;

    if (cxUnitName.Text <> '') then
      cxGridView.Bands[1].Caption := cxMeasureName.Text + ', ' + cxUnitName.Text
    else
      cxGridView.Bands[1].Caption := cxMeasureName.Text;
    cxGridView.Bands[1].Width := 150;
    cxGridView.Bands[2].Visible := false;
    cxGridView.Bands[3].Visible := false;
    cxGridView.Bands[4].Visible := false;
    if (cxUnitName.Text <> '') then
      cxGridView.Bands[5].Caption := 'Скользящий размах' + ', ' + cxUnitName.Text
    else
      cxGridView.Bands[5].Caption := 'Скользящий размах';
    cxGridView.Bands[5].Width := 150;
    cxGridView.Bands[6].Visible := false;
    cxTabSheet3.Caption := 'Карта индивидуальных значений';
    cxTabSheet4.TabVisible := false;
    cxTabSheet5.TabVisible := false;
    cxTabSheet10.TabVisible := false;
    cxTabSheet11.TabVisible := false;
    cxTabSheet6.Caption := 'Карта скользящих размахов';

    cxGroupBoxC.Visible := false;

    cxAnalyzeType.ItemIndex := 0;
    cxAnalyzeType.Enabled := false;
  end;

  for i := 1 to cxGridView.ColumnCount - 1 do begin
    if (cxGridView.Columns[i].Name = 'cxColumn1') then continue;

    case cxDigitCnt.ItemIndex of
    0:
      begin
        cxCurrencyEdit.TcxCustomCurrencyEditProperties(cxGridView.Columns[i].Properties).DecimalPlaces := 1;
        cxCurrencyEdit.TcxCustomCurrencyEditProperties(cxGridView.Columns[i].Properties).DisplayFormat := ',0.0;-,0.0';
      end;
    1:
      begin
        cxCurrencyEdit.TcxCustomCurrencyEditProperties(cxGridView.Columns[i].Properties).DecimalPlaces := 2;
        cxCurrencyEdit.TcxCustomCurrencyEditProperties(cxGridView.Columns[i].Properties).DisplayFormat := ',0.00;-,0.00';
      end;
    2:
      begin
        cxCurrencyEdit.TcxCustomCurrencyEditProperties(cxGridView.Columns[i].Properties).DecimalPlaces := 3;
        cxCurrencyEdit.TcxCustomCurrencyEditProperties(cxGridView.Columns[i].Properties).DisplayFormat := ',0.000;-,0.000';
      end;
    3:
      begin
        cxCurrencyEdit.TcxCustomCurrencyEditProperties(cxGridView.Columns[i].Properties).DecimalPlaces := 4;
        cxCurrencyEdit.TcxCustomCurrencyEditProperties(cxGridView.Columns[i].Properties).DisplayFormat := ',0.0000;-,0.0000';
      end;

    end;
  end;

  cxGridView.EndUpdate();
end;

procedure TShuhartForm.InitDataMCI(bInitGrid: boolean);
begin
  cxPageControl3.ActivePageIndex := 0;

  cxLaboratoryCnt.ItemIndex := 0;
  cxOKCnt.ItemIndex := 0;

  cxMCILaboratoryType.ItemIndex := 0;
  cxMCIMeasureName.Text := '';
  cxMCIUnitName.Text := '';
  cxMCIMeasureMethod.Text := '';
  cxMCIDateBegin.Text := '';
  cxMCIDateEnd.Text := '';
  cxMCIModel.ItemIndex := 0;

  cxGridLaboratoryView.DataController.BeginUpdate;
  cxGridLaboratoryView.DataController.RecordCount := 0;
  cxGridLaboratoryView.DataController.EndUpdate;

  cxGridResultView.DataController.BeginUpdate;
  cxGridResultView.DataController.RecordCount := 0;
  cxGridResultView.DataController.EndUpdate;

end;

procedure TShuhartForm.InitData(bInitGrid: boolean);
var i, j: integer;
begin
  //cxMainTab.ActivePageIndex := 0;
  cxPageControl1.ActivePageIndex := 0;
  cxPageControl2.ActivePageIndex := 0;

  cxCardType.ItemIndex := 0;
  cxAnalyzeType.ItemIndex := 0;
  cxDigitCnt.ItemIndex := 3;
  cxMeasureCnt.ItemIndex := 0;
  cxRowCnt.Text := '20';
  cxChartFontSize.ItemIndex := 1;

  cxCardName.Text := '';
  cxLaboratoryType.ItemIndex := 0;
  cxMeasureName.Text := '';
  cxUnitName.Text := '';
  cxMeasureMethod.Text := '';
  cxDateBegin.Text := '';
  cxDateEnd.Text := '';

  cxGridView.DataController.BeginUpdate;

  cxGridView.DataController.RecordCount := 20;

  for i := 1 to cxGridView.ColumnCount - 1 do begin
    for j := 0 to cxGridView.DataController.RecordCount - 1 do begin
      cxGridView.DataController.Values[j, i] := Null;
    end;
  end;

   cxGridView.DataController.EndUpdate;

  EnableStandardValues();


  cxX0.Value := 0;
  cxR0.Value := 0;
  cxS0.Value := 0;
  cxSigma0.Value := 0;

  cxRContr.Value := 5;
  cxR2Contr.Value := 10;

  if (bInitGrid) then InitGrid();
end;

procedure TShuhartForm.mLoadClick(Sender: TObject);
var
  node, node2, node3, node_main : IXmlNode;
  i, j, iCnt: integer;
  str: string;
  val: variant;
begin
  DlgOpen.Title := 'Загрузить данные';
  DlgOpen.InitialDir := GetCurrentDir;
  DlgOpen.Filter := 'TI-EES file|*.ti-ees';
  DlgOpen.DefaultExt := 'ti-ees';

  if (DlgOpen.Execute()) then begin

    InitData(false);
    InitDataMCI(false);

    cxGridView.BeginUpdate();
    cxGridLaboratoryView.BeginUpdate();
    cxGridResultView.BeginUpdate();

    try
      XMLDoc.Active := true;
      XMLDoc.LoadFromFile(DlgOpen.FileName);

      if (XMLDoc.ChildNodes.Count <> 1) then begin
        MessageDLG('Некорректный формат данных.',mtError,[mbOk], 0);
        exit;
      end;

      node_main := XMLDoc.ChildNodes.Get(0);
      {if (node_main.ChildNodes.Count <> 2) then begin
        ShowMessage('Некорректный формат данных.');
        exit;
      end;}

      node := node_main.ChildNodes.Get(0);

      {if (node.NodeName <> 'inner_lab') then begin
        MessageDLG('Некорректный формат данных.',mtError,[mbOk], 0);
        exit;
      end;

      if (node.ChildNodes.Count <> 2) then begin
        MessageDLG('Некорректный формат данных.',mtError,[mbOk], 0);
        exit;
      end;}

      node2 := node.ChildNodes.Get(0);
      {if (node2.NodeName <> 'parameters') then begin
        MessageDLG('Некорректный формат данных.',mtError,[mbOk], 0);
        exit;
      end; }

      cxCardName.Text := node2.GetAttributeNS('card_name', '');
      if (node2.HasAttribute('laboratory_type')) then
        cxLaboratoryType.ItemIndex := node2.GetAttributeNS('laboratory_type', '')
      else
        cxLaboratoryType.ItemIndex := 0;
      cxMeasureName.Text := node2.GetAttributeNS('measure_name', '');
      cxUnitName.Text := node2.GetAttributeNS('unit_name', '');
      cxCardType.ItemIndex := node2.GetAttributeNS('card_type', '');

      if (node2.HasAttribute('measure_method')) then
        cxMeasureMethod.Text := node2.GetAttributeNS('measure_method', '')
      else
        cxMeasureMethod.Text := '';

      if (node2.HasAttribute('measure_date_begin')) then
        cxDateBegin.Text := node2.GetAttributeNS('measure_date_begin', '')
      else
        cxDateBegin.Text := '';

      if (node2.HasAttribute('measure_date_end')) then
        cxDateEnd.Text := node2.GetAttributeNS('measure_date_end', '')
      else
        cxDateEnd.Text := '';

      if (node2.HasAttribute('analyze_type')) then
        cxAnalyzeType.ItemIndex := node2.GetAttributeNS('analyze_type', '')
      else
        cxAnalyzeType.ItemIndex := 0;

      cxDigitCnt.ItemIndex := node2.GetAttributeNS('digit_cnt', '');
      cxMeasureCnt.ItemIndex := node2.GetAttributeNS('measure_cnt', '');
      if (node2.HasAttribute('row_cnt')) then begin
        cxRowCnt.Text := node2.GetAttributeNS('row_cnt', '');
        if (cxRowCnt.Text = '') or (StrToInt(cxRowCnt.Text) < 2) then cxRowCnt.Text := '2';
      end
      else
        cxRowCnt.Text := '20';
      if (node2.HasAttribute('axis_font_size')) then
        cxChartFontSize.ItemIndex := node2.GetAttributeNS('axis_font_size', '')
      else
        cxChartFontSize.ItemIndex := 1;
      cxStandardValues.Checked := node2.GetAttributeNS('use_standard_values', '');
      cxX0.Value := node2.GetAttributeNS('X0', '');
      cxR0.Value := node2.GetAttributeNS('R0', '');
      cxS0.Value := node2.GetAttributeNS('S0', '');
      cxSigma0.Value := node2.GetAttributeNS('Sigma0', '');
      if (cxStandardValues.Checked) then begin
        cxX0.Enabled := true;
        cxR0.Enabled := true;
        cxS0.Enabled := true;
        cxSigma0.Enabled := true;
      end
      else begin
        cxX0.Enabled := false;
        cxR0.Enabled := false;
        cxS0.Enabled := false;
        cxSigma0.Enabled := false;
      end;
      if (node2.HasAttribute('r_contr')) then
        cxRContr.Value := node2.GetAttributeNS('r_contr', '')
      else
        cxRContr.Value := 5;
      if (node2.HasAttribute('r2_contr')) then
        cxR2Contr.Value := node2.GetAttributeNS('r2_contr', '')
      else
        cxR2Contr.Value := 10;

      if (node2.HasAttribute('att_value')) then
        cxC.Value := node2.GetAttributeNS('att_value', '')
      else
        cxC.Value := 0;

      if (node2.HasAttribute('att_value_delta')) then
        cxCDelta.Value := node2.GetAttributeNS('att_value_delta', '')
      else
        cxCDelta.Value := 0;

      node2 := node.ChildNodes.Get(1);
      {if (node2.NodeName <> 'data') then begin
        MessageDLG('Некорректный формат данных.',mtError,[mbOk], 0);
        exit;
      end;}

      iCnt := cxMeasureCnt.ItemIndex + 2;

      cxGridView.DataController.RecordCount := StrToInt(cxRowCnt.Text);

      for i := 0 to cxGridView.DataController.RecordCount - 1 do begin
        if (i >= node2.ChildNodes.Count) then break;
        node3 := node2.ChildNodes.Get(i);

        for j := 0 to iCnt - 1 do begin
          val := node3.GetAttributeNS('X' + IntToStr(j + 1), '');
          if (val = '') then
            cxGridView.DataController.Values[i, 1 + j] := Null
          else
            cxGridView.DataController.Values[i, 1 + j] := val;
        end;

        RecalcRowData(i);
      end;

      if (node_main.ChildNodes.Count >= 2) then begin

        node := node_main.ChildNodes.Get(1);

        {if (node.NodeName <> 'outer_lab') then begin
          MessageDLG('Некорректный формат данных.',mtError,[mbOk], 0);
          exit;
        end;

        if (node.ChildNodes.Count <> 3) then begin
          MessageDLG('Некорректный формат данных.',mtError,[mbOk], 0);
          exit;
        end; }

        node2 := node.ChildNodes.Get(0);
        {if (node2.NodeName <> 'parameters') then begin
          MessageDLG('Некорректный формат данных.',mtError,[mbOk], 0);
          exit;
        end;}

        cxLaboratoryCnt.ItemIndex := node2.GetAttributeNS('laboratory_cnt', '');
        cxOKCnt.ItemIndex := node2.GetAttributeNS('OK_cnt', '');
        if (node2.HasAttribute('unit_name')) then
          cxMCIUnitName.Text := node2.GetAttributeNS('unit_name', '')
        else
          cxMCIUnitName.Text := '';

        if (node2.HasAttribute('laboratory_type')) then
          cxMCILaboratoryType.ItemIndex := node2.GetAttributeNS('laboratory_type', '')
        else
          cxMCILaboratoryType.ItemIndex := 0;

        if (node2.HasAttribute('model_program')) then
          cxMCIModel.ItemIndex := node2.GetAttributeNS('model_program', '')
        else
          cxMCIModel.ItemIndex := 0;

        if (node2.HasAttribute('measure_name')) then
          cxMCIMeasureName.Text := node2.GetAttributeNS('measure_name', '')
        else
          cxMCIMeasureName.Text := '';

        if (node2.HasAttribute('measure_method')) then
          cxMCIMeasureMethod.Text := node2.GetAttributeNS('measure_method', '')
        else
          cxMCIMeasureMethod.Text := '';

        if (node2.HasAttribute('measure_date_begin')) then
          cxMCIDateBegin.Text := node2.GetAttributeNS('measure_date_begin', '')
        else
          cxMCIDateBegin.Text := '';

        if (node2.HasAttribute('measure_date_end')) then
          cxMCIDateEnd.Text := node2.GetAttributeNS('measure_date_end', '')
        else
          cxMCIDateEnd.Text := '';

        if (node2.HasAttribute('descript')) then
          cxMCIDescript.Text := node2.GetAttributeNS('descript', '')
        else
          cxMCIDescript.Text := '';

        node2 := node.ChildNodes.Get(1);
        {if (node2.NodeName <> 'laboratory_data') then begin
          MessageDLG('Некорректный формат данных.',mtError,[mbOk], 0);
          exit;
        end;}

        cxGridLaboratoryView.DataController.RecordCount := cxLaboratoryCnt.ItemIndex;

        for i := 0 to cxGridLaboratoryView.DataController.RecordCount - 1 do begin
          if (i >= node2.ChildNodes.Count) then break;
          node3 := node2.ChildNodes.Get(i);

          val := node3.GetAttributeNS('name', '');
          cxGridLaboratoryView.DataController.Values[i, 1] := val;
        end;

        InitResultMCIGrid(false);

        node2 := node.ChildNodes.Get(2);
        {if (node2.NodeName <> 'data') then begin
          MessageDLG('Некорректный формат данных.',mtError,[mbOk], 0);
          exit;
        end; }

        for i := 0 to cxGridResultView.DataController.RecordCount - 1 do begin
          if (i >= node2.ChildNodes.Count) then break;
          node3 := node2.ChildNodes.Get(i);

          val := node3.GetAttributeNS('val2', '');
          if (val = '') then
            cxGridResultView.DataController.Values[i, 2] := Null
          else
            cxGridResultView.DataController.Values[i, 2] := val;

          val := node3.GetAttributeNS('val3', '');
          if (val = '') then
            cxGridResultView.DataController.Values[i, 3] := Null
          else
            cxGridResultView.DataController.Values[i, 3] := val;

          val := node3.GetAttributeNS('val4', '');
          if (val = '') then
            cxGridResultView.DataController.Values[i, 4] := Null
          else
            cxGridResultView.DataController.Values[i, 4] := val;

          val := node3.GetAttributeNS('val5', '');
          if (val = '') then
            cxGridResultView.DataController.Values[i, 5] := Null
          else
            cxGridResultView.DataController.Values[i, 5] := val;

          val := node3.GetAttributeNS('val6', '');
          if (val = '') then
            cxGridResultView.DataController.Values[i, 6] := Null
          else
            cxGridResultView.DataController.Values[i, 6] := val;

          RecalcResultRowData(i);
        end;
      end;

      InitGrid();
    except
      MessageDLG('Ошибка при открытии файла.',mtError,[mbOk], 0);
    end;

    cxGridView.EndUpdate();
    cxGridLaboratoryView.EndUpdate();
    cxGridResultView.EndUpdate();
  end;
end;

procedure TShuhartForm.mNewClick(Sender: TObject);
begin
  InitData(true);
  InitDataMCI(true);
end;

procedure TShuhartForm.mSaveClick(Sender: TObject);
var
  node, node2, node3, node_main : IXmlNode;
  i, j, iCnt: integer;
begin
  DlgSave.Title := 'Сохранить данные';
  DlgSave.InitialDir := GetCurrentDir;
  DlgSave.Filter := 'TI-EES file|*.ti-ees';
  DlgSave.DefaultExt := 'ti-ees';

  if (DlgSave.Execute()) then begin
    XMLDoc.Active := true;
    XMLDoc.ChildNodes.Clear;
    node_main := XMLDoc.AddChild('main');
    node := node_main.AddChild('inner_lab');
    node2 := node.AddChild('parameters');
    node2.SetAttributeNS('card_name', '', cxCardName.Text);
    node2.SetAttributeNS('laboratory_type', '', cxLaboratoryType.ItemIndex);
    node2.SetAttributeNS('measure_name', '', cxMeasureName.Text);
    node2.SetAttributeNS('unit_name', '', cxUnitName.Text);
    node2.SetAttributeNS('card_type', '', cxCardType.ItemIndex);
    node2.SetAttributeNS('measure_method', '', cxMeasureMethod.Text);
    node2.SetAttributeNS('measure_date_begin', '', cxDateBegin.Text);
    node2.SetAttributeNS('measure_date_end', '', cxDateEnd.Text);
    node2.SetAttributeNS('analyze_type', '', cxAnalyzeType.ItemIndex);
    node2.SetAttributeNS('digit_cnt', '', cxDigitCnt.ItemIndex);
    node2.SetAttributeNS('measure_cnt', '', cxMeasureCnt.ItemIndex);
    node2.SetAttributeNS('row_cnt', '', cxRowCnt.Text);
    node2.SetAttributeNS('axis_font_size', '', cxChartFontSize.ItemIndex);
    node2.SetAttributeNS('use_standard_values', '', cxStandardValues.Checked);
    node2.SetAttributeNS('X0', '', cxX0.Value);
    node2.SetAttributeNS('R0', '', cxR0.Value);
    node2.SetAttributeNS('S0', '', cxS0.Value);
    node2.SetAttributeNS('Sigma0', '', cxSigma0.Value);
    node2.SetAttributeNS('att_value', '', cxC.Value);
    node2.SetAttributeNS('att_value_delta', '', cxCDelta.Value);
    node2.SetAttributeNS('r_contr', '', cxRContr.Value);
    node2.SetAttributeNS('r2_contr', '', cxR2Contr.Value);

    node2 := node.AddChild('data');

    iCnt := cxMeasureCnt.ItemIndex + 2;
    for i := 0 to cxGridView.DataController.RecordCount - 1 do begin

      node3 := node2.AddChild('row_' + IntToStr(i));

      for j := 0 to iCnt - 1 do begin
        if (cxGridView.DataController.Values[i, 1 + j] <> Null) then
          node3.SetAttributeNS('X' + IntToStr(j + 1), '', cxGridView.DataController.Values[i, 1 + j])
        else
          node3.SetAttributeNS('X' + IntToStr(j + 1), '', '');
      end;
    end;

    node := node_main.AddChild('outer_lab');
    node2 := node.AddChild('parameters');
    node2.SetAttributeNS('laboratory_cnt', '', cxLaboratoryCnt.ItemIndex);
    node2.SetAttributeNS('laboratory_type', '', cxMCILaboratoryType.ItemIndex);
    node2.SetAttributeNS('measure_name', '', cxMCIMeasureName.Text);
    node2.SetAttributeNS('measure_method', '', cxMCIMeasureMethod.Text);
    node2.SetAttributeNS('model_program', '', cxMCIModel.ItemIndex);
    node2.SetAttributeNS('measure_date_begin', '', cxMCIDateBegin.Text);
    node2.SetAttributeNS('measure_date_end', '', cxMCIDateEnd.Text);
    node2.SetAttributeNS('descript', '', cxMCIDescript.Text);
    node2.SetAttributeNS('OK_cnt', '', cxOKCnt.ItemIndex);
    node2.SetAttributeNS('unit_name', '', cxMCIUnitName.Text);
    node2 := node.AddChild('laboratory_data');
    for i := 0 to cxGridLaboratoryView.DataController.RecordCount - 1 do begin

      node3 := node2.AddChild('row_' + IntToStr(i));
      if (cxGridLaboratoryView.DataController.Values[i, 1] <> Null) then
        node3.SetAttributeNS('name', '', cxGridLaboratoryView.DataController.Values[i, 1])
      else
        node3.SetAttributeNS('name', '', '')
    end;

    node2 := node.AddChild('data');
    for i := 0 to cxGridResultView.DataController.RecordCount - 1 do begin

      node3 := node2.AddChild('row_' + IntToStr(i));
      for j := 2 to 6 do begin
        if (cxGridResultView.DataController.Values[i, 1] <> Null) then
          node3.SetAttributeNS('val' +  IntToStr(j), '', cxGridResultView.DataController.Values[i, j])
        else
          node3.SetAttributeNS('val' +  IntToStr(j), '', '')
      end;
    end;

    XMLDoc.SaveToFile(DlgSave.FileName);
    MessageDLG('Данные успешно сохранены.',mtConfirmation,[mbOk], 0);
  end;
end;

procedure TShuhartForm.N3Click(Sender: TObject);
begin
  frAbout.ShowModal();
end;

procedure TShuhartForm.FormActivate(Sender: TObject);
var res: TModalResult;
begin
  fLogin := TfLogin.Create(Self);
  res := fLogin.ShowModal;
  fLogin.Free;

  if (res <> mrOk) then begin
    halt;
    exit;
  end;
end;

procedure TShuhartForm.FormCreate(Sender: TObject);
var i: integer;
begin

  {fLogin := TfLogin.Create(Self);
  fLogin.ShowModal;
  fLogin.Free;

  Self.Visible := false;
  halt;
  exit;}

  // n: 2-25
  SetLength(A_koef, 3);
  for i := 0 to Length(A_koef) - 1 do SetLength(A_koef[i], 25 - 1);
  SetLength(B_koef, 4);
  for i := 0 to Length(B_koef) - 1 do SetLength(B_koef[i], 25 - 1);
  SetLength(D_koef, 4);
  for i := 0 to Length(D_koef) - 1 do SetLength(D_koef[i], 25 - 1);
  SetLength(C_koef_middle, 2);
  for i := 0 to Length(C_koef_middle) - 1 do SetLength(C_koef_middle[i], 25 - 1);
  SetLength(d_koef_middle, 2);
  for i := 0 to Length(d_koef_middle) - 1 do SetLength(d_koef_middle[i], 25 - 1);

  // n: 2-10
  SetLength(A_Me_koef, 10 - 1);

  // n: 2-10
  SetLength(Q_koef, 10 - 1);

  // n-1: 1-5, L: 2-40
  SetLength(Kohren_koef, 5);
  for i := 0 to Length(Kohren_koef) - 1 do SetLength(Kohren_koef[i], 40 - 1);

  // L-1: 1-120
  SetLength(Student_koef, 120);

  // n: 2-5
  SetLength(SCO_koef, 3);
  for i := 0 to Length(SCO_koef) - 1 do SetLength(SCO_koef[i], 5 - 1);

  A_koef[0, 0] := 2.121; A_koef[1, 0] := 1.880; A_koef[2, 0] := 2.659; B_koef[0, 0] := 0.000; B_koef[1, 0] := 3.267; B_koef[2, 0] := 0.000; B_koef[3, 0] := 2.606; D_koef[0, 0] := 0.000; D_koef[1, 0] := 3.686; D_koef[2, 0] := 0.000; D_koef[3, 0] := 3.267; C_koef_middle[0, 0] := 0.7979; C_koef_middle[1, 0] := 1.2533; d_koef_middle[0, 0] := 1.128; d_koef_middle[1, 0] := 0.8865;
  A_koef[0, 1] := 1.732; A_koef[1, 1] := 1.023; A_koef[2, 1] := 1.954; B_koef[0, 1] := 0.000; B_koef[1, 1] := 2.568; B_koef[2, 1] := 0.000; B_koef[3, 1] := 2.276; D_koef[0, 1] := 0.000; D_koef[1, 1] := 4.358; D_koef[2, 1] := 0.000; D_koef[3, 1] := 2.574; C_koef_middle[0, 1] := 0.8886; C_koef_middle[1, 1] := 1.1284; d_koef_middle[0, 1] := 1.693; d_koef_middle[1, 1] := 0.5907;
  A_koef[0, 2] := 1.500; A_koef[1, 2] := 0.729; A_koef[2, 2] := 1.628; B_koef[0, 2] := 0.000; B_koef[1, 2] := 2.266; B_koef[2, 2] := 0.000; B_koef[3, 2] := 2.088; D_koef[0, 2] := 0.000; D_koef[1, 2] := 4.696; D_koef[2, 2] := 0.000; D_koef[3, 2] := 2.282; C_koef_middle[0, 2] := 0.9213; C_koef_middle[1, 2] := 1.0854; d_koef_middle[0, 2] := 2.059; d_koef_middle[1, 2] := 0.4857;
  A_koef[0, 3] := 1.342; A_koef[1, 3] := 0.577; A_koef[2, 3] := 1.427; B_koef[0, 3] := 0.000; B_koef[1, 3] := 2.089; B_koef[2, 3] := 0.000; B_koef[3, 3] := 1.964; D_koef[0, 3] := 0.000; D_koef[1, 3] := 4.918; D_koef[2, 3] := 0.000; D_koef[3, 3] := 2.114; C_koef_middle[0, 3] := 0.9400; C_koef_middle[1, 3] := 1.0638; d_koef_middle[0, 3] := 2.326; d_koef_middle[1, 3] := 0.4299;
  A_koef[0, 4] := 1.225; A_koef[1, 4] := 0.483; A_koef[2, 4] := 1.287; B_koef[0, 4] := 0.030; B_koef[1, 4] := 1.970; B_koef[2, 4] := 0.029; B_koef[3, 4] := 1.874; D_koef[0, 4] := 0.000; D_koef[1, 4] := 5.078; D_koef[2, 4] := 0.000; D_koef[3, 4] := 2.004; C_koef_middle[0, 4] := 0.9515; C_koef_middle[1, 4] := 1.0510; d_koef_middle[0, 4] := 2.534; d_koef_middle[1, 4] := 0.3946;
  A_koef[0, 5] := 1.134; A_koef[1, 5] := 0.419; A_koef[2, 5] := 1.182; B_koef[0, 5] := 0.118; B_koef[1, 5] := 1.882; B_koef[2, 5] := 0.113; B_koef[3, 5] := 1.806; D_koef[0, 5] := 0.204; D_koef[1, 5] := 5.204; D_koef[2, 5] := 0.076; D_koef[3, 5] := 1.924; C_koef_middle[0, 5] := 0.9594; C_koef_middle[1, 5] := 1.0423; d_koef_middle[0, 5] := 2.704; d_koef_middle[1, 5] := 0.3698;
  A_koef[0, 6] := 1.061; A_koef[1, 6] := 0.373; A_koef[2, 6] := 1.099; B_koef[0, 6] := 0.185; B_koef[1, 6] := 1.815; B_koef[2, 6] := 0.179; B_koef[3, 6] := 1.751; D_koef[0, 6] := 0.388; D_koef[1, 6] := 5.306; D_koef[2, 6] := 0.136; D_koef[3, 6] := 1.864; C_koef_middle[0, 6] := 0.9650; C_koef_middle[1, 6] := 1.0363; d_koef_middle[0, 6] := 2.847; d_koef_middle[1, 6] := 0.3512;
  A_koef[0, 7] := 1.000; A_koef[1, 7] := 0.337; A_koef[2, 7] := 1.032; B_koef[0, 7] := 0.239; B_koef[1, 7] := 1.761; B_koef[2, 7] := 0.232; B_koef[3, 7] := 1.707; D_koef[0, 7] := 0.547; D_koef[1, 7] := 5.393; D_koef[2, 7] := 0.184; D_koef[3, 7] := 1.816; C_koef_middle[0, 7] := 0.9693; C_koef_middle[1, 7] := 1.0317; d_koef_middle[0, 7] := 2.970; d_koef_middle[1, 7] := 0.3367;
  A_koef[0, 8] := 0.949; A_koef[1, 8] := 0.308; A_koef[2, 8] := 0.975; B_koef[0, 8] := 0.284; B_koef[1, 8] := 1.716; B_koef[2, 8] := 0.276; B_koef[3, 8] := 1.669; D_koef[0, 8] := 0.687; D_koef[1, 8] := 5.469; D_koef[2, 8] := 0.223; D_koef[3, 8] := 1.777; C_koef_middle[0, 8] := 0.9727; C_koef_middle[1, 8] := 1.0281; d_koef_middle[0, 8] := 3.078; d_koef_middle[1, 8] := 0.3249;
  A_koef[0, 9] := 0.905; A_koef[1, 9] := 0.285; A_koef[2, 9] := 0.927; B_koef[0, 9] := 0.321; B_koef[1, 9] := 1.679; B_koef[2, 9] := 0.313; B_koef[3, 9] := 1.637; D_koef[0, 9] := 0.811; D_koef[1, 9] := 5.535; D_koef[2, 9] := 0.256; D_koef[3, 9] := 1.744; C_koef_middle[0, 9] := 0.9754; C_koef_middle[1, 9] := 1.0252; d_koef_middle[0, 9] := 3.173; d_koef_middle[1, 9] := 0.3152;
  A_koef[0, 10] := 0.866; A_koef[1, 10] := 0.266; A_koef[2, 10] := 0.886; B_koef[0, 10] := 0.354; B_koef[1, 10] := 1.646; B_koef[2, 10] := 0.346; B_koef[3, 10] := 1.610; D_koef[0, 10] := 0.922; D_koef[1, 10] := 5.594; D_koef[2, 10] := 0.283; D_koef[3, 10] := 1.717; C_koef_middle[0, 10] := 0.9776; C_koef_middle[1, 10] := 1.0229; d_koef_middle[0, 10] := 3.258; d_koef_middle[1, 10] := 0.3069;
  A_koef[0, 11] := 0.832; A_koef[1, 11] := 0.249; A_koef[2, 11] := 0.850; B_koef[0, 11] := 0.382; B_koef[1, 11] := 1.618; B_koef[2, 11] := 0.374; B_koef[3, 11] := 1.585; D_koef[0, 11] := 1.025; D_koef[1, 11] := 5.647; D_koef[2, 11] := 0.307; D_koef[3, 11] := 1.693; C_koef_middle[0, 11] := 0.9794; C_koef_middle[1, 11] := 1.0210; d_koef_middle[0, 11] := 3.336; d_koef_middle[1, 11] := 0.2998;
  A_koef[0, 12] := 0.802; A_koef[1, 12] := 0.235; A_koef[2, 12] := 0.817; B_koef[0, 12] := 0.406; B_koef[1, 12] := 1.594; B_koef[2, 12] := 0.399; B_koef[3, 12] := 1.563; D_koef[0, 12] := 1.118; D_koef[1, 12] := 5.696; D_koef[2, 12] := 0.328; D_koef[3, 12] := 1.672; C_koef_middle[0, 12] := 0.9810; C_koef_middle[1, 12] := 1.0194; d_koef_middle[0, 12] := 3.407; d_koef_middle[1, 12] := 0.2935;
  A_koef[0, 13] := 0.775; A_koef[1, 13] := 0.223; A_koef[2, 13] := 0.789; B_koef[0, 13] := 0.428; B_koef[1, 13] := 1.572; B_koef[2, 13] := 0.421; B_koef[3, 13] := 1.544; D_koef[0, 13] := 1.203; D_koef[1, 13] := 5.741; D_koef[2, 13] := 0.347; D_koef[3, 13] := 1.653; C_koef_middle[0, 13] := 0.9823; C_koef_middle[1, 13] := 1.0180; d_koef_middle[0, 13] := 3.472; d_koef_middle[1, 13] := 0.2880;
  A_koef[0, 14] := 0.750; A_koef[1, 14] := 0.212; A_koef[2, 14] := 0.763; B_koef[0, 14] := 0.448; B_koef[1, 14] := 1.552; B_koef[2, 14] := 0.440; B_koef[3, 14] := 1.526; D_koef[0, 14] := 1.282; D_koef[1, 14] := 5.782; D_koef[2, 14] := 0.363; D_koef[3, 14] := 1.637; C_koef_middle[0, 14] := 0.9835; C_koef_middle[1, 14] := 1.0168; d_koef_middle[0, 14] := 3.532; d_koef_middle[1, 14] := 0.2831;
  A_koef[0, 15] := 0.728; A_koef[1, 15] := 0.203; A_koef[2, 15] := 0.739; B_koef[0, 15] := 0.466; B_koef[1, 15] := 1.534; B_koef[2, 15] := 0.458; B_koef[3, 15] := 1.511; D_koef[0, 15] := 1.356; D_koef[1, 15] := 5.820; D_koef[2, 15] := 0.378; D_koef[3, 15] := 1.622; C_koef_middle[0, 15] := 0.9845; C_koef_middle[1, 15] := 1.0157; d_koef_middle[0, 15] := 3.588; d_koef_middle[1, 15] := 0.2784;
  A_koef[0, 16] := 0.707; A_koef[1, 16] := 0.194; A_koef[2, 16] := 0.718; B_koef[0, 16] := 0.482; B_koef[1, 16] := 1.518; B_koef[2, 16] := 0.475; B_koef[3, 16] := 1.496; D_koef[0, 16] := 1.424; D_koef[1, 16] := 5.856; D_koef[2, 16] := 0.391; D_koef[3, 16] := 1.608; C_koef_middle[0, 16] := 0.9854; C_koef_middle[1, 16] := 1.0148; d_koef_middle[0, 16] := 3.640; d_koef_middle[1, 16] := 0.2747;
  A_koef[0, 17] := 0.688; A_koef[1, 17] := 0.187; A_koef[2, 17] := 0.698; B_koef[0, 17] := 0.497; B_koef[1, 17] := 1.503; B_koef[2, 17] := 0.490; B_koef[3, 17] := 1.483; D_koef[0, 17] := 1.487; D_koef[1, 17] := 5.891; D_koef[2, 17] := 0.403; D_koef[3, 17] := 1.597; C_koef_middle[0, 17] := 0.9862; C_koef_middle[1, 17] := 1.0140; d_koef_middle[0, 17] := 3.689; d_koef_middle[1, 17] := 0.2711;
  A_koef[0, 18] := 0.671; A_koef[1, 18] := 0.180; A_koef[2, 18] := 0.680; B_koef[0, 18] := 0.510; B_koef[1, 18] := 1.490; B_koef[2, 18] := 0.504; B_koef[3, 18] := 1.470; D_koef[0, 18] := 1.549; D_koef[1, 18] := 5.921; D_koef[2, 18] := 0.415; D_koef[3, 18] := 1.585; C_koef_middle[0, 18] := 0.9869; C_koef_middle[1, 18] := 1.0133; d_koef_middle[0, 18] := 3.735; d_koef_middle[1, 18] := 0.2677;
  A_koef[0, 19] := 0.655; A_koef[1, 19] := 0.173; A_koef[2, 19] := 0.663; B_koef[0, 19] := 0.523; B_koef[1, 19] := 1.477; B_koef[2, 19] := 0.516; B_koef[3, 19] := 1.459; D_koef[0, 19] := 1.605; D_koef[1, 19] := 5.951; D_koef[2, 19] := 0.425; D_koef[3, 19] := 1.575; C_koef_middle[0, 19] := 0.9876; C_koef_middle[1, 19] := 1.0126; d_koef_middle[0, 19] := 3.778; d_koef_middle[1, 19] := 0.2647;
  A_koef[0, 20] := 0.640; A_koef[1, 20] := 0.167; A_koef[2, 20] := 0.647; B_koef[0, 20] := 0.534; B_koef[1, 20] := 1.466; B_koef[2, 20] := 0.528; B_koef[3, 20] := 1.448; D_koef[0, 20] := 1.659; D_koef[1, 20] := 5.979; D_koef[2, 20] := 0.434; D_koef[3, 20] := 1.566; C_koef_middle[0, 20] := 0.9882; C_koef_middle[1, 20] := 1.0119; d_koef_middle[0, 20] := 3.819; d_koef_middle[1, 20] := 0.2618;
  A_koef[0, 21] := 0.626; A_koef[1, 21] := 0.162; A_koef[2, 21] := 0.633; B_koef[0, 21] := 0.545; B_koef[1, 21] := 1.455; B_koef[2, 21] := 0.539; B_koef[3, 21] := 1.438; D_koef[0, 21] := 1.710; D_koef[1, 21] := 6.006; D_koef[2, 21] := 0.443; D_koef[3, 21] := 1.557; C_koef_middle[0, 21] := 0.9887; C_koef_middle[1, 21] := 1.0114; d_koef_middle[0, 21] := 3.858; d_koef_middle[1, 21] := 0.2592;
  A_koef[0, 22] := 0.612; A_koef[1, 22] := 0.157; A_koef[2, 22] := 0.619; B_koef[0, 22] := 0.555; B_koef[1, 22] := 1.445; B_koef[2, 22] := 0.549; B_koef[3, 22] := 1.429; D_koef[0, 22] := 1.759; D_koef[1, 22] := 6.031; D_koef[2, 22] := 0.451; D_koef[3, 22] := 1.548; C_koef_middle[0, 22] := 0.9892; C_koef_middle[1, 22] := 1.0109; d_koef_middle[0, 22] := 3.895; d_koef_middle[1, 22] := 0.2567;
  A_koef[0, 23] := 0.600; A_koef[1, 23] := 0.153; A_koef[2, 23] := 0.606; B_koef[0, 23] := 0.565; B_koef[1, 23] := 1.434; B_koef[2, 23] := 0.559; B_koef[3, 23] := 1.420; D_koef[0, 23] := 1.806; D_koef[1, 23] := 6.056; D_koef[2, 23] := 0.459; D_koef[3, 23] := 1.541; C_koef_middle[0, 23] := 0.9896; C_koef_middle[1, 23] := 1.0105; d_koef_middle[0, 23] := 3.931; d_koef_middle[1, 23] := 0.2544;

  A_Me_koef[0] := 1.88;
  A_Me_koef[1] := 1.19;
  A_Me_koef[2] := 0.80;
  A_Me_koef[3] := 0.69;
  A_Me_koef[4] := 0.55;
  A_Me_koef[5] := 0.51;
  A_Me_koef[6] := 0.43;
  A_Me_koef[7] := 0.41;
  A_Me_koef[8] := 0.36;

  Q_koef[0] := 2.77;
  Q_koef[1] := 3.31;
  Q_koef[2] := 3.63;
  Q_koef[3] := 3.86;
  Q_koef[4] := 4.03;
  Q_koef[5] := 4.17;
  Q_koef[6] := 4.29;
  Q_koef[7] := 4.39;
  Q_koef[8] := 4.47;

  // [n - 1, L], n0 = 1, L0 = 2
  Kohren_koef[0, 0] := 0.999; Kohren_koef[1, 0] := 0.975; Kohren_koef[2, 0] := 0.939; Kohren_koef[3, 0] := 0.906; Kohren_koef[4, 0] := 0.877;
  Kohren_koef[0, 1] := 0.967; Kohren_koef[1, 1] := 0.871; Kohren_koef[2, 1] := 0.798; Kohren_koef[3, 1] := 0.746; Kohren_koef[4, 1] := 0.707;
  Kohren_koef[0, 2] := 0.906; Kohren_koef[1, 2] := 0.768; Kohren_koef[2, 2] := 0.684; Kohren_koef[3, 2] := 0.629; Kohren_koef[4, 2] := 0.590;
  Kohren_koef[0, 3] := 0.841; Kohren_koef[1, 3] := 0.684; Kohren_koef[2, 3] := 0.598; Kohren_koef[3, 3] := 0.544; Kohren_koef[4, 3] := 0.506;
  Kohren_koef[0, 4] := 0.781; Kohren_koef[1, 4] := 0.616; Kohren_koef[2, 4] := 0.532; Kohren_koef[3, 4] := 0.480; Kohren_koef[4, 4] := 0.445;
  Kohren_koef[0, 5] := 0.727; Kohren_koef[1, 5] := 0.561; Kohren_koef[2, 5] := 0.480; Kohren_koef[3, 5] := 0.431; Kohren_koef[4, 5] := 0.397;
  Kohren_koef[0, 6] := 0.680; Kohren_koef[1, 6] := 0.516; Kohren_koef[2, 6] := 0.438; Kohren_koef[3, 6] := 0.391; Kohren_koef[4, 6] := 0.360;
  Kohren_koef[0, 7] := 0.638; Kohren_koef[1, 7] := 0.478; Kohren_koef[2, 7] := 0.403; Kohren_koef[3, 7] := 0.358; Kohren_koef[4, 7] := 0.329;
  Kohren_koef[0, 8] := 0.602; Kohren_koef[1, 8] := 0.445; Kohren_koef[2, 8] := 0.373; Kohren_koef[3, 8] := 0.331; Kohren_koef[4, 8] := 0.303;
  Kohren_koef[0, 9] := 0.570; Kohren_koef[1, 9] := 0.417; Kohren_koef[2, 9] := 0.348; Kohren_koef[3, 9] := 0.308; Kohren_koef[4, 9] := 0.281;
  Kohren_koef[0, 10] := 0.541; Kohren_koef[1, 10] := 0.392; Kohren_koef[2, 10] := 0.326; Kohren_koef[3, 10] := 0.288; Kohren_koef[4, 10] := 0.262;
  Kohren_koef[0, 11] := 0.515; Kohren_koef[1, 11] := 0.371; Kohren_koef[2, 11] := 0.307; Kohren_koef[3, 11] := 0.271; Kohren_koef[4, 11] := 0.243;
  Kohren_koef[0, 12] := 0.492; Kohren_koef[1, 12] := 0.352; Kohren_koef[2, 12] := 0.291; Kohren_koef[3, 12] := 0.255; Kohren_koef[4, 12] := 0.232;
  Kohren_koef[0, 13] := 0.471; Kohren_koef[1, 13] := 0.335; Kohren_koef[2, 13] := 0.276; Kohren_koef[3, 13] := 0.242; Kohren_koef[4, 13] := 0.220;
  Kohren_koef[0, 14] := 0.452; Kohren_koef[1, 14] := 0.319; Kohren_koef[2, 14] := 0.262; Kohren_koef[3, 14] := 0.230; Kohren_koef[4, 14] := 0.208;
  Kohren_koef[0, 15] := 0.434; Kohren_koef[1, 15] := 0.305; Kohren_koef[2, 15] := 0.250; Kohren_koef[3, 15] := 0.219; Kohren_koef[4, 15] := 0.198;
  Kohren_koef[0, 16] := 0.418; Kohren_koef[1, 16] := 0.293; Kohren_koef[2, 16] := 0.240; Kohren_koef[3, 16] := 0.209; Kohren_koef[4, 16] := 0.189;
  Kohren_koef[0, 17] := 0.403; Kohren_koef[1, 17] := 0.281; Kohren_koef[2, 17] := 0.230; Kohren_koef[3, 17] := 0.200; Kohren_koef[4, 17] := 0.181;
  Kohren_koef[0, 18] := 0.389; Kohren_koef[1, 18] := 0.270; Kohren_koef[2, 18] := 0.220; Kohren_koef[3, 18] := 0.192; Kohren_koef[4, 18] := 0.174;
  Kohren_koef[0, 19] := 0.377; Kohren_koef[1, 19] := 0.261; Kohren_koef[2, 19] := 0.212; Kohren_koef[3, 19] := 0.185; Kohren_koef[4, 19] := 0.167;
  Kohren_koef[0, 20] := 0.365; Kohren_koef[1, 20] := 0.252; Kohren_koef[2, 20] := 0.204; Kohren_koef[3, 20] := 0.178; Kohren_koef[4, 20] := 0.160;
  Kohren_koef[0, 21] := 0.354; Kohren_koef[1, 21] := 0.243; Kohren_koef[2, 21] := 0.197; Kohren_koef[3, 21] := 0.172; Kohren_koef[4, 21] := 0.155;
  Kohren_koef[0, 22] := 0.343; Kohren_koef[1, 22] := 0.235; Kohren_koef[2, 22] := 0.191; Kohren_koef[3, 22] := 0.166; Kohren_koef[4, 22] := 0.149;
  Kohren_koef[0, 23] := 0.334; Kohren_koef[1, 23] := 0.228; Kohren_koef[2, 23] := 0.185; Kohren_koef[3, 23] := 0.160; Kohren_koef[4, 23] := 0.144;
  Kohren_koef[0, 24] := 0.325; Kohren_koef[1, 24] := 0.221; Kohren_koef[2, 24] := 0.179; Kohren_koef[3, 24] := 0.155; Kohren_koef[4, 24] := 0.140;
  Kohren_koef[0, 25] := 0.316; Kohren_koef[1, 25] := 0.215; Kohren_koef[2, 25] := 0.173; Kohren_koef[3, 25] := 0.150; Kohren_koef[4, 25] := 0.135;
  Kohren_koef[0, 26] := 0.308; Kohren_koef[1, 26] := 0.209; Kohren_koef[2, 26] := 0.168; Kohren_koef[3, 26] := 0.146; Kohren_koef[4, 26] := 0.131;
  Kohren_koef[0, 27] := 0.300; Kohren_koef[1, 27] := 0.203; Kohren_koef[2, 27] := 0.164; Kohren_koef[3, 27] := 0.142; Kohren_koef[4, 27] := 0.127;
  Kohren_koef[0, 28] := 0.293; Kohren_koef[1, 28] := 0.198; Kohren_koef[2, 28] := 0.159; Kohren_koef[3, 28] := 0.138; Kohren_koef[4, 28] := 0.124;
  Kohren_koef[0, 29] := 0.286; Kohren_koef[1, 29] := 0.193; Kohren_koef[2, 29] := 0.155; Kohren_koef[3, 29] := 0.134; Kohren_koef[4, 29] := 0.120;
  Kohren_koef[0, 30] := 0.280; Kohren_koef[1, 30] := 0.188; Kohren_koef[2, 30] := 0.151; Kohren_koef[3, 30] := 0.131; Kohren_koef[4, 30] := 0.117;
  Kohren_koef[0, 31] := 0.273; Kohren_koef[1, 31] := 0.184; Kohren_koef[2, 31] := 0.147; Kohren_koef[3, 31] := 0.127; Kohren_koef[4, 31] := 0.114;
  Kohren_koef[0, 32] := 0.267; Kohren_koef[1, 32] := 0.179; Kohren_koef[2, 32] := 0.144; Kohren_koef[3, 32] := 0.124; Kohren_koef[4, 32] := 0.111;
  Kohren_koef[0, 33] := 0.262; Kohren_koef[1, 33] := 0.175; Kohren_koef[2, 33] := 0.140; Kohren_koef[3, 33] := 0.121; Kohren_koef[4, 33] := 0.108;
  Kohren_koef[0, 34] := 0.256; Kohren_koef[1, 34] := 0.172; Kohren_koef[2, 34] := 0.137; Kohren_koef[3, 34] := 0.118; Kohren_koef[4, 34] := 0.106;
  Kohren_koef[0, 35] := 0.251; Kohren_koef[1, 35] := 0.168; Kohren_koef[2, 35] := 0.134; Kohren_koef[3, 35] := 0.116; Kohren_koef[4, 35] := 0.103;
  Kohren_koef[0, 36] := 0.246; Kohren_koef[1, 36] := 0.164; Kohren_koef[2, 36] := 0.131; Kohren_koef[3, 36] := 0.113; Kohren_koef[4, 36] := 0.101;
  Kohren_koef[0, 37] := 0.242; Kohren_koef[1, 37] := 0.161; Kohren_koef[2, 37] := 0.129; Kohren_koef[3, 37] := 0.111; Kohren_koef[4, 37] := 0.099;
  Kohren_koef[0, 38] := 0.237; Kohren_koef[1, 38] := 0.158; Kohren_koef[2, 38] := 0.126; Kohren_koef[3, 38] := 0.108; Kohren_koef[4, 38] := 0.097;

  // [L - 1], L0 = 1
  Student_koef[0] := 2.71;
  Student_koef[1] := 2.30;
  Student_koef[2] := 2.18;
  Student_koef[3] := 2.78;
  Student_koef[4] := 2.57;
  Student_koef[5] := 2.45;
  Student_koef[6] := 2.37;
  Student_koef[7] := 2.31;
  Student_koef[8] := 2.26;
  Student_koef[9] := 2.23;
  Student_koef[10] := 2.20;
  Student_koef[11] := 2.18;
  Student_koef[12] := 2.16;
  Student_koef[13] := 2.15;
  Student_koef[14] := 2.14;
  Student_koef[15] := 2.12;
  Student_koef[16] := 2.11;
  Student_koef[17] := 2.10;
  Student_koef[18] := 2.09;
  Student_koef[19] := 2.09;
  Student_koef[20] := 2.08;
  Student_koef[21] := 2.07;
  Student_koef[22] := 2.07;
  Student_koef[23] := 2.06;
  Student_koef[24] := 2.06;
  Student_koef[25] := 2.06;
  Student_koef[26] := 2.05;
  Student_koef[27] := 2.05;
  Student_koef[28] := 2.04;
  Student_koef[29] := 2.04;
  for i := 30 to 38 do Student_koef[i] := 2.04;
  for i := 39 to 58 do Student_koef[i] := 2.02;
  for i := 59 to 118 do Student_koef[i] := 2.00;
  Student_koef[119] := 1.98;

  SCO_koef[0, 0] := 1.128; SCO_koef[1, 0] := 2.834; SCO_koef[2, 0] := 3.686;
  SCO_koef[0, 1] := 1.693; SCO_koef[1, 1] := 3.469; SCO_koef[2, 1] := 4.358;
  SCO_koef[0, 2] := 2.059; SCO_koef[1, 2] := 3.819; SCO_koef[2, 2] := 4.698;
  SCO_koef[0, 3] := 2.326; SCO_koef[1, 3] := 4.054; SCO_koef[2, 3] := 4.918;

  SetLength(arrCriteries, 8);
  for i := 0 to 7 do arrCriteries[i] := TList.Create;

  cxGridCriteriaView.DataController.RecordCount := 8;
  cxGridCriteriaView.DataController.Values[0, 0] := 1;
  cxGridCriteriaView.DataController.Values[1, 0] := 2;
  cxGridCriteriaView.DataController.Values[2, 0] := 3;
  cxGridCriteriaView.DataController.Values[3, 0] := 4;
  cxGridCriteriaView.DataController.Values[4, 0] := 5;
  cxGridCriteriaView.DataController.Values[5, 0] := 6;
  cxGridCriteriaView.DataController.Values[6, 0] := 7;
  cxGridCriteriaView.DataController.Values[7, 0] := 8;

  cxGridCriteriaView.DataController.Values[0, 1] := 'Одна точка вне зоны A';
  cxGridCriteriaView.DataController.Values[1, 1] := 'Девять точек подряд в зоне C или по одну сторону от центральной линии';
  cxGridCriteriaView.DataController.Values[2, 1] := 'Шесть возрастающих или убывающих точек подряд';
  cxGridCriteriaView.DataController.Values[3, 1] := 'Четырнадцать попеременно возрастающих и убывающих точек';
  cxGridCriteriaView.DataController.Values[4, 1] := 'Две из трех последовательных точек в зоне A или вне ее';
  cxGridCriteriaView.DataController.Values[5, 1] := 'Четыре из пяти последовательных точек в зоне B или вне ее';
  cxGridCriteriaView.DataController.Values[6, 1] := 'Пятнадцать последовательных точек в зоне C выше и ниже центральной линии';
  cxGridCriteriaView.DataController.Values[7, 1] := 'Восемь последовательных точек по обеим сторонам центральной линии и ни одной в зоне C';

  cxGridCriteriaView.DataController.Values[0, 2] := 'Критический';
  cxGridCriteriaView.DataController.Values[1, 2] := 'Критический';
  cxGridCriteriaView.DataController.Values[2, 2] := 'Критический';
  cxGridCriteriaView.DataController.Values[3, 2] := 'Предупреждающий';
  cxGridCriteriaView.DataController.Values[4, 2] := 'Предупреждающий';
  cxGridCriteriaView.DataController.Values[5, 2] := 'Предупреждающий';
  cxGridCriteriaView.DataController.Values[6, 2] := 'Критический';
  cxGridCriteriaView.DataController.Values[7, 2] := 'Предупреждающий';

  dataMeasureName := TStringList.Create;
  dataMeasureName.Add(Format('%s=%s', ['Влагосодержание', 'влагосодержания']));
  dataMeasureName.Add(Format('%s=%s', ['Концентрация водорода H2', 'концентрации водорода H2']));
  dataMeasureName.Add(Format('%s=%s', ['Концентрация кислорода О2', 'концентрации кислорода О2']));
  dataMeasureName.Add(Format('%s=%s', ['Концентрация метана СH4', 'концентрации метана СH4']));
  dataMeasureName.Add(Format('%s=%s', ['Концентрация оксида углерода CO', 'концентрации оксида углерода CO']));
  dataMeasureName.Add(Format('%s=%s', ['Концентрация углекислого газа CO2', 'концентрации углекислого газа CO2']));
  dataMeasureName.Add(Format('%s=%s', ['Концентрация этилена С2H4', 'концентрации этилена С2H4']));
  dataMeasureName.Add(Format('%s=%s', ['Концентрация этана С2Н6', 'концентрации этана С2Н6']));
  dataMeasureName.Add(Format('%s=%s', ['Концентрация ацетилена С2H2', 'концентрации ацетилена С2H2']));
  dataMeasureName.Add(Format('%s=%s', ['Концентрация азота N2', 'концентрации азота N2']));
  dataMeasureName.Add(Format('%s=%s', ['Кислотное число', 'кислотного числа']));
  dataMeasureName.Add(Format('%s=%s', ['Температура вспышки', 'температуры вспышки']));
  dataMeasureName.Add(Format('%s=%s', ['Пробивное напряжение', 'пробивного напряжения']));
  dataMeasureName.Add(Format('%s=%s', ['Тангенс угла диэлектрических потерь', 'тангенса угла диэлектрических потерь']));
  dataMeasureName.Add(Format('%s=%s', ['Содержание присадок', 'содержания присадок']));
  dataMeasureName.Add(Format('%s=%s', ['Содержание растворенного шлама', 'содержания растворенного шлама']));
  dataMeasureName.Add(Format('%s=%s', ['Содержание фурановых производных (FAL)', 'содержания фурановых производных (FAL)']));
  dataMeasureName.Add(Format('%s=%s', ['Содержание фурановых производных (ACF)', 'содержания фурановых производных (ACF)']));
  dataMeasureName.Add(Format('%s=%s', ['Содержание фурановых производных (MEF)', 'содержания фурановых производных (MEF)']));
  dataMeasureName.Add(Format('%s=%s', ['Содержание фурановых производных (FOL)', 'содержания фурановых производных (FOL)']));
  dataMeasureName.Add(Format('%s=%s', ['Сопротивления изоляции', 'сопротивлений изоляции']));
  dataMeasureName.Add(Format('%s=%s', ['Частичные разряды', 'частичных разрядов']));
  dataMeasureName.Add(Format('%s=%s', ['Тангенс угла диэлектрических потерь', 'тангенса угла диэлектрических потерь']));
  dataMeasureName.Add(Format('%s=%s', ['Коэффициент трансформации', 'коэффициента трансформации']));
  dataMeasureName.Add(Format('%s=%s', ['Сопротивление КЗ', 'сопротивления КЗ']));
  dataMeasureName.Add(Format('%s=%s', ['Размер образца', 'размера образца']));
  dataMeasureName.Add(Format('%s=%s', ['Размер дефекта', 'размера дефекта']));
  dataMeasureName.Add(Format('%s=%s', ['Толщина металла', 'толщины металла']));
  dataMeasureName.Add(Format('%s=%s', ['Твердость', 'твердости']));

end;

procedure TShuhartForm.FormDestroy(Sender: TObject);
begin
    dataMeasureName.Free;
end;

procedure TShuhartForm.FormShow(Sender: TObject);
var i: integer;
  res: TModalResult;
begin
  {fLogin := TfLogin.Create(Self);
  res := fLogin.ShowModal;
  fLogin.Free;

  if (res <> mrOk) then begin
    halt;
    exit;
  end; }


  cxStyle1.Font.Height := ShuhartForm.Font.Height;
  cxStyle2.Font.Height := ShuhartForm.Font.Height;
  cxStyle3.Font.Height := ShuhartForm.Font.Height;
  cxStyle4.Font.Height := ShuhartForm.Font.Height;
  cxStyle5.Font.Height := ShuhartForm.Font.Height;

  X_Chart.Title.Font.Name := ShuhartForm.Font.Name;
  X_Chart.Title.Font.Height := ShuhartForm.Font.Height - 2;

  R_Chart.Title.Font.Name := ShuhartForm.Font.Name;
  R_Chart.Title.Font.Height := ShuhartForm.Font.Height - 2;

  S_Chart.Title.Font.Name := ShuhartForm.Font.Name;
  S_Chart.Title.Font.Height := ShuhartForm.Font.Height - 2;

  Me_Chart.Title.Font.Name := ShuhartForm.Font.Name;
  Me_Chart.Title.Font.Height := ShuhartForm.Font.Height - 2;

  SKO_Chart.Title.Font.Name := ShuhartForm.Font.Name;
  SKO_Chart.Title.Font.Height := ShuhartForm.Font.Height - 2;
  SKO_Chart.Legend.Font.Name := ShuhartForm.Font.Name;
  SKO_Chart.Legend.Font.Height := ShuhartForm.Font.Height - 2;

  cxGridView.OptionsView.BandHeaderHeight := Round(cxGridView.OptionsView.BandHeaderHeight * ShuhartForm.PixelsPerInch / 96);
  for i := 0 to cxGridView.Bands.Count - 1 do begin
    cxGridView.Bands.Items[i].Width := Round(cxGridView.Bands.Items[i].Width * ShuhartForm.PixelsPerInch / 96);
  end;


  for i := 0 to cxGridView.ColumnCount - 1 do begin
    cxGridView.Columns[i].Width := Round(cxGridView.Columns[i].Width * ShuhartForm.PixelsPerInch / 96);
  end;

  cxGridCriteriaView.OptionsView.DataRowHeight := Round(cxGridCriteriaView.OptionsView.DataRowHeight * ShuhartForm.PixelsPerInch / 96);
  for i := 0 to cxGridCriteriaView.ColumnCount - 1 do begin
    cxGridCriteriaView.Columns[i].Width := Round(cxGridCriteriaView.Columns[i].Width * ShuhartForm.PixelsPerInch / 96);
  end;

  for i := 0 to cxGridLaboratoryView.ColumnCount - 1 do begin
    cxGridLaboratoryView.Columns[i].Width := Round(cxGridLaboratoryView.Columns[i].Width * ShuhartForm.PixelsPerInch / 96);
  end;

  for i := 0 to cxGridResultView.ColumnCount - 1 do begin
    cxGridResultView.Columns[i].Width := Round(cxGridResultView.Columns[i].Width * ShuhartForm.PixelsPerInch / 96);
  end;

  InitData(true);
  InitDataMCI(true);

end;

end.

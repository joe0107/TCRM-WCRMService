unit PrtServiceRpt;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms, Dialogs, cDLL_Base120,
  ADODB, BetterADODataSet, cxPropertiesStore, JcCxGridResStr, DB, RzTabs, RzBmpBtn, ExtCtrls, Mask,
  Buttons, RzButton, RzPanel, DBClient, StdCtrls, RzEdit, DBCtrls, RzDBEdit, kbmMemTable, RzDBNav,
  DateUtils, cxClasses;

type
  TfmcWCPH1030 = class(TfmcDLL_Base120)
    btnOk: TRzBitBtn;
    btnAbort: TRzBitBtn;
    cdsReport: TClientDataSet;
    cdsReportREPORT_NO: TIntegerField;
    cdsReportrb_YEAR: TStringField;
    cdsReportrb_MONTH: TStringField;
    cdsReportrb_DAY: TStringField;
    cdsReportrb_IPCE002: TDateTimeField;
    cdsReportrb_CUT1002: TStringField;
    cdsReportrb_CUT1001: TStringField;
    cdsReportrb_SYSTEM: TStringField;
    cdsReportrb_ADDRESS: TStringField;
    cdsReportrb_SER_CE: TStringField;
    cdsReportrb_SER_DEP: TStringField;
    cdsReportrb_SER_LINE1: TStringField;
    cdsReportrb_SER_LINE2: TStringField;
    cdsReportrb_SER_LINE3: TStringField;
    cdsReportrb_SER_LINE4: TStringField;
    cdsReportrb_SER_LINE5: TStringField;
    cdsReportrb_SER_LINE6: TStringField;
    cdsReportrb_KIND1: TBooleanField;
    cdsReportrb_KIND2: TBooleanField;
    cdsReportrb_KIND3: TBooleanField;
    cdsReportrb_MEMO1: TStringField;
    cdsReportrb_MEMO2: TStringField;
    cdsReportrb_MEMO3: TStringField;
    cdsReportrb_CUT1005: TStringField;
    cdsReportrb_IPCE004: TStringField;
    cdsReportrb_CUT1035: TStringField;
    Label2: TLabel;
    edtCustId: TRzDBEdit;
    Label1: TLabel;
    edtCustName: TRzDBEdit;
    Label3: TLabel;
    edtCallFixDate: TRzDBDateTimeEdit;
    Label4: TLabel;
    edtContact: TRzDBEdit;
    Label5: TLabel;
    edtPhone: TRzDBEdit;
    Label6: TLabel;
    edtAgreeOver: TRzDBDateTimeEdit;
    Label7: TLabel;
    edtAddress: TRzDBEdit;
    Label8: TLabel;
    edtEmpName: TRzDBEdit;
    Label9: TLabel;
    edtDeptName: TRzDBEdit;
    Label10: TLabel;
    edtSysName: TRzDBEdit;
    Label11: TLabel;
    edtARAmount: TRzDBNumericEdit;
    Label12: TLabel;
    edtServiceItems: TRzDBMemo;
    dseSource: TDataSource;
    navSource: TRzDBNavigator;
    tblSource: TkbmMemTable;
    tblSourceCallFixDate: TDateTimeField;
    tblSourceAgreeOver: TDateTimeField;
    tblSourceCustId: TStringField;
    tblSourceCustName: TStringField;
    tblSourcePhone: TStringField;
    tblSourceContact: TStringField;
    tblSourceSysName: TStringField;
    tblSourceAddress: TStringField;
    tblSourceEmpName: TStringField;
    tblSourceDeptName: TStringField;
    tblSourceARAmount: TFloatField;
    tblSourceServiceItems: TMemoField;
    btnClear: TSpeedButton;
    labDataCount: TLabel;
    Label15: TLabel;
    labNowRec: TLabel;
    tblSourcePrintDate: TDateTimeField;
    Label13: TLabel;
    RzDBDateTimeEdit1: TRzDBDateTimeEdit;
    tblSourceCallFixCount: TStringField;
    cdsReportrb_CALLFIX_COUNT: TStringField;
    tblSourceSafeFixItems: TStringField;
    cdsReportrb_SAFEITEM_1: TStringField;
    cdsReportrb_SAFEITEM_2: TStringField;
    cdsReportrb_SAFEITEM_3: TStringField;
    cdsReportrb_SAFEITEM_4: TStringField;
    cdsReportrb_SAFEITEM_5: TStringField;
    cdsReportrb_SAFEITEM_6: TStringField;
    cdsReportrb_SAFEITEM_7: TStringField;
    cdsReportrb_SAFEITEM_8: TStringField;
    procedure btnOkClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure btnCancelClick(Sender: TObject);
    procedure FormKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure FormActivate(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure btnClearClick(Sender: TObject);
    procedure tblSourceAfterPost(DataSet: TDataSet);
    procedure tblSourceAfterDelete(DataSet: TDataSet);
    procedure dseSourceDataChange(Sender: TObject; Field: TField);
    procedure tblSourceAfterOpen(DataSet: TDataSet);
  private
    { Private declarations }
  public
    { Public declarations }

    class Procedure StartService; override;
    class procedure PrintServiceReport(fCallFixDate, fAgreeOver, APrintDate: TDateTime;
                                      fCustId, fCustName, fPhone, fContact,
                                      fSysName, fAddress, fEmpName, fDeptName: String;
                                      fARAmount: Double; fServiceItems, CallFixCount,
                                      SafeFixItems: String);
    procedure PrepareReportData(var Accept: Boolean); override;
  end;

var
  fmcWCPH1030: TfmcWCPH1030;

implementation

uses
	cController, cDLLDm, cUtility, DLL_COMMON, DLL_PUBLIC, cDB_Manager, vPubReport, CallLibrary, cBase;

{$R *.dfm}

{ TfmcWCPH1030 }

procedure TfmcWCPH1030.PrepareReportData(var Accept: Boolean);
const
  cMemo3Fmt = '至%s為止,尚有未收款:%s 元';
var
  //xCUT1035: Variant;
  xReportNo, K, M, i: Integer;
  xYear, xMonth, xDay: String;
  xFieldName, S: string;
  xList, aList: TStringList;

  procedure GetNowDate(var aYear, aMonth, aDay: String);
//  var
//     X1, X2, X3: Word;
  begin
    // Modified by C45 2013/05/10 上午 11:02:21
    (*
    aYear := ''; aMonth := ''; aDay := '';
    DecodeDate(Date, X1, X2, X3);
    X1 := X1 - 1911;
    aYear := IntToStr(X1);
    aMonth := IntToStr(X2);
    aDay := IntToStr(X3);
    *)
    aYear  := IntToStr(YearOf(tblSourcePrintDate.AsDateTime)-1911);
    aMonth := IntToStr(MonthOf(tblSourcePrintDate.AsDateTime));
    aDay   := IntToStr(DayOf(tblSourcePrintDate.AsDateTime));
  end;
begin
  xReportNo := 0;
  xList := TStringList.Create;
  aList := TStringList.Create;
  
  try
    GetNowDate(xYear, xMonth, xDay);

    cdsReport.EmptyDataSet;
    Inc(xReportNo);

    tblSource.First;
    while not tblSource.Eof do
    begin

      cdsReport.Append;
      cdsReport['REPORT_NO'] := xReportNo;
      cdsReport['rb_YEAR'] := xYear;
      cdsReport['rb_MONTH'] := xMonth;
      cdsReport['rb_DAY'] := xDay;

      cdsReport['rb_IPCE002'] := edtCallFixDate.Date;
      cdsReport['rb_CUT1001'] := edtCustId.Text;
      cdsReport['rb_IPCE004'] := edtContact.Text;
      cdsReport['rb_CUT1002'] := edtCustName.Text;
      if edtPhone.Text <> '' then
      cdsReport['rb_CUT1005'] := '電話:' + edtPhone.Text;

      if Trim(edtAgreeOver.Text) <> '' then
      cdsReport['rb_CUT1035'] := Format('硬體合約到期日: %s', [DatetimeToStr(edtAgreeOver.Date)]);

      cdsReport['rb_SYSTEM'] := edtSysName.Text;
      cdsReport['rb_ADDRESS'] := edtAddress.Text;
      cdsReport['rb_SER_CE'] := edtEmpName.Text;
      cdsReport['rb_SER_DEP'] := edtDeptName.Text;

      M := 0;
      xList.Assign(edtServiceItems.Lines);

      for K := 0 to xList.Count - 1 do
      begin
        S := Trim(xList.Strings[K]);
        While Length(S) >= 70 do
        begin
          INC(M);
          if M > 6 then Break;
          xFieldName := 'rb_SER_LINE' + IntToStr(M);
          cdsReport[xFieldName] := Trim(Copy(S, 1, 70));
          System.Delete(S, 1, 70);
          S := Trim(S);
        end;

        if Length(S) > 0 then
        begin
          INC(M);
          if M > 6 then Break;
          xFieldName := 'rb_SER_LINE' + IntToStr(M);
          cdsReport[xFieldName] := S;
          S := '';
        end;
      end;

      cdsReport['rb_KIND1'] := False;
      cdsReport['rb_KIND2'] := False;
      cdsReport['rb_KIND3'] := False;
      cdsReport['rb_MEMO1'] := '';
      cdsReport['rb_MEMO2'] := '';

      if not ((edtARAmount.Text = '') or (edtARAmount.Text = '0')) then
      	cdsReport['rb_MEMO3'] := Format(cMemo3Fmt, [DateTimeToStr(Date-1), edtARAmount.Text]);
      //2014.02.27
      cdsReport['rb_CALLFIX_COUNT'] := tblSourceCallFixCount.AsString;

      aList.Delimiter := ';';
      aList.DelimitedText := tblSourceSafeFixItems.AsString;

			for i := 0 to aList.Count-1 do
      begin
        if aList[i] = '1' then
      		cdsReportrb_SAFEITEM_1.AsString := 'V'
        else if aList[i] = '2' then
      		cdsReportrb_SAFEITEM_2.AsString := 'V'
        else if aList[i] = '3' then
      		cdsReportrb_SAFEITEM_3.AsString := 'V'
        else if aList[i] = '4' then
      		cdsReportrb_SAFEITEM_4.AsString := 'V'
        else if aList[i] = '5' then
      		cdsReportrb_SAFEITEM_5.AsString := 'V'
        else if aList[i] = '6' then
      		cdsReportrb_SAFEITEM_6.AsString := 'V'
        else if aList[i] = '7' then
      		cdsReportrb_SAFEITEM_7.AsString := 'V'
        else if aList[i] = '8' then
      		cdsReportrb_SAFEITEM_8.AsString := 'V';
      end;
      //-----------------------------------------------------------------------

      cdsReport.Post;
        tblSource.Next;
    end;
    Finally
    xList.Free;
    aList.Free;
  end;

  inherited;
  FPubReport.dsMaster.DataSet := cdsReport;
end;

class Procedure TfmcWCPH1030.StartService;
begin
  inherited;
  if not Assigned(fmcWCPH1030) then
  fmcWCPH1030 := TfmcWCPH1030.Create(Application);

  fmcWCPH1030.tblSource.Close;
  fmcWCPH1030.tblSource.Open;
  fmcWCPH1030.ShowModal;

  fmcWCPH1030.Free;
  fmcWCPH1030 := nil;
end;

procedure TfmcWCPH1030.btnOkClick(Sender: TObject);
begin
  if tblSource.RecordCount = 0 then
  begin
    MessageDlg('請至少輸入一筆要列印的資料.', mtError, [mbOk], 0);
    Abort;
  end;

  inherited;

  try
    tblSource.CheckBrowseMode;
    TfmvPubReport.CreateService(Self);
  except
  end;
end;

procedure TfmcWCPH1030.FormCreate(Sender: TObject);
begin
  inherited;
  cdsReport.CreateDataSet;
end;

procedure TfmcWCPH1030.btnCancelClick(Sender: TObject);
begin
  inherited;
  Close;
end;

procedure TfmcWCPH1030.FormKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
//  inherited;
  if not edtServiceItems.Focused then
  if (Key = VK_RETURN) then
  PostMessage(Handle, WM_KEYDOWN, VK_TAB, 0);
end;

class procedure TfmcWCPH1030.PrintServiceReport(fCallFixDate, fAgreeOver, APrintDate: TDateTime;
                                  fCustId, fCustName, fPhone, fContact,
                                  fSysName, fAddress, fEmpName, fDeptName: string;
                                  fARAmount: Double; fServiceItems, CallFixCount,
                                  SafeFixItems: string);
begin
  if not Assigned(fmcWCPH1030) then
  begin
    fmcWCPH1030 := TfmcWCPH1030.Create(Application);

    if Assigned(ADD_FUNCTION) then
    	ADD_FUNCTION(fmcWCPH1030);
  end;

  with fmcWCPH1030 do
  begin
    with tblSource do
    begin
      if not Active then
        Open;

      Append;
      FieldByName('CallFixDate').AsDateTime := fCallFixDate;
      FieldByName('PrintDate').AsDateTime := APrintDate;

      if fAgreeOver <> 0 then
      	FieldByName('AgreeOver').AsDatetime := fAgreeOver;

      FieldByName('CustId').AsString := fCustId;
      FieldByName('CustName').AsString := fCustName;
      FieldByName('Phone').AsString := fPhone;
      FieldByName('Contact').AsString := fContact;
      FieldByName('SysName').AsString := fSysName;
      FieldByName('Address').AsString := fAddress;
      FieldByName('EmpName').AsString := fEmpName;
      FieldByName('DeptName').AsString := fDeptName;
      FieldByName('ARAmount').AsFloat := fARAmount;
      FieldByName('ServiceItems').AsString := fServiceItems;
      FieldByName('CallFixCount').AsString := CallFixCount;
      FieldByName('SafeFixItems').AsString := SafeFixItems;
      Post;
    end;
    Show;

    if Assigned(SET_FUNCTION) then
    	SET_FUNCTION(fmcWCPH1030);
  end;
end;

procedure TfmcWCPH1030.FormActivate(Sender: TObject);
begin
  if Assigned(SET_FUNCTION) then
  SET_FUNCTION(fmcWCPH1030);
  inherited;
end;

procedure TfmcWCPH1030.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
  if Assigned(DEL_FUNCTION) then
  DEL_FUNCTION(fmcWCPH1030);
  inherited;

  Action := caFree;
  fmcWCPH1030 := nil;
end;

procedure TfmcWCPH1030.btnClearClick(Sender: TObject);
begin
  inherited;
  tblSource.Close;
  tblSource.Open;
end;

procedure TfmcWCPH1030.tblSourceAfterPost(DataSet: TDataSet);
begin
  inherited;
  labNowRec.Caption := IntToStr(tblSource.RecNo);
  labDataCount.Caption := IntToStr(tblSource.RecordCount);
end;

procedure TfmcWCPH1030.tblSourceAfterDelete(DataSet: TDataSet);
begin
  inherited;
  labNowRec.Caption := IntToStr(tblSource.RecNo);  
  labDataCount.Caption := IntToStr(tblSource.RecordCount);
end;

procedure TfmcWCPH1030.dseSourceDataChange(Sender: TObject; Field: TField);
begin
  inherited;
  labNowRec.Caption := IntToStr(tblSource.RecNo);
end;

procedure TfmcWCPH1030.tblSourceAfterOpen(DataSet: TDataSet);
begin
  inherited;
  labNowRec.Caption := '0';
  labDataCount.Caption := '0';
end;

end.

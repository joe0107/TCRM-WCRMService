library WCRMService;

{ Important note about DLL memory management: ShareMem must be the
  first unit in your library's USES clause AND your project's (select
  Project-View Source) USES clause if your DLL exports any procedures or
  functions that pass strings as parameters or function results. This
  applies to all strings passed to and from your DLL--even those that
  are nested in records and classes. ShareMem is the interface unit to
  the BORLNDMM.DLL shared memory manager, which must be deployed along
  with your DLL. To avoid using BORLNDMM.DLL, pass string information
  using PChar or ShortString parameters. }

uses
  ShareMem,
  SysUtils,
  Classes,
  Forms,
  Variants,
  Dialogs,
  cDLLDm in '..\Public_Template\cDLLDm.pas',
  cUtility in '..\Public_Template\cUtility.pas',
  cDB_Manager in '..\Public_Template\cDB_Manager.pas',
  CallLibrary in '..\Public_Template\CallLibrary.pas',
  DLL_COMMON in '..\Public_Template\DLL_COMMON.pas',
  DLL_PUBLIC in '..\Public_Template\DLL_PUBLIC.pas',
  cBase in '..\Public_Template\cBase.pas' {fmcBase},
  cDLL_Base in '..\Public_Template\cDLL_Base.pas' {fmcDLL_Base},
  cDLL_Base10 in '..\Public_Template\cDLL_Base10.pas' {fmcDLL_Base10},
  cDLL_Base100 in '..\Public_Template\cDLL_Base100.pas' {fmcDLL_Base100},
  cDLL_Base120 in '..\Public_Template\cDLL_Base120.pas' {fmcDLL_Base120},
  cController in '..\Public_Template\cController.pas' {dmPubController: TDataModule},
  vRB_BASE in '..\Public_Template\vRB_BASE.pas' {fmvRBBase},
  vRB_REPORT in '..\Public_Template\vRB_REPORT.pas' {fmvRBReport},
  vPubReport in '..\Public_Template\vPubReport.pas' {fmvPubReport},
  vTemplate in '..\Public_Template\vTemplate.pas' {fmTemplate},
  vGetTemplateFile in '..\Public_Template\vGetTemplateFile.pas' {fmGetTemplateFile},
  vMakeUpPrintPage in '..\Public_Template\vMakeUpPrintPage.pas' {fmMakeUpPrintSetup},
  vPrintToFileDlg in '..\Public_Template\vPrintToFileDlg.pas' {frmPrintToFileDlg},
  PrtServiceRpt in 'PrtServiceRpt.pas' {fmcWCPH1030},
  cDLL_Base210 in '..\Public_Template\cDLL_Base210.pas' {fmcDLL_Base210};

var
  SaveExit: Pointer;
  OldApHandle: THandle;

{$R *.RES}

procedure InitializeLib(AP: TApplication); stdcall;
begin
  Application.Handle := AP.Handle;
  Application.OnException := AP.OnException;
  CallerApp := AP;  // Added by Joe 2018/07/17 10:40:39
end;

procedure FinalizeLib; stdcall;
begin

end;

function CreateService: Integer; stdcall;
begin
  Result := 1;
end;

function StartService(ParamValues: OleVariant): Integer; stdcall;
begin
  {取得固定傳進來的兩個參數}
  FShow_Modal := StrToIntDef(DllParams.GetParamAsString('SHOW_MODAL'), 0);
  FService_NO := StrToIntDef(DllParams.GetParamAsString('SERVICE_NO'), 0);
  {呼叫要發動的From}
//  if varArrayHighBound(ParamValues, 1) = -1 then
    TfmcWCPH1030.StartService;
//  else
//    TfmcWCPH1030.FoundDataService(ParamValues);
	Result := 1;
end;

function StopService: Integer; stdcall;
begin
  {結束呼叫的From}
  if Assigned(fmcWCPH1030) then
     fmcWCPH1030.Free;
  Result := 1;
end;

procedure PrintServiceReport(fCallFixDate, fAgreeOver, APrintDate: TDateTime;
                             fCustId, fCustName, fPhone, fContact,
                             fSysName, fAddress, fEmpName, fDeptName: String;
                             fARAmount: Double; fServiceItems, CallFixCount, SafeFixItems: String); stdcall;
begin
  TfmcWCPH1030.PrintServiceReport(fCallFixDate, fAgreeOver, APrintDate,
                                  fCustId, fCustName, fPhone, fContact,
                                  fSysName, fAddress, fEmpName, fDeptName,
                                  fARAmount, fServiceItems, CallFixCount, SafeFixItems);
end;

procedure RegisterFormMethod(fAddForm, fDelForm, fSetForm: TRegisterForm); stdcall;
begin
  ADD_FUNCTION := fAddForm;
  DEL_FUNCTION := fDelForm;
  SET_FUNCTION := fSetForm;
end;

exports
  InitializeLib,
  FinalizeLib,
  SetADOConnection,
  SetDataSet,
  SetCallBackProc,
  SetParam,
  CreateService,
  StartService,
  StopService,
  PrintServiceReport,
  RegisterFormMethod;

procedure LibExit;
begin
  Application.Handle := OldApHandle;
  ExitProc := SaveExit;  // restore exit procedure chain
end;

begin
  OldApHandle := Application.Handle;
  SaveExit := ExitProc;  // save exit procedure chain
  ExitProc := @LibExit;  // install LibExit exit procedure
end.


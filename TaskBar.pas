﻿unit TaskBar;

//Установка значения
//TaskBar.TaskBarProgress :=50;

//TaskBarState:=2;

//NoProgress    = 0
//Indeterminate = 1 переливающаяся полоса
//Normal        = 2 Зеленый
//Error         = 3 Красный
//Paused        = 4 Желтый

//TaskBar.AlertState:=True //Красный полностью


interface
uses
  {$IFDEF MSWINDOWS}
  System.Win.ComObj,
  Winapi.ShlObj, System.Classes, Winapi.Windows, FMX.Platform.Win,
  System.SysUtils;
  {$ENDIF}
//--------﻿-----------------------------------------
type
  TTaskBarState = record
    name : string;
    value : byte;
  end;

  TFMXTaskBar = class(TPersistent)
  private
    fTaskBarState : integer;
    fTaskBarList : ITaskBarList3;
    fTaskBarProgress : integer;
    fAlertState : Boolean;
    procedure setTaskBarState(newState:integer);
    procedure setTaskBarProgress(newValue : integer);
    procedure setAlertState(const Value: Boolean);
  protected
  public
    property AlertState: Boolean read fAlertState write setAlertState;
    property TaskBarState : integer read fTaskBarState write setTaskBarState;
    property TaskBarProgress : integer read fTaskBarProgress write setTaskBarProgress;
    constructor Create;
    destructor Destroy; override;
  end;
//-------------------------------------------------

const
  taskStates : array [0..4] of TTaskBarState = (
            (name: 'TBPF_NOPROGRESS';    value: TBPF_NOPROGRESS),
            (name: 'TBPF_INDETERMINATE'; value: TBPF_INDETERMINATE),
            (name: 'TBPF_NORMAL';        value: TBPF_NORMAL),
            (name: 'TBPF_ERROR';         value: TBPF_ERROR),
            (name: 'TBPF_PAUSED';        value: TBPF_PAUSED));
//-------------------------------------------------
var
  MainTaskBar: TFMXTaskBar;
//-------------------------------------------------
implementation
{ TFMXTaskBar }

constructor TFMXTaskBar.Create;
var
  tbList : ITaskBarList;
  hr : HRESULT;
begin
  tbList := CreateComObject(CLSID_TaskBarList) as ITaskBarList;
  hr := tbList.QueryInterface(IID_ITaskBarList3, fTaskBarList);
  if hr <> S_OK then
  begin
    fTaskBarList := nil;
    tbList._Release;
  end;
end;

destructor TFMXTaskBar.Destroy;
begin

  inherited;
end;

procedure TFMXTaskBar.setAlertState(const Value: Boolean);
begin
  if assigned(fTaskBarList) then
    if fAlertState <> Value then
    begin
      fAlertState := Value;

      if Value then
      begin
        setTaskBarState(3);       // TBPF_ERROR
        setTaskBarProgress(100);
      end else
      begin
        setTaskBarState(0);
        setTaskBarProgress(0);
      end;
    end;
end;

procedure TFMXTaskBar.setTaskBarProgress(newValue: integer);
var AppHandle:HWnd;
begin
  AppHandle := ApplicationHWND;
  if assigned(fTaskBarList) then
    fTaskBarList.SetProgressValue(ApplicationHWND, newValue, 100);
end;

procedure TFMXTaskBar.setTaskBarState(newState: integer);
begin
  if assigned(fTaskBarList) then
    fTaskBarList.SetProgressState(ApplicationHWND, taskStates[newState].value);
end;
//-------------------------------------------------

initialization
  MainTaskBar := TFMXTaskBar.Create;

finalization
  if Assigned(MainTaskBar) then
    FreeAndNil(MainTaskBar);

end.

unit Main;

interface

uses
  System.SysUtils, System.Types, System.UITypes, System.Classes, System.Variants,
  FMX.Types, FMX.Controls, FMX.Forms, FMX.Graphics, FMX.Dialogs, FMX.Edit,
  FMX.Controls.Presentation, System.Rtti, FMX.Grid.Style, FMX.Grid,
  FMX.ScrollBox,
  System.Win.ComObj, FMX.Platform,
  StrUtils, FMX.StdCtrls;

type
  TForm2 = class(TForm)
    Edit1: TEdit;
    SpinEditButton1: TSpinEditButton;
    StringGrid1: TStringGrid;
    StringColumn1: TStringColumn;
    StringColumn2: TStringColumn;
    StringColumn3: TStringColumn;
    StringColumn4: TStringColumn;
    OpenDialog1: TOpenDialog;
    Timer1: TTimer;
    CheckBox1: TCheckBox;
    CheckBox2: TCheckBox;
    Label1: TLabel;

    procedure Edit1Typing(Sender: TObject);
    procedure Edit1Click(Sender: TObject);

    procedure SearchInGrid(SearchText, param: string);
    procedure GetSearchResultQty(SearchText: string);
    function GetCurrentSearchPos(direction: string): string;

    procedure LoadXLS(XLSFile:string; Grid:TStringGrid);

    procedure FormActivate(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure Timer1Timer(Sender: TObject);
    procedure SpinEditButton1DownClick(Sender: TObject);
    procedure SpinEditButton1UpClick(Sender: TObject);
    procedure Edit1MouseWheel(Sender: TObject; Shift: TShiftState;
      WheelDelta: Integer; var Handled: Boolean);
    procedure FormDeactivate(Sender: TObject);
    procedure FormKeyDown(Sender: TObject; var Key: Word; var KeyChar: Char;
      Shift: TShiftState);

  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
    Form2: TForm2;
    SearchRow, SearchCol: integer;
    CurrentSearchResult, SearchResultQty: integer;
    SelectText: boolean;
implementation

{$R *.fmx}

procedure TForm2.FormCreate(Sender: TObject);
begin
    Timer1.Enabled:=True;
    SelectText:=True;
end;

procedure TForm2.FormActivate(Sender: TObject);
var Svc: IFMXClipboardService;
    s: string;
begin
    try
        if TPlatformServices.Current.SupportsPlatformService(IFMXClipboardService, Svc)
        then s:=Svc.GetClipboard.AsString;
        if s<>''
        then begin
            Edit1.Text:=Trim(s);
            SearchInGrid(s, 'first');
        end;
    except

    end;
end;

procedure TForm2.FormDeactivate(Sender: TObject);
var Svc: IFMXClipboardService;
    s: string;
begin
    if CheckBox1.IsChecked
    then begin
        try
            if TPlatformServices.Current.SupportsPlatformService(IFMXClipboardService, Svc)
            then s:=Svc.GetClipboard.AsString;
            if s<>''
            then begin
                s:=ansiuppercase(s[1])+Copy(s,2,Length(s));
                Svc.SetClipboard(s);
            end;
        except

        end;
    end;
end;

procedure TForm2.Timer1Timer(Sender: TObject);
begin
    Timer1.Enabled:=False;
    if FileExists(ExtractFileDir(ParamStr(0))+PathDelim +'enru.xls')
    then begin
        Edit1.Enabled:=False;
        LoadXLS( ExtractFileDir(ParamStr(0))+PathDelim +'enru.xls', StringGrid1);
        Edit1.Enabled:=True;
    end
    else
        if OpenDialog1.Execute
        then begin
            Edit1.Enabled:=False;
            LoadXLS (OpenDialog1.FileName, StringGrid1);
            Edit1.Enabled:=True;
        end
        else Application.Terminate;
end;

{$REGION 'XLS'}
  procedure TForm2.LoadXLS(XLSFile:string; Grid:TStringGrid);
   const
    xlCellTypeLastCell = $0000000B;
  var
    ExlApp, Sheet: OLEVariant;
    i, j, r, c: integer;
    namestr: string;
    checkdata: string;
    tempstr: string;
  begin
      ExlApp := CreateOleObject('Excel.Application');
      ExlApp.Visible := false;
      ExlApp.Workbooks.Open(XLSFile);
      Sheet := ExlApp.Workbooks[ExtractFileName(XLSFile)].WorkSheets[1];
      Sheet.Cells.SpecialCells(xlCellTypeLastCell, EmptyParam).Activate;
      r := ExlApp.ActiveCell.Row;
      c := ExlApp.ActiveCell.Column;
      Grid.RowCount:=r;
      c :=Grid.ColumnCount;
      Form2.Caption:='Загрузка 0%';
      Application.ProcessMessages;
      for j:= 1 to r do
      begin
          try
              if (j mod 1000 = 0)
              then begin
                  Form2.Caption:='Загрузка '+FormatFloat('0.00%',(j/r*100));
                  Application.ProcessMessages;
              end;

              for i:= 1 to c do
              begin
                  Grid.Cells[i-1,j-1]:=sheet.cells[j,i];
              end;
          except on e:exception do
              //ShowMessage(inttostr(j)+' '+e.Message);
          end;
      end;
      Form2.Caption:='Загрузка 100%';
      ExlApp.Quit;
      ExlApp := Unassigned;
      Sheet := Unassigned;
  end;

{$REGION 'Search'}
  procedure TForm2.SearchInGrid(SearchText, param: string);
  var i, k: integer;
  StartCol: integer;
  TextFound: boolean;
  begin
      //SearchInGrid(Edit1.Text, 'first');
      if Param='first'
      then begin
          SearchRow:=-1;
          SearchCol:=-1;
          CurrentSearchResult:=0;

          for I := 0 to StringGrid1.RowCount-1 do
          begin
              for k := 0 to StringGrid1.ColumnCount-1 do
              begin
                  TextFound:=false;

                  if CheckBox2.IsChecked
                  then TextFound:=(SearchText=StringGrid1.Cells[k,i])
                  else TextFound:=( Pos( AnsiLowerCase(SearchText), AnsiLowerCase(StringGrid1.Cells[k,i]) )>0 );

                  if TextFound
                  then begin
                      StringGrid1.SelectCell(k,i);
                      StringGrid1.ScrollToSelectedCell;
                      SearchRow:=i;
                      SearchCol:=k;
                      CurrentSearchResult:=1;
                      GetSearchResultQty(SearchText);
                      Caption:='Найдено результатов: '+IntToStr(SearchResultQty);
                      exit;
                  end;
              end;
          end;
          Caption:='Не найдено результатов / '+SearchText;
      end;

      if Param='next'
      then begin
          if searchcol=StringGrid1.ColumnCount-1
          then begin
              SearchRow:=SearchRow+1;
              SearchCol:=0;
          end
          else begin
              SearchCol:=SearchCol+1;
          end;

          for I:= SearchRow to StringGrid1.RowCount-1 do
          begin
              if i=searchRow
              then StartCol:=SearchCol
              else StartCol:=0;

              for k:= StartCol to StringGrid1.ColumnCount-1 do
              begin
                  if Pos( AnsiLowerCase(SearchText), AnsiLowerCase(StringGrid1.Cells[k,i]) )>0
                  then begin
                      StringGrid1.SelectCell(k,i);
                      StringGrid1.ScrollToSelectedCell;
                      SearchRow:=i;
                      SearchCol:=k;
                      Caption:='Результат: '+GetCurrentSearchPos('next')+' из '+IntToStr(SearchResultQty);
                      exit;
                  end;
              end;
          end;
          Caption:='Достигнут конец результатов. Возврат поиска в начало.';
          CurrentSearchResult:=0;
          StringGrid1.SelectCell(0,0);
          StringGrid1.ScrollToSelectedCell;
          SearchRow:=-1;
          SearchCol:=-1;
      end;

      if Param='prev'
      then begin
          if searchcol=0
          then begin
              SearchRow:=SearchRow-1;
              SearchCol:=StringGrid1.ColumnCount-1;
          end
          else begin
              SearchCol:=SearchCol-1;
          end;

          for I := SearchRow downto 0 do
          begin
              if i=searchRow
              then StartCol:=SearchCol
              else StartCol:=StringGrid1.ColumnCount-1;

              for k := StartCol downto 0 do
              begin
                  if Pos( AnsiLowerCase(SearchText), AnsiLowerCase(StringGrid1.Cells[k,i]) )>0
                  then begin
                      StringGrid1.SelectCell(k,i);
                      StringGrid1.ScrollToSelectedCell;
                      SearchRow:=i;
                      SearchCol:=k;
                      Caption:='Результат: '+GetCurrentSearchPos('prev')+' из '+IntToStr(SearchResultQty);
                      exit;
                  end;
              end;
          end;
          Caption:= 'Достигнуто начало результатов. Возврат поиска в конец.';
          CurrentSearchResult:=SearchResultQty;
          StringGrid1.SelectCell(StringGrid1.ColumnCount-1, StringGrid1.RowCount-1);
          StringGrid1.ScrollToSelectedCell;
          SearchRow:=StringGrid1.RowCount-1;
          SearchCol:=StringGrid1.ColumnCount;
      end;
  end;


procedure TForm2.GetSearchResultQty(SearchText: string);
  var i, k: integer;
  begin
      SearchResultQty:=0;
      for I := 0 to StringGrid1.RowCount-1 do
      begin
          for k := 0 to StringGrid1.ColumnCount-1 do
          begin
              if Pos( AnsiLowerCase(SearchText), AnsiLowerCase(StringGrid1.Cells[k,i]) )>0
              then begin
                  SearchResultQty:=SearchResultQty+1;
              end;
          end;
      end;
  end;

  function TForm2.GetCurrentSearchPos(direction: string): string;
  begin
      if direction='next'
      then begin
          CurrentSearchResult:=CurrentSearchResult+1;
          if CurrentSearchResult>SearchResultQty
          then CurrentSearchResult:=1;
          Result:=inttostr(CurrentSearchResult);
      end
      else begin
          CurrentSearchResult:=CurrentSearchResult-1;
          if CurrentSearchResult<1
          then CurrentSearchResult:=SearchResultQty;
          Result:=inttostr(CurrentSearchResult);
      end;
  end;
{$ENDREGION}

procedure TForm2.Edit1Click(Sender: TObject);
begin
    if SelectText
    then begin
        Edit1.SelectAll;
        SelectText:=False;
    end;
end;

procedure TForm2.Edit1Typing(Sender: TObject);
begin
    //Button2.Enabled:=False;
    SearchInGrid(Edit1.Text, 'first');
end;

procedure TForm2.SpinEditButton1UpClick(Sender: TObject);
begin
    SearchInGrid(Edit1.Text, 'prev');
end;

  procedure TForm2.SpinEditButton1DownClick(Sender: TObject);
begin
    SearchInGrid(Edit1.Text, 'next');
end;

procedure TForm2.Edit1MouseWheel(Sender: TObject; Shift: TShiftState;
  WheelDelta: Integer; var Handled: Boolean);
begin
    if WheelDelta > 0
    then SearchInGrid(Trim(Edit1.Text), 'prev')
    else SearchInGrid(Trim(Edit1.Text), 'next');
end;

procedure TForm2.FormKeyDown(Sender: TObject; var Key: Word; var KeyChar: Char;
  Shift: TShiftState);
begin
    case key of
        113: SearchInGrid(Trim(Edit1.Text), 'prev');
        114: SearchInGrid(Trim(Edit1.Text), 'next');
    end;
end;

end.

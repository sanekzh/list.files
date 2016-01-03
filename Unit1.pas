unit Unit1;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, Grids, ComObj, StdCtrls, DB, ADODB, Menus, ComCtrls, ToolWin,
  ExtCtrls;

type
  TForm1 = class(TForm)
    StringGrid1: TStringGrid;
    Button1: TButton;
    Edit1: TEdit;
    OpenDialog1: TOpenDialog;
    ADODataSet1: TADODataSet;
    ADOConnection1: TADOConnection;
    ADOCommand1: TADOCommand;
    ADOQuery1: TADOQuery;
    Button2: TButton;
    MainMenu1: TMainMenu;
    File1: TMenuItem;
    OpenFile1: TMenuItem;
    Exit1: TMenuItem;
    Edit3: TMenuItem;
    Label1: TLabel;
    Memo1: TMemo;
    Label2: TLabel;
    Label3: TLabel;
    Label4: TLabel;
    Label5: TLabel;
    Edit4: TEdit;
    Edit5: TEdit;
    Edit6: TEdit;
    Edit2: TEdit;
    ProgressBar1: TProgressBar;
    StatusBar1: TStatusBar;
    N1: TMenuItem;
    N2: TMenuItem;
    Button3: TButton;
    procedure Button2Click(Sender: TObject);
    procedure Button1Click(Sender: TObject);
    procedure Exit1Click(Sender: TObject);
    procedure OpenFile1Click(Sender: TObject);
 
    procedure N2Click(Sender: TObject);
    procedure Button3Click(Sender: TObject);
  
  private
  function OpenFile():integer;
  function ClearTF():integer;
    { Private declarations }
  public
  s, sg, nomer, obname, info, grupp, username : string;
  btn : integer;
    { Public declarations }
  end;

var
  Form1: TForm1;
 // Form2: TForm1;

implementation

uses Unit2;

{$R *.dfm}
function Xls_To_StringGrid(AGrid: TStringGrid; AXLSFile: string): Boolean;
const
  xlCellTypeLastCell = $0000000B;
var
  XLApp, Sheet: OLEVariant;
  RangeMatrix: Variant;
  x, y, k, r: Integer;
begin
  Result := False;
  // Create Excel-OLE Object
  XLApp := CreateOleObject('Excel.Application');
  try
    // Hide Excel
    XLApp.Visible := False;
    try
    // Open the Workbook
    XLApp.Workbooks.Open(AXLSFile);
    except
      MessageDlg('Cannot open MS Excel document!'+AXLSFile+')',mtError,[mbOk],0); Exit;
    end;
    // Sheet := XLApp.Workbooks[1].WorkSheets[1];
    Sheet := XLApp.Workbooks[ExtractFileName(AXLSFile)].WorkSheets[1];

    // In order to know the dimension of the WorkSheet, i.e the number of rows
    // and the number of columns, we activate the last non-empty cell of it

    Sheet.Cells.SpecialCells(xlCellTypeLastCell, EmptyParam).Activate;
    // Get the value of the last row
    x := XLApp.ActiveCell.Row;
    // Get the value of the last column
    y := XLApp.ActiveCell.Column;

    // Set Stringgrid's row &col dimensions.

    AGrid.RowCount := x;
    AGrid.ColCount := y;

    // Assign the Variant associated with the WorkSheet to the Delphi Variant

    RangeMatrix := XLApp.Range['A1', XLApp.Cells.Item[X, Y]].Value;
    //  Define the loop for filling in the TStringGrid
    k := 1;
    repeat
      for r := 1 to y do
        AGrid.Cells[(r - 1), (k - 1)] := RangeMatrix[K, R];
      Inc(k, 1);
      AGrid.RowCount := k + 1;
    until k > x;
    // Unassign the Delphi Variant Matrix
    RangeMatrix := Unassigned;

  finally
    // Quit Excel
    if not VarIsEmpty(XLApp) then
    begin
      // XLApp.DisplayAlerts := False;
      XLApp.Quit;
      XLAPP := Unassigned;
      Sheet := Unassigned;
      Result := True;
    end;
  end;
end;

procedure TForm1.Button2Click(Sender: TObject);
var
      i, j, vnes, sovpad, vsego : integer;
begin
    //if Xls_To_StringGrid(StringGrid1, 'd:\test.xls') then
  //  if Xls_To_StringGrid(StringGrid1, s) then
//    ShowMessage('Table has been exported!');
    // Удалить пустые столбцы
    // for i:=StringGrid1.ColCount-1 downto 0 do
    // begin
    //    if StringGrid1.Cells[i,0]='' then
    //    begin
    //      StringGrid1.ColCount:=StringGrid1.ColCount-1;
   //     end;
   //  end;
     Edit2.Text:= IntToStr(StringGrid1.RowCount-1);
     vnes:=0;
     vsego:=0;
     sovpad:=0;
//  begin
  //        MessageDlg('Ошибка доступа к базе SQL', mtError, [mbOK], 0);
    //      break;
      //    end;
     //  ADODataSet1.Active:=False;
  //ADODataSet1.Active:=True;
 //  begin
  //  MessageDlg('Ошибка доступа к базе SQL!', mtError, [mbOK], 0);

  // end else
 //  begin
     for i:=1 to StringGrid1.RowCount-1 do
     begin
     ProgressBar1.Position:=Round(100*i/StringGrid1.RowCount);
          nomer:= StringGrid1.Cells[1,i];
          vsego:=vsego+1;
          for j:=Length(nomer) downto 1 do
          begin
             if not(nomer[j] in ['0'..'9']) then
             begin
                 Delete(nomer,j,1);
             end;
          end;
         // Edit2.Text:=nomer;
         // ShowMessage('Проверка номера '+nomer);
         if Length(nomer) < 10 then
         begin
        // ShowMessage('Неверный номер: '+nomer+'');
           vsego:=vsego-1;
         end;
         if Length(nomer) > 10 then
         begin
          ADODataSet1.Active:=False;
          ADODataset1.CommandText:='select * from `Arhiv`.`spisok` where nomer like ''%'+nomer+'''';
          ADODataSet1.Active:=True;
         // Edit2.Text:='123';
          if ADODataset1.IsEmpty then
           begin
              obname:= StringGrid1.Cells[2,i];
              for j:=0 to Length(obname) do
               begin
                 if obname[j]= #39 then
                   begin
                     obname[j]:='*';
                   end;
                 if obname[j]= #13 then
                   begin
                     obname[j]:=' ';
                   end;
               end;
               info:= StringGrid1.Cells[3,i];
               if Length(info)>0 then
               begin
                     for j:=0 to Length(info) do
                     begin
                         if info[j]= #39 then
                         begin
                            info[j]:='*';
                         end;
                     end;
               end else info:='  ';

               grupp:= 'root';
               username:= 'root';
                ADOQuery1.Close;
                ADOQuery1.SQL.Text:='INSERT INTO `Arhiv`.`spisok`(`nomer`,`kom`,`inf`,`grup`,`data`,`username`)'+ 'VALUES('''+nomer+''','''+obname+''','''+info+''','''+grupp+''','''+FormatDateTime('yyyy.mm.dd',Date)+''','''+username+''')';
                ADOQuery1.ExecSQL;
               vnes:=vnes+1;
            end else
              begin
                Memo1.Lines.Add(nomer);
                sovpad:=sovpad+1;
              end;
      //ShowMessage('Этот '+nomer+'номер уже внесен в базу.');
          end;
         Edit4.Text:=IntToStr(vsego);
         Edit5.Text:=IntToStr(vnes);
         Edit6.Text:=IntToStr(sovpad);
     end;
     ShowMessage('Table has been exported!');
    //end;
end;

function TForm1.ClearTF(): Integer;
var
    i : Integer;
begin
    Edit1.Text:='';
    Edit4.Text:='';
    Edit5.Text:='';
    Edit6.Text:='';
    btn:=0;
    ProgressBar1.Position:=0;
    Memo1.Text:='';
    Button2.Enabled:=false;
    with StringGrid1 do
     begin
       for i:=0 to ColCount-1 do
        Cols[i].Clear;
     end;
end;

Function TForm1.OpenFile(): Integer;
var
     i, exec: Integer;
     fn : String;
begin
    Edit1.Text:='';
    Edit4.Text:='';
    Edit5.Text:='';
    Edit6.Text:='';
    btn:=0;
    exec:=0;
    ProgressBar1.Position:=0;
    Memo1.Text:='';
     with StringGrid1 do
     begin
       for i:=0 to ColCount-1 do
        Cols[i].Clear;
     end;
     if OpenDialog1.Execute then
       begin
         s:=OpenDialog1.Files.Strings[0];
         fn:=OpenDialog1.Files.Strings[0];
         fn:=ExtractFileName(fn);
         Edit1.Text:=fn;
         exec:=1;
       end;
    if exec=1 then
    begin
     if Xls_To_StringGrid(StringGrid1, s) then
       // ShowMessage('Table has been exported!');
     for i:=StringGrid1.RowCount-1 downto 0 do
        begin
          if StringGrid1.Cells[0,i]='' then
            begin
               StringGrid1.RowCount:=StringGrid1.RowCount-1;
            end;
        end;
        btn:=1;
       if btn=1 then
        begin
          Button2.Enabled:=True;
        end;
    end;    
end;
procedure TForm1.Button1Click(Sender: TObject);
begin
     OpenFile();
end;
procedure TForm1.Exit1Click(Sender: TObject);
begin
      Close();
end;
procedure TForm1.OpenFile1Click(Sender: TObject);
begin
      OpenFile();
end;
procedure TForm1.N2Click(Sender: TObject);
begin
     ClearTF();
end;
procedure TForm1.Button3Click(Sender: TObject);
begin
  // if (not Assigned(Form2)) then
  // Form2:=TForm.Create(Self);
   Form2.Show;
end;

end.

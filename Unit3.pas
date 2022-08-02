unit Unit3;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs,   Vcl.ExtCtrls, Vcl.Grids,
  Vcl.StdCtrls;

type
  TForm3 = class(TForm)
    Panel1: TPanel;
    StringGrid1: TStringGrid;
    OpenDialog1: TOpenDialog;
    Button1: TButton;
    procedure RangeRead;
    //procedure Button2Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form3: TForm3;

implementation

{$R *.dfm}

 uses ComObj;

{procedure TForm3.Button2Click(Sender: TObject);
var
  exApp, exBook, exSh, exUsRng, exTblRng : Variant;
  i, j, Row1, Col1, Row2, Col2 : Integer;
  Sg : TStringGrid;
  Od : TOpenDialog;
begin
  Od := OpenDialog1;
  Sg := StringGrid1;
  if Od.InitialDir = '' then
    Od.InitialDir := ExtractFilePath(ParamStr(0))
  ;
  if not Od.Execute then Exit;
  if not FileExists(Od.FileName) then begin
    Exit;
    ShowMessage('Файл с заданным именем не найден. Действие отменено.');
  end;
  //Пытаемся запустить Excel и подключиться к нему.
  try
    exApp := CreateOleObject('Excel.Application');
  except
    ShowMessage('Не удалось подключиться к MS Excel. Действие отменено.');
    Exit;
  end;

  //Очистка StringGrid1. - Для устранения последствий бага, при котором
  //компонент типа TStringGrid может не удалять строки и столбцы, а скрывать их.
  for j := 0 to Sg.ColCount - 1 do begin
    Sg.Cols[j].Clear;
  end;

  //На время отладки делаем окно Excel видимым.
  exApp.Visible := True;
  //Открываем файл рабочей книги.
  exBook := exApp.Workbooks.Open(FileName:=Od.FileName);
  //Подключение к первому листу рабочей книги.
  exSh := exBook.Worksheets[1];
  //Определяем рабочий диапазон.
  exUsRng := exSh.UsedRange;
  //Предположим, мы знаем, что левый верхний край таблицы должен находиться
  //в координатах: Row = 3, Col = 2.
  Row1 := 1;
  Col1 :=  1;
  Row2 := exUsRng.Row + exUsRng.Rows.Count - 1;
  Col2 :=  exUsRng.Column + exUsRng.Columns.Count - 1;
  if (Row1 > Row2) or (Col1 > Col2) then begin
    ShowMessage('Таблица не обнаружена. Действие отменено.');
    Exit;
  end;

  //Диапазон таблицы.
  exTblRng := exSh.Range[exSh.Cells[Row1, Col1], exSh.Cells[Row2, Col2]];

  //Переносим размер таблицы.
  Sg.FixedRows := 0;
  Sg.FixedCols := 0;
  Sg.RowCount := exTblRng.Rows.Count;
  Sg.ColCount := exTblRng.Columns.Count;

  //Переносим шапку. Предположим, нам известно, что шапка занимает
  //первую строку в таблице.
  for j := 0 to Sg.ColCount - 1 do begin
    Sg.Cells[j, 0] := exTblRng.Cells[1, 1 + j].Text;
  end;
  //TStringGrid обязательно должна содержать хотябы одну нефиксированную строку.
  if Sg.RowCount = 1 then Sg.RowCount := Sg.RowCount + 1;
  Sg.FixedRows := 1;

  //Перенос данных.
  for i := Sg.FixedRows to Sg.RowCount - 1 do begin
    for j := 0 to Sg.ColCount - 1 do begin
      Sg.Cells[j, i] := exTblRng.Cells[i + 1, j + 1].Text;
    end;
  end;

  //Закрываем книгу и выходим из Excel.
  //На время отладки отключено.
  //exBook.Close;
  //exApp.Quit;

  end;}

  procedure TForm3.RangeRead;
var Rows, Cols, i,j: integer;
    WorkSheet: OLEVariant;
    FData: OLEVariant;
    d: TDateTime;
begin
  //открываем книгу
  ExcelApp.Workbooks.Open(edFile.Text);
  //получаем активный лист
  WorkSheet:=ExcelApp.ActiveWorkbook.ActiveSheet;
  //определяем количество строк и столбцов таблицы
  Rows:=WorkSheet.UsedRange.Rows.Count;
  Cols:=WorkSheet.UsedRange.Columns.Count;

  //считываем данные всего диапазона
  FData:=WorkSheet.UsedRange.Value;

  StringGrid1.RowCount:=Rows;
  StringGrid1.ColCount:=Cols;

//выводим данные в таблицу
  for I := 0 to Rows-1 do
    for j := 0 to Cols-1 do
        StringGrid1.Cells[J,I]:=FData[I+1,J+1];

end;
end.

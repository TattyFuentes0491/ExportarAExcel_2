unit Exportar;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls, Vcl.Grids, XLSSheetData5,
  XLSReadWriteII5, Xc12Utils5, Data.DB, Datasnap.DBClient, ShellApi,
  IBX.IBCustomDataSet, IBX.IBDatabase, Vcl.ExtCtrls, Vcl.DBCtrls, IBX.IBQuery,
  Vcl.XPMan, Xc12DataStyleSheet5, XLSCmdFormat5, XLSComment5, XLSDrawing5, 
  XLSCmdFormatValues5, XLSRelCells5, XLSCellAreas5, XLSTools5, XLSRow5, XLSColumn5, 
  XLSAutofilter5, XLSNames5, XLSValidate5;

type
  TForm1 = class(TForm)
    Button1: TButton;
    Button2: TButton;
    Button3: TButton;
    tabla: TStringGrid;
    XLSExcel: TXLSReadWriteII5;
    IBDatabase1: TIBDatabase;
    IBQuery1: TIBQuery;
    IBQuery1CANT_EXA: TIntegerField;
    IBQuery1PROM_SEG_TOMA: TIntegerField;
    IBQuery1PROM_SEG_REPORT: TIntegerField;
    IBQuery1HMS_TOMA: TIBStringField;
    IBQuery1HMS_REPORT: TIBStringField;
    IBTransaction1: TIBTransaction;
    XPManifest1: TXPManifest;
    qry: TIBQuery;
    IBTransaction2: TIBTransaction;
    IBQuery1GRUPO_1: TIBStringField;
    Button4: TButton;
    Button5: TButton;
    procedure Button5Click(Sender: TObject);

    const archivo= 'D:\Documents\Embarcadero\Studio\Projects\ExportarAExcel_2\tiempos.xlsx';

    procedure Button3Click(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure Button1Click(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure Button2Click(Sender: TObject);
    procedure Button4Click(Sender: TObject);
  private
    { Private declarations }
    procedure llenarTabla;
  public
    { Public declarations }
    procedure precargarHojaExcel;
    procedure cargarDatos;
    procedure cargarTitulos;
    procedure formatoTipoAcumulado;
    procedure escribirExcel;
    procedure formatoTitulos(col,row: Integer);
    procedure formatoSubtitulos(col, row: Integer);
    procedure formatoEncabezado;
    procedure limpiarArreglo(arrayStr: array of Variant);
  end;

var
  Form1: TForm1;

implementation

{$R *.dfm}

{ TForm1 }

procedure TForm1.Button1Click(Sender: TObject);
begin
  cargarDatos;
end;

procedure TForm1.Button2Click(Sender: TObject);
begin
  llenarTabla;
end;

procedure TForm1.Button3Click(Sender: TObject);
var
  i: Integer;
begin
  for i := 0 to tabla.RowCount-1 do begin
    tabla.Rows[i].Clear;
  end;
end;

procedure TForm1.cargarDatos;
var
  i,j: Integer;
begin
  try
    IBQuery1.Open;
    with IBQuery1, tabla do begin
      FetchAll;//cantidad de registros que trae la consulta
      RowCount:= RecordCount+1+1+1; //Filas de la Grilla = Cant de Registros + título + fila 
      Cols[RowCount - 1].Clear;
      ColCount:= FieldCount;
      First;//primer registro   
      for i:=2 to (RowCount-2) do begin //recorremos
        for j:=0 to (ColCount-1) do begin
          Cells[j,i]:= Fields[j].AsString;
        end;
         Next;
      end;
    end;
    cargarTitulos;
  except on E: Exception do
    MessageDlg('Error ',mtError,[mbCancel],0, mbOK);
  end;
end;

procedure TForm1.escribirExcel;
begin
  XLSExcel.Filename:= archivo;
  XLSExcel.Sheets[0].Name:= 'Tiempos';
  XLSExcel.Write;
  ShowMessage('Archivo creado');
end;

procedure TForm1.Button4Click(Sender: TObject);
begin
  ShellExecute(Handle,'open', 'excel.exe',PChar(archivo), nil, SW_SHOWNORMAL);
end;

procedure TForm1.Button5Click(Sender: TObject);
begin
  formatoTipoAcumulado;
end;

procedure TForm1.formatoTitulos(col,row: Integer);
var
  I: Integer;
begin
  //formato para titulos
  XLSExcel.CmdFormat.BeginEdit(XLSExcel[0]);
  XLSExcel.CmdFormat.Border.Color.RGB := $FF0000;
  XLSExcel.CmdFormat.Border.Style := cbsThick;
  XLSExcel.CmdFormat.Font.Style:= [xfsBold];
  XLSExcel.CmdFormat.Font.Name:= 'Arial';
  XLSExcel.CmdFormat.Font.Size:= 10;
  XLSExcel.CmdFormat.Alignment.Horizontal:= chaCenter;
  XLSExcel.CmdFormat.Fill.BackgroundColor.RGB:=  $C8EAD7;
  XLSExcel.CmdFormat.Apply(col,row);
end;

procedure TForm1.formatoEncabezado;
begin
  //XLSExcel[0].AsString[0,0]:= 'LABORATORIO XXXX';  //no toma cambios en la celda 0,0 por temas de copyRigth
  XLSExcel.CmdFormat.BeginEdit(XLSExcel[0]);
  XLSExcel.CmdFormat.Font.Style:= [xfsBold];
  XLSExcel.CmdFormat.Font.Name:= 'Arial';
  XLSExcel.CmdFormat.Font.Size:= 12;
  XLSExcel.CmdFormat.Fill.BackgroundColor.RGB:= $F6EFE9;
  XLSExcel.CmdFormat.Apply(0,0);
end;

procedure TForm1.formatoSubtitulos(col, row: Integer);
begin
  //formato para subtitulos de la columna cero
  XLSExcel.CmdFormat.BeginEdit(XLSExcel[0]);
  XLSExcel.CmdFormat.Fill.BackgroundColor.RGB:= $EFE9F6;
  XLSExcel.CmdFormat.Font.Style:= [xfsBold];
  XLSExcel.CmdFormat.Font.Name:= 'Arial';
  XLSExcel.CmdFormat.Font.Size:= 10;
  XLSExcel.CmdFormat.Alignment.Horizontal:= chaLeft;
  for col := 0 to 5 do   //recorrer columnas
    XLSExcel.CmdFormat.Apply(col,row);
end;

procedure TForm1.formatoTipoAcumulado;
var
  i,j,k,l,m,n,o,p,fil_1,fil_2,fil_3,fil_4,fil_5,  y,x: Integer;
  nomAcum1,nomAcum2,nomAcum3,nomAcum4,nomAcum5,nomAcum6: string;
  arrayStrGp2: TArray<Variant>;
begin
  try
  qry.Open;
    with qry do begin
      j:=0; k:=0; l:=0; m:=0; n:=0; o:=0; p:=0; arrayStrGp2:= nil;
      First;
      while not Eof do begin
        for i:=0 to FieldCount-2 do begin
          case FieldByName('TIPO_ACUM').Value of
            1: begin
                 k:= k+1;
                 nomAcum1:= FieldByName('NOM_ACUM').Value;
               end;
            2: begin
                 l:= l+1;
                 nomAcum2:= FieldByName('NOM_ACUM').Value;
               end;
            3: begin
                 m:= m+1;
                 nomAcum3:= FieldByName('NOM_ACUM').Value;
               end;
            4: begin
                 n:= n+1;
                 nomAcum4:= FieldByName('NOM_ACUM').Value;
               end;
            5: begin
                 o:= o+1;
                 nomAcum5:= FieldByName('NOM_ACUM').Value;
//                 SetLength(arrayStrGp2,o);
//                 arrayStrGp2[o]:= VarToStr(FieldByName('GRUPO_3').Value);
//                 ShowMessage('regis ' + VarToStr(arrayStrGp2[o]));
//                 ShowMessage('regis ' + arrayStrGp2[o]);
               end;
            6: begin
                 p:= p+1;
                 nomAcum6:= FieldByName('NOM_ACUM').Value;
               end;
          end;
          Next;
        end;

//        for y := 0 to qry.FieldCount-1 do begin
//          if qry.FieldByName('TIPO_ACUM').Value = 5 then begin
//            arrayStrGp2[o]:= qry.FieldByName('GRUPO_2').Value;
//          end;
//          ShowMessage('regis ' + arrayStrGp2[y]);
//          Next
//        end;
      end;
    end;

//    SetLength(arrayStrGp2, o);
//     x:= High(arrayStrGp2)+1;
//      for y:= 0 to x-1 do begin
//        arrayStrGp2:= qry.FieldByName('GRUPO_2').Value;
//        ShowMessage('regis ' + arrayStrGp2[y]);
//      end;

//    for o := 0 to o-1 do begin
//      ShowMessage('regis ' + VarToStr(arrayStrGp2[o]));
//    end;

    //ShowMessage('regis ' + IntToStr(o));

    XLSExcel[0].AsString[1,1]:= 'Total exámenes'; //titulo Total exámenes

    j:=2; // 2 posiciones despues de la fila 0
    XLSExcel[0].AsString[0,j-1]:= nomAcum1; //titulo seccion
    formatoSubtitulos(0,j-1);

    fil_1:= k+j;//+4 lineas iniciales
    XLSExcel.InsertRows(0,fil_1,2);
    XLSExcel[0].AsString[0,fil_1+1]:= nomAcum2; //titulo examen
    formatoSubtitulos(0,fil_1+1);

    fil_2:= l+fil_1+2;//+2 -> filas añadidas
    XLSExcel.InsertRows(0,fil_2,2);
    XLSExcel[0].AsString[0,fil_2+1]:= nomAcum3; //titulo servicio
    formatoSubtitulos(0,fil_2+1);

    fil_3:= m+fil_2+2;//+2 -> filas añadidas
    XLSExcel.InsertRows(0,fil_3,2);
    XLSExcel[0].AsString[0,fil_3+1]:= nomAcum4; //titulo Usuario que validó
    formatoSubtitulos(0,fil_3+1);

    fil_4:= n+fil_3+2;//+2 -> filas añadidas
    XLSExcel.InsertRows(0,fil_4,2);
    XLSExcel[0].AsString[0,fil_4+1]:= nomAcum5; //titulo Usuario que validó y sección
    formatoSubtitulos(0,fil_4+1);

    fil_5:= o+fil_4+2;//+2 -> filas añadidas
    XLSExcel.InsertRows(0,fil_5,2);
    XLSExcel[0].AsString[0,fil_5+1]:= nomAcum6; //titulo Usuario que validó y examen
    formatoSubtitulos(0,fil_5+1);
  finally
    qry.Close;
  end;
end;

procedure TForm1.FormCreate(Sender: TObject);
begin
  IBQuery1.Open;
end;

procedure TForm1.FormDestroy(Sender: TObject);
begin
  IBQuery1.Close;
end;

procedure TForm1.cargarTitulos;
var
  j: integer;
  anchos:array of integer;
  titulo: array of string;
begin
  with tabla do
  begin
    titulo:= ['GRUPO_1','CANT_EXA','PROM_SEG_TOMA','PROM_SEG_REPORT','HMS_TOMA','HMS_REPORT'];
    anchos:= [300,100,150,150,100,100];
    ColCount:= High(titulo)+1;
    for j:=0 to (ColCount-1) do //recorremos los titulos
    begin
      Rows[1].Add(titulo[j]);
      formatoTitulos(j,Row);
      ColWidths[j]:= anchos[j];
    end;
  end;
end;

procedure TForm1.limpiarArreglo(arrayStr: array of Variant);
var
  i: Integer;
begin
  for i := Low(arrayStr) to High(arrayStr) do
    arrayStr[i]:= EmptyStr;
end;

procedure TForm1.llenarTabla;
var
  col, fil, C, R: Integer;
  Ref: string;
  Cnt: integer;
  CellType: TXLSCellType;
  I: Integer;
  FrmAlig: TXLSCmdFormatAlignment;
begin
  for fil := 0 to 163 do begin  //hay que establecer las filas del stringGrid
    for col := 0 to 10 do begin //hay que establecer las columnas del stringGrid
       XLSExcel[0].AsString[col,fil]:= tabla.Cells[col,fil];
       XLSExcel.Calculate;
       Cnt := 0;
      XLSExcel[0].CalcDimensions;
        for R := (XLSExcel[0].FirstRow) to XLSExcel[0].LastRow do begin
          for C := XLSExcel[0].FirstCol to XLSExcel[0].LastCol do begin
          Inc(Cnt);
        end;
      end;
    end;
  end;
//  formatoEncabezado;
  formatoTipoAcumulado;
  escribirExcel;
end;

procedure TForm1.precargarHojaExcel;
var
  c,f: Integer;
begin
  //inicalizamos la hoja de excel
  for c := 0 to 16 do //columnas
    tabla.Cells[c,0] := ColRowToRefStr(c,0);
  for f := 0 to 300 do //filas
    tabla.Cells[0,f] := ColRowToRefStr(0,f);
end;





end.

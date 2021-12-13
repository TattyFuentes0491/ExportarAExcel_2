object Form1: TForm1
  Left = 0
  Top = 0
  Width = 900
  Height = 448
  AutoScroll = True
  Caption = 'Exportar a Excel - XLSReadWrite5 II'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  OnCreate = FormCreate
  OnDestroy = FormDestroy
  PixelsPerInch = 96
  TextHeight = 13
  object Button1: TButton
    Left = 8
    Top = 0
    Width = 241
    Height = 40
    Caption = 'Cargar Datos de un archivo Excel'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -12
    Font.Name = 'Tahoma'
    Font.Style = [fsBold]
    ParentFont = False
    TabOrder = 0
    OnClick = Button1Click
  end
  object Button2: TButton
    Left = 266
    Top = 0
    Width = 200
    Height = 40
    Caption = 'Exportar Archivo Excel'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -12
    Font.Name = 'Tahoma'
    Font.Style = [fsBold]
    ParentFont = False
    TabOrder = 1
    OnClick = Button2Click
  end
  object Button3: TButton
    Left = 488
    Top = 0
    Width = 120
    Height = 40
    Caption = 'Limpiar Tabla'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -12
    Font.Name = 'Tahoma'
    Font.Style = [fsBold]
    ParentFont = False
    TabOrder = 2
    OnClick = Button3Click
  end
  object tabla: TStringGrid
    Left = 0
    Top = 46
    Width = 884
    Height = 363
    Align = alBottom
    ColCount = 6
    FixedCols = 0
    RowCount = 170
    TabOrder = 3
  end
  object Button4: TButton
    Left = 632
    Top = 0
    Width = 113
    Height = 40
    Caption = 'Abrir Archivo'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -12
    Font.Name = 'Tahoma'
    Font.Style = [fsBold]
    ParentFont = False
    TabOrder = 4
    OnClick = Button4Click
  end
  object Button5: TButton
    Left = 768
    Top = 8
    Width = 75
    Height = 25
    Caption = 'Button5'
    TabOrder = 5
    OnClick = Button5Click
  end
  object XLSExcel: TXLSReadWriteII5
    ComponentVersion = '6.01.10a'
    Version = xvExcel2007
    DirectRead = False
    DirectWrite = False
    DoNotReadSheets = False
    Left = 696
    Top = 72
  end
  object IBDatabase1: TIBDatabase
    Connected = True
    DatabaseName = 'localhost:D:\SL_SOFTWARE\Lab Dilab\DATABASE\DBLAB_DILAB.FDB'
    Params.Strings = (
      'user_name=sysdba'
      'password=masterkey'
      'lc_ctype=ISO8859_1')
    LoginPrompt = False
    ServerType = 'IBServer'
    Left = 24
    Top = 72
  end
  object IBQuery1: TIBQuery
    Database = IBDatabase1
    Transaction = IBTransaction1
    Active = True
    BufferChunks = 1000
    CachedUpdates = False
    ParamCheck = True
    SQL.Strings = (
      'WITH TIEMPOS_REP AS ('
      #9'SELECT RC.FECHA_RECEPCION'
      
        #9'  , IIF(RL.MUESTRAENTREGADA = '#39'T'#39', COALESCE(RL.FECHAHORAMUESTRA' +
        ', RC.FECHA_RECEPCION), RC.FECHA_RECEPCION) AS FECHA_MUESTRA'
      
        #9'  , IIF(RL.EST_DIGITAL = 1, RL.FECHA_PDF, IIF(RL.REPORTADO = '#39'T' +
        #39', RL.FECHA_REPORTADO, NULL)) AS FECHA_REPORTADO'
      
        #9'  , RL.REPORTADO, RL.VALIDADO, RL.EST_DIGITAL, RL.ID_PDF, RL.FE' +
        'CHA_PDF'
      #9'  , DATEDIFF(SECOND FROM RC.FECHA_RECEPCION'
      
        #9#9#9#9'TO IIF(RL.MUESTRAENTREGADA = '#39'T'#39', COALESCE(RL.FECHAHORAMUEST' +
        'RA, RC.FECHA_RECEPCION), RC.FECHA_RECEPCION)) AS SEG_TOMA'
      
        #9'  , DATEDIFF(SECOND FROM IIF(RL.MUESTRAENTREGADA = '#39'T'#39', COALESC' +
        'E(RL.FECHAHORAMUESTRA, RC.FECHA_RECEPCION), RC.FECHA_RECEPCION)'
      
        #9#9#9#9'TO IIF(RL.EST_DIGITAL = 1, RL.FECHA_PDF, IIF(RL.REPORTADO = ' +
        #39'T'#39', RL.FECHA_REPORTADO, NULL))) AS SEG_REPORT'
      #9'  , RL.IDRECEPCION, RL.COD_EXAMEN, EX.NOMBRE AS NOM_EXAMEN'
      
        #9'  , IIF(RL.EST_DIGITAL = 1, '#39'web_adjunto'#39', RL.ACEPTADOPOR) AS A' +
        'CEPTADOPOR'
      
        #9'  , RC.COD_SALA, SA.NOM_SALA, PT.COD_SECCION, SE.NOMBRE AS NOM_' +
        'SECCION'
      
        #9'  , IIF(RL.EST_DIGITAL = 1, '#39'PDF Adjuntado en Portal Web'#39', (BC.' +
        'NOMBRES||'#39' '#39'||BC.APELLIDOS)) AS NOM_BACT_VALIDA'
      #9'FROM RELACION RL'
      
        #9'  INNER JOIN RECEPCION RC     ON  RC.IDRECEPCION = RL.IDRECEPCI' +
        'ON'
      #9'  INNER JOIN EXAMEN EX        ON  RL.COD_EXAMEN = EX.CODIGO'
      #9'  LEFT JOIN  SALA SA          ON  RC.COD_SALA = SA.COD_SALA'
      #9'  LEFT JOIN  PROTOCOLO PT     ON  EX.COD_PROTOCOLO = PT.CODIGO'
      #9'  LEFT JOIN  SECCION SE       ON  PT.COD_SECCION = SE.CODIGO'
      #9'  LEFT JOIN  BACTERIOLOGO BC  ON  RL.ACEPTADOPOR = BC.USUARIO'
      #9'WHERE RC.FECHA_RECEPCION >= '#39'01.10.2021'#39
      #9'  AND RC.FECHA_RECEPCION <  '#39'31.10.2021'#39
      #9'ORDER BY RL.IDRECEPCION, RL.COD_EXAMEN'
      ')'
      ''
      
        '  SELECT 1 AS TIPO_ACUM, CAST('#39'Secci'#243'n'#39' AS VARCHAR(35)) AS NOM_A' +
        'CUM'
      
        '    , TR.NOM_SECCION AS GRUPO_1, TR.NOM_SECCION AS GRUPO_2, TR.N' +
        'OM_SECCION AS GRUPO_3'
      
        '    , COUNT(*) AS CANT_EXA, CAST(AVG(TR.SEG_TOMA) AS NUMERIC) AS' +
        ' PROM_SEG_TOMA, CAST(AVG(TR.SEG_REPORT) AS NUMERIC) AS PROM_SEG_' +
        'REPORT'
      
        '    , (AVG(TR.SEG_TOMA)/3600)||'#39':'#39'||(MOD(AVG(TR.SEG_TOMA),3600)/' +
        '60)||'#39':'#39'||MOD(AVG(TR.SEG_TOMA),60) AS HMS_TOMA'
      
        '    , (AVG(TR.SEG_REPORT)/3600)||'#39':'#39'||(MOD(AVG(TR.SEG_REPORT),36' +
        '00)/60)||'#39':'#39'||MOD(AVG(TR.SEG_REPORT),60) AS HMS_REPORT'
      '  FROM TIEMPOS_REP TR'
      '  GROUP BY TIPO_ACUM, NOM_ACUM, GRUPO_1, GRUPO_2, GRUPO_3'
      ''
      '  UNION'
      
        '  SELECT 2 AS TIPO_ACUM, CAST('#39'Examen'#39' AS VARCHAR(35)) AS NOM_AC' +
        'UM'
      
        '    , TR.NOM_EXAMEN AS GRUPO_1, TR.NOM_EXAMEN AS GRUPO_2, TR.NOM' +
        '_EXAMEN AS GRUPO_3'
      
        '    , COUNT(*) AS CANT_EXA, CAST(AVG(TR.SEG_TOMA) AS NUMERIC) AS' +
        ' PROM_SEG_TOMA, CAST(AVG(TR.SEG_REPORT) AS NUMERIC) AS PROM_SEG_' +
        'REPORT'
      
        '    , (AVG(TR.SEG_TOMA)/3600)||'#39':'#39'||(MOD(AVG(TR.SEG_TOMA),3600)/' +
        '60)||'#39':'#39'||MOD(AVG(TR.SEG_TOMA),60) AS HMS_TOMA'
      
        '    , (AVG(TR.SEG_REPORT)/3600)||'#39':'#39'||(MOD(AVG(TR.SEG_REPORT),36' +
        '00)/60)||'#39':'#39'||MOD(AVG(TR.SEG_REPORT),60) AS HMS_REPORT'
      '  FROM TIEMPOS_REP TR'
      '  GROUP BY TIPO_ACUM, NOM_ACUM, GRUPO_1, GRUPO_2, GRUPO_3'
      ''
      '  UNION'
      
        '  SELECT 3 AS TIPO_ACUM, CAST('#39'Servicio'#39' AS VARCHAR(35)) AS NOM_' +
        'ACUM'
      
        '    , TR.NOM_SALA AS GRUPO_1, TR.NOM_SALA AS GRUPO_2, TR.NOM_SAL' +
        'A AS GRUPO_3'
      
        '    , COUNT(*) AS CANT_EXA, CAST(AVG(TR.SEG_TOMA) AS NUMERIC) AS' +
        ' PROM_SEG_TOMA, CAST(AVG(TR.SEG_REPORT) AS NUMERIC) AS PROM_SEG_' +
        'REPORT'
      
        '    , (AVG(TR.SEG_TOMA)/3600)||'#39':'#39'||(MOD(AVG(TR.SEG_TOMA),3600)/' +
        '60)||'#39':'#39'||MOD(AVG(TR.SEG_TOMA),60) AS HMS_TOMA'
      
        '    , (AVG(TR.SEG_REPORT)/3600)||'#39':'#39'||(MOD(AVG(TR.SEG_REPORT),36' +
        '00)/60)||'#39':'#39'||MOD(AVG(TR.SEG_REPORT),60) AS HMS_REPORT'
      '  FROM TIEMPOS_REP TR'
      '  GROUP BY TIPO_ACUM, NOM_ACUM, GRUPO_1, GRUPO_2, GRUPO_3'
      ''
      '  UNION'
      
        '  SELECT 4 AS TIPO_ACUM, CAST('#39'Usuario que valid'#243#39' AS VARCHAR(35' +
        ')) AS NOM_ACUM'
      
        '    , TR.ACEPTADOPOR AS GRUPO_1, TR.NOM_BACT_VALIDA AS GRUPO_2, ' +
        'TR.NOM_BACT_VALIDA AS GRUPO_3'
      
        '    , COUNT(*) AS CANT_EXA, CAST(AVG(TR.SEG_TOMA) AS NUMERIC) AS' +
        ' PROM_SEG_TOMA, CAST(AVG(TR.SEG_REPORT) AS NUMERIC) AS PROM_SEG_' +
        'REPORT'
      
        '    , (AVG(TR.SEG_TOMA)/3600)||'#39':'#39'||(MOD(AVG(TR.SEG_TOMA),3600)/' +
        '60)||'#39':'#39'||MOD(AVG(TR.SEG_TOMA),60) AS HMS_TOMA'
      
        '    , (AVG(TR.SEG_REPORT)/3600)||'#39':'#39'||(MOD(AVG(TR.SEG_REPORT),36' +
        '00)/60)||'#39':'#39'||MOD(AVG(TR.SEG_REPORT),60) AS HMS_REPORT'
      '  FROM TIEMPOS_REP TR'
      '  GROUP BY TIPO_ACUM, NOM_ACUM, GRUPO_1, GRUPO_2, GRUPO_3'
      ''
      '  UNION'
      
        '  SELECT 5 AS TIPO_ACUM, CAST('#39'Usuario que valid'#243' y secci'#243'n'#39' AS ' +
        'VARCHAR(35)) AS NOM_ACUM'
      
        '    , TR.ACEPTADOPOR AS GRUPO_1, TR.NOM_BACT_VALIDA AS GRUPO_2, ' +
        'TR.NOM_SECCION AS GRUPO_3'
      
        '    , COUNT(*) AS CANT_EXA, CAST(AVG(TR.SEG_TOMA) AS NUMERIC) AS' +
        ' PROM_SEG_TOMA, CAST(AVG(TR.SEG_REPORT) AS NUMERIC) AS PROM_SEG_' +
        'REPORT'
      
        '    , (AVG(TR.SEG_TOMA)/3600)||'#39':'#39'||(MOD(AVG(TR.SEG_TOMA),3600)/' +
        '60)||'#39':'#39'||MOD(AVG(TR.SEG_TOMA),60) AS HMS_TOMA'
      
        '    , (AVG(TR.SEG_REPORT)/3600)||'#39':'#39'||(MOD(AVG(TR.SEG_REPORT),36' +
        '00)/60)||'#39':'#39'||MOD(AVG(TR.SEG_REPORT),60) AS HMS_REPORT'
      '  FROM TIEMPOS_REP TR'
      '  GROUP BY TIPO_ACUM, NOM_ACUM, GRUPO_1, GRUPO_2, GRUPO_3'
      ''
      '  UNION'
      
        '  SELECT 6 AS TIPO_ACUM, CAST('#39'Usuario que valid'#243' y examen'#39' AS V' +
        'ARCHAR(35)) AS NOM_ACUM'
      
        '    , TR.ACEPTADOPOR AS GRUPO_1, TR.NOM_BACT_VALIDA AS GRUPO_2, ' +
        'TR.NOM_EXAMEN AS GRUPO_3'
      
        '    , COUNT(*) AS CANT_EXA, CAST(AVG(TR.SEG_TOMA) AS NUMERIC) AS' +
        ' PROM_SEG_TOMA, CAST(AVG(TR.SEG_REPORT) AS NUMERIC) AS PROM_SEG_' +
        'REPORT'
      
        '    , (AVG(TR.SEG_TOMA)/3600)||'#39':'#39'||(MOD(AVG(TR.SEG_TOMA),3600)/' +
        '60)||'#39':'#39'||MOD(AVG(TR.SEG_TOMA),60) AS HMS_TOMA'
      
        '    , (AVG(TR.SEG_REPORT)/3600)||'#39':'#39'||(MOD(AVG(TR.SEG_REPORT),36' +
        '00)/60)||'#39':'#39'||MOD(AVG(TR.SEG_REPORT),60) AS HMS_REPORT'
      '  FROM TIEMPOS_REP TR'
      '  GROUP BY TIPO_ACUM, NOM_ACUM, GRUPO_1, GRUPO_2, GRUPO_3')
    Left = 88
    Top = 72
    object IBQuery1GRUPO_1: TIBStringField
      FieldName = 'GRUPO_1'
      ProviderFlags = []
      Size = 60
    end
    object IBQuery1CANT_EXA: TIntegerField
      FieldName = 'CANT_EXA'
      ProviderFlags = []
    end
    object IBQuery1PROM_SEG_TOMA: TIntegerField
      FieldName = 'PROM_SEG_TOMA'
      ProviderFlags = []
    end
    object IBQuery1PROM_SEG_REPORT: TIntegerField
      FieldName = 'PROM_SEG_REPORT'
      ProviderFlags = []
    end
    object IBQuery1HMS_TOMA: TIBStringField
      FieldName = 'HMS_TOMA'
      ProviderFlags = []
      Size = 62
    end
    object IBQuery1HMS_REPORT: TIBStringField
      FieldName = 'HMS_REPORT'
      ProviderFlags = []
      Size = 62
    end
  end
  object IBTransaction1: TIBTransaction
    Active = True
    DefaultDatabase = IBDatabase1
    Left = 160
    Top = 72
  end
  object XPManifest1: TXPManifest
    Left = 600
    Top = 104
  end
  object qry: TIBQuery
    Database = IBDatabase1
    Transaction = IBTransaction2
    BufferChunks = 1000
    CachedUpdates = False
    ParamCheck = True
    SQL.Strings = (
      'WITH TIEMPOS_REP AS ('
      #9'SELECT RC.FECHA_RECEPCION'
      
        #9'  , IIF(RL.MUESTRAENTREGADA = '#39'T'#39', COALESCE(RL.FECHAHORAMUESTRA' +
        ', RC.FECHA_RECEPCION), RC.FECHA_RECEPCION) AS FECHA_MUESTRA'
      
        #9'  , IIF(RL.EST_DIGITAL = 1, RL.FECHA_PDF, IIF(RL.REPORTADO = '#39'T' +
        #39', RL.FECHA_REPORTADO, NULL)) AS FECHA_REPORTADO'
      
        #9'  , RL.REPORTADO, RL.VALIDADO, RL.EST_DIGITAL, RL.ID_PDF, RL.FE' +
        'CHA_PDF'
      #9'  , DATEDIFF(SECOND FROM RC.FECHA_RECEPCION'
      
        #9#9#9#9'TO IIF(RL.MUESTRAENTREGADA = '#39'T'#39', COALESCE(RL.FECHAHORAMUEST' +
        'RA, RC.FECHA_RECEPCION), RC.FECHA_RECEPCION)) AS SEG_TOMA'
      
        #9'  , DATEDIFF(SECOND FROM IIF(RL.MUESTRAENTREGADA = '#39'T'#39', COALESC' +
        'E(RL.FECHAHORAMUESTRA, RC.FECHA_RECEPCION), RC.FECHA_RECEPCION)'
      
        #9#9#9#9'TO IIF(RL.EST_DIGITAL = 1, RL.FECHA_PDF, IIF(RL.REPORTADO = ' +
        #39'T'#39', RL.FECHA_REPORTADO, NULL))) AS SEG_REPORT'
      #9'  , RL.IDRECEPCION, RL.COD_EXAMEN, EX.NOMBRE AS NOM_EXAMEN'
      
        #9'  , IIF(RL.EST_DIGITAL = 1, '#39'web_adjunto'#39', RL.ACEPTADOPOR) AS A' +
        'CEPTADOPOR'
      
        #9'  , RC.COD_SALA, SA.NOM_SALA, PT.COD_SECCION, SE.NOMBRE AS NOM_' +
        'SECCION'
      
        #9'  , IIF(RL.EST_DIGITAL = 1, '#39'PDF Adjuntado en Portal Web'#39', (BC.' +
        'NOMBRES||'#39' '#39'||BC.APELLIDOS)) AS NOM_BACT_VALIDA'
      #9'FROM RELACION RL'
      
        #9'  INNER JOIN RECEPCION RC     ON  RC.IDRECEPCION = RL.IDRECEPCI' +
        'ON'
      #9'  INNER JOIN EXAMEN EX        ON  RL.COD_EXAMEN = EX.CODIGO'
      #9'  LEFT JOIN  SALA SA          ON  RC.COD_SALA = SA.COD_SALA'
      #9'  LEFT JOIN  PROTOCOLO PT     ON  EX.COD_PROTOCOLO = PT.CODIGO'
      #9'  LEFT JOIN  SECCION SE       ON  PT.COD_SECCION = SE.CODIGO'
      #9'  LEFT JOIN  BACTERIOLOGO BC  ON  RL.ACEPTADOPOR = BC.USUARIO'
      #9'WHERE RC.FECHA_RECEPCION >= '#39'01.10.2021'#39
      #9'  AND RC.FECHA_RECEPCION <  '#39'31.10.2021'#39
      #9'ORDER BY RL.IDRECEPCION, RL.COD_EXAMEN'
      ')'
      ''
      
        '  SELECT 1 AS TIPO_ACUM, CAST('#39'Secci'#243'n'#39' AS VARCHAR(35)) AS NOM_A' +
        'CUM'
      
        '    , TR.NOM_SECCION AS GRUPO_1, TR.NOM_SECCION AS GRUPO_2, TR.N' +
        'OM_SECCION AS GRUPO_3'
      
        '    , COUNT(*) AS CANT_EXA, CAST(AVG(TR.SEG_TOMA) AS NUMERIC) AS' +
        ' PROM_SEG_TOMA, CAST(AVG(TR.SEG_REPORT) AS NUMERIC) AS PROM_SEG_' +
        'REPORT'
      
        '    , (AVG(TR.SEG_TOMA)/3600)||'#39':'#39'||(MOD(AVG(TR.SEG_TOMA),3600)/' +
        '60)||'#39':'#39'||MOD(AVG(TR.SEG_TOMA),60) AS HMS_TOMA'
      
        '    , (AVG(TR.SEG_REPORT)/3600)||'#39':'#39'||(MOD(AVG(TR.SEG_REPORT),36' +
        '00)/60)||'#39':'#39'||MOD(AVG(TR.SEG_REPORT),60) AS HMS_REPORT'
      '  FROM TIEMPOS_REP TR'
      '  GROUP BY TIPO_ACUM, NOM_ACUM, GRUPO_1, GRUPO_2, GRUPO_3'
      ''
      '  UNION'
      
        '  SELECT 2 AS TIPO_ACUM, CAST('#39'Examen'#39' AS VARCHAR(35)) AS NOM_AC' +
        'UM'
      
        '    , TR.NOM_EXAMEN AS GRUPO_1, TR.NOM_EXAMEN AS GRUPO_2, TR.NOM' +
        '_EXAMEN AS GRUPO_3'
      
        '    , COUNT(*) AS CANT_EXA, CAST(AVG(TR.SEG_TOMA) AS NUMERIC) AS' +
        ' PROM_SEG_TOMA, CAST(AVG(TR.SEG_REPORT) AS NUMERIC) AS PROM_SEG_' +
        'REPORT'
      
        '    , (AVG(TR.SEG_TOMA)/3600)||'#39':'#39'||(MOD(AVG(TR.SEG_TOMA),3600)/' +
        '60)||'#39':'#39'||MOD(AVG(TR.SEG_TOMA),60) AS HMS_TOMA'
      
        '    , (AVG(TR.SEG_REPORT)/3600)||'#39':'#39'||(MOD(AVG(TR.SEG_REPORT),36' +
        '00)/60)||'#39':'#39'||MOD(AVG(TR.SEG_REPORT),60) AS HMS_REPORT'
      '  FROM TIEMPOS_REP TR'
      '  GROUP BY TIPO_ACUM, NOM_ACUM, GRUPO_1, GRUPO_2, GRUPO_3'
      ''
      '  UNION'
      
        '  SELECT 3 AS TIPO_ACUM, CAST('#39'Servicio'#39' AS VARCHAR(35)) AS NOM_' +
        'ACUM'
      
        '    , TR.NOM_SALA AS GRUPO_1, TR.NOM_SALA AS GRUPO_2, TR.NOM_SAL' +
        'A AS GRUPO_3'
      
        '    , COUNT(*) AS CANT_EXA, CAST(AVG(TR.SEG_TOMA) AS NUMERIC) AS' +
        ' PROM_SEG_TOMA, CAST(AVG(TR.SEG_REPORT) AS NUMERIC) AS PROM_SEG_' +
        'REPORT'
      
        '    , (AVG(TR.SEG_TOMA)/3600)||'#39':'#39'||(MOD(AVG(TR.SEG_TOMA),3600)/' +
        '60)||'#39':'#39'||MOD(AVG(TR.SEG_TOMA),60) AS HMS_TOMA'
      
        '    , (AVG(TR.SEG_REPORT)/3600)||'#39':'#39'||(MOD(AVG(TR.SEG_REPORT),36' +
        '00)/60)||'#39':'#39'||MOD(AVG(TR.SEG_REPORT),60) AS HMS_REPORT'
      '  FROM TIEMPOS_REP TR'
      '  GROUP BY TIPO_ACUM, NOM_ACUM, GRUPO_1, GRUPO_2, GRUPO_3'
      ''
      '  UNION'
      
        '  SELECT 4 AS TIPO_ACUM, CAST('#39'Usuario que valid'#243#39' AS VARCHAR(35' +
        ')) AS NOM_ACUM'
      
        '    , TR.ACEPTADOPOR AS GRUPO_1, TR.NOM_BACT_VALIDA AS GRUPO_2, ' +
        'TR.NOM_BACT_VALIDA AS GRUPO_3'
      
        '    , COUNT(*) AS CANT_EXA, CAST(AVG(TR.SEG_TOMA) AS NUMERIC) AS' +
        ' PROM_SEG_TOMA, CAST(AVG(TR.SEG_REPORT) AS NUMERIC) AS PROM_SEG_' +
        'REPORT'
      
        '    , (AVG(TR.SEG_TOMA)/3600)||'#39':'#39'||(MOD(AVG(TR.SEG_TOMA),3600)/' +
        '60)||'#39':'#39'||MOD(AVG(TR.SEG_TOMA),60) AS HMS_TOMA'
      
        '    , (AVG(TR.SEG_REPORT)/3600)||'#39':'#39'||(MOD(AVG(TR.SEG_REPORT),36' +
        '00)/60)||'#39':'#39'||MOD(AVG(TR.SEG_REPORT),60) AS HMS_REPORT'
      '  FROM TIEMPOS_REP TR'
      '  GROUP BY TIPO_ACUM, NOM_ACUM, GRUPO_1, GRUPO_2, GRUPO_3'
      ''
      '  UNION'
      
        '  SELECT 5 AS TIPO_ACUM, CAST('#39'Usuario que valid'#243' y secci'#243'n'#39' AS ' +
        'VARCHAR(35)) AS NOM_ACUM'
      
        '    , TR.ACEPTADOPOR AS GRUPO_1, TR.NOM_BACT_VALIDA AS GRUPO_2, ' +
        'TR.NOM_SECCION AS GRUPO_3'
      
        '    , COUNT(*) AS CANT_EXA, CAST(AVG(TR.SEG_TOMA) AS NUMERIC) AS' +
        ' PROM_SEG_TOMA, CAST(AVG(TR.SEG_REPORT) AS NUMERIC) AS PROM_SEG_' +
        'REPORT'
      
        '    , (AVG(TR.SEG_TOMA)/3600)||'#39':'#39'||(MOD(AVG(TR.SEG_TOMA),3600)/' +
        '60)||'#39':'#39'||MOD(AVG(TR.SEG_TOMA),60) AS HMS_TOMA'
      
        '    , (AVG(TR.SEG_REPORT)/3600)||'#39':'#39'||(MOD(AVG(TR.SEG_REPORT),36' +
        '00)/60)||'#39':'#39'||MOD(AVG(TR.SEG_REPORT),60) AS HMS_REPORT'
      '  FROM TIEMPOS_REP TR'
      '  GROUP BY TIPO_ACUM, NOM_ACUM, GRUPO_1, GRUPO_2, GRUPO_3'
      ''
      '  UNION'
      
        '  SELECT 6 AS TIPO_ACUM, CAST('#39'Usuario que valid'#243' y examen'#39' AS V' +
        'ARCHAR(35)) AS NOM_ACUM'
      
        '    , TR.ACEPTADOPOR AS GRUPO_1, TR.NOM_BACT_VALIDA AS GRUPO_2, ' +
        'TR.NOM_EXAMEN AS GRUPO_3'
      
        '    , COUNT(*) AS CANT_EXA, CAST(AVG(TR.SEG_TOMA) AS NUMERIC) AS' +
        ' PROM_SEG_TOMA, CAST(AVG(TR.SEG_REPORT) AS NUMERIC) AS PROM_SEG_' +
        'REPORT'
      
        '    , (AVG(TR.SEG_TOMA)/3600)||'#39':'#39'||(MOD(AVG(TR.SEG_TOMA),3600)/' +
        '60)||'#39':'#39'||MOD(AVG(TR.SEG_TOMA),60) AS HMS_TOMA'
      
        '    , (AVG(TR.SEG_REPORT)/3600)||'#39':'#39'||(MOD(AVG(TR.SEG_REPORT),36' +
        '00)/60)||'#39':'#39'||MOD(AVG(TR.SEG_REPORT),60) AS HMS_REPORT'
      '  FROM TIEMPOS_REP TR'
      '  GROUP BY TIPO_ACUM, NOM_ACUM, GRUPO_1, GRUPO_2, GRUPO_3')
    Left = 520
    Top = 200
  end
  object IBTransaction2: TIBTransaction
    DefaultDatabase = IBDatabase1
    Left = 520
    Top = 256
  end
end

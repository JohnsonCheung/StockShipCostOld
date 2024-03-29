Attribute VB_Name = "MAdo_Cat_Tbl"
Option Compare Database
Option Explicit
Function CatTbl(A As Catalog, W) As ADOX.Table
Set CatTbl = A.Tables(W & "$")
End Function

Function CatTblFny(A As ADOX.Table) As String()
CatTblFny = ItrNy(A.Columns)
End Function

Function CatTblTyAy(A As ADOX.Table) As String()
CatTblTyAy = ItrMapSy(A.Columns, "CatColTy")
End Function

Function CatColTy$(A As ADOX.Column)
CatColTy = AdoTyTy(A.Type)
End Function

Function AdoTyTy$(A As ADODB.DataTypeEnum)
Dim O$
Select Case A
Case ADODB.DataTypeEnum.adTinyInt: O = "Byt"
Case ADODB.DataTypeEnum.adInteger: O = "Lng"
Case ADODB.DataTypeEnum.adSmallInt: O = "Int"
Case ADODB.DataTypeEnum.adDate: O = "Dte"
Case ADODB.DataTypeEnum.adVarChar: O = "Txt"
Case ADODB.DataTypeEnum.adBoolean: O = "Yes"
Case ADODB.DataTypeEnum.adDouble: O = "Dbl"
Case ADODB.DataTypeEnum.adCurrency: O = "Cur"
Case ADODB.DataTypeEnum.adSingle: O = "Sng"
Case ADODB.DataTypeEnum.adDecimal: O = "Dec"
Case ADODB.DataTypeEnum.adVarWChar: O = "Mem"
Case Else: O = "?" & A & "?"
End Select
AdoTyTy = O
End Function

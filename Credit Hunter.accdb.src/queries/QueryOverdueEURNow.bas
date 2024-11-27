Operation =1
Option =0
Where ="(((Tbl_Invoices.Overdue_Date)<=Now()) And ((Tbl_Invoices.Currency)=\"EUR\"))"
Having ="(((Tbl_Invoices.Customer_ID)=[CUSTID]) And ((Tbl_Invoices.Update_date)=[date]))"
Begin InputTables
    Name ="Tbl_Invoices"
End
Begin OutputColumns
    Alias ="SommaDiAmount"
    Expression ="Sum(Tbl_Invoices.Amount)"
End
Begin Groups
    Expression ="Tbl_Invoices.Customer_ID"
    GroupLevel =0
    Expression ="Tbl_Invoices.Update_date"
    GroupLevel =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="SommaDiAmount"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =0
    Top =0
    Right =1705
    Bottom =837
    Left =-1
    Top =-1
    Right =1689
    Bottom =305
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =38
        Top =6
        Right =275
        Bottom =242
        Top =0
        Name ="Tbl_Invoices"
        Name =""
    End
End

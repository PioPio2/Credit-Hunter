SELECT
  Sum(Tbl_Invoices.Amount) AS SommaDiAmount
FROM
  Tbl_Invoices
WHERE
  (
    (
      (Tbl_Invoices.Overdue_Date)<= Now()
    )
    And (
      (Tbl_Invoices.Currency)= "EUR"
    )
  )
GROUP BY
  Tbl_Invoices.Customer_ID,
  Tbl_Invoices.Update_date
HAVING
  (
    (
      (Tbl_Invoices.Customer_ID)= [CUSTID]
    )
    And (
      (Tbl_Invoices.Update_date)= [date]
    )
  );

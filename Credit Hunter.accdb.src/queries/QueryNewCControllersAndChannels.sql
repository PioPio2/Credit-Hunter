INSERT INTO Tbl_Cash_Target (Channel, CControllerID)
SELECT
  QueryCControllersInUseToday.RetailOEM,
  Tbl_Users.ID
FROM
  (
    QueryCControllersInUseToday
    LEFT JOIN Tbl_Cash_Target ON (
      QueryCControllersInUseToday.Credit_controller = Tbl_Cash_Target.CControllerID
    )
    AND (
      QueryCControllersInUseToday.RetailOEM = Tbl_Cash_Target.Channel
    )
  )
  LEFT JOIN Tbl_Users ON QueryCControllersInUseToday.Credit_controller = Tbl_Users.ID
WHERE
  (
    (
      (
        Tbl_Cash_Target.CashTargetInEUR
      ) Is Null
    )
  );

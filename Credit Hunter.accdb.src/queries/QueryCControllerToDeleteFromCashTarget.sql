DELETE DISTINCTROW Tbl_Cash_Target.*,
QueryCControllersInUseToday.Credit_controller
FROM
  Tbl_Cash_Target
  LEFT JOIN QueryCControllersInUseToday ON (
    Tbl_Cash_Target.Channel = QueryCControllersInUseToday.RetailOEM
  )
  AND (
    Tbl_Cash_Target.CControllerID = QueryCControllersInUseToday.Credit_controller
  )
WHERE
  (
    (
      (
        QueryCControllersInUseToday.Credit_controller
      ) Is Null
    )
  );

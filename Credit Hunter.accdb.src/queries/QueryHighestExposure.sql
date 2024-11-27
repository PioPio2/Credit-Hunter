SELECT
  Tbl_DailyExposure.CustomerID,
  Max(Tbl_DailyExposure.ARExposure) AS MaxOfARExposure
FROM
  Tbl_DailyExposure
GROUP BY
  Tbl_DailyExposure.CustomerID;

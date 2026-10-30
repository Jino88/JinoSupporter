import DataBrowser from "./qr/DataBrowser";

/** /app/qr/gauss-data — port of QrGaussDataPage.razor. */
export default function QrGaussData() {
  return (
    <DataBrowser
      tableKey="GSTD"
      kicker="GAUSS TEST RESULTS · DB RECORDS"
      description="Select one or more dates to show the matching TCDB.dbo.GSTD rows together. The test time column is OTIME and TestResult is 1 (OK) / 0 (NG)."
    />
  );
}

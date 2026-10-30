import DataBrowser from "./qr/DataBrowser";

/** /app/qr/bako-data — port of QrBakoDataPage.razor. ?date= and ?qr= come from QR SUMMARY. */
export default function QrBakoData() {
  return (
    <DataBrowser
      tableKey="BKTD"
      kicker="BAKO TEST RESULTS · DB RECORDS"
      description="Select one or more dates to show the matching TCDB.dbo.BKTD rows together. TestResult is PASS / FAIL."
    />
  );
}

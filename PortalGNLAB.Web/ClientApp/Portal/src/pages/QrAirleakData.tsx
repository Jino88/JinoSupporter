import DataBrowser from "./qr/DataBrowser";

/** /app/qr/airleak-data — port of QrAirleakDataPage.razor. */
export default function QrAirleakData() {
  return (
    <DataBrowser
      tableKey="ALTD"
      kicker="AIR LEAK TEST RESULTS · DB RECORDS"
      description="Select one or more dates to show the matching TCDB.dbo.ALTD rows together. TestResult is OK / AL (leak)."
    />
  );
}

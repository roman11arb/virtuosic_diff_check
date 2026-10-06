import pandas as pd
import os
import argparse
from datetime import datetime, timedelta


def _report_int(value):
    if isinstance(value, str):
        value = value.replace(",", "").strip()
        if value.startswith("(") and value.endswith(")"):
            value = f"-{value[1:-1]}"
    return int(value)


# If we add tD (from reportingDate) + MtD (from previous day) I get MtD Check then I substract MtD cehck - MtD (from reporting date) = I get the Diff Mtd should be 0
# Also check capital diff, Capital from reporting date - Capital from prev date = If is 0 is Ok


class GetVirtuosic:
    def __init__(self, dataFolder):
        """
        Get the files that we need to perform the diff check
        """
        self.dataFolder = dataFolder
        self.sageCloudFolder = os.path.join(dataFolder, "Sage Cloud")
        self.rawReportFolder = os.path.join(
            dataFolder,
            "DailyReports",
            "Crypto_Daily",
            "DailyReportVirtuosic",
        )
        self.reportFile = None
        self.previousReportFile = None
        self.virtuosicReportPattern = "DailyReportVirtuosicFund_%s.xlsx"

    def getXlsxFile(self, reportDate):
        """Build the Sage Cloud Excel path for the requested report date."""
        return os.path.join(
            self.sageCloudFolder,
            self.virtuosicReportPattern % reportDate,
        )

    def getCurrentCsvFile(self, reportDate):
        """Build the raw-report CSV path for the requested report date."""
        currentCsvName = (
            os.path.splitext(self.virtuosicReportPattern % reportDate)[0] + ".csv"
        )
        return os.path.join(self.rawReportFolder, currentCsvName)

    def getPreviousCsvFile(self, reportDate):
        """Build the raw-report CSV path for the preceding calendar day."""
        previousDate = (
            datetime.strptime(reportDate, "%Y%m%d") - timedelta(days=1)
        ).strftime("%Y%m%d")

        previousCsvName = (
            os.path.splitext(self.virtuosicReportPattern % previousDate)[0] + ".csv"
        )
        return os.path.join(self.rawReportFolder, previousCsvName)

    def checkComplete(self, reportDate):
        """Check for a current source report and the preceding day's raw CSV."""
        currentXlsxPath = self.getXlsxFile(reportDate)
        currentCsvPath = self.getCurrentCsvFile(reportDate)
        previousCsvPath = self.getPreviousCsvFile(reportDate)

        if not os.path.isfile(currentXlsxPath) and not os.path.isfile(currentCsvPath):
            print(f"{os.path.basename(currentXlsxPath)} is not ready")
            return False

        if not os.path.isfile(previousCsvPath):
            print(f"{os.path.basename(previousCsvPath)} is not ready")
            return False

        return True

    def xlsxToCsv(self, xlsxPath):
        """Convert one Sage Cloud workbook into the raw-report folder."""
        csvName = os.path.splitext(os.path.basename(xlsxPath))[0] + ".csv"
        csvPath = os.path.join(self.rawReportFolder, csvName)
        report_df = pd.read_excel(xlsxPath)
        report_df.to_csv(csvPath, index=False)
        os.remove(xlsxPath)
        return csvPath

    def getFiles(self, reportDate):
        """Convert or reuse the current report and return both raw CSV paths."""
        currentXlsxPath = self.getXlsxFile(reportDate)
        if os.path.isfile(currentXlsxPath):
            currentCsvPath = self.xlsxToCsv(currentXlsxPath)
        else:
            currentCsvPath = self.getCurrentCsvFile(reportDate)

        return {
            "VirtuosicFund": currentCsvPath,
            "VirtuosicFundPrev": self.getPreviousCsvFile(reportDate),
        }


class CheckDiff:
    def __init__(self, fileDict):
        """
        Make the acctual diff check for the virtuosic
        """
        self.fileDict = fileDict
        self.cleanDfs = self.cleanData()
        self.virtuosic_df = self.cleanDfs["VirtuosicFund"]
        self.virtuosic_prev_df = self.cleanDfs["VirtuosicFundPrev"]

    def cleanData(self):
        """Skip the first nine physical rows in every CSV in ``fileDict``."""
        cleanDfs = {}

        for fileName, filePath in self.fileDict.items():
            report_df = pd.read_csv(filePath, skiprows=9, thousands=",")
            cleanDfs[fileName] = report_df.reset_index(drop=True)

        return cleanDfs

    def diffCheck(self):
        current_df = self.cleanDfs["VirtuosicFund"]
        previous_df = self.cleanDfs["VirtuosicFundPrev"]

        currentDate = os.path.splitext(
            os.path.basename(self.fileDict["VirtuosicFund"])
        )[0].rsplit("_", 1)[-1]
        previousDate = os.path.splitext(
            os.path.basename(self.fileDict["VirtuosicFundPrev"])
        )[0].rsplit("_", 1)[-1]

        self.diff_df = pd.DataFrame(
            {
                "Capital": [
                    _report_int(previous_df.iloc[0]["Capital"]),
                    _report_int(current_df.iloc[0]["Capital"]),
                ],
                "MtD": [
                    _report_int(previous_df.iloc[0]["M TrdResult"]),
                    _report_int(current_df.iloc[0]["M TrdResult"]),
                ],
                "tD": [
                    _report_int(previous_df.iloc[0]["D TrdResult"]),
                    _report_int(current_df.iloc[0]["D TrdResult"]),
                ],
            },
            index=[previousDate, currentDate],
        )
        self.diff_df["MtD Check"] = pd.array(
            [
                pd.NA,
                self.diff_df.loc[previousDate, "MtD"]
                + self.diff_df.loc[currentDate, "tD"],
            ],
            dtype="Int64",
        )
        self.diff_df["Diff MtD"] = pd.array(
            [
                pd.NA,
                self.diff_df.loc[currentDate, "MtD Check"]
                - self.diff_df.loc[currentDate, "MtD"],
            ],
            dtype="Int64",
        )
        self.diff_df["CapitalChangeIn"] = pd.array(
            [
                pd.NA,
                self.diff_df.loc[currentDate, "Capital"]
                - self.diff_df.loc[previousDate, "Capital"],
            ],
            dtype="Int64",
        )

        differences = self.diff_df.loc[currentDate, ["Diff MtD", "CapitalChangeIn"]]
        differences = differences[differences != 0]

        if differences.empty:
            print("No diff found")
            return False

        print("Diff found")
        diffReportFolder = os.path.dirname(self.fileDict["VirtuosicFund"])
        diffReportPath = os.path.join(
            diffReportFolder,
            f"virtuosic_diff_{currentDate}.csv",
        )
        self.diff_df.to_csv(diffReportPath)
        print(f"Diff report saved to {diffReportPath}")

        for columnName, diffAmount in differences.items():
            print(f"{columnName}: {int(diffAmount)}")

        return True


def main():
    parser = argparse.ArgumentParser(description="Check the diff for virtousic")
    parser.add_argument(
        "--date",
        type=str,
        help="Date in YYYYMMDD format",
        required=False,
        default="20261004",
    )
    parser.add_argument(
        "--data-folder",
        type=str,
        help="Reporting Data root containing Sage Cloud and DailyReports",
        default=os.getenv("Reporting-Data"),
    )

    args = parser.parse_args()

    report_date = args.date
    virtuosic_files = GetVirtuosic(args.data_folder)

    if virtuosic_files.checkComplete(report_date):
        VirtuosicCheck = CheckDiff(fileDict=virtuosic_files.getFiles(report_date))
        return VirtuosicCheck.diffCheck()

    return False


def run():
    raise SystemExit(1 if main() else 0)


if __name__ == "__main__":
    run()

import pandas as pd
import os
import argparse
from datetime import datetime, timedelta

# If we add tD (from reportingDate) + MtD (from previous day) I get MtD Check then I substract MtD cehck - MtD (from reporting date) = I get the Diff Mtd should be 0
# Also check capital diff, Capital from reporting date - Capital from prev date = If is 0 is Ok


class GetVirtuosic:
    def __init__(self, dataFolder):
        """
        Get the files that we need to perform the diff check
        """
        self.dataFolder = dataFolder
        self.reportFile = None
        self.previousReportFile = None
        self.virtuosicReportPattern = "DailyReportVirtuosicFund_%s.xlsx"

    def getXlsxFiles(self, reportDate):
        """Build the Excel file paths for the report date and previous day."""
        previousDate = (
            datetime.strptime(reportDate, "%Y%m%d") - timedelta(days=1)
        ).strftime("%Y%m%d")

        return {
            "VirtuosicFund": os.path.join(
                self.dataFolder, self.virtuosicReportPattern % reportDate
            ),
            "VirtuosicFundPrev": os.path.join(
                self.dataFolder, self.virtuosicReportPattern % previousDate
            ),
        }

    def checkComplete(self, reportDate):
        """Check for the Excel reports for the date and preceding calendar day."""
        isComplete = True
        for full_path in self.getXlsxFiles(reportDate).values():
            if not os.path.isfile(full_path):
                print(f"{os.path.basename(full_path)} is not ready")
                isComplete = False
                break
        return isComplete

    def getFiles(self, reportDate):
        """
        Convert the Excel reports to CSV and return the CSV file paths.
        """
        fileDict = {}

        for fileName, xlsxPath in self.getXlsxFiles(reportDate).items():
            csvPath = os.path.splitext(xlsxPath)[0] + ".csv"
            report_df = pd.read_excel(xlsxPath)
            report_df.to_csv(csvPath, index=False)
            fileDict[fileName] = csvPath

        return fileDict


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
            report_df = pd.read_csv(filePath, skiprows=9)
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
                    int(previous_df.iloc[0]["Capital"]),
                    int(current_df.iloc[0]["Capital"]),
                ],
                "MtD": [
                    int(previous_df.iloc[0]["M TrdResult"]),
                    int(current_df.iloc[0]["M TrdResult"]),
                ],
                "tD": [
                    int(previous_df.iloc[0]["D TrdResult"]),
                    int(current_df.iloc[0]["D TrdResult"]),
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
        default="20260922",
    )
    parser.add_argument(
        "--data-folder",
        type=str,
        help="Folder containing the Virtuosic daily report Excel files",
        default="C:/Users/Roman Lupan/Desktop/crypto_check/",
    )

    args = parser.parse_args()

    report_date = args.date
    virtuosic_files = GetVirtuosic(args.data_folder)

    if virtuosic_files.checkComplete(report_date):
        VirtuosicCheck = CheckDiff(fileDict=virtuosic_files.getFiles(report_date))
        return VirtuosicCheck.diffCheck()

    return False


if __name__ == "__main__":
    main()

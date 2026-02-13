import pandas as pd


def main():
    file_name = "Account Statement.xlsx"
    df = pd.read_excel(
        file_name,
        header=None,
        usecols="A:D",
        names=["Date", "Description", "Amount(QAR)", "Balance(QAR)"]
    )
    df = df.iloc[13:].reset_index(drop=True)
    df = df[df["Date"].astype(str).str.strip() != "TXN Date"].reset_index(drop=True)

    # Parse dates strictly as dd/mm/yyyy and report any mismatches.
    raw_dates = df["Date"]
    parsed_dates = pd.to_datetime(
        raw_dates,
        format="%d/%m/%Y",
        errors="coerce"
    )
    mismatches = parsed_dates.isna()
    if mismatches.any():
        print("Date mismatches found (expected dd/mm/yyyy):")
        print(df.loc[mismatches])
        raise ValueError("Unparseable date(s) found; expected dd/mm/yyyy.")

    df["Date"] = parsed_dates

    # Change the datetime format to mm/dd/yyyy for CSV output.
    df["Date"] = df["Date"].dt.strftime("%m/%d/%Y")

    df = df.iloc[::-1]

    df["Amount(QAR)"] = (
        df["Amount(QAR)"]
        .astype(str)
        .str.replace(",", "", regex=False)
        .astype(float)
    )

    df["Balance(QAR)"] = (
        df["Balance(QAR)"]
        .astype(str)
        .str.replace(",", "", regex=False)
        .astype(float)
    )

    df["Type"] = df["Amount(QAR)"].apply(
        lambda x: "Income" if x > 0 else "Expense"
    )

    # Remove sign safely (make positive)
    df["Amount(QAR)"] = df["Amount(QAR)"].abs()

    print(df)

    df.to_csv("Statement.csv", sep=",", encoding="utf-8", index=False, header=True)


if __name__ == "__main__":
    main()

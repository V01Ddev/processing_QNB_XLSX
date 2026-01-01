import pandas as pd


def main():
    file_name = "Account Statement.xlsx"
    df = pd.read_excel(
        file_name,
        header=None,
        skiprows=lambda x: x in range(0, 13),
        usecols="A:D",
        names=["Date", "Description", "Amount(QAR)", "Balance(QAR)"]
    )

    df["Date"] = pd.to_datetime(df["Date"])

    # change the datetime format
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

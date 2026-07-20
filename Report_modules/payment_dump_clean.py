import pandas as pd


def generate_payment_dump(input_file, output_file):

    # CSV Read
    df = pd.read_csv(
        input_file,
        encoding="latin1",
        low_memory=False
    )

    # AnalystName Create
    df["AnalystName"] = (
        df["FirstName"].fillna("").astype(str).str.strip()
        + " "
        + df["LastName"].fillna("").astype(str).str.strip()
    ).str.strip()

    # Date Format
    if "offercreationdate" in df.columns:

        df["offercreationdate"] = pd.to_datetime(
            df["offercreationdate"],
            errors="coerce"
        ).dt.strftime("%d-%m-%Y")

    required_columns = [
        "BankName",
        "BusinessProduct",
        "AccountStatus",
        "AnalystName",
        "EmployeeId",
        "ef10",
        "TradelineFullName",
        "AccountNumber",
        "AmountPaid",
        "PaymentDate",
        "CreationDate",
        "offercreationdate",
        "PaymentStatus",
        "LivePayment",
        "PaymentMode",
        "Resolution",
        "TransactionNo"
    ]

    final_columns = [
        col for col in required_columns
        if col in df.columns
    ]

    if "AnalystName" not in final_columns:
        final_columns.insert(3, "AnalystName")

    output_df = df[final_columns]

    output_df.to_excel(
        output_file,
        index=False
    )

    return output_file
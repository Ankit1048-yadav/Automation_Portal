import pandas as pd
import msoffcrypto
import io
from pyxlsb import open_workbook
from datetime import datetime

SETT_PAID_RECO = 0.00  # CR (keep default or update if needed)

def generate_axis_message(FILE_PATH, PASSWORD):

    decrypted = io.BytesIO()

    with open(FILE_PATH, "rb") as f:
        office = msoffcrypto.OfficeFile(f)
        office.load_key(password=PASSWORD)
        office.decrypt(decrypted)

    decrypted.seek(0)

    with open_workbook(decrypted) as wb:
        sheet = wb.get_sheet(wb.sheets[0])
        rows = [[cell.v for cell in row] for row in sheet.rows()]

    df = pd.DataFrame(rows[1:], columns=rows[0])
    df.columns = df.columns.str.strip()

    # ---------------- CALCULATIONS ----------------

    upgrade_pos = df.loc[
        df["UPGRADE/RECOVERY"].str.upper() == "UPGRADE", "POS"
    ].sum()
    upgrade_cr = round(upgrade_pos / 1e7, 2)

    biu_reco = round(
        df.loc[
            df["UPGRADE/RECOVERY"].str.upper() == "RECOVERY", "Reporting"
        ].sum(),
        2
    )

    # Total
    total = round(upgrade_cr + biu_reco + SETT_PAID_RECO, 2)

    # B4 RSL %
    b4_total = df.loc[df["Bkt"] == "B4(O)", "POS"].sum()

    b4_resolved = df.loc[
        (df["Bkt"] == "B4(O)") &
        (df["RSL STATUS"].str.upper() == "RESOLVED"),
        "POS"
    ].sum()

    b4_rsl_pct = round((b4_resolved / b4_total) * 100, 2) if b4_total else 0

    # ---------------- DATE ----------------
    today = datetime.now().strftime("%d %b %Y")

    # ---------------- MESSAGE ----------------
    message = f"""
Axis Cards OPN/ADD ({today})

Upgrade ₹ {upgrade_cr} CR
BIU Reco ₹ {biu_reco} CR
Sett Paid Reco ₹ {SETT_PAID_RECO} CR
Total (U+R+Sett) ₹ {total} CR
B4 OPN RSL % {b4_rsl_pct}%
""".strip()

    return message
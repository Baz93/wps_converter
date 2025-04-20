import sys
import pandas as pd
from datetime import datetime


def convert(src_path: str, dst_path: str):
    df = pd.read_excel(src_path, header=None)

    edr_df = df[df[0] == 'EDR'].copy()

    def from_excel_datetime(dt: datetime) -> float:
        base = datetime(1899, 12, 30)  # Excel origin (with leap bug)
        delta = dt - base
        return round(delta.days + delta.seconds / 86400, 6)

    edr_pay_start = pd.to_datetime(edr_df[5])
    edr_pay_end = pd.to_datetime(edr_df[6])
    edr_pay_middle = edr_pay_start + (edr_pay_end - edr_pay_start) / 2
    year, = set(edr_pay_middle.dt.year)
    month, = set(edr_pay_middle.dt.month)
    month_abbr = datetime(year, month, 1).strftime("%b").upper()

    edr_final = pd.DataFrame({
        'Type': edr_df[0],
        'Employee_ID': edr_df[1],
        'Routing_Code': edr_df[3],
        'Employee_IBAN': edr_df[4],
        'PayStart_Date': edr_pay_start.dt.strftime("%Y-%m-%d"),
        'PayEnd_Date': edr_pay_end.dt.strftime("%Y-%m-%d"),
        'Days_In_Period': edr_df[7],
        'Fixed_Income': edr_df[8].apply(from_excel_datetime),
        'Variable_Income': edr_df[9].astype(float),
        'Leave_Days': edr_df[10]
    })

    scr_row = df[df[0] == 'SCR'].copy()

    assert((scr_row[5] == month_abbr).all())

    scr_final = pd.DataFrame({
        'Type': 'SCR',
        'Company_ID': scr_row[1],
        'Company_Bank_Code': scr_row[2],
        'Transfer_Date': pd.to_datetime(scr_row[3]).dt.strftime("%Y-%m-%d"),
        'Reference_Number': scr_row[4],
        'Salary_Month': f"{month:02}{year:04}",
        'Employee_Count': scr_row[6],
        'Total_Amount': scr_row[7].astype(float),
        'Currency': scr_row[8]
    })

    edr_csv = edr_final.to_csv(index=False, float_format='%.2f')
    scr_csv = scr_final.to_csv(index=False, header=False, float_format='%.2f')

    combined_csv = edr_csv.strip() + "\n" + scr_csv.strip()

    with open(dst_path, "w", encoding="utf-8") as f:
        f.write(combined_csv)


def main():
    assert len(sys.argv) == 2, "Usage: python convert.py <input_file>"
    src_path = sys.argv[1]
    dst_path = src_path + ".csv"
    convert(src_path, dst_path)


if __name__ == "__main__":
    main()
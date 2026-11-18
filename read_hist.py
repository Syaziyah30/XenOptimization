import pandas as pd

Hist1 = r"D:\XenOptimization\Data\MT Chiller\Hist Value1 - MTCHW.xlsx"


def read_hist(path, sheet_name=0):
    """Read a historian export: DateTime column + one column per tag."""
    df = pd.read_excel(path, sheet_name=sheet_name)
    df["DateTime"] = pd.to_datetime(df["DateTime"])
    df = df.set_index("DateTime").sort_index()

    # Shorten tag names: "C10A1_1_C101C_MTCHW_CH01.CondPressS" -> "CH01.CondPressS"
    df.columns = [c.replace("C10A1_1_C101C_MTCHW_", "") for c in df.columns]
    df = df.apply(pd.to_numeric, errors="coerce")
    return df


if __name__ == "__main__":
    df = read_hist(Hist1)

    print(f"Shape: {df.shape}")
    print(f"Range: {df.index.min()} -> {df.index.max()}")
    print("\nColumns:")
    for c in df.columns:
        print(f"  {c:<22} non-null={df[c].notna().sum()}")
    print("\nHead:")
    print(df.head())

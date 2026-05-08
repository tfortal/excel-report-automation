from pathlib import Path

import pandas as pd


BASE_DIR = Path(__file__).resolve().parent.parent
DATA_PATH = BASE_DIR / "data_sample" / "sample_sales_data.csv"
OUTPUT_DIR = BASE_DIR / "output"
OUTPUT_FILE = OUTPUT_DIR / "monthly_sales_report.xlsx"


def load_data(file_path: Path) -> pd.DataFrame:
    """Load sales data from a CSV file."""
    if not file_path.exists():
        raise FileNotFoundError(f"Input file not found: {file_path}")

    df = pd.read_csv(file_path)
    return df


def clean_data(df: pd.DataFrame) -> pd.DataFrame:
    """Clean and prepare sales data."""
    required_columns = {
        "date",
        "order_id",
        "customer",
        "product",
        "category",
        "sales_channel",
        "quantity",
        "unit_price",
    }

    missing_columns = required_columns - set(df.columns)
    if missing_columns:
        raise ValueError(f"Missing required columns: {missing_columns}")

    df = df.copy()

    df["date"] = pd.to_datetime(df["date"], errors="coerce")
    df["quantity"] = pd.to_numeric(df["quantity"], errors="coerce")
    df["unit_price"] = pd.to_numeric(df["unit_price"], errors="coerce")

    df = df.dropna(subset=["date", "quantity", "unit_price"])

    df["total_amount"] = df["quantity"] * df["unit_price"]
    df["month"] = df["date"].dt.to_period("M").astype(str)

    return df


def calculate_summary(df: pd.DataFrame) -> pd.DataFrame:
    """Calculate high-level business KPIs."""
    total_revenue = df["total_amount"].sum()
    total_orders = df["order_id"].nunique()
    total_units = df["quantity"].sum()
    average_ticket = total_revenue / total_orders if total_orders else 0

    summary = pd.DataFrame(
        {
            "Metric": [
                "Total Revenue",
                "Total Orders",
                "Total Units Sold",
                "Average Ticket",
            ],
            "Value": [
                total_revenue,
                total_orders,
                total_units,
                average_ticket,
            ],
        }
    )

    return summary


def calculate_category_report(df: pd.DataFrame) -> pd.DataFrame:
    """Calculate revenue and quantity by product category."""
    category_report = (
        df.groupby("category", as_index=False)
        .agg(
            total_revenue=("total_amount", "sum"),
            total_quantity=("quantity", "sum"),
            total_orders=("order_id", "nunique"),
        )
        .sort_values(by="total_revenue", ascending=False)
    )

    return category_report


def calculate_channel_report(df: pd.DataFrame) -> pd.DataFrame:
    """Calculate revenue and orders by sales channel."""
    channel_report = (
        df.groupby("sales_channel", as_index=False)
        .agg(
            total_revenue=("total_amount", "sum"),
            total_orders=("order_id", "nunique"),
        )
        .sort_values(by="total_revenue", ascending=False)
    )

    return channel_report


def calculate_monthly_report(df: pd.DataFrame) -> pd.DataFrame:
    """Calculate revenue and orders by month."""
    monthly_report = (
        df.groupby("month", as_index=False)
        .agg(
            total_revenue=("total_amount", "sum"),
            total_orders=("order_id", "nunique"),
            total_quantity=("quantity", "sum"),
        )
        .sort_values(by="month")
    )

    return monthly_report


def export_report(
    summary: pd.DataFrame,
    category_report: pd.DataFrame,
    channel_report: pd.DataFrame,
    monthly_report: pd.DataFrame,
    cleaned_data: pd.DataFrame,
    output_file: Path,
) -> None:
    """Export all reports to an Excel workbook."""
    OUTPUT_DIR.mkdir(exist_ok=True)

    with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
        summary.to_excel(writer, sheet_name="Summary", index=False)
        category_report.to_excel(writer, sheet_name="By Category", index=False)
        channel_report.to_excel(writer, sheet_name="By Channel", index=False)
        monthly_report.to_excel(writer, sheet_name="Monthly Report", index=False)
        cleaned_data.to_excel(writer, sheet_name="Cleaned Data", index=False)


def main() -> None:
    """Run the complete report automation workflow."""
    print("Loading data...")
    raw_data = load_data(DATA_PATH)

    print("Cleaning data...")
    cleaned_data = clean_data(raw_data)

    print("Calculating KPIs...")
    summary = calculate_summary(cleaned_data)
    category_report = calculate_category_report(cleaned_data)
    channel_report = calculate_channel_report(cleaned_data)
    monthly_report = calculate_monthly_report(cleaned_data)

    print("Exporting report...")
    export_report(
        summary=summary,
        category_report=category_report,
        channel_report=channel_report,
        monthly_report=monthly_report,
        cleaned_data=cleaned_data,
        output_file=OUTPUT_FILE,
    )

    print(f"Report generated successfully: {OUTPUT_FILE}")


if __name__ == "__main__":
    main()

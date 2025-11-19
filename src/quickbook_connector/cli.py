# cli.py
import argparse
from quickbook_connector.customer_excel_qb_sync import process_customers, read_customers_from_excel, get_qb_customers
from quickbook_connector.reporting import write_json_report

def main():
    parser = argparse.ArgumentParser(description="QuickBooks Excel Sync CLI")
    parser.add_argument("excel_file", help="Path to the Excel file")
    parser.add_argument("--report", default="customer_report.json", help="Output JSON file name")
    args = parser.parse_args()

    comparison = process_customers(args.excel_file)
    excel_count = len(read_customers_from_excel(args.excel_file))
    qb_count = len(get_qb_customers())

    write_json_report(comparison, excel_count, qb_count, args.report)

if __name__ == "__main__":
    main()

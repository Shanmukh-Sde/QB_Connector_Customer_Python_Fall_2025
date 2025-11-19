# report_writer.py
import json
from quickbook_connector.customer_excel_qb_sync import CustomerComparison, Customer

def write_json_report(comparison: CustomerComparison, excel_count: int, qb_count: int, output_file="customer_report.json"):
    report_data = {
        "summary": {
            "excel_total": excel_count,
            "qb_total": qb_count,
            "matching_count": comparison.matching_count,
            "conflict_count": len(comparison.same_id_diff_data),
            "only_in_excel_count": len(comparison.only_in_excel),
            "only_in_qb_count": len(comparison.only_in_qb),
        },
        "conflicts": [
            {
                "customer_id": cid,
                "excel_name": excel_name,
                "qb_name": qb_name
            }
            for excel_name, qb_name, cid in comparison.same_id_diff_data
        ],
        "only_in_excel": [
            {"customer_id": c.customer_id, "name": c.name}
            for c in comparison.only_in_excel
        ],
        "only_in_qb": [
            {"customer_id": c.customer_id, "name": c.name}
            for c in comparison.only_in_qb
        ]
    }

    with open(output_file, "w") as f:
        json.dump(report_data, f, indent=4)

    print(f"\nJSON report created: {output_file}")

import openpyxl
import json

def parse_nested_excel_by_position(
    file_path: str,
    sheet_name: str = "Sheet1",
    levels: int = 1
):
    wb = openpyxl.load_workbook(file_path)
    ws = wb[sheet_name]

    extra_field_names = ["description", "tag", "conformityTopic", "status", "thresholdValue", "performanceLevel", "category"]
    current_levels = [None] * levels
    tree = []

    for row in ws.iter_rows(max_col=levels + 1 + len(extra_field_names)):
        # Parse extra fields
        extra_data = {}
        for i, field_name in enumerate(extra_field_names):
            col_idx = levels + 1 + i  # starts after the last name column
            value = row[col_idx].value if col_idx < len(row) else None
            if field_name == 'tag':
                value = [item.strip() for item in value.split(",\n")]

            if value:
                extra_data[field_name] = value

        for level in range(levels):
            id_cell = row[level].value                     # ID at index `level`
            name_cell = row[level + 1].value               # Name at index `level + 1`

            if id_cell:
                if str(id_cell).strip() == 'ID': continue   # Don't read headers
                node = {
                    "type": ["Criterion"],
                    "id": str(id_cell).strip(),
                    "name": str(name_cell).strip() if name_cell else ""
                }
                node.update(extra_data)
                node["subCriterion"] = []

                current_levels[level] = node
                for j in range(level + 1, levels):
                    current_levels[j] = None  # Clear deeper levels

                if level == 0:
                    tree.append(node)
                else:
                    parent = current_levels[level - 1]
                    if parent:
                        parent["subCriterion"].append(node)

                break  # Only one level per row is processed

    return tree


# Example usage
file_path = "./files/test_xls_to_csv_file_2.xlsx"
json_result = parse_nested_excel_by_position(
    file_path=file_path,
    sheet_name="Sheet1",
    levels=3,
)

# print(json.dumps(json_result, indent=2, ensure_ascii=False))

# Export JSON to a file
output_json_path = "./files/output/output_result.json"
with open(output_json_path, "w", encoding="utf-8") as f:
    json.dump(json_result, f, indent=2, ensure_ascii=False)

print(f"JSON result exported to {output_json_path}")
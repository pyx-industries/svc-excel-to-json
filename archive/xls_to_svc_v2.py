import openpyxl
import json
import os

def parse_excel_to_svc_json(
    file_path: str,
    sheet_name: str = "Sheet1",
    levels: int = 1,
    extra_field_names: list = None
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

def main():
    print("=== Excel to Nested JSON Converter ===")
    input_path = input("Enter input Excel file path (.xlsx): ").strip()
    while not os.path.isfile(input_path):
        input_path = input("File not found. Please enter a valid file path: ").strip()

    try:
        levels = int(input("Enter number of levels (e.g. 3): ").strip())
    except ValueError:
        print("Invalid number. Defaulting to 3.")
        levels = 1

    output_path = input("Enter output JSON file path: ").strip()
    if not output_path.lower().endswith(".json"):
        output_path += ".json"

    try:
        json_result = parse_excel_to_svc_json(
            file_path=input_path,
            levels=levels
        )

        with open(output_path, "w", encoding="utf-8") as f:
            json.dump(json_result, f, indent=2, ensure_ascii=False)

        print(f"\n✅ JSON exported to: {output_path}")
    except Exception as e:
        print(f"\n❌ Error: {e}")

if __name__ == "__main__":
    main()

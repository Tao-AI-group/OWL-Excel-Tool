import pandas as pd
from owlready2 import get_ontology
import argparse

def extract_properties_to_excel(owl_path, output_path):
    onto = get_ontology(owl_path).load()

    records = []

    # 1. ObjectProperty
    for prop in onto.object_properties():
        label = prop.label.first() if prop.label else prop.name
        domains = [cls.label.first() if cls.label else cls.name for cls in prop.domain]
        ranges = [cls.label.first() if cls.label else cls.name for cls in prop.range]
        records.append({
            "name": label,
            "property": "object property",
            "domain": "; ".join(domains) if domains else "",
            "range": "; ".join(ranges) if ranges else ""
        })

    # 2. DataProperty
    for prop in onto.data_properties():
        label = prop.label.first() if prop.label else prop.name
        domains = [cls.label.first() if cls.label else cls.name for cls in prop.domain]
        ranges = [r.name if hasattr(r, "name") else str(r) for r in prop.range]
        records.append({
            "name": label,
            "property": "data property",
            "domain": "; ".join(domains) if domains else "",
            "range": "; ".join(ranges) if ranges else ""
        })

    # 3. Save to Excel
    df = pd.DataFrame(records)
    df.to_excel(output_path, index=False)
    print(f"save: {output_path}")


def excel_to_txt(excel_path, txt_path):
    df = pd.read_excel(excel_path)
    df.to_csv(txt_path, sep="\t", index=False)
    print(f" transform {excel_path} into: {txt_path}")

    

if __name__ == "__main__":

    parser = argparse.ArgumentParser(description="Extract OWL Object/Data properties to Excel and optional TXT")
    parser.add_argument("-i", "--input", required=True, help="Path to input OWL file")
    parser.add_argument("-o", "--output", required=True, help="Path to output Excel file")
    parser.add_argument("--txt", help="Optional: Path to output TXT file (UTF-8 tab-delimited)")

    args = parser.parse_args()

    df = extract_properties_to_excel(args.input, args.output)

    if args.txt:
        save_excel_to_txt(df, args.txt)



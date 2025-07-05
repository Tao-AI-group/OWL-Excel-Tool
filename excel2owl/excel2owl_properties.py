import pandas as pd
import types
import re
from tqdm import tqdm
from owlready2 import get_ontology, ObjectProperty, DataProperty, Thing
import argparse

def clean_excel_to_utf8_txt(input_file: str, output_file: str):
    import pandas as pd

    df = pd.read_excel(input_file)
    df_cleaned = df.replace({r'[\r\n]+': ' '}, regex=True)

    # Save as UTF-8 tab-separated file
    df_cleaned.to_csv(output_file, sep="\t", index=False, encoding="utf-8")
    print(f"Saved cleaned UTF-8 TXT to: {output_file}")

def get_next_property_index(onto, prefix):
    max_index = 0
    all_props = list(onto.object_properties()) + list(onto.data_properties())
    for prop in all_props:
        match = re.match(rf"{prefix}(\d+)", prop.name)
        if match:
            idx = int(match.group(1))
            max_index = max(max_index, idx)
    return max_index + 1

def property_exists_by_label(prop_list, label_text):
    for prop in prop_list:
        if label_text in prop.label:
            return prop
    return None

def get_or_create_property_by_label(onto, label_text, prop_type, next_index):
    prop_list = onto.object_properties() if prop_type == "object" else onto.data_properties()
    existing = property_exists_by_label(prop_list, label_text)
    if existing:
        print(f"exist {prop_type} property label='{label_text}' -> IRI: {existing.iri}")
        return existing, next_index

    prefix = "R" if prop_type == "object" else "D"
    iri_name = f"{prefix}{next_index:03d}"
    

    base_class = ObjectProperty if prop_type == "object" else DataProperty
    prop = types.new_class(iri_name, (base_class,))
    prop.label = [label_text]
    
    print(f"creat new {prop_type} property: {prop.iri}")
    return prop, next_index + 1

def get_class_by_label(onto, label):
    results = list(onto.search(label=label))
    if results:
        entity = results[0]
        if isinstance(entity, type) and issubclass(entity, Thing):
            print(f"existed label='{label}' -> class: {entity.name}")
            return entity
        else:
            raise TypeError(f"label='{label}' this is not class but: {type(entity)}")
    

def add_properties_from_txt(owl_path, txt_path, output_path):
    df = pd.read_csv(txt_path, sep="\t", dtype=str)
    df.columns = df.columns.str.strip().str.lower().str.replace(" ", "_")

    onto = get_ontology(owl_path).load()
    onto.base_iri = "https://github.com/Tao-AI-group/BSO_AD#"
        
    with onto:
        next_obj_index = get_next_property_index(onto, "R")
        next_data_index = get_next_property_index(onto, "D")
        print(f"object property starting index R{next_obj_index:03d}, data property starting index D{next_data_index:03d}")

        for _, row in tqdm(df.iterrows(), total=len(df), desc="Adding properties"):
            prop_label = row['name'].strip()
            prop_type = row['property'].strip().lower()
            domain_label = row['domain'].strip()
            range_label = row['range'].strip()

            DomainClass = get_class_by_label(onto, domain_label)

            if prop_type == "object property":
                RangeClass = get_class_by_label(onto, range_label)
                prop, next_obj_index = get_or_create_property_by_label(onto, prop_label, "object", next_obj_index)
                prop.domain = [DomainClass]
                prop.range = [RangeClass]
            elif prop_type == "data property":
                prop, next_data_index = get_or_create_property_by_label(onto, prop_label, "data", next_data_index)
                prop.domain = [DomainClass]
                
                
            else:
                print(f"unknow type: {prop_type} skip")

    onto.save(file=output_path, format="rdfxml")
    print(f"\n ontology is saved into : {output_path}")

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Convert Excel relations to OWL via TXT, and save to new OWL file.")
    parser.add_argument("-e", "--excel", required=True, help="Path to input Excel file with relations")
    parser.add_argument("-i", "--input", required=True, help="Path to input OWL file")
    parser.add_argument("-o", "--output", required=True, help="Path to output OWL file")
    parser.add_argument("-t", "--txt", help="Optional: path to intermediate TXT file (default will be based on Excel name)")

    args = parser.parse_args()

    # 1. Step: Excel -> UTF-8 TXT
    if args.txt:
        txt_path = args.txt
    else:
        txt_path = args.excel.replace(".xlsx", "_utf8.txt")

    clean_excel_to_utf8_txt(args.excel, txt_path)

    # 2. Step: Add properties to OWL
    add_properties_from_txt(args.input, txt_path, args.output)


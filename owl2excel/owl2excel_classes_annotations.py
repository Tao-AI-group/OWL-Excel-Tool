from owlready2 import get_ontology, Thing, rdfs, AnnotationProperty
import csv
import math
import pandas as pd
import argparse

# obtain label or name
def get_label_or_name(cls):
    return cls.label.first() if cls.label else cls.name

#extract annotation
def get_annotation_summary(cls, props):
    values = []
    for prop in props:
        if isinstance(prop, str):
            # support skos__definition, UMLS_cui etc
            val_list = getattr(cls, prop, [])
        else:
            val_list = getattr(cls, prop.python_name, [])
        
        values.append(val_list[0] if val_list else "")
    return values  # to ensure every annotation property in one column and align with the header order 

# traverse 
def traverse_class(cls, path, all_paths, props, depth=1):
    subclasses = list(cls.subclasses())
    label = get_label_or_name(cls)
    annotation_values = get_annotation_summary(cls, props)
    new_path = path + [label]+ annotation_values

    if not subclasses:
        all_paths.append((depth, new_path))
    else:
        for sub in subclasses:
            traverse_class(sub, new_path, all_paths, props, depth+1)
    
def extract_class_hierarchy_with_annotations(owl_file, output_txt):
    ontology = get_ontology(owl_file).load()
    skos = get_ontology("http://www.w3.org/2004/02/skos/core").load()

    annotation_property_dict = {}

    for prop in ontology.annotation_properties():
        annotation_property_dict[prop.name] = prop
    for prop in skos.annotation_properties():
        annotation_property_dict[prop.name] = prop
    
    for cls in ontology.classes():
        val = annotation_property_dict["definition"][cls]
        print(cls.name, "skos:definition:", val[0] if val else "")
    
    print(annotation_property_dict)
    annotation_props = [
        rdfs.comment,
        annotation_property_dict["definition"],
        annotation_property_dict["altLabel"],
        annotation_property_dict["ICD10CM"],
        annotation_property_dict["UMLS_CUI"],
        annotation_property_dict["UMLS_Semantic_Types"]
    ] # can add more annotation properties 

    all_paths = []

    top_level_classes = [
        cls for cls in ontology.classes()
        if Thing in cls.is_a or cls.is_a == []
    ]

    for top_cls in top_level_classes:
        traverse_class(top_cls, [], all_paths, annotation_props)

    max_depth = max(depth for depth, _ in all_paths)

    num_annos = len(annotation_props)
    header = []
    for i in range(max_depth):
        header.append(f"Level{i+1}_Label")
        for j in range(num_annos):
            header.append(f"Level{i+1}_Anno{j+1}")

    with open(output_txt, "w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f, delimiter="\t")
        writer.writerow(header)
        for _, path in all_paths:
            row = []
            unit_size = 1 + num_annos
            total_units = len(path) // unit_size

            for level in range(max_depth):
                if level < total_units:
                    start = level * unit_size
                    row.extend(path[start : start + unit_size])  # label + annotation1 + annotation2 + ...
                else:
                    row.extend([""] * unit_size)

            writer.writerow(row)
        
    print(f"Saved to: {output_txt}")


def txt2excel(txt_file, excel_file):
    df = pd.read_csv(txt_file, sep="\t", index_col=None, dtype=str)

    df.to_excel(excel_file, index=False)

    print("completed")

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Extract OWL classes and annotation properties to Excel")
    parser.add_argument("-i", "--input", required=True, help="Path to OWL file")
    parser.add_argument("-o", "--output", required=True, help="Output Excel file name")
    args = parser.parse_args()

    output_txt = args.output.replace(".xlsx", ".txt")
    extract_class_hierarchy_with_annotations(args.input, output_txt)
    txt2excel(output_txt, args.output)

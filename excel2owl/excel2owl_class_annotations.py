import pandas as pd
from owlready2 import *
from rdflib.namespace import RDFS, SKOS
import argparse

def clean_excel_to_utf8_txt(input_file: str, output_file: str):
    import pandas as pd

    df = pd.read_excel(input_file)
    df_cleaned = df.replace({r'[\r\n]+': ' '}, regex=True)

    # Save as UTF-8 tab-separated file
    df_cleaned.to_csv(output_file, sep="\t", index=False, encoding="utf-8")
    print(f"Saved cleaned UTF-8 TXT to: {output_file}")

def formalize_label(label):
    exceptions = {"a", "an", "the", "and", "but", "or", "for", "nor", "on", "at", "to", "from", "by", "of", "in", "with"}
    words = label.strip().replace("_", " ").replace("-", " ").split(" ")

    result = []
    for i, word in enumerate(words):
        word_lower = word.lower()
        if i == 0 or word_lower not in exceptions:
            result.append(word.capitalize())
        else:
            result.append(word_lower)

    return "_".join(result)

def build_ontology_with_standard_annotations(txt_path: str, ontology_uri: str, output_path: str):
    # Step 1: Load the tab-separated .txt file
    df = pd.read_csv(txt_path, sep="\t", dtype=str).fillna("")

    # Step 2: Create the main ontology
    onto = get_ontology(ontology_uri)

    # Load SKOS ontology outside the with block
    skos = get_ontology("http://www.w3.org/2004/02/skos/core").load()

    # === ID generation setup ===
    id_counter = 0
    label_to_id = {}  # Maps label text to 000001-style IDs

    with onto:
        # Step 3: Define skos:definition annotation property in SKOS namespace
        class definition(AnnotationProperty):
            namespace = skos
        
        class ICD10CM(AnnotationProperty):
            namespace = onto
        
        class UMLS_CUI(AnnotationProperty):
            namespace = onto
        
        class UMLS_Semantic_Types(AnnotationProperty):
            namespace = onto

        created_classes = {}     # key: class_id -> value: owlready2 class object
        class_parents_map = {}   # key: class_id -> set of parent classes

        # Step 4: Iterate through each row
        for _, row in df.iterrows():
            hierarchy = []

            # Build hierarchy: list of (class_id, label, level_num)
            for level in range(1, 9):
                label = row.get(f"level_{level}", "").strip()
                label = formalize_label(label)
                if label:
                    if label not in label_to_id:
                        label_to_id[label] = f"{id_counter:05d}"
                        id_counter += 1
                    iri_suffix = label_to_id[label]
                    hierarchy.append((iri_suffix, label, level))
          
            print(f"hierarchy:", hierarchy)

            # Step 5: Build or update class hierarchy with multiple parent support
            parent = Thing
            for class_id, label, level_num in hierarchy:
                if class_id not in created_classes:
                    # First time seen: create class
                    new_class = types.new_class(class_id, (parent,))
                    created_classes[class_id] = new_class
                    class_parents_map[class_id] = {parent}

                    # Annotations
                    # rdfs:label
                    new_class.label.append(label)

                    # rdfs:comment
                    comment_text = row.get(f"level_{level_num} comment", "")
                    new_class.comment.extend(
                        item.strip().replace("\\n", "\n") for item in comment_text.split('|') if item.strip()
                    )

                    # skos:definition
                    definition_text = row.get(f"level_{level_num} definition", "")
                    new_class.definition.extend(
                        item.strip().replace("\\n", "\n") for item in definition_text.split('|') if item.strip()
                    )

                    # skos:altLabel
                    synonym_text = row.get(f"level_{level_num} synonym", "")
                    new_class.altLabel.extend(
                        item.strip().replace("\\n", "\n") for item in synonym_text.split('|') if item.strip()
                    )

                    # ICD10CM 
                    ICD10CM_text = row.get(f"level_{level_num} ICD10CM", "")
                    if ICD10CM_text.strip():
                        new_class.ICD10CM.append(ICD10CM_text.strip()) 
                    
                    # UMLS_CUI
                    UMLS_CUI_text = row.get(f"level_{level_num} UMLS_CUI", "")
                    if UMLS_CUI_text.strip():
                        new_class.UMLS_CUI.append(UMLS_CUI_text.strip()) 

                    #UMLS_Semantic_Types
                    UMLS_Semantic_Types_text = row.get(f"level_{level_num} UMLS_Semantic_Types", "")
                    if UMLS_Semantic_Types_text.strip():
                        new_class.UMLS_Semantic_Types.append(UMLS_Semantic_Types_text.strip()) 


                else:
                    # Already exists, check if new parent needs to be added
                    existing_class = created_classes[class_id]
                    if parent not in class_parents_map[class_id]:
                        # Add new parent (multi-inheritance)
                        new_parents = set(existing_class.is_a) | {parent}
                        existing_class.is_a = list(new_parents)
                        class_parents_map[class_id].add(parent)

                # Move down to next level
                parent = created_classes[class_id]
                
                
    for cls in onto.classes():
        cls.is_a = [parent for parent in cls.is_a if parent != owl.Thing]

    # Step 6: Save the ontology
    onto.save(file=output_path, format="rdfxml")
    print(f"Ontology saved to: {output_path}")



if __name__ == "__main__":
    
    parser = argparse.ArgumentParser(description="Convert Excel to OWL ontology via UTF-8 Txt.")
    parser.add_argument("-e", "--excel", required=True, help="Input Excel file with ontology data")
    parser.add_argument("-u", "--uri", required=True, help="Ontology base URI (e.g., https://yourdomain.org/ontology#)")
    parser.add_argument("-o", "--output", required=True, help="Path to output OWL file")
    parser.add_argument("-t", "--txt", help="Optional: output intermediate TXT file path")

    args = parser.parse_args()

    # Step 1: Excel -> TXT
    if args.txt:
        txt_path = args.txt
    else:
        txt_path = args.excel.replace(".xlsx", "_utf8.txt")

    clean_excel_to_utf8_txt(args.excel, txt_path)

    # Step 2: TXT -> OWL
    build_ontology_with_standard_annotations(txt_path, args.uri, args.output)

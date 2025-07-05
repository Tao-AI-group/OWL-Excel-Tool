from owlready2 import *
from collections import deque
import re
from owlready2 import AnnotationProperty, rdfs
import json
import argparse

def get_all_subclasses(cls):
    subclasses = set(cls.subclasses())
    for sub in cls.subclasses():
        subclasses.update(get_all_subclasses(sub))
    return subclasses

def topological_sort(classes):
    graph = {cls: set(cls.is_a) for cls in classes}
    in_degree = {cls: 0 for cls in classes}

    for cls, parents in graph.items():
        for parent in parents:
            if parent in in_degree:
                in_degree[cls] += 1

    queue = deque([cls for cls, deg in in_degree.items() if deg == 0])
    ordered = []

    while queue:
        cls = queue.popleft()
        ordered.append(cls)
        for child in classes:
            if cls in graph.get(child, set()):
                in_degree[child] -= 1
                if in_degree[child] == 0:
                    queue.append(child)
    return ordered

def get_existing_max_id(onto, base_iri):
    max_id = -1
    for cls in onto.classes():
        if cls.iri.startswith(base_iri):
            suffix = cls.iri.split("#")[-1]
            if suffix.isdigit():
                max_id = max(max_id, int(suffix))
    return max_id

def get_existing_max_objprop_id(onto, base_iri):
    max_id = -1
    for prop in onto.object_properties():
        if prop.iri.startswith(base_iri):
            suffix = prop.iri.split("#")[-1]
            if suffix.startswith("R") and suffix[1:].isdigit():
                max_id = max(max_id, int(suffix[1:]))
    return max_id

def get_existing_max_dataprop_id(onto, base_iri):
    max_id = -1
    for prop in onto.data_properties():
        if prop.iri.startswith(base_iri):
            suffix = prop.iri.split("#")[-1]
            if suffix.startswith("D") and suffix[1:].isdigit():
                max_id = max(max_id, int(suffix[1:]))
    return max_id


def format_label(label):
    # Convert label to Title_Case_With_Underscores
    exceptions = {"a", "an", "the", "and", "but", "or", "for", "nor", "on", "at", "to", "from", "by", "of", "in", "with"}
    words = label.strip().split("_")

    result = []
    for i, word in enumerate(words):
        word_lower = word.lower()
        if i == 0 or word_lower not in exceptions:
            result.append(word.capitalize())
        else:
            result.append(word_lower)

    return "_".join(result)

def get_used_annotation_properties(source_cls):
    onto = source_cls.namespace.ontology
    used_props = set()
    for prop in onto.annotation_properties():
        if getattr(source_cls, prop.name, None):
            used_props.add(prop)
    used_props.update({rdfs.label, rdfs.comment})
    return used_props

def get_all_annotation_properties_to_copy(source_cls):
    onto = source_cls.namespace.ontology
    defined_props = set(onto.annotation_properties())
    used_props = get_used_annotation_properties(source_cls)

    return defined_props | used_props  

def copy_annotation_properties(source_cls, target_cls):
    for ann in get_all_annotation_properties_to_copy(source_cls):
        ann_name = ann.python_name
        source_values = getattr(source_cls, ann_name, [])
        if source_values:
            target_values = getattr(target_cls, ann_name, [])
            for val in source_values:
                if val not in target_values:
                    target_values.append(val)

def preserve_valid_object_properties(onto_base, onto_import, cls_map, base_iri, iri_objprop_counter):
    with onto_base:
        for prop in onto_import.object_properties():
            domains = list(prop.domain)
            ranges = list(prop.range)
            domain_ok = all((d in cls_map or d in onto_base.classes()) for d in domains)
            range_ok = all((r in cls_map or r in onto_base.classes()) for r in ranges)
            if domain_ok and range_ok:
                new_prop = types.new_class(prop.name, (ObjectProperty,))
                new_prop.namespace = onto_base
                iri_suffix = f"R{iri_objprop_counter:05d}"
                new_prop.iri = base_iri + iri_suffix
                iri_objprop_counter += 1
                new_prop.domain = [cls_map.get(d, d) for d in domains]
                new_prop.range = [cls_map.get(r, r) for r in ranges]
                new_prop.label = [prop.name]
                if hasattr(prop, 'comment') and prop.comment:
                    new_prop.comment = list(prop.comment)
                new_prop.comment.append(f"Original IRI: {prop.iri}")
                print(f"Created and preserved: {new_prop.name} ({new_prop.iri})")
    return iri_objprop_counter

def preserve_valid_data_properties(onto_base, onto_import, cls_map, base_iri, iri_dataprop_counter):
    with onto_base:
        for prop in onto_import.data_properties():
            domains = list(prop.domain)
            if all((d in cls_map or d in onto_base.classes()) for d in domains):
                new_prop = types.new_class(prop.name, (DataProperty,))
                new_prop.namespace = onto_base
                iri_suffix = f"D{iri_dataprop_counter:05d}"
                new_prop.iri = base_iri + iri_suffix
                iri_dataprop_counter += 1
                new_prop.domain = [cls_map.get(d, d) for d in domains]
                print('data property range')
                print(prop.range)
                new_prop.range = list(prop.range)
                new_prop.label = [prop.name]
                if hasattr(prop, 'comment') and prop.comment:
                    new_prop.comment = list(prop.comment)
                new_prop.comment.append(f"Original IRI: {prop.iri}")
                print(f"Created DataProperty: {new_prop.name} ({new_prop.iri})")
    return iri_dataprop_counter

def merge_importOnto_importClass_to_ontoBase(onto_base_path, import_ontology_path, merge_tasks, base_iri, output_dir, final_merged_file):
    # Load base ontology
    onto_base = get_ontology(onto_base_path).load()

    # Load import ontology
    onto_import = get_ontology(import_ontology_path).load()

    # Load SKOS ontology and define skos:definition as AnnotationProperty
    skos = get_ontology("http://www.w3.org/2004/02/skos/core").load()
    with skos:
        class definition(AnnotationProperty):
            namespace = skos

    with onto_import:
        sync_reasoner()

    # Initialize IRI counter
    iri_counter = get_existing_max_id(onto_base, base_iri) + 1
    print(f"Starting IRI counter from: {iri_counter}")
    iri_objprop_counter = get_existing_max_objprop_id(onto_base, base_iri) + 1
    iri_dataprop_counter = get_existing_max_dataprop_id(onto_base, base_iri) + 1

    cls_map = {}

    # Process each task
    for task in merge_tasks:
        import_class_iri = task["import_class_iri"]
        base_parent_iri = task["base_parent_iri"]
        output_file = task["output_file"]

        import_class = onto_import.search_one(iri=import_class_iri)
        if not import_class:
            print(f"Class not found in import ontology: {import_class_iri}")
            continue

        all_import_classes = {import_class} | get_all_subclasses(import_class)
        sorted_classes = topological_sort(all_import_classes)

        with onto_base:
            for cls in sorted_classes:
                if cls == import_class:
                    base_parent = onto_base.search_one(iri=base_parent_iri)
                    if not base_parent:
                        print(f"Parent class not found in base ontology: {base_parent_iri}")
                        break
                    new_cls = types.new_class(cls.name, (base_parent,))
                else:
                    original_parents = [p for p in cls.is_a if p in cls_map]
                    parent_classes = [cls_map[p] for p in original_parents] or [Thing]
                    new_cls = types.new_class(cls.name, tuple(parent_classes))

                new_cls.iri = base_iri + f"{iri_counter:05d}"
                iri_counter += 1
                
                # Label
                if cls.label:
                    for label in cls.label:
                        new_cls.label.append(format_label(label))
                elif not cls.name.isdigit():
                    new_cls.label.append(format_label(cls.name))

                # Copy annotations
                copy_annotation_properties(cls, new_cls)
                new_cls.comment.append(f"Original IRI: {cls.iri}")
                cls_map[cls] = new_cls
                

        onto_base.save(file=f"{output_dir}/{output_file}", format="rdfxml")
        print(f"Saved: {output_file}")
    

    # Preserve valid object properties
    iri_objprop_counter = preserve_valid_object_properties(onto_base, onto_import, cls_map, base_iri, iri_objprop_counter)
    # Preserve valid data properties
    iri_dataprop_counter = preserve_valid_data_properties(onto_base, onto_import, cls_map, base_iri, iri_dataprop_counter)

    print("Object properties before save:")
    for prop in onto_base.object_properties():
        print(f"- {prop.name} | domain: {[d.name for d in prop.domain]} | range: {[r.name for r in prop.range]}")

    print("Data properties before save:")
    for prop in onto_base.data_properties():
        print(f"- {prop.name} | domain: {[d.name for d in prop.domain]} | range: {[r for r in prop.range]}")

    
    onto_base.save(file=f"{output_dir}/{final_merged_file}", format="rdfxml")
    print(f"Saved: {final_merged_file}")
    

if __name__ == "__main__":
    
    parser = argparse.ArgumentParser(description="Merge selected branches from import ontology into base ontology.")
    parser.add_argument("--base", required=True, help="Path to base ontology OWL file")
    parser.add_argument("--import_onto", required=True, help="Path to import ontology OWL file")
    parser.add_argument("--tasks", required=True, help="Path to JSON file with merge task definitions")
    parser.add_argument("--base_iri", required=True, help="Base IRI for the merged ontology")
    parser.add_argument("--output_dir", required=True, help="Directory to save intermediate and final outputs")
    parser.add_argument("--final_output", required=True, help="Filename for final merged OWL file")

    args = parser.parse_args()

    
    with open(args.tasks, "r", encoding="utf-8") as f:
        merge_tasks = json.load(f)

    merge_importOnto_importClass_to_ontoBase(
        onto_base_path=args.base,
        import_ontology_path=args.import_onto,
        merge_tasks=merge_tasks,
        base_iri=args.base_iri,
        output_dir=args.output_dir,
        final_merged_file=args.final_output
    )

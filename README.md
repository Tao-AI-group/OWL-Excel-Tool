
## Overview

This tool provides bidirectional OWL-Excel conversion, supporting both:
- Class hierarchies with annotation properties
- Object and data properties

It also supports selective ontology merging. 

## Requirements

This project uses **Python 3.10.16**.

Install dependencies:
```bash
pip install -r requirements.txt
```

## Usage

### Basic Command:

#### 1. owl2excel: extract OWL to Excel

##### extract classes and annotation properties
```bash
python owl2excel/owl2excel_class_annotations.py -i your.owl -o output.xlsx
```

##### extract  object and data properties
```bash
python owl2excel/owl2excel_properties.py -i your.owl -o properties.xlsx

```

#### 2. excel2owl: convert Excel to OWL
##### Build OWL classes and annotation properties from Excel
```bash
python excel2owl/excel2owl_class_annotations.py  -e your.xlsx -u your_uri -o output.owl 
```
##### Add object and data properties from Excel to existing OWL
```bash
python excel2owl/excel2owl_properties.py -e your_relation.xlsx -i your.owl -o new_owl_relations.owl
```

#### 3. selective owl merging: Merge selected branches from imported ontology into base ontology
```bash
sh selective_owl_merging/selective_owl_merging.sh
```

Note: Please update the configuration in both merge_branches.json and selective_owl_merging.sh before running.

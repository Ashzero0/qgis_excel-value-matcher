from qgis.PyQt.QtWidgets import QFileDialog, QInputDialog, QMessageBox
from qgis.core import *
from PyQt5.QtCore import QVariant
import pandas as pd
import os

# 1. Get the active layer
layer = iface.activeLayer()
if not layer:
    raise Exception("No active layer found.")

# 2. Select the field from the layer
layer_fields = [f.name() for f in layer.fields()]
layer_field, ok = QInputDialog.getItem(None, "QGIS Field", "Select the layer field to compare:", layer_fields, 0, False)
if not ok:
    raise Exception("No field selected.")

# 3. Select multiple Excel files (possibly from different folders)
excel_files = []
while True:
    files, _ = QFileDialog.getOpenFileNames(None, "Select Excel files (Cancel to stop)", "", "Excel Files (*.xlsx)")
    if not files:
        break
    excel_files.extend(files)

if not excel_files:
    raise Exception("No file selected.")

# 4. Extract unique values from each Excel file
values_by_file = {}
for file in excel_files:
    df = pd.read_excel(file)
    columns = df.columns.tolist()

    excel_field, ok = QInputDialog.getItem(None, f"Field for {os.path.basename(file)}", "Select the column to compare:", columns, 0, False)
    if not ok:
        continue

    values = set(map(str, df[excel_field].dropna().unique()))
    values_by_file[os.path.basename(file)] = values

# 5. Create an in-memory layer
geom_type = QgsWkbTypes.displayString(layer.wkbType())
crs = layer.crs().authid()

memory_layer = QgsVectorLayer(f"{geom_type}?crs={crs}", "Matches", "memory")
provider = memory_layer.dataProvider()

# Copy fields from the source layer + add the "matched_files" field
provider.addAttributes(layer.fields())
provider.addAttributes([QgsField("matched_files", QVariant.String)])
memory_layer.updateFields()

# 6. Build a reverse index for faster matching {value: [files]}
value_index = {}
for file_name, values in values_by_file.items():
    for v in values:
        value_index.setdefault(v, []).append(file_name)

# 7. Iterate over layer features and add matches
matching_features = []
layer_field_index = layer.fields().indexFromName(layer_field)

for feature in layer.getFeatures():
    val = str(feature[layer_field_index])
    matched_files = value_index.get(val)

    if matched_files:
        new_feat = QgsFeature(memory_layer.fields())
        new_feat.setGeometry(feature.geometry())
        new_feat.setAttributes(feature.attributes() + ["; ".join(matched_files)])
        matching_features.append(new_feat)

# Add matching features to the memory layer
if matching_features:
    provider.addFeatures(matching_features)
    memory_layer.updateExtents()
    QgsProject.instance().addMapLayer(memory_layer)

QMessageBox.information(None, "Process Finished", f"{len(matching_features)} features added with matches.")

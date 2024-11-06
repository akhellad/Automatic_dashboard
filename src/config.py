import os
import json

BASE_IMAGE_PATH = os.path.join(os.getcwd(), "temp_images")
JSON_PATHS = ["image_data.json", "elements_to_export.json"]

with open(JSON_PATHS[0], "r", encoding="utf-8") as f:
    image_data = json.load(f)

with open(JSON_PATHS[1], "r", encoding="utf-8") as f:
    elements_to_export = json.load(f)
"""Downloads the Superstore sales dataset from Kaggle into data/.

If this fails, download the CSV by hand from
https://www.kaggle.com/datasets/vivek468/superstore-dataset-final
and save it as data/superstore.csv
"""
import shutil
from pathlib import Path

DATA_DIR = Path(__file__).resolve().parent / "data"
DATA_DIR.mkdir(exist_ok=True)
TARGET = DATA_DIR / "superstore.csv"

try:
    import kagglehub
    folder = Path(kagglehub.dataset_download("vivek468/superstore-dataset-final"))
    csv_files = list(folder.rglob("*.csv"))
    if not csv_files:
        raise FileNotFoundError("no CSV in the downloaded folder")
    shutil.copy(csv_files[0], TARGET)
    print("Saved:", TARGET)
except Exception as err:
    print("Automatic download failed:", err)
    print("Download the CSV manually (link at the top of this file) and save it as", TARGET)

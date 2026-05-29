import zipfile
import re
from pathlib import Path
from core.plugins.base import BasePlugin

class ArchiveIngester(BasePlugin):
    def run(self):
        conf = self.config['ingester']
        src_dir = Path(conf['source_dir'])
        target_base = Path(conf['target_tree'])
        report_type = conf['report_type']
        year = conf['year']

        if not src_dir.exists():
            print(f"Error: Directory {src_dir} does not exist.")
            return

        # Search for archives like "Branch_51_27.05.zip"
        for zip_path in src_dir.glob("*.zip"):
            filename = zip_path.stem
            
            # Parse branch and month
            branch_match = next((b for b in conf['branches'] if b.lower() in filename.lower()), None)
            month_match = re.search(r'[\._](\d{2})\b', filename)

            if not branch_match or not month_match:
                print(f"Skipped: Unrecognized archive format -> {zip_path.name}")
                continue

            month = month_match.group(1)
            
            # Build rigid tree: Y/<branch>/reports/<report_type>/<yyyy>/<mm>/
            dest_dir = target_base / branch_match / "reports" / report_type / year / month
            dest_dir.mkdir(parents=True, exist_ok=True)

            print(f"Extracting {zip_path.name} -> {dest_dir}")
            
            try:
                with zipfile.ZipFile(zip_path, 'r') as zip_ref:
                    zip_ref.extractall(dest_dir)
            except zipfile.BadZipFile:
                print(f"Error: Corrupted archive -> {zip_path.name}")

        print("Ingestion completed.")
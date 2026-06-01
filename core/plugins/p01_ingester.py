import logging
import os
from datetime import datetime
import zipfile
import re
from pathlib import Path
from core.plugins.base import BasePlugin


class ArchiveIngester(BasePlugin):
    def __init__(self, config):
        self.config = config.get("ingester", {})
        self.setup_logging()
        # Consolidating all parameters here
        self.source_dir = Path(self.config.get("source_dir", "./incoming"))
        self.target_tree = Path(self.config.get("target_tree", "./company_tree"))
        self.branches = self.config.get("branches", [])
        self.report_type = self.config.get("report_type", {})

    def setup_logging(self):
        log_file = self.config.get("log_file", "pipeline.log")

        logging.basicConfig(
            level=logging.INFO,
            format="%(asctime)s [%(levelname)s] %(message)s",
            datefmt="%Y-%m-%d %H:%M:%S",
            handlers=[
                logging.FileHandler(log_file, encoding="utf-8"),
                logging.StreamHandler(),
            ],
        )
        self.logger = logging.getLogger(__name__)

    def run(self):
        self.logger.setLevel(logging.DEBUG)
        self.logger.info("Starting Incoming archive ingestion process...")
        conf = self.config
        src_dir = Path(conf["source_dir"])
        target_base = Path(conf["target_tree"])
        report_type = conf["report_type"]
        year = conf["year"]

        if not src_dir.exists():
            self.logger.error(f"Error: Directory {src_dir} does not exist.")
            return

        # Search for archives like "Branch_51_27.05.zip"
        for zip_path in src_dir.glob("*.zip"):
            filename = zip_path.stem

            # Parse branch
            branch_match = next(
                (b for b in self.branches if b.lower() in filename.lower()), None
            )
            
            # Match a full date pattern: DD.MM (e.g., 01.06)
            date_match = re.search(r"(\d{2})\.(\d{2})", filename)

            if date_match:
                day = date_match.group(1)
                month = date_match.group(2)
            else:
                # Fallback to current month if no date pattern is found
                month = f"{datetime.now().month:02d}"
                self.logger.warning(
                    f"Could not parse date from {filename}. Defaulting to current month: {month}"
                )

            # If branch doesn't match, skip it. 
            # (We don't skip for missing date anymore, because we have the fallback above)
            if not branch_match:
                self.logger.debug(
                    f"Skipped: Unrecognized branch in archive name -> {filename}"
                )
                continue

            # Build rigid tree: Y/<branch>/reports/<report_type>/<yyyy>/<mm>/
            dest_dir = (
                target_base / branch_match / "reports" / report_type / year / month
            )
            dest_dir.mkdir(parents=True, exist_ok=True)

            self.logger.info(f"Extracting {zip_path.name} -> {dest_dir}")

            try:
                with zipfile.ZipFile(zip_path, "r") as zip_ref:
                    zip_ref.extractall(dest_dir)
            except zipfile.BadZipFile:
                self.logger.error(f"Error: Corrupted archive -> {zip_path.name}")

        self.logger.info("Ingestion completed.")

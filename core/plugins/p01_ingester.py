import logging
import os
import time
from datetime import datetime
import zipfile
import re
from pathlib import Path
from core.plugins.base import BasePlugin


class ArchiveIngester(BasePlugin):
    def __init__(self, config):
        self.config = config.get("ingester", {})
        # Consolidating all parameters here
        self.log_file = self.config.get("log_file", "pipeline.log")
        self.src_dir = Path(self.config.get("source_dir", "./incoming"))
        self.target_tree = Path(self.config.get("target_tree", "./company_tree"))
        self.branches = self.config.get("branches", [])
        self.all_reports_folder = self.config.get("report_group", "reports")
        self.report_types = self.config.get("report_types", {})
        # Dynamic fallback for the year
        self.year = str(self.config.get("year", datetime.now().year))

        self.setup_logging()

    def setup_logging(self):

        logging.basicConfig(
            level=logging.INFO,
            format="%(asctime)s [%(levelname)s] %(message)s",
            datefmt="%Y-%m-%d %H:%M:%S",
            handlers=[
                logging.FileHandler(self.log_file, encoding="utf-8"),
                logging.StreamHandler(),
            ],
        )
        self.logger = logging.getLogger(__name__)

    def run(self):
        self.logger.setLevel(logging.DEBUG)
        self.logger.info("Starting Incoming archive ingestion process...")

        if not self.src_dir.exists():
            self.logger.error(f"Error: Directory '{self.src_dir}' does not exist.")
            return

        # Search for archives like "Branch_51_27.05.zip"
        for zip_path in self.src_dir.glob("*.zip"):
            filename = zip_path.stem

            # 1. PARSE & CHECK BRANCH FIRST
            branch_match = next(
                (b for b in self.branches if b.lower() in filename.lower()), None
            )
            if not branch_match:
                self.logger.debug(
                    f"Skipped: Unrecognized branch in archive name -> '{filename}'"
                )
                continue

            # 2. PARSE DATE
            # Match a full date pattern: DD.MM (e.g., 01.06)
            date_match = re.search(r"(\d{2})\.(\d{2})", filename)
            if date_match:
                day = date_match.group(1)
                month = date_match.group(2)
            else:
                # Fallback to current month if no date pattern is found
                month = f"{datetime.now().month:02d}"
                self.logger.warning(
                    f"Could not parse date from '{filename}'. Defaulting to current month: '{month}'"
                )

            # 3. DETERMINE SPECIFIC REPORT FOLDER
            specific_report_folder = None
            for folder_name, match_strings in self.report_types.items():
                if any(s.lower() in filename.lower() for s in match_strings):
                    specific_report_folder = folder_name
                    break

            if not specific_report_folder:
                self.logger.debug(
                    f"Skipping '{filename}': No match strings found for any report type."
                )
                continue

            # 4. BUILD RIGID TREE
            # Y/<branch>/<all reports>/<report_type>/<yyyy>/<mm>/
            dest_dir = (
                self.target_tree
                / branch_match
                / self.all_reports_folder
                / specific_report_folder
                / self.year
                / month
            )
            dest_dir.mkdir(parents=True, exist_ok=True)

            self.logger.info(f"Extracting '{zip_path.name}' -> '{dest_dir}'")

            # 5. UNZIP TO DESTINATION WITH DATE/TIME PRESERVATION
            try:
                with zipfile.ZipFile(zip_path, "r") as zip_ref:
                    # Iterate through each file inside the archive to preserve original timestamps
                    for zip_info in zip_ref.infolist():

                        # Extract the file to the destination directory
                        extracted_path = zip_ref.extract(zip_info, dest_dir)

                        # Retrieve the original date tuple from the archive metadata
                        # Format: (year, month, day, hour, minute, second)
                        date_time = zip_info.date_time

                        # Convert the tuple into a system-readable timestamp
                        dt = datetime(*date_time)
                        timestamp = time.mktime(dt.timetuple())

                        # Forcibly apply the original access and modification times to the newly extracted file
                        os.utime(extracted_path, (timestamp, timestamp))

            except zipfile.BadZipFile:
                self.logger.error(f"Error: Corrupted archive -> '{zip_path.name}'")

        self.logger.info("Ingestion completed.")

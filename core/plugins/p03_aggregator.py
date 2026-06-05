import logging
import shutil
import re
from pathlib import Path
from datetime import datetime
import xlwings as xw


class Aggregator:
    def __init__(self, config):
        self.config = config.get("aggregator", {})

        # Paths & Topology
        self.target_tree = Path(
            self.config.get("target_tree", "./company_tree")
        ).resolve()
        self.total_node = self.config.get("total_node", "Total")
        self.reports_group = self.config.get("reports_group", "reports")
        self.branches = self.config.get("branches", [])

        self.report_types = self.config.get("report_types", {})
        self.excel_visible = self.config.get("show_excel", True)

        self._setup_dates()
        self._setup_logging()

    def _setup_dates(self):
        raw_date = self.config.get("target_date", "current")
        if raw_date.lower() == "current":
            self.target_datetime = datetime.now()
        else:
            self.target_datetime = datetime.strptime(raw_date, "%d.%m.%Y")

        self.target_date_str = self.target_datetime.strftime("%d.%m.%Y")
        self.folder_year = self.target_datetime.strftime("%Y")
        self.folder_month = self.target_datetime.strftime("%m")
        self.folder_day = self.target_datetime.strftime("%d")

    def _setup_logging(self):
        log_file = self.config.get("log_file", "aggregator.log")
        logging.basicConfig(
            level=logging.INFO,
            format="%(asctime)s [%(levelname)s] %(message)s",
            handlers=[
                logging.FileHandler(log_file, encoding="utf-8"),
                logging.StreamHandler(),
            ],
        )
        self.logger = logging.getLogger(__name__)

    def _find_previous_report_dir(self, report_base_dir):
        """Scans to find the most recent date before the target date."""
        if not report_base_dir.exists():
            return None
        date_folders = []
        for path in report_base_dir.rglob("*/*/*"):
            if path.is_dir():
                try:
                    parts = path.parts[-3:]
                    folder_date = datetime(int(parts[0]), int(parts[1]), int(parts[2]))
                    if folder_date.date() < self.target_datetime.date():
                        date_folders.append((folder_date, path))
                except ValueError:
                    continue
        if not date_folders:
            return None
        date_folders.sort(key=lambda x: x[0], reverse=True)
        return date_folders[0][1]

    def run(self):
        self.logger.info(
            f"--- Starting Aggregation Cycle for {self.target_date_str} ---"
        )

        for report_name, rules in self.report_types.items():
            self.logger.info(f"Aggregating '{report_name}'...")
            self._process_report(report_name, rules)

    def _process_report(self, report_name, rules):
        report_base_dir = (
            self.target_tree / self.total_node / self.reports_group / report_name
        )
        report_dir = (
            report_base_dir / self.folder_year / self.folder_month / self.folder_day
        )

        report_dir.mkdir(parents=True, exist_ok=True)

        # 1. Handle Historical Files Copy (So we have a file to update!)
        prev_dir = self._find_previous_report_dir(report_base_dir)
        if prev_dir:
            rel_prev = prev_dir.relative_to(self.target_tree)
            rel_target = report_dir.relative_to(self.target_tree)

            print(f"\n[?] FOUND PREVIOUS MASTER AGGREGATOR: {report_name}")
            print(f"    Source:      {rel_prev}")
            print(f"    Destination: {rel_target}")
            response = input(f"    Copy master file to today's folder? [Y/n]: ")

            if response.strip().lower() not in ["n", "no"]:
                for item in prev_dir.iterdir():
                    if item.is_file():
                        shutil.copy2(item, report_dir / item.name)
                self.logger.info(f"  Copied master files from {rel_prev}")
            else:
                self.logger.info("  Skipped historical files based on user input.")
        else:
            self.logger.warning(
                "  No previous master report found! Aggregation may fail if templates are missing."
            )

        # 2. Mutate Excel Links
        self._mutate_excel_links(report_dir, rules)

    def _mutate_excel_links(self, report_dir, rules):
        target_files = rules.get("target_files_to_update", [])
        if not target_files:
            return

        self.logger.info(
            f"  Starting xlwings link mutation engine. Visible={self.excel_visible}"
        )

        # Regex to precisely find the date folder segment in the path
        date_pattern = re.compile(r"[\\/]\d{4}[\\/]\d{2}[\\/]\d{2}[\\/]")

        with xw.App(visible=self.excel_visible) as app:
            if self.excel_visible:
                app.activate(steal_focus=True)
            app.display_alerts = False

            for file_name in target_files:
                target_path = report_dir / file_name
                if not target_path.exists():
                    self.logger.error(f"  Target file missing: {target_path}")
                    continue

                self.logger.info(f"  Opening master file: {target_path.name}")
                try:
                    wb_target = app.books.open(target_path, update_links=False)
                    links = wb_target.api.LinkSources(1)

                    if links:
                        for old_link in links:
                            match = date_pattern.search(old_link)
                            if match:
                                old_date_str = match.group()
                                # Grab the slash type (\ or /) from the original link
                                sep = old_date_str[0]
                                new_date_str = f"{sep}{self.folder_year}{sep}{self.folder_month}{sep}{self.folder_day}{sep}"

                                # Use literal string replacement to avoid regex escape crashes
                                new_link = old_link.replace(old_date_str, new_date_str)

                                if old_link != new_link:
                                    if Path(new_link).exists():
                                        self.logger.info(
                                            f"    Mutating Link: ...{Path(new_link).parent.name}"
                                        )
                                        wb_target.api.ChangeLink(
                                            Name=old_link, NewName=new_link, Type=1
                                        )
                                    else:
                                        self.logger.warning(
                                            f"    Cannot update link! Source file missing: {new_link}"
                                        )

                    app.calculate()

                    # 1. Save the Master File (with live formulas)
                    wb_target.save()
                    self.logger.info(f"  {target_path.name} master aggregation saved.")

                    # 2. Handle Step 4: Create Flat "Values-Only" Copy
                    # I assumed you meant YYYY.MM.DD based on your standard format
                    flat_name = f"{target_path.stem} - {self.folder_year}.{self.folder_month}.{self.folder_day}{target_path.suffix}"
                    flat_path = report_dir / flat_name

                    self.logger.info(
                        f"  Creating flattened values-only copy: {flat_name}"
                    )

                    # Save-As to create the new file
                    wb_target.save(flat_path)

                    # Loop through all sheets and overwrite formulas with their calculated values
                    for sheet in wb_target.sheets:
                        used_range = sheet.used_range
                        if used_range.address != "$A$1" or used_range.value is not None:
                            # Setting value = value instantly kills formulas
                            used_range.value = used_range.value

                    # Save the flat file and close
                    wb_target.save()
                    wb_target.close()
                    self.logger.info(f"  Values-only copy successfully generated.")

                except Exception as e:
                    self.logger.error(
                        f"  Failed to process master file {target_path.name}: {e}"
                    )

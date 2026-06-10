import logging
import shutil
from pathlib import Path
from datetime import datetime
import xlwings as xw


class Generator:
    def __init__(self, config):
        self.config = config.get("generator", {})

        # 1. Base Paths
        self.target_tree = Path(
            self.config.get("target_tree", "./company_tree")
        ).resolve()
        self.hq_branch = self.config.get("hq_branch", "HQ_51")
        self.reports_group = self.config.get("reports_group", "reports")

        # Branch-specific relative path resolution
        hq_path = self.target_tree / self.hq_branch

        # Strip leading slashes to prevent pathlib from treating it as a root C:/ path
        raw_ext_dir = self.config.get("external_inputs_dir", "external_inputs").lstrip(
            "/"
        )
        self.external_inputs_dir = (hq_path / raw_ext_dir).resolve()

        # 2. Excel & Data Rules
        self.constants_file = self.config.get("constants_file", "constants.xlsx")
        self.constants_sheet = self.config.get("constants_sheet", "Constants")
        self.constants_date_cell = self.config.get("constants_date_cell", "$A$2")
        self.excel_visible = self.config.get("show_excel", True)

        self.report_types = self.config.get("report_types", {})

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
        log_file = self.config.get("log_file", "generator.log")
        logging.basicConfig(
            level=logging.INFO,
            format="%(asctime)s [%(levelname)s] %(message)s",
            handlers=[
                logging.FileHandler(log_file, encoding="utf-8"),
                logging.StreamHandler(),
            ],
        )
        self.logger = logging.getLogger(__name__)

    def _is_scheduled_for_target(self, schedule_config):
        if not schedule_config:
            return True

        freq = schedule_config.get("frequency", "daily")

        if freq == "daily":
            return True
        if freq == "weekly":
            return self.target_datetime.strftime("%A") in schedule_config.get(
                "days_of_week", []
            )
        if freq == "monthly":
            return self.target_datetime.day in schedule_config.get("days_of_month", [])
        if freq == "exact_dates":
            return self.target_date_str in schedule_config.get("dates", [])

        return False

    def _find_previous_report_dir(self, report_base_dir):
        """Scans YYYY/MM/DD subfolders to find the most recent date before the target date."""
        if not report_base_dir.exists():
            return None

        date_folders = []
        # rglob finds all directories 3 levels deep (YYYY/MM/DD)
        for path in report_base_dir.rglob("*/*/*"):
            if path.is_dir():
                try:
                    parts = path.parts[-3:]
                    folder_date = datetime(int(parts[0]), int(parts[1]), int(parts[2]))

                    # Only consider folders older than our current target generation date
                    if folder_date.date() < self.target_datetime.date():
                        date_folders.append((folder_date, path))
                except ValueError:
                    continue  # Ignore non-date folders

        if not date_folders:
            return None

        # Sort descending by datetime and return the path of the most recent one
        date_folders.sort(key=lambda x: x[0], reverse=True)
        return date_folders[0][1]

    def _safe_save(self, wb, save_path=None, file_name=""):
        """Attempts to save the workbook. Pauses and prompts the user if the file is locked."""
        while True:
            try:
                if save_path:
                    wb.save(save_path)
                else:
                    wb.save()
                return True
            except PermissionError:
                # Catches native Python I/O locks
                print(f"\n[⚠️  FILE LOCKED] Windows cannot save '{file_name}'.")
                print("    It is currently open in Excel or locked by another process.")
                input(
                    "    -> Please close the file, then press [ENTER] here to retry..."
                )
            except Exception as e:
                error_str = str(e)
                # Catches Microsoft Excel COM locks (-2146827284 / 0x800A03EC)
                if "Cannot access" in error_str or "-2146827284" in error_str:
                    print(f"\n[⚠️  FILE LOCKED] Excel cannot overwrite '{file_name}'.")
                    print(
                        "    The destination file is currently open in another window."
                    )
                    input("    -> Please close it, then press [ENTER] here to retry...")
                else:
                    self.logger.error(f"  Unexpected error saving {file_name}: {e}")
                    return False

    def _safe_copy(self, src, dst, file_name=""):
        """Attempts to copy a file safely. Pauses if the destination is locked."""
        while True:
            try:
                shutil.copy2(src, dst)
                return True
            except PermissionError:
                print(f"\n[⚠️  FILE LOCKED] Windows cannot overwrite '{file_name}'.")
                print("    The destination file is currently open in Excel.")
                input("    -> Please close it, then press [ENTER] here to retry...")
            except Exception as e:
                self.logger.error(f"  Unexpected error copying {file_name}: {e}")
                return False

    def run(self):
        self.logger.info(
            f"--- Starting Generation Cycle for {self.target_date_str} ---"
        )

        for report_name, rules in self.report_types.items():
            schedule = rules.get("schedule", {})

            if not self._is_scheduled_for_target(schedule):
                self.logger.info(
                    f"Skipping '{report_name}' (Not scheduled for target date: {self.target_date_str})"
                )
                continue

            self.logger.info(f"Processing '{report_name}'...")
            self._process_report(report_name, rules)

    def _process_report(self, report_name, rules):
        report_base_dir = (
            self.target_tree / self.hq_branch / self.reports_group / report_name
        )
        report_dir = (
            report_base_dir / self.folder_year / self.folder_month / self.folder_day
        )

        # DRY Fix: Create the target directory ONCE at the very beginning
        report_dir.mkdir(parents=True, exist_ok=True)

        # 1. Handle Historical Files Copy
        prev_dir = self._find_previous_report_dir(report_base_dir)
        if prev_dir:
            # Format clean relative paths starting from the HQ branch
            rel_prev = prev_dir.relative_to(self.target_tree)
            rel_target = report_dir.relative_to(self.target_tree)

            # Interactive Terminal Prompt
            print(f"\n[?] FOUND PREVIOUS REPORT: {report_name}")
            print(f"    Source:      {rel_prev}")
            print(f"    Destination: {rel_target}")
            response = input("    Copy all files from source to destination? [Y/n]: ")

            if response.strip().lower() not in ["n", "no"]:
                for item in prev_dir.iterdir():
                    if item.is_file():
                        self._safe_copy(
                            item, report_dir / item.name, file_name=item.name
                        )

                self.logger.info(f"  Copied historical files from {rel_prev}")
            else:
                self.logger.info("  Skipped historical files based on user input.")
        else:
            self.logger.info("  No previous report found to copy.")

        # 2. Inject External Manual Files
        external_files = rules.get("external_files_to_copy", [])
        for file_name in external_files:
            source_path = (self.external_inputs_dir / file_name).resolve()
            dest_path = report_dir / Path(file_name).name
            input("    -> Please press [ENTER] to continue...")

            if source_path.exists():
                self._safe_copy(source_path, dest_path)
                self.logger.info(f"  Copied external input: {file_name}")
            else:
                self.logger.warning(f"  MISSING external input file: {source_path}")

        # 3. Trigger Excel modifications
        self._update_excel_links(report_dir, rules)

    def _update_excel_links(self, report_dir, rules):
        target_files = rules.get("target_files_to_update", [])
        constants_path = report_dir / self.constants_file

        if not target_files and not constants_path.exists():
            return

        self.logger.info(f"  Starting xlwings engine. Visible={self.excel_visible}")

        with xw.App(visible=self.excel_visible) as app:
            if self.excel_visible:
                try:
                    # Attempt to force Excel to the front
                    app.activate(steal_focus=True)
                except Exception:
                    # If it's a headless/zombie process, ignore the UI error and keep working
                    self.logger.warning(
                        "  Could not steal Windows focus. Excel will process in the background."
                    )

            # -------------------------------------------------------------
            # 1. PROCESS CONSTANTS FILE
            # -------------------------------------------------------------
            if constants_path.exists():
                self.logger.info(f"  Injecting date into {self.constants_file}")
                try:
                    # Open WITHOUT updating links to prevent old-date data pulls
                    wb_const = app.books.open(constants_path, update_links=False)
                    sheet = wb_const.sheets[self.constants_sheet]

                    # Inject the native Python date object
                    sheet.range(
                        self.constants_date_cell
                    ).value = self.target_datetime.date()

                    # Ask the native Windows COM API for all external link sources
                    links = wb_const.api.LinkSources()
                    if links:
                        self.logger.info(
                            "  Refreshing external links in constants file..."
                        )
                        for link in links:
                            wb_const.api.UpdateLink(Name=link)

                    # Force a full formula recalculation
                    app.calculate()
                    self._safe_save(wb_const, file_name=self.constants_file)
                    wb_const.close()
                    self.logger.info("  Constants successfully updated and saved.")
                except Exception as e:
                    self.logger.error(f"  Failed to process constants file: {e}")

            # -------------------------------------------------------------
            # 2. PROCESS TARGET FILES
            # -------------------------------------------------------------
            for file_name in target_files:
                target_path = report_dir / file_name
                self.logger.info(f"  Processing target file: {target_path.name}")
                try:
                    # We open these WITH link updates enabled, so they immediately
                    # pull the fresh data from the constants file we just saved.
                    wb_target = app.books.open(target_path, update_links=True)

                    # Force recalculation of internal formulas
                    app.calculate()
                    self._safe_save(wb_target, file_name=file_name)
                    wb_target.close()
                    self.logger.info(
                        f"  {target_path.name} successfully updated and saved."
                    )
                except Exception as e:
                    self.logger.error(
                        f"  Failed to process target file {target_path.name}: {e}"
                    )

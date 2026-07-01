import logging
import shutil
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
        # Link path mutation settings (Fallback to None if not present)
        self.path_replace_from = self.config.get("path_replace_from", None)
        self.path_replace_to = self.config.get("path_replace_to", None)

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

            print(f"\n[?] FOUND PREVIOUS MASTER AGGREGATION: {report_name}")
            print(f"    Source: {rel_prev}")
            response = input(
                f"    Copy to target destination ({self.target_date_str})? [Y/n]: "
            )

            if response.strip().lower() not in ["n", "no"]:
                for item in prev_dir.iterdir():
                    if item.is_file():
                        self._safe_copy(
                            item, report_dir / item.name, file_name=item.name
                        )
                self.logger.info(f"  Copied previous master files from {rel_prev}")
            else:
                self.logger.info("  Skipped previous master files based on user input.")
        else:
            self.logger.warning(
                "  No previous master files found! Aggregation may fail if templates are missing."
            )

        # 2. Mutate Excel Links
        self._mutate_excel_links(report_dir, report_name, rules)

    def _mutate_excel_links(self, report_dir, report_name, rules):
        target_files = rules.get("target_files_to_update", [])
        if not target_files:
            return

        self.logger.info(
            f"  Starting xlwings link mutation engine. Visible={self.excel_visible}"
        )

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
                            # 1. Identify which branch this link belongs to
                            target_branch = None
                            for branch in self.branches:
                                # Look for \Branch\ or /Branch/ in the old legacy link
                                if (
                                    f"\\{branch}\\" in old_link
                                    or f"/{branch}/" in old_link
                                ):
                                    target_branch = branch
                                    break

                            # If we can't identify the branch, fallback to the old mutation logic
                            if not target_branch:
                                self.logger.warning(
                                    f"    Could not identify branch in link: {old_link}"
                                )
                                continue

                            # 2. Extract the target file name
                            target_file_name = Path(old_link).name

                            # 3. Dynamically reconstruct the pristine absolute path using our rigid structure
                            # target_tree is already absolute: e.g., F:\A0807\04_WORKPLACE\Company
                            new_link_path = (
                                self.target_tree
                                / target_branch
                                / self.reports_group
                                / report_name
                                / self.folder_year
                                / self.folder_month
                                / self.folder_day
                                / target_file_name
                            ).resolve()

                            new_link = str(new_link_path)

                            # 3. Apply ChangeLink if any mutation occurred
                            if old_link != new_link:
                                if Path(new_link).exists():
                                    # Grabs the last 8 parts of the path (e.g., "HQ/Branch/Reports/Report/2026/06/04/file.name")
                                    short_path = "\\".join(Path(new_link).parts[-8:])
                                    self.logger.info(
                                        f"    Mutating Link -> ...\\{short_path}"
                                    )

                                    wb_target.api.ChangeLink(
                                        Name=old_link, NewName=new_link, Type=1
                                    )
                                else:
                                    self.logger.warning(
                                        f"    Cannot update link! Source file missing: {new_link}"
                                    )

                    # Force calculation now that links are updated
                    app.calculate()

                    # 1. Save the Master File safely
                    self._safe_save(wb_target, file_name=target_path.name)
                    self.logger.info(f"  {target_path.name} master aggregation saved.")

                    # 2. Handle Step 4: Create Flat "Values-Only" Copy
                    flat_name = f"{target_path.stem} - {self.folder_year}.{self.folder_month}.{self.folder_day}{target_path.suffix}"
                    flat_path = report_dir / flat_name

                    self.logger.info(
                        f"  Creating flattened values-only copy: {flat_name}"
                    )

                    # Save-As safely
                    self._safe_save(wb_target, save_path=flat_path, file_name=flat_name)

                    # Overwrite formulas with values
                    for sheet in wb_target.sheets:
                        used_range = sheet.used_range
                        if used_range.address != "$A$1" or used_range.value is not None:
                            # Setting value = value instantly kills formulas
                            used_range.value = used_range.value

                    # Save the flat file safely and close
                    self._safe_save(wb_target, file_name=flat_name)
                    wb_target.close()
                    self.logger.info("  Values-only copy successfully generated.")

                except Exception as e:
                    self.logger.error(
                        f"  Failed to process master file {target_path.name}: {e}"
                    )

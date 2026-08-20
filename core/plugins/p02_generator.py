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
        # Support old variable 'hq_branch' and new array 'branches'
        legacy_hq = self.config.get("hq_branch", "HQ_51")
        raw_branches = self.config.get("branches", legacy_hq)
        self.branches = (
            raw_branches if isinstance(raw_branches, list) else [raw_branches]
        )

        self.reports_group = self.config.get("reports_group", "reports")

        # Strip leading slashes to prevent pathlib from treating it as a root C:/ path
        # Save only relative path. Absolute path will be dynamically created for every  branch.
        self.raw_ext_dir = self.config.get(
            "external_inputs_dir", "external_inputs"
        ).lstrip("/")

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
        if freq == "quarterly":
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

            for branch in self.branches:
                self.logger.info(f"Processing '{report_name}' for '{branch}'...")
                self._process_report(branch, report_name, rules)

    def _process_report(self, branch, report_name, rules):
        report_base_dir = self.target_tree / branch / self.reports_group / report_name
        report_dir = (
            report_base_dir / self.folder_year / self.folder_month / self.folder_day
        )

        # DRY Fix: Create the target directory ONCE at the very beginning
        report_dir.mkdir(parents=True, exist_ok=True)

        # 1. Recursive Historical Copy (Handles nested folders)
        prev_dir = self._find_previous_report_dir(report_base_dir)
        if prev_dir:
            rel_prev = prev_dir.relative_to(self.target_tree)
            print(f"\n[?] FOUND PREVIOUS REPORT GENERATION: {report_name}")
            print(f"    Source: {rel_prev}")
            response = input(
                f"    Copy to target destination ({self.target_date_str})? [Y/n]: "
            )

            if response.strip().lower() not in ["n", "no"]:
                # Use rglob to recursively grab all files in subfolders
                for item in prev_dir.rglob("*"):
                    if item.is_file():
                        rel_path = item.relative_to(prev_dir)
                        dst_path = report_dir / rel_path
                        dst_path.parent.mkdir(parents=True, exist_ok=True)
                        self._safe_copy(item, dst_path, file_name=item.name)
                self.logger.info(f"  Copied previous report files from {rel_prev}")
            else:
                self.logger.warning(
                    "  Skipped previous report files based on user input."
                )
        else:
            self.logger.warning(
                "  No previous report files found! Generation may fail if templates are missing."
            )

        # 2. Copy Standard External Inputs (Look in the current branch dir)
        branch_ext_dir = (self.target_tree / branch / self.raw_ext_dir).resolve()
        print("\n[?] EXTERNAL FILES PROCESSING")
        print(f"    Source: {branch_ext_dir}")
        response = input(
            f"    Copy to target destination ({self.target_date_str})? [Y/n]: "
        )

        if response.strip().lower() not in ["n", "no"]:
            ext_files = rules.get("external_files_to_copy", [])
            for file_name in ext_files:
                src_file = branch_ext_dir / file_name
                if src_file.exists():
                    self._safe_copy(
                        src_file, report_dir / src_file.name, file_name=src_file.name
                    )
                    self.logger.info(f"  Copied external input: {src_file}")
                else:
                    self.logger.warning(f"  External input missing: {src_file}")
        else:
            self.logger.warning("  Skipped external files based on user input.")

        # 3. Handle Special Conditional Routing (e.g., 1st of month fallback)
        special_files = rules.get("special_files", [])
        for sf in special_files:
            target_path = report_dir / sf["target_path"]
            # If historical copy didn't bring it over, trigger fallback
            if not target_path.exists():
                if sf.get("date_rule") == "first_of_month":
                    fallback_date = self.target_datetime.replace(day=1)
                else:
                    fallback_date = self.target_datetime

                alt_dir = (
                    self.target_tree
                    / branch
                    / sf["fallback_group"]
                    / sf["fallback_report"]
                    / fallback_date.strftime("%Y")
                    / fallback_date.strftime("%m")
                    / fallback_date.strftime("%d")
                )
                src_file = alt_dir / sf["fallback_file"]

                if src_file.exists():
                    target_path.parent.mkdir(parents=True, exist_ok=True)
                    self._safe_copy(src_file, target_path, file_name=src_file.name)
                    self.logger.info(
                        f"  Routed special fallback file to {sf['target_path']}"
                    )
                else:
                    self.logger.warning(f"  Special fallback missing at: {src_file}")

        # 4. Trigger Priority-Tiered Excel Engine
        self._update_excel_links(report_dir, rules)

    def _update_excel_links(self, report_dir, rules):
        update_sequence = rules.get("update_sequence", [])
        legacy_targets = rules.get("target_files_to_update", [])

        # Automatically fall back to global constants config unless explicitly overridden
        constants_cfg = rules.get("constants", {})
        const_file_name = constants_cfg.get("file", self.constants_file)
        const_sheet = constants_cfg.get("sheet", self.constants_sheet)
        const_date_cell = constants_cfg.get("date_cell", self.constants_date_cell)

        const_file = report_dir / const_file_name

        if not const_file.exists() and not update_sequence and not legacy_targets:
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
            app.display_alerts = False

            # --- PRIORITY 1: CONSTANTS ---
            # Automatically runs if the file exists in the folder, maintaining backward compatibility
            if const_file.exists():
                self.logger.info(
                    f"  [Priority 1] Updating Constants: {const_file.name}"
                )
                try:
                    wb_const = app.books.open(const_file, update_links=False)
                    wb_const.sheets[const_sheet].range(
                        const_date_cell
                    ).value = self.target_datetime

                    links = wb_const.api.LinkSources(1)
                    if links:
                        for link in links:
                            wb_const.api.UpdateLink(Name=link)

                    app.calculate()
                    self._safe_save(wb_const, file_name=const_file.name)
                    wb_const.close()
                except Exception as e:
                    self.logger.error(f"  Failed to process constants: {e}")

            # --- PRIORITY TIERS: INTERMEDIATES & TARGETS ---
            tiers = update_sequence if update_sequence else [legacy_targets]

            for tier_idx, file_group in enumerate(tiers, start=2):
                if not file_group:
                    continue
                self.logger.info(
                    f"  [Priority {tier_idx}] Executing Tier Update Sequence..."
                )

                for file_name in file_group:
                    target_path = report_dir / file_name
                    if not target_path.exists():
                        self.logger.warning(f"  Missing file in sequence: {file_name}")
                        continue

                    self.logger.info(
                        f"    -> Updating links & calculating: {target_path.name}"
                    )
                    try:
                        wb_target = app.books.open(target_path, update_links=True)
                        app.calculate()
                        self._safe_save(wb_target, file_name=target_path.name)
                        wb_target.close()
                    except Exception as e:
                        self.logger.error(
                            f"    Failed to process {target_path.name}: {e}"
                        )

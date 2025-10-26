import os
import re
import pandas as pd

from PyQt6.QtWidgets import QMessageBox, QInputDialog, QDialog
from ui.file_ui import FileUI
from models import TestScenarioModel, ObjRepoModel


class FileController:
    def __init__(self, main):
        self.main = main

    def new_file(self):
        module_name, ok = QInputDialog.getText(self.main, "Module Name", "Enter module name: eg:- QuickSanity")
        if not ok:
            return
        if not module_name or not module_name.strip():
            QMessageBox.warning(self.main, "Invalid Input", "Enter Module Name")
            return
        module_name = module_name.strip()

        module_abbr, ok = QInputDialog.getText(self.main, "Module Abbreviation", f"Enter 2-letter abbreviation for {module_name}:  eg:- QS")
        if not ok:
            return
        module_abbr = module_abbr.strip().upper()[:2] if module_abbr.strip() else module_name[:2].upper()

        tab = self.main.create_new_tab(module_name)

        test_scenario_df = TestScenarioModel.create_preset(module_name, module_abbr)
        tab.model1.loadData(test_scenario_df)

        objects_df = ObjRepoModel.create_preset()
        tab.model2.loadData(objects_df)

        tab.table1.auto_adjust_cells()
        tab.table2.auto_adjust_cells()

    def open_file(self):
        ui = FileUI(main=self.main,mode="open")
        if ui.exec() != QDialog.DialogCode.Accepted:
            return

        scenario_path, obj_path = ui.get_selected_files_path()

        if not (scenario_path and obj_path):
            return

        try:
            scenario_name = os.path.basename(scenario_path)
            module_match = re.search(r'Automation_Module_([^.]+)\.xlsx', scenario_name)
            module_name = module_match.group(1) if module_match else "Untitled"
            project_name = ui.selected_project

            tab = self.main.create_new_tab(module_name, scenario_path, obj_path, project_name)
            tab.scenario_path = scenario_path
            tab.obj_path = obj_path

            test_scenario_df = pd.read_excel(
                scenario_path,
                sheet_name="TestScenario",
                header=None,
                dtype=object,
                keep_default_na=False,
                engine='openpyxl'
            )
            objects_df = pd.read_excel(
                obj_path,
                sheet_name="Objects",
                header=None,
                dtype=object,
                keep_default_na=False,
                engine='openpyxl'
            )

            tab.model1.loadData(test_scenario_df)
            tab.model2.loadData(objects_df)

            scenario_table_view = tab.table1.table
            scenario_table_view.setColumnWidth(0, 50)  # Type
            scenario_table_view.setColumnWidth(1, 130) # ID
            scenario_table_view.setColumnWidth(2, 50)  # Skip
            scenario_table_view.setColumnWidth(3, 300) # Description
            scenario_table_view.setColumnWidth(4, 350) # Steps Performed
            scenario_table_view.setColumnWidth(5, 50) # Expected Results
            scenario_table_view.setColumnWidth(6, 250) # Command
            for col_index in range(7, 12):
                scenario_table_view.setColumnWidth(col_index, 250)

            tab.table2.auto_adjust_cells()

            tab.mark_saved()

        except Exception as e:
            QMessageBox.critical(self.main, "Error", f"Failed to load files:\n{str(e)}")

    def save_file(self):
        tab = self.main.get_current_tab()
        if tab:
            if tab.model1.df.empty and tab.model2.df.empty:
                return

            scenario_path = tab.scenario_path
            obj_path = tab.obj_path

            if scenario_path and obj_path:
                try:
                    self._write_files(scenario_path, obj_path)
                    tab.mark_saved()
                except (FileNotFoundError, PermissionError) as e:
                    QMessageBox.critical(self.main, "Error", f"Permission denied or path not found when saving files:\n{str(e)}")
                    return False
                except Exception as e:
                    QMessageBox.critical(self.main, "Error", f"An unexpected error occurred while saving files:\n{str(e)}")
                    return False
            else:
                return self.save_file_as(tab)

    def save_file_as(self, tab):
        if not tab:
            tab = self.main.get_current_tab()

        ui = FileUI(main=self.main,mode="save")
        if ui.exec() != QDialog.DialogCode.Accepted:
            return

        scenario_path, obj_path = ui.get_selected_files_path()
        if not (scenario_path and obj_path):
            return

        project_name = ui.selected_project
        module_match = re.search(r'Automation_Module_([^.]+)\.xlsx', os.path.basename(scenario_path))
        new_module_name = module_match.group(1) if module_match else "Untitled"

        existing_tab = self.main.find_existing_tab(project_name, new_module_name)
        if existing_tab and existing_tab != tab:
            QMessageBox.warning(
                self.main,
                "File Already Open",
                f"'{project_name}/{new_module_name}' is already open in another tab. "
                "Close that tab first or choose a different name."
            )
            return False

        if os.path.exists(scenario_path) or os.path.exists(obj_path):
            reply = QMessageBox.question(
                self.main,
                "Warning",
                "File with same name already exists! Overwrite?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
            )
            if reply != QMessageBox.StandardButton.Yes:
                return False

        try:
            os.makedirs(os.path.dirname(scenario_path), exist_ok = True)
            os.makedirs(os.path.dirname(obj_path), exist_ok = True)

            self._write_files(scenario_path, obj_path)

            tab.scenario_path  = scenario_path
            tab.obj_path = obj_path
            tab.project_name = ui.selected_project
            tab.mark_saved()
            return True

        except Exception as e:
            QMessageBox.critical(
                self.main, "Error",
                f"Failed to save files:\n{str(e)}"
            )
            return False

    def _write_files(self, scenario_path, obj_path):
        tab = self.main.get_current_tab()
        if not tab:
            return
        try:
            with pd.ExcelWriter(scenario_path, engine="openpyxl") as Writer:
                if not tab.model1.df.empty:
                    tab.model1.df.to_excel(
                        Writer, sheet_name="TestScenario",
                        index=False, header=False
                    )
            with pd.ExcelWriter(obj_path, engine="openpyxl") as Writer:
                if not tab.model2.df.empty:
                    tab.model2.df.to_excel(
                        Writer, sheet_name="Objects",
                        index=False, header=False
                    )
        except Exception as e:
            raise Exception(f"Failed to write Excel files: {str(e)}")

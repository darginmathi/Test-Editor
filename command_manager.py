import pandas as pd
import os
import re
from typing import List, Dict, Tuple

class CommandManager:
    def __init__(self, data_directory=None):
        self.commands_df = None
        self.all_commands = []
        self.command_acronyms = {}
        self.object_names = set()
        self.data_directory = data_directory
        self.load_commands()

    def set_data_directory(self, data_directory):
        """Set the data directory and reload commands"""
        self.data_directory = data_directory
        self.load_commands()

    def load_commands(self, file_path=None):
        """Load commands from E2E Commands file"""
        if file_path is None:
            # First try: relative to data directory
            if self.data_directory:
                data_dir_path = os.path.join(self.data_directory, "E2E Commands.xlsx")
                if os.path.exists(data_dir_path):
                    file_path = data_dir_path
                    print(f"✓ Found E2E Commands file in data directory: {file_path}")

            # Second try: various relative paths
            if not file_path:
                search_paths = [
                    "./data/E2E Commands.xlsx",
                    "../data/E2E Commands.xlsx",
                    "../../data/E2E Commands.xlsx",
                    "E2E Commands.xlsx",
                    "./E2E Commands.xlsx",
                ]

                for path in search_paths:
                    if os.path.exists(path):
                        file_path = path
                        print(f"✓ Found E2E Commands file at: {path}")
                        break

        if file_path and os.path.exists(file_path):
            try:
                self.commands_df = pd.read_excel(file_path, sheet_name="Sheet1")
                self._process_commands()
                print(f"✓ Loaded {len(self.all_commands)} commands from {file_path}")
            except Exception as e:
                print(f"❌ Error loading commands file: {e}")
                self._create_default_commands()
        else:
            print("❌ E2E Commands file not found")
            print("Please ensure 'E2E Commands.xlsx' is in your data directory")
            self._create_default_commands()

    def _process_commands(self):
        """Process the commands dataframe"""
        self.all_commands = []
        self.command_acronyms = {}

        if self.commands_df is not None:
            for index, row in self.commands_df.iterrows():
                # Skip empty rows and header rows
                if pd.isna(row.iloc[0]) or str(row.iloc[0]).startswith('('):
                    continue

                command = str(row.iloc[0]).strip()
                description = str(row.iloc[1]) if pd.notna(row.iloc[1]) else ""

                if command and command.startswith(('FWC_', 'EC_')):
                    self.all_commands.append(command)
                    # Generate acronym
                    acronym = self._generate_acronym(command)
                    if acronym:
                        self.command_acronyms[command] = acronym

            print(f"✓ Processed {len(self.all_commands)} valid commands")

    def _create_default_commands(self):
        """Create default commands if file not found"""
        default_commands = [
            "FWC_ClickButton", "FWC_SetTextBox", "FWC_SetDropdown", "FWC_VerifyText",
            "FWC_VerifyButtonExist", "FWC_VerifyTextBoxExist", "EC_Login", "EC_SetTextBox",
            "EC_VerifyCellData", "EC_TakeScreenShot", "FWC_OpenURL", "FWC_WaitForSecond"
        ]
        self.all_commands = default_commands
        for cmd in default_commands:
            acronym = self._generate_acronym(cmd)
            if acronym:
                self.command_acronyms[cmd] = acronym
        print("⚠ Using default commands")

    def _generate_acronym(self, command: str) -> str:
        """Generate acronym from command (e.g., FWC_ClickButton -> CB)"""
        # Remove prefix
        clean_command = command.replace('FWC_', '').replace('EC_', '')

        # Split by capital letters and take first letter of each word
        words = re.findall('[A-Z][a-z]*', clean_command)
        if words:
            return ''.join(word[0] for word in words)
        return ""

    def search_commands(self, search_text: str) -> List[str]:
        """Search commands with both basic and acronym matching"""
        if not search_text:
            return self.all_commands

        search_text = search_text.upper()
        results = []

        # Basic search (case-insensitive contains)
        basic_matches = [cmd for cmd in self.all_commands if search_text in cmd.upper()]

        # Acronym search
        acronym_matches = []
        for cmd, acronym in self.command_acronyms.items():
            if search_text in acronym:
                acronym_matches.append(cmd)

        # Combine and remove duplicates, prioritize acronym matches
        combined = list(dict.fromkeys(acronym_matches + basic_matches))
        return combined

    def get_command_description(self, command: str) -> str:
        """Get description for a command"""
        if self.commands_df is not None:
            for index, row in self.commands_df.iterrows():
                if str(row.iloc[0]) == command:
                    return str(row.iloc[1]) if pd.notna(row.iloc[1]) else ""
        return ""

    def update_object_names(self, objects_df):
        """Update object names from Objects spreadsheet - use column B (index 1)"""
        self.object_names.clear()
        if objects_df is not None and not objects_df.empty:
            print(f"Processing Objects dataframe with {len(objects_df)} rows and {len(objects_df.columns)} columns")

            # Column B is index 1 - "User friendly name of Object"
            for index, row in objects_df.iterrows():
                if len(row) > 1 and pd.notna(row.iloc[1]):
                    obj_name = str(row.iloc[1]).strip()
                    # Skip header rows and empty names
                    if (obj_name and
                        obj_name != "User friendly name of Object" and
                        obj_name != "Type" and
                        not obj_name.startswith('---')):
                        self.object_names.add(obj_name)
                        print(f"Added object: {obj_name}")

            # Convert to sorted list for consistent ordering
            self.object_names = set(sorted(self.object_names))
            print(f"✓ Updated object names: {len(self.object_names)} objects")

    def search_objects(self, search_text: str) -> List[str]:
        """Search object names"""
        if not search_text:
            return sorted(list(self.object_names))

        search_text = search_text.lower()
        matches = [obj for obj in self.object_names if search_text in obj.lower()]
        return sorted(matches)

# Global command manager instance
command_manager = CommandManager()

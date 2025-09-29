# Constants and configurations

# Column configurations for Test Scenario
TEST_SCENARIO_COLUMNS = [
    "Type", "ID", "Skip", "Description", "Steps Performed",
    "Expected Results", "Command", "Data1", "Data2", "Data3"
]

OBJECTS_COLUMNS = [
    "Type", "User friendly name of Object", "By-Type", "Webdriver friendly name of Object"
]

# File patterns
SCENARIO_FILE_PATTERN = "Automation_Module_{}.xlsx"
OBJECTS_FILE_PATTERN = "ObjRep_Module_{}_Test.xlsx"

# Directory names
TEST_SUITES_DIR = "testSuites"
OBJECT_REPOSITORIES_DIR = "objectRepositories"

# Settings file for remembering last directory
SETTINGS_FILE = "test_editor_settings.ini"

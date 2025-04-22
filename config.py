import pandas as pd
from typing import Dict, Any


class ConfigLoader:
    """Handles loading environment variables and configuration from Excel"""

    def __init__(self, xlsx_path: str):
        self.xlsx_path = xlsx_path

    def load_environment(self) -> Dict[str, str]:
        """Load environment variables from the first sheet of the Excel file"""
        environment_vars = {}
        try:
            # Use header=None to ensure it reads from the very first row
            env_df = pd.read_excel(self.xlsx_path, sheet_name=0, header=None)
            # Take only first two columns as key-value pairs
            # Filter out rows where the first column (key) is NaN or empty after stripping
            env_df = env_df.dropna(subset=[0])
            env_df[0] = env_df[0].astype(str).str.strip()  # Ensure key is string and strip whitespace
            env_df = env_df[env_df[0] != '']  # Remove rows where key is empty after strip

            for _, row in env_df.iterrows():
                key = row.iloc[0]
                # Handle potential NaN in the second column by treating it as empty string
                value = str(row.iloc[1]) if pd.notna(row.iloc[1]) else ""
                environment_vars[key] = value
            print(f"Loaded {len(environment_vars)} environment variables")
        except FileNotFoundError:
            print(f"Error: Environment file not found at '{self.xlsx_path}' when loading environment.")
        except Exception as e:
            print(f"Error loading environment variables from sheet 1: {e}")

        return environment_vars
import pandas as pd

class DataLoaderAgent:
    def __init__(self):
        self.df = None
        self.num_cols = 0
        self.other_cols = []
        self.data_types = {}
        self.stats = {}

    def load_data(self, csv_file):
        csv_file.seek(0)
        try:
            self.df = pd.read_csv(csv_file)
            if self.df.empty:
                return False, "CSV file is empty."
            self.num_cols = len(self.df.columns)
            self.detect_data_types()
            self.analyze_data()
            return True, f"Loaded with {len(self.df)} rows and {self.num_cols} columns."
        except Exception as e:
            return False, f"Error reading CSV: {str(e)}"

    def detect_data_types(self):
        self.data_types = {col: str(self.df[col].dtype) for col in self.df.columns}

    def analyze_data(self):
        self.stats = {}
        for col in self.df.columns:
            stats = {}
            if pd.api.types.is_numeric_dtype(self.df[col]):
                stats['mean'] = f"{self.df[col].mean():.2f}"
                stats['min'] = f"{self.df[col].min():.2f}"
                stats['max'] = f"{self.df[col].max():.2f}"
                stats['std'] = f"{self.df[col].std():.2f}"
            stats['unique'] = str(self.df[col].nunique())
            stats['top'] = str(self.df[col].mode()[0]) if not self.df[col].mode().empty else "N/A"
            self.stats[col] = stats
        # Compute correlations
        numeric_cols = [col for col in self.df.columns if pd.api.types.is_numeric_dtype(self.df[col])]
        if len(numeric_cols) > 1:
            corr_matrix = self.df[numeric_cols].corr()
            for col1 in numeric_cols:
                for col2 in numeric_cols:
                    if col1 != col2:
                        self.stats[col1][f"corr_with_{col2}"] = f"{corr_matrix.loc[col1, col2]:.2f}"

    def set_column(self, col):
        self.other_cols = [c for c in self.df.columns if c != col]
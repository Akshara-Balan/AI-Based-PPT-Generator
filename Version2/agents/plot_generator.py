# agents/plot_generator.py
import matplotlib.pyplot as plt
import tempfile
import pandas as pd

class PlotGeneratorAgent:
    def generate_plot(self, df, col, other_col, plot_type):
        with tempfile.NamedTemporaryFile(suffix='.png', delete=False) as tmp:
            chart_path = tmp.name
            plt.figure(figsize=(8, 5))
            is_numeric_col = pd.api.types.is_numeric_dtype(df[col])
            is_numeric_other = pd.api.types.is_numeric_dtype(df[other_col])
            actual_plot_type = plot_type
            
            if is_numeric_col and is_numeric_other:
                if plot_type == "Scatter":
                    df.plot.scatter(x=col, y=other_col, color='teal', alpha=0.5)
                    plt.title(f"{col} vs {other_col}", fontsize=12)
                elif plot_type == "Hexbin":
                    plt.hexbin(df[col], df[other_col], gridsize=20, cmap='Blues', mincnt=1)
                    plt.colorbar(label='Count')
                    plt.title(f"{col} vs {other_col}", fontsize=12)
                elif plot_type == "Box":
                    bins = pd.cut(df[col], bins=3)
                    df_box = df[[other_col]].copy()
                    df_box['Binned_' + col] = bins
                    df_box.boxplot(column=other_col, by='Binned_' + col, grid=False, patch_artist=True)
                    plt.title(f"{other_col} by {col}", fontsize=12)
                    plt.suptitle('')
                elif plot_type == "Bar":
                    bins = pd.cut(df[col], bins=3)
                    df.groupby(bins)[other_col].mean().plot(kind='bar', color='lightcoral')
                    plt.title(f"Mean {other_col} by {col}", fontsize=12)
            elif is_numeric_other:
                actual_plot_type = "Bar"
                df.groupby(col)[other_col].mean().plot(kind='bar', color='lightgreen')
                plt.title(f"Mean {other_col} by {col}", fontsize=12)
            elif is_numeric_col:
                actual_plot_type = "Bar"
                df.groupby(other_col)[col].count().plot(kind='bar', color='lightblue')
                plt.title(f"Count of {col} by {other_col}", fontsize=12)
            else:
                actual_plot_type = "Stacked Bar"
                pd.crosstab(df[col], df[other_col]).plot(kind='bar', stacked=True, colormap='Set2')
                plt.title(f"{col} vs {other_col}", fontsize=12)
            
            plt.xlabel(col, fontsize=10)
            plt.ylabel(other_col, fontsize=10)
            plt.xticks(rotation=45, ha='right', fontsize=8)
            plt.savefig(chart_path, bbox_inches='tight')
            plt.close()
            return chart_path, actual_plot_type
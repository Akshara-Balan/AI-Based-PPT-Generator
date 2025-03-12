# AI Based PPT Generator

This version is an AI-based PowerPoint presentation generator that analyzes CSV data and creates a comprehensive report with slides. The application uses various agents to load data, generate content, create plots, and assemble the final report.

This version wants the user to select a single column and it will be compared with all other columns in the dataset or csv file.


## Agents

- **DataLoaderAgent**: Loads and analyzes CSV data.
- **ContentGeneratorAgent**: Generates content based on data analysis using the LLaMA model.
- **SlideBuilderAgent**: Creates PowerPoint slides with different layouts and themes.
- **PlotGeneratorAgent**: Generates plots from the data using Matplotlib.
- **ReportAssemblerAgent**: Assembles the report and converts it to different formats.
- **UIHandlerAgent**: Handles the user interface using Streamlit.

## Installation

 Install the required dependencies:
 
    ```sh
    pip install -r requirements.txt
    ```

## Usage

1. Run the application:
    ```sh
    streamlit run code.py
    ```

2. Open the Streamlit interface in your browser and follow the instructions to upload a CSV file and customize your report.

## Example

1. Upload a CSV file.
2. Select the column to analyze and the plot type.
3. Customize the theme, font style, and minimum number of slides.
4. Generate the draft report.
5. Edit the slides if needed.
6. Finalize and export the report in the desired format (ODP, PDF, DOCX).

## Note!!
There is something wrong here, it is giving exception cases exception handling.

# agents/ui_handler.py
import streamlit as st
from .data_loader import DataLoaderAgent
from .content_generator import ContentGeneratorAgent
from .slide_builder import SlideBuilderAgent
from .plot_generator import PlotGeneratorAgent
from .report_assembler import ReportAssemblerAgent

class UIHandlerAgent:
    def run(self):
        st.title("AI Based PPT Generator")
        st.markdown("Customize your report with theme, font, and export options.")
        
        # Theme and font selection
        theme = st.selectbox("Select Theme", ["light", "dark", "blue", "green"])
        font_style = st.selectbox("Select Font Style", ["Arial", "Calibri", "Times New Roman", "Verdana"])
        
        uploaded_file = st.file_uploader("Choose a CSV file", type="csv")
        
        if uploaded_file:
            data_loader = DataLoaderAgent()
            content_gen = ContentGeneratorAgent()
            slide_builder = SlideBuilderAgent()
            plot_gen = PlotGeneratorAgent()
            report_assembler = ReportAssemblerAgent()
            
            success, message = data_loader.load_data(uploaded_file)
            if not success:
                st.error(message)
                return
            
            st.success(message)
            st.write("Detected Data Types:", data_loader.data_types)
            col = st.selectbox("Select Column to Analyze", data_loader.df.columns)
            plot_type = st.selectbox("Select Plot Type", ["Scatter", "Hexbin", "Box", "Bar"])
            min_slides = st.number_input("Minimum Number of Slides", min_value=3, value=5, step=1)
            user_prompt = st.text_area("Optional: Customize PPT (e.g., 'add summary slide')", 
                                       "Default analysis of one column vs others", height=100)
            
            if st.button("Generate Draft Report"):
                with st.spinner("Generating draft report with LLaMA..."):
                    success, slide_titles = report_assembler.assemble_report(
                        uploaded_file, col, plot_type, min_slides, user_prompt,
                        theme, font_style, data_loader, content_gen, slide_builder, plot_gen
                    )
                    if success:
                        st.success("Draft report generated!")
                        st.session_state['slide_titles'] = slide_titles
                        st.session_state['draft_generated'] = True
                    else:
                        st.error(f"Error: {slide_titles}")
                        return
            
            if st.session_state.get('draft_generated', False):
                st.subheader("Edit Slides")
                edited_slides = {}
                for title in st.session_state['slide_titles']:
                    edited_content = st.text_area(f"Edit {title}", value="\n".join([f"• Point {i+1}" for i in range(5)]), height=150)
                    edited_slides[title] = edited_content.split('\n')
                
                export_format = st.selectbox("Select Export Format", ["odp", "pdf", "docx"])
                if st.button("Finalize and Export Report"):
                    with st.spinner(f"Exporting report as {export_format}..."):
                        success, result = report_assembler.finalize_report(export_format)
                        if success:
                            st.success("Report exported successfully!")
                            mime_types = {"odp": "application/vnd.oasis.opendocument.presentation", "pdf": "application/pdf", "docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document"}
                            st.download_button(
                                label=f"Download {export_format.upper()} Report",
                                data=result,
                                file_name=f"one_column_eda_report.{export_format}",
                                mime=mime_types[export_format]
                            )
                        else:
                            st.error(f"Error: {result}")
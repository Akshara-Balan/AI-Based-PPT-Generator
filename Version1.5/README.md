# AI-Based PPT Generator V1.5

## Overview
The **AI-Based PPT Generator** is a Streamlit application that automates the process of creating PowerPoint presentations from user input. The application uses llama3.2 to generate slides based on given csv file.

## Features
- **Generate PPTs from csv file:** Users can upload csv file, and the tool will generate corresponding PowerPoint slides.
- **AI-Powered Content Generation:** Uses llama3.2 to generate slide content.
- **Customizable Slide Layouts:** Includes title slides, bullet points, and content slides.
- **Download PPTs:** Users can download the generated PowerPoint presentation.
- **User-Friendly Interface:** Built using Streamlit for an intuitive and seamless experience.

## Installation
### Prerequisites
- Python 3.8+
- pip (Python package manager)


### Steps to Generate PPT
1. Open the web app in your browser.
2. Enter the csv file, select a column and give suggestions as prompt.
3. Click on the "Generate PPT" button.
4. Download the generated PowerPoint file.

## Dependencies
The project uses the following Python libraries:
- `streamlit` - For building the web interface
- `python-pptx` - For creating PowerPoint presentations
- `Pandas` - For data handling

### This works well except for the conclusion slide.
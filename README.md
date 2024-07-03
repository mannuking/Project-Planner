# Chatbot Project Plan Creator & Analyzer

## Description

This Streamlit application provides a chatbot interface for creating, analyzing, and modifying project plans. It includes features like:

- **Project Plan Creation:** Guide users through creating a new project plan with milestones.
- **Project Plan Analysis:** Analyze existing project plans, identify potential delays, and visualize progress.
- **Project Plan Modification:** Modify existing project plans by adding, editing, or deleting milestones.
- **Dependency Management:** Add dependencies between projects. 
- **Interactive Visualizations:** Visualize project timelines and milestone progress using Plotly charts.

## How to Use

1. **Installation:**
   - Make sure you have Python 3.7 or higher installed.
   - Install the required libraries: `pip install -r requirements.txt`
2. **Running the App:**
   - Open a terminal or command prompt.
   - Navigate to the project directory.
   - Run the app using Streamlit: `streamlit run app.py`
   - The app will open in your web browser.

## App Features

### 1. Create a New Project Plan
   - The chatbot will guide you through entering:
     - Project Name
     - Description
     - Stakeholders
     - Target End Date
     - Milestones (Name, Start Date, End Date, Milestone Owner, Progress) 

### 2. Analyze an Existing Plan
   - Upload an existing project plan in .xlsx format.
   - The chatbot will analyze the plan and provide:
     - Project Duration
     - Milestone Status (On Track, Behind Schedule, Completed)
     - Potential Delay Warnings
     - Visualizations (Gantt chart, progress bars)

### 3. Modify an Existing Plan
   - Upload an existing project plan in .xlsx format.
   - You can:
     - Add New Milestones
     - Modify Existing Milestones 
     - Delete Milestones
     - Add Dependencies on other projects

### 4. Provide Feedback
   - After creating, analyzing, or modifying a plan, you can provide feedback on your experience with the chatbot. 

## Project Structure
├── chatbot.py # Main Streamlit application code
└── requirements.txt # List of required Python libraries

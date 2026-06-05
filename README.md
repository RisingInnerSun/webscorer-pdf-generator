Webscorer PDF Generator
A Streamlit application for converting Webscorer Excel race results into professionally formatted PDF reports.
The application extracts race result sections from a Webscorer Excel export and generates a PDF report with:
•	Race name and date header
•	Sri Chinmoy Races branding
•	Automatically formatted result tables
•	Consistent pagination and page numbering
•	Overall, gender and age-category results
•	Exclusion of DNS (Did Not Start) and DNF (Did Not Finish) competitors
•	Automatic sorting of race sections by distance and category
________________________________________
Features
•	Simple browser-based interface using Streamlit
•	Upload Webscorer Excel files directly
•	Generate PDF reports with a single click
•	Download the completed PDF immediately
•	No manual editing of results required
________________________________________
Installation
1. Clone the repository
git clone https://github.com/<your-username>/webscorer-pdf-generator.git
cd webscorer-pdf-generator
2. Install required packages
pip install -r requirements.txt
________________________________________
Running the Application
Start the Streamlit application:
streamlit run app.py
Your browser should automatically open to:
http://localhost:8501
If it does not, copy and paste the address into your browser.
________________________________________
Usage
1.	Open the application.
2.	Upload a Webscorer Excel results file (.xlsx).
3.	Enter the race name and race date (if required).
4.	Click Generate PDF.
5.	Download the generated PDF.
________________________________________
Project Structure
webscorer-pdf-generator
│
├── app.py
├── race_results_to_pdf.py
├── sri_chinmoy_races_logo.png
├── requirements.txt
└── README.md
Files
File	Description
app.py	Streamlit user interface
race_results_to_pdf.py	PDF generation logic
sri_chinmoy_races_logo.png	Header logo
requirements.txt	Python dependencies
________________________________________
Dependencies
•	Streamlit
•	Pandas
•	OpenPyXL
•	ReportLab
________________________________________
Example Workflow
Webscorer Excel Export
            ↓
      Upload File
            ↓
      Generate PDF
            ↓
      Download Report
________________________________________
License
This project is intended for internal and community race administration purposes.
________________________________________
Author
Balarka Robinson


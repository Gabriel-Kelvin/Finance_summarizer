# FinanceSummarizer AI  

## Overview  
FinanceSummarizer AI is a generative AI-powered application that processes annual financial reports in PDF format and generates concise, accurate, and professional summaries. The application uses Azure OpenAI and Azure AI Document Intelligence services to extract and analyze financial data, enabling users to gain valuable insights at a glance.  

## Features  
- **Performance Summary Generation**: Creates a high-level overview of a company's financial performance for a given fiscal year.  
- **Financial Metrics Extraction**: Calculates key financial metrics, including:  
  - Current Ratio  
  - Debt-to-Equity Ratio  
  - Return on Equity (ROE)  
  - Return on Assets (ROA)  
  - Gross Profit Margin  
  - Net Profit Margin  
  - Earnings Per Share (EPS)  
- **Risk Factor Analysis**: Summarizes key risk factors mentioned in the financial report.  
- **Output Formats**:  
  - **Excel File**: Tabular format with calculated financial metrics for all years available in the input report.  
  - **Word Document**: Comprehensive performance and risk factor summary.  

## Technology Stack  
- **Azure OpenAI**: For generating concise summaries and risk factor analysis.  
- **Azure AI Document Intelligence**: For extracting and processing financial data from PDF reports.  
- **Python**: For data extraction, transformation, and report generation.  
- **Pandas**: For data manipulation and calculations.  
- **OpenPyXL**: For creating Excel reports.  
- **Docx**: For generating Word documents.  

## Use Case  
This project was developed as part of a hackathon challenge with the following objectives:  
- Process and analyze financial data from Amazon's 2023 Annual Financial Report (PDF format).  
- Deliver a one-page performance summary, financial metrics table, and risk factor analysis within 3 hours.  

## Installation  
1. Clone the repository:  
   ```bash  
   git clone https://github.com/Gabriel-Kelvin/FinanceSummarizer-AI.git  
   cd FinanceSummarizer-AI  
   ```  
2. Create and activate a Python virtual environment:  
   ```bash  
   python -m venv venv  
   source venv/bin/activate  # On Windows: venv\Scripts\activate  
   ```  
3. Install the required dependencies:  
   ```bash  
   pip install -r requirements.txt  
   ```  

# Zendesk RMA Manager

![Python](https://img.shields.io/badge/Python-3.8%2B-blue.svg)
![Contributions Welcome](https://img.shields.io/badge/Contributions-Welcome-brightgreen.svg)
![Zendesk API](https://img.shields.io/badge/Zendesk-API-orange)
![Expandable ERP](https://img.shields.io/badge/Expandable-ERP-yellow)

## 📜 Description
The **Zendesk RMA Manager** is a Python-based application designed to automate the process of Return Merchandise Authorization (RMA) management. It integrates the **Zendesk API**, the **Expandable ERP** system, and a specified **Excel-based system** to fetch customer data, manage RMA requests, and generate shipping documents.

> **Note:** This software is a custom solution developed to address inefficiencies and waste caused by manual RMA processes. It is not meant to be downloaded publicly as it was tailored for a specific workflow.

This system simplifies RMA creation, automates customer responses, and outputs RMA and shipping information in PDF format, helping organizations streamline their RMA process efficiently.

## ✨ Key Features
- **Customer Data Management**: Retrieve and manage customer details via the Zendesk API and update them automatically in the RMA form.
- **ERP Integration (Expandable)**: Seamlessly integrates with the Expandable ERP system to automate RMA creation based on customer and product information.
- **RMA Creation**: Effortlessly create and store RMA details in an Excel sheet, increment RMA numbers, and update customer information.
- **Automated Responses**: Automatically generate Cadex responses based on service type and customer data.
- **PDF Output**: Generate shipping documents and commercial invoices in PDF format to prevent unwanted customs charges.

## 🏗 Project Structure
- **Customer Class**: Handles customer details such as name, email, company information, and address.
- **ZendeskClient Class**: Manages the interaction with the Zendesk API for fetching user and organization details.
- **RMAmanagement Class**: Updates and manages RMA information in an Excel workbook and generates shipping reports.
- **ERP Integration**: Interfaces with the Expandable ERP system for automating RMA creation and customer interactions.
- **Timestamp Class**: Manages timestamps for RMA creation and date formatting for Excel sheets.

## 🔄 Workflow Overview
1. **🔍 Search Customer**: Retrieves customer data using an email query through the Zendesk API.
2. **⚙️ Automate RMA Creation**: Based on product details and customer input, the system automatically generates the RMA.
3. **📝 Create RMA**: Input product details, customer complaints, and Cadex replies. Parts of the RMA form are auto-completed using dropdown menu selections.
4. **📄 Generate PDF Documents**: Automatically generate RMA and shipping documents in PDF format to streamline shipping processes.

## 📂 Excel File Availability
The **Excel file** used for managing the RMA data and workflow is only available upon request. Since this was a custom-built solution to solve a specific problem, the software and its components are not publicly distributed. The solution was developed to address inefficiencies and waste due to excessive manual work in the RMA process.

## 🛠 Prerequisites
- ![Python](https://img.shields.io/badge/Python-3.8%2B-blue.svg) **Python 3.8+**
- ![Zendesk API](https://img.shields.io/badge/Zendesk-API-orange) **Zendesk API Access**
- ![Expandable ERP](https://img.shields.io/badge/Expandable-ERP-yellow) **Expandable ERP API Access**
- **OpenPyXL** for Excel management
- **Requests** for making HTTP requests
- **MS SQL Server** (Expandable for ERP integration)

## ⚙️ Installation
1. Clone the repository:
    ```bash
    git clone https://github.com/yourusername/zendesk-rma-manager.git
    ```
2. Install dependencies:
    ```bash
    pip install -r requirements.txt
    ```

3. Configure your environment:
   - Set up your Zendesk subdomain, email, API password, and ERP (Expandable) API credentials in the code.
   - Update the file path for your Excel workbook containing customer and RMA data.

## 🚀 Usage
1. **Run the script**:
    ```bash
    python rma_manager.py
    ```
2. Follow the prompts to fetch customer data, automate RMA creation, and generate PDFs.
3. RMA details are stored in an Excel workbook, and shipping documents are generated in PDF format.

## 🔮 Future Enhancements
- **Web Interface**: A future release will include a web interface for dynamic customer search and RMA form completion, similar to Zendesk's search bar using WebSockets.
- **Enhanced ERP Integration**: Deeper integration with the Expandable ERP system for managing orders, inventory, and customer interactions.
- **Database Integration**: Transition from Excel to a more robust database for storing RMA and customer information.

## 🤝 Contributing
We welcome contributions from the community! Please fork the repository and submit a pull request.

---

### ✉️ Contact
For any questions or to request access to the Excel file, feel free to reach out at **your-email@example.com**.


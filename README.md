# Automated Document Generation and Email System

This Python automation system streamlines the process of creating and distributing personalized documents. It reads recipient data from an Excel file, generates customized documents using templates, converts them to PDF format, and automatically emails them to recipients. Perfect for batch processing tasks like certificates, personalized letters, reports, or any document that needs to be customized and distributed to multiple recipients.

## Features

- 📊 Excel Integration
  - Reads recipient data from Excel spreadsheet
  - Supports multiple recipients in batch processing
  - Flexible data structure for various use cases

- 📝 Document Generation
  - Creates personalized DOCX documents using templates
  - Supports dynamic content insertion
  - Maintains consistent formatting across all documents

- 📄 PDF Conversion
  - Automatically converts DOCX to PDF format
  - Maintains document formatting and layout
  - Creates professional-ready documents

- 📧 Automated Email Distribution
  - Sends personalized emails to each recipient
  - Includes PDF attachments automatically
  - Supports custom email messages with recipient's name
  
- 🔒 Security Features
  - Secure email configuration using environment variables
  - Support for Gmail App Passwords
  - Protected sensitive information management

## Requirements

- Python 3.x
- Required Python packages:
  ```sh
  pandas==2.1.0          # Data handling from Excel
  python-docx-template   # Document template processing
  docx2pdf              # PDF conversion
  python-dotenv         # Environment variable management
  ```

## Project Structure

```
Writing-From-a-Dataset/
│
├── script.py           # Main automation script
├── student.docx        # Document template for personalization
├── Sample.xlsx         # Data source with recipient information
├── .env               # Email configuration (not in version control)
├── .gitignore         # Git ignore configuration
│
└── Files/             # Generated output directory
    ├── docx/         # Generated DOCX documents
    └── pdfs/         # Converted PDF files
```

## Data File Format

Your `Sample.xlsx` should include these columns:
- `Name`: Recipient's full name
- `Email`: Recipient's email address

Additional columns can be added and referenced in the template using `{{ column_name }}`.

## Installation & Setup

1. **Clone the Repository**
   ```sh
   git clone https://github.com/Krushnakant-08/Writing-From-a-Dataset.git
   cd Writing-From-a-Dataset
   ```

2. **Set Up Virtual Environment (Recommended)**
   ```sh
   python -m venv venv
   # On Windows
   .\venv\Scripts\activate
   # On Unix or MacOS
   source venv/bin/activate
   ```

3. **Install Dependencies**
   ```sh
   pip install pandas python-docx-template docx2pdf python-dotenv
   ```

4. **Configure Email Settings**
   Create a `.env` file in the project root:
   ```env
   SENDER_EMAIL=your_email@example.com
   APP_PASSWORD=your_app_password
   SMTP_SERVER=smtp.gmail.com
   ```
   For Gmail users:
   1. Enable 2-Factor Authentication
   2. Generate an App Password
   3. Use the App Password in the `.env` file

5. **Create Output Directories**
   ```sh
   mkdir -p Files/docx Files/pdfs
   ```

## Usage Guide

1. **Prepare Your Data**
   - Populate `Sample.xlsx` with recipient information
   - Required columns:
     - `Name`: Full name of the recipient
     - `Email`: Valid email address
   - Add any additional columns needed for your template

2. **Customize Document Template**
   - Edit `student.docx` with your desired content and formatting
   - Use template variables in double curly braces: `{{ Name }}`
   - Available variables:
     - `{{ Name }}`: Recipient's name
     - `{{ Email }}`: Recipient's email
     - Add more by including corresponding columns in Excel

3. **Run the System**
   ```sh
   # Activate virtual environment if used
   source venv/bin/activate  # Unix/MacOS
   .\venv\Scripts\activate   # Windows
   
   # Run the script
   python script.py
   ```

4. **Monitor Progress**
   The script provides real-time feedback:
   - Document generation status
   - PDF conversion confirmation
   - Email delivery status
   - Any errors or issues encountered

5. **Check Outputs**
   - Generated DOCX: `Files/docx/<Name>.docx`
   - Generated PDF: `Files/pdfs/<Name>.pdf`
   - Review console output for success/failure logs

## Security Best Practices

- 🔒 Environment Variables
  - Never commit `.env` file to version control
  - Keep email credentials secure
  - Use App Passwords for additional security

- 📁 File Management
  - Keep sensitive data out of version control
  - Regularly clean up generated files
  - Back up templates and data files

## Error Handling

The system includes comprehensive error handling for:
- ✅ File operations (read/write)
- ✅ Template processing
- ✅ PDF conversion
- ✅ Email delivery
- ✅ Data validation

Errors are:
- Logged to console
- Non-blocking (script continues processing)
- Informative (clear error messages)

## Troubleshooting

Common issues and solutions:
1. **Email Sending Fails**
   - Check internet connection
   - Verify email credentials in `.env`
   - Ensure correct SMTP settings

2. **PDF Conversion Issues**
   - Verify Microsoft Word is installed
   - Check file permissions
   - Ensure enough disk space

3. **Template Processing Errors**
   - Verify variable names match Excel columns
   - Check template file format
   - Ensure proper template syntax

## Contributing

1. Fork the repository
2. Create your feature branch (`git checkout -b feature/AmazingFeature`)
3. Commit your changes (`git commit -m 'Add some AmazingFeature'`)
4. Push to the branch (`git push origin feature/AmazingFeature`)
5. Open a Pull Request

## License

This project is licensed under the MIT License - see the LICENSE file for details.

## Acknowledgments

- Thanks to all contributors
- Built with Python and open-source libraries
- Inspired by the need for efficient document automation

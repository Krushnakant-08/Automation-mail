# Document Automation System

This Python-based automation system generates personalized documents from Excel data, converts them to PDF, and automatically emails them to recipients. Perfect for batch processing of personalized documents like certificates, letters, or reports.

## Features

- 📊 Reads data from Excel spreadsheet
- 📝 Generates personalized DOCX documents using templates
- 📄 Automatically converts DOCX to PDF
- 📧 Sends automated emails with PDF attachments
- 🔒 Secure email configuration using environment variables
- ⚡ Batch processing for multiple recipients

## Prerequisites

Before running the script, you need:

- Python 3.x
- Required Python packages:
  ```sh
  pip install pandas python-docx-template docx2pdf python-dotenv
  ```

## Project Structure

```
Writing-From-a-Dataset/
│
├── script.py              # Main automation script
├── student.docx          # Document template
├── Sample.xlsx           # Data source
├── .env                  # Email configuration
│
└── Files/                # Generated files
    ├── docx/            # DOCX outputs
    └── pdfs/            # PDF outputs
```

## Setup

1. **Clone the Repository**
   ```sh
   git clone https://github.com/Krushnakant-08/Writing-From-a-Dataset.git
   cd Writing-From-a-Dataset
   ```

2. **Install Dependencies**
   ```sh
   pip install pandas python-docx-template docx2pdf python-dotenv
   ```

3. **Configure Email Settings**
   Create a `.env` file with:
   ```
   SENDER_EMAIL=your_email@example.com
   APP_PASSWORD=your_app_password
   SMTP_SERVER=smtp.gmail.com
   ```
   Note: For Gmail, use an App Password instead of your regular password.

4. **Prepare Directories**
   ```sh
   mkdir -p Files/docx Files/pdfs
   ```

## Usage

1. **Prepare Your Data**
   - Update `Sample.xlsx` with recipient information
   - Required columns:
     - `Name`: Recipient's name
     - `Email`: Recipient's email address

2. **Customize Template**
   - Modify `student.docx` with your desired layout
   - Use placeholders like `{{ Name }}` for dynamic content
   - Placeholders must match Excel column names

3. **Run the Automation**
   ```sh
   python script.py
   ```

4. **Check Results**
   - DOCX files appear in `Files/docx/`
   - PDF files appear in `Files/pdfs/`
   - Console shows email delivery status

## Security

- ✅ Never commit `.env` file
- ✅ Use App Passwords for email services
- ✅ Sensitive files ignored in `.gitignore`

## Error Handling

The script includes robust error handling for:
- File operations
- PDF conversion
- Email sending
- Data processing

Failed operations are logged but won't stop the batch process.

## Contributing

1. Fork the repository
2. Create your feature branch
3. Commit your changes
4. Push to the branch
5. Create a Pull Request

## License

This project is open source and available under the MIT License.

# ANANTYA Participation Certificate Automation

This script automates the generation and distribution of participation certificates for ANANTYA events. It creates personalized PowerPoint certificates, converts them to PDF, and emails them to participants using Mailjet API.

## 📋 Prerequisites

- Python 3.8 or higher
- Microsoft PowerPoint (required for PDF conversion)
- Mailjet account with verified sender email
- Windows OS (for PowerPoint COM automation)

## ⚠️ Important: Daily Email Limit

**Mailjet allows sending only 200 emails per day to comply with best practices and avoid spam filters.**

If you have more than 200 participants:
1. The script will send emails to the first 200 participants
2. After completion, **manually remove the 200 participants who received emails** from `Sample.xlsx`
3. Save the updated `Sample.xlsx` with only the remaining participants
4. Run the script again the next day to send emails to the next batch
5. Repeat this process until all participants receive their certificates

## 🚀 Getting Started for Event Coordinators

### Step 1: Prepare Your Event Data

1. Open the main **`anantya.xlsx`** file
2. Filter the data to show only your event participants:
   - Click on the **Event** column header
   - Apply filter to show only your event name
3. Select and copy all filtered rows (including the header row with Name, Email, Event columns)
4. Create a new Excel file named **`Sample.xlsx`** in the project folder
5. Paste your event data into `Sample.xlsx`
6. Save and close the file

**Important:** Make sure your `Sample.xlsx` has these columns:
- `Name` - Participant's full name
- `Email` - Participant's email address
- `Event` - Your event name

### Step 2: Certificate Template

The certificate template **`student.pptx`** is already provided in the repository and ready to use.

The template contains placeholders that will be automatically replaced:
- `{{Name}}` - Replaced with participant's name
- `{{Event}}` - Replaced with event name

**No customization needed** - just ensure the file exists in the project folder.

### Step 3: Set Up Mailjet API

1. Go to [Mailjet](https://www.mailjet.com/) and create an account (if you don't have one)
2. On the dashboard, navigate to **API** → **API Key Management**
3. Copy your **API Key** (Public) and **Secret Key** (Private) 

   **More info given in .env.local**

### Step 4: Configure Environment Variables

1. Locate the **`.env.example`** file in the project folder
2. Copy it and rename the copy to **`.env`**
3. Open **`.env`** file in a text editor
4. Fill in your credentials:
   ```
   MJ_APIKEY_PUBLIC=your_actual_public_key_here
   MJ_APIKEY_PRIVATE=your_actual_private_key_here
   SENDER_EMAIL=your-verified-email@yourdomain.com
   SENDER_NAME=Team Anantya
   ```
5. Save and close the file

**Important:** Never share your `.env` file with anyone!

### Step 5: Install Dependencies

1. Open PowerShell or Command Prompt in the project folder
2. Run the following command:
   ```powershell
   pip install pandas openpyxl python-pptx pywin32 mailjet_rest python-dotenv
   ```
3. Wait for installation to complete

### Step 6: Run the Script

1. Ensure Microsoft PowerPoint is installed and not running
2. Double-check that:
   - `Sample.xlsx` has **only** your event data
   - `certificate.pptx` template exists (already provided in repo)
   - `.env` file is configured
3. Open PowerShell in the project folder
4. Run the script:
   ```powershell
   python script.py
   ```

### Step 7: Handling Large Participant Lists (More than 200)

**Important:** Due to the 200 email per day limit:

1. **After the script completes** and sends emails to 200 participants:
   - Open `Sample.xlsx`
   - **Manually delete the rows of all participants who received emails** (first 200 rows)
   - Save the file with only the remaining participants
   
2. **The next day:**
   - Run the script again: `python script.py`
   - It will send emails to the next 200 participants
   
3. **Repeat this process** until all participants receive their certificates

**Tip:** Keep a backup of your original `Sample.xlsx` before making changes, or maintain a separate tracking sheet to know which participants have received emails.

### What Happens When You Run the Script

The script will:
1. ✅ Read participant data from `Sample.xlsx`
2. ✅ Create personalized PPTX certificates in `Files/pptx/` folder
3. ✅ Convert each PPTX to PDF in `Files/pdfs/` folder
4. ✅ Send personalized emails with PDF certificates to each participant (up to 200 per day)
5. ✅ Display progress messages for each step

**Note:** If you have more than 200 participants, the script will process all certificates but you should send emails in batches of 200 per day. After sending 200 emails, remove those participants from `Sample.xlsx` and run the script again the next day.

### Expected Output

```
Saved: ./Files/pptx/John_Doe_HackathonEvent.pptx
Exported to PDF: ./Files/pdfs/John_Doe_HackathonEvent.pdf
-> Successfully sent email with PDF attachment to: john@example.com

Saved: ./Files/pptx/Jane_Smith_HackathonEvent.pptx
Exported to PDF: ./Files/pdfs/Jane_Smith_HackathonEvent.pdf
-> Successfully sent email with PDF attachment to: jane@example.com

All documents have been processed and emails sent.
```

## 📁 Project Structure

```
Writing-From-a-Dataset/
│
├── script.py              # Main automation script
├── Sample.xlsx            # Your event participant data (you create this)
├── student.pptx          # Certificate template (provided)
├── .env                  # Your credentials (you create from .env.example)
├── .env.example          # Template for environment variables
├── README.md             # This file
│
└── Files/
    ├── pptx/             # Generated PPTX certificates
    └── pdfs/             # Generated PDF certificates
```

## ⚠️ Troubleshooting

### "FileNotFoundError: Sample.xlsx"
- Make sure you created `Sample.xlsx` in the project root folder
- Check the file name spelling (case-sensitive)

### "FileNotFoundError: certificate.pptx"
- The template should already be in the repository
- If missing, contact **Krushnakant Patil** at [krushnakant.patil24@pccoepune.org](mailto:krushnakant.patil24@pccoepune.org) to get the template file
- Ensure it's placed in the project root folder

### Email not sending / Mailjet errors
- Verify your sender email is validated in Mailjet dashboard
- Check your API keys are correct in `.env` file
- Ensure you have sufficient Mailjet email credits

### PowerPoint errors
- Make sure Microsoft PowerPoint is installed
- Close PowerPoint before running the script
- Try running the script with administrator privileges

### "Module not found" errors
- Run: `pip install pandas openpyxl python-pptx pywin32 mailjet_rest python-dotenv`

## 📧 Email Template

Participants will receive an email with:
- **Subject:** Participation Certificate - [Event Name] | ANANTYA
- **Content:** Professional message congratulating them
- **Attachment:** PDF certificate

**Daily Limit:** Maximum 200 emails per day

## 📊 Managing Large Participant Lists

For events with more than 200 participants:

1. **Day 1:**
   - Run the script with all participants in `Sample.xlsx`
   - Script sends emails to first 200 participants
   - Open `Sample.xlsx` and delete the 200 rows of participants who received emails
   - Save the file

2. **Day 2:**
   - Run the script again with the updated `Sample.xlsx`
   - Script sends emails to the next 200 participants
   - Again, remove those 200 participants from `Sample.xlsx`

3. **Continue** this process until all participants are notified

**Best Practice:** Maintain a master copy of your complete participant list separately to track who has received certificates.

## 🔒 Security Notes

- Never commit your `.env` file to version control
- Keep your Mailjet API keys confidential
- The `.env` file should be added to `.gitignore`

## 📞 Support

If you encounter any issues:
1. Check the troubleshooting section above
2. Verify all prerequisites are met
3. Contact **Krushnakant Patil** with error messages:
   - 📧 Email: [krushnakant.patil24@pccoepune.org](mailto:krushnakant.patil24@pccoepune.org)

---

**Made for ANANTYA Event Coordinators** 🎉
# ERF Email Automation System

A Python-based automation system for sending ERF (Equipment Request Form) status updates via Outlook.

## Features

- ✅ Process Excel files with ERF data
- ✅ Filter items by status ('On order', 'Received')
- ✅ Group items by requester to send consolidated emails
- ✅ Generate formatted status update emails
- ✅ Integrate with corporate Outlook
- ✅ Test mode for safe preview
- ✅ Comprehensive logging
- ✅ Modular architecture

## Quick Start

1. **Install dependencies:**

   ```bash
   pip install -r requirements.txt
   ```

2. **Run the application:**

   ```bash
   python main.py path/to/your/excel/file.xlsx
   ```

3. **Follow the prompts:**
   - The system will process your Excel file
   - Preview emails in test mode first
   - Confirm to send actual emails

## Project Structure

- `config/` - Configuration settings
- `src/data/` - Excel processing logic
- `src/email/` - Email templates and Outlook service
- `src/services/` - Main automation orchestration
- `src/utils/` - Logging and validation utilities
- `tests/` - Unit tests
- `logs/` - Application logs

## Configuration

Edit `config/settings.py` to customize:

- Required Excel columns
- Target statuses
- Email templates
- File paths

## Future Enhancements

- SharePoint live data integration
- Web UI interface
- Email scheduling
- Advanced reporting

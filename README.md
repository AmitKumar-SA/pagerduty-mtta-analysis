# PagerDuty MTTA Analysis

A Python script that fetches Mean Time to Acknowledge (MTTA) metrics from PagerDuty's Analytics API and updates an Excel spreadsheet with the data.

## Overview

This tool automates the process of collecting PagerDuty escalation policy metrics and populating them into an Excel file for analysis and reporting. It retrieves the mean time to first acknowledgment for incidents within specified date ranges and converts the data from seconds to minutes.

## Features

- 🔄 **Automated Data Collection**: Fetches MTTA metrics from PagerDuty Analytics API
- 📊 **Excel Integration**: Updates existing Excel files with numerical data (formatted to 2 decimal places)
- 🛡️ **Robust Error Handling**: Includes retry logic, rate limiting, and connection error handling
- 🧪 **Mock Mode**: Test functionality without making real API calls
- ⚡ **Selective Processing**: Process specific rows or date ranges
- 🔒 **Secure Authentication**: Uses environment variables for API tokens
- 📅 **Flexible Date Ranges**: Support for any month/year combination

## Prerequisites

- Python 3.6+
- Required Python packages:
  ```bash
  pip install openpyxl requests
  ```
- PagerDuty API token with analytics access
- Excel file with proper structure (see [File Structure](#file-structure))

## Installation

1. Clone or download this repository
2. Install required dependencies:
   ```bash
   pip install -r requirements.txt
   ```
3. Set up your PagerDuty API token as an environment variable:
   ```bash
   export PAGERDUTY_API_TOKEN="your_pagerduty_api_token_here"
   ```

## File Structure

The script expects an Excel file at `src/MTTA_calc.xlsx` with the following structure:

| Column | Description |
|--------|-------------|
| First column | Escalation policy name |
| `id` column | PagerDuty escalation policy ID |
| Month columns | `Jan`, `Feb`, `Mar`, etc. (where data will be populated) |

## Usage

### Basic Usage

```bash
# Update January data for all escalation policies
python src/update_pagerduty_metrics.py

# Update specific month
python src/update_pagerduty_metrics.py --month Mar
```

### Command Line Options

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| `--mock` | flag | False | Use mock data instead of real API calls |
| `--month` | string | "Jan" | Month column to update (Jan-Dec) |
| `--force` | flag | False | Force update even if values already exist |
| `--year` | integer | 2025 | Year to get data for |
| `--delay` | float | 2.5 | Delay between API requests (seconds) |
| `--start-row` | integer | None | Start processing from this row number |
| `--end-row` | integer | None | Stop processing at this row number |
| `--no-verify-ssl` | flag | False | Disable SSL verification (testing only) |

### Example Commands

```bash
# Test with mock data
python src/update_pagerduty_metrics.py --mock --month Feb

# Force update March data even if values exist
python src/update_pagerduty_metrics.py --month Mar --force

# Process only rows 5-10 with custom delay
python src/update_pagerduty_metrics.py --start-row 5 --end-row 10 --delay 3.0

# Get data for 2024 instead of default 2025
python src/update_pagerduty_metrics.py --month Jun --year 2024

# Increase delay to avoid rate limiting
python src/update_pagerduty_metrics.py --delay 5.0
```

## Configuration

### Environment Variables

| Variable | Required | Description |
|----------|----------|-------------|
| `PAGERDUTY_API_TOKEN` | Yes* | Your PagerDuty API token (*Not required in mock mode) |

### Getting a PagerDuty API Token

1. Log in to your PagerDuty account
2. Go to **Configuration** → **API Access**
3. Click **Create New API Key**
4. Give it a description and ensure it has the necessary permissions
5. Copy the token and set it as an environment variable

## Output

The script will:

1. **Read** the Excel file and identify escalation policies
2. **Fetch** MTTA data from PagerDuty for each policy
3. **Convert** response times from seconds to minutes (2 decimal places)
4. **Update** the specified month column in Excel
5. **Save** the updated file

### Sample Output

```
Loading Excel file from: /path/to/src/MTTA_calc.xlsx
Using PagerDuty API token (first 4 chars: abcd...)
Processing rows 2 to 15 (total: 14 rows)
Processing: Production Alerts (ID: P123ABC)
  Making API request to https://api.pagerduty.com/analytics/metrics/incidents/all
  Found mean_seconds_to_first_ack: 450
  Updated Jan column: 7.50 minutes
Excel file updated successfully: /path/to/src/MTTA_calc.xlsx
```

## Error Handling

The script includes comprehensive error handling:

- **Rate Limiting**: Automatically retries with exponential backoff
- **Connection Errors**: Retries with jitter to avoid synchronized requests
- **Timeouts**: Configurable timeout with retry logic
- **Data Validation**: Handles missing or malformed API responses
- **File Errors**: Validates Excel file existence and structure

## Testing

Use mock mode for testing without making API calls:

```bash
python src/update_pagerduty_metrics.py --mock
```

Mock mode generates random but realistic MTTA values between 2 minutes and 2 hours.

## Troubleshooting

### Common Issues

1. **"PAGERDUTY_API_TOKEN environment variable not set"**
   - Solution: Set the environment variable with your API token

2. **"Excel file not found"**
   - Solution: Ensure `src/MTTA_calc.xlsx` exists in the correct location

3. **"Column 'Month' not found"**
   - Solution: Verify the Excel file has the correct month column headers

4. **Rate limiting (429 errors)**
   - Solution: Increase the `--delay` parameter (e.g., `--delay 5.0`)

5. **SSL certificate errors**
   - Solution: Use `--no-verify-ssl` flag (testing only) or fix SSL configuration

### Debug Mode

For verbose output, the script automatically prints:
- API request details
- Response processing information
- Error stack traces when issues occur

## Security Considerations

- ✅ API tokens are read from environment variables (not hardcoded)
- ✅ SSL certificate verification is enabled by default
- ✅ No sensitive data is logged to console
- ⚠️ Only disable SSL verification for testing purposes

## Contributing

1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Test with mock mode
5. Submit a pull request

## License

This tool is for internal use only. Do not distribute or share this code without proper authorization.

## Support

For issues or questions:
1. Check the [Troubleshooting](#troubleshooting) section
2. Review the console output for error details
3. Test with `--mock` mode to isolate API issues
4. Contact the development team

---

**Version**: 1.0  
**Last Updated**: August 2025  
**Author**: @Akumar6

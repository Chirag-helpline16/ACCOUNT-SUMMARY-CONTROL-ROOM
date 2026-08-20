# Account Summary Control Room

[![Tests](https://github.com/Chirag-helpline16/ACCOUNT-SUMMARY-CONTROL-ROOM/actions/workflows/tests.yml/badge.svg)](https://github.com/Chirag-helpline16/ACCOUNT-SUMMARY-CONTROL-ROOM/actions/workflows/tests.yml)
[![Code Quality](https://github.com/Chirag-helpline16/ACCOUNT-SUMMARY-CONTROL-ROOM/actions/workflows/code-quality.yml/badge.svg)](https://github.com/Chirag-helpline16/ACCOUNT-SUMMARY-CONTROL-ROOM/actions/workflows/code-quality.yml)

A comprehensive Flask-based financial transaction analysis and reporting system designed for analyzing layered financial transactions, tracking money flow, and generating detailed audit reports.

## 📋 Table of Contents

- [Features](#features)
- [Quick Start](#quick-start)
- [Installation](#installation)
- [Usage](#usage)
- [Batch Processing](#batch-processing)
- [File Structure](#file-structure)
- [Configuration](#configuration)
- [API Documentation](#api-documentation)
- [Data Formats](#data-formats)
- [Troubleshooting](#troubleshooting)
- [Contributing](#contributing)

## ✨ Features

### 1. Hierarchical Transaction Analysis
- Multi-layer transaction tracking (Layer 1 through Layer 6+)
- Account-wise summaries with credited/debited amounts
- Bank-wise grouping and analysis
- Interactive expandable/collapsible transaction trees
- Real-time transaction filtering and search

### 2. Money Flow Visualization
- Visual tree representation of transaction flows
- Color-coded account types:
  - 🔴 **Red**: Victim/Entry accounts
  - 🔵 **Blue**: Intermediate accounts
  - 🟠 **Orange**: Cash-out accounts
  - ⚫ **Gray**: Frozen accounts
- Transaction status indicators:
  - 🔴 Red: PENDING
  - 🟠 Orange: PARTIAL
  - 🟢 Green: COMPLETE
- Interactive nodes with detailed information
- Downloadable HTML flow diagrams

### 3. Cash-Out Channel Analysis
- ATM withdrawals with location tracking
- Cheque withdrawal analysis
- POS transactions with merchant details
- Transaction holds and frozen account tracking
- Micro-transactions (< ₹500) grouping
- Automatic account number deduplication
- Disputed amount tracking

### 4. Comprehensive Reporting
- Excel-based account summary reports
- Bank-wise transaction analysis
- Flow diagram HTML exports
- Duplicate transaction detection and logging
- Audit trail with detailed logging
- Real-time dashboard integration

### 5. Batch Processing Engine
- Overnight scheduled batch processing
- Resumable multi-workbook processing
- Deduplication with audit logging
- Parallel processing support (4 processes)
- SQLite database for consolidated data
- Workbook cache management
- Direct Excel export capabilities

## 🚀 Quick Start

### Prerequisites
- Python 3.8+
- pip (Python package manager)
- Git

### Installation

1. **Clone the repository:**
```bash
git clone https://github.com/Chirag-helpline16/ACCOUNT-SUMMARY-CONTROL-ROOM.git
cd ACCOUNT-SUMMARY-CONTROL-ROOM
```

2. **Install dependencies:**
```bash
pip install -r requirements.txt
```

3. **Prepare data files:**
   - Place your Excel transaction file in the project directory
   - Ensure the file contains required sheets: Main, ATM, Cheque, POS, Hold, Others
   - Verify column formats match the expected data structure

## 💻 Usage

### Web Dashboard

1. **Start the application:**
```bash
python app_account.py
```

2. **Access the dashboard:**
```
http://localhost:5000
```

3. **Features available:**
   - Upload and analyze Excel files
   - View hierarchical transactions
   - Generate flow diagrams
   - Export reports
   - Download analysis data

### Interactive Views
- **Main View**: Hierarchical transaction analysis with layer-wise breakdown
- **Flow Diagram**: Visual representation of money flow with interactive nodes
- **Reports**: Download comprehensive analysis in various formats
- **Search**: Filter transactions by ACK, amount, date, or account number

## 📅 Batch Processing

### Overnight Batch Processing

Automate the analysis of multiple workbooks with the batch processor:

```bash
# Run the batch processor
python batch_account_summaries.py
```

Or use the provided batch file (Windows):
```batch
START_OVERNIGHT_AND_DASHBOARD.bat
```

This will:
- Process all workbooks in `C:\Users\admin\Desktop\bank_trails`
- Create a consolidated SQLite database
- Exclude duplicate credited amounts while preserving debit/recovery rows
- Generate a summary dashboard
- Export results to Excel

### Configuration

Edit `batch_account_summaries.py` to customize:
- Input directory path
- Database location
- Output file format
- Processing threads
- Deduplication rules

### Deduplication Logic

The batch processor automatically deduplicates transactions matched by:
- **Acknowledgement Number (ACK)**
- **Credited Transaction ID**
- **Last 4 digits of Credited Account**

When duplicates are found:
- First occurrence is retained
- Subsequent occurrences are logged in **Duplicate Entry Info**
- Only subsequent credited amounts are excluded from **Total Credited Amount**
- Debit totals and other-sheet/recovery matching continue to use every row
- Audit counts are recorded in SQLite

For a complete reprocessing without cache:
```batch
REPROCESS_ALL_STRICT_ONCE.bat
```

## 📁 File Structure

```
ACCOUNT-SUMMARY-CONTROL-ROOM/
├── app_account.py                        # Main Flask application
├── batch_account_summaries.py            # Batch processing engine
├── summary_database.py                   # Database management
├── templates/
│   └── account_summary_dashboard.html    # Dashboard HTML template
├── static/                               # Static assets
├── .github/
│   └── workflows/
│       ├── overnight-batch.yml           # Scheduled batch processing
│       ├── tests.yml                     # Test runner
│       └── code-quality.yml              # Linting and code quality
├── requirements.txt                      # Python dependencies
├── ACCOUNT_SUMMARY_BATCH_GUIDE.md       # Detailed batch guide
├── START_OVERNIGHT_AND_DASHBOARD.bat    # Windows batch launcher
├── REPROCESS_ALL_STRICT_ONCE.bat        # Force reprocessing script
├── .gitignore                           # Git configuration
└── README.md                            # This file
```

## ⚙️ Configuration

### Environment Variables

Create a `.env` file in the project root:
```env
FLASK_ENV=development
FLASK_DEBUG=True
DATABASE_PATH=./data/account_summary.db
INPUT_DIRECTORY=C:\Users\admin\Desktop\bank_trails
OUTPUT_DIRECTORY=./exports
MAX_PROCESSES=4
LOG_LEVEL=INFO
```

### Database Configuration

The system uses SQLite for consolidated data storage. Key tables:
- `account_summaries`: Consolidated account-wise summaries
- `deduplication_audit`: Duplicate transaction logs
- `processing_log`: Batch processing audit trail
- `workbook_cache`: Processed workbook metadata

## 📡 API Documentation

### Endpoints

#### Get Transactions by Layer
```
GET /api/transactions?layer=1&ack=ACK123
```
Returns transactions for specified layer and ACK

#### Get Flow Diagram Data
```
GET /api/flow-diagram/<transaction_id>
```
Returns hierarchical flow diagram data as JSON

#### Get Other Sheets Summary
```
GET /api/other-sheets-total?ack=ACK123
```
Returns summary of ATM, Cheque, POS, Hold, Others sheets

#### Get Sheet Details
```
GET /api/sheet-details/<sheet_name>?ack=ACK123
```
Returns detailed data from specified sheet

#### Upload File
```
POST /api/upload
Content-Type: multipart/form-data
file: <Excel file>
```
Uploads and processes Excel file

## 📊 Data Formats

### Main Transaction Sheet
Expected columns:
| Column | Name | Format | Example |
|--------|------|--------|---------|
| 0 | Serial Number | Integer | 1 |
| 1 | Acknowledgement Number (ACK) | String | ACK-2024-001 |
| 2 | Debited Account | String | ACC123456789 |
| 3 | Debited Transaction ID | String | TXN-2024-001 |
| 4 | Debited Amount | Decimal | 50000.00 |
| 5 | Layer | Integer | 1 |
| 6 | Credited Account | String | ACC987654321 |
| 7 | Credited Transaction ID | String | TXN-2024-002 |
| 8 | Credited Amount | Decimal | 50000.00 |

### Other Sheets (ATM, Cheque, POS, Hold)
- Column 6: Disputed amount
- Column 8: Location/Merchant/Details
- Column 11: Bank name
- Column 1: Account number
- Column 2: Transaction ID
- Column 4: Amount
- Column 5: Date

## 🔧 Troubleshooting

### Common Issues

**Issue**: Database is locked
```
Solution: Close all Excel files and restart the batch processor
```

**Issue**: File not found error
```
Solution: Verify input directory path in configuration
         Check that Excel file is in correct location
```

**Issue**: Duplicate credits are still counted
```
Solution: Run REPROCESS_ALL_STRICT_ONCE.bat for a full credit-only rebuild
         Check the duplicate identity helpers in app_account.py
```

**Issue**: Flask port already in use
```
Solution: Change port in app_account.py (default: 5000)
         Or kill existing process using the port
```

### Logging

Check logs in `logs/` directory:
```bash
tail -f logs/batch_processing.log
```

## 🤝 Contributing

We welcome contributions! Please follow these guidelines:

1. **Fork the repository**
2. **Create a feature branch:**
```bash
git checkout -b feature/your-feature-name
```
3. **Make your changes and commit:**
```bash
git commit -m "Add description of changes"
```
4. **Push to your fork:**
```bash
git push origin feature/your-feature-name
```
5. **Create a Pull Request**

### Code Standards
- Follow PEP 8 style guide
- Write docstrings for functions
- Add tests for new features
- Update documentation

## 📝 License

This project is licensed under the MIT License - see LICENSE file for details.

## 📞 Support

For issues and questions:
- **Issues**: [GitHub Issues](https://github.com/Chirag-helpline16/ACCOUNT-SUMMARY-CONTROL-ROOM/issues)
- **Discussions**: [GitHub Discussions](https://github.com/Chirag-helpline16/ACCOUNT-SUMMARY-CONTROL-ROOM/discussions)
- **Email**: chirag.helpline16@example.com

## 🎯 Roadmap

- [ ] Web UI improvements (responsive design)
- [ ] Advanced filtering and search
- [ ] Real-time notification system
- [ ] Performance optimizations for large datasets
- [ ] Multi-language support
- [ ] API authentication and rate limiting
- [ ] Docker containerization
- [ ] Cloud deployment options

---

**Last Updated**: August 2026
**Version**: 2.0

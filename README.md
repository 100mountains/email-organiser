# Email Organizer

A Python script to organize and process email files, handling various email formats including .eml files.

By default, the script scans the EMAIL-MAIN folder for email files to process and organize.

## Setup
1. Create a Python virtual environment:
```bash
python -m venv .venv
source .venv/bin/activate  # On Linux/Mac
# or
.\venv\Scripts\activate  # On Windows
```

2. Install requirements:
```bash
pip install -r requirements.txt
```

## Usage

```bash
python email_organizer.py
```

# Email Processing System Analysis

## Core Components

1. EmailProcessor (Main Class)
   - Manages the overall email processing workflow
   - Handles file discovery, processing, and organization
   - Creates organized directory structure based on domains and dates
   - Coordinates attachment handling

2. Data Classes
   - EmailDomains: Stores from/to/cc domains with special handling for .gov.uk
   - EmailDateInfo: Manages datetime information and formatting
   - EmailProgress: Handles progress tracking and display

3. Helper Functions
   - extract_email_details(): Parses HTML emails
   - extract_eml_details(): Parses EML files
   - copy_attachments(): Handles file attachments
   - extract_embedded_attachments(): Processes embedded EML attachments

## Key Features

1. File Organization
   - Organizes emails by domain and year
   - Creates separate copies for each .gov.uk domain involved
   - Preserves original timestamps
   - Handles both HTML and EML formats

2. Attachment Handling
   - Extracts embedded attachments from EML files
   - Processes linked attachments in HTML files
   - Handles recursive EML attachments
   - Creates detailed attachment logs

3. Progress Tracking
   - Real-time progress bar
   - Counts processed files
   - Tracks .gov.uk emails
   - Monitors attachment processing

4. Error Handling & Logging
   - Comprehensive error logging
   - Multiple encoding support
   - Fallback parsing methods
   - Safe file operations

## Security Features

1. Path Safety
   - Validates file paths
   - Sanitizes filenames
   - Prevents directory traversal
   - Handles malformed inputs

2. File Processing
   - Multiple encoding attempts
   - Validates email-like content
   - Safe attachment extraction
   - Secure file copying

## Performance Considerations

1. Efficiency
   - Incremental processing
   - Progress tracking
   - Memory-efficient file handling
   - Streaming file operations

2. Robustness
   - Multiple date parsing formats
   - Multiple file encoding support
   - Fallback parsing methods
   - Comprehensive error handling

## Areas for Potential Enhancement

1. Performance
   - Parallel processing for large datasets
   - Batch file operations
   - Caching of common operations

2. Features
   - Additional file format support
   - More detailed metadata extraction
   - Enhanced search capabilities
   - Duplicate detection

3. Monitoring
   - Enhanced progress metrics
   - Performance statistics
   - Processing summaries
   - Detailed error reporting

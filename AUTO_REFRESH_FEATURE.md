# GURI Welcome Tab Auto-Refresh Feature

## Overview

The Welcome tab now automatically refreshes every 60 seconds to keep email statistics and important emails up-to-date.

## Features

### 1. **Automatic Refresh**
- **Interval:** Every 60 seconds (1 minute)
- **What Refreshes:**
  - Email statistics (today and yesterday)
  - Important emails list
  - Last refresh timestamp

### 2. **Last Refresh Display**
- Shows the exact date and time of the last refresh
- Format: `YYYY-MM-DD HH:MM:SS`
- Located at the top of the Email Statistics section
- Example: `Last Refresh: 2025-11-04 14:35:22`

### 3. **Visual Indicators**
- **Last Refresh Time:** Shows when data was last updated
- **Auto-refresh Status:** Displays "Auto-refresh: Every 60 seconds" in blue text

## How It Works

1. **Initial Load:**
   - When you open the application, the Welcome tab loads immediately
   - Data is fetched from the database
   - First refresh timestamp is set

2. **Automatic Updates:**
   - Timer triggers every 60 seconds
   - Queries database for new email records
   - Updates all statistics and important emails
   - Updates the "Last Refresh" timestamp

3. **Clean Shutdown:**
   - Auto-refresh timer is properly cancelled when closing the application
   - No background processes continue running

## Manual Refresh

You can still manually refresh at any time using:
- **🔄 Refresh Statistics** button - Updates email statistics
- **🔄 Refresh Important Emails** button - Updates important emails list

Manual refreshes also update the "Last Refresh" timestamp.

## Benefits

- **Real-time Monitoring:** Stay updated on email activity without manual intervention
- **Always Current:** Important emails are always up-to-date
- **Timestamp Visibility:** Know exactly when data was last refreshed
- **Efficient:** Only refreshes when application is running

## Technical Details

- Uses tkinter's `after()` method for scheduling
- Non-blocking - doesn't freeze the UI
- Error handling prevents crashes if database is unavailable
- Timer is properly cleaned up on application exit

## Performance

- **Refresh Time:** Typically < 1 second for databases with thousands of records
- **Database Impact:** Minimal - only queries email records (document_type = '01')
- **Memory Usage:** Negligible additional overhead


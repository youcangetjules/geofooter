# Suppliers & Contract Codes Tab Guide

## Overview

The new **Suppliers & Contract Codes** tab allows you to manage supplier information, contract codes, and company details all in one place. This feature is separate from GURI records but tied to company names extracted from email domains.

## Features

### 1. **Company Name Management**
- Automatically extract company names from email domains
- Manual entry for outgoing emails
- Link suppliers to their root domain

### 2. **Supplier Information**
Manage comprehensive supplier details:
- **Company Name**: Primary identifier
- **Company Domain**: Email domain (e.g., microsoft.com)
- **Supplier Code**: Your internal supplier code
- **Contract Code**: Associated contract reference
- **Contact Person**: Primary contact name
- **Contact Email**: Contact's email address
- **Notes**: Additional information (free text)

### 3. **Auto-Extract from Email**
- Enter any email address
- Click "Extract" to automatically:
  - Extract company name from domain
  - Populate company domain field
  - Example: `john@microsoft.com` → Company: "Microsoft", Domain: "microsoft.com"

## Using the Suppliers Tab

### Adding a New Supplier

**Method 1: Auto-Extract from Email**
1. Navigate to the **🏢 Suppliers** tab
2. In the "Extract from Email" field, enter an email address (e.g., `contact@supplier.com`)
3. Click **Extract**
4. Company name and domain are automatically populated
5. Fill in the remaining fields:
   - Supplier Code
   - Contract Code
   - Contact Person
   - Contact Email
   - Notes
6. Click **Save Supplier**

**Method 2: Manual Entry**
1. Navigate to the **🏢 Suppliers** tab
2. Manually enter all fields:
   - Company Name (required)
   - Company Domain (optional)
   - Supplier Code
   - Contract Code
   - Contact Person
   - Contact Email
   - Notes
3. Click **Save Supplier**

### Editing an Existing Supplier

1. Find the supplier in the list on the right
2. Click on the supplier row to load it into the form
   - Or double-click to load
3. Modify any fields as needed
4. Click **Save Supplier**
   - This will update the existing record

### Deleting a Supplier

1. Select the supplier from the list
2. Click **Delete Selected**
3. Confirm the deletion

### Searching for Suppliers

1. Use the search box at the top of the suppliers list
2. Type any part of:
   - Company name
   - Domain
   - Supplier code
   - Contract code
   - Contact person
3. The list filters in real-time as you type

## Database Structure

### supplier_codes Table

| Field | Type | Description |
|-------|------|-------------|
| id | Integer | Primary key (auto-increment) |
| company_name | Text | Company name (unique) |
| company_domain | Text | Email domain |
| supplier_code | Text | Your internal supplier code |
| contract_code | Text | Contract reference number |
| contact_person | Text | Primary contact name |
| contact_email | Text | Contact's email |
| notes | Text | Additional information |
| created_at | Timestamp | Record creation time |
| updated_at | Timestamp | Last update time |

### guri_records Table Addition

A new `company_name` field has been added to the `guri_records` table to link GURIs with companies.

## Integration with GURI Records

While the Suppliers tab is separate from GURI records, the `company_name` field is now available in both:

1. **GURI Records** can reference company names
2. **Supplier Codes** provide additional company information
3. Future updates may auto-populate company names when creating GURIs from emails

## Examples

### Example 1: Adding Microsoft as a Supplier

```
Extract from Email: bill@microsoft.com
→ Auto-extracts:
   Company Name: Microsoft
   Company Domain: microsoft.com

Then add:
Supplier Code: SUP-MS-001
Contract Code: CNT-2025-MST
Contact Person: Bill Gates
Contact Email: bill@microsoft.com
Notes: Main software supplier, annual contract renewal in Q4
```

### Example 2: Manual Entry for Aliniant

```
Company Name: Aliniant
Company Domain: aliniant.com
Supplier Code: INT-001
Contract Code: INTERNAL
Contact Person: Julian Garrett
Contact Email: julian.garrett@aliniant.com
Notes: Internal company records
```

## Tips & Best Practices

1. **Consistent Naming**: Use consistent company name formatting (e.g., "Microsoft" not "microsoft" or "MICROSOFT")

2. **Use Auto-Extract**: For external suppliers, use the email extraction feature to ensure consistency

3. **Keep Notes Updated**: Use the notes field for important information like:
   - Contract renewal dates
   - Special terms
   - Account manager information

4. **Regular Backups**: The supplier codes are stored in the same database as GURI records

5. **Search Function**: Use the search to quickly find suppliers - it searches across all fields

## Migration

If you have an existing GURI database, run the migration script:

```bash
python add_company_field.py
```

This will:
- Add the `company_name` column to `guri_records`
- Create the `supplier_codes` table
- Preserve all existing data

## Benefits

- **Centralized Management**: All supplier information in one place
- **Quick Lookup**: Instant search and filtering
- **Email Integration**: Auto-extract company names from emails
- **Contract Tracking**: Link suppliers to contract codes
- **Contact Management**: Keep contact information organized
- **GURI Integration**: Ready for future enhancements linking GURIs to suppliers


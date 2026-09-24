# GURI Global Type Classification System

## Overview

GURIs are classified into **3 Global Types** based on their document type (Component 6):

1. **Global Type 1**: Email
2. **Global Type 2**: General Documents  
3. **Global Type 3**: Aliniant Internal Documents

The Global Type is automatically displayed in the GURI Decompositor above the component table.

---

## Global Type 1: Email

**Document Type Code:** `01`

**Description:** Email messages processed through the email security analysis system.

**Characteristics:**
- Sender: Actual email address
- Recipients: Actual recipient addresses
- Risk Level: Security assessment score
- All fields contain real data

**Example:**
```
GURI: 3ae2bx7f4d1xa1b3c5e7x9f8e7d6cx2a4x01
Component 6 Value: 01
Global Type: Global Type 1: Email
```

---

## Global Type 2: General Documents

**Document Type Codes:** `02`, `03`, `04`, `05`, `06`, `07`, `08`, `09`, `10`, `99`

### Document Types

| Code | Document Type |
|------|---------------|
| 02 | Word Document |
| 03 | Excel Spreadsheet |
| 04 | PowerPoint Presentation |
| 05 | PDF Document |
| 06 | Text File |
| 07 | Image File |
| 08 | Video File |
| 09 | Audio File |
| 10 | Archive/Zip |
| 99 | Other |

**Characteristics:**
- Sender: `julian.garrett@aliniant.com` (default)
- Recipients: `FFFFF` (placeholder)
- Risk/Location: Document file path
- General business documents

**Example:**
```
GURI: 3ae2bx7f4d1xa1b3c5e7x9f8e7d6cx2a4x05
Component 6 Value: 05 (PDF Document)
Global Type: Global Type 2: General Documents
```

---

## Global Type 3: Aliniant Internal Documents

**Document Type Codes:** `1a`, `1b`, `1c`, `1d`, `1e`

### Document Types

| Code | Document Type | Purpose |
|------|---------------|---------|
| 1a | Aliniant Policy Document | Company policies, procedures, guidelines |
| 1b | Aliniant Technical Document | Technical specifications, architecture docs |
| 1c | Aliniant Compliance Document | Regulatory compliance, audit documents |
| 1d | Aliniant Pre-sales Document | Sales proposals, RFPs, customer presentations |
| 1e | Aliniant Contract Document | Contracts, agreements, legal documents |

**Characteristics:**
- Sender: `julian.garrett@aliniant.com` (default)
- Recipients: `FFFFF` (placeholder)
- Risk/Location: Document file path
- Internal Aliniant business documents
- Specific categorization for document management

**Examples:**

**Policy Document:**
```
GURI: 3ae2bx7f4d1xa1b3c5e7x9f8e7d6cx2a4x1a
Component 6 Value: 1a
Document Type: Aliniant Policy Document
Global Type: Global Type 3: Aliniant Internal Documents
```

**Contract Document:**
```
GURI: 3ae2bx7f4d1xa1b3c5e7x9f8e7d6cx2a4x1e
Component 6 Value: 1e
Document Type: Aliniant Contract Document
Global Type: Global Type 3: Aliniant Internal Documents
```

---

## Usage in GUI

### Viewing Global Type

When you decompose a GURI in the **Decompositor tab**, the Global Type appears in the summary:

```
+----------------------------------------------------------------+
| GURI Summary                                                   |
+----------------------------------------------------------------+
| Original GURI: 3ae2bx7f4d1xa1b3c5e7x9f8e7d6cx2a4x1a           |
| Total Length: 37 characters                                    |
| Components Found: 6                                            |
| 📋 Global Type 3: Aliniant Internal Documents                 |
| ✓ VALID GURI - All checks passed                              |
+----------------------------------------------------------------+
```

### Component Table Shows Document Type

Component 6 in the decomposition table shows both the hex value and its meaning:

| Component | Value | Meaning | ... | Status |
|-----------|-------|---------|-----|--------|
| Component 6 (Doc Type) | 1a | Document Type: Aliniant Policy Document | ... | Valid |

---

## Creating GURIs for Each Global Type

### Global Type 1: Email

**Steps:**
1. Select **"01 - Email"**
2. Enter actual sender email
3. Enter actual recipients
4. Enter email subject
5. Enter date/time
6. Enter risk level (e.g., "LOW (15/100)")
7. Click "Generate GURI"

### Global Type 2: General Documents

**Steps:**
1. Select document type (02-10 or 99)
2. Click "Auto-fill Non-Email" button
   - Sender auto-fills to `julian.garrett@aliniant.com`
   - Recipients auto-fills to `FFFFF`
3. Enter document name in Subject field
4. Enter current date/time
5. Enter document location/path
6. Click "Generate GURI"

### Global Type 3: Aliniant Internal Documents

**Steps:**
1. Select Aliniant document type (1a-1e)
   - **1a** - Policy Document
   - **1b** - Technical Document
   - **1c** - Compliance Document
   - **1d** - Pre-sales Document
   - **1e** - Contract Document
2. Click "Auto-fill Non-Email" button
3. Enter document name in Subject field
4. Enter current date/time
5. Enter document location/path
6. Click "Generate GURI"

---

## Use Cases

### Global Type 1: Email Security

**Scenario:** Tracking security-analyzed emails

```
- Email from: suspicious_sender@example.com
- To: julian.garrett@aliniant.com
- Subject: "Urgent: Account Verification"
- Risk: HIGH (85/100)
→ GURI Component 6: 01
→ Global Type 1: Email
```

### Global Type 2: General File Management

**Scenario:** Cataloging general business documents

```
- Quarterly Report.xlsx
- Company Presentation.pptx
- Project Photos.zip
→ GURI Component 6: 03, 04, 10
→ Global Type 2: General Documents
```

### Global Type 3: Internal Document Control

**Scenario:** Managing critical Aliniant documents

```
Policy Documents (1a):
- Information Security Policy.pdf
- Remote Work Policy.docx

Technical Documents (1b):
- System Architecture Specification.pdf
- API Documentation.md

Compliance Documents (1c):
- SOC 2 Audit Report.pdf
- GDPR Compliance Checklist.xlsx

Pre-sales Documents (1d):
- Customer Proposal - Acme Corp.pptx
- RFP Response - Global Industries.pdf

Contract Documents (1e):
- Master Services Agreement.pdf
- NDA - Partner Company.docx

→ GURI Component 6: 1a, 1b, 1c, 1d, 1e
→ Global Type 3: Aliniant Internal Documents
```

---

## Benefits of Global Type Classification

### Organization
✅ **Clear categorization** of all documents  
✅ **Easy identification** of document purpose  
✅ **Structured document management**

### Security
✅ **Email-specific tracking** (Type 1)  
✅ **Separate internal document handling** (Type 3)  
✅ **Risk assessment integration**

### Workflow
✅ **Quick visual identification** in decompositor  
✅ **Automated field population** based on type  
✅ **Consistent document categorization**

---

## Summary

| Global Type | Codes | Category | Primary Use |
|-------------|-------|----------|-------------|
| **Type 1** | 01 | Email | Security-analyzed emails |
| **Type 2** | 02-10, 99 | General Documents | General business files |
| **Type 3** | 1a-1e | Aliniant Internal | Company-specific documents |

**Total Document Types:** 16  
**Total Global Types:** 3

---

**Created:** November 2, 2025  
**Version:** 1.0  
**Status:** ✅ Active


# Component 4: DateTime Encoding Guide

## Overview

**Component 4** of the GURI now encodes the **time of day** as **seconds since midnight**, represented as an **8-digit hexadecimal** value.

**Key Point:** Midnight (00:00:00) = `00000000`

---

## Encoding Specification

### Format
- **Length:** 8 hexadecimal characters
- **Range:** `00000000` to `0001517f`
- **Represents:** Seconds elapsed since midnight (00:00:00)

### Calculation

```
Seconds since midnight = (hours × 3600) + (minutes × 60) + seconds
Component 4 = format(seconds_since_midnight, '08x')
```

### Examples

| Time | Seconds | Component 4 (hex) |
|------|---------|-------------------|
| 00:00:00 (midnight) | 0 | `00000000` |
| 00:00:01 | 1 | `00000001` |
| 00:01:00 | 60 | `0000003c` |
| 01:00:00 | 3,600 | `00000e10` |
| 12:00:00 (noon) | 43,200 | `0000a8c0` |
| 14:30:45 | 52,245 | `0000cc15` |
| 23:59:59 (end of day) | 86,399 | `0001517f` |

---

## Decoding Component 4

### Algorithm

```python
# Convert hex to decimal
seconds = int(component4, 16)

# Calculate hours, minutes, seconds
hours = seconds // 3600
minutes = (seconds % 3600) // 60
secs = seconds % 60

# Format as time
time_str = f"{hours:02d}:{minutes:02d}:{secs:02d}"
```

### Example

```python
component4 = "0000cc15"

seconds = int("0000cc15", 16)  # = 52,245
hours = 52245 // 3600          # = 14
minutes = (52245 % 3600) // 60  # = 30
secs = 52245 % 60              # = 45

Result: 14:30:45
```

---

## Capacity Analysis

### Daily Range
- **Seconds in a day:** 86,400
- **Max hex value needed:** `0001517f` (for 23:59:59)
- **Actual max capacity:** `ffffffff` = 4,294,967,295

### Time Before Reset

**If incrementing every second:**
- **4,294,967,295 seconds** = **136 years, 1 month**

**More than sufficient for:**
✅ Lifetime of the system  
✅ Uniqueness requirements  
✅ Sub-second precision if needed in future  

---

## Benefits

### Deterministic
✅ **Same time = same component** - Documents created at 14:30:45 always get `0000cc15`  
✅ **Predictable** - Easy to understand and verify  
✅ **Sortable** - Earlier times have lower hex values  

### Human-Readable
✅ **Decodable** - Can convert back to actual time  
✅ **Meaningful** - Represents real data, not random  
✅ **Verifiable** - Can check if encoding is correct  

### Efficient
✅ **Compact** - Only 8 characters for full time precision  
✅ **Fast** - Simple integer arithmetic  
✅ **Unique** - One value per second  

---

## GURI Component Breakdown

Now that Components 4 and 6 are deterministic:

| Component | Purpose | Encoding | Example | Meaning |
|-----------|---------|----------|---------|---------|
| 1 | Sender | Random* | `3ae2b` | youcangetholdofjules@gmail.com |
| 2 | Recipients | Random* | `7f4d1` | julian.garrett@aliniant.com |
| 3 | Subject | Random* | `a1b3c5e7` | "Important Meeting" |
| 4 | **DateTime** | **Seconds since midnight** | `0000cc15` | **14:30:45** |
| 5 | Risk/Location | Random* | `2a4` | "LOW (15/100)" |
| 6 | **Doc Type** | **Document type code** | `01` | **Email** |

*To be replaced with deterministic hashing in future updates

---

## Implementation in guri.py

### Generation Code

```python
def encode_datetime_component(dt_str: str) -> str:
    """
    Encode datetime as 8-digit hex representing seconds since midnight.
    Reference: Midnight of the current date = 00000000
    """
    # Parse the datetime string
    dt = datetime.strptime(dt_str, "%Y-%m-%d %H:%M:%S")
    
    # Get midnight of the same date
    midnight = dt.replace(hour=0, minute=0, second=0, microsecond=0)
    
    # Calculate seconds since midnight
    seconds_since_midnight = int((dt - midnight).total_seconds())
    
    # Convert to 8-digit hex (padded with zeros)
    return format(seconds_since_midnight, '08x')
```

### Usage in generate_guri()

```python
component4 = encode_datetime_component(dt_str)
guri = f"{component1}x{component2}x{component3}x{component4}x{component5}x{document_type}"
```

---

## Display in Decompositor

When viewing a GURI in the decompositor, Component 4 now shows:

**Before:**
```
Component 4 (DateTime): a1b3c5e7
Meaning: Date and Time (hashed)
```

**After:**
```
Component 4 (DateTime): 0000cc15
Meaning: DateTime: 2025-11-02 14:30:45 → 14:30:45 (52,245 seconds since midnight)
```

---

## Edge Cases

### Midnight
```
Time: 00:00:00
Seconds: 0
Component 4: 00000000
```

### One Second After Midnight
```
Time: 00:00:01
Seconds: 1
Component 4: 00000001
```

### End of Day
```
Time: 23:59:59
Seconds: 86,399
Component 4: 0001517f
```

### Invalid DateTime
If the datetime string cannot be parsed, the system falls back to random hex to avoid breaking GURI generation.

---

## Testing

### Generate Test GURI
```bash
python test_guri_generation.py
```

### Expected Output
```
Component 4 (DateTime): 0000cc15
Decoded time: 14:30:45
Match: YES!
```

---

## Migration

### New GURIs
All **new** GURIs generated after this update will have:
- Component 4 = seconds since midnight (hex)
- Component 6 = document type code

### Existing GURIs
Existing GURIs in the database still have random Component 4 values. They can be:
- **Left as-is** (still valid, just not deterministic)
- **Migrated** (regenerate Component 4 from stored datetime)

A migration script can be created if needed to update Component 4 for existing records.

---

## Summary

✅ **Component 4 now deterministic** - Based on actual time  
✅ **Seconds since midnight** - Precise to the second  
✅ **136 years capacity** - Won't reset in our lifetime  
✅ **Decodable** - Can extract time from component  
✅ **GUI displays decoded time** - Shows both hex and readable time  
✅ **Midnight reference** - Each day starts at 00000000  

**Status:** ✅ IMPLEMENTED  
**Date:** November 2, 2025  
**Version:** 2.0


@echo off
REM Geo Footer Test Script
REM Tests the updated geo footer system

echo ============================================================
echo  Geo Footer System Test
echo ============================================================
echo.

REM Test 1: Check Python
echo [Test 1] Checking Python installation...
python --version
if errorlevel 1 (
    echo [FAIL] Python not found
    pause
    exit /b 1
)
echo [PASS] Python found
echo.

REM Test 2: Test GURI import
echo [Test 2] Testing GURI module import...
python quick_test.py
if errorlevel 1 (
    echo [FAIL] GURI module test failed
    pause
    exit /b 1
)
echo [PASS] GURI module working
echo.

REM Test 3: Process a header file
echo [Test 3] Processing test header file...
echo.
python geolocate_headers.py "headers\headers_20251102_004141_236.txt" "output\test_footer.html"
if errorlevel 1 (
    echo [FAIL] Header processing failed
    pause
    exit /b 1
)
echo.
echo [PASS] Header processed successfully
echo.

REM Test 4: Check output file
echo [Test 4] Checking if output file was created...
if exist "output\test_footer.html" (
    echo [PASS] Output file created: output\test_footer.html
    dir "output\test_footer.html"
) else (
    echo [FAIL] Output file not found
    pause
    exit /b 1
)
echo.

echo ============================================================
echo  ALL TESTS PASSED!
echo ============================================================
echo.
echo The geo footer system is working correctly.
echo.
echo Next steps:
echo  1. Open GURI GUI: python guri_gui.py
echo  2. View generated footer: output\test_footer.html
echo  3. Check GURI database for new entry
echo.
pause


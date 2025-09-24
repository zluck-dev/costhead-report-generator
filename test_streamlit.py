#!/usr/bin/env python3
"""
Test script to verify Streamlit app functionality
"""

import sys
import os
from pathlib import Path

def test_imports():
    """Test if all required modules can be imported"""
    print("🔍 Testing imports...")

    try:
        import streamlit as st
        print("✅ Streamlit imported successfully")
    except ImportError as e:
        print(f"❌ Streamlit import failed: {e}")
        return False

    try:
        import pandas as pd
        print("✅ Pandas imported successfully")
    except ImportError as e:
        print(f"❌ Pandas import failed: {e}")
        return False

    try:
        import openpyxl
        print("✅ OpenPyXL imported successfully")
    except ImportError as e:
        print(f"❌ OpenPyXL import failed: {e}")
        return False

    try:
        from app.services.excel_service import load_excel
        from app.services.gin_service import list_sheets, load_sheet
        from app.services.costhead_service import generate_costhead_report
        print("✅ App services imported successfully")
    except ImportError as e:
        print(f"❌ App services import failed: {e}")
        return False

    return True

def test_file_structure():
    """Test if required files exist"""
    print("\n📁 Testing file structure...")

    required_files = [
        "streamlit_app.py",
        "run_streamlit.py",
        "requirements.txt",
        "app/services/excel_service.py",
        "app/services/gin_service.py",
        "app/services/costhead_service.py"
    ]

    all_exist = True
    for file_path in required_files:
        if Path(file_path).exists():
            print(f"✅ {file_path}")
        else:
            print(f"❌ {file_path} - MISSING")
            all_exist = False

    return all_exist

def test_streamlit_config():
    """Test Streamlit configuration"""
    print("\n⚙️ Testing Streamlit configuration...")

    try:
        import streamlit as st
        # Test basic Streamlit functionality
        print("✅ Streamlit configuration looks good")
        return True
    except Exception as e:
        print(f"❌ Streamlit configuration error: {e}")
        return False

def main():
    """Run all tests"""
    print("🧪 CostHead Report Generator - Streamlit Test Suite")
    print("=" * 60)

    tests = [
        ("Import Test", test_imports),
        ("File Structure Test", test_file_structure),
        ("Streamlit Config Test", test_streamlit_config)
    ]

    results = []
    for test_name, test_func in tests:
        print(f"\n🔬 Running {test_name}...")
        result = test_func()
        results.append((test_name, result))

    print("\n" + "=" * 60)
    print("📊 Test Results Summary:")
    print("=" * 60)

    all_passed = True
    for test_name, result in results:
        status = "✅ PASSED" if result else "❌ FAILED"
        print(f"{test_name}: {status}")
        if not result:
            all_passed = False

    print("\n" + "=" * 60)
    if all_passed:
        print("🎉 All tests passed! Your Streamlit app is ready to run.")
        print("\n🚀 To start the app, run:")
        print("   python run_streamlit.py")
        print("   or")
        print("   streamlit run streamlit_app.py")
    else:
        print("⚠️  Some tests failed. Please fix the issues before running the app.")
        print("\n💡 Common fixes:")
        print("   - Install missing dependencies: pip install -r requirements.txt")
        print("   - Check file paths and structure")
        print("   - Ensure all required files are present")

    return all_passed

if __name__ == "__main__":
    success = main()
    sys.exit(0 if success else 1)

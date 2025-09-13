#!/usr/bin/env python3
"""
Test script for DOCA table reference extraction - FULL OUTPUT
"""

import requests
import os

def test_doca_full_output():
    """Test DOCA extraction with input3.xlsx and show FULL output"""
    
    test_file = "input3.xlsx"
    
    if not os.path.exists(test_file):
        print(f"❌ Test file {test_file} not found")
        return False
    
    print(f"🧪 Testing DOCA extraction with {test_file} - FULL OUTPUT")
    
    url = "http://localhost:5000/upload"
    
    with open(test_file, 'rb') as f:
        files = {'file': (test_file, f, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')}
        data = {'mode': 'DOCA'}
        
        try:
            response = requests.post(url, files=files, data=data, timeout=30)
            
            if response.status_code == 200:
                result = response.json()
                
                if result.get('success'):
                    print("✅ DOCA extraction test successful!")
                    print(f"📄 Message: {result.get('message', 'No message')}")
                    
                    # Get the FULL preview
                    preview = result.get('preview', '')
                    
                    print("\n" + "="*80)
                    print("📋 COMPLETE DOCA OUTPUT:")
                    print("="*80)
                    print(preview)
                    print("="*80)
                    
                    # Analyze the output
                    lines = preview.split('\n')
                    total_lines = len(lines)
                    doca_lines = 0
                    
                    for line in lines:
                        if 'TABLE REFERENCE DATA (DOCA MODE)' in line or 'RECORD TYPE' in line or 'BBK' in line or 'QBK' in line:
                            doca_lines += 1
                    
                    print(f"\n📊 Analysis:")
                    print(f"  Total lines: {total_lines}")
                    print(f"  DOCA-related lines: {doca_lines}")
                    
                    return True
                else:
                    print(f"❌ DOCA extraction failed: {result.get('error', 'Unknown error')}")
                    return False
            else:
                print(f"❌ HTTP error: {response.status_code}")
                return False
                
        except requests.exceptions.RequestException as e:
            print(f"❌ Request failed: {e}")
            return False

if __name__ == "__main__":
    print("🚀 Testing DOCA table reference extraction - FULL OUTPUT...")
    print("=" * 80)
    
    success = test_doca_full_output()
    
    print("=" * 80)
    if success:
        print("🎉 DOCA extraction test completed!")
    else:
        print("❌ DOCA extraction test failed.")

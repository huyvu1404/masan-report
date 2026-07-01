#!/usr/bin/env python3
"""
Script phân tích và list tất cả chart trong template để debug
"""

import json
from pathlib import Path

def analyze_template_structure():
    """Phân tích cấu trúc template để hiểu rõ chart mapping"""
    
    with open("template_structure.json", "r", encoding="utf-8") as f:
        structure = json.load(f)
    
    print("=" * 80)
    print("PHÂN TÍCH CHART STRUCTURE")
    print("=" * 80)
    
    for slide_name, slide_data in structure.items():
        charts = slide_data.get("charts", [])
        if not charts:
            continue
            
        print(f"\n📄 {slide_name.upper()} - {len(charts)} charts:")
        print("-" * 50)
        
        for i, chart in enumerate(charts, 1):
            rid = chart.get("rId", "")
            file = chart.get("file", "")
            types = ", ".join(chart.get("types", []))
            series_count = len(chart.get("series", []))
            
            print(f"  Chart {i}: {rid} -> {file}")
            print(f"    Type: {types}")
            print(f"    Series: {series_count}")
            
            # Print first series info for context
            series = chart.get("series", [])
            if series:
                first_series = series[0]
                title = first_series.get("title", "")
                cats_count = len(first_series.get("cats", []))
                vals_count = len(first_series.get("vals", []))
                print(f"    First series: '{title}' ({cats_count} cats, {vals_count} vals)")
                
                # Show categories for context
                cats = first_series.get("cats", [])[:3]  # First 3 only
                if cats:
                    print(f"    Categories: {cats}{'...' if len(first_series.get('cats', [])) > 3 else ''}")
                    
            print()

def analyze_current_code():
    """Phân tích code hiện tại trong populate_pptx để tìm lỗi"""
    
    print("=" * 80)
    print("MAPPING TRONG CODE HIỆN TẠI")
    print("=" * 80)
    
    # Slide 2 mappings từ code
    slide2_mapping = {
        "rId3": "chart6: share of voice doughnut",
        "rId4": "chart7: daily line chart", 
        "rId5": "chart8: brand discussion comparison bar"
    }
    
    print("\n📄 SLIDE 2:")
    for rid, desc in slide2_mapping.items():
        print(f"  {rid} -> {desc}")
    
    # Slide 3 mappings từ code  
    slide3_sentiment_rids = ["rId3", "rId4", "rId5", "rId6", "rId7", "rId8"]
    slide3_source_rids = ["rId9", "rId10", "rId11", "rId12"]
    
    print("\n📄 SLIDE 3:")
    print("  Sentiment charts:", slide3_sentiment_rids)
    print("  Source charts:", slide3_source_rids)
    
    # Slide 4 mappings từ code
    slide4_mapping = {
        "rId6": "chart19: verbatim bar",
        "rId7": "chart20: topic sentiment bar",
        "rId8": "chart21: doughnut sentiment", 
        "rId9": "chart22: top sources bar",
        "rId10": "chart23: daily line"
    }
    
    print("\n📄 SLIDE 4:")
    for rid, desc in slide4_mapping.items():
        print(f"  {rid} -> {desc}")

if __name__ == "__main__":
    analyze_template_structure()
    analyze_current_code()
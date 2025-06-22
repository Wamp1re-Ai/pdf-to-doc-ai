#!/usr/bin/env python3
"""
Advanced Features Module for PDF to Word Converter
Additional enhancements and utilities
"""

import os
import json
import time
from datetime import datetime
from pathlib import Path
import logging

logger = logging.getLogger(__name__)

class ConversionStats:
    """Track conversion statistics and performance"""
    
    def __init__(self):
        self.stats_file = "conversion_stats.json"
        self.stats = self.load_stats()
    
    def load_stats(self):
        """Load existing statistics"""
        try:
            if os.path.exists(self.stats_file):
                with open(self.stats_file, 'r') as f:
                    return json.load(f)
        except Exception as e:
            logger.warning(f"Failed to load stats: {e}")
        
        return {
            "total_conversions": 0,
            "successful_conversions": 0,
            "failed_conversions": 0,
            "total_pages_processed": 0,
            "average_processing_time": 0,
            "models_used": {},
            "extraction_methods_used": {},
            "last_conversion": None
        }
    
    def save_stats(self):
        """Save statistics to file"""
        try:
            with open(self.stats_file, 'w') as f:
                json.dump(self.stats, f, indent=2)
        except Exception as e:
            logger.warning(f"Failed to save stats: {e}")
    
    def record_conversion(self, success=True, pages=0, processing_time=0, 
                         model_used=None, extraction_method=None):
        """Record a conversion attempt"""
        self.stats["total_conversions"] += 1
        
        if success:
            self.stats["successful_conversions"] += 1
            self.stats["total_pages_processed"] += pages
            
            # Update average processing time
            if self.stats["successful_conversions"] > 1:
                current_avg = self.stats["average_processing_time"]
                new_avg = ((current_avg * (self.stats["successful_conversions"] - 1)) + processing_time) / self.stats["successful_conversions"]
                self.stats["average_processing_time"] = new_avg
            else:
                self.stats["average_processing_time"] = processing_time
        else:
            self.stats["failed_conversions"] += 1
        
        # Track model usage
        if model_used:
            if model_used not in self.stats["models_used"]:
                self.stats["models_used"][model_used] = 0
            self.stats["models_used"][model_used] += 1
        
        # Track extraction method usage
        if extraction_method:
            if extraction_method not in self.stats["extraction_methods_used"]:
                self.stats["extraction_methods_used"][extraction_method] = 0
            self.stats["extraction_methods_used"][extraction_method] += 1
        
        self.stats["last_conversion"] = datetime.now().isoformat()
        self.save_stats()
    
    def get_success_rate(self):
        """Calculate success rate percentage"""
        if self.stats["total_conversions"] == 0:
            return 0
        return (self.stats["successful_conversions"] / self.stats["total_conversions"]) * 100
    
    def get_summary(self):
        """Get formatted statistics summary"""
        return f"""
📊 Conversion Statistics:
• Total conversions: {self.stats['total_conversions']}
• Successful: {self.stats['successful_conversions']} ({self.get_success_rate():.1f}%)
• Failed: {self.stats['failed_conversions']}
• Pages processed: {self.stats['total_pages_processed']}
• Average time: {self.stats['average_processing_time']:.1f}s
• Most used model: {max(self.stats['models_used'], key=self.stats['models_used'].get) if self.stats['models_used'] else 'None'}
"""

class DocumentAnalyzer:
    """Analyze document structure and provide insights"""
    
    @staticmethod
    def analyze_text_structure(text):
        """Analyze text structure and provide insights"""
        lines = text.split('\n')
        
        analysis = {
            "total_lines": len(lines),
            "non_empty_lines": len([l for l in lines if l.strip()]),
            "potential_headings": 0,
            "paragraphs": 0,
            "word_count": 0,
            "character_count": len(text),
            "average_line_length": 0,
            "structure_quality": "Unknown"
        }
        
        # Count words
        words = text.split()
        analysis["word_count"] = len(words)
        
        # Analyze line structure
        non_empty_lines = [l for l in lines if l.strip()]
        if non_empty_lines:
            analysis["average_line_length"] = sum(len(l) for l in non_empty_lines) / len(non_empty_lines)
        
        # Count potential headings and paragraphs
        for line in lines:
            line = line.strip()
            if not line:
                continue
                
            # Potential heading detection
            if (len(line) < 80 and 
                (line.isupper() or 
                 line.endswith(':') or
                 any(line.startswith(prefix) for prefix in ['Chapter', 'Section', 'Part']))):
                analysis["potential_headings"] += 1
            else:
                analysis["paragraphs"] += 1
        
        # Assess structure quality
        if analysis["potential_headings"] > 0 and analysis["paragraphs"] > 0:
            heading_ratio = analysis["potential_headings"] / analysis["non_empty_lines"]
            if 0.05 <= heading_ratio <= 0.3:  # 5-30% headings is good structure
                analysis["structure_quality"] = "Good"
            elif heading_ratio < 0.05:
                analysis["structure_quality"] = "Poor (few headings)"
            else:
                analysis["structure_quality"] = "Poor (too many headings)"
        else:
            analysis["structure_quality"] = "Poor (no clear structure)"
        
        return analysis
    
    @staticmethod
    def suggest_improvements(analysis):
        """Suggest improvements based on analysis"""
        suggestions = []
        
        if analysis["structure_quality"] == "Poor (few headings)":
            suggestions.append("Consider adding more section headings to improve document structure")
        
        if analysis["structure_quality"] == "Poor (too many headings)":
            suggestions.append("Some headings might be misidentified - review heading detection")
        
        if analysis["average_line_length"] > 200:
            suggestions.append("Very long lines detected - may indicate formatting issues")
        
        if analysis["word_count"] < 100:
            suggestions.append("Document seems very short - verify extraction was complete")
        
        return suggestions

class QualityChecker:
    """Check conversion quality and detect issues"""
    
    @staticmethod
    def check_spacing_quality(text):
        """Check for spacing issues in text"""
        issues = []
        
        # Check for merged words (common patterns)
        merged_patterns = [
            r'[a-z][A-Z]',  # camelCase
            r'[a-zA-Z]\d',  # word+number
            r'\d[a-zA-Z]',  # number+word
            r'[.!?][A-Z]',  # punctuation+capital
        ]
        
        import re
        total_issues = 0
        for pattern in merged_patterns:
            matches = re.findall(pattern, text)
            if matches:
                total_issues += len(matches)
        
        if total_issues > 0:
            issues.append(f"Found {total_issues} potential spacing issues")
        
        # Check for excessive spaces
        excessive_spaces = re.findall(r' {3,}', text)
        if excessive_spaces:
            issues.append(f"Found {len(excessive_spaces)} instances of excessive spacing")
        
        return issues
    
    @staticmethod
    def check_text_quality(text):
        """Overall text quality assessment"""
        issues = []
        
        # Check for garbled text
        import re
        garbled_patterns = [
            r'[^\w\s\.,!?;:()\-\'\"]{3,}',  # 3+ special chars in a row
            r'\w{50,}',  # Very long words (likely merged)
        ]
        
        for pattern in garbled_patterns:
            matches = re.findall(pattern, text)
            if matches:
                issues.append(f"Potential garbled text detected: {len(matches)} instances")
        
        return issues

def create_conversion_report(input_file, output_file, stats, analysis, quality_issues):
    """Create a detailed conversion report"""
    report = f"""
# PDF to Word Conversion Report

**Generated:** {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}

## File Information
- **Input:** {input_file}
- **Output:** {output_file}
- **File Size:** {os.path.getsize(output_file) if os.path.exists(output_file) else 'Unknown'} bytes

## Document Analysis
- **Total Lines:** {analysis['total_lines']}
- **Word Count:** {analysis['word_count']}
- **Character Count:** {analysis['character_count']}
- **Potential Headings:** {analysis['potential_headings']}
- **Paragraphs:** {analysis['paragraphs']}
- **Structure Quality:** {analysis['structure_quality']}
- **Average Line Length:** {analysis['average_line_length']:.1f} characters

## Quality Assessment
"""
    
    if quality_issues:
        report += "**Issues Found:**\n"
        for issue in quality_issues:
            report += f"- {issue}\n"
    else:
        report += "✅ No quality issues detected\n"
    
    report += f"\n{stats.get_summary()}"
    
    return report

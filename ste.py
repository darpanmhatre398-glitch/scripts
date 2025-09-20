import streamlit as st
import spacy
from pathlib import Path
import os
import time
from typing import List, Dict, Any, Optional
import logging
from dataclasses import dataclass
from concurrent.futures import ThreadPoolExecutor, as_completed
import traceback
import json
from datetime import datetime
import io

# Configure environment and logging
os.environ["STREAMLIT_WATCHER_TYPE"] = "none"
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

from components.ste_word_checker import STEReplacer, extract_text_from_file
from components.llm_utills import rewrite_to_active
from components.post_active_processor import process_active_and_polish
from components.punctuation import detect_punctuation_violations
from components.hyphen_suggester import detect_hyphen_suggestions
from components.si_unit_checker import check_si_units
from components.multiword_noun_checker import MultiwordNounChecker

@dataclass
class CheckResult:
    """Data class for storing check results"""
    category: str
    violations: List[Dict[str, Any]]
    success_message: str
    error_message: Optional[str] = None

class STEDocumentProcessor:
    """Main processor class for STE document checking"""
    
    def __init__(self):
        self.nlp = None
        self.ste = None
        self._initialize_components()
    
    def _initialize_components(self):
        """Initialize spaCy and STE components with error handling"""
        try:
            with st.spinner("Loading language model..."):
                self.nlp = spacy.load("en_core_web_sm")
                if "parser" not in self.nlp.pipe_names and "sentencizer" not in self.nlp.pipe_names:
                    self.nlp.add_pipe("sentencizer")
            
            with st.spinner("Loading STE replacer..."):
                self.ste = STEReplacer(json_dir="data")
                
        except Exception as e:
            st.error(f"Failed to initialize components: {str(e)}")
            st.stop()
    
    def process_document(self, text: str) -> Dict[str, CheckResult]:
        """Process document through all STE checks"""
        paragraphs = [p.strip() for p in text.split("\n") if p.strip()]
        
        # Use progress bar for better UX
        progress_bar = st.progress(0)
        status_text = st.empty()
        
        checks = [
            ("STE Word Replacement", self._check_ste_words),
            ("Passive Voice", self._check_passive_voice),
            ("Punctuation", self._check_punctuation),
            ("Hyphenation", self._check_hyphenation),
            ("SI Units", self._check_si_units),
            ("Multi-word Nouns", self._check_multiword_nouns)
        ]
        
        results = {}
        
        for i, (check_name, check_func) in enumerate(checks):
            status_text.text(f"Running {check_name} check...")
            progress_bar.progress((i + 1) / len(checks))
            
            try:
                results[check_name] = check_func(paragraphs)
            except Exception as e:
                logger.error(f"Error in {check_name}: {str(e)}")
                results[check_name] = CheckResult(
                    category=check_name,
                    violations=[],
                    success_message="",
                    error_message=f"Error during {check_name} check: {str(e)}"
                )
        
        progress_bar.empty()
        status_text.empty()
        
        return results
    
    def _check_ste_words(self, paragraphs: List[str]) -> CheckResult:
        """Check for STE word violations"""
        flagged = []
        
        for line_num, line in enumerate(paragraphs, start=1):
            try:
                doc = self.ste.nlp(line)
                for token in doc:
                    if hasattr(token._, 'was_replaced') and token._.was_replaced:
                        flagged.append({
                            "line": line_num,
                            "original": token.text,
                            "pos": token.pos_,
                            "replacement": getattr(token._, 'ste_replacement', 'N/A'),
                            "context": line
                        })
            except Exception as e:
                logger.warning(f"Error processing line {line_num}: {str(e)}")
        
        return CheckResult(
            category="STE Word Replacement",
            violations=flagged,
            success_message="✅ No STE violations found."
        )
    
    def _check_passive_voice(self, paragraphs: List[str]) -> CheckResult:
        """Check for passive voice and provide rewrites"""
        passive_results = []
        
        for para_num, para in enumerate(paragraphs, start=1):
            try:
                doc = self.nlp(para)
                for sent in doc.sents:
                    for token in sent:
                        if token.tag_ == "VBN":
                            auxiliaries = [child.text for child in token.children 
                                         if child.dep_ in ("aux", "auxpass")]
                            if auxiliaries:
                                phrase = " ".join(auxiliaries + [token.text])
                                passive_results.append({
                                    "paragraph": para_num,
                                    "phrase": phrase.strip(),
                                    "sentence": sent.text.strip()
                                })
                                break
            except Exception as e:
                logger.warning(f"Error processing paragraph {para_num}: {str(e)}")
        
        return CheckResult(
            category="Passive Voice",
            violations=passive_results,
            success_message="✅ No passive voice detected."
        )
    
    def _check_punctuation(self, paragraphs: List[str]) -> CheckResult:
        """Check for punctuation violations"""
        try:
            violations = detect_punctuation_violations(paragraphs)
            return CheckResult(
                category="Punctuation",
                violations=violations,
                success_message="✅ No punctuation issues detected."
            )
        except Exception as e:
            return CheckResult(
                category="Punctuation",
                violations=[],
                success_message="",
                error_message=f"Error checking punctuation: {str(e)}"
            )
    
    def _check_hyphenation(self, paragraphs: List[str]) -> CheckResult:
        """Check for hyphenation suggestions"""
        try:
            suggestions = detect_hyphen_suggestions(paragraphs)
            violations = []
            
            for line_text, entries in suggestions.items():
                for entry in entries:
                    violations.append({
                        "line_number": entry.get('line_number'),
                        "suggestion": entry.get('suggestion'),
                        "original": entry.get('original'),
                        "context": line_text
                    })
            
            return CheckResult(
                category="Hyphenation",
                violations=violations,
                success_message="✅ No hyphenation suggestions found."
            )
        except Exception as e:
            return CheckResult(
                category="Hyphenation",
                violations=[],
                success_message="",
                error_message=f"Error checking hyphenation: {str(e)}"
            )
    
    def _check_si_units(self, paragraphs: List[str]) -> CheckResult:
        """Check for SI unit compliance"""
        try:
            violations = check_si_units(paragraphs)
            return CheckResult(
                category="SI Units",
                violations=violations,
                success_message="✅ All units use proper SI formatting."
            )
        except Exception as e:
            return CheckResult(
                category="SI Units",
                violations=[],
                success_message="",
                error_message=f"Error checking SI units: {str(e)}"
            )
    
    def _check_multiword_nouns(self, paragraphs: List[str]) -> CheckResult:
        """Check for multi-word noun violations"""
        try:
            checker = MultiwordNounChecker()
            checker.process(paragraphs)
            report_lines = checker.report()
            
            violations = [{"report_line": line} for line in report_lines] if report_lines else []
            
            return CheckResult(
                category="Multi-word Nouns",
                violations=violations,
                success_message="✅ No multi-word noun violations."
            )
        except Exception as e:
            return CheckResult(
                category="Multi-word Nouns",
                violations=[],
                success_message="",
                error_message=f"Error checking multi-word nouns: {str(e)}"
            )

def display_results(results: Dict[str, CheckResult], processor: STEDocumentProcessor):
    """Display results in an organized manner with expandable sections"""
    
    # Summary section
    st.subheader("📊 Summary")
    
    total_violations = sum(len(result.violations) for result in results.values())
    
    if total_violations == 0:
        st.success("🎉 Document passes all STE checks!")
    else:
        st.warning(f"⚠️ Found {total_violations} issues across {len([r for r in results.values() if r.violations])} categories")
    
    # Create summary metrics
    col1, col2, col3 = st.columns(3)
    
    with col1:
        st.metric("Total Issues", total_violations)
    
    with col2:
        categories_with_issues = len([r for r in results.values() if r.violations])
        st.metric("Categories with Issues", categories_with_issues)
    
    with col3:
        compliance_rate = ((len(results) - categories_with_issues) / len(results)) * 100
        st.metric("Compliance Rate", f"{compliance_rate:.1f}%")
    
    # Detailed results
    st.subheader("🔍 Detailed Results")
    
    for category, result in results.items():
        with st.expander(f"{get_category_icon(category)} {category} ({len(result.violations)} issues)", 
                        expanded=bool(result.violations)):
            
            if result.error_message:
                st.error(result.error_message)
                continue
            
            if not result.violations:
                st.success(result.success_message)
                continue
            
            # Display violations based on category
            if category == "STE Word Replacement":
                display_ste_violations(result.violations)
            elif category == "Passive Voice":
                display_passive_voice_violations(result.violations, processor)
            elif category == "Punctuation":
                display_punctuation_violations(result.violations)
            elif category == "Hyphenation":
                display_hyphenation_violations(result.violations)
            elif category == "SI Units":
                display_si_unit_violations(result.violations)
            elif category == "Multi-word Nouns":
                display_multiword_noun_violations(result.violations)

def get_category_icon(category: str) -> str:
    """Get appropriate icon for each category"""
    icons = {
        "STE Word Replacement": "🛠️",
        "Passive Voice": "⚠️",
        "Punctuation": "🔤",
        "Hyphenation": "➖",
        "SI Units": "📏",
        "Multi-word Nouns": "🧠"
    }
    return icons.get(category, "📋")

def display_ste_violations(violations: List[Dict]):
    """Display STE word replacement violations"""
    for entry in violations:
        st.markdown(f"🔹 **Line {entry['line']}**: `{entry['original']}` ({entry['pos']}) → `{entry['replacement']}`")
        st.markdown(f"   *Context: {entry['context']}*")

def display_passive_voice_violations(violations: List[Dict], processor: STEDocumentProcessor):
    """Display passive voice violations with rewrites"""
    for i, info in enumerate(violations, 1):
        st.markdown(f"**{i}.** Passive Phrase: `{info['phrase']}`")
        st.markdown(f"➡️ **Original**: *{info['sentence']}*")
        
        try:
            with st.spinner("Rewriting to active voice..."):
                active = rewrite_to_active(info['sentence'])
                st.markdown(f"🔁 **Active Voice**: *{active}*")
                
                result = process_active_and_polish(active)
                if result:
                    st.markdown(f"🟡 **With STE highlights**: {result.get('highlighted_ste', 'N/A')}")
                    st.markdown(f"🛠️ **Replacements**: `{result.get('replacements', 'None')}`")
                    st.markdown(f"✨ **Final polished**: *{result.get('final_polished', active)}*")
        except Exception as e:
            st.error(f"Error rewriting sentence: {str(e)}")
        
        st.markdown("---")

def display_punctuation_violations(violations: List[Dict]):
    """Display punctuation violations"""
    for v in violations:
        line = v.get("line_number", "Unknown")
        text = v.get("text", "")
        punctuation = v.get("punctuation", "")
        st.markdown(f"⚠️ **Line {line}**: contains `{punctuation}` — *{text}*")
        st.markdown(f"💡 **Suggestion**: Avoid using `{punctuation}` — rewrite using a period or conjunction.")

def display_hyphenation_violations(violations: List[Dict]):
    """Display hyphenation suggestions"""
    for entry in violations:
        st.markdown(f"💡 **Line {entry['line_number']}**: `{entry['suggestion']}` instead of `{entry['original']}`")
        st.markdown(f"   *Context: {entry['context']}*")

def display_si_unit_violations(violations: List[Dict]):
    """Display SI unit violations"""
    for entry in violations:
        line = entry.get("line", "Unknown")
        text = entry.get("text", "")
        suggestion = entry.get("suggestion", "")
        st.markdown(f"💡 **Line {line}**: `{suggestion}` in — *{text}*")
        for issue in entry.get("issues", []):
            st.markdown(f"   ⚠️ {issue}")

def display_multiword_noun_violations(violations: List[Dict]):
    """Display multi-word noun violations"""
    for entry in violations:
        st.markdown(entry["report_line"])

def generate_report(results: Dict[str, CheckResult], filename: str, doc_stats: Dict) -> Dict[str, str]:
    """Generate comprehensive reports in multiple formats"""
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    base_filename = Path(filename).stem
    
    # Calculate summary statistics
    total_violations = sum(len(result.violations) for result in results.values())
    categories_with_issues = len([r for r in results.values() if r.violations])
    compliance_rate = ((len(results) - categories_with_issues) / len(results)) * 100
    
    # Generate JSON Report
    json_report = {
        "metadata": {
            "filename": filename,
            "timestamp": timestamp,
            "total_violations": total_violations,
            "categories_checked": len(results),
            "categories_with_issues": categories_with_issues,
            "compliance_rate": round(compliance_rate, 2),
            "document_stats": doc_stats
        },
        "results": {}
    }
    
    for category, result in results.items():
        json_report["results"][category] = {
            "violation_count": len(result.violations),
            "violations": result.violations,
            "success_message": result.success_message,
            "error_message": result.error_message
        }
    
    # Generate Markdown Report
    md_report = f"""# STE Analysis Report

**Document:** {filename}  
**Generated:** {timestamp}  
**Total Issues:** {total_violations}  
**Compliance Rate:** {compliance_rate:.1f}%

## Document Statistics
- **Lines:** {doc_stats.get('lines', 0)}
- **Words:** {doc_stats.get('words', 0)}
- **Characters:** {doc_stats.get('chars', 0)}

## Summary
- Categories Checked: {len(results)}
- Categories with Issues: {categories_with_issues}
- Overall Compliance: {compliance_rate:.1f}%

## Detailed Results

"""
    
    for category, result in results.items():
        icon = get_category_icon(category)
        md_report += f"### {icon} {category}\n"
        
        if result.error_message:
            md_report += f"**Error:** {result.error_message}\n\n"
            continue
            
        if not result.violations:
            md_report += f"{result.success_message}\n\n"
            continue
            
        md_report += f"**Issues Found:** {len(result.violations)}\n\n"
        
        # Add category-specific details
        if category == "STE Word Replacement":
            for entry in result.violations:
                md_report += f"- **Line {entry['line']}:** `{entry['original']}` ({entry['pos']}) → `{entry['replacement']}`\n"
                md_report += f"  - *Context:* {entry['context']}\n"
        
        elif category == "Passive Voice":
            for i, entry in enumerate(result.violations, 1):
                md_report += f"- **{i}.** Passive phrase: `{entry['phrase']}`\n"
                md_report += f"  - *Sentence:* {entry['sentence']}\n"
        
        elif category == "Punctuation":
            for entry in result.violations:
                md_report += f"- **Line {entry.get('line_number', 'Unknown')}:** Contains `{entry.get('punctuation', '')}` in: {entry.get('text', '')}\n"
        
        elif category == "Hyphenation":
            for entry in result.violations:
                md_report += f"- **Line {entry['line_number']}:** Use `{entry['suggestion']}` instead of `{entry['original']}`\n"
        
        elif category == "SI Units":
            for entry in result.violations:
                md_report += f"- **Line {entry.get('line', 'Unknown')}:** {entry.get('suggestion', '')}\n"
                for issue in entry.get('issues', []):
                    md_report += f"  - {issue}\n"
        
        elif category == "Multi-word Nouns":
            for entry in result.violations:
                md_report += f"- {entry['report_line']}\n"
        
        md_report += "\n"
    
    md_report += f"""
---
*Report generated by STE Document Checker on {timestamp}*
"""
    
    # Generate Plain Text Report
    txt_report = f"""STE ANALYSIS REPORT
{'='*50}

Document: {filename}
Generated: {timestamp}
Total Issues: {total_violations}
Compliance Rate: {compliance_rate:.1f}%

DOCUMENT STATISTICS
{'-'*20}
Lines: {doc_stats.get('lines', 0)}
Words: {doc_stats.get('words', 0)}
Characters: {doc_stats.get('chars', 0)}

SUMMARY
{'-'*20}
Categories Checked: {len(results)}
Categories with Issues: {categories_with_issues}
Overall Compliance: {compliance_rate:.1f}%

DETAILED RESULTS
{'-'*20}

"""
    
    for category, result in results.items():
        txt_report += f"{category.upper()}\n{'-' * len(category)}\n"
        
        if result.error_message:
            txt_report += f"Error: {result.error_message}\n\n"
            continue
            
        if not result.violations:
            txt_report += f"{result.success_message}\n\n"
            continue
            
        txt_report += f"Issues Found: {len(result.violations)}\n\n"
        
        # Add simplified details for text format
        for i, violation in enumerate(result.violations, 1):
            if category == "STE Word Replacement":
                txt_report += f"{i}. Line {violation['line']}: {violation['original']} -> {violation['replacement']}\n"
            elif category == "Passive Voice":
                txt_report += f"{i}. {violation['phrase']} in: {violation['sentence']}\n"
            elif category == "Punctuation":
                txt_report += f"{i}. Line {violation.get('line_number', 'Unknown')}: {violation.get('text', '')}\n"
            elif category == "Hyphenation":
                txt_report += f"{i}. Line {violation['line_number']}: {violation['suggestion']}\n"
            elif category == "SI Units":
                txt_report += f"{i}. Line {violation.get('line', 'Unknown')}: {violation.get('suggestion', '')}\n"
            elif category == "Multi-word Nouns":
                txt_report += f"{i}. {violation['report_line']}\n"
        
        txt_report += "\n"
    
    txt_report += f"\nReport generated by STE Document Checker on {timestamp}\n"
    
    return {
        "json": json.dumps(json_report, indent=2),
        "markdown": md_report,
        "text": txt_report,
        "base_filename": base_filename
    }

def save_reports_automatically(reports: Dict[str, str], base_filename: str) -> List[str]:
    """Save all report formats automatically and return saved file paths"""
    saved_files = []
    reports_dir = Path("reports")
    reports_dir.mkdir(exist_ok=True)
    
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    
    formats = {
        "json": "json",
        "markdown": "md", 
        "text": "txt"
    }
    
    for format_name, extension in formats.items():
        if format_name in reports:
            filename = f"{base_filename}_report_{timestamp}.{extension}"
            filepath = reports_dir / filename
            
            try:
                with open(filepath, 'w', encoding='utf-8') as f:
                    f.write(reports[format_name])
                saved_files.append(str(filepath))
            except Exception as e:
                st.error(f"Failed to save {format_name} report: {str(e)}")
    
    return saved_files

def process_uploaded_file(uploaded_file) -> Optional[str]:
    """Process uploaded file and extract text"""
    original_name = uploaded_file.name
    suffix = Path(original_name).suffix.lower()
    
    if suffix not in [".txt", ".pdf", ".adoc"]:
        st.error("❌ Unsupported file format. Please upload a .txt, .pdf, or .adoc file.")
        return None
    
    temp_path = Path(f"uploaded_temp_{int(time.time())}{suffix}")
    
    try:
        with open(temp_path, "wb") as f:
            f.write(uploaded_file.read())
        
        with st.spinner(f"Extracting text from {original_name}..."):
            input_text = extract_text_from_file(str(temp_path))
            
        return input_text
        
    except ValueError as e:
        st.error(f"❌ Error extracting text: {str(e)}")
        return None
    except Exception as e:
        st.error(f"❌ Unexpected error: {str(e)}")
        return None
    finally:
        if temp_path.exists():
            try:
                os.remove(temp_path)
            except Exception as e:
                logger.warning(f"Could not remove temp file: {str(e)}")

def main():
    """Main application function"""
    # Page configuration
    st.set_page_config(
        page_title="STE Document Checker",
        page_icon="✍️",
        layout="wide",
        initial_sidebar_state="expanded"
    )
    
    # Header
    st.title("✍️ STE Document Checker")
    st.markdown("**Simplified Technical English (STE) compliance checker for technical documents**")
    
    # Sidebar with information
    with st.sidebar:
        st.header("📋 About STE Checker")
        st.markdown("""
        This tool checks your technical documents for compliance with 
        **Simplified Technical English (STE)** standards including:
        
        - 🛠️ **Word Replacement**: Non-approved words
        - ⚠️ **Passive Voice**: Detection and active rewrites  
        - 🔤 **Punctuation**: Prohibited punctuation marks
        - ➖ **Hyphenation**: Compound word suggestions
        - 📏 **SI Units**: Proper unit formatting
        - 🧠 **Multi-word Nouns**: Noun phrase violations
        """)
        
        st.header("📊 Document Stats")
        if 'doc_stats' in st.session_state:
            stats = st.session_state.doc_stats
            st.metric("Lines", stats.get('lines', 0))
            st.metric("Words", stats.get('words', 0))
            st.metric("Characters", stats.get('chars', 0))
    
    # File upload section
    st.subheader("📄 Upload Document")
    uploaded_file = st.file_uploader(
        "Choose a file",
        type=["txt", "pdf", "adoc"],
        help="Supported formats: .txt, .pdf, .adoc"
    )
    
    if uploaded_file:
        # Process file
        input_text = process_uploaded_file(uploaded_file)
        
        if input_text:
            # Store document stats
            lines = len(input_text.split('\n'))
            words = len(input_text.split())
            chars = len(input_text)
            
            st.session_state.doc_stats = {
                'lines': lines,
                'words': words,
                'chars': chars
            }
            
            # Show document preview
            with st.expander("📖 Document Preview", expanded=False):
                preview_text = input_text[:1000] + "..." if len(input_text) > 1000 else input_text
                st.text_area("Document Content", preview_text, height=200, disabled=True)
            
            # Initialize processor and run checks
            if st.button("🚀 Run STE Analysis", type="primary"):
                processor = STEDocumentProcessor()
                
                with st.spinner("Processing document..."):
                    results = processor.process_document(input_text)
                
                # Display results
                display_results(results, processor)
                
                # Export results option
                st.subheader("📥 Export Results")
                if st.button("📋 Copy Summary to Clipboard"):
                    total_issues = sum(len(result.violations) for result in results.values())
                    summary = f"STE Analysis Summary\n{'='*20}\n"
                    summary += f"Total Issues: {total_issues}\n"
                    summary += f"Categories Checked: {len(results)}\n\n"
                    
                    for category, result in results.items():
                        if result.violations:
                            summary += f"{category}: {len(result.violations)} issues\n"
                    
                    st.code(summary, language="text")
                    st.success("📋 Summary ready to copy!")

if __name__ == "__main__":
    main()

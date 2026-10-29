import io
import re
from typing import List, Dict, Any
import PyPDF2
import docx
import openpyxl

def clean_text(text: str) -> str:
    """Clean up whitespace and unprintable characters."""
    if not text:
        return ""
    text = re.sub(r'[\r\t]+', ' ', text)
    text = re.sub(r'\n{3,}', '\n\n', text)
    text = re.sub(r'[ ]{2,}', ' ', text)
    return text.strip()

def detect_section_header(text: str) -> str:
    """Detect section or policy titles at the beginning of chunks/pages."""
    lines = [line.strip() for line in text.split("\n") if line.strip()]
    for line in lines[:3]:
        # Matches e.g. "Section 4.1", "Chapter 2", "Policy: Leave", "1. Overview", "ARTICLE III"
        if re.match(r'^(section|chapter|article|policy|clause|part)\b', line, re.IGNORECASE):
            return line[:100]
        if re.match(r'^[0-9]+(\.[0-9]+)*\s+[A-Z]', line):
            return line[:100]
        if len(line) < 80 and line.isupper():
            return line[:100]
    return ""

def chunk_text(text: str, max_chunk_size: int = 1000, overlap: int = 200) -> List[str]:
    """Split text into overlapping chunks keeping paragraph integrity."""
    if not text:
        return []
    
    if len(text) <= max_chunk_size:
        return [text]
    
    chunks = []
    start = 0
    while start < len(text):
        end = start + max_chunk_size
        if end >= len(text):
            chunks.append(text[start:].strip())
            break
        
        # Try to break on paragraph or sentence boundary
        split_point = text.rfind("\n\n", start, end)
        if split_point == -1 or split_point < start + max_chunk_size // 2:
            split_point = text.rfind(". ", start, end)
        if split_point == -1 or split_point < start + max_chunk_size // 2:
            split_point = text.rfind("\n", start, end)
        if split_point == -1 or split_point < start + max_chunk_size // 2:
            split_point = end
        else:
            split_point += 1  # include punctuation
            
        chunk = text[start:split_point].strip()
        if chunk:
            chunks.append(chunk)
        start = max(start + 1, split_point - overlap)
        
    return chunks

def parse_pdf(file_bytes: bytes, file_name: str) -> List[Dict[str, Any]]:
    """Parse PDF document and return chunks with page numbers and section headers."""
    chunks_data = []
    try:
        reader = PyPDF2.PdfReader(io.BytesIO(file_bytes))
        current_section = ""
        for page_idx, page in enumerate(reader.pages):
            page_text = page.extract_text() or ""
            cleaned = clean_text(page_text)
            if not cleaned:
                continue
            
            section = detect_section_header(cleaned)
            if section:
                current_section = section

            page_chunks = chunk_text(cleaned, max_chunk_size=1000, overlap=150)
            for chunk_idx, chk in enumerate(page_chunks):
                header_prefix = f"[{current_section}]\n" if current_section else ""
                chunks_data.append({
                    "text": header_prefix + chk,
                    "raw_text": chk,
                    "section": current_section,
                    "page": page_idx + 1,
                    "total_pages": len(reader.pages),
                    "chunk_index": chunk_idx,
                    "file_name": file_name
                })
    except Exception as e:
        chunks_data.append({
            "text": f"Error parsing PDF {file_name}: {str(e)}",
            "page": 1,
            "total_pages": 1,
            "chunk_index": 0,
            "file_name": file_name
        })
    return chunks_data

def parse_docx(file_bytes: bytes, file_name: str) -> List[Dict[str, Any]]:
    """Parse DOCX document and return chunks."""
    chunks_data = []
    try:
        doc = docx.Document(io.BytesIO(file_bytes))
        paragraphs = []
        for p in doc.paragraphs:
            t = p.text.strip()
            if t:
                paragraphs.append(t)
        
        # Also extract table text
        for table in doc.tables:
            for row in table.rows:
                row_cells = [cell.text.strip() for cell in row.cells if cell.text.strip()]
                if row_cells:
                    paragraphs.append(" | ".join(row_cells))
        
        full_text = "\n\n".join(paragraphs)
        cleaned = clean_text(full_text)
        doc_chunks = chunk_text(cleaned, max_chunk_size=1000, overlap=150)
        
        for idx, chk in enumerate(doc_chunks):
            section = detect_section_header(chk)
            header_prefix = f"[{section}]\n" if section else ""
            chunks_data.append({
                "text": header_prefix + chk,
                "raw_text": chk,
                "section": section,
                "page": 1,
                "total_pages": 1,
                "chunk_index": idx,
                "file_name": file_name
            })
    except Exception as e:
        chunks_data.append({
            "text": f"Error parsing DOCX {file_name}: {str(e)}",
            "page": 1,
            "total_pages": 1,
            "chunk_index": 0,
            "file_name": file_name
        })
    return chunks_data

def parse_xlsx(file_bytes: bytes, file_name: str) -> List[Dict[str, Any]]:
    """Parse Excel spreadsheet."""
    chunks_data = []
    try:
        wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)
        for sheet_name in wb.sheetnames:
            sheet = wb[sheet_name]
            rows_text = []
            for row in sheet.iter_rows(values_only=True):
                non_empty = [str(cell) for cell in row if cell is not None and str(cell).strip()]
                if non_empty:
                    rows_text.append(", ".join(non_empty))
            sheet_content = f"Sheet: {sheet_name}\n" + "\n".join(rows_text[:200])
            cleaned = clean_text(sheet_content)
            chunks = chunk_text(cleaned, max_chunk_size=1000, overlap=100)
            for idx, chk in enumerate(chunks):
                chunks_data.append({
                    "text": chk,
                    "page": sheet_name,
                    "total_pages": len(wb.sheetnames),
                    "chunk_index": idx,
                    "file_name": file_name
                })
    except Exception as e:
        chunks_data.append({
            "text": f"Error parsing Excel {file_name}: {str(e)}",
            "page": "1",
            "total_pages": 1,
            "chunk_index": 0,
            "file_name": file_name
        })
    return chunks_data

def parse_plain_text(file_bytes: bytes, file_name: str) -> List[Dict[str, Any]]:
    """Parse text, CSV, markdown, JSON, etc."""
    try:
        text = file_bytes.decode("utf-8", errors="replace")
    except Exception:
        text = file_bytes.decode("latin-1", errors="replace")
        
    cleaned = clean_text(text)
    chunks = chunk_text(cleaned, max_chunk_size=1000, overlap=150)
    chunks_data = []
    for idx, chk in enumerate(chunks):
        chunks_data.append({
            "text": chk,
            "page": 1,
            "total_pages": 1,
            "chunk_index": idx,
            "file_name": file_name
        })
    return chunks_data

def parse_file(file_bytes: bytes, file_name: str) -> List[Dict[str, Any]]:
    """Main router to parse document files based on extension."""
    ext = file_name.lower().split(".")[-1] if "." in file_name else ""
    if ext == "pdf":
        return parse_pdf(file_bytes, file_name)
    elif ext in ["docx", "doc"]:
        return parse_docx(file_bytes, file_name)
    elif ext in ["xlsx", "xls"]:
        return parse_xlsx(file_bytes, file_name)
    elif ext in ["txt", "md", "csv", "json", "log", "xml", "html"]:
        return parse_plain_text(file_bytes, file_name)
    else:
        return parse_plain_text(file_bytes, file_name)

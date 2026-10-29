import re
import math
from collections import Counter
from typing import List, Dict, Any, Optional

STOPWORDS = {
    "a", "about", "above", "after", "again", "against", "all", "am", "an", "and", "any", "are", 
    "as", "at", "be", "because", "been", "before", "being", "below", "between", "both", "but", 
    "by", "can", "could", "did", "do", "does", "doing", "down", "during", "each", "few", "for", 
    "from", "further", "had", "has", "have", "having", "he", "her", "here", "hers", "herself", 
    "him", "himself", "his", "how", "i", "if", "in", "into", "is", "it", "its", "itself", "just", 
    "me", "more", "most", "my", "myself", "no", "nor", "not", "now", "of", "off", "on", "once", 
    "only", "or", "other", "ought", "our", "ours", "ourselves", "out", "over", "own", "same", 
    "she", "should", "so", "some", "such", "than", "that", "the", "their", "theirs", "them", 
    "themselves", "then", "there", "these", "they", "this", "those", "through", "to", "too", 
    "under", "until", "up", "very", "was", "we", "were", "what", "when", "where", "which", 
    "while", "who", "whom", "why", "with", "would", "you", "your", "yours", "yourself", "yourselves"
}

def tokenize(text: str) -> List[str]:
    """Tokenize string into lowercase alphanumeric words."""
    if not text:
        return []
    tokens = re.findall(r'\b[a-zA-Z0-9_\-\.]+\b', text.lower())
    return [t for t in tokens if len(t) > 1 and t not in STOPWORDS]

class BM25Index:
    """Fast in-memory BM25 index for document chunks and list records."""
    def __init__(self, k1: float = 1.5, b: float = 0.75):
        self.k1 = k1
        self.b = b
        self.docs: List[Dict[str, Any]] = []
        self.doc_lengths: List[int] = []
        self.avg_doc_len: float = 0.0
        self.doc_frequencies: Dict[str, int] = Counter()
        self.inverted_index: Dict[str, List[int]] = {}

    def build(self, documents: List[Dict[str, Any]], text_field: str = "text"):
        self.text_field = text_field
        self.docs = documents
        self.doc_lengths = []
        self.doc_frequencies = Counter()
        self.inverted_index = {}
        
        total_len = 0
        for doc_id, doc in enumerate(documents):
            text = doc.get(text_field, "")
            tokens = tokenize(text)
            self.doc_lengths.append(len(tokens))
            total_len += len(tokens)
            
            seen_terms = set(tokens)
            for term in seen_terms:
                self.doc_frequencies[term] += 1
                if term not in self.inverted_index:
                    self.inverted_index[term] = []
                self.inverted_index[term].append(doc_id)
                
        self.avg_doc_len = total_len / max(len(documents), 1)

    def search(self, query: str, top_k: int = 5, filter_fn=None) -> List[Dict[str, Any]]:
        if not self.docs or not query.strip():
            return []
            
        query_tokens = tokenize(query)
        if not query_tokens:
            return []
            
        n_docs = len(self.docs)
        scores = [0.0] * n_docs
        field = getattr(self, "text_field", "text")
        
        for term in query_tokens:
            if term not in self.inverted_index:
                continue
            df = self.doc_frequencies[term]
            idf = math.log((n_docs - df + 0.5) / (df + 0.5) + 1.0)
            
            for doc_id in self.inverted_index[term]:
                doc_len = self.doc_lengths[doc_id]
                # count term occurrences in doc text
                doc_text = self.docs[doc_id].get(field, "").lower()
                tf = doc_text.count(term)
                
                score_num = tf * (self.k1 + 1.0)
                score_denom = tf + self.k1 * (1.0 - self.b + self.b * (doc_len / max(self.avg_doc_len, 1.0)))
                scores[doc_id] += idf * (score_num / score_denom)

        # Exact phrase bonus
        query_clean = query.strip().lower()
        if len(query_clean) > 3:
            for doc_id in range(n_docs):
                doc_text = self.docs[doc_id].get(field, "").lower()
                if query_clean in doc_text:
                    scores[doc_id] += 5.0

        # Rank documents
        scored_docs = []
        for doc_id, score in enumerate(scores):
            if score > 0:
                doc = self.docs[doc_id]
                if filter_fn and not filter_fn(doc):
                    continue
                doc_copy = dict(doc)
                doc_copy["relevance_score"] = round(score, 3)
                scored_docs.append(doc_copy)
                
        scored_docs.sort(key=lambda x: x["relevance_score"], reverse=True)
        return scored_docs[:top_k]

class SharePointIndexManager:
    """Manages search indices for both Document Libraries and Lists."""
    def __init__(self):
        self.doc_index = BM25Index()
        self.list_index = BM25Index()
        self.all_document_chunks: List[Dict[str, Any]] = []
        self.all_list_records: List[Dict[str, Any]] = []

    def set_documents(self, chunks: List[Dict[str, Any]]):
        self.all_document_chunks = chunks
        self.doc_index.build(chunks, text_field="text")

    def set_lists(self, records: List[Dict[str, Any]]):
        self.all_list_records = records
        self.list_index.build(records, text_field="searchable_text")

    def search_documents(self, query: str, library_name: Optional[str] = None, top_k: int = 5) -> List[Dict[str, Any]]:
        filter_fn = None
        if library_name:
            lib_lower = library_name.strip().lower()
            # Check if it matches any library name or document file name
            has_matching_lib = any(lib_lower in d.get("library_name", "").lower() for d in self.all_document_chunks)
            has_matching_file = any(lib_lower in d.get("file_name", "").lower() for d in self.all_document_chunks)
            
            if has_matching_lib:
                filter_fn = lambda d: lib_lower in d.get("library_name", "").lower()
            elif has_matching_file:
                filter_fn = lambda d: lib_lower in d.get("file_name", "").lower()

        results = self.doc_index.search(query, top_k=top_k, filter_fn=filter_fn)
        # If filtered search yielded 0 results, fallback to search across all documents
        if not results and filter_fn:
            results = self.doc_index.search(query, top_k=top_k)
        return results

    def search_lists(self, query: str, list_name: Optional[str] = None, top_k: int = 10) -> List[Dict[str, Any]]:
        filter_fn = None
        if list_name:
            l_lower = list_name.strip().lower()
            filter_fn = lambda r: r.get("list_name", "").lower() == l_lower
        return self.list_index.search(query, top_k=top_k, filter_fn=filter_fn)

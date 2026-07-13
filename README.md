# FastDocSearch - Lightning Fast Document Search Tool

**Purpose**: Quickly validate if an LLM response is grounded in your source documents (PDFs, DOCX, TXT, MD, PPTX, etc.). Instead of manually opening files and Ctrl+F, index once and search instantly across thousands of pages. Perfect for RAG/LLM output validation, compliance checks, research, or finding quotes in large document sets.

## Why it's fast
- Uses **Whoosh** (high-performance pure-Python search engine) with BM25 ranking.
- Granular indexing: PDF pages, DOCX paragraphs/slides, text chunks.
- Persistent index (search in <100ms even on large collections after initial indexing).
- Rich terminal output with highlighted snippets, file locations, scores, and context.

## Supported Formats
- **PDF** (.pdf) - page-level extraction with page numbers
- **Word** (.docx) - paragraph + table level
- **PowerPoint** (.pptx) - slide level
- **Text/Markdown** (.txt, .md, .rst, .html) - paragraph level
- **Excel** (.xlsx, .xlsm) - sheet + row text (optional)

## Installation (one-time)

```bash
pip install whoosh pdfplumber python-docx python-pptx rich click pypdf
```

( `pdfplumber` recommended for best PDF text quality; falls back to `pypdf` if issues. )

## Quick Start

```bash
# 1. Index your documents (run once or when docs change)
python cli.py index --docs-dir /path/to/your/documents --index-dir ./my_project_index

# 2. Search for phrases from the LLM response (exact or keywords)
python cli.py search --query '"large language model"' --index-dir ./my_project_index --limit 5

# Or with context and nice formatting
python cli.py search --query 'RAG validation grounding' --index-dir ./my_project_index --context-chars 300 --show-score
```

## Example Use Case (LLM Validation)
LLM says: "According to the policy document, employees must complete ethics training annually."

You run:
```bash
python cli.py search --query 'ethics training annually' --index-dir ./policies_index
```

It instantly shows:
- policy_v3.pdf (Page 12): "...All employees **must complete the mandatory ethics training annually** by December 31st..."
- onboarding_guide.docx (Paragraph 45): "Annual ethics certification is required..."

You can verify grounding in <5 seconds instead of 10+ minutes of manual searching.

## Commands

### `index`
Builds or rebuilds the search index from a directory (recursive).

Options:
- `--docs-dir` (required): Root folder containing documents (scans subfolders too)
- `--index-dir` (default: `.docsearch_index`): Where to store the Whoosh index
- `--force`: Rebuild from scratch even if index exists
- `--file-types`: Comma-separated extensions to include (default: pdf,docx,txt,md,pptx,xlsx)
- `--max-chars-per-unit`: For very long pages/sections, split into chunks (default: no split, let Whoosh handle)

Example:
```bash
python cli.py index --docs-dir ./project_docs --index-dir ./indexes/project1 --force
```

### `search`
Fast search with highlighting and rich output.

Options:
- `--query` (required): Search terms. Use `"exact phrase"` for precise matches. Supports AND, OR, wildcards (*), fuzzy (~), etc.
- `--index-dir`: Path to the index
- `--limit` / `-k`: Max results (default 10)
- `--context-chars`: Characters of surrounding context to show in snippet (default 250)
- `--show-score`: Display BM25 relevance score
- `--file-filter`: Only search in files matching glob (e.g. `*policy*.pdf`)

Example advanced:
```bash
python cli.py search --query '"annual ethics training" OR "code of conduct"' --index-dir ./indexes/project1 -k 20 --context-chars 400
```

### `validate`
Check whether **one English LLM response** is grounded in your index.

The whole input is treated as a **single response**. It is split into sentences/claims; for each claim the tool extracts search phrases (numbers, multi-word windows, keywords via simple English stopword heuristics — no spaCy/YAKE) and runs a Whoosh search ladder (exact phrase → soft keywords). Each claim is labeled **grounded**, **partial**, or **not_found**.

```bash
# From a file (one response per file)
python cli.py validate --response-file answer.txt --index-dir ./my_project_index

# Inline
python cli.py validate --text "Employees must complete ethics training annually." --index-dir ./policies_index

# Stdin
cat answer.txt | python cli.py validate --index-dir ./my_project_index --show-queries
```

Options:
- `--response-file` / `-f`: path to the LLM answer file
- `--text` / `-t`: inline response text
- `--index-dir` / `-i`: index path
- `--limit` / `-k`: evidence hits per claim (default 2)
- `--max-claims`: cap claims processed (default 20)
- `--context-chars`: snippet size
- `--show-queries`: print generated queries (debug)

**v1 limits:** English only; one response per run; lexical search only (no embeddings / second LLM).

### `info`
Show index statistics (number of documents indexed, last update, etc.)

```bash
python cli.py info --index-dir ./my_index
```

### `clear`
Delete an index (frees space).

## Tips for Best Results (LLM Grounding Validation)
- Use **exact phrases** from the LLM output in double quotes for highest precision (`search`), or use `validate` to automate claim-level checks.
- Search key facts, numbers, names, or unique phrases first.
- If no results: the claim may be hallucinated or paraphrased too loosely — try broader keywords or check for synonyms.
- Re-index after adding/updating documents in the folder.
- For very large collections (1000+ PDFs), indexing may take a few minutes initially — it's a one-time cost.

## Performance Notes
- Indexing speed: ~50-200 pages/sec depending on PDF complexity (on modern laptop).
- Search speed: Usually <50ms for first results.
- Index size: Typically 10-30% of original text size (compressed + inverted index).
- Tested with 500+ PDFs / 50k+ pages — still instant search.

## Future Enhancements (possible)
- Optional YAKE / spaCy phrase extraction backends
- Multi-language validate
- Semantic search (embeddings) toggle
- Export results to JSON/CSV for automation
- Web UI (Gradio/Streamlit)
- Diff/highlight changes between document versions
- Integration with Obsidian / VS Code

## License & Credits
Built for fast, local, private document search and LLM output validation.
Uses Whoosh, pdfplumber, python-docx, python-pptx, rich, click.

Happy validating! If the LLM can't point to the exact page/paragraph, it probably made it up. 🚀

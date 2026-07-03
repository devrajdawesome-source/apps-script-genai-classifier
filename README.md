# RAG-Augmented GenAI Document Classifier & Parser

An asynchronous, serverless Document Parsing automation engineered in Google Apps Script. This system processes high-volume student marksheet uploads via multi-stage OCR REST endpoints and leverages Large Language Models (LLMs) for dynamic, zero-shot layout classification.

## 🛠️ System Architecture & Workflow
1. **Asynchronous Ingestion:** Time-driven Google Apps Script triggers batch-process incoming spreadsheet rows to handle platform timeout constraints.
2. **OCR Feature Extraction:** Document URLs are dispatched to structured REST endpoints to convert image artifacts into raw text tokens.
3. **Retrieval-Augmented Generation (RAG):** The system cross-references extracted text features against organizational mapping rules hosted natively within the sheet environment to augment data contextualization.
4. **LLM Classification:** Clean tokens are evaluated via targeted prompt constraints for precise zero-shot educational board validation.
5. **Telemetry & Reporting:** Formatted analytics and regional distribution cohorts are dynamically pushed to a dedicated "Campus Summary" live dashboard.

## 📝 Status Note
*Production API endpoints and live enterprise API access keys have been decoupled from this repository for data privacy and security. The underlying core architecture, execution logic, and multi-sheet telemetry handlers remain fully documented within `Code.js` for evaluation.*

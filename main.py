import os
import fitz  # PyMuPDF
from fastapi import FastAPI, HTTPException, Request
from fastapi.responses import JSONResponse, FileResponse
from pydantic import BaseModel
import os
from fastapi.middleware.cors import CORSMiddleware
from a2wsgi import ASGIMiddleware
 
# Initialize FastAPI app
app = FastAPI()
origins=["*"]
 
app.add_middleware(
    CORSMiddleware,
    allow_origins=origins,
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)
wsgi_app = ASGIMiddleware(app)

class SearchRequest(BaseModel):
    folder_name: str
    keyword: str

def load_pdfs(folder_path):
    documents = []
    for filename in os.listdir(folder_path):
        if filename.endswith('.pdf') and not filename.endswith('_highlighted.pdf'):
            file_path = os.path.join(folder_path, filename)
            try:
                doc = fitz.open(file_path)
                full_text = ""
                for page_num in range(len(doc)):
                    page = doc.load_page(page_num)
                    text = page.get_text("text")  # Extract plain text
                    full_text += text
                documents.append({"filename": filename, "filepath": file_path, "content": full_text})
                doc.close()
            except Exception as e:
                print(f'Error processing {filename}: {str(e)}')
    return documents


def count_keyword_occurrences(documents, keyword):
    keyword_counts = {}
    keyword_contexts = {}
    for document in documents:
        text = document["content"]
        count = 0
        contexts = []
        start = 0
        while (index := text.lower().find(keyword.lower(), start)) != -1:
            count += 1
            start = index + len(keyword)
            # Extract context
            tokens = text.split()
            word_index = text[:index].count(' ')
            start_context = max(word_index - 5, 0)
            end_context = min(word_index + 11, len(tokens))
            context = ' '.join(tokens[start_context:end_context])
            contexts.append(context)
        if count > 0:
            keyword_counts[document["filename"]] = count
            keyword_contexts[document["filename"]] = contexts
    return keyword_counts, keyword_contexts

def sort_files_by_frequency(keyword_counts):
    sorted_files = sorted(keyword_counts.items(), key=lambda item: item[1], reverse=True)
    return sorted_files

def highlight_keyword_in_pdf(file_path, keyword):
    base, ext = os.path.splitext(file_path)
    highlighted_file_path = f"{base}_{keyword}_highlighted{ext}"
    
    if os.path.exists(highlighted_file_path):
        return highlighted_file_path
    
    doc = fitz.open(file_path)
    keyword_lower = keyword.lower()
    for page_num in range(len(doc)):
        page = doc.load_page(page_num)
        text_instances = page.search_for(keyword)
        if text_instances:
            print(f"Found {len(text_instances)} instances of keyword '{keyword}' on page {page_num + 1}")
        for inst in text_instances:
            highlight = page.add_highlight_annot(inst)
            highlight.update()
    
    doc.save(highlighted_file_path, garbage=4, deflate=True)
    doc.close()
    return highlighted_file_path


@app.post('/search')
async def search(search_request: SearchRequest, request: Request):
    folder_name = search_request.folder_name
    keyword = search_request.keyword
    with open("path.txt", 'r') as file:
        path = file.readline().strip()
    folder_path = os.path.join(path, folder_name)
    
    if not os.path.exists(folder_path):
        raise HTTPException(status_code=404, detail="Folder not found")
    
    documents = load_pdfs(folder_path)
    keyword_counts, keyword_contexts = count_keyword_occurrences(documents, keyword)
    sorted_files = sort_files_by_frequency(keyword_counts)

    result = []
    base_url = str(request.base_url)
    for filename, count in sorted_files:
        for doc in documents:
            if doc["filename"] == filename:
                #original_file_path = doc["filepath"]
                #original_url = f"{base_url}pdfs/{folder_name}/{os.path.basename(original_file_path)}"
                highlighted_file_path = highlight_keyword_in_pdf(doc["filepath"], keyword)
                highlighted_url = f"{base_url}highlighted_pdfs/{folder_name}/{os.path.basename(highlighted_file_path)}"
                result.append({
                    'filename': filename,
                    'filepath': doc["filepath"],
                    'count': count,
                    'highlighted_path': highlighted_url,
                    'contexts': keyword_contexts[filename]
                })

    return JSONResponse(content=result)

@app.get('/highlighted_pdfs/{folder_name}/{filename:path}')
async def open_file(folder_name: str, filename: str):
    with open("path.txt", 'r') as file:
        path = file.readline().strip()
    folder_path = os.path.join(path, folder_name)

    file_path = os.path.join(folder_path, filename)
    
    if not os.path.exists(file_path):
        raise HTTPException(status_code=404, detail="File not found")
    
    return FileResponse(path=file_path, filename=filename, media_type='application/pdf', headers={"Content-Disposition": "inline"})

@app.get('/pdfs/{folder_name}/{filename:path}')
async def open_file(folder_name: str, filename: str):
    with open("path.txt", 'r') as file:
        path = file.readline().strip()
    folder_path = os.path.join(path, folder_name)

    file_path = os.path.join(folder_path, filename)
    
    if not os.path.exists(file_path):
        raise HTTPException(status_code=404, detail="File not found")
    
    return FileResponse(path=file_path, filename=filename, media_type='application/pdf', headers={"Content-Disposition": "inline"})

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000)


'''
http://127.0.0.1:8000/search
body - {
  "folder_name": "SOPKK007",
  "keyword": "title"
}

http://127.0.0.1:8000/highlighted_pdfs/SOPKK007/aa_highlighted.pdf
'''
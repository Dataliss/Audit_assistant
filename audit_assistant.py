import gradio as gr
import pandas as pd
import PyPDF2
from docx import Document
import requests
import json
import os
import tempfile
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH

# API Configuration
GROQ_API_KEY = "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"
GROQ_API_URL = "https://api.groq.com/openai/v1/chat/completions"

# Constants for models
ANALYZER_MODEL = "llama3-70b-8192"  # Llama model for analysis
REASONER_MODEL = "llama3-70b-8192"  # Model for reasoning (CA equivalent)

# Document parsing functions
def parse_pdf(file_path):
    # Using PyPDF2 instead of PyMuPDF
    with open(file_path, 'rb') as file:
        pdf_reader = PyPDF2.PdfReader(file)
        text = ""
        for page_num in range(len(pdf_reader.pages)):
            page = pdf_reader.pages[page_num]
            text += page.extract_text() or ""
    return text

def parse_csv(file_path):
    df = pd.read_csv(file_path)
    return df

def parse_document(file_path):
    if file_path.endswith('.pdf'):
        return parse_pdf(file_path)
    elif file_path.endswith('.csv'):
        return parse_csv(file_path)
    elif file_path.endswith('.docx'):
        doc = Document(file_path)
        return "\n".join([para.text for para in doc.paragraphs])
    else:
        return "Unsupported file format"

# API call function
def call_groq_api(prompt, model_name, system_message="You are a helpful assistant"):
    headers = {
        "Authorization": f"Bearer {GROQ_API_KEY}",
        "Content-Type": "application/json"
    }
    
    data = {
        "model": model_name,
        "messages": [
            {"role": "system", "content": system_message},
            {"role": "user", "content": prompt}
        ],
        "temperature": 0.7
    }
    
    response = requests.post(GROQ_API_URL, headers=headers, json=data)
    
    if response.status_code == 200:
        return response.json()['choices'][0]['message']['content']
    else:
        return f"Error: {response.status_code}, {response.text}"

# Analyzer agent
def analyzer_agent(document_text, audit_type):
    system_message = """You are an expert financial document analyzer. 
    Your task is to extract key information, summarize important points, 
    and highlight any anomalies or patterns in financial documents. 
    Focus on being thorough and precise."""
    
    prompt = f"""Analyze the following document for a {audit_type}:
    
    {document_text[:10000]}  # Limiting text length to avoid token limits
    
    Please provide:
    1. A summary of the key financial information
    2. Important highlights that require attention
    3. Any potential anomalies or patterns
    4. Key metrics and figures that are relevant for this audit type
    
    Format your response clearly with sections and bullet points where appropriate."""
    
    return call_groq_api(prompt, ANALYZER_MODEL, system_message)

# Reasoner agent (CA equivalent)
def reasoner_agent(analysis_results, audit_type):
    system_message = """You are a senior Indian Chartered Accountant with extensive 
    experience in auditing and financial compliance. Your expertise includes deep knowledge 
    of Indian GAAP, Companies Act 2013, Income Tax Act, and GST regulations. 
    Provide expert reasoning and cite relevant regulations in your explanations."""
    
    prompt = f"""Based on the following analysis of financial documents for a {audit_type}, 
    provide your expert reasoning:
    
    {analysis_results}
    
    Please:
    1. Explain the significance of the key findings
    2. Identify compliance issues or concerns
    3. Provide reasoning for any anomalies
    4. Cite relevant Indian financial regulations and standards
    5. Offer professional recommendations
    
    Format your response in a professional audit report style with proper citations."""
    
    return call_groq_api(prompt, REASONER_MODEL, system_message)

# Report generation functions
def create_pdf_report(content, output_path):
    doc = Document()
    
    # Add title
    title = doc.add_heading('Audit Report', 0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    # Add date
    from datetime import datetime
    date_paragraph = doc.add_paragraph()
    date_paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    date_paragraph.add_run(f"Date: {datetime.now().strftime('%d-%m-%Y')}").bold = True
    
    doc.add_paragraph("\n")
    
    # Add content with formatting
    sections = content.split("\n\n")
    for section in sections:
        if section.strip().upper() == section.strip() and len(section.strip()) > 0:
            # This looks like a header
            doc.add_heading(section, level=1)
        else:
            p = doc.add_paragraph()
            for line in section.split("\n"):
                if line.startswith("•") or line.startswith("-"):
                    p = doc.add_paragraph(line, style='List Bullet')
                elif line.startswith("1.") or line.startswith("2."):
                    p = doc.add_paragraph(line, style='List Number')
                else:
                    p.add_run(line)
                    p.add_run("\n")
    
    # Save the document
    doc.save(output_path)
    return output_path

# Conversational audit assistant
def audit_assistant(message, history, uploaded_file=None, state=None):
    # Initialize state if it's the first message
    if state is None:
        state = {
            "step": "greeting",
            "audit_type": None,
            "document_path": None,
            "analysis_results": None,
            "reasoning_results": None,
            "constraints": None
        }
    
    # Create a copy of the history to avoid modifying the input
    new_history = history.copy() if history else []
    
    # Add user message to history if provided
    if message:
        new_history.append({"role": "user", "content": message})
    
    # Step 1: Greeting
    if state["step"] == "greeting":
        state["step"] = "ask_audit_type"
        new_history.append({
            "role": "assistant", 
            "content": "👋 Hello! I'm your AI Audit Assistant. I can help you perform different types of audits on your financial documents. What type of audit would you like to perform today? (e.g., Financial Audit, Tax Audit, Compliance Audit)"
        })
        return new_history, state, None
    
    # Step 2: Ask audit type
    elif state["step"] == "ask_audit_type":
        state["audit_type"] = message
        state["step"] = "suggest_documents"
        
        document_suggestions = {
            "financial audit": "balance sheet, income statement, cash flow statement, general ledger",
            "tax audit": "income tax returns, GST returns, expense records, tax computation sheets",
            "compliance audit": "regulatory filings, board minutes, policy documents, compliance certificates"
        }
        
        audit_type_lower = message.lower()
        for key in document_suggestions:
            if key in audit_type_lower:
                suggested_docs = document_suggestions[key]
                break
        else:
            suggested_docs = "financial statements, ledgers, and supporting documentation"
        
        # Fixed f-string by using explicit newlines instead of escape sequences
        response = f"You've selected: {message}. For this type of audit, I'll need to analyze relevant documents."
        response += "\n\nTypically, a {0} requires the following documents:".format(message.lower())
        
        # Convert comma-separated list to bullet points
        for doc in suggested_docs.split(", "):
            response += f"\n• {doc}"
        
        response += "\n\nPlease upload the documents you'd like me to analyze."
        
        new_history.append({"role": "assistant", "content": response})
        return new_history, state, None
    
    # Step 3: Document upload and constraints
    # Step 3: Document upload and constraints (modified)
    if state["step"] == "suggest_documents":
        if uploaded_file:
            # Handle multiple files
            files_to_process = uploaded_file if isinstance(uploaded_file, list) else [uploaded_file]
            state["document_paths"] = []
            
            for file in files_to_process:
                if isinstance(file, str) and os.path.exists(file):
                    state["document_path"].append(file)
                elif hasattr(file, "name"):
                    temp_dir = tempfile.mkdtemp()
                    temp_path = os.path.join(temp_dir, file.name)
                    
                    try:
                        if hasattr(file, "save"):
                            file.save(temp_path)
                        else:
                            with open(temp_path, "wb") as f:
                                if hasattr(file, "read"):
                                    f.write(file.read())
                                elif hasattr(file, "data"):
                                    f.write(file.data)
                        
                        state["document_path"].append(temp_path)
                    except Exception as e:
                        response = f"Error processing file {file.name}: {str(e)}"
                        new_history.append({"role": "assistant", "content": response})
                        return new_history, state, None
            
            state["step"] = "ask_constraints"
            response = "Thank you for uploading the document. Are there any specific constraints or areas you'd like me to focus on during this audit? (e.g., specific time periods, accounts, or compliance requirements)"
            new_history.append({"role": "assistant", "content": response})
            return new_history, state, None
        else:
            response = "I don't see any uploaded files. Please upload the documents you'd like me to analyze for the audit."
            new_history.append({"role": "assistant", "content": response})
            return new_history, state, None
    
    # Step 4: Ask constraints
    elif state["step"] == "ask_constraints":
        state["constraints"] = message
        state["step"] = "analyze_documents"
        
        response = "Thank you for providing those details. I'll now analyze your documents. This may take a moment..."
        new_history.append({"role": "assistant", "content": response})
        return new_history, state, None
    
    # Step 5: Analyze documents
    elif state["step"] == "analyze_documents":
        if state["document_path"]:
            document_text = parse_document(state["document_path"])
            
            if isinstance(document_text, pd.DataFrame):
                # Convert DataFrame to string for analysis
                document_text = f"CSV Data Summary:\n{document_text.describe().to_string()}\n\nFirst 10 rows:\n{document_text.head(10).to_string()}"
            
            analysis_results = analyzer_agent(document_text, state["audit_type"])
            state["analysis_results"] = analysis_results
            state["step"] = "ask_for_reasoning"
            
            response = f"## Document Analysis Complete\n\n{analysis_results}\n\nWould you like me to continue with reasoning and provide expert insights on these findings?"
            new_history.append({"role": "assistant", "content": response})
            return new_history, state, None
        else:
            state["step"] = "suggest_documents"
            response = "I don't have any documents to analyze. Please upload your documents first."
            new_history.append({"role": "assistant", "content": response})
            return new_history, state, None
    
    # Step 6: Ask for reasoning
    elif state["step"] == "ask_for_reasoning":
        if "yes" in message.lower() or "continue" in message.lower() or "proceed" in message.lower():
            state["step"] = "provide_reasoning"
            response = "I'll now provide expert reasoning on the analysis. This may take a moment..."
            new_history.append({"role": "assistant", "content": response})
            return new_history, state, None
        else:
            state["step"] = "ask_for_report"
            response = "Would you like me to generate a final audit report based on the analysis so far?"
            new_history.append({"role": "assistant", "content": response})
            return new_history, state, None
    
    # Step 7: Provide reasoning
    elif state["step"] == "provide_reasoning":
        reasoning_results = reasoner_agent(state["analysis_results"], state["audit_type"])
        state["reasoning_results"] = reasoning_results
        state["step"] = "ask_for_report"
        
        response = f"## Expert Reasoning\n\n{reasoning_results}\n\nWould you like me to generate a final audit report that includes both the analysis and reasoning?"
        new_history.append({"role": "assistant", "content": response})
        return new_history, state, None
    
    # Step 8: Generate report
    elif state["step"] == "ask_for_report":
        if "yes" in message.lower() or "generate" in message.lower() or "report" in message.lower():
            # Combine analysis and reasoning for the report
            report_content = f"# AUDIT REPORT\n\n"
            report_content += f"## Type of Audit: {state['audit_type']}\n\n"
            
            if state["constraints"]:
                report_content += f"## Audit Constraints\n{state['constraints']}\n\n"
            
            report_content += f"## Document Analysis\n{state['analysis_results']}\n\n"
            
            if state["reasoning_results"]:
                report_content += f"## Expert Reasoning and Recommendations\n{state['reasoning_results']}\n\n"
            
            # Generate report file
            temp_dir = tempfile.mkdtemp()
            report_path = os.path.join(temp_dir, "Audit_Report.docx")
            create_pdf_report(report_content, report_path)
            
            state["step"] = "complete"
            response = f"I've prepared your audit report. You can download it using the button below.\n\nIs there anything else you'd like me to help you with?"
            new_history.append({"role": "assistant", "content": response})
            return new_history, state, report_path
        else:
            state["step"] = "complete"
            response = "Thank you for using the Audit Assistant. If you need any other audit services, feel free to start a new conversation."
            new_history.append({"role": "assistant", "content": response})
            return new_history, state, None
    
    # Final step or fallback
    else:
        if "new audit" in message.lower() or "start over" in message.lower():
            state = {
                "step": "greeting",
                "audit_type": None,
                "document_path": None,
                "analysis_results": None,
                "reasoning_results": None,
                "constraints": None
            }
            response = "👋 Hello! I'm your AI Audit Assistant. I can help you perform different types of audits on your financial documents. What type of audit would you like to perform today? (e.g., Financial Audit, Tax Audit, Compliance Audit)"
            new_history.append({"role": "assistant", "content": response})
            return new_history, state, None
        else:
            response = "Thank you for using the Audit Assistant. If you'd like to start a new audit, just let me know!"
            new_history.append({"role": "assistant", "content": response})
            return new_history, state, None

# Gradio interface
with gr.Blocks() as demo:
    gr.Markdown("# AI Audit Assistant")
    gr.Markdown("Upload your financial documents and chat with the assistant to perform various types of audits.")
    
    state = gr.State(None)
    
    with gr.Row():
        with gr.Column(scale=3):
            chatbot = gr.Chatbot(height=500, type="messages")
            msg = gr.Textbox(label="Your message", placeholder="Type your message here...")
            with gr.Row():
                submit = gr.Button("Send")
                clear = gr.Button("Clear")
        
        with gr.Column(scale=2):
            file_input = gr.File(label="Upload Documents", file_count="multiple")
            download_button = gr.File(label="Download Audit Report", interactive=True)
    
    # Handle interactions
    submit.click(
        audit_assistant, 
        inputs=[msg, chatbot, file_input, state], 
        outputs=[chatbot, state, download_button]
    )
    
    msg.submit(
        audit_assistant, 
        inputs=[msg, chatbot, file_input, state], 
        outputs=[chatbot, state, download_button]
    )
    
    # Clear function that properly resets the chat with the correct message format
    def clear_chat():
        return [], None, None
    
    clear.click(clear_chat, outputs=[chatbot, state, download_button])

# Launch the app
if __name__ == "__main__":
    demo.launch()




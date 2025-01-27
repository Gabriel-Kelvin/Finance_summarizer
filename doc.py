from azure.ai.documentintelligence import DocumentIntelligenceClient
from azure.core.credentials import AzureKeyCredential
from dotenv import load_dotenv
import os

load_dotenv()
Endpoint = os.getenv("DOCUMENTINTELLIGENCE_ENDPOINT")
Key = os.getenv("DOCUMENTINTELLIGENCE_API_KEY")
document_intelligence_client = DocumentIntelligenceClient(endpoint=Endpoint, credential=AzureKeyCredential(Key))
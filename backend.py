import json
import os
from typing import Optional
import msal
import requests
from fastapi import FastAPI, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from dotenv import load_dotenv
import base64

load_dotenv()  # Load environment variables from .env file

# ============================================================
# AAD Configuration
# ============================================================

AAD_CLIENT_ID = os.getenv("AAD_CLIENT_ID")
AAD_TENANT_ID = os.getenv("AAD_TENANT_ID")
AAD_CLIENT_SECRET = os.getenv("AAD_CLIENT_SECRET")
AAD_AUTHORITY = f"https://login.microsoftonline.com/{AAD_TENANT_ID}"
SCOPES = ["https://graph.microsoft.com/.default"]

if not all([AAD_CLIENT_ID, AAD_TENANT_ID, AAD_CLIENT_SECRET]):
    raise ValueError("缺少 AAD 環境變數: AAD_CLIENT_ID, AAD_TENANT_ID, AAD_CLIENT_SECRET")

# ============================================================
# FastAPI Setup
# ============================================================

app = FastAPI(title="Invoice Processing Backend")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# ============================================================
# MSAL Token Management
# ============================================================

def get_access_token():
    """從 AAD 獲取 access token (Client Credentials Flow)"""
    app_client = msal.ConfidentialClientApplication(
        AAD_CLIENT_ID,
        authority=AAD_AUTHORITY,
        client_credential=AAD_CLIENT_SECRET,
    )

    token_response = app_client.acquire_token_for_client(scopes=SCOPES)

    if "access_token" in token_response:
        return token_response["access_token"]
    else:
        raise HTTPException(
            status_code=401,
            detail=f"AAD authentication failed: {token_response.get('error_description')}"
        )

# ============================================================
# Microsoft Graph API Calls
# ============================================================

def get_graph_headers(user_token: Optional[str] = None):
    """取得 Graph API 請求頭"""
    if user_token:
        token = user_token
    else:
        token = get_access_token()
    return {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json",
    }

@app.get("/api/emails")
def get_emails(user_token: Optional[str] = None):
    """
    從 Outlook 獲取郵件
    返回郵件列表和附件信息
    user_token: 用戶的 AAD access token（可選，如果提供則使用用戶token，否則使用service account）
    """
    try:
        if not user_token:
            raise HTTPException(status_code=401, detail="User token required. Please log in first.")

        # 使用用戶 token（已登入的使用者）
        headers = get_graph_headers(user_token)

        # 獲取用戶信息
        user_response = requests.get(
            "https://graph.microsoft.com/v1.0/me",
            headers=headers
        )
        user_response.raise_for_status()
        user_data = user_response.json()
        user_id = user_data.get("id", "unknown")

        messages_url = f"https://graph.microsoft.com/v1.0/me/mailFolders/inbox/messages"

        messages_response = requests.get(
            f'{messages_url}',
            params={
                "$filter": "isRead eq false and contains(subject, 'Invoice')",
                "$select": "id,conversationId,subject,receivedDateTime,from,hasAttachments, bodyPreview",
                "$top": 50
            },
            headers=headers
        )
        messages_response.raise_for_status()
        messages_data = messages_response.json()
        emails = []
        for message in messages_data.get("value", []):
            # 先取得原始對話串
            conversation_id = message.get("conversationId")
            conversation_url = "https://graph.microsoft.com/v1.0/me/messages"
            conversation_response = requests.get(
                conversation_url,
                params={
                    "$filter": f"conversationId eq '{conversation_id}'"
                },
                headers=headers
            )
            conversation_response.raise_for_status()
            conversation_data = conversation_response.json()

        for convesation in conversation_data.get("value", []):
            # 獲取附件
            attachments = []
            if convesation.get("hasAttachments"):
                attachments_response = requests.get(
                    f"https://graph.microsoft.com/v1.0/me/messages/{convesation['id']}/attachments",
                    headers=headers
                )
                if attachments_response.status_code == 200:
                    attachments_data = attachments_response.json()
                    for attachment in attachments_data.get("value", []):
                        attachments.append({
                            "id": attachment.get("id"),
                            "name": attachment.get("name"),
                            "size": attachment.get("size"),
                        })

            email_entry = {
                "message_id": convesation["id"],
                "subject": convesation.get("subject"),
                "from": convesation.get("from", {}).get("emailAddress", {}).get("address"),
                "received_time": convesation.get("receivedDateTime"),
                "body_preview": convesation.get("bodyPreview"),
                "has_attachments": convesation.get("hasAttachments"),
                "attachments": attachments,
            }
            emails.append(email_entry)

        return {
            "user_id": user_id,
            "conversation data": conversation_data,
            "emails": emails,
        }

    except requests.exceptions.RequestException as e:
        raise HTTPException(status_code=500, detail=f"Graph API error: {str(e)}")
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@app.get("/api/attachment/{message_id}/{attachment_id}")
def get_attachment(message_id: str, attachment_id: str, user_token: Optional[str] = None):
    """
    下載郵件附件的內容
    """
    try:
        if not user_token:
            raise HTTPException(status_code=401, detail="User token required")

        headers = get_graph_headers(user_token)

        # 獲取附件
        attachment_response = requests.get(
            f"https://graph.microsoft.com/v1.0/me/messages/{message_id}/attachments/{attachment_id}",
            headers=headers
        )
        attachment_response.raise_for_status()
        attachment_data = attachment_response.json()

        # 如果是 itemAttachment，需要特殊處理
        if attachment_data.get("@odata.type") == "#microsoft.graph.itemAttachment":
            return {
                "error": "Item attachment not supported",
                "type": attachment_data.get("@odata.type")
            }

        # 獲取檔案內容（如果是 fileAttachment）
        if attachment_data.get("@odata.type") == "#microsoft.graph.fileAttachment":
            content_bytes = base64.b64decode(attachment_data.get("contentBytes", ""))
            return {
                "name": attachment_data.get("name"),
                "content": base64.b64encode(content_bytes).decode("utf-8"),
                "content_type": attachment_data.get("contentType"),
            }

        return {
            "name": attachment_data.get("name"),
            "size": attachment_data.get("size"),
        }

    except requests.exceptions.RequestException as e:
        raise HTTPException(status_code=500, detail=f"Graph API error: {str(e)}")
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))



@app.post("/api/analyze-invoice")
def analyze_invoice(
    message_id: str,
    attachment_id: str,
    attachment_name: str,
    user_token: Optional[str] = None
):
    """
    使用 Azure AI Document Intelligence 分析郵件附件中的發票
    """
    try:
        if not user_token:
            raise HTTPException(
                status_code=401,
                detail="User token required"
            )

        # ==========================================
        # 1. 從 Microsoft Graph 下載附件
        # ==========================================
        headers = get_graph_headers(user_token)

        attachment_response = requests.get(
            f"https://graph.microsoft.com/v1.0/me/messages/"
            f"{message_id}/attachments/{attachment_id}",
            headers=headers
        )

        attachment_response.raise_for_status()

        attachment_data = attachment_response.json()

        print("取得附件成功")

        # ==========================================
        # 2. 確認是 File Attachment
        # ==========================================
        if attachment_data.get("@odata.type") != "#microsoft.graph.fileAttachment":
            return {
                "file": attachment_name,
                "is_invoice": False,
                "error": "Not a file attachment"
            }

        # ==========================================
        # 3. 確認 PDF
        # ==========================================
        if not attachment_name.lower().endswith(".pdf"):
            return {
                "file": attachment_name,
                "is_invoice": False,
                "error": "Not a PDF file"
            }

        # ==========================================
        # 4. Decode Outlook attachment
        # ==========================================
        pdf_content = base64.b64decode(
            attachment_data.get("contentBytes", "")
        )

        # ==========================================
        # 5. Azure Document Intelligence
        # ==========================================
        from azure.ai.documentintelligence import DocumentIntelligenceClient
        from azure.core.credentials import AzureKeyCredential

        endpoint = os.getenv("AZURE_DOCUMENT_INTELLIGENCE_ENDPOINT")
        key = os.getenv("AZURE_DOCUMENT_INTELLIGENCE_KEY")

        if not endpoint or not key:
            raise HTTPException(
                status_code=500,
                detail="Azure Document Intelligence credentials not configured"
            )

        document_client = DocumentIntelligenceClient(
            endpoint=endpoint,
            credential=AzureKeyCredential(key)
        )

        poller = document_client.begin_analyze_document(
            "prebuilt-invoice",
            body=pdf_content
        )

        result = poller.result()

        print("Azure Document Intelligence analysis completed")

        # ==========================================
        # 6. Extract Invoice Fields
        # ==========================================
        invoices = []

        for document in result.documents:

            fields = document.fields

            invoice_data = {
                "invoice_number": None,
                "invoice_date": None,
                "vendor_name": None,
                "customer_name": None,
                "total_amount": None,
                "currency": None,
                "due_date": None,
            }

            if fields.get("InvoiceId"):
                invoice_data["invoice_number"] = (
                    fields["InvoiceId"].value_string
                )

            if fields.get("InvoiceDate"):
                invoice_data["invoice_date"] = (
                    fields["InvoiceDate"].value_date.isoformat()
                )

            if fields.get("VendorName"):
                invoice_data["vendor_name"] = (
                    fields["VendorName"].value_string
                )

            if fields.get("CustomerName"):
                invoice_data["customer_name"] = (
                    fields["CustomerName"].value_string
                )

            if fields.get("InvoiceTotal"):
                invoice_total = fields["InvoiceTotal"].value_currency

                if invoice_total:
                    invoice_data["total_amount"] = invoice_total.amount
                    invoice_data["currency"] = invoice_total.currency_code

            if fields.get("DueDate"):
                invoice_data["due_date"] = (
                    fields["DueDate"].value_date.isoformat()
                )

            invoices.append(invoice_data)

        # ==========================================
        # 7. Return result
        # ==========================================
        return {
            "file": attachment_name,
            "is_invoice": len(invoices) > 0,
            "invoices": invoices
        }

    except HTTPException:
        raise

    except Exception as e:
        print(f"Invoice analysis error: {str(e)}")

        raise HTTPException(
            status_code=500,
            detail=str(e)
        )

@app.get("/api/health")
def health():
    """健康檢查"""
    return {"status": "ok"}

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000)

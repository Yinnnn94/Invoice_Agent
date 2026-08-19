import json
import os
from typing import Optional
import msal
import requests
from fastapi import FastAPI, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from dotenv import load_dotenv
import base64
import tempfile
import pymysql


load_dotenv()  # Load environment variables from .env file

# ============================================================
# AAD Configuration
# ============================================================

AAD_CLIENT_ID = os.getenv("AAD_CLIENT_ID")
AAD_TENANT_ID = os.getenv("AAD_TENANT_ID")
AAD_CLIENT_SECRET = os.getenv("AAD_CLIENT_SECRET")
AAD_AUTHORITY = f"https://login.microsoftonline.com/{AAD_TENANT_ID}"
SCOPES = ["https://graph.microsoft.com/.default"]
SYSTEM_MAILBOX = os.getenv(
    "SYSTEM_MAILBOX",
    "your-system-mailbox@company.com"
)
DB_USER = os.getenv("DB_USER")
DB_PW = os.getenv("DB_PW")
DB_HOST = os.getenv("DB_HOST")
DB_NAME = os.getenv("DB_NAME")
DB_PORT = os.getenv("DB_PORT")

### DB必須的基本資料
db_settings = {
    "user": DB_USER,
    "password": DB_PW,
    "host": DB_HOST,
    "database": DB_NAME,
    "port": DB_PORT,
    "read_timeout": 5,
}
def get_db_connection(config):
    return pymysql.connect(**db_settings)


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

        # 過濾多個關鍵字：發票、Invoice、invoice、數電票
        keywords = ["發票", "Invoice", "invoice", "數電票"]
        filter_conditions = [f"contains(subject, '{kw}')" for kw in keywords]
        filter_query = f"isRead eq false and ({' or '.join(filter_conditions)})"

        messages_response = requests.get(
            f'{messages_url}',
            params={
                "$filter": filter_query,
<<<<<<< Updated upstream
                "$select": "id,conversationId,subject,receivedDateTime,from,hasAttachments, bodyPreview",
                "$top": 50
=======
                "$select": "id,conversationId,subject,receivedDateTime,from,hasAttachments, bodyPreview"
>>>>>>> Stashed changes
            },
            headers=headers
        )
        messages_response.raise_for_status()
        messages_data = messages_response.json()
        emails = []

        # 按 conversationId 分組，避免重複查詢同一個對話串
        conversations_cache = {}

        for message in messages_data.get("value", []):
            conversation_id = message.get("conversationId")

            # 如果還沒查過這個對話串，就查一次
            if conversation_id not in conversations_cache:
                try:
                    conv_response = requests.get(
                        f"https://graph.microsoft.com/v1.0/me/messages",
                        params={
                            "$filter": f"conversationId eq '{conversation_id}'",
                            "$select": "id,subject,receivedDateTime,from,hasAttachments,bodyPreview",
                            "$top": 100
                        },
                        headers=headers,
                        timeout=10
                    )
                    if conv_response.status_code == 200:
                        conversations_cache[conversation_id] = conv_response.json().get("value", [])
                    else:
                        conversations_cache[conversation_id] = []
                except:
                    conversations_cache[conversation_id] = []

            # 獲取該對話串的所有郵件
            conversation_messages = conversations_cache[conversation_id]
<<<<<<< Updated upstream

            # 只查當前郵件的附件（不是整個對話串）
            attachments = []
            if message.get("hasAttachments"):
                try:
                    attachments_response = requests.get(
                        f"https://graph.microsoft.com/v1.0/me/messages/{message['id']}/attachments",
=======
            # 只查當前郵件的附件（不是整個對話串）
        print(conversation_messages)
        for conversation in conversation_messages:
            attachments = []
            if conversation.get("hasAttachments"):
                try:
                    attachments_response = requests.get(
                        f"https://graph.microsoft.com/v1.0/me/messages/{conversation.get('id')}/attachments",
>>>>>>> Stashed changes
                        headers=headers,
                        timeout=5
                    )
                    if attachments_response.status_code == 200:
                        attachments_data = attachments_response.json()
                        for attachment in attachments_data.get("value", []):
                            attachments.append({
                                "id": attachment.get("id"),
                                "name": attachment.get("name"),
                                "size": attachment.get("size"),
                            })
                except:
                    pass

            email_entry = {
<<<<<<< Updated upstream
                "message_id": message["id"],
                "subject": message.get("subject"),
                "from": message.get("from", {}).get("emailAddress", {}).get("address"),
                "received_time": message.get("receivedDateTime"),
                "body_preview": message.get("bodyPreview"),
                "has_attachments": message.get("hasAttachments"),
=======
                "subject": conversation.get("subject"),
                "from": conversation.get("from", {}).get("emailAddress", {}).get("address"),
                "received_time": conversation.get("receivedDateTime"),
                "body_preview": conversation.get("bodyPreview"),
                "has_attachments": conversation.get("hasAttachments"),
>>>>>>> Stashed changes
                "attachments": attachments,
                "conversation_id": conversation_id,
                "conversation_messages": [
                    {
                        "id": m["id"],
                        "subject": m.get("subject"),
                        "received_time": m.get("receivedDateTime"),
                        "from": m.get("from", {}).get("emailAddress", {}).get("address"),
                    }
                    for m in conversation_messages
                ],
            }
            emails.append(email_entry)

        return {
            "user_id": user_id,
            "emails": emails
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

@app.get("/api/conversation/{message_id}")
def get_conversation(message_id: str, user_token: Optional[str] = None):
    """
    取得特定郵件所屬的完整對話串
    message_id: 郵件 ID
    """
    try:
        if not user_token:
            raise HTTPException(status_code=401, detail="User token required")

        headers = get_graph_headers(user_token)

        # 1. 先取得該郵件的基本信息和 conversationId
        message_response = requests.get(
            f"https://graph.microsoft.com/v1.0/me/messages/{message_id}",
            params={
                "$select": "id,conversationId,subject,receivedDateTime,from,hasAttachments,bodyPreview"
            },
            headers=headers
        )
        message_response.raise_for_status()
        message_data = message_response.json()
        conversation_id = message_data.get("conversationId")

        # 2. 取得整個對話串的所有郵件
        conversation_response = requests.get(
            f"https://graph.microsoft.com/v1.0/me/messages",
            params={
                "$filter": f"conversationId eq '{conversation_id}'",
                "$select": "id,subject,receivedDateTime,from,hasAttachments,bodyPreview",
                "$orderby": "receivedDateTime asc",
                "$top": 100
            },
            headers=headers
        )
        conversation_response.raise_for_status()
        conversation_data = conversation_response.json()

        # 3. 逐一取附件
        conversation_messages = []
        for msg in conversation_data.get("value", []):
            attachments = []
            if msg.get("hasAttachments"):
                try:
                    att_response = requests.get(
                        f"https://graph.microsoft.com/v1.0/me/messages/{msg['id']}/attachments",
                        headers=headers,
                        timeout=5
                    )
                    if att_response.status_code == 200:
                        att_data = att_response.json()
                        for att in att_data.get("value", []):
                            attachments.append({
                                "id": att.get("id"),
                                "name": att.get("name"),
                                "size": att.get("size"),
                            })
                except:
                    pass

            conversation_messages.append({
                "message_id": msg["id"],
                "subject": msg.get("subject"),
                "from": msg.get("from", {}).get("emailAddress", {}).get("address"),
                "received_time": msg.get("receivedDateTime"),
                "body_preview": msg.get("bodyPreview"),
                "has_attachments": msg.get("hasAttachments"),
                "attachments": attachments,
            })

        return {
            "conversation_id": conversation_id,
            "messages": conversation_messages,
        }

    except requests.exceptions.RequestException as e:
        raise HTTPException(status_code=500, detail=f"Graph API error: {str(e)}")
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@app.post("/api/mark-processed")
def mark_processed(message_id: str, invoice_number: str, amount: str = "", user_token: Optional[str] = None):
    """
    標記發票已處理，移動郵件到歸檔資料夾
    """
    try:
        if not user_token:
            raise HTTPException(status_code=401, detail="User token required")

        headers = get_graph_headers(user_token)

        # 1. 先建立或取得「已處理發票」資料夾
        folders_response = requests.get(
            "https://graph.microsoft.com/v1.0/me/mailFolders",
            params={"$filter": "displayName eq '已處理發票'"},
            headers=headers
        )
        folders_response.raise_for_status()
        folders = folders_response.json().get("value", [])

        if folders:
            folder_id = folders[0]["id"]
        else:
            # 建立新資料夾
            create_response = requests.post(
                "https://graph.microsoft.com/v1.0/me/mailFolders",
                json={"displayName": "已處理發票"},
                headers=headers
            )
            create_response.raise_for_status()
            folder_id = create_response.json()["id"]

        # 2. 移動郵件到該資料夾
        move_response = requests.post(
            f"https://graph.microsoft.com/v1.0/me/messages/{message_id}/move",
            json={"destinationId": folder_id},
            headers=headers
        )
        move_response.raise_for_status()

        return {
            "status": "success",
            "message": f"✅ 發票 {invoice_number} (金額: {amount}) 已標記為已處理，郵件已移動到「已處理發票」資料夾",
            "message_id": message_id,
            "folder_id": folder_id,
        }

    except requests.exceptions.RequestException as e:
        raise HTTPException(status_code=500, detail=f"Graph API error: {str(e)}")
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

<<<<<<< Updated upstream
=======


# ============================================================
# Database
# ============================================================
@app.get("/api/reconciliation-status")
def get_reconciliation_status(invoice_number: str):
    """
    查詢發票目前的對帳狀態。
    TODO: 修改 SQL statement 與 DB connection。
    """

    try:
        # ====================================================
        # TODO: 改成你自己的 DB connection
        # ====================================================
        conn = get_db_connection()

        cursor = conn.cursor()

        # ====================================================
        # TODO: 修改這裡的 SQL
        # ====================================================
        sql = """
            SELECT
                invoice_number,
                result
            FROM n8n_result_from_process_time
            WHERE invoice_number = %s
            and is_latest = 1
        """

        cursor.execute(sql, (invoice_number,))
        row = cursor.fetchone()

        cursor.close()
        conn.close()

        if not row:
            return {
                "status": "not_found",
                "invoice_number": invoice_number,
                "can_auto_process": False,
            }

        # 如果你的 DB 回傳 tuple
        db_invoice_number = row[0]
        db_status = row[1]

        # ====================================================
        # TODO: 根據你們實際 status 修改
        # ====================================================
        can_auto_process = db_status in [
            "ERP_NO_Data",
            "R049_No_Data",
            "Reconciliation Fail",
        ]

        return {
            "status": "success",
            "invoice_number": db_invoice_number,
            "reconciliation_status": db_status,
            "can_auto_process": can_auto_process,
        }

    except Exception as e:
        return {
            "status": "error",
            "error": str(e),
            "can_auto_process": False,
        }


@app.get("/api/download-invoice")
def download_invoice_attachment(
    message_id: str,
    attachment_id: str,
    attachment_name: str,
    user_token: str,
):
    """
    從 Microsoft Graph 下載指定附件，
    暫存到 backend 的 temporary directory。
    """

    try:
        if not user_token:
            return {
                "status": "error",
                "error": "Missing user token",
            }

        url = (
            "https://graph.microsoft.com/v1.0/"
            f"me/messages/{message_id}/attachments/{attachment_id}"
        )

        headers = {
            "Authorization": f"Bearer {user_token}",
        }

        response = requests.get(
            url,
            headers=headers,
            timeout=30,
        )

        response.raise_for_status()

        attachment_data = response.json()

        # Microsoft Graph fileAttachment
        if attachment_data.get("@odata.type") != "#microsoft.graph.fileAttachment":
            return {
                "status": "error",
                "error": "Attachment is not a file attachment",
            }

        content_bytes = attachment_data.get("contentBytes")

        if not content_bytes:
            return {
                "status": "error",
                "error": "Attachment does not contain contentBytes",
            }

        file_data = base64.b64decode(content_bytes)

        # 暫存檔
        temp_dir = tempfile.gettempdir()

        safe_filename = os.path.basename(attachment_name)

        file_path = os.path.join(
            temp_dir,
            safe_filename,
        )

        with open(file_path, "wb") as f:
            f.write(file_data)

        return {
            "status": "success",
            "file_path": file_path,
            "file_name": safe_filename,
            "size": len(file_data),
        }

    except requests.exceptions.RequestException as e:
        return {
            "status": "error",
            "error": f"Microsoft Graph error: {str(e)}",
        }

    except Exception as e:
        return {
            "status": "error",
            "error": str(e),
        }


@app.post("/api/send-invoice")
def send_invoice_to_system_mailbox(
    file_path: str,
    file_name: str,
    invoice_number: str,
    user_token: str,
):
    """
    將確認後的發票附件寄到系統信箱。
    """

    try:
        if not user_token:
            return {
                "status": "error",
                "error": "Missing user token",
            }

        if not os.path.exists(file_path):
            return {
                "status": "error",
                "error": f"File not found: {file_path}",
            }

        # 讀取附件
        with open(file_path, "rb") as f:
            file_bytes = f.read()

        encoded_file = base64.b64encode(file_bytes).decode("utf-8")

        # Microsoft Graph Send Mail
        url = (
            "https://graph.microsoft.com/v1.0/"
            "me/sendMail"
        )

        headers = {
            "Authorization": f"Bearer {user_token}",
            "Content-Type": "application/json",
        }

        mail_payload = {
            "message": {
                "subject": f"Invoice - {invoice_number}",
                "body": {
                    "contentType": "Text",
                    "content": (
                        "This invoice has been identified and "
                        "processed by the AI Invoice Agent."
                    ),
                },
                "toRecipients": [
                    {
                        "emailAddress": {
                            "address": SYSTEM_MAILBOX
                        }
                    }
                ],
                "attachments": [
                    {
                        "@odata.type": "#microsoft.graph.fileAttachment",
                        "name": file_name,
                        "contentType": "application/pdf",
                        "contentBytes": encoded_file,
                    }
                ],
            }
        }

        response = requests.post(
            url,
            headers=headers,
            json=mail_payload,
            timeout=30,
        )

        response.raise_for_status()

        # Graph sendMail 成功通常會回 202
        return {
            "status": "success",
            "invoice_number": invoice_number,
            "file_name": file_name,
            "sent_to": SYSTEM_MAILBOX,
        }

    except requests.exceptions.RequestException as e:
        return {
            "status": "error",
            "error": f"Microsoft Graph error: {str(e)}",
        }

    except Exception as e:
        return {
            "status": "error",
            "error": str(e),
        }

    finally:
        # 清除暫存檔
        try:
            if file_path and os.path.exists(file_path):
                os.remove(file_path)
        except Exception:
            pass



>>>>>>> Stashed changes
@app.get("/api/health")
def health():
    """健康檢查"""
    return {"status": "ok"}

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000)
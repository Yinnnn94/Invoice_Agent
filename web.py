import json
from typing import Optional
import streamlit as st
import requests
from openai import OpenAI
import msal

# ============================================================
# Configuration
# ============================================================

st.set_page_config(
    page_title="AI Invoice Processing Agent",
    page_icon="🧾",
    layout="wide",
)

BACKEND_URL = st.secrets.get("BACKEND_URL", "http://localhost:8000")
AAD_CLIENT_ID = st.secrets.get("AAD_CLIENT_ID")
AAD_CLIENT_SECRET = st.secrets.get("AAD_CLIENT_SECRET")
AAD_TENANT_ID = st.secrets.get("AAD_TENANT_ID")
AAD_AUTHORITY = f"https://login.microsoftonline.com/{AAD_TENANT_ID}"
REDIRECT_URI = "http://localhost:8501/"

client = OpenAI(api_key=st.secrets.get("OPENAI_API_KEY"))

# ============================================================
# Session State
# ============================================================

if "user_token" not in st.session_state:
    st.session_state.user_token = None

if "user_info" not in st.session_state:
    st.session_state.user_info = None

if "chat_messages" not in st.session_state:
    st.session_state.chat_messages = []

# ============================================================
# AAD Authentication
# ============================================================

def get_msal_app():
    """建立 MSAL 應用程式"""
    return msal.ConfidentialClientApplication(
        AAD_CLIENT_ID,
        client_credential=AAD_CLIENT_SECRET,
        authority=AAD_AUTHORITY,
    )

def login_with_microsoft():
    """用 Microsoft 帳號登入"""
    app = get_msal_app()

    # 生成授權碼 URL
    auth_url = app.get_authorization_request_url(
        scopes=["Mail.Read", "User.Read"],
        redirect_uri=REDIRECT_URI,
    )

    st.write("請點擊下方連結用 Microsoft 帳號登入:")
    st.link_button("🔐 用 Microsoft 登入", auth_url)

def handle_oauth_callback(auth_code: str):
    """處理 OAuth 回呼"""
    app = get_msal_app()

    try:
        token_response = app.acquire_token_by_authorization_code(
            auth_code,
            scopes=["Mail.Read", "User.Read"],
            redirect_uri=REDIRECT_URI,
        )

        if "access_token" in token_response:
            st.session_state.user_token = token_response["access_token"]
            st.session_state.user_info = token_response.get("id_token_claims", {})
            st.success(f"✅ 登入成功！歡迎 {st.session_state.user_info.get('name', 'User')}")
            st.rerun()
        else:
            st.error(f"❌ 登入失敗: {token_response.get('error_description')}")
    except Exception as e:
        st.error(f"❌ 認證錯誤: {str(e)}")

def logout():
    """登出"""
    st.session_state.user_token = None
    st.session_state.user_info = None
    st.success("✅ 已登出")
    st.rerun()

# ============================================================
# Backend API Functions
# ============================================================

def get_emails_from_outlook(user_token: Optional[str] = None):
    """從 Outlook 取得郵件"""
    try:
        params = {}
        if user_token:
            params["user_token"] = user_token

        response = requests.get(
            f"{BACKEND_URL}/api/emails",
            params=params,
            timeout=30,
        )
        response.raise_for_status()
        return response.json()
    except requests.exceptions.RequestException as e:
        st.error(f"無法連接後端：{str(e)}")
        return None

def get_attachment_content(message_id: str, attachment_id: str, attachment_name: str, user_token: Optional[str] = None):
    """下載附件內容"""
    try:
        params = {
            "message_id": message_id,
            "attachment_id": attachment_id,
        }
        if user_token:
            params["user_token"] = user_token

        response = requests.get(
            f"{BACKEND_URL}/api/attachment/{message_id}/{attachment_id}",
            params=params,
            timeout=10,
        )
        response.raise_for_status()
        return response.json()
    except requests.exceptions.RequestException as e:
        st.error(f"無法下載附件：{str(e)}")
        return None

def analyze_invoice_attachment(message_id: str, attachment_id: str, attachment_name: str, user_token: Optional[str] = None):
    """分析發票附件"""
    try:
        params = {
            "message_id": message_id,
            "attachment_id": attachment_id,
            "attachment_name": attachment_name,
        }
        if user_token:
            params["user_token"] = user_token

        response = requests.post(
            f"{BACKEND_URL}/api/analyze-invoice",
            params=params,
            timeout=30,
        )
        response.raise_for_status()
        return response.json()
    except requests.exceptions.RequestException as e:
        st.error(f"無法分析發票：{str(e)}")
        return None

def get_reconciliation_status(
    invoice_number: str,
):
    try:
        response = requests.get(
            f"{BACKEND_URL}/api/reconciliation-status",
            params={
                "invoice_number": invoice_number,
            },
            timeout=10,
        )

        response.raise_for_status()
        return response.json()

    except requests.exceptions.RequestException as e:
        return {
            "status": "error",
            "error": str(e),
        }
    
def download_invoice_attachment(
    message_id: str,
    attachment_id: str,
    attachment_name: str,
    user_token: Optional[str] = None,
):
    try:
        response = requests.get(
            f"{BACKEND_URL}/api/download-invoice",
            params={
                "message_id": message_id,
                "attachment_id": attachment_id,
                "attachment_name": attachment_name,
                "user_token": user_token,
            },
            timeout=30,
        )

        response.raise_for_status()
        return response.json()

    except requests.exceptions.RequestException as e:
        return {
            "status": "error",
            "error": str(e),
        }

def send_invoice_to_system_mailbox(
    file_path: str,
    file_name: str,
    invoice_number: str,
    user_token: Optional[str] = None,
):
    try:
        response = requests.post(
            f"{BACKEND_URL}/api/send-invoice",
            params={
                "file_path": file_path,
                "file_name": file_name,
                "invoice_number": invoice_number,
                "user_token": user_token,
            },
            timeout=30,
        )

        response.raise_for_status()
        return response.json()

    except requests.exceptions.RequestException as e:
        return {
            "status": "error",
            "error": str(e),
        }


# ============================================================
# Tool definitions for the LLM
# ============================================================

TOOLS = [
    {
        "type": "function",
        "function": {
            "name": "get_emails",
            "description": "從 Outlook 取得所有郵件和附件信息",
            "parameters": {
                "type": "object",
                "properties": {},
                "required": [],
            },
        },
    },
    {
        "type": "function",
        "function": {
            "name": "analyze_invoice",
            "description": "分析郵件中的 PDF 發票附件，提取發票號碼、金額等信息",
            "parameters": {
                "type": "object",
                "properties": {
                    "message_id": {
                        "type": "string",
                        "description": "郵件 ID",
                    },
                    "attachment_id": {
                        "type": "string",
                        "description": "附件 ID",
                    },
                    "attachment_name": {
                        "type": "string",
                        "description": "附件名稱",
                    },
                },
                "required": ["message_id", "attachment_id", "attachment_name"],
            },
        },
    },
    {
        "type": "function",
        "function": {
            "name": "mark_invoice_processed",
            "description": "標記發票已處理，自動移動郵件到歸檔資料夾",
            "parameters": {
                "type": "object",
                "properties": {
                    "message_id": {
                        "type": "string",
                        "description": "郵件 ID",
                    },
                    "invoice_number": {
                        "type": "string",
                        "description": "發票號碼",
                    },
                    "amount": {
                        "type": "string",
                        "description": "發票金額",
                    },
                },
                "required": ["message_id", "invoice_number"],
            },
        },
    },
    {
    "type": "function",
    "function": {
        "name": "get_reconciliation_status",
        "description": (
            "查詢指定發票目前在發票對帳資料庫中的狀態。"
            "在處理更新版發票之前，必須先查詢此狀態。"
        ),
        "parameters": {
            "type": "object",
            "properties": {
                "invoice_number": {
                    "type": "string",
                    "description": "發票號碼",
                }
            },
            "required": ["invoice_number"],
        },
    },
},
{
    "type": "function",
    "function": {
        "name": "download_invoice_attachment",
        "description": (
            "從 Outlook 下載指定的發票附件。"
            "只能下載已確認為有效發票的附件。"
        ),
        "parameters": {
            "type": "object",
            "properties": {
                "message_id": {
                    "type": "string",
                    "description": "Email message ID",
                },
                "attachment_id": {
                    "type": "string",
                    "description": "Attachment ID",
                },
                "attachment_name": {
                    "type": "string",
                    "description": "Attachment file name",
                },
            },
            "required": [
                "message_id",
                "attachment_id",
                "attachment_name",
            ],
        },
    },
},
{
    "type": "function",
    "function": {
        "name": "send_invoice_to_system_mailbox",
        "description": (
            "將已確認的發票附件寄送到指定的系統信箱，"
            "供後續發票對帳流程使用。"
        ),
        "parameters": {
            "type": "object",
            "properties": {
                "file_path": {
                    "type": "string",
                    "description": "下載後的暫存檔案路徑",
                },
                "file_name": {
                    "type": "string",
                    "description": "發票檔案名稱",
                },
                "invoice_number": {
                    "type": "string",
                    "description": "發票號碼",
                },
            },
            "required": [
                "file_path",
                "file_name",
                "invoice_number",
            ],
        },
    },
},
]

# ============================================================
# Tool executor
# ============================================================

def mark_invoice_processed(message_id: str, invoice_number: str, amount: str = "", user_token: Optional[str] = None):
    """標記發票已處理，移動到歸檔資料夾"""
    try:
        response = requests.post(
            f"{BACKEND_URL}/api/mark-processed",
            params={
                "message_id": message_id,
                "invoice_number": invoice_number,
                "amount": amount,
                "user_token": user_token,
            },
            timeout=10,
        )
        response.raise_for_status()
        return response.json()
    except requests.exceptions.RequestException as e:
        return {"error": f"無法標記為已處理：{str(e)}"}

def execute_tool(name: str, arguments: dict) -> str:
    try:
        if name == "get_emails":
            result = get_emails_from_outlook(st.session_state.user_token)
            return json.dumps(result, ensure_ascii=False)

        if name == "analyze_invoice":
            result = analyze_invoice_attachment(
                arguments["message_id"],
                arguments["attachment_id"],
                arguments["attachment_name"],
                st.session_state.user_token,
            )
            return json.dumps(result, ensure_ascii=False)

        if name == "mark_invoice_processed":
            result = mark_invoice_processed(
                arguments["message_id"],
                arguments["invoice_number"],
                arguments.get("amount", ""),
                st.session_state.user_token,
            )
            return json.dumps(result, ensure_ascii=False)
        
        if name == "get_reconciliation_status":
            result = get_reconciliation_status(
                arguments["invoice_number"]
            )
            return json.dumps(
                result,
                ensure_ascii=False
            )
        
        if name == "download_invoice_attachment":
            result = download_invoice_attachment(
                arguments["message_id"],
                arguments["attachment_id"],
                arguments["attachment_name"],
                st.session_state.user_token,
            )

            return json.dumps(
                result,
                ensure_ascii=False
            )
        
        if name == "send_invoice_to_system_mailbox":
            result = send_invoice_to_system_mailbox(
                arguments["file_path"],
                arguments["file_name"],
                arguments["invoice_number"],
                st.session_state.user_token,
            )

            return json.dumps(
                result,
                ensure_ascii=False
            )

        return json.dumps({"error": f"Unknown tool: {name}"})
    except Exception as e:
        return json.dumps({"error": str(e)})

# ============================================================
# Agent
# ============================================================

SYSTEM_PROMPT = """
        You are an AI Invoice Processing Agent.

        Your job is to identify the latest valid invoice from an email
        conversation in Outlook.

        You have access to tools:

        1. get_emails
        Retrieve all unread invoice-related emails from Outlook inbox.

        2. analyze_invoice(message_id, attachment_id, attachment_name)
        Analyze a PDF invoice attachment to extract invoice information.

        3. get_conversation(message_id)
        Get the complete conversation thread for a specific message to see
        the full history and understand if the invoice was already processed.

        Important rules for invoice identification:

        - Do not assume that the first invoice is the valid invoice.
        - Inspect the entire email thread using get_conversation if needed.
        - Consider email timestamps - LATEST invoice is usually the valid one.

        CRITICAL: Identify invoice updates/corrections:
        - Same invoice number + different version = UPDATE (use latest)
        - Same supplier + similar invoice number = LIKELY SAME INVOICE
        - Look for keywords: "更新", "修正", "更正", "corrected", "updated", "revised"
        - Compare amounts: if amount differs, it's likely an update
        - Compare line items: if products/quantities change, it's an update
        - Timeline: if supplier sends new invoice days after first one, treat as update

        Decision logic:
        - If multiple versions of same invoice exist, select LATEST version
        - If unsure if it's an update, check the conversation context
        - Only select the FINAL, MOST RECENT valid invoice
        When a valid invoice has been identified:

        1. Analyze the invoice and identify the invoice number.

        2. ALWAYS call get_reconciliation_status(invoice_number)
        before taking any action.

        3. If the reconciliation status indicates that the invoice
        can be automatically processed:

        a. Download the selected invoice attachment using
            download_invoice_attachment.

        b. Send the downloaded invoice attachment to the
            system mailbox using send_invoice_to_system_mailbox.

        4. If sending the invoice succeeds:
        Report the successful submission.

        5. If sending the invoice fails:
        DO NOT mark the invoice as processed.
        DO NOT move the original email.
        Report the failure.

        6. If the reconciliation status indicates that the invoice
        has already been successfully processed:
        DO NOT automatically replace or resend the invoice.
        Return "Review Required".

        IMPORTANT:
        The reconciliation database is the source of truth for the
        current invoice processing state.

        The LLM is responsible for understanding email context and
        selecting the appropriate tools.

        The backend is responsible for enforcing business rules.

        - Only analyze PDF attachments (ignore images, documents, etc).
        - If a later email explicitly says an earlier invoice should be disregarded, exclude it.
        - If multiple different invoices exist, report all but indicate which is primary.
        - If there is uncertainty, mark the result as "Review Required".

        User can ask in natural language:
        - "Show me the second email from the 3 emails"
        - "Which of these emails has invoice attachments?"
        - "Check the complete conversation for message X"
        - "Find all PDF attachments in message Y"

        You should understand the user's intent and use the appropriate tools.
        When user refers to email position (first, second, third), use the order
        from the get_emails result.

        IMPORTANT: When examining any specific message:
        - ALWAYS use get_conversation(message_id) to read the ENTIRE conversation thread
        - Never analyze a single message in isolation
        - You need to see the full context to understand invoice updates/corrections
        - Check all messages in the thread to identify which is the latest valid version

        After identifying the valid invoice:
        1. Get the complete conversation thread using get_conversation
        2. Analyze all versions in the thread
        3. Select the LATEST valid invoice
        4. Mark it as processed (moves to archive folder automatically)
        5. Report the decision with complete reasoning from the thread context

        At the end, provide:

        1. Selected invoice file name
        2. Invoice number
        3. Amount
        4. Supplier
        5. Processing status
        6. Reason for the decision

        Possible statuses:

        - Completed
        - Review Required
        - Failed
"""

def run_agent(user_request: str):
    """執行 AI Agent，保留多輪對話"""

    # 如果是第一次對話，建立 system prompt
    if not st.session_state.chat_messages:
        st.session_state.chat_messages.append(
            {
                "role": "system",
                "content": SYSTEM_PROMPT,
            }
        )

    # 加入使用者訊息
    st.session_state.chat_messages.append(
        {
            "role": "user",
            "content": user_request,
        }
    )

    execution_trace = []

    while True:
        response = client.chat.completions.create(
            model="gpt-5.6-luna",
            messages=st.session_state.chat_messages,
            tools=TOOLS,
            tool_choice="auto",
        )

        message = response.choices[0].message

        # 將 assistant message 存進 conversation
        st.session_state.chat_messages.append(
            message.model_dump(exclude_none=True)
        )

        # 沒有 tool call → Agent 最終回答
        if not message.tool_calls:
            return message.content, execution_trace

        # 執行 tools
        for tool_call in message.tool_calls:
            tool_name = tool_call.function.name
            arguments = json.loads(tool_call.function.arguments)

            execution_trace.append(
                {
                    "tool": tool_name,
                    "arguments": arguments,
                }
            )

            result = execute_tool(tool_name, arguments)

            st.session_state.chat_messages.append(
                {
                    "role": "tool",
                    "tool_call_id": tool_call.id,
                    "content": result,
                }
            )
# ============================================================
# Streamlit UI
# ============================================================

st.title("🧾 AI Invoice Processing Agent")
st.caption("使用 AAD 認證的 Outlook 郵件和發票處理")

st.divider()

# 檢查 URL 參數中的授權碼
query_params = st.query_params
if "code" in query_params and "processed_code" not in st.session_state:
    auth_code = query_params["code"]
    st.session_state.processed_code = True
    handle_oauth_callback(auth_code)
    # 清除 URL 參數
    st.query_params.clear()

# ============================================================
# 登入/登出部分
# ============================================================

col1, col2, col3 = st.columns([2, 1, 1])

with col1:
    if st.session_state.user_token:
        user_name = st.session_state.user_info.get("name", "User")
        st.success(f"✅ 已登入: {user_name}")
    else:
        st.warning("⚠️ 未登入")

with col3:
    if st.session_state.user_token:
        if st.button("🚪 登出"):
            logout()
    else:
        st.write("")

st.divider()
if st.button("🔄 Refresh Emails"):
    st.rerun()
# ============================================================
# 未讀發票郵件摘要（登入後顯示）
# ============================================================

if st.session_state.user_token:
    st.subheader("📬 未讀發票郵件")

    with st.spinner("載入郵件中..."):
        emails_data = get_emails_from_outlook(st.session_state.user_token)

    if emails_data and "emails" in emails_data:
        emails_list = emails_data.get("emails", [])

        if emails_list:
            st.success(f"✉️ 找到 **{len(emails_list)}** 封未讀發票郵件")

            # 顯示郵件列表
            for idx, email in enumerate(emails_list, 1):
                with st.expander(f"📧 [{idx}] {email['subject']}", expanded=(idx == 1)):
                    col1, col2 = st.columns([3, 1])
                    with col1:
                        st.caption(f"**寄件者**: {email['from']}")
                        st.caption(f"**時間**: {email['received_time']}")
                        if email['body_preview']:
                            st.caption(f"**預覽**: {email['body_preview'][:100]}...")

                    with col2:
                        if email['has_attachments']:
                            st.markdown(f"📎 **{len(email['attachments'])} 個附件**")
                            for att in email['attachments']:
                                st.caption(f"• {att['name']}")
                        else:
                            st.caption("無附件")
        else:
            st.info("✅ 目前沒有未讀的發票郵件")
    else:
        st.error("❌ 無法載入郵件")

st.divider()

# ============================================================
# 未讀發票郵件摘要（登入後顯示）
# ============================================================

if st.session_state.user_token:
    st.subheader("📬 未讀發票郵件")

    with st.spinner("載入郵件中..."):
        emails_data = get_emails_from_outlook(st.session_state.user_token)

    if emails_data and "emails" in emails_data:
        emails_list = emails_data.get("emails", [])

        if emails_list:
            st.success(f"✉️ 找到 **{len(emails_list)}** 封未讀發票郵件")

            # 顯示郵件列表
            for idx, email in enumerate(emails_list, 1):
                with st.expander(f"📧 [{idx}] {email['subject']}", expanded=(idx == 1)):
                    col1, col2 = st.columns([3, 1])
                    with col1:
                        st.caption(f"**寄件者**: {email['from']}")
                        st.caption(f"**時間**: {email['received_time']}")
                        if email['body_preview']:
                            st.caption(f"**預覽**: {email['body_preview'][:100]}...")

                    with col2:
                        if email['has_attachments']:
                            st.markdown(f"📎 **{len(email['attachments'])} 個附件**")
                            for att in email['attachments']:
                                st.caption(f"• {att['name']}")
                        else:
                            st.caption("無附件")
        else:
            st.info("✅ 目前沒有未讀的發票郵件")
    else:
        st.error("❌ 無法載入郵件")

st.divider()

# Dashboard
st.subheader("📊 系統狀態")

col1, col2, col3 = st.columns(3)
with col1:
    try:
        backend_health = requests.get(f"{BACKEND_URL}/api/health", timeout=5).status_code == 200
    except:
        backend_health = False

    st.metric(
        "後端連接",
        "✅ 正常" if backend_health else "❌ 失敗",
    )
with col2:
    st.metric(
        "認證狀態",
        "✅ 已登入" if st.session_state.user_token else "❌ 未登入",
    )
with col3:
    st.metric(
        "認證方式",
        "AAD (MSAL)",
    )

st.divider()

# ============================================================
# 登入提示
# ============================================================

if not st.session_state.user_token:
    st.info("請用 Microsoft 帳號登入來存取你的 Outlook 郵件")
    login_with_microsoft()
else:
    # ============================================================
    # Agent 請求
    # ============================================================

    # st.subheader("🤖 Agent 請求")
    # user_request = st.chat_input(
    #     "請輸入要求:"
    # )

    # if st.button("🚀 執行 Agent", type="primary"):
    #     if not backend_health:
    #         st.error("❌ 無法連接後端。請確保後端正在運行: python backend.py")
    #     else:
    #         st.divider()
    #         st.subheader("🔎 Agent 執行過程")
    #         with st.spinner("Agent 正在處理..."):
    #             try:
    #                 answer, trace = run_agent(user_request)

    #                 # Execution trace
    #                 for index, step in enumerate(trace, start=1):
    #                     tool = step["tool"]
    #                     arguments = step["arguments"]

    #                     st.write(f"**步驟 {index}:** `{tool}`")
    #                     if arguments:
    #                         st.caption(f"參數: {json.dumps(arguments, ensure_ascii=False)}")

    #                 # Final result
    #                 st.divider()
    #                 st.subheader("📋 Agent 結果")
    #                 st.markdown(answer)

    #             except Exception as e:
    #                 st.error(f"❌ 執行失敗: {str(e)}")


    st.subheader("🤖 Invoice Agent")

    # 顯示歷史對話
    for message in st.session_state.chat_messages:

        # 不顯示 system prompt
        if message["role"] == "system":
            continue

        # 不直接顯示 tool message
        if message["role"] == "tool":
            continue

        if message["role"] == "user":
            with st.chat_message("user"):
                st.markdown(message["content"])

        elif message["role"] == "assistant":
            with st.chat_message("assistant"):
                st.markdown(message.get("content", ""))


    # Chat input
    user_request = st.chat_input(
        "例如：幫我找最新的有效發票"
    )

    if user_request:

        # 顯示使用者訊息
        with st.chat_message("user"):
            st.markdown(user_request)

        # Agent 回覆
        with st.chat_message("assistant"):
            with st.spinner("Agent 正在處理..."):
                try:
                    answer, trace = run_agent(user_request)

                    st.markdown(answer)

                    # 顯示 Agent tool execution
                    if trace:
                        with st.expander("🔎 Agent Execution Trace"):
                            for index, step in enumerate(trace, start=1):
                                tool = step["tool"]
                                arguments = step["arguments"]

                                st.write(
                                    f"**Step {index}:** `{tool}`"
                                )

                                if arguments:
                                    st.caption(
                                        json.dumps(
                                            arguments,
                                            ensure_ascii=False
                                        )
                                    )

                except Exception as e:
                    st.error(f"❌ 執行失敗: {str(e)}")

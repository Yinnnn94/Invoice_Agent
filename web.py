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
            timeout=10,
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
]

# ============================================================
# Tool executor
# ============================================================

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

You have access to two tools:

1. get_emails
   Retrieve all emails and attachments from Outlook inbox.

2. analyze_invoice(message_id, attachment_id, attachment_name)
   Analyze a PDF invoice attachment to extract invoice information.

Important rules:

- Do not assume that the first invoice is the valid invoice.
- Inspect the entire email thread.
- Consider email timestamps.
- Identify corrected or replaced invoices.
- Only analyze PDF attachments (ignore images, documents, etc).
- If a later email explicitly says that an earlier invoice should
  be disregarded, do not select the earlier invoice.
- If multiple invoice versions exist, select the latest valid version.
- If there is uncertainty, mark the result as "Review Required".

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
    """執行 AI Agent"""
    messages = [
        {
            "role": "system",
            "content": SYSTEM_PROMPT,
        },
        {
            "role": "user",
            "content": user_request,
        },
    ]

    execution_trace = []

    while True:
        response = client.chat.completions.create(
            model="gpt-5.6-luna",
            messages=messages,
            tools=TOOLS,
            tool_choice="auto",
        )

        message = response.choices[0].message
        messages.append(message)

        # No tool call → final answer
        if not message.tool_calls:
            return message.content, execution_trace

        # Execute tools
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

            messages.append(
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

    st.subheader("🤖 Agent 請求")
    user_request = st.text_input(
        "請輸入要求:",
        value="從 Outlook 找出最新的有效發票",
    )

    if st.button("🚀 執行 Agent", type="primary"):
        if not backend_health:
            st.error("❌ 無法連接後端。請確保後端正在運行: python backend.py")
        else:
            st.divider()
            st.subheader("🔎 Agent 執行過程")
            with st.spinner("Agent 正在處理..."):
                try:
                    answer, trace = run_agent(user_request)

                    # Execution trace
                    for index, step in enumerate(trace, start=1):
                        tool = step["tool"]
                        arguments = step["arguments"]

                        st.write(f"**步驟 {index}:** `{tool}`")
                        if arguments:
                            st.caption(f"參數: {json.dumps(arguments, ensure_ascii=False)}")

                    # Final result
                    st.divider()
                    st.subheader("📋 Agent 結果")
                    st.markdown(answer)

                except Exception as e:
                    st.error(f"❌ 執行失敗: {str(e)}")

st.divider()
st.caption("💡 提示: 確保環境變數已設置 (AAD_CLIENT_ID, AAD_TENANT_ID, AAD_CLIENT_SECRET, OPENAI_API_KEY)")

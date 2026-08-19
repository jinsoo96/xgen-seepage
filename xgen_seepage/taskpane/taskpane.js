/* xgen-seepage 채팅 패널. office.js에 의존하지 않는다 - 이 페이지는 Office
   JavaScript API를 하나도 쓰지 않는다(셀 편집은 커넥터의 WS 브릿지가 하지
   이 페이지가 하는 게 아니다). 그래서 Office 호스트를 기다릴 이유가 없고,
   XGEN 연결(로컬 커넥터의 /workflows·/chat/stream)에만 의존한다. 순수
   DOM 로드 시점에 바로 뜬다. */

const statusEl = document.getElementById("status");
const serverInfoEl = document.getElementById("serverinfo");
const messagesEl = document.getElementById("messages");
const form = document.getElementById("composer");
const input = document.getElementById("input");
const sendBtn = document.getElementById("send");
const agentSel = document.getElementById("agent");
const reloadBtn = document.getElementById("reload");
const toastEl = document.getElementById("toast");

function setStatus(text, cls) {
  statusEl.textContent = text;
  statusEl.className = "status" + (cls ? " " + cls : "");
}

/** 화면 하단에 잠깐 뜨는 알림 토스트. type: "warn"|"error"|"info". */
let _toastTimer = null;
function showToast(text, type) {
  toastEl.textContent = text;
  toastEl.className = "toast show" + (type ? " " + type : "");
  if (_toastTimer) clearTimeout(_toastTimer);
  _toastTimer = setTimeout(() => { toastEl.className = "toast"; }, 6000);
}

/** 실행 에러가 LLM 모델 미설정으로 보이면, 캔버스에서 노드 모델을 설정하라고
 * 토스트로 안내한다(패널엔 모델 토글이 없다 - 모델은 에이전트플로우의
 * 에이전트 노드 설정을 그대로 쓴다). 그 외 에러는 일반 에러 토스트. */
function maybeLlmToast(detail) {
  const s = String(detail || "").toLowerCase();
  if (/no model loaded|model|provider|connection attempts failed|llm|503/.test(s)) {
    showToast(
      "이 에이전트에 LLM 모델이 설정돼 있지 않은 것 같습니다. XGEN 캔버스에서 " +
      "에이전트 노드의 모델을 설정한 뒤 다시 시도하세요.",
      "warn"
    );
  } else {
    showToast(`실행 에러: ${detail}`.slice(0, 200), "error");
  }
}

function addMessage(text, cls) {
  const div = document.createElement("div");
  div.className = "msg " + cls;
  div.textContent = text;
  messagesEl.appendChild(div);
  messagesEl.scrollTop = messagesEl.scrollHeight;
  return div;
}

/** 로그인된 XGEN 계정이 가진 워크플로우(에이전트)를 드롭다운에 채운다.
 * 이걸로 CLI에서 미리 하나 박아둘 필요 없이 패널 안에서 바로 골라 쓴다. */
async function loadAgents() {
  agentSel.disabled = true;
  reloadBtn.disabled = true;
  agentSel.innerHTML = '<option value="">불러오는 중...</option>';
  try {
    const resp = await fetch("/workflows");
    if (!resp.ok) {
      const body = await resp.json().catch(() => ({}));
      agentSel.innerHTML = '<option value="">(목록 불러오기 실패)</option>';
      setStatus(`에이전트 목록 실패: ${body.message || resp.status}`, "error");
      return;
    }
    const data = await resp.json();
    const list = data.workflows || [];
    if (!list.length) {
      agentSel.innerHTML = '<option value="">(에이전트 없음 - XGEN 캔버스에서 먼저 만드세요)</option>';
      return;
    }
    agentSel.innerHTML = "";
    for (const w of list) {
      const opt = document.createElement("option");
      opt.value = w.workflow_id;
      opt.textContent = w.workflow_name || w.workflow_id;
      if (w.workflow_id === data.current) opt.selected = true;
      agentSel.appendChild(opt);
    }
    agentSel.disabled = false;
    reloadBtn.disabled = false;
  } catch (e) {
    agentSel.innerHTML = '<option value="">(연결 실패)</option>';
    setStatus(`에이전트 목록 연결 실패: ${e.message}`, "error");
  }
}

/** XGEN 커넥터(로컬 서버)에 붙어 있는지 확인하고 준비 상태로 만든다.
 * office.js가 아니라 이 연결이 이 패널의 유일한 의존성이다. */
/** 지금 붙은 XGEN 서버와 로그인 계정을 헤더에 보여준다. 엉뚱한 서버(jeju
 * 대신 prod 등)에 붙어 403이 나는 상황을, 사용자가 눈으로 바로 잡게 한다. */
async function loadServerInfo() {
  try {
    const resp = await fetch("/server");
    if (!resp.ok) return;
    const data = await resp.json();
    if (data.server_url) {
      const who = data.username ? ` · ${data.username}` : "";
      serverInfoEl.textContent = `${data.server_url}${who}`;
    }
  } catch (e) {
    /* 서버 정보 실패는 치명적이지 않다 */
  }
}

async function init() {
  input.disabled = false;
  sendBtn.disabled = false;
  input.focus();
  try {
    const resp = await fetch("/health");
    if (resp.ok) {
      setStatus("XGEN 연결됨", "ready");
    } else {
      setStatus(`커넥터 응답 이상: ${resp.status}`, "error");
    }
  } catch (e) {
    setStatus(`커넥터에 연결할 수 없습니다. \`xgen-seepage run\`이 켜져 있는지 확인하세요.`, "error");
  }
  loadServerInfo();
  loadAgents();
}

reloadBtn.addEventListener("click", loadAgents);
window.addEventListener("DOMContentLoaded", init);

/** SSE 원시 텍스트 프레임을 파싱한다. `event:`/`data:` 줄 + 빈 줄로 프레임
 * 구분(표준 SSE). XGEN이 보내는 실제 이벤트 이름(log/node_status/tool/
 * execution_io)과 무명 data 프레임({"type":"data"|"end"|"error",...})을
 * 그대로 넘긴다 - 여기서 XGEN의 실제 형태를 재해석하지 않는다. */
function parseSseFrame(frame) {
  let eventName = "message";
  const dataLines = [];
  for (const line of frame.split("\n")) {
    if (line.startsWith("event:")) eventName = line.slice(6).trim();
    else if (line.startsWith("data:")) dataLines.push(line.slice(5).trim());
  }
  const dataText = dataLines.join("\n");
  let data = null;
  if (dataText) {
    try {
      data = JSON.parse(dataText);
    } catch (e) {
      data = dataText;
    }
  }
  return { event: eventName, data };
}

function handleEvent(eventName, data, assistantDiv) {
  if (eventName === "tool") {
    const t = data && data.type;
    if (t === "tool_start" || t === "tool_call") {
      addMessage(`[도구 호출] ${data.tool_name || data.tool || ""}`, "tool");
    } else if (t === "tool_result") {
      addMessage(`[도구 결과] ${data.tool_name || ""} 완료`, "tool");
    }
    // 답변 블록을 끊는다: 도구 호출 다음에 이어지는 답변 토큰이 도구 메시지
    // 위의 기존 블록에 붙지 않고, 도구 아래에 새 블록으로 흐르게 한다(도구
    // 활동은 발생 순서대로 위에, 최종 답변은 그 아래에).
    return null;
  }
  if (eventName === "node_status") {
    return assistantDiv; // 상세 로그, 채팅창엔 안 보여줌
  }
  if (eventName === "log") {
    return assistantDiv;
  }
  // 무명 data 프레임: {"type":"data","content":...} 토큰 스트림 /
  // {"type":"end",...} 완료 / {"type":"error",...} 에러
  if (data && typeof data === "object") {
    if (data.type === "data" && typeof data.content === "string") {
      if (!assistantDiv) assistantDiv = addMessage("", "assistant");
      assistantDiv.textContent += data.content;
      messagesEl.scrollTop = messagesEl.scrollHeight;
      return assistantDiv;
    }
    if (data.type === "error") {
      const detail = data.detail || JSON.stringify(data);
      addMessage(`[에러] ${detail}`, "error");
      maybeLlmToast(detail);
      return assistantDiv;
    }
    if (data.type === "end") {
      return assistantDiv;
    }
  }
  return assistantDiv;
}

async function sendMessage(message) {
  const workflowId = agentSel.value;
  if (!workflowId) {
    addMessage("[안내] 먼저 위에서 에이전트를 선택하세요.", "error");
    return;
  }
  addMessage(message, "user");
  input.value = "";
  input.disabled = true;
  sendBtn.disabled = true;
  let assistantDiv = null;
  try {
    const resp = await fetch("/chat/stream", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ message, workflow_id: workflowId }),
    });
    if (!resp.ok) {
      const body = await resp.json().catch(() => ({}));
      const detail = body.message || resp.statusText;
      addMessage(`[에러] ${resp.status} ${detail}`, "error");
      maybeLlmToast(`${resp.status} ${detail}`);
      return;
    }
    const reader = resp.body.getReader();
    const decoder = new TextDecoder("utf-8");
    let buffer = "";
    while (true) {
      const { done, value } = await reader.read();
      if (done) break;
      buffer += decoder.decode(value, { stream: true });
      let sep;
      while ((sep = buffer.indexOf("\n\n")) !== -1) {
        const frame = buffer.slice(0, sep);
        buffer = buffer.slice(sep + 2);
        if (!frame.trim()) continue;
        const { event, data } = parseSseFrame(frame);
        assistantDiv = handleEvent(event, data, assistantDiv);
      }
    }
  } catch (e) {
    addMessage(`[연결 실패] ${e.message}`, "error");
    showToast(`커넥터 연결 실패: ${e.message}`, "error");
  } finally {
    input.disabled = false;
    sendBtn.disabled = false;
    input.focus();
  }
}

form.addEventListener("submit", (e) => {
  e.preventDefault();
  const message = input.value.trim();
  if (message) sendMessage(message);
});

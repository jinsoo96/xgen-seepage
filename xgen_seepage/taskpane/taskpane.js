/* xgen-seepage 채팅 패널. office.js에 의존하지 않는다 - 이 페이지는 Office
   JavaScript API를 하나도 쓰지 않는다(셀 편집은 커넥터의 WS 브릿지가 하지
   이 페이지가 하는 게 아니다). 그래서 Office 호스트를 기다릴 이유가 없고,
   XGEN 연결(로컬 커넥터의 /workflows·/chat/stream)에만 의존한다. 순수
   DOM 로드 시점에 바로 뜬다. */

const statusEl = document.getElementById("status");
const messagesEl = document.getElementById("messages");
const form = document.getElementById("composer");
const input = document.getElementById("input");
const sendBtn = document.getElementById("send");
const agentSel = document.getElementById("agent");
const reloadBtn = document.getElementById("reload");
const providerSel = document.getElementById("provider");

function setStatus(text, cls) {
  statusEl.textContent = text;
  statusEl.className = "status" + (cls ? " " + cls : "");
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
/** provider/model 드롭다운을 채운다. 서버에 설정된 provider의 모델을 골라
 * 실행에 주입 - 워크플로우 provider가 비어 "No model loaded 503"이 나는 걸
 * 피한다. "서버 기본값"을 그대로 두면 주입 안 하고 서버가 알아서 고른다. */
async function loadProviders() {
  try {
    const resp = await fetch("/providers");
    if (!resp.ok) return;
    const data = await resp.json();
    const list = data.providers || [];
    providerSel.innerHTML = '<option value="">서버 기본값</option>';
    for (const p of list) {
      const opt = document.createElement("option");
      opt.value = p.provider;
      opt.dataset.model = p.default_model || "";
      opt.textContent = p.default_model ? `${p.provider} (${p.default_model})` : p.provider;
      providerSel.appendChild(opt);
    }
    providerSel.disabled = false;
  } catch (e) {
    /* provider 목록 실패는 치명적이지 않다 - 서버 기본값으로 진행 */
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
  loadAgents();
  loadProviders();
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
    return assistantDiv;
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
      addMessage(`[에러] ${data.detail || JSON.stringify(data)}`, "error");
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
    const provider = providerSel.value;
    const model = provider ? (providerSel.selectedOptions[0].dataset.model || "") : "";
    const resp = await fetch("/chat/stream", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ message, workflow_id: workflowId, provider, model }),
    });
    if (!resp.ok) {
      const body = await resp.json().catch(() => ({}));
      addMessage(`[에러] ${resp.status} ${body.message || resp.statusText}`, "error");
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

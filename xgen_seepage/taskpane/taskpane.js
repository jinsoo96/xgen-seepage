/* xgen-seepage 태스크팬 채팅. 셀은 직접 안 만진다 - 그건 이미 도는
   connector 프로세스의 WS 브릿지(tools.call_tool -> live_adapter/
   libreoffice_adapter)가 한다. 이 스크립트는 /chat/stream을 호출해서
   XGEN이 실제로 보내는 SSE를 그대로 파싱해 보여주는 채팅창일 뿐이다. */

const statusEl = document.getElementById("status");
const messagesEl = document.getElementById("messages");
const form = document.getElementById("composer");
const input = document.getElementById("input");
const sendBtn = document.getElementById("send");

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

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    setStatus(`준비됨 (Excel, ${Office.context.diagnostics.version})`, "ready");
  } else {
    setStatus(`준비됨 (host=${info.host})`, "ready");
  }
  input.disabled = false;
  sendBtn.disabled = false;
  input.focus();
});

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
  addMessage(message, "user");
  input.value = "";
  input.disabled = true;
  sendBtn.disabled = true;
  let assistantDiv = null;
  try {
    const resp = await fetch("/chat/stream", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ message }),
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

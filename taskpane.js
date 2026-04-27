const PROXY = "http://127.0.0.1:8765";

let messages = [];
let isKbMode = false;
let kbConversationId = "";
let quotedText = "";
let isSending = false;

const dom = {};

function mdToHtml(text) {
  if (!text) return "";
  let s = text
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");

  s = s.replace(/```(\w*)\n?([\s\S]*?)```/g, (_, lang, code) => `<pre><code>${code.trim()}</code></pre>`);
  s = s.replace(/`([^`]+)`/g, "<code>$1</code>");
  s = s.replace(/\*\*(.+?)\*\*/g, "<strong>$1</strong>");
  s = s.replace(/\*(.+?)\*/g, "<em>$1</em>");
  s = s.replace(/^[ \t]*[-*] (.+)$/gm, "<li>$1</li>");
  s = s.replace(/(<li>[\s\S]*?<\/li>)/g, "<ul>$1</ul>");
  s = s.replace(/<\/ul>\s*<ul>/g, "");
  s = s.replace(/^[ \t]*\d+\. (.+)$/gm, "<li>$1</li>");
  s = s.replace(/\n\n+/g, "</p><p>");
  s = s.replace(/\n/g, "<br>");
  s = `<p>${s}</p>`;
  s = s.replace(/<p>\s*<\/p>/g, "");
  return s;
}

function renderMessage(role, content, isTyping = false) {
  const wrap = document.createElement("div");
  wrap.className = `message ${role === "user" ? "user" : "assistant"}`;

  const avatar = document.createElement("div");
  avatar.className = "avatar";
  avatar.textContent = role === "user" ? "我" : "AI";

  const bubble = document.createElement("div");
  bubble.className = "bubble";

  if (isTyping) {
    bubble.innerHTML = `<div class="typing-indicator"><span></span><span></span><span></span></div>`;
  } else {
    bubble.innerHTML = mdToHtml(content);
  }

  wrap.appendChild(avatar);
  wrap.appendChild(bubble);
  dom.messages.appendChild(wrap);
  dom.messages.scrollTop = dom.messages.scrollHeight;
  return bubble;
}

function updateBubble(bubble, content) {
  bubble.innerHTML = mdToHtml(content);
  dom.messages.scrollTop = dom.messages.scrollHeight;
}

function showWelcome() {
  const welcome = [
    "你好，我是 **AI智能问答助手**。",
    "- 支持普通对话与知识库问答",
    "- 支持引用当前 WPS 选中内容",
    "- 新增公文功能：`体检` 与 `规范化`"
  ].join("\n");
  renderMessage("assistant", welcome);
}

async function sendMessage() {
  if (isSending) return;

  const rawInput = dom.input.value.trim();
  if (!rawInput) return;

  isSending = true;
  dom.btnSend.disabled = true;

  let userContent = rawInput;
  if (quotedText) {
    userContent = `【引用文档内容】\n${quotedText}\n\n【问题】\n${rawInput}`;
    clearQuote();
  }

  dom.input.value = "";
  renderMessage("user", userContent);
  messages.push({ role: "user", content: userContent });
  const typingBubble = renderMessage("assistant", "", true);

  try {
    let answer = "";
    if (isKbMode) {
      const resp = await fetch(`${PROXY}/api/kb`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ query: userContent, conversation_id: kbConversationId })
      });
      if (!resp.ok) throw new Error(`HTTP ${resp.status}`);
      const data = await resp.json();
      if (data.error) throw new Error(data.error);
      answer = data.answer || "（知识库未返回内容）";
      kbConversationId = data.conversation_id || kbConversationId;
    } else {
      const systemMsg = {
        role: "system",
        content: "你是一位专业的企业内部AI助手，请用中文给出简洁、准确、可执行的答复。"
      };
      const history = messages.slice(-20);
      const resp = await fetch(`${PROXY}/api/chat`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ messages: [systemMsg, ...history] })
      });
      if (!resp.ok) throw new Error(`HTTP ${resp.status}`);
      const data = await resp.json();
      if (data.error) throw new Error(data.error);
      answer = data.choices?.[0]?.message?.content || "（未收到回复）";
    }

    updateBubble(typingBubble, answer);
    messages.push({ role: "assistant", content: answer });
  } catch (e) {
    typingBubble.parentElement.classList.add("error");
    updateBubble(
      typingBubble,
      `请求失败：${e.message}\n\n请确认代理服务已启动：\n\`\`\`\npython proxy.py\n\`\`\``
    );
    messages.pop();
  }

  isSending = false;
  dom.btnSend.disabled = false;
  dom.input.focus();
}

function quoteFromWPS() {
  let text = "";

  try {
    if (typeof Application !== "undefined") {
      try {
        text = Application.ActiveDocument?.ActiveWindow?.Selection?.Text || "";
      } catch (e) {}

      if (!text.trim()) {
        try {
          const sel = Application.Selection;
          if (sel && sel.Rows) {
            const rows = [];
            for (let r = 1; r <= sel.Rows.Count; r++) {
              const cols = [];
              for (let c = 1; c <= sel.Columns.Count; c++) {
                const v = sel.Cells(r, c).Value;
                if (v !== null && v !== undefined) cols.push(String(v));
              }
              if (cols.length) rows.push(cols.join("\t"));
            }
            text = rows.join("\n");
          }
        } catch (e) {}
      }
    }
  } catch (e) {}

  if (!text.trim()) {
    showToast("没有获取到选中内容，请先在文档中选中要引用的文字");
    return;
  }

  setQuote(text.trim());
  showToast("已引用选中内容");
}

function setQuote(text) {
  quotedText = text;
  dom.quotePreviw.textContent = `${text.substring(0, 80)}${text.length > 80 ? "..." : ""}`;
  dom.quoteBar.classList.remove("hidden");
  dom.input.focus();
}

function clearQuote() {
  quotedText = "";
  dom.quoteBar.classList.add("hidden");
  dom.quotePreviw.textContent = "";
}

function clearChat() {
  messages = [];
  kbConversationId = "";
  dom.messages.innerHTML = "";
  showWelcome();
}

function toggleMode() {
  isKbMode = !isKbMode;
  kbConversationId = "";
  if (isKbMode) {
    dom.modeBar.className = "mode-kb";
    dom.modeLabel.textContent = "📚 知识库模式";
    dom.btnKbToggle.classList.add("active");
    showToast("已切换到知识库模式");
  } else {
    dom.modeBar.className = "mode-normal";
    dom.modeLabel.textContent = "💬 普通对话";
    dom.btnKbToggle.classList.remove("active");
    showToast("已切换到普通对话模式");
  }
}

let toastTimer = null;
function showToast(msg) {
  let toast = document.getElementById("status-toast");
  if (!toast) {
    toast = document.createElement("div");
    toast.id = "status-toast";
    document.body.appendChild(toast);
  }
  toast.textContent = msg;
  toast.classList.add("show");
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => toast.classList.remove("show"), 2500);
}

function formatScanResultMarkdown(result) {
  const issues = result.issues || [];
  const lines = [
    "## 公文格式体检结果",
    `- 标准：${result.standard || "GB/T 9704-2012"}`,
    `- 段落数：${result.paragraphCount ?? 0}`,
    `- 问题数：${issues.length}`
  ];

  if (!issues.length) {
    lines.push("", "未发现明显格式问题。");
    return lines.join("\n");
  }

  lines.push("", "### 问题清单");
  issues.slice(0, 20).forEach((item, idx) => {
    const prefix = item.level === "warn" ? "⚠️" : "ℹ️";
    lines.push(`${idx + 1}. ${prefix} ${item.message}`);
  });

  if (issues.length > 20) {
    lines.push(`\n其余 ${issues.length - 20} 条请在文档中继续人工核对。`);
  }
  return lines.join("\n");
}

async function runGovScan() {
  if (!window.GovDocFormatter) {
    renderMessage("assistant", "公文格式模块未加载。");
    return;
  }
  const typing = renderMessage("assistant", "正在执行公文格式体检...", true);
  try {
    const result = await window.GovDocFormatter.scan();
    if (!result.ok) {
      updateBubble(typing, `体检失败：${result.message}`);
      return;
    }
    updateBubble(typing, formatScanResultMarkdown(result));
  } catch (e) {
    updateBubble(typing, `体检异常：${e.message}`);
  }
}

function formatApplyResultMarkdown(result) {
  const remain = result.remainingIssues || [];
  const diagnostics = result.diagnostics || {};
  const page = diagnostics.page || {};
  const lines = [
    "## 一键规范化完成",
    `- 标准：${result.standard || "GB/T 9704-2012"}`,
    `- 已调整项：${result.changed ?? 0}`,
    `- 剩余提示项：${remain.length}`,
    `- 行距写入：${diagnostics.lineSpacingVerified ? "已执行" : "未确认"}`,
    `- 页码状态：${diagnostics.pageNumberVerified ? "已存在" : "未检测到"}`,
    `- 页码诊断：奇数页=${page.footer1PageNums ?? 0}个, 偶数页=${page.footer3PageNums ?? 0}个`
  ];
  if (remain.length) {
    lines.push("", "### 仍需人工确认");
    remain.slice(0, 8).forEach((item, idx) => {
      lines.push(`${idx + 1}. ${item.message}`);
    });
  }
  return lines.join("\n");
}

async function runGovApply() {
  if (!window.GovDocFormatter) {
    renderMessage("assistant", "公文格式模块未加载。");
    return;
  }

  const confirmed = window.confirm("将按公文规则统一页面、标题和正文样式，是否继续？");
  if (!confirmed) return;

  const typing = renderMessage("assistant", "正在执行一键规范化...", true);
  try {
    const result = await window.GovDocFormatter.apply();
    if (!result.ok) {
      updateBubble(typing, `规范化失败：${result.message}`);
      return;
    }
    updateBubble(typing, formatApplyResultMarkdown(result));
  } catch (e) {
    updateBubble(typing, `规范化异常：${e.message}`);
  }
}

function handleUrlParams() {
  const params = new URLSearchParams(window.location.search);

  if (params.get("action") === "clear") {
    clearChat();
    return;
  }
  if (params.get("action") === "gov-scan") {
    runGovScan();
  }
  if (params.get("action") === "gov-apply") {
    runGovApply();
  }

  const quote = params.get("quote");
  if (quote) {
    setQuote(decodeURIComponent(quote));
  }

  if (params.toString()) {
    const cleanUrl = window.location.href.split("?")[0];
    history.replaceState({}, "", cleanUrl);
  }
}

async function checkProxy() {
  try {
    const resp = await fetch(`${PROXY}/api/status`, { method: "GET" });
    if (!resp.ok) throw new Error(`status ${resp.status}`);
    const data = await resp.json();
    showToast(`代理已就绪，模型：${data.model}`);
  } catch (e) {
    renderMessage(
      "assistant",
      [
        "⚠️ 无法连接代理服务。",
        "",
        "请先运行：",
        "```",
        "python proxy.py",
        "```",
        `错误：${e.message}`
      ].join("\n")
    );
  }
}

document.addEventListener("DOMContentLoaded", () => {
  dom.messages = document.getElementById("messages");
  dom.input = document.getElementById("input");
  dom.btnSend = document.getElementById("btn-send");
  dom.btnKbToggle = document.getElementById("btn-kb-toggle");
  dom.btnClear = document.getElementById("btn-clear");
  dom.btnQuoteWps = document.getElementById("btn-quote-wps");
  dom.modeBar = document.getElementById("mode-bar");
  dom.modeLabel = document.getElementById("mode-label");
  dom.quoteBar = document.getElementById("quote-bar");
  dom.quotePreviw = document.getElementById("quote-preview-text");
  dom.btnRemoveQ = document.getElementById("btn-remove-quote");
  dom.btnGovScan = document.getElementById("btn-gov-scan");
  dom.btnGovApply = document.getElementById("btn-gov-apply");

  dom.btnSend.addEventListener("click", sendMessage);
  dom.btnKbToggle.addEventListener("click", toggleMode);
  dom.btnClear.addEventListener("click", clearChat);
  dom.btnQuoteWps.addEventListener("click", quoteFromWPS);
  dom.btnRemoveQ.addEventListener("click", clearQuote);
  dom.btnGovScan.addEventListener("click", runGovScan);
  dom.btnGovApply.addEventListener("click", runGovApply);

  dom.input.addEventListener("keydown", (e) => {
    if (e.ctrlKey && e.key === "Enter") {
      e.preventDefault();
      sendMessage();
    }
  });

  dom.input.addEventListener("input", () => {
    dom.input.style.height = "auto";
    dom.input.style.height = `${Math.min(dom.input.scrollHeight, 120)}px`;
  });

  handleUrlParams();
  showWelcome();
  checkProxy();
});

/**
 * WPS plugin entry for ribbon actions.
 */

const TASKPANE_URL = "http://192.168.9.226:8765/taskpane.html?v=8";
let g_taskPane = null;

function OnRibbonLoad(ribbon) {}

function OpenAIPanel(control) {
  try {
    if (g_taskPane && g_taskPane.Visible) {
      g_taskPane.Visible = false;
      return;
    }
    if (!g_taskPane) g_taskPane = _createTaskPane(TASKPANE_URL);
    if (g_taskPane) {
      g_taskPane.Visible = true;
    } else {
      Application.ShowDialog(TASKPANE_URL, "AI 智能问答助手");
    }
  } catch (e) {
    Application.Alert(`打开 AI 助手失败：${e.message}`);
  }
}

function QuoteSelectedText(control) {
  try {
    const text = _getSelectedText();
    if (!text || !text.trim()) {
      Application.Alert("没有选中文本，请先在文档中选中要引用的内容。");
      return;
    }
    const quoteUrl = `${TASKPANE_URL}?quote=${encodeURIComponent(text.trim().substring(0, 3000))}`;
    _navigateTaskPane(quoteUrl);
  } catch (e) {
    Application.Alert(`引用失败：${e.message}`);
  }
}

function ClearChatHistory(control) {
  try {
    _navigateTaskPane(`${TASKPANE_URL}?action=clear`);
  } catch (e) {
    Application.Alert(`清空失败：${e.message}`);
  }
}

function ScanGovDocFormat(control) {
  try {
    _navigateTaskPane(`${TASKPANE_URL}?action=gov-scan`);
  } catch (e) {
    Application.Alert(`公文体检失败：${e.message}`);
  }
}

function ApplyGovDocFormat(control) {
  try {
    _navigateTaskPane(`${TASKPANE_URL}?action=gov-apply`);
  } catch (e) {
    Application.Alert(`公文规范化失败：${e.message}`);
  }
}

function _navigateTaskPane(url) {
  if (g_taskPane) {
    try {
      g_taskPane.Navigate(url);
      g_taskPane.Visible = true;
      return;
    } catch (e) {
      g_taskPane = null;
    }
  }

  g_taskPane = _createTaskPane(url);
  if (g_taskPane) {
    g_taskPane.Visible = true;
  } else {
    Application.ShowDialog(url, "AI 智能问答助手");
  }
}

function _createTaskPane(url) {
  try {
    if (typeof Application.CreateTaskPane === "function") {
      const pane = Application.CreateTaskPane(url);
      try {
        pane.Width = 380;
      } catch (e) {}
      return pane;
    }
  } catch (e) {}

  try {
    const win =
      (Application.ActiveDocument && Application.ActiveDocument.ActiveWindow) ||
      (Application.ActiveWorkbook && Application.ActiveWorkbook.Windows(1)) ||
      (Application.ActivePresentation && Application.ActivePresentation.ActiveWindow);

    if (win && typeof win.CreateTaskPane === "function") {
      const pane = win.CreateTaskPane(url);
      try {
        pane.Width = 380;
      } catch (e) {}
      return pane;
    }
  } catch (e) {}

  return null;
}

function _getSelectedText() {
  try {
    const doc = Application.ActiveDocument;
    if (doc) {
      const text = doc.ActiveWindow ? doc.ActiveWindow.Selection.Text : Application.Selection.Text;
      if (text && text.trim()) return text;
    }
  } catch (e) {}

  try {
    const wb = Application.ActiveWorkbook;
    if (wb) {
      const sel = Application.Selection;
      if (sel && sel.Rows) {
        const rows = [];
        for (let r = 1; r <= sel.Rows.Count; r++) {
          const cols = [];
          for (let c = 1; c <= sel.Columns.Count; c++) {
            const v = sel.Cells(r, c).Value;
            if (v !== null && v !== undefined) cols.push(String(v));
          }
          if (cols.length > 0) rows.push(cols.join("\t"));
        }
        if (rows.length > 0) return rows.join("\n");
      }
    }
  } catch (e) {}

  try {
    const ppt = Application.ActivePresentation;
    if (ppt) {
      const slide = ppt.ActiveSlide || ppt.Slides(1);
      const texts = [];
      for (let i = 1; i <= slide.Shapes.Count; i++) {
        try {
          const shape = slide.Shapes.Item(i);
          if (shape.HasTextFrame) {
            const t = shape.TextFrame.TextRange.Text;
            if (t && t.trim()) texts.push(t.trim());
          }
        } catch (e) {}
      }
      if (texts.length > 0) return texts.join("\n\n");
    }
  } catch (e) {}

  return "";
}
